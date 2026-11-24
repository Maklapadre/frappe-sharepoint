"""Exercise backup transfer recovery with local files and a fake HTTP boundary."""

import json
import tempfile
import unittest
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import Mock, patch

import requests

from frappe_sharepoint.utils import verified_upload as upload


def response(status, payload=None, headers=None):
	result = requests.Response()
	result.status_code = status
	result._content = json.dumps(payload or {}).encode()
	result.headers.update(headers or {})
	return result


class TestVerifiedUpload(unittest.TestCase):
	def setUp(self):
		self.temp = tempfile.TemporaryDirectory()
		self.addCleanup(self.temp.cleanup)
		self.path = Path(self.temp.name) / "backup.sql.gz"
		self.path.write_bytes(b"database")
		self.client = SimpleNamespace(base_url="https://graph.test/v1.0/drives/drive", settings=Mock())
		for target, value in (
			("frappe_sharepoint.utils.verified_upload.get_request_header", Mock(return_value={"Authorization": "Bearer secret"})),
			("frappe_sharepoint.utils.verified_upload.time.sleep", Mock()),
			("frappe_sharepoint.utils.verified_upload.CHUNK_SIZE", 327680),
		):
			patcher = patch(target, value)
			patcher.start()
			self.addCleanup(patcher.stop)
		patcher = patch("requests.request")
		self.http = patcher.start()
		self.addCleanup(patcher.stop)

	def upload(self):
		return upload.upload_file(self.client, "folder", self.path)

	def test_small_file_streams_and_rewinds_after_throttle(self):
		bodies = []
		def send(method, url, **kwargs):
			bodies.append(kwargs["data"].read())
			return response(429, headers={"Retry-After": "2"}) if len(bodies) == 1 else response(201, {"id": "file", "size": 8})
		self.http.side_effect = send
		self.assertEqual(self.upload()["id"], "file")
		self.assertEqual(bodies, [b"database", b"database"])
		upload.time.sleep.assert_called_once_with(2)

	def test_http_rejection_keeps_status_code_and_sanitized_graph_details(self):
		self.http.return_value = response(403, {"error": {
			"code": "accessDenied", "message": "Denied https://session.test/?token=hidden client_secret=secret",
			"innerError": {"request-id": "request-123"},
		}})
		with self.assertRaises(upload.SharePointTransferError) as error:
			self.upload()
		message = str(error.exception)
		for expected in ("backup.sql.gz", "8 bytes", "HTTP 403", "accessDenied", "request-123"):
			self.assertIn(expected, message)
		for forbidden in ("https://", "token=hidden", "client_secret=secret", "no HTTP response"):
			self.assertNotIn(forbidden, message)
		self.assertEqual(self.http.call_count, 1)

	def test_non_json_failure_does_not_log_raw_body(self):
		result = response(400)
		result._content = b"<html>secret credentials</html>"
		message = str(upload.response_error(result, "Upload"))
		self.assertIn("HTTP 400", message)
		self.assertNotIn("credentials", message)
		self.assertIn("no HTTP response", str(upload.response_error(None, "Upload")))

	def test_retries_are_bounded_for_service_and_transport_failures(self):
		for failure in (response(503), requests.Timeout("https://secret-session.test/token")):
			with self.subTest(failure=type(failure).__name__):
				self.http.reset_mock()
				self.http.side_effect = [failure] * upload.MAX_ATTEMPTS
				with self.assertRaises(upload.SharePointTransferError) as error:
					self.upload()
				self.assertNotIn("secret-session", str(error.exception))
				self.assertEqual(self.http.call_count, upload.MAX_ATTEMPTS)

	def test_long_retry_after_is_not_retried_early(self):
		self.http.return_value = response(429, headers={"Retry-After": "3600"})
		with self.assertRaisesRegex(upload.SharePointTransferError, "retry this job later"):
			self.upload()
		self.assertEqual(self.http.call_count, 1)
		upload.time.sleep.assert_not_called()

	def test_empty_file_and_wrong_final_size_cannot_succeed(self):
		self.http.return_value = response(201, {"id": "file", "size": 7})
		with self.assertRaisesRegex(upload.SharePointTransferError, "confirm"):
			self.upload()
		self.path.write_bytes(b"")
		self.http.reset_mock()
		with self.assertRaisesRegex(upload.SharePointTransferError, "empty"):
			self.upload()
		self.http.assert_not_called()

	def test_large_file_uses_sequential_aligned_chunks_without_bearer_token(self):
		content = b"a" * (327680 * 2 + 17)
		self.path.write_bytes(content)
		self.http.side_effect = [
			response(200, {"uploadUrl": "https://upload.test/secret"}),
			response(202, {"nextExpectedRanges": ["327680-"]}),
			response(202, {"nextExpectedRanges": ["655360-"]}),
			response(201, {"id": "item", "size": len(content)}),
		]
		self.assertEqual(self.upload()["size"], len(content))
		calls = self.http.call_args_list
		self.assertTrue(calls[0].args[1].endswith(":/createUploadSession"))
		self.assertEqual(calls[0].kwargs["json"]["item"]["@microsoft.graph.conflictBehavior"], "replace")
		self.assertEqual(b"".join(call.kwargs["data"] for call in calls[1:]), content)
		for index, call in enumerate(calls[1:]):
			self.assertNotIn("Authorization", call.kwargs["headers"])
			self.assertLessEqual(len(call.kwargs["data"]), 327680)
			self.assertTrue(call.kwargs["headers"]["Content-Range"].startswith(f"bytes {index * 327680}-"))

	def test_session_queries_expected_range_after_ambiguous_chunk(self):
		self.path.write_bytes(b"a" * 327681)
		self.http.side_effect = [
			response(200, {"uploadUrl": "https://upload.test/secret"}), response(416),
			response(200, {"nextExpectedRanges": ["327680-"]}), response(201, {"id": "file", "size": 327681}),
		]
		self.upload()
		self.assertEqual(self.http.call_args_list[2].args[0], "GET")
		self.assertIsNone(self.http.call_args_list[2].kwargs["headers"])
		self.assertEqual(self.http.call_args_list[3].kwargs["data"], b"a")

	def test_lost_final_response_requires_remote_size_confirmation(self):
		self.path.write_bytes(b"a" * 327681)
		self.http.side_effect = [
			response(200, {"uploadUrl": "https://upload.test/secret"}),
			response(202, {"nextExpectedRanges": ["327680-"]}),
			*[requests.Timeout()] * upload.MAX_ATTEMPTS,
			response(404), response(200, {"id": "file", "size": 327681}),
		]
		self.assertEqual(self.upload()["id"], "file")
		self.assertEqual(self.http.call_args.args[:2], (
			"GET", "https://graph.test/v1.0/drives/drive/items/folder:/backup.sql.gz",
		))

	def test_session_cannot_loop_forever_without_progress(self):
		self.path.write_bytes(b"a" * 327681)
		self.http.side_effect = [response(200, {"uploadUrl": "https://upload.test/secret"})] + [
			response(202, {"nextExpectedRanges": ["0-"]}) for _ in range(upload.MAX_ATTEMPTS)
		]
		with self.assertRaisesRegex(upload.SharePointTransferError, "no progress"):
			self.upload()

	def test_large_upload_honors_long_server_backoff_without_recovering_early(self):
		self.path.write_bytes(b"a" * 327681)
		self.http.side_effect = [response(200, {"uploadUrl": "https://upload.test/secret"}),
			response(429, headers={"Retry-After": "3600"})]
		with self.assertRaises(upload.SharePointRetryLater) as error:
			self.upload()
		self.assertIn("backup.sql.gz (327681 bytes)", str(error.exception))
		self.assertEqual(self.http.call_count, 2)

	def test_quoted_credentials_are_redacted(self):
		self.assertNotIn("hidden", upload.sanitize_message('"client_secret":"hidden" password=hidden Bearer hidden'))

	def test_folder_listing_reads_all_pages_and_rejects_external_next_link(self):
		self.http.side_effect = [response(200, {"value": [{"id": "1", "folder": {}}],
			"@odata.nextLink": self.client.base_url + "/items/folder/children?$skiptoken=next"}),
			response(200, {"value": [{"id": "2", "file": {}}]})]
		self.assertEqual(len(upload.list_children(self.client, "folder")), 2)
		self.http.side_effect = None
		self.http.return_value = response(200, {"value": [], "@odata.nextLink": "https://other.test/steal"})
		with self.assertRaisesRegex(upload.SharePointTransferError, "pagination"):
			upload.list_children(self.client, "folder")
