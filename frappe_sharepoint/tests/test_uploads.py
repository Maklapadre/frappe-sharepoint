"""Upload regressions using real local files and a simulated Graph boundary.

Run: python -m unittest discover -s frappe_sharepoint/tests -v
No live Microsoft credentials or network access are used.
"""

import io
import tempfile
import unittest
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import Mock, patch

import frappe
import requests
from frappe.core.doctype.file.file import File

from frappe_sharepoint import utils
from frappe_sharepoint.controllers.file_controller import file_upload, get_file_path
from frappe_sharepoint.utils.sharepoint import SharePoint, trigger_sharepoint_upload


def graph_response(status=201, size=0, **values):
	response = requests.Response()
	response.status_code = status
	response._content = b"Graph test response"
	response.json = Mock(return_value={"id": "remote-file", "size": size, **values})
	return response


class TestUploads(unittest.TestCase):
	def setUp(self):
		self.temp = tempfile.TemporaryDirectory()
		self.addCleanup(self.temp.cleanup)
		self.root = Path(self.temp.name)
		self.content = b"original attachment\x00\xff\r\n" * 100
		self.path = self.root / "private" / "files" / "invoice_12ab34.pdf"
		self.path.parent.mkdir(parents=True)
		self.path.write_bytes(self.content)
		self.file = SimpleNamespace(
			name="FILE-1", file_name="invoice.pdf", file_url="/private/files/invoice_12ab34.pdf",
			is_private=1, is_folder=0, is_remote_file=False, uploaded_to_sharepoint=0,
			attached_to_doctype="Communication", attached_to_name="COMM-1",
			get_full_path=Mock(return_value=str(self.path)),
		)
		self.settings = SimpleNamespace(
			enable_file_sync=1, sharepoint_drive_id="drive", root_folder_path="",
			folder_structure="Module/DocType/Document", graph_api_url="https://graph.test/v1.0",
			replace_file_link=0, tenant_id="tenant", client_id="client",
			get_password=lambda name: "test-secret",
		)
		self.db = Mock()
		for target, value in (
			("frappe.db", self.db),
			("frappe.session", SimpleNamespace(user="Administrator")),
			("frappe.logger", Mock(return_value=Mock())),
			("frappe.log_error", Mock()),
			("frappe.get_doc", Mock(return_value=self.file)),
			("frappe.get_single", Mock(return_value=self.settings)),
			("frappe_sharepoint.utils.sharepoint.get_request_header", Mock(return_value={})),
		):
			patcher = patch(target, value)
			patcher.start()
			self.addCleanup(patcher.stop)
		# Any accidental real network operation fails the test.
		patcher = patch("requests.sessions.Session.request", side_effect=AssertionError("Unexpected network"))
		patcher.start()
		self.addCleanup(patcher.stop)

	def test_hook_enqueues_only_id_after_commit(self):
		with patch("frappe.enqueue") as enqueue:
			file_upload(self.file, "after_insert")
			enqueue.assert_called_once_with(
				"frappe_sharepoint.utils.sharepoint.trigger_sharepoint_upload",
				queue="long", filedoc="FILE-1", enqueue_after_commit=True,
			)
		self.file.get_full_path.assert_not_called()

	def test_hook_skips_remote_already_uploaded_unattached_and_disabled(self):
		for field, value in (("is_remote_file", True), ("uploaded_to_sharepoint", 1),
			("attached_to_name", None), ("is_folder", 1)):
			with self.subTest(field=field), patch("frappe.enqueue") as enqueue:
				previous = getattr(self.file, field)
				setattr(self.file, field, value)
				file_upload(self.file, "after_insert")
				enqueue.assert_not_called()
				setattr(self.file, field, previous)
		self.settings.enable_file_sync = 0
		with patch("frappe.enqueue") as enqueue:
			file_upload(self.file, "after_insert")
			enqueue.assert_not_called()

	def test_real_frappe_path_resolution_uses_stored_public_and_private_suffix(self):
		def files_path(*parts, is_private=0):
			return str(self.root.joinpath("private" if is_private else "public", "files", *parts))
		with patch("frappe.core.doctype.file.file.get_url", return_value="https://erp.test"), \
			patch("frappe.core.doctype.file.file.get_files_path", side_effect=files_path), \
			patch("frappe.core.doctype.file.file.is_safe_path", return_value=True):
			for private, prefix in ((0, "/files/"), (1, "/private/files/")):
				with self.subTest(private=private):
					self.file.is_private = private
					self.file.file_url = prefix + "invoice_12ab34.pdf"
					self.file.get_full_path = lambda: File.get_full_path(self.file)
					self.assertEqual(get_file_path(self.file), files_path("invoice_12ab34.pdf", is_private=private))

	def test_worker_reloads_current_path_and_communication_ignoring_legacy_arguments(self):
		with patch.object(SharePoint, "build_folder_structure", return_value="folder") as folders:
			streams = []
			def put(url, headers, data, timeout):
				streams.append(data)
				self.assertFalse(data.closed)
				self.assertTrue(url.endswith("/invoice_12ab34.pdf:/content"))
				self.assertEqual(data.read(), self.content)
				return graph_response(size=len(self.content))
			with patch("requests.put", side_effect=put):
				result = trigger_sharepoint_upload("OldType", "OldDoc", "/stale/invoice.pdf", "FILE-1")
			self.assertEqual(result["status"], "uploaded")
			self.assertTrue(streams[0].closed)
			folders.assert_called_once()
		frappe.get_doc.assert_called_once_with("File", "FILE-1")
		self.db.set_value.assert_called_once_with("File", "FILE-1", {"uploaded_to_sharepoint": 1})
		self.assertEqual(self.path.read_bytes(), self.content)
		self.assertEqual(self.file.file_url, "/private/files/invoice_12ab34.pdf")

	def test_worker_uses_current_attachment_target(self):
		with patch("frappe_sharepoint.utils.sharepoint.SharePoint") as sharepoint:
			trigger_sharepoint_upload("Old", "Old", "/old", "FILE-1")
			sharepoint.assert_called_once_with(
				doctype="Communication", docname="COMM-1", filepath=str(self.path), filedoc="FILE-1",
			)

	def test_missing_local_file_fails_job_without_remote_calls_or_success_flag(self):
		self.path.unlink()
		with patch.object(SharePoint, "build_folder_structure") as folders:
			with self.assertRaises(FileNotFoundError):
				trigger_sharepoint_upload(filedoc="FILE-1")
			folders.assert_not_called()
		self.db.set_value.assert_not_called()

	def test_remote_url_is_explicitly_rejected_if_file_changes_after_enqueue(self):
		self.file.is_remote_file = True
		with self.assertRaisesRegex(ValueError, "remote URL"):
			trigger_sharepoint_upload(filedoc="FILE-1")
		self.file.get_full_path.assert_not_called()
		self.db.set_value.assert_not_called()

	def test_deleted_file_fails_job(self):
		frappe.get_doc.side_effect = frappe.DoesNotExistError("FILE-1")
		with self.assertRaises(frappe.DoesNotExistError):
			trigger_sharepoint_upload(filedoc="FILE-1")
		self.db.set_value.assert_not_called()

	def test_worker_skips_previously_uploaded_file(self):
		self.file.uploaded_to_sharepoint = 1
		result = trigger_sharepoint_upload(filedoc="FILE-1")
		self.assertEqual(result["reason"], "already_uploaded")
		self.file.get_full_path.assert_not_called()
		self.db.set_value.assert_not_called()

	def test_auth_failure_closes_stream_and_fails_job(self):
		opened = []
		real_open = open
		def track_open(*args, **kwargs):
			stream = real_open(*args, **kwargs)
			opened.append(stream)
			return stream
		with patch("builtins.open", side_effect=track_open), \
			patch.object(SharePoint, "build_folder_structure", return_value="folder"), \
			patch("frappe_sharepoint.utils.sharepoint.get_request_header", side_effect=frappe.ValidationError("Auth failed")):
			with self.assertRaises(frappe.ValidationError):
				trigger_sharepoint_upload(filedoc="FILE-1")
		self.assertTrue(opened[0].closed)
		self.db.set_value.assert_not_called()

	def test_network_and_http_failures_close_stream_and_fail_job(self):
		for outcome in (requests.Timeout("timeout"), requests.ConnectionError("offline"),
			graph_response(status=401), graph_response(status=500), graph_response(status=202)):
			with self.subTest(outcome=outcome):
				streams = []
				def put(url, headers, data, timeout):
					streams.append(data)
					if isinstance(outcome, Exception):
						raise outcome
					return outcome
				with patch.object(SharePoint, "build_folder_structure", return_value="folder"), \
					patch("requests.put", side_effect=put):
					with self.assertRaises(RuntimeError):
						trigger_sharepoint_upload(filedoc="FILE-1")
				self.assertTrue(streams[0].closed)
				self.db.set_value.assert_not_called()

	def test_missing_or_wrong_upload_confirmation_never_marks_success(self):
		for response in (graph_response(size=-1), graph_response(size=len(self.content), id=None)):
			with self.subTest(response=response), \
				patch.object(SharePoint, "build_folder_structure", return_value="folder"), \
				patch("requests.put", return_value=response):
				with self.assertRaisesRegex(RuntimeError, "confirm"):
					trigger_sharepoint_upload(filedoc="FILE-1")
				self.db.set_value.assert_not_called()

	def test_replacing_link_retains_local_bytes(self):
		self.settings.replace_file_link = 1
		with patch.object(SharePoint, "build_folder_structure", return_value="folder"), \
			patch("requests.put", return_value=graph_response(size=len(self.content), webUrl="https://sharepoint.test/file")):
			trigger_sharepoint_upload(filedoc="FILE-1")
		self.db.set_value.assert_called_once_with("File", "FILE-1", {
			"uploaded_to_sharepoint": 1, "file_url": "https://sharepoint.test/file",
		})
		self.assertEqual(self.path.read_bytes(), self.content)

	def test_put_passes_bytes_and_stream_unchanged_without_reading_for_logging(self):
		for body in (self.content, io.BytesIO(self.content)):
			with self.subTest(body_type=type(body).__name__):
				with patch("requests.put", return_value=graph_response(size=len(self.content))) as put:
					utils.make_request("PUT", "https://graph.test/content", {}, body)
					self.assertIs(put.call_args.kwargs["data"], body)
					if hasattr(body, "tell"):
						self.assertEqual(body.tell(), 0)
						self.assertFalse(body.closed)

	def test_bytes_bundle_path_sends_original_content(self):
		with patch("requests.put", return_value=graph_response(size=len(self.content))) as put:
			sharepoint = SharePoint(doctype="Communication", docname="COMM-1")
			self.assertTrue(sharepoint.upload_file_to_folder("folder", str(self.path), "invoice #1.pdf"))
			self.assertEqual(put.call_args.kwargs["data"], self.content)
			self.assertTrue(put.call_args.args[0].endswith("/invoice%20%231.pdf:/content"))

	def test_folder_lookup_failure_does_not_create_duplicate_folders(self):
		with patch("frappe_sharepoint.utils.sharepoint.make_request", return_value=graph_response(status=503)), \
			patch.object(SharePoint, "create_sharepoint_folder") as create:
			sharepoint = SharePoint()
			with self.assertRaises(RuntimeError):
				sharepoint.get_or_create_folder("parent", "Communication")
			create.assert_not_called()


if __name__ == "__main__":
	unittest.main()
