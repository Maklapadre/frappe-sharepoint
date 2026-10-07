"""Opt-in Frappe database tests; Graph and RQ are replaced at their boundaries.

SHAREPOINT_TEST_SITE=your-dev-site python -m unittest discover -s frappe_sharepoint/tests -v
Use a development site with frappe_sharepoint installed. Temporary records are
removed afterwards; no Microsoft calls or real background jobs are submitted.
"""

import os
import unittest
from contextlib import ExitStack
from pathlib import Path
from unittest.mock import Mock, patch

import frappe

from frappe_sharepoint.utils.sharepoint import SharePoint, trigger_sharepoint_upload


@unittest.skipUnless(os.environ.get("SHAREPOINT_TEST_SITE"), "Requires an explicit development test site")
class TestFrappeUploads(unittest.TestCase):
	def setUp(self):
		self.original_cwd = os.getcwd()
		self.bench = Path(__file__).resolve().parents[4]
		os.chdir(self.bench / "sites")
		frappe.init(site=os.environ["SHAREPOINT_TEST_SITE"])
		frappe.connect()
		frappe.set_user("Administrator")
		self.records = []
		self.paths = []
		self.stack = ExitStack()
		self.addCleanup(self.cleanup)
		self.assertIn("frappe_sharepoint", frappe.get_installed_apps())
		self.queue = Mock()
		self.stack.enter_context(patch("frappe.utils.background_jobs.get_queue", return_value=self.queue))
		self.stack.enter_context(patch("frappe.utils.background_jobs._check_queue_size"))
		self.stack.enter_context(patch("requests.sessions.Session.request", side_effect=AssertionError("Unexpected network")))
		self.stack.enter_context(patch("frappe.sendmail"))
		self.stack.enter_context(patch("frappe.model.document.Document.run_notifications"))
		self.settings = frappe._dict(
			enable_file_sync=1, sharepoint_drive_id="test-drive", folder_structure="Flat",
			graph_api_url="https://graph.test/v1.0", replace_file_link=0,
		)
		get_single = frappe.get_single
		self.stack.enter_context(patch("frappe.get_single", side_effect=lambda name, *a, **kw:
			self.settings if name == "SharePoint Settings" else get_single(name, *a, **kw)))
		self.stack.enter_context(patch("frappe_sharepoint.utils.sharepoint.get_request_header", return_value={}))
		self.stack.enter_context(patch.object(SharePoint, "build_folder_structure", return_value="test-folder"))

	def cleanup(self):
		try:
			frappe.db.rollback()
			for doctype, name in reversed(self.records):
				if frappe.db.exists(doctype, name):
					frappe.delete_doc(doctype, name, ignore_permissions=True, force=True)
			frappe.db.commit()
			for path in self.paths:
				Path(path).unlink(missing_ok=True)
		finally:
			self.stack.close()
			frappe.destroy()
			os.chdir(self.original_cwd)

	def communication(self):
		doc = frappe.get_doc({
			"doctype": "Communication", "communication_type": "Communication",
			"communication_medium": "Other", "subject": "SharePoint upload regression",
			"content": "Disposable integration test", "sent_or_received": "Received",
		}).insert(ignore_permissions=True)
		self.records.append(("Communication", doc.name))
		return doc

	def attachment(self, communication, name, content, private):
		doc = frappe.get_doc({
			"doctype": "File", "file_name": name, "content": content,
			"is_private": private, "attached_to_doctype": "Communication",
			"attached_to_name": communication.name,
		}).insert(ignore_permissions=True)
		self.records.append(("File", doc.name))
		self.paths.append(doc.get_full_path())
		return doc

	def test_committed_public_private_and_suffixed_communication_uploads(self):
		communication = self.communication()
		files = []
		for private in (0, 1):
			name = "sharepoint-regression-" + frappe.generate_hash(length=12) + ".bin"
			for number in (1, 2):
				content = f"private={private}, file={number}".encode() + b"\x00\xff"
				doc = self.attachment(communication, name, content, private)
				files.append((doc.name, doc.file_url, doc.get_full_path(), content))
			self.assertNotEqual(files[-1][1], files[-2][1], "Frappe should suffix the second stored file")
			# Reproduce a display filename that differs from its stored URL.
			frappe.db.set_value("File", files[-1][0], "file_name", name)
		self.queue.enqueue_call.assert_not_called()
		frappe.db.commit()
		self.assertEqual(self.queue.enqueue_call.call_count, 4)
		jobs = [call.kwargs["kwargs"]["kwargs"] for call in self.queue.enqueue_call.call_args_list]
		self.assertEqual(jobs, [{"filedoc": file[0]} for file in files])

		for (name, url, path, content), job in zip(files, jobs):
			with self.subTest(private="private" in url, name=name):
				streams = []
				def put(url, headers, data, timeout):
					streams.append(data)
					self.assertEqual(data.read(), content)
					return Mock(status_code=201, ok=True, json=lambda: {"id": "item", "size": len(content)})
				with patch("requests.put", side_effect=put):
					trigger_sharepoint_upload(**job)
				frappe.db.commit()
				self.assertTrue(streams[0].closed)
				file = frappe.get_doc("File", name)
				self.assertEqual(file.uploaded_to_sharepoint, 1)
				self.assertEqual(file.file_url, url)
				self.assertEqual(file.attached_to_name, communication.name)
				self.assertEqual(file.get_content(), content)
				self.assertEqual(Path(path).read_bytes(), content)

	def test_rollback_never_submits_upload(self):
		communication = self.communication()
		self.attachment(communication, "sharepoint-rollback-" + frappe.generate_hash() + ".bin", b"rollback", 1)
		self.queue.enqueue_call.assert_not_called()
		frappe.db.rollback()
		frappe.db.commit()
		self.queue.enqueue_call.assert_not_called()
