import frappe

SETTINGS = "SharePoint Settings"


def file_upload(doc, method):
	"""Queue local attachments by ID only after their transaction commits."""
	if (
		method != "after_insert"
		or not doc.attached_to_doctype
		or not doc.attached_to_name
		or doc.uploaded_to_sharepoint
		or doc.is_folder
		or doc.is_remote_file
		or not frappe.db.exists("DocType", SETTINGS)
	):
		return

	if frappe.get_single(SETTINGS).enable_file_sync:
		frappe.enqueue(
			"frappe_sharepoint.utils.sharepoint.trigger_sharepoint_upload",
			queue="long",
			filedoc=doc.name,
			enqueue_after_commit=True,
		)


def get_file_path(doc):
	"""Resolve the stored local File URL, including Frappe's unique suffixes."""
	if doc.is_remote_file:
		raise ValueError(f"File {doc.name} is a remote URL; automatic copying requires a local file")
	return doc.get_full_path()
