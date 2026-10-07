import frappe
from frappe import _
from frappe_sharepoint.utils import get_request_header, make_request

import os
from urllib.parse import quote

from frappe_sharepoint.controllers.file_controller import get_file_path

'''
	SharePoint file synchronization using Direct Drive API
'''

SETTINGS = "SharePoint Settings"
ContentType = {"Content-Type": "application/json"}


def trigger_sharepoint_upload(doctype=None, docname=None, filepath=None, filedoc=None):
	"""Reload the committed File; legacy queued arguments must not override it."""
	if not filedoc:
		raise ValueError("A File document ID is required for automatic SharePoint uploads")

	file = frappe.get_doc("File", filedoc)
	if file.uploaded_to_sharepoint:
		return {"status": "skipped", "reason": "already_uploaded"}
	if file.is_folder or not file.attached_to_doctype or not file.attached_to_name:
		return {"status": "skipped", "reason": "not_an_attachment"}
	if not frappe.get_single(SETTINGS).enable_file_sync:
		return {"status": "skipped", "reason": "sync_disabled"}

	sharepoint = SharePoint(
		doctype=file.attached_to_doctype,
		docname=file.attached_to_name,
		filepath=get_file_path(file),
		filedoc=file.name,
	)
	return sharepoint.run_sharepoint_upload()


def upload_document_bundle(doctype, docname, files):
	"""
	Upload multiple files (document PDF + attachments) to SharePoint
	
	Args:
		doctype: Document type (e.g., "Expense Claim")
		docname: Document name (e.g., "HR-EXP-2025-00033")
		files: List of file dicts with keys: filepath, filename, is_temp
		
	Returns:
		dict: Upload status with success flag and SharePoint folder URL
	"""
	try:
		frappe.logger().info(f"[SharePoint Bundle] Starting upload for {doctype}: {docname} with {len(files)} files")
		
		sharepoint = SharePoint(doctype=doctype, docname=docname, filepath=None, filedoc=None)
		frappe.logger().info(f"[SharePoint Bundle] SharePoint instance created. Drive ID: {sharepoint.drive_id}")
		frappe.logger().info(f"[SharePoint Bundle] Root folder: {sharepoint.root_folder}, Folder structure: {sharepoint.folder_structure}")
		
		# Build the folder structure first
		frappe.logger().info(f"[SharePoint Bundle] Building folder structure...")
		target_folder_id = sharepoint.build_folder_structure()
		frappe.logger().info(f"[SharePoint Bundle] Target folder ID: {target_folder_id}")
		
		if not target_folder_id:
			frappe.logger().error(f"[SharePoint Bundle] Failed to determine target folder")
			return {
				'success': False,
				'message': 'Could not determine target folder in SharePoint'
			}
		
		# Upload each file
		uploaded_count = 0
		failed_files = []
		
		for idx, file_info in enumerate(files):
			filepath = file_info.get('filepath')
			filename = file_info.get('filename')
			
			frappe.logger().info(f"[SharePoint Bundle] File {idx+1}/{len(files)}: {filename}")
			frappe.logger().info(f"[SharePoint Bundle] File path: {filepath}")
			
			if not filepath or not filename:
				frappe.logger().warning(f"[SharePoint Bundle] Skipping file {idx+1} - missing filepath or filename")
				continue
			
			# Upload file with overwrite behavior
			frappe.logger().info(f"[SharePoint Bundle] Uploading {filename} to folder {target_folder_id}")
			success = sharepoint.upload_file_to_folder(
				target_folder_id=target_folder_id,
				filepath=filepath,
				filename=filename
			)
			
			if success:
				uploaded_count += 1
				frappe.logger().info(f"[SharePoint Bundle] Successfully uploaded {filename}")
				# Update File doc if this is an attachment
				if file_info.get('file_doc'):
					frappe.db.set_value("File", file_info['file_doc'], "uploaded_to_sharepoint", 1)
					frappe.logger().info(f"[SharePoint Bundle] Marked File {file_info['file_doc']} as uploaded")
			else:
				failed_files.append(filename)
				frappe.logger().error(f"[SharePoint Bundle] Failed to upload {filename}")
		
		# Get SharePoint folder URL
		frappe.logger().info(f"[SharePoint Bundle] Getting folder URL for {target_folder_id}")
		folder_url = sharepoint.get_folder_url(target_folder_id)
		frappe.logger().info(f"[SharePoint Bundle] Folder URL: {folder_url}")
		
		if uploaded_count > 0:
			frappe.logger().info(f"[SharePoint Bundle] Upload completed: {uploaded_count} succeeded, {len(failed_files)} failed")
			return {
				'success': True,
				'uploaded_count': uploaded_count,
				'failed_count': len(failed_files),
				'folder_url': folder_url,
				'message': f'Successfully uploaded {uploaded_count} file(s) to SharePoint'
			}
		else:
			frappe.logger().error(f"[SharePoint Bundle] All uploads failed. Failed files: {failed_files}")
			return {
				'success': False,
				'message': 'Failed to upload files to SharePoint',
				'failed_files': failed_files
			}
			
	except Exception as e:
		frappe.logger().error(f"[SharePoint Bundle] Exception: {str(e)}")
		frappe.log_error("Document Bundle Upload Error", str(e))
		return {
			'success': False,
			'message': f'Error: {str(e)}'
		}


class SharePoint(object):
	def __init__(self, **kwargs):
		self.user = frappe.session.user
		self.doctype = kwargs.get("doctype")
		self.docname = kwargs.get("docname")
		self.filepath = kwargs.get("filepath")
		self.filedoc = kwargs.get("filedoc")
		self.settings = frappe.get_single(SETTINGS)
		
		# Validate required settings
		if not self.settings.sharepoint_drive_id:
			frappe.throw(_("SharePoint Drive ID not configured in SharePoint Settings"))
		
		self.drive_id = self.settings.sharepoint_drive_id
		self.root_folder = self.settings.root_folder_path or ""
		self.folder_structure = self.settings.folder_structure or "Module/DocType/Document"
		self.base_url = f'{self.settings.graph_api_url}/drives/{self.drive_id}'

	def get_sharepoint_folder_items(self, folder_id):
		'''
			Fetch folder contents from SharePoint Drive
		'''
		folder_items = []
		headers = get_request_header(self.settings)
		headers.update(ContentType)
		url = f'{self.base_url}/items/{folder_id}/children'

		response = make_request('GET', url, headers, None)
		if response.status_code == 200 or response.ok:
			for item in response.json()['value']:
				folder_items.append({"name": item["name"], "id": item["id"]})
		else:
			raise RuntimeError(f"SharePoint folder lookup failed ({response.status_code})")

		return folder_items

	def create_sharepoint_folder(self, parent_folder_id, folder_name):
		'''
			Create a folder in SharePoint Drive
		'''
		frappe.logger().info(f"[Create Folder] Creating '{folder_name}' in parent {parent_folder_id}")
		
		headers = get_request_header(self.settings)
		headers.update(ContentType)
		url = f'{self.base_url}/items/{parent_folder_id}/children'
		frappe.logger().info(f"[Create Folder] URL: {url}")
		
		body = {
			"name": f'{folder_name}',
			"folder": {},
			"@microsoft.graph.conflictBehavior": "rename"
		}

		response = make_request('POST', url, headers, body)
		frappe.logger().info(f"[Create Folder] Response status: {response.status_code if response else 'None'}")
		
		if not response.ok:
			frappe.logger().error(f"[Create Folder] Failed to create '{folder_name}': {response.text if response else 'No response'}")
			raise RuntimeError(f"SharePoint folder creation failed ({response.status_code})")
		else:
			folder_id = response.json()["id"]
			frappe.logger().info(f"[Create Folder] Successfully created '{folder_name}' with ID: {folder_id}")
			return folder_id

	def get_folder_id_by_name(self, parent_folder_id, folder_name):
		'''
			Get folder ID by name within a parent folder
		'''
		folder_id = None
		folder_items = self.get_sharepoint_folder_items(parent_folder_id)
		for item in folder_items:
			if folder_name == item['name']:
				folder_id = item['id']
				break
		return folder_id

	def get_or_create_folder(self, parent_folder_id, folder_name):
		'''
			Get existing folder or create new one
		'''
		frappe.logger().info(f"[Get/Create Folder] Looking for '{folder_name}' in parent {parent_folder_id}")
		folder_id = self.get_folder_id_by_name(parent_folder_id, folder_name)
		
		if not folder_id:
			frappe.logger().info(f"[Get/Create Folder] Folder '{folder_name}' not found, creating...")
			folder_id = self.create_sharepoint_folder(parent_folder_id, folder_name)
			frappe.logger().info(f"[Get/Create Folder] Created folder '{folder_name}' with ID: {folder_id}")
		else:
			frappe.logger().info(f"[Get/Create Folder] Found existing folder '{folder_name}' with ID: {folder_id}")
		
		return folder_id

	def get_root_folder_id(self):
		'''
			Get or create the root folder for uploads
		'''
		frappe.logger().info(f"[Get Root Folder] Starting - root_folder: '{self.root_folder}'")
		
		if self.root_folder:
			# Navigate to root folder path
			frappe.logger().info(f"[Get Root Folder] Fetching root folder from path: {self.root_folder}")
			headers = get_request_header(self.settings)
			url = f'{self.base_url}/root:/{self.root_folder}'
			frappe.logger().info(f"[Get Root Folder] URL: {url}")
			
			response = make_request('GET', url, headers, None)
			frappe.logger().info(f"[Get Root Folder] Response status: {response.status_code if response else 'None'}")
			
			if response.ok:
				folder_id = response.json()["id"]
				frappe.logger().info(f"[Get Root Folder] Found existing root folder ID: {folder_id}")
				return folder_id
			elif response.status_code == 404:
				# Create root folder only after a confirmed not-found response.
				frappe.logger().warning(f"[Get Root Folder] Root folder not found, creating: {self.root_folder}")
				folder_id = self.create_sharepoint_folder("root", self.root_folder)
				frappe.logger().info(f"[Get Root Folder] Created root folder ID: {folder_id}")
				return folder_id
			raise RuntimeError(f"SharePoint root folder lookup failed ({response.status_code})")
		else:
			# Use drive root
			frappe.logger().info(f"[Get Root Folder] No root folder specified, using drive root")
			headers = get_request_header(self.settings)
			url = f'{self.base_url}/root'
			response = make_request('GET', url, headers, None)
			if response.ok:
				folder_id = response.json()["id"]
				frappe.logger().info(f"[Get Root Folder] Drive root ID: {folder_id}")
				return folder_id
			raise RuntimeError(f"SharePoint drive lookup failed ({response.status_code})")

	def build_folder_structure(self):
		'''
			Build folder structure based on settings
			Returns the final folder ID where file should be uploaded
		'''
		frappe.logger().info(f"[Build Folders] Starting - structure: {self.folder_structure}")
		current_folder_id = self.get_root_folder_id()
		frappe.logger().info(f"[Build Folders] Root folder ID: {current_folder_id}")

		if self.folder_structure == "Flat":
			# No additional folders, upload directly to root
			frappe.logger().info(f"[Build Folders] Using flat structure - no additional folders")
			return current_folder_id
		
		# Module/DocType/Document structure
		doctype_module = frappe.db.get_value("DocType", {"name": self.doctype}, "module")
		frappe.logger().info(f"[Build Folders] DocType module: {doctype_module}")
		
		# Create Module folder
		if doctype_module:
			frappe.logger().info(f"[Build Folders] Creating/getting module folder: {doctype_module}")
			module_id = self.get_or_create_folder(current_folder_id, doctype_module)
			current_folder_id = module_id
			frappe.logger().info(f"[Build Folders] Module folder ID: {module_id}")
		
		# Create DocType folder
		frappe.logger().info(f"[Build Folders] Creating/getting doctype folder: {self.doctype}")
		doctype_id = self.get_or_create_folder(current_folder_id, self.doctype)
		current_folder_id = doctype_id
		frappe.logger().info(f"[Build Folders] DocType folder ID: {doctype_id}")
		
		# Create Document folder (docname)
		if self.docname:
			frappe.logger().info(f"[Build Folders] Creating/getting document folder: {self.docname}")
			document_id = self.get_or_create_folder(current_folder_id, self.docname)
			current_folder_id = document_id
			frappe.logger().info(f"[Build Folders] Document folder ID: {document_id}")
		
		frappe.logger().info(f"[Build Folders] Final target folder ID: {current_folder_id}")
		return current_folder_id

	def run_sharepoint_upload(self):
		"""Upload the current local file, letting errors fail the background job."""
		if not self.filepath or not self.filedoc:
			raise ValueError("A local file path and File document ID are required")

		# Open before remote operations, and close on every success/error path.
		with open(self.filepath, "rb") as file_content:
			target_folder_id = self.build_folder_structure()
			if not target_folder_id:
				raise RuntimeError("Could not determine the SharePoint target folder")

			headers = get_request_header(self.settings)
			headers.update({"Content-Type": "application/octet-stream"})
			file_name = quote(os.path.basename(self.filepath), safe="")
			url = f"{self.base_url}/items/{target_folder_id}:/{file_name}:/content"
			response = make_request("PUT", url, headers, file_content)
			item = self.confirm_upload(response, os.fstat(file_content.fileno()).st_size)

		# Keep local copies: Communication HTML and other File records can still
		# reference the original URL, even when this File's link is replaced.
		values = {"uploaded_to_sharepoint": 1}
		if self.settings.replace_file_link:
			if not item.get("webUrl"):
				raise RuntimeError("SharePoint upload returned no replacement file URL")
			values["file_url"] = item["webUrl"]
		frappe.db.set_value("File", self.filedoc, values)
		return {"status": "uploaded", "filedoc": self.filedoc, "item_id": item["id"]}

	@staticmethod
	def confirm_upload(response, expected_size):
		"""Require a completed Graph upload and matching size before marking a copy."""
		if response is None or response.status_code not in (200, 201):
			status = response.status_code if response is not None else "no response"
			raise RuntimeError(f"SharePoint upload failed ({status})")
		item = response.json()
		if not item.get("id") or item.get("size") != expected_size:
			raise RuntimeError("SharePoint upload did not confirm the file ID and expected size")
		return item

	def upload_file_to_folder(self, target_folder_id, filepath, filename):
		'''
			Upload a single file to a specific SharePoint folder
			
			Args:
				target_folder_id: SharePoint folder ID
				filepath: Local file path
				filename: Name for the file in SharePoint
				
			Returns:
				bool: True if upload successful, False otherwise
		'''
		try:
			frappe.logger().info(f"[Upload File] Starting upload: {filename}")
			frappe.logger().info(f"[Upload File] Source path: {filepath}")
			frappe.logger().info(f"[Upload File] Target folder ID: {target_folder_id}")
			
			# Read file content
			frappe.logger().info(f"[Upload File] Reading file content from disk")
			with open(filepath, 'rb') as f:
				file_content = f.read()
			
			frappe.logger().info(f"[Upload File] File size: {len(file_content)} bytes")
			
			if not file_content:
				frappe.logger().error(f"[Upload File] File {filename} is empty")
				frappe.log_error("SharePoint Upload Error", f"File {filename} is empty")
				return False
			
			# Upload file with replace behavior
			frappe.logger().info(f"[Upload File] Getting authentication headers")
			headers = get_request_header(self.settings)
			headers.update({"Content-Type": "application/octet-stream"})
			
			url = f'{self.base_url}/items/{target_folder_id}:/{quote(filename, safe="")}:/content'
			frappe.logger().info(f"[Upload File] Upload URL: {url}")
			
			frappe.logger().info(f"[Upload File] Making PUT request to SharePoint")
			response = make_request('PUT', url, headers, file_content)
			
			frappe.logger().info(f"[Upload File] Response status: {response.status_code if response else 'None'}")
			
			if not response.ok:
				frappe.logger().error(f"[Upload File] Upload failed for {filename}")
				frappe.logger().error(f"[Upload File] Response: {response.text if response else 'No response'}")
				frappe.log_error("SharePoint File Upload Error", f"File: {filename}, Status: {response.status_code}, Error: {response.text}")
				return False
			
			self.confirm_upload(response, len(file_content))
			frappe.logger().info(f"[Upload File] Successfully uploaded {filename}")
			return True
			
		except Exception as e:
			frappe.logger().error(f"[Upload File] Exception while uploading {filename}: {str(e)}")
			frappe.log_error("File Upload Error", f"File: {filename}, Error: {str(e)}")
			return False
	
	def get_folder_url(self, folder_id):
		'''
			Get web URL for a SharePoint folder
			
			Args:
				folder_id: SharePoint folder ID
				
			Returns:
				str: Web URL to the folder or None
		'''
		try:
			frappe.logger().info(f"[Get Folder URL] Fetching URL for folder ID: {folder_id}")
			
			headers = get_request_header(self.settings)
			url = f'{self.base_url}/items/{folder_id}'
			frappe.logger().info(f"[Get Folder URL] Request URL: {url}")
			
			response = make_request('GET', url, headers, None)
			frappe.logger().info(f"[Get Folder URL] Response status: {response.status_code if response else 'None'}")
			
			if response.ok:
				web_url = response.json().get('webUrl')
				frappe.logger().info(f"[Get Folder URL] Retrieved web URL: {web_url}")
				return web_url
			else:
				frappe.logger().error(f"[Get Folder URL] Failed to get URL: {response.text if response else 'No response'}")
			
			return None
			
		except Exception as e:
			frappe.logger().error(f"[Get Folder URL] Exception: {str(e)}")
			frappe.log_error("Get Folder URL Error", str(e))
			return None
