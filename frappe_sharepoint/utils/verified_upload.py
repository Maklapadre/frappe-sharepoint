"""Bounded Graph transfers for backups; never log preauthenticated upload URLs."""

import os
import re
import time
from datetime import datetime, timezone
from email.utils import parsedate_to_datetime
from urllib.parse import quote

import requests

from frappe_sharepoint.utils import get_request_header

CHUNK_SIZE = 10 * 1024 * 1024  # A multiple of Graph's required 320 KiB.
MAX_ATTEMPTS = 4
MAX_RETRY_DELAY = 60
REQUEST_TIMEOUT = (10, 120)
TRANSIENT_STATUS = {408, 429, 500, 502, 503, 504}


class SharePointTransferError(RuntimeError):
	"""A sanitized error suitable for job logs and operator status."""


class SharePointRetryLater(SharePointTransferError):
	"""Server backoff exceeds this job's retry budget."""


def sanitize_message(value):
	"""Remove URLs, credential assignments and control characters from diagnostics."""
	text = re.sub(r"https?://\S+", "[redacted URL]", str(value or ""))
	text = re.sub(r"(?i)bearer\s+\S+", "Bearer [redacted]", text)
	text = re.sub(
		r'''(?i)((?:access_token|refresh_token|client_secret|password|token|secret)["']?\s*[=:]\s*["']?)[^"'\s,;]+''',
		r"\1[redacted]", text,
	)
	return " ".join(text.split())[:1000]


def response_json(response):
	"""Return an object without exposing a non-JSON response body."""
	try:
		value = response.json()
		return value if isinstance(value, dict) else {}
	except (ValueError, AttributeError):
		return {}


def response_error(response, context):
	"""Preserve false-valued HTTP failures, with only selected Graph diagnostics."""
	if response is None:
		return SharePointTransferError(f"{context}: no HTTP response")
	payload = response_json(response)
	error = payload.get("error") or {}
	if not isinstance(error, dict):
		error = {}
	inner = error.get("innerError") or error.get("innererror") or {}
	request_id = response.headers.get("request-id") or (inner.get("request-id") if isinstance(inner, dict) else "")
	detail = sanitize_message(error.get("message") or "No structured Graph error message")
	return SharePointTransferError(
		f"{context}: HTTP {response.status_code}; code={sanitize_message(error.get('code'))}; "
		f"request_id={sanitize_message(request_id)}; {detail}"
	)


def retry_delay(response, attempt):
	"""Honor Retry-After; refuse excessive waits instead of retrying too early."""
	value = response.headers.get("Retry-After") if response is not None else None
	try:
		delay = float(value) if value else 2 ** attempt
	except ValueError:
		try:
			delay = (parsedate_to_datetime(value) - datetime.now(timezone.utc)).total_seconds()
		except (TypeError, ValueError):
			delay = 2 ** attempt
	if delay > MAX_RETRY_DELAY:
		raise SharePointRetryLater(f"SharePoint requested Retry-After={delay:g}s; retry this job later")
	return max(0, delay)


def request(method, url, *, context, headers=None, data=None, json=None):
	"""Retry transient responses with a fixed budget and rewind streamed bodies."""
	response = None
	for attempt in range(MAX_ATTEMPTS):
		if hasattr(data, "seek"):
			data.seek(0)
		try:
			response = requests.request(
				method, url, headers=headers, data=data, json=json, timeout=REQUEST_TIMEOUT,
			)
		except (requests.Timeout, requests.ConnectionError):
			response = None
		except requests.RequestException:
			raise SharePointTransferError(f"{context}: HTTP transport failed") from None
		if response is not None and response.status_code not in TRANSIENT_STATUS:
			return response
		if attempt < MAX_ATTEMPTS - 1:
			try:
				time.sleep(retry_delay(response, attempt))
			except SharePointRetryLater as exc:
				raise SharePointRetryLater(f"{context}: {exc}") from None
	raise response_error(response, context) from None


def graph_request(sharepoint, method, path, *, context, **kwargs):
	"""Send authenticated Graph requests without using the connector's URL logger."""
	headers = get_request_header(sharepoint.settings)
	return request(method, f"{sharepoint.base_url}{path}", context=context, headers=headers, **kwargs)


def confirm_item(response, size, context):
	"""Require completed upload metadata with the expected byte count."""
	if response is None or response.status_code not in (200, 201):
		raise response_error(response, context)
	item = response_json(response)
	if not item.get("id") or item.get("size") != size:
		raise SharePointTransferError(f"{context}: upload did not confirm item ID and size={size}")
	return item


def upload_file(sharepoint, folder_id, path, filename=None):
	"""Stream small files or resume bounded fragments, then verify final metadata."""
	filename = filename or os.path.basename(path)
	with open(path, "rb") as source:
		size = os.fstat(source.fileno()).st_size
		context = f"Upload {sanitize_message(filename)} ({size} bytes)"
		if size <= 0:
			raise SharePointTransferError(f"{context}: file is empty")
		item_path = f"/items/{folder_id}:/{quote(filename, safe='')}"
		if size <= CHUNK_SIZE:
			headers = {**get_request_header(sharepoint.settings), "Content-Type": "application/octet-stream"}
			response = request("PUT", f"{sharepoint.base_url}{item_path}:/content",
				context=context, headers=headers, data=source)
			return confirm_item(response, size, context)

		response = graph_request(sharepoint, "POST", f"{item_path}:/createUploadSession", context=context,
			json={"item": {"@microsoft.graph.conflictBehavior": "replace", "name": filename}})
		if response.status_code not in (200, 201):
			raise response_error(response, context)
		upload_url = response_json(response).get("uploadUrl")
		if not isinstance(upload_url, str) or not upload_url.startswith("https://"):
			raise SharePointTransferError(f"{context}: missing secure upload session URL")

		offset = 0
		recoveries = 0
		last_failure = None
		while offset < size:
			source.seek(offset)
			chunk = source.read(min(CHUNK_SIZE, size - offset))
			if not chunk:
				raise SharePointTransferError(f"{context}: local file changed during upload")
			headers = {"Content-Length": str(len(chunk)),
				"Content-Range": f"bytes {offset}-{offset + len(chunk) - 1}/{size}"}
			# Upload URLs already contain credentials. Never attach a bearer token.
			try:
				response = request("PUT", upload_url, context=context, headers=headers, data=chunk)
			except SharePointRetryLater:
				raise
			except SharePointTransferError as exc:
				last_failure = exc
				response = None
			if response is not None and response.status_code in (200, 201):
				return confirm_item(response, size, context)
			if response is None or response.status_code in (404, 416):
				recoveries += 1
				if recoveries >= MAX_ATTEMPTS:
					raise last_failure or SharePointTransferError(f"{context}: upload-session recovery budget exhausted")
				response = request("GET", upload_url, context=context)
				if response.status_code == 404:
					# The final PUT may have committed before its response was lost.
					response = graph_request(sharepoint, "GET", item_path, context=context)
					return confirm_item(response, size, context)
				if response.status_code != 200:
					raise response_error(response, context)
			elif response.status_code != 202:
				raise response_error(response, context)
			ranges = response_json(response).get("nextExpectedRanges") or []
			try:
				next_offset = int(ranges[0].split("-", 1)[0])
			except (IndexError, TypeError, ValueError, AttributeError):
				raise SharePointTransferError(f"{context}: invalid upload-session ranges") from None
			if not 0 <= next_offset < size:
				raise SharePointTransferError(f"{context}: invalid upload-session offset")
			if next_offset <= offset:
				recoveries += 1
				if recoveries >= MAX_ATTEMPTS:
					raise last_failure or SharePointTransferError(f"{context}: upload session made no progress")
			offset = next_offset


def list_children(sharepoint, folder_id):
	"""Read every child page; preserve file and folder metadata for verification."""
	url = f"{sharepoint.base_url}/items/{folder_id}/children"
	items = []
	seen = set()
	while url:
		if url in seen or not url.startswith(f"{sharepoint.base_url}/"):
			raise SharePointTransferError("Invalid SharePoint directory pagination")
		seen.add(url)
		response = request("GET", url, context="List backup folder", headers=get_request_header(sharepoint.settings))
		if response.status_code != 200:
			raise response_error(response, "List backup folder")
		payload = response_json(response)
		if not isinstance(payload.get("value"), list):
			raise SharePointTransferError("SharePoint directory response has no item list")
		items.extend(payload["value"])
		url = payload.get("@odata.nextLink")
	return items
