<div align="center">
    <h1>Frappe SharePoint Integration</h1>
</div>

A universal SharePoint file synchronization solution for Frappe/ERPNext. This app automatically uploads files from your Frappe system to your SharePoint site, with flexible folder structure options and OAuth2 authentication.

## Features

- **Universal SharePoint Integration**: Connect to any SharePoint site using your own Azure AD tenant
- **Automatic File Sync**: Automatically upload files to SharePoint when they're attached to documents
- **Flexible Folder Structure**: Choose between hierarchical (Module/DocType/Document) or flat folder organization
- **Simple Authentication**: Direct credential configuration with Azure AD App Registration
- **Optional File Replacement**: Keep files on SharePoint only or maintain local copies
- **Supports Frappe v13 and v14**

## Why Use SharePoint Integration?

**Centralized File Management:** Keep all your files in SharePoint for better organization and easier sharing with external stakeholders.

**Enhanced Security:** Leverage SharePoint's enterprise-grade security features and compliance tools.

**Better Collaboration:** Share files easily with team members using SharePoint's built-in sharing capabilities.

**Reduced Storage Costs:** Optionally remove local file copies after uploading to SharePoint to save storage space.

**Backup and Recovery:** Benefit from SharePoint's built-in versioning and backup capabilities.

---

## Installation

### Self Hosting:

```bash
# Get the app
bench get-app https://github.com/yourusername/frappe-sharepoint.git

# Install on your site
bench --site [your.site.name] install-app frappe_sharepoint

# Run migrations
bench --site [your.site.name] migrate

# Restart
bench restart
```

---

## Setup Instructions

### 1. Azure AD App Registration

1. Go to [Azure Portal](https://portal.azure.com) → Azure Active Directory → App registrations
2. Click "New registration"
3. Configure your app:
   - **Name**: Frappe SharePoint Sync (or any name you prefer)
   - **Supported account types**: Accounts in this organizational directory only
   - **Redirect URI**: Not required (leave blank)

<img src="./app_registration.png" height="480">

4. After creation, note down:
   - **Application (client) ID**
   - **Directory (tenant) ID**

5. Go to "Certificates & secrets" → Create a new client secret
   - Note down the **Value** (you won't be able to see it again)

6. Go to "API permissions" → Add the following Microsoft Graph **Application** permissions:
   - `Files.ReadWrite.All`
   - `Sites.ReadWrite.All`

7. Click "Grant admin consent" for your organization

### 2. Configure SharePoint Settings

1. Go to **SharePoint Settings** in ERPNext
2. Fill in the following fields:

   **Azure AD Credentials:**
   - **Tenant ID**: Your Azure AD tenant ID
   - **Client ID**: Your app registration client ID
   - **Client Secret**: Your app registration client secret
   - Click **Test Connection** to verify your credentials

   **SharePoint Configuration:**
   - **Graph API URL**: `https://graph.microsoft.com/v1.0` (default)
   - **Enable File Sync**: Check to enable automatic file upload
   - **SharePoint Site URL**: Full URL of your SharePoint site (e.g., `https://yourtenant.sharepoint.com/sites/YourSite`)
   - Click **Fetch SharePoint Details** button to automatically retrieve Site ID and Drive ID
   - **Root Folder Path**: (Optional) Specify a root folder within the drive (e.g., `/Frappe Files`)

   **File Handling:**
   - **Replace File Link**: Check to replace local files with SharePoint links (saves local storage)
   - **Folder Structure**: Choose between:
     - `Module/DocType/Document`: Creates hierarchical folders
     - `Flat`: Uploads all files to root folder

<img src="./m365_settings.png" height="580">

3. Save the settings

---

## Usage

Once configured, the app will automatically:

1. Upload any new files attached to Frappe documents to SharePoint
2. Create the folder structure based on your settings
3. Mark files as "Uploaded to SharePoint"
4. Optionally replace the local file with a SharePoint link

### Folder Structure Examples

**Module/DocType/Document:**
```
SharePoint Drive
└── [Root Folder Path]
    └── [Module Name]
        └── [DocType Name]
            └── [Document Name]
                └── [File]
```

**Flat:**
```
SharePoint Drive
└── [Root Folder Path]
    └── [File]
```

---

## Troubleshooting

### Files not uploading?

1. Check that "Enable File Sync" is enabled in SharePoint Settings
2. Click "Test Connection" to verify your Azure AD credentials
3. Check Error Log in Frappe for specific error messages
4. Ensure the SharePoint Site ID and Drive ID are correctly fetched

### Permission errors?

1. Verify all required Microsoft Graph **Application** permissions are granted
2. Ensure admin consent was granted in Azure AD
3. Check that your Azure AD app has access to the SharePoint site

### Can't fetch SharePoint details?

1. Verify the SharePoint Site URL is correct
2. Click "Test Connection" to verify your credentials
3. Ensure your Azure AD app has proper permissions

---

## Dependencies

- [Frappe Framework](https://github.com/frappe/frappe) v13 or v14
- Microsoft 365 subscription with SharePoint Online
- Azure AD tenant with app registration permissions

---

## Bug Reports

Please create an issue on [GitHub Issues](https://github.com/yourusername/frappe-sharepoint/issues/new)

---

## License

MIT

## Automatic upload reliability and deployment

Automatic uploads queue only the File ID after the insert transaction commits.
The worker reloads the File, resolves its stored URL with `get_full_path()`, and
opens the current local file with a context manager. Old queued jobs containing
`filepath`, `doctype`, or `docname` remain supported, but those saved values are
ignored in favor of the current File record. Already-uploaded Files are skipped.
Remote URL attachments are not downloaded or queued as local files; a queued File
that later becomes remote fails with an explicit message.

The upload flag is set only after Graph confirms a completed upload with an item
ID and matching byte size. Missing files, authentication errors, network failures,
and invalid upload confirmations fail the background job. A timeout can occur
after Graph received the file: inspect SharePoint before retrying such a job.

Local copies are retained, including when **Replace File Link** updates the File
record to a SharePoint URL. Existing Communication content or other File records
may still use the original local URL. Automatic uploads no longer delete those
local copies. This intentionally retains local storage to preserve attachment
accessibility. For copies that must retain the original ERPNext attachment URLs,
keep **Replace File Link** disabled. Enabling it still changes the File URL and
requires SharePoint permissions; retaining the bytes alone does not preserve old
private URLs, whose access checks depend on a matching File record.

### Regression tests

From `apps/frappe_sharepoint`, using the bench Python environment:

```sh
../../env/bin/python -m unittest discover -s frappe_sharepoint/tests -v
```

To also exercise real File/Communication inserts, commit/rollback timing, and
attachment reads, use a **development** site with this app installed:

```sh
SHAREPOINT_TEST_SITE=your-development-site \
  ../../env/bin/python -m unittest discover -s frappe_sharepoint/tests -v
```

The integration tests replace Graph and RQ at their boundaries. They create and
clean up temporary records; they do not send live uploads. Live download/hash
verification must therefore also be performed on the deployment target.

### Frappe Cloud rollout and selective reconciliation

1. Configure the bench group's app source as `Maklapadre/frappe-sharepoint`,
   branch `andre_main`. Record the exact tested commit SHA. Follow the
   [Frappe Cloud app/site update procedure](https://docs.frappe.io/cloud/sites/how-to-update-an-app-site-on-a-private-bench)
   to deploy that app revision and update the intended site. Verify the deployed
   SHA and that the site's workers run the new revision before retrying uploads.
2. Confirm **Replace File Link** is disabled when preserving ERPNext attachment
   URLs, then run a small canary: a public attachment, a private attachment with a generated
   filename suffix, and an attachment to a Communication. Download each uploaded
   copy from SharePoint and compare its SHA-256 with the local original. Verify
   attachment access in ERPNext under the intended user's permissions.
3. Build a bounded manifest of **specific File IDs** from the affected failure
   window and Error Logs/failed jobs. Older swallowed exceptions may have left
   jobs appearing successful. Include each stored URL, current attachment target,
   local availability, upload flag, local SHA-256, expected SharePoint folder and
   stored basename, remote item ID/hash if present, and proposed action. A zero
   upload flag alone is not proof that the remote copy is absent.
4. Inspect the expected remote location before changing anything. If an identical
   copy exists, reconcile its flag without uploading it again. If contents differ,
   or there are multiple candidate copies, leave that File for investigation; do
   not overwrite it. Missing local files and remote-URL-only records require their
   own recovery. Do not bulk-reset flags or retry every attachment.
5. Queue only manifest entries confirmed to have a readable local original and
   no remote copy, initially in a small batch. Use the existing job with the File
   ID and `enqueue_after_commit=True` (as in `file_controller.file_upload`). For
   old failed jobs, the worker ignores stale path arguments. Recheck ambiguous
   timeouts against SharePoint before any subsequent retry.
6. After each batch, compare downloaded content hashes, confirm upload flags,
   inspect failed jobs, and open the ERPNext/Communication attachments again.
   Stop on unexpected differences and retain the manifest as the reconciliation
   record. Do not use a site migration to re-upload attachments.

Graph's upload response contract is documented in
[Upload or replace driveItem content](https://learn.microsoft.com/en-us/graph/api/driveitem-put-content?view=graph-rest-1.0).
