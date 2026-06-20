🚀 Google Drive Admin

A robust, "Low-Code" Google Sheets & Apps Script solution to audit, manage, and bulk-update Google Drive permissions recursively.

Unlike basic scripts, this tool fully supports Shared Drives, bypasses the Google Apps Script 6-minute execution limit, and uses native Google Sheets formulas to compare your theoretical access matrix with the actual Drive reality.

🎯 The Problem

Managing permissions across hundreds of folders and files in Google Drive is a nightmare. Native UI is too slow for bulk actions, and identifying who has access to what (especially inherited rights in Shared Drives) is nearly impossible without third-party enterprise tools.

💡 The Solution

This tool uses Google Sheets as a database and Apps Script as the execution engine:

Audit: Recursively scans a Drive folder (or Shared Drive) and maps all files and permissions into relational Sheets.

Diffing (Low-Code): Uses a Template tab with Google Sheets formulas to cross-reference the audit results against your theoretical Matrix of roles.

Bulk Update: Applies the calculated changes (Add/Modify/Delete permissions, block downloads, restrict sharing) via the Drive API.

✨ Key Features

Shared Drive (v3) Compatible: Accurately detects and handles inherited permissions in Shared Drives.

Timeout-Proof: Uses a Queue system (Breadth-First Search) and time-based triggers to automatically pause and resume large audits, entirely bypassing the 6-minute Apps Script limit.

Rate-Limit Safe: Includes targeted delays (Utilities.sleep) to prevent HTTP 429 API errors.

Data Privacy First: 100% runs within your Google Workspace. No external APIs, no data exfiltration.

🛠️ Installation & Setup

Option A: The 1-Click Install (Recommended)

Copy the Template at this address :

https://docs.google.com/spreadsheets/d/1kNNWGqMxn4839VFATvZ46jpX5wKLgrqQkRJqNxnL4Nc/copy

The Apps Script is already included in the copy.

Go to Extensions > Apps Script, open the Services left menu, click +, add the Drive API and ensure Version V3 is selected.

Option B: Manual Install

Create a new Google Sheet.

Go to Extensions > Apps Script and paste the Code.gs from this repository.

Add the Drive API (v3) in the Services menu.

Save and refresh your spreadsheet. A custom menu 🚀 Admin Drive Pro (V3) will appear.

📖 How to Use (The Workflow)

Set the Target: In the script (CONFIG.STARTING_FOLDER_ID), paste the ID of the folder you want to audit (or leave it as "root").

Initialize: Click Admin Drive Pro > 1️⃣ Initialize Sheets.

Audit: Click 2️⃣ Start / Resume Audit. Wait for the script to recursively fetch all data.

Compare: Use the Template tab (with your own formulas like VLOOKUP) to compare the 🔑 Permissions tab against your desired rules.

Prepare Actions: Copy the anomalies (e.g., missing users) and paste them into the 🔑 Permissions tab. Set the Required_Action column to TO_ADD, TO_MODIFY, or TO_DELETE, and specify the New_Role.

Execute: Click 4️⃣ UPDATE: User Permissions to apply the changes directly to Google Drive.

⚠️ Security & Limitations

You must have "Owner" or "Manager" rights on the files/drives you are attempting to modify.

Inherited rights (from a parent folder) cannot be modified at the child file level. The script will safely skip these and alert you.

🤝 Contributing

Feel free to fork this project, submit pull requests, or open an issue if you find a bug!
