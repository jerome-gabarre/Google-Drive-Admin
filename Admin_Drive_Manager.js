/**
 * @fileoverview Centralized Audit and Permission Management System for Google Drive.
 * ⚠️ VERSION V3: Uses Drive API V3 for native inheritance detection and Shared Drives support.
 * * * MANDATORY PREREQUISITES:
 * 1. Go to "Services" (left menu in Apps Script) > "+" > Add "Drive API".
 * 2. IMPORTANT: In the "Version" dropdown, you MUST select "v3".
 */

const CONFIG = {
  SHEET_FILES: "📁 Files",
  SHEET_PERMS: "🔑 Permissions",
  STARTING_FOLDER_ID: "PUT_YOUR_FOLDER_ID_HERE", // <-- REPLACE THIS WITH YOUR FOLDER ID OR "root"
  MAX_EXEC_TIME_MS: 250000,
  
  DRIVE_FIELDS: "nextPageToken, files(id, name, mimeType, webViewLink, parents, owners, writersCanShare, copyRequiresWriterPermission, inheritedPermissionsDisabled, permissions(id, type, emailAddress, role, permissionDetails))",
  
  STATUS_COL_HEADER: "Update_Status",
  ACTION_COL_HEADER: "Required_Action" 
};

/**
 * Creates the custom menu when the spreadsheet is opened.
 * @return {void}
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🚀 Admin Drive Pro (V3)')
    .addItem('1️⃣ Initialize Sheets', 'setupRelationalSheets')
    .addSeparator()
    .addItem('2️⃣ Start / Resume Audit', 'startAudit')
    .addItem('🧹 Clear Audit Cache', 'clearAuditCache')
    .addSeparator()
    .addItem('3️⃣ UPDATE: File Settings', 'startBulkUpdateFiles')
    .addItem('4️⃣ UPDATE: User Permissions', 'startBulkUpdatePermissions')
    .addToUi();
}

/** =========================================================================
 * MODULE 1: SHEETS INITIALIZATION
 * ========================================================================= */

/**
 * Sets up the required relational tabs (Files and Permissions) with headers and data validation.
 * @return {void}
 */
function setupRelationalSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. Setup Files tab
  let sheetFiles = ss.getSheetByName(CONFIG.SHEET_FILES);
  if (!sheetFiles) sheetFiles = ss.insertSheet(CONFIG.SHEET_FILES);
  
  const headersFiles = [
    "File_ID", "Name", "MIME_Type", "Path", "Link", "Parent_ID", 
    "Editors_Can_Share (Current)", "Viewers_Can_Download (Current)", "Inherited_Access_Limited (Current)",
    CONFIG.ACTION_COL_HEADER, 
    "New_Editors_Can_Share (Yes/No)", "New_Viewers_Can_Download (Yes/No)", "New_Limit_Access (Yes/No)",
    CONFIG.STATUS_COL_HEADER
  ];
  sheetFiles.getRange(1, 1, 1, headersFiles.length).setValues([headersFiles]).setFontWeight("bold").setBackground("#e0f7fa");
  
  const ruleActionFiles = SpreadsheetApp.newDataValidation().requireValueInList(["TO_MODIFY", "-"], true).build();
  const ruleYesNo = SpreadsheetApp.newDataValidation().requireValueInList(["Yes", "No", "-"], true).build();
  
  sheetFiles.getRange(2, 10, 10000, 1).setDataValidation(ruleActionFiles);
  sheetFiles.getRange(2, 11, 10000, 3).setDataValidation(ruleYesNo);
  sheetFiles.setFrozenRows(1);

  // 2. Setup Permissions tab
  let sheetPerms = ss.getSheetByName(CONFIG.SHEET_PERMS);
  if (!sheetPerms) sheetPerms = ss.insertSheet(CONFIG.SHEET_PERMS);
  
  const headersPerms = [
    "File_ID", "File_Name", "Path", "Permission_ID", "Entity_Type", 
    "Email", "Current_Role", "Inherited_Rights", 
    CONFIG.ACTION_COL_HEADER, "New_Role", CONFIG.STATUS_COL_HEADER
  ];
  sheetPerms.getRange(1, 1, 1, headersPerms.length).setValues([headersPerms]).setFontWeight("bold").setBackground("#fff3e0");
  
  const ruleActionPerms = SpreadsheetApp.newDataValidation().requireValueInList(["TO_ADD", "TO_MODIFY", "TO_DELETE", "-"], true).build();
  sheetPerms.getRange(2, 9, 10000, 1).setDataValidation(ruleActionPerms);
  sheetPerms.setFrozenRows(1);
  
  SpreadsheetApp.getActive().toast("Relational sheets ready for audit and modification.", "Success");
}

/** =========================================================================
 * MODULE 2: AUDIT ENGINE (V3 API READING)
 * ========================================================================= */

/**
 * Starts or resumes the Google Drive permission audit.
 * Includes data lifecycle management to prevent row duplication.
 * @return {void}
 */
function startAudit() {
  const props = PropertiesService.getScriptProperties();
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  deleteTriggers_('processAuditQueue');

  let queueStr = props.getProperty('DRIVE_QUEUE');
  let lastFolderId = props.getProperty('LAST_STARTING_FOLDER');
  
  // If a queue exists but the target folder changed, we reset the queue.
  if (queueStr && lastFolderId !== CONFIG.STARTING_FOLDER_ID) {
    queueStr = null; 
  }

  let queue;
  
  // IF NO QUEUE EXISTS -> IT'S A NEW AUDIT
  if (!queueStr) {
    
    // --- 🛠️ BUG FIX: Check for existing data and clear it ---
    let sheetFiles = ss.getSheetByName(CONFIG.SHEET_FILES);
    if (sheetFiles && sheetFiles.getLastRow() > 1) {
      let response = ui.alert(
        "⚠️ Restart Audit", 
        "Existing data found. Starting a new audit will clear the current sheets. Do you want to continue?", 
        ui.ButtonSet.YES_NO
      );
      if (response !== ui.Button.YES) return; // Abort if user says no
    }

    clearSheetData_(); // Wipe the slate clean before starting
    // --------------------------------------------------------

    let rootName = "My Drive";
    if (CONFIG.STARTING_FOLDER_ID !== "root") {
      try {
        let rootMeta = Drive.Files.get(CONFIG.STARTING_FOLDER_ID, {supportsAllDrives: true, fields: "name"});
        rootName = rootMeta.name;
      } catch (e) { rootName = "Root Folder"; }
    }
    
    queue = [{id: CONFIG.STARTING_FOLDER_ID, path: rootName}];
    props.setProperty('DRIVE_QUEUE', JSON.stringify(queue));
    props.setProperty('LAST_STARTING_FOLDER', CONFIG.STARTING_FOLDER_ID);
    
    ss.toast("Starting V3 Audit...", "Info");
    
  } else {
    // IF QUEUE EXISTS -> RESUME AUDIT
    queue = JSON.parse(queueStr);
    ss.toast("Resuming audit in progress...", "Info");
  }

  if (queue.length === 0) return ui.alert("Audit is already complete! Clear the cache to restart.");
  
  processAuditQueue();
}

/**
 * Iteratively processes the audit queue. Handles Shared Drive fallbacks.
 * @return {void}
 */
function processAuditQueue() {
  const startTime = Date.now();
  const props = PropertiesService.getScriptProperties();
  let queue = JSON.parse(props.getProperty('DRIVE_QUEUE') || '[]');
  let batchFiles = [], batchPerms = [];

  const requestOptions = { pageSize: 1000, supportsAllDrives: true, includeItemsFromAllDrives: true, fields: CONFIG.DRIVE_FIELDS };

  while (queue.length > 0) {
    if (Date.now() - startTime > CONFIG.MAX_EXEC_TIME_MS) {
      props.setProperty('DRIVE_QUEUE', JSON.stringify(queue));
      flushDataToSheets_(batchFiles, batchPerms);
      scheduleNextRun_('processAuditQueue');
      return;
    }

    let currentItem = queue.shift();
    let currentFolderId = typeof currentItem === 'string' ? currentItem : currentItem.id;
    let currentFolderPath = typeof currentItem === 'string' ? "Unknown Path" : currentItem.path;
    let pageToken = null;
    
    do {
      try {
        requestOptions.q = currentFolderId === "root" ? "trashed = false" : `"${currentFolderId}" in parents and trashed = false`;
        requestOptions.pageToken = pageToken;
        let response = Drive.Files.list(requestOptions);
        let files = response.files || [];
        
        files.forEach(file => {
          let parentId = (file.parents && file.parents.length > 0) ? file.parents[0] : "root";
          
          let canEditorsShare = file.writersCanShare !== false ? "Yes" : "No";
          let canViewersDownload = file.copyRequiresWriterPermission === true ? "No" : "Yes";
          let isAccessLimited = file.inheritedPermissionsDisabled === true ? "Yes" : "No";
          
          if (file.mimeType === "application/vnd.google-apps.folder") {
            canViewersDownload = "N/A (Folder)"; 
          } else {
            isAccessLimited = "N/A (File)"; 
          }
          
          batchFiles.push([
            "'" + file.id, file.name, file.mimeType, currentFolderPath, file.webViewLink, parentId,
            canEditorsShare, canViewersDownload, isAccessLimited, 
            "-", "", "", "", ""
          ]);

          let itemPermissions = file.permissions;

          // Shared Drive Fallback: Fetch explicitly if skipped by list()
          if (!itemPermissions) {
            try {
              let permResponse = Drive.Permissions.list(file.id, {
                supportsAllDrives: true,
                fields: "permissions(id, type, emailAddress, role, permissionDetails)"
              });
              itemPermissions = permResponse.permissions || [];
              Utilities.sleep(100); // Prevent 429 Rate Limit Errors
            } catch (e) {
              console.warn(`Info: Cannot extract permissions for ${file.name} (${file.id}). Cause: ${e.message}`);
              itemPermissions = [];
            }
          }

          if (itemPermissions && itemPermissions.length > 0) {
            itemPermissions.forEach(perm => {
              let isInherited = "No";
              if (perm.permissionDetails && perm.permissionDetails.length > 0 && perm.permissionDetails[0].inherited) isInherited = "Yes (Inherited)";

              batchPerms.push([
                "'" + file.id, file.name, currentFolderPath, "'" + perm.id, perm.type,
                perm.emailAddress || (perm.type === 'anyone' ? 'Public Link' : 'Domain'),
                perm.role, isInherited, "-", "", ""   
              ]);
            });
          }

          if (file.mimeType === "application/vnd.google-apps.folder") {
            queue.push({ id: file.id, path: currentFolderPath + " / " + file.name });
          }
        });
        
        pageToken = response.nextPageToken;
      } catch (e) {
        console.error(`Error accessing folder ${currentFolderId} : ${e.message}`);
        pageToken = null; 
      }
    } while (pageToken);
    
    props.setProperty('DRIVE_QUEUE', JSON.stringify(queue));
  }

  flushDataToSheets_(batchFiles, batchPerms);
  props.deleteProperty('DRIVE_QUEUE');
  SpreadsheetApp.getActive().toast("🎉 V3 Audit completed successfully!", "Done");
}

/** =========================================================================
 * MODULE 3A: UPDATE ENGINE - FILE SETTINGS (Drive.Files)
 * ========================================================================= */

/**
 * Applies requested metadata updates (Share/Download rules) to files.
 * @param {Object} e - Event object if triggered by time-based trigger.
 * @return {void}
 */
function startBulkUpdateFiles(e) {
  const isTrigger = (e && e.triggerUid) ? true : false;
  if (isTrigger) ScriptApp.getProjectTriggers().forEach(t => { if (t.getUniqueId() === e.triggerUid) ScriptApp.deleteTrigger(t); });

  let ui;
  if (!isTrigger) {
    ui = SpreadsheetApp.getUi();
    const response = ui.alert('⚠️ UPDATE Files', 'Apply parameter modifications (Sharing/Download/Inheritance)?', ui.ButtonSet.YES_NO);
    if (response !== ui.Button.YES) return;
  }

  const sheet = SpreadsheetApp.getActive().getSheetByName(CONFIG.SHEET_FILES);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  
  const idxFileId = headers.indexOf("File_ID");
  const idxMimeType = headers.indexOf("MIME_Type"); 
  const idxAction = headers.indexOf(CONFIG.ACTION_COL_HEADER);
  const idxNewShare = headers.indexOf("New_Editors_Can_Share (Yes/No)");
  const idxNewDL = headers.indexOf("New_Viewers_Can_Download (Yes/No)");
  const idxNewLimit = headers.indexOf("New_Limit_Access (Yes/No)");
  const idxStatus = headers.indexOf(CONFIG.STATUS_COL_HEADER);

  const startTime = Date.now();
  let countProcessed = 0;

  for (let i = 1; i < data.length; i++) {
    let row = data[i];
    if (row[idxAction] !== "TO_MODIFY" || row[idxStatus] === "OK") continue;

    if (Date.now() - startTime > CONFIG.MAX_EXEC_TIME_MS) {
      if (!isTrigger) ui.alert(`Paused. ${countProcessed} processed. Please restart.`);
      else scheduleNextRun_('startBulkUpdateFiles');
      return;
    }

    let fileId = String(row[idxFileId]).replace(/^'/, "");
    let mimeType = row[idxMimeType];
    let newShareStr = row[idxNewShare];
    let newDLStr = row[idxNewDL];
    let newLimitStr = row[idxNewLimit];
    
    try {
      let resourceUpdate = {};
      let hasUpdate = false;

      // 1. Sharing Rights
      if (newShareStr === "Yes" || newShareStr === "No") {
        resourceUpdate.writersCanShare = (newShareStr === "Yes");
        hasUpdate = true;
      }
      
      // 2. Download/Copy Rights
      if (newDLStr === "Yes" || newDLStr === "No") {
        if (mimeType !== "application/vnd.google-apps.folder") {
          resourceUpdate.copyRequiresWriterPermission = (newDLStr === "No");
          hasUpdate = true;
        } else if (newDLStr !== "-") {
          console.log(`Ignored: Download blocking does not apply to folders (${fileId})`);
        }
      }

      // 3. Inheritance Limitation
      if (newLimitStr === "Yes" || newLimitStr === "No") {
        if (mimeType !== "application/vnd.google-apps.folder") {
            throw new Error("Action denied: Inheritance limits only apply to Folders (Shared Drives), not Files.");
        }
        resourceUpdate.inheritedPermissionsDisabled = (newLimitStr === "Yes");
        hasUpdate = true;
      }

      if (!hasUpdate) throw new Error("Please enter 'Yes' or 'No' (or modification is not applicable).");

      Drive.Files.update(resourceUpdate, fileId, null, { supportsAllDrives: true });
      
      sheet.getRange(i + 1, idxStatus + 1).setValue("OK").setBackground("#d4edda");
      countProcessed++;

    } catch (e) {
      let errorMsg = e.message;
      if (errorMsg.includes("Bad Request") && newLimitStr === "Yes") {
          errorMsg = "Failed (My Drive): Inheritance limits require a Shared Drive.";
      } else if (errorMsg.includes("sufficient permissions")) {
          errorMsg = "Failed: You must be the File Owner to modify security rules.";
      }
      sheet.getRange(i + 1, idxStatus + 1).setValue("Error: " + errorMsg).setBackground("#f8d7da");
    }
  }
  if (!isTrigger) ui.alert(`Done. ${countProcessed} items updated.`);
}

/** =========================================================================
 * MODULE 3B: UPDATE ENGINE - PERMISSIONS (Drive.Permissions)
 * ========================================================================= */

/**
 * Applies requested permission modifications to specific files.
 * @param {Object} e - Event object if triggered by time-based trigger.
 * @return {void}
 */
function startBulkUpdatePermissions(e) {
  const isTrigger = (e && e.triggerUid) ? true : false;
  if (isTrigger) ScriptApp.getProjectTriggers().forEach(t => { if (t.getUniqueId() === e.triggerUid) ScriptApp.deleteTrigger(t); });

  let ui;
  if (!isTrigger) {
    ui = SpreadsheetApp.getUi();
    const response = ui.alert('⚠️ UPDATE Permissions', 'Apply modifications to users and access rights?', ui.ButtonSet.YES_NO);
    if (response !== ui.Button.YES) return;
  }

  const sheet = SpreadsheetApp.getActive().getSheetByName(CONFIG.SHEET_PERMS);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  
  const idxFileId = headers.indexOf("File_ID");
  const idxPermId = headers.indexOf("Permission_ID");
  const idxType = headers.indexOf("Entity_Type");
  const idxEmail = headers.indexOf("Email");
  const idxCurrentRole = headers.indexOf("Current_Role"); 
  const idxInherited = headers.indexOf("Inherited_Rights");
  const idxAction = headers.indexOf(CONFIG.ACTION_COL_HEADER);
  const idxNewRole = headers.indexOf("New_Role");
  const idxStatus = headers.indexOf(CONFIG.STATUS_COL_HEADER);

  const startTime = Date.now();
  let countProcessed = 0;

  for (let i = 1; i < data.length; i++) {
    let row = data[i];
    let action = row[idxAction];
    let status = row[idxStatus];

    if (action === "-" || action === "" || status === "OK" || status === "OK (Already deleted)") continue;

    if (Date.now() - startTime > CONFIG.MAX_EXEC_TIME_MS) {
      if (!isTrigger) ui.alert(`Paused. ${countProcessed} processed. Please restart.`);
      else scheduleNextRun_('startBulkUpdatePermissions');
      return;
    }

    let fileId = String(row[idxFileId]).replace(/^'/, "");
    let permId = String(row[idxPermId]).replace(/^'/, "");
    let email = row[idxEmail];
    let currentRole = row[idxCurrentRole]; 
    let newRole = row[idxNewRole];
    let type = row[idxType];
    let isInherited = row[idxInherited];

    try {
      if (isInherited === "Yes (Inherited)" && (action === "TO_MODIFY" || action === "TO_DELETE")) {
        throw new Error("Action blocked: Inherited right. Modify the parent folder.");
      }
      if (currentRole === "owner" && (action === "TO_MODIFY" || action === "TO_DELETE")) {
        throw new Error("Action blocked: Cannot delete the owner.");
      }

      switch (action) {
        case "TO_ADD":
          if (!email || !newRole) throw new Error("Email and New Role are mandatory.");
          Drive.Permissions.create({ emailAddress: email, type: type || 'user', role: newRole }, fileId, { supportsAllDrives: true, sendNotificationEmail: false });
          break;
        case "TO_MODIFY":
          if (!permId || !newRole) throw new Error("Permission ID and New Role are mandatory.");
          Drive.Permissions.update({ role: newRole }, fileId, permId, { supportsAllDrives: true });
          break;
        case "TO_DELETE":
          if (!permId) throw new Error("Permission ID is mandatory.");
          Drive.Permissions.remove(fileId, permId, { supportsAllDrives: true });
          break;
        default:
          throw new Error("Unrecognized action.");
      }
      sheet.getRange(i + 1, idxStatus + 1).setValue("OK").setBackground("#d4edda");
      countProcessed++;

    } catch (e) {
      let errorMsg = e.message;
      if (action === "TO_DELETE" && errorMsg.includes("Permission not found")) {
        sheet.getRange(i + 1, idxStatus + 1).setValue("OK (Already deleted)").setBackground("#d4edda");
        countProcessed++;
      } else {
        sheet.getRange(i + 1, idxStatus + 1).setValue("Error: " + errorMsg).setBackground("#f8d7da");
      }
    }
  }
  if (!isTrigger) ui.alert(`Done. ${countProcessed} permissions processed.`);
}

/** =========================================================================
 * UTILITY FUNCTIONS
 * ========================================================================= */

/**
 * Clears data from the relational sheets without deleting headers.
 * @return {void}
 */
function clearSheetData_() {
  const ss = SpreadsheetApp.getActive();
  [CONFIG.SHEET_FILES, CONFIG.SHEET_PERMS].forEach(name => {
    let s = ss.getSheetByName(name);
    if (s) {
      let lastRow = s.getLastRow();
      let lastCol = s.getLastColumn();
      if (lastRow > 1 && lastCol > 0) {
        let rangeToClear = s.getRange(2, 1, lastRow - 1, lastCol);
        rangeToClear.clearContent();
        rangeToClear.setBackground(null);
      }
    }
  });
}

/**
 * Flushes batched data to the spreadsheet.
 * @param {Array<Array<any>>} filesData 
 * @param {Array<Array<any>>} permsData 
 * @return {void}
 */
function flushDataToSheets_(filesData, permsData) {
  const ss = SpreadsheetApp.getActive();
  if (filesData.length > 0) {
    const sheetF = ss.getSheetByName(CONFIG.SHEET_FILES);
    sheetF.getRange(sheetF.getLastRow() + 1, 1, filesData.length, filesData[0].length).setValues(filesData);
  }
  if (permsData.length > 0) {
    const sheetP = ss.getSheetByName(CONFIG.SHEET_PERMS);
    sheetP.getRange(sheetP.getLastRow() + 1, 1, permsData.length, permsData[0].length).setValues(permsData);
  }
}

/**
 * Schedules a continuation trigger to bypass Apps Script timeout limits.
 * @param {string} functionName 
 * @return {void}
 */
function scheduleNextRun_(functionName) {
  ScriptApp.newTrigger(functionName).timeBased().after(60000).create();
}

/**
 * Deletes existing triggers for a specific function.
 * @param {string} functionName 
 * @return {void}
 */
function deleteTriggers_(functionName) {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === functionName) ScriptApp.deleteTrigger(t);
  });
}

/**
 * Clears the audit cache, deletes triggers, and wipes sheet data.
 * @return {void}
 */
function clearAuditCache() {
  const props = PropertiesService.getScriptProperties();
  props.deleteProperty('DRIVE_QUEUE');
  props.deleteProperty('LAST_STARTING_FOLDER');
  deleteTriggers_('processAuditQueue');
  deleteTriggers_('startBulkUpdateFiles');
  deleteTriggers_('startBulkUpdatePermissions');
  
  clearSheetData_();
  SpreadsheetApp.getUi().alert('Cache and sheets have been cleared.');
}
