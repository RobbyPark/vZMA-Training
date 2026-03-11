// ============================================================
//  MIAMI ARTCC – TRAINING QUEUE BACKEND
//  Google Apps Script – Code.gs
//
//  SETUP INSTRUCTIONS:
//  1. Go to script.google.com and create a new project
//  2. Paste this entire file into Code.gs
//  3. Fill in SHEET_ID and VATUSA_API_KEY below
//  4. Run setupSheets() once manually to create the sheet tabs
//  5. Deploy > New Deployment > Web App
//       Execute as: Me
//       Who has access: Anyone
//  6. Copy the Web App URL into your signup form HTML
// ============================================================

// ─── CONFIGURATION (fill these in) ──────────────────────────
const CONFIG = {
  SHEET_ID:       'YOUR_GOOGLE_SHEET_ID_HERE',   // From your Sheet URL
  VATUSA_API_KEY: 'YOUR_VATUSA_API_KEY_HERE',     // From your VATUSA facility dashboard
  ADMIN_EMAIL:    'your-email@example.com',       // Where new signup alerts go
  ARTCC_NAME:     'Miami ARTCC',
  ARTCC_CODE:     'ZMA',
};

// Sheet tab names
const TABS = {
  HOME:       'Home Queue',
  VISITOR:    'Visitor Queue',
  INELIGIBLE: 'Ineligible',
  PENDING:    'Pending Approval',
  ARCHIVED:   'Completed',
  REMOVED:    'Removed',
  DENIED:     'Denied & Ineligible',
  TRAINERS:      'Trainer Ratings',
  TRAINER_CACHE: 'Trainer Cache',
  LOG:        'Log',
};

// Column layout for Training Queue sheet
const COLS = {
  TIMESTAMP:       1,
  POSITION:        2,
  STATUS:          3,
  FIRST_NAME:      4,
  LAST_NAME:       5,
  CID:             6,
  RATING:          7,
  EMAIL:           8,
  NOTES:           9,
  UPDATED_AT:      10,
  QUEUE_TYPE:      11,
  REMOVAL_REASON:  12,
  EXAM_GRADE:      13,
  ZMA_EXAM_GRADE:      14,
  ASSIGNED_TRAINER:     15,
  TRAINER_ASSIGNED_AT:  16,
};


// ─── CORS HEADERS ────────────────────────────────────────────
function corsHeaders() {
  return ContentService.createTextOutput(JSON.stringify({ status: 'ok' }))
    .setMimeType(ContentService.MimeType.JSON);
}

function setHeaders(output) {
  return output; // Apps Script handles headers; CORS via no-cors on client
}


// ─── MAIN ENTRY POINTS ───────────────────────────────────────

function doGet(e) {
  const action = e.parameter.action || 'dashboard';
  if (action === 'signup') {
    return HtmlService.createHtmlOutputFromFile('SignupForm')
      .setTitle('ZMA Training Request')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
  if (action === 'examrequest') {
    return HtmlService.createHtmlOutputFromFile('ExamRequestForm')
      .setTitle('ZMA Exam Request')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
  if (action === 'examportal') {
    return HtmlService.createHtmlOutputFromFile('ExamPortal')
      .setTitle('ZMA Exam Center')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
  if (action === 'dashboard') {
    return HtmlService.createHtmlOutputFromFile('AdminDashboard')
      .setTitle('ZMA Training Admin')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
  // API calls from admin dashboard
  return handleApiGet(e);
}

// Called directly from SignupForm via google.script.run
function handleSignupFromForm(data) {
  return handleSignup(data);
}

// Called directly from AdminDashboard via google.script.run
function getSheetDataPublic(tabName) {
  return getSheetData(tabName);
}

function getTrainersPublic() {
  return fetchZmaTrainers();
}

function getPendingPublic() {
  return getSheetData(TABS.PENDING);
}

function getDeniedPublic() {
  return getSheetData(TABS.DENIED);
}

function loadTrainerRatingsPublic() {
  return loadAllTrainerRatings();
}

// Called directly from AdminDashboard via google.script.run
function handleAdminAction(data) {
  var action = data.action;
  if (action === 'updateStatus')  return handleUpdateStatus(data);
  if (action === 'deleteEntry')   return handleDelete(data);
  if (action === 'sendFollowUp')  return handleFollowUp(data);
  if (action === 'markComplete')  return handleMarkComplete(data);
  if (action === 'moveEntry')     return handleMoveEntry(data);
  if (action === 'reorderQueue')  return handleReorderQueue(data);
  if (action === 'removeEntry')    return handleRemoveEntry(data);
  if (action === 'saveNotes')      return handleSaveNotes(data);
  if (action === 'refreshZmaGrade') return handleRefreshZmaGrade(data);
  if (action === 'permanentDelete')        return handlePermanentDelete(data);
  if (action === 'permanentDeleteArchived') return handlePermanentDeleteArchived(data);
  if (action === 'permanentDeleteDenied')   return handlePermanentDeleteDenied(data);
  if (action === 'addCompletedRecord')      return handleAddCompletedRecord(data);
  if (action === 'approvePending')          return handleApprovePending(data);
  if (action === 'denyPending')             return handleDenyPending(data);
  if (action === 'sendTrainerAssignment') return handleSendTrainerAssignment(data);
  if (action === 'saveTrainerRatings')    return saveTrainerRatingsToSheet(data);
  if (action === 'loadTrainerRatings')    return loadAllTrainerRatings();
  if (action === 'saveAssignedTrainer')   return handleSaveAssignedTrainer(data);
  if (action === 'examRequestAction')      return handleExamRequestAction(data);
  if (action === 'examCenter')              return handleExamCenterAction(data);
  return { success: false, error: 'Unknown action' };
}

function doPost(e) {
  try {
    let data;
    // Handle both JSON and form-encoded submissions
    if (e.postData.type === 'application/json') {
      data = JSON.parse(e.postData.contents);
    } else {
      // Form submission - data is in a 'data' parameter
      data = JSON.parse(e.parameter.data || e.postData.contents);
    }
    const action = data.action || 'signup';

    if (action === 'signup')          return handleSignup(data);
    if (action === 'updateStatus')    return handleUpdateStatus(data);
    if (action === 'deleteEntry')     return handleDelete(data);
    if (action === 'sendFollowUp')    return handleFollowUp(data);
    if (action === 'markComplete')    return handleMarkComplete(data);
    if (action === 'moveEntry')       return handleMoveEntry(data);

    return jsonResponse({ success: false, error: 'Unknown action' });
  } catch (err) {
    logError('doPost', err);
    return jsonResponse({ success: false, error: err.message });
  }
}

function handleApiGet(e) {
  const action = e.parameter.action;
  if (action === 'getHome')       return jsonResponse(getSheetData(TABS.HOME));
  if (action === 'getVisitor')    return jsonResponse(getSheetData(TABS.VISITOR));
  if (action === 'getIneligible') return jsonResponse(getSheetData(TABS.INELIGIBLE));
  if (action === 'getArchived')   return jsonResponse(getSheetData(TABS.ARCHIVED));
  if (action === 'getRemoved')    return jsonResponse(getSheetData(TABS.REMOVED));
  // Legacy support
  if (action === 'getQueue')      return jsonResponse(getSheetData(TABS.HOME));
  return jsonResponse({ success: false, error: 'Unknown action' });
}


// ─── SIGNUP HANDLER ──────────────────────────────────────────

function handleSignup(data) {
  try {
  const { firstName, lastName, cid, rating, email: formEmail, timestamp } = data;

  if (!firstName || !lastName || !cid || !rating) {
    return { success: false, error: 'Missing required fields' };
  }

  // Look up member via VATUSA API
  let studentEmail = formEmail || '';
  let vatusaRating  = '';
  let member = null;
  try {
    member = fetchVatusaMember(cid);
    if (member) {
      studentEmail = member.email || studentEmail;
      vatusaRating  = member.rating_short || '';
    }
  } catch (err) {
    logError('VATUSA lookup', err);
    // Non-fatal — continue with form email
  }

  // Determine which queue based on VATUSA facility membership
  const queueType = determineQueue(cid, member);
  const tabName   = queueType === 'Home'       ? TABS.HOME
                  : queueType === 'Visitor'    ? TABS.VISITOR
                  : TABS.INELIGIBLE;

  const now      = new Date();

  // Fetch academy exam grade
  var examGrade    = '';
  var zmaExamGrade = '';
  try { examGrade    = fetchAcademyExamGrade(cid, rating); } catch(e) { logError('examGrade', e); }
  try { zmaExamGrade = fetchZmaExamGrade(cid, rating);     } catch(e) { logError('zmaExamGrade', e); }

  // Route to Pending unless this is an admin manual add (skipApproval flag)
  if (!data.skipApproval) {
    var pendingSheet = getSheet(TABS.PENDING);
    var pendingPos   = getNextPosition(pendingSheet);
    pendingSheet.appendRow([
      timestamp || now.toISOString(),
      pendingPos,
      'Pending',
      firstName,
      lastName,
      cid,
      rating,
      studentEmail,
      '',
      now.toISOString(),
      queueType,
      '',
      examGrade,
      zmaExamGrade,
      '',
      '',
    ]);
    logAction('PENDING', firstName + ' ' + lastName + ' (' + cid + ') pending approval for ' + rating);
    // Send emails separately — don't let email failure block the sheet write
    if (!data.skipEmail) {
      try { sendPendingEmail(studentEmail, firstName, rating); } catch(e) { logError('sendPendingEmail', e); }
      try { sendAdminNotification(firstName, lastName, cid, rating, queueType, studentEmail); } catch(e) { logError('sendAdminNotification', e); }
    }
    return { success: true, queued: true, pending: true, message: 'Your request has been submitted and is pending approval.' };
  }

  const sheet    = getSheet(resolveTabName(tabName));
  const position = getNextPosition(sheet);

  sheet.appendRow([
    timestamp || now.toISOString(),
    position,
    'Waiting',
    firstName,
    lastName,
    cid,
    rating,
    studentEmail,
    '',                // Notes
    now.toISOString(),
    queueType,         // Col 11 — queue type
    '',                // Col 12 — removal reason
    examGrade,         // Col 13 — exam grade
    zmaExamGrade,      // Col 14 — ZMA exam grade
    '',                // Col 15 — assigned trainer
    '',                // Col 16 — trainer assigned at
  ]);

  const lastRow = sheet.getLastRow();
  formatQueueRow(sheet, lastRow, 'Waiting');

  if (studentEmail && !data.skipEmail) {
    try { sendPendingEmail(studentEmail, firstName, rating); } catch(e) { logError('sendPendingEmail', e); }
  }
  try { sendAdminNotification(firstName, lastName, cid, rating, position, studentEmail, queueType); } catch(e) { logError('sendAdminNotification', e); }

  logAction('SIGNUP', `${firstName} ${lastName} (CID: ${cid}) → ${queueType} Queue (#${position})`);

  return { success: true, position: position, queueType: queueType, email: studentEmail ? 'sent' : 'no_email' };

  } catch(e) {
    logError('handleSignup', e);
    return { success: false, error: e.message };
  }
}


// ─── STATUS HANDLERS ─────────────────────────────────────────

function handleUpdateStatus(data) {
  const { rowIndex, status, notes, tabName } = data;
  const sheet = getSheet(resolveTabName(tabName) || TABS.HOME);
  const now   = new Date();

  sheet.getRange(rowIndex, COLS.STATUS).setValue(status);
  if (notes !== undefined) sheet.getRange(rowIndex, COLS.NOTES).setValue(notes);
  sheet.getRange(rowIndex, COLS.UPDATED_AT).setValue(now.toISOString());

  formatQueueRow(sheet, rowIndex, status);

  logAction('STATUS_UPDATE', `Row ${rowIndex} → ${status}`);
  return jsonResponse({ success: true });
}

function handleMarkComplete(data) {
  const { rowIndex, tabName } = data;
  const qSheet = getSheet(resolveTabName(tabName) || TABS.HOME);
  const aSheet = getSheet(TABS.ARCHIVED);

  // Copy row to Archive (all 16 cols)
  const lastCol = Math.max(qSheet.getLastColumn(), 16);
  const rowData = qSheet.getRange(rowIndex, 1, 1, lastCol).getValues()[0];
  while (rowData.length < 16) rowData.push('');
  rowData[COLS.STATUS - 1] = 'Completed';
  rowData[COLS.UPDATED_AT - 1] = new Date().toISOString();
  aSheet.appendRow(rowData);

  // Delete from queue
  qSheet.deleteRow(rowIndex);

  // Renumber positions
  renumberQueue(qSheet);

  logAction('COMPLETE', `Row ${rowIndex} moved to archive`);
  return jsonResponse({ success: true });
}

function handleDelete(data) {
  const { rowIndex, tabName } = data;
  const sheet = getSheet(resolveTabName(tabName) || TABS.HOME);
  sheet.deleteRow(rowIndex);
  renumberQueue(sheet);
  logAction('DELETE', `Row ${rowIndex} deleted from ${tabName}`);
  return jsonResponse({ success: true });
}

function handleFollowUp(data) {
  const { rowIndex, message, tabName } = data;
  const sheet     = getSheet(resolveTabName(tabName) || TABS.HOME);
  const email     = sheet.getRange(rowIndex, COLS.EMAIL).getValue();
  const firstName = sheet.getRange(rowIndex, COLS.FIRST_NAME).getValue();

  if (!email) return jsonResponse({ success: false, error: 'No email on file' });

  MailApp.sendEmail({
    to:       email,
    subject:  `[ZMA] Training Department Message`,
    body:     `Hello ${firstName},\n\n${message}\n\nBest regards,\nZMA Training Team`,
    htmlBody: buildFollowUpEmail(firstName, message),
    name:     'ZMA Training Team',
  });

  logAction('FOLLOW_UP', `Email sent to ${email} (row ${rowIndex})`);
  return jsonResponse({ success: true });
}


// ─── DATA GETTERS ─────────────────────────────────────────────

function getSheetData(tabName) {
  const sheet = getSheet(resolveTabName(tabName));
  const data  = sheet.getDataRange().getValues();
  if (data.length <= 1) return { success: true, rows: [] };

  const rows = data.slice(1).map(function(row, i) {
    return {
      rowIndex:   i + 2,
      timestamp:  row[0],
      position:   row[1],
      status:     row[2],
      firstName:  row[3],
      lastName:   row[4],
      cid:        row[5],
      rating:     row[6],
      email:      row[7],
      notes:      row[8],
      updatedAt:  row[9],
      queueType:       row[10] || tabName,
      tabName:         tabName,
      removalReason:   row[11] || '',
      examGrade:        row[12] || '',
      zmaExamGrade:     row[13] || '',
      assignedTrainer:      row[14] || '',
      trainerAssignedAt:    row[15] || '',
    };
  });

  return { success: true, rows: rows };
}

// Legacy alias
function getQueueData() { return getSheetData(TABS.HOME); }
function getArchivedData() { return getSheetData(TABS.ARCHIVED); }


// ─── VATUSA API ───────────────────────────────────────────────

function fetchVatusaMember(cid) {
  const url = `https://api.vatusa.net/v2/user/${cid}?apikey=${CONFIG.VATUSA_API_KEY}`;
  const res = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  if (res.getResponseCode() !== 200) return null;
  const json = JSON.parse(res.getContentText());
  return json.data || null;
}


// ─── REMOVE ENTRY (move to Removed tab) ─────────────────────

function handleRemoveEntry(data) {
  const { rowIndex, tabName, emailTemplate, customMessage } = data;
  const fromSheet = getSheet(resolveTabName(tabName) || TABS.HOME);
  // Ineligible removals go to Denied & Ineligible tab; all others go to Removed
  const toSheet   = (tabName === TABS.INELIGIBLE) ? getSheet(TABS.DENIED) : getSheet(TABS.REMOVED);
  const now       = new Date();

  // Read full row
  const lastCol = fromSheet.getLastColumn();
  const numCols = Math.max(lastCol, 16);
  var rowData = fromSheet.getRange(rowIndex, 1, 1, numCols).getValues()[0];
  // Pad if needed
  while (rowData.length < 16) rowData.push('');

  const email     = rowData[COLS.EMAIL - 1];
  const firstName = rowData[COLS.FIRST_NAME - 1];
  const rating    = rowData[COLS.RATING - 1];

  // Update status and removal reason
  rowData[COLS.STATUS - 1]          = 'Removed';
  rowData[COLS.UPDATED_AT - 1]      = now.toISOString();
  rowData[COLS.REMOVAL_REASON - 1]  = emailTemplate || 'Manual';

  // Append to Removed sheet
  toSheet.appendRow(rowData);
  formatQueueRow(toSheet, toSheet.getLastRow(), 'Removed');

  // Delete from source
  fromSheet.deleteRow(rowIndex);
  renumberQueue(fromSheet);

  // Send email if template specified
  if (email && emailTemplate && emailTemplate !== 'none') {
    sendRemovalEmail(email, firstName, rating, emailTemplate, customMessage);
  }

  logAction('REMOVE', `${firstName} (CID: ${rowData[COLS.CID-1]}) removed from ${tabName} — ${emailTemplate || 'no email'}`);
  return jsonResponse({ success: true });
}


function resolveTabName(tabName) {
  var map = {
    'home':      TABS.HOME,
    'visitor':   TABS.VISITOR,
    'ineligible':TABS.INELIGIBLE,
    'archived':  TABS.ARCHIVED,
    'removed':   TABS.REMOVED,
    'denied':    TABS.DENIED,
    'pending':   TABS.PENDING
  };
  return map[tabName] || tabName; // if already a full name, pass through
}

function handleSaveAssignedTrainer(data) {
  var rowIndex     = data.rowIndex;
  var tabName      = resolveTabName(data.tabName);
  var trainerName  = data.trainerName || '';
  var sheet        = getSheet(resolveTabName(tabName));

  // Read student info for trainer notification
  var studentFirstName = sheet.getRange(rowIndex, COLS.FIRST_NAME).getValue();
  var studentLastName  = sheet.getRange(rowIndex, COLS.LAST_NAME).getValue();
  var studentCid       = sheet.getRange(rowIndex, COLS.CID).getValue();
  var rating           = sheet.getRange(rowIndex, COLS.RATING).getValue();

  sheet.getRange(rowIndex, COLS.ASSIGNED_TRAINER).setValue(trainerName);
  sheet.getRange(rowIndex, COLS.STATUS).setValue(trainerName ? 'Training' : 'Waiting');
  sheet.getRange(rowIndex, COLS.TRAINER_ASSIGNED_AT).setValue(trainerName ? new Date().toISOString() : '');
  sheet.getRange(rowIndex, COLS.UPDATED_AT).setValue(new Date().toISOString());

  // Notify trainer if requested
  if (data.notifyTrainer && data.trainerCid) {
    try { sendTrainerNotificationEmail(data.trainerCid, trainerName, studentFirstName + ' ' + studentLastName, studentCid, rating); }
    catch(e) { logError('trainerNotify', e); }
  }

  logAction('ASSIGN_TRAINER', 'Row ' + rowIndex + ' in ' + tabName + ' assigned to ' + trainerName);
  return jsonResponse({ success: true });
}

function handleSaveNotes(data) {
  const { rowIndex, tabName, notes } = data;
  const sheet = getSheet(resolveTabName(tabName) || TABS.HOME);
  sheet.getRange(rowIndex, COLS.NOTES).setValue(notes || '');
  sheet.getRange(rowIndex, COLS.UPDATED_AT).setValue(new Date().toISOString());
  logAction('NOTES', `Notes saved for row ${rowIndex} in ${tabName}`);
  return jsonResponse({ success: true });
}

function handleRefreshZmaGrade(data) {
  const { rowIndex, tabName, cid, trainingType } = data;
  const sheet = getSheet(resolveTabName(tabName));
  const grade = fetchZmaExamGrade(cid, trainingType);
  sheet.getRange(rowIndex, COLS.ZMA_EXAM_GRADE).setValue(grade);
  sheet.getRange(rowIndex, COLS.UPDATED_AT).setValue(new Date().toISOString());
  return jsonResponse({ success: true, grade: grade });
}

function handlePermanentDelete(data) {
  const { rowIndex } = data;
  const sheet = getSheet(TABS.REMOVED);
  sheet.deleteRow(rowIndex);
  logAction('PERM_DELETE', `Row ${rowIndex} permanently deleted from Removed`);
  return jsonResponse({ success: true });
}


// ─── PENDING APPROVAL HANDLERS ───────────────────────────────

function handleApprovePending(data) {
  try {
    var rowIndex = data.rowIndex;
    var pendingSheet = getSheet(TABS.PENDING);
    var lastCol = Math.max(pendingSheet.getLastColumn(), 16);
    var rowData = pendingSheet.getRange(rowIndex, 1, 1, lastCol).getValues()[0];
    while (rowData.length < 16) rowData.push('');

    var queueType = rowData[COLS.QUEUE_TYPE - 1] || 'Home';
    var tabName   = queueType === 'Home'    ? TABS.HOME
                  : queueType === 'Visitor' ? TABS.VISITOR
                  : TABS.INELIGIBLE;

    var destSheet = getSheet(resolveTabName(tabName));
    var position  = getNextPosition(destSheet);
    rowData[COLS.POSITION  - 1] = position;
    rowData[COLS.STATUS    - 1] = 'Waiting';
    rowData[COLS.UPDATED_AT- 1] = new Date().toISOString();

    destSheet.appendRow(rowData);
    formatQueueRow(destSheet, destSheet.getLastRow(), 'Waiting');
    pendingSheet.deleteRow(rowIndex);
    renumberQueue(pendingSheet);

    // Notify student
    var email     = rowData[COLS.EMAIL      - 1];
    var firstName = rowData[COLS.FIRST_NAME - 1];
    var rating    = rowData[COLS.RATING     - 1];
    if (email) sendApprovalEmail(email, firstName, rating, tabName);

    logAction('APPROVE', firstName + ' approved into ' + tabName);
    return jsonResponse({ success: true });
  } catch(e) {
    logError('handleApprovePending', e);
    return jsonResponse({ success: false, error: e.message });
  }
}

function handleDenyPending(data) {
  try {
    var rowIndex = data.rowIndex;
    var pendingSheet = getSheet(TABS.PENDING);
    var lastCol = Math.max(pendingSheet.getLastColumn(), 16);
    var rowData = pendingSheet.getRange(rowIndex, 1, 1, lastCol).getValues()[0];
    while (rowData.length < 16) rowData.push('');

    var email     = rowData[COLS.EMAIL      - 1];
    var firstName = rowData[COLS.FIRST_NAME - 1];
    var rating    = rowData[COLS.RATING     - 1];
    var now       = new Date();

    // Move to Denied & Ineligible tab
    var deniedSheet = getSheet(TABS.DENIED);
    rowData[COLS.STATUS     - 1] = 'Denied';
    rowData[COLS.UPDATED_AT - 1] = now.toISOString();
    rowData[COLS.REMOVAL_REASON - 1] = 'Denied from pending';
    deniedSheet.appendRow(rowData);

    pendingSheet.deleteRow(rowIndex);
    renumberQueue(pendingSheet);

    var template = data.emailTemplate || 'none';
    var custom   = data.customMessage  || '';
    if (template !== 'none' && email) {
      try { sendRemovalEmail(email, firstName, rating, template, custom); } catch(e) { logError('denyEmail', e); }
    }
    logAction('DENY', firstName + ' denied and moved to Denied & Ineligible tab');
    return jsonResponse({ success: true });
  } catch(e) {
    logError('handleDenyPending', e);
    return jsonResponse({ success: false, error: e.message });
  }
}

function sendPendingEmail(email, firstName, rating) {
  if (!email) return;
  var subject = '[ZMA] Training Request Received';
  var bodyContent = '<p style="margin:0 0 16px 0;font-size:16px;color:#1a2a3a;">Hello <strong>' + firstName + '</strong>,</p>' +
    '<p style="margin:0 0 16px 0;font-size:15px;color:#2a3a4a;line-height:1.7;">Thank you for submitting your request to join the Miami ARTCC Training Queue. Your request has been received, and you will receive a follow-up email once an admin reviews your request.</p>';
  var plain = 'Hello ' + firstName + ',\n\nThank you for submitting your request to join the Miami ARTCC Training Queue. Your request has been received, and you will receive a follow-up email once an admin reviews your request.\n\nZMA Training Team';
  MailApp.sendEmail({ to: email, subject: subject, body: plain, htmlBody: buildEmailHtml('Training Request', 'REQUEST RECEIVED', bodyContent), name: 'ZMA Training Team' });
}

function getExamInfo(rating) {
  // Map training type to exam tier and codes. Returns null if no exam required.
  var map = {
    // Tier 2
    'S1 Training (Tier 2)':                              { tier: 'T2', label: 'S1 Tier 2',                      codes: 'S1 Tier 2 - poJHH4246' },
    'S2 Training (Tier 2)':                              { tier: 'T2', label: 'S2 Tier 2',                      codes: 'S2 Tier 2 - qnd53K' },
    'S3 Training (Tier 2)':                              { tier: 'T2', label: 'S3 Tier 2',                      codes: 'S3 Tier 2 - dg;BN43' },
    'C1 Training':                                       { tier: 'T2', label: 'C1 Training / Visitor',          codes: 'C1 Training, Visitor/Transfer - Miami Center Domestic - :diu81G' },
    'Visitor/Transfer - Miami Center Domestic Training': { tier: 'T2', label: 'C1 Training / Visitor',          codes: 'C1 Training, Visitor/Transfer - Miami Center Domestic - :diu81G' },
    // Tier 1
    'Tier 1 MIA GND/DEL':                               { tier: 'T1', label: 'Tier 1 MIA GND/DEL',             codes: 'Miami GND/DEL - Gdc36fz' },
    'Tier 1 MIA LCL':                                   { tier: 'T1', label: 'Tier 1 MIA LCL',                 codes: 'Miami LCL - JHger10c' },
    'Tier 1 MIA TRACON':                                { tier: 'T1', label: 'Tier 1 MIA TRACON',              codes: 'Tier 1 MIA TRACON - Kjdo47GA' },
    'Miami Center Oceanic':                              { tier: 'T1', label: 'Miami Center Oceanic',           codes: 'Miami Center Oceanic - sDf135' },
    'Visitor/Transfer - Miami CAB':                      { tier: 'T1', label: 'Visitor/Transfer - Miami CAB',   codes: 'Miami GND/DEL - Gdc36fz, Miami LCL - JHger10c' },
  };
  return map[rating] || null;
}

function sendApprovalEmail(email, firstName, rating, queueType) {
  if (!email) return;
  var subject = '[ZMA] Training Request Approved';

  var examInfo = getExamInfo(rating);
  var examParaHtml = '';
  var examParaPlain = '';
  if (examInfo) {
    var tierLabel = examInfo.tier === 'T1' ? 'ZMA T1' : 'ZMA T2';
    examParaHtml = '<p style="margin:16px 0 0 0;font-size:15px;color:#2a3a4a;line-height:1.7;">In addition, you are required to complete the <strong>' + tierLabel + '</strong> exam. The code for ' + examInfo.label + ' is: <strong>' + examInfo.codes + '</strong>. You will have 30 days to complete this exam from the time you receive this email. Please note that you will not be eligible to receive your rating upon completion of training without completing this written exam. If you have any questions, please email <a href="mailto:ta@zmaartcc.net" style="color:#0ea5c8;">ta@zmaartcc.net</a>.</p>';
    examParaPlain = '\n\nIn addition, you are required to complete the ' + tierLabel + ' exam. The code for ' + examInfo.label + ' is: ' + examInfo.codes + '. You will have 30 days to complete this exam from the time you receive this email. Please note that you will not be eligible to receive your rating upon completion of training without completing this written exam. If you have any questions, please email ta@zmaartcc.net.';
  }

  var bodyContent =
    '<p style="margin:0 0 16px 0;font-size:16px;color:#1a2a3a;">Hello <strong>' + firstName + '</strong>,</p>' +
    '<p style="margin:0 0 16px 0;font-size:15px;color:#2a3a4a;line-height:1.7;">Your training request for the <strong>' + rating + '</strong> rating has been approved.</p>' +
    '<p style="margin:0 0 16px 0;font-size:15px;color:#2a3a4a;line-height:1.7;">The Miami ARTCC Training Team is actively working to assign trainers and certify students as efficiently as possible while maintaining our high training standards. We appreciate your patience during this process.</p>' +
    '<p style="margin:0 0 0 0;font-size:15px;color:#2a3a4a;line-height:1.7;">While you await trainer assignment, please begin reviewing all relevant training materials for your requested rating. Per VATUSA and ZMA policy, students are expected to arrive at training sessions fully prepared and with a baseline understanding of all required concepts. Please ensure you review the current training policy in its entirety, as it provides important guidance and clarity on the training process and expectations.</p>' +
    examParaHtml +
    '<p style="margin:16px 0 0 0;font-size:15px;color:#2a3a4a;line-height:1.7;">Thank you again for your patience. We look forward to training with you soon.</p>';

  var plain = 'Hello ' + firstName + ',\n\n' +
    'Your training request for the ' + rating + ' rating has been approved.\n\n' +
    'The Miami ARTCC Training Team is actively working to assign trainers and certify students as efficiently as possible while maintaining our high training standards. We appreciate your patience during this process.\n\n' +
    'While you await trainer assignment, please begin reviewing all relevant training materials for your requested rating. Per VATUSA and ZMA policy, students are expected to arrive at training sessions fully prepared and with a baseline understanding of all required concepts. Please ensure you review the current training policy in its entirety, as it provides important guidance and clarity on the training process and expectations.' +
    examParaPlain + '\n\n' +
    'Thank you again for your patience. We look forward to training with you soon.\n\nBest regards,\nZMA Training Team';

  MailApp.sendEmail({ to: email, subject: subject, body: plain, htmlBody: buildEmailHtml('Training Request', 'REQUEST APPROVED', bodyContent), name: 'ZMA Training Team' });
}


function sendRemovalEmail(email, firstName, rating, template, customMessage) {
  var subject, bodyHtml, plainBody;

  if (template === 'inactivity') {
    subject = '[ZMA] Training Queue Removal Notice';
    var bodyContent = `
      <p style="margin:0 0 24px 0;font-size:16px;color:#1a2a3a;line-height:1.5;">Hello <strong>${firstName}</strong>,</p>
      <p style="margin:0 0 18px 0;font-size:15px;color:#2a3a4a;line-height:1.7;">This email is to inform you that you have been removed from the Miami ARTCC Training Queue due to <strong style="color:#0a3d6b;">continued inactivity</strong>.</p>
      <p style="margin:0 0 18px 0;font-size:15px;color:#2a3a4a;line-height:1.7;">If you wish to resume training in the future, you may submit a new training request once you are able to fully commit to the expectations outlined in VATUSA and ZMA training policies.</p>
      <p style="margin:0 0 0 0;font-size:15px;color:#2a3a4a;line-height:1.7;">If you have questions regarding this decision, you may contact the Training Administrator for clarification.</p>`;
    bodyHtml  = buildEmailHtml('Training Queue', 'REMOVAL NOTICE', bodyContent);
    plainBody = `Hello ${firstName},

This email is to inform you that you have been removed from the Miami ARTCC Training Queue due to continued inactivity.

If you wish to resume training in the future, you may submit a new training request once you are able to fully commit to the expectations outlined in VATUSA and ZMA training policies.

If you have questions regarding this decision, you may contact the Training Administrator for clarification.

Best regards,
ZMA Training Team`;

  } else if (template === 'ineligibility') {
    subject = '[ZMA] Training Queue Removal Notice';
    var bodyContent = `
      <p style="margin:0 0 24px 0;font-size:16px;color:#1a2a3a;line-height:1.5;">Hello <strong>${firstName}</strong>,</p>
      <p style="margin:0 0 18px 0;font-size:15px;color:#2a3a4a;line-height:1.7;">This email is to inform you that you have been removed from the Miami ARTCC Training Queue due to <strong style="color:#0a3d6b;">ineligibility per T01</strong>.</p>
      <p style="margin:0 0 18px 0;font-size:15px;color:#2a3a4a;line-height:1.7;">If you wish to resume training in the future, you may submit a new training request once you are able to fully commit to the expectations outlined in VATUSA and ZMA training policies.</p>
      <p style="margin:0 0 0 0;font-size:15px;color:#2a3a4a;line-height:1.7;">If you have questions regarding this decision, you may contact the Training Administrator for clarification.</p>`;
    bodyHtml  = buildEmailHtml('Training Queue', 'REMOVAL NOTICE', bodyContent);
    plainBody = `Hello ${firstName},

This email is to inform you that you have been removed from the Miami ARTCC Training Queue due to ineligibility per T01.

If you wish to resume training in the future, you may submit a new training request once you are able to fully commit to the expectations outlined in VATUSA and ZMA training policies.

If you have questions regarding this decision, you may contact the Training Administrator for clarification.

Best regards,
ZMA Training Team`;

  } else if (template === 'custom' && customMessage) {
    subject = '[ZMA] Training Queue Notice';
    var bodyContent = `
      <p style="margin:0 0 24px 0;font-size:16px;color:#1a2a3a;line-height:1.5;">Hello <strong>${firstName}</strong>,</p>
      <table width="100%" cellpadding="0" cellspacing="0" style="margin-bottom:18px;">
        <tr><td style="background:#f0f6ff;border-left:4px solid #0099bb;border-radius:4px;padding:16px 20px;font-size:15px;color:#1a2a3a;line-height:1.7;">${customMessage}</td></tr>
      </table>`;
    bodyHtml  = buildEmailHtml('Training Queue', 'NOTICE', bodyContent);
    plainBody = `Hello ${firstName},

${customMessage}

Best regards,
ZMA Training Team`;
  } else {
    return;
  }

  MailApp.sendEmail({ to: email, subject: subject, body: plainBody, htmlBody: bodyHtml, name: 'ZMA Training Team' });
}


function buildFollowUpEmail(firstName, message) {
  const bodyHtml = `
    <p style="margin:0 0 24px 0;font-size:16px;color:#1a2a3a;line-height:1.5;">Hello <strong>${firstName}</strong>,</p>
    <p style="margin:0 0 18px 0;font-size:15px;color:#2a3a4a;line-height:1.7;">You have a message from the Miami ARTCC Training Department:</p>
    <table width="100%" cellpadding="0" cellspacing="0" style="margin-bottom:18px;">
      <tr><td style="background:#f0f6ff;border-left:4px solid #0099bb;border-radius:4px;padding:16px 20px;font-size:15px;color:#1a2a3a;line-height:1.7;">${message}</td></tr>
    </table>
    <p style="margin:0;font-size:15px;color:#2a3a4a;line-height:1.7;">If you have any questions, please reply to this email or reach out via Discord.</p>`;
  return buildEmailHtml('Training Department', 'MESSAGE', bodyHtml);
}

function sendAdminNotification(firstName, lastName, cid, rating, position, email, queueType) {
  const subject = `[ZMA Admin] New ${queueType} Training Request – ${firstName} ${lastName}`;
  const bodyHtml = `
    <p style="margin:0 0 24px 0;font-size:16px;color:#1a2a3a;line-height:1.5;">A new training request has been submitted and requires your attention.</p>
    <table width="100%" cellpadding="0" cellspacing="0" style="margin-bottom:24px;">
      <tr><td style="background:#f0f6ff;border-radius:8px;padding:20px 24px;">
        <table width="100%" cellpadding="0" cellspacing="0">
          <tr><td style="padding:6px 0;font-size:14px;color:#6a8aaa;width:120px;">Name</td><td style="padding:6px 0;font-size:14px;color:#1a2a3a;font-weight:600;">${firstName} ${lastName}</td></tr>
          <tr><td style="padding:6px 0;font-size:14px;color:#6a8aaa;">CID</td><td style="padding:6px 0;font-size:14px;color:#1a2a3a;font-family:monospace;">${cid}</td></tr>
          <tr><td style="padding:6px 0;font-size:14px;color:#6a8aaa;">Training</td><td style="padding:6px 0;font-size:14px;color:#1a2a3a;">${rating}</td></tr>
          <tr><td style="padding:6px 0;font-size:14px;color:#6a8aaa;">Queue</td><td style="padding:6px 0;font-size:14px;color:#0a3d6b;font-weight:700;">${queueType}</td></tr>
          <tr><td style="padding:6px 0;font-size:14px;color:#6a8aaa;">Position</td><td style="padding:6px 0;font-size:14px;color:#1a2a3a;">#${position}</td></tr>
          <tr><td style="padding:6px 0;font-size:14px;color:#6a8aaa;">Email</td><td style="padding:6px 0;font-size:14px;color:#1a2a3a;">${email || 'Not found'}</td></tr>
        </table>
      </td></tr>
    </table>
    <p style="margin:0;font-size:15px;color:#2a3a4a;line-height:1.7;">Log in to the admin dashboard to review and manage this request.</p>`;
  const htmlBody = buildEmailHtml('New Training Request', queueType.toUpperCase() + ' QUEUE', bodyHtml);
  const plainBody = `New training request:\n\nName: ${firstName} ${lastName}\nCID: ${cid}\nRating: ${rating}\nQueue: ${queueType}\nPosition: #${position}\nEmail: ${email || 'Not found'}`;
  MailApp.sendEmail({ to: CONFIG.ADMIN_EMAIL, subject, body: plainBody, htmlBody, name: 'ZMA Training Admin' });
}


// ─── SHEET HELPERS ────────────────────────────────────────────

function getSheet(name) {
  const ss    = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  let   sheet = ss.getSheetByName(name);
  if (!sheet) {
    // Auto-create missing sheet with standard headers rather than crash
    sheet = ss.insertSheet(name);
    const qHeaders = ['Submitted','Position','Status','First Name','Last Name','CID','Training Type','Email','Notes','Updated At','Queue Type','Removal Reason','Exam Grade (VATUSA)','ZMA Exam Grade','Assigned Trainer','Trainer Assigned At'];
    const headerRange = sheet.getRange(1, 1, 1, qHeaders.length);
    headerRange.setValues([qHeaders]);
    headerRange.setFontWeight('bold');
    headerRange.setBackground('#1565c0');
    headerRange.setFontColor('#ffffff');
    sheet.setFrozenRows(1);
    logAction('AUTO_CREATE_SHEET', 'Created missing sheet: ' + name);
  }
  return sheet;
}

function getNextPosition(sheet) {
  const last = sheet.getLastRow();
  if (last <= 1) return 1;
  const positions = sheet.getRange(2, COLS.POSITION, last - 1, 1).getValues();
  const max = Math.max(0, ...positions.map(r => Number(r[0]) || 0));
  return max + 1;
}

function renumberQueue(sheet) {
  const last = sheet.getLastRow();
  if (last <= 1) return;
  for (let r = 2; r <= last; r++) {
    sheet.getRange(r, COLS.POSITION).setValue(r - 1);
  }
}

function formatQueueRow(sheet, rowIndex, status) {
  const range = sheet.getRange(rowIndex, 1, 1, 10);
  const colors = {
    'Pending':   '#fff3cd',
    'Approved':  '#d4edda',
    'Rejected':  '#f8d7da',
    'Completed': '#d1ecf1',
  };
  range.setBackground(colors[status] || '#ffffff');
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function logAction(type, message) {
  try {
    const sheet = getSheet(TABS.LOG);
    sheet.appendRow([new Date().toISOString(), type, message]);
  } catch (e) { /* Log sheet missing — non-fatal */ }
}

function logError(context, err) {
  console.error(`[${context}] ${err.message}`);
  logAction('ERROR', `${context}: ${err.message}`);
}


// ─── ONE-TIME SETUP ───────────────────────────────────────────
// Run this function ONCE manually from the Apps Script editor

function setupSheets() {
  const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);

  function ensureSheet(name, headers) {
    let sheet = ss.getSheetByName(name);
    if (!sheet) {
      // Only create and format if the sheet doesn't exist yet
      sheet = ss.insertSheet(name);
      const headerRange = sheet.getRange(1, 1, 1, headers.length);
      headerRange.setValues([headers]);
      headerRange.setFontWeight('bold');
      headerRange.setBackground('#1565c0');
      headerRange.setFontColor('#ffffff');
      sheet.setFrozenRows(1);
    }
    // If sheet already exists, leave all data intact
    return sheet;
  }

  const qHeaders = ['Submitted', 'Position', 'Status', 'First Name', 'Last Name', 'CID', 'Training Type', 'Email', 'Notes', 'Updated At', 'Queue Type', 'Removal Reason', 'Exam Grade', 'ZMA Exam Grade'];
  ensureSheet(TABS.HOME,       qHeaders);
  ensureSheet(TABS.VISITOR,    qHeaders);
  ensureSheet(TABS.INELIGIBLE, qHeaders);
  ensureSheet(TABS.ARCHIVED,   qHeaders);
  ensureSheet(TABS.PENDING,    qHeaders);
  ensureSheet(TABS.REMOVED,    qHeaders);
  ensureSheet(TABS.DENIED,     qHeaders);
  ensureSheet(TABS.TRAINERS,   ['CID', 'Name', 'Ratings', 'Max Students']);
  ensureSheet(TABS.LOG,        ['Timestamp', 'Type', 'Message']);
  getExamRequestSheet(); // ensure Exam Requests tab exists with correct headers

  SpreadsheetApp.getUi().alert('✅ Sheets created successfully! You can now deploy the Web App.');
}


// ─── QUEUE DETERMINATION ─────────────────────────────────────────────────────

function determineQueue(cid, member) {
  try {
    if (!member) {
      member = fetchVatusaMember(cid);
    }
    if (!member) return 'Ineligible';

    const facility = member.facility || '';

    // Home controller = primary facility is ZMA
    if (facility === 'ZMA') return 'Home';

    // Visitor = has a visiting relationship with ZMA
    const visiting = member.visiting || [];
    let isVisitor = false;
    visiting.forEach(function(v) {
      if ((v.facility || '').toUpperCase() === 'ZMA') isVisitor = true;
    });
    if (isVisitor) return 'Visitor';

    // Also check via facility_id if available
    if (member.facility_id && member.facility_id === 'ZMA') return 'Home';

    return 'Ineligible';
  } catch(err) {
    logError('determineQueue', err);
    return 'Ineligible';
  }
}


// ─── VATUSA / ZMA EXAM GRADE FETCHING ────────────────────────────────────────

function fetchAcademyExamGrade(cid, trainingType) {
  try {
    var url = 'https://api.vatusa.net/v2/user/' + cid + '/exam/history?apikey=' + CONFIG.VATUSA_API_KEY;
    var resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    if (resp.getResponseCode() !== 200) return '';
    var data = JSON.parse(resp.getContentText());
    var exams = (data.data || data || []);
    if (!Array.isArray(exams)) return '';
    // Return the most recent score as a string
    if (exams.length === 0) return '';
    var latest = exams[exams.length - 1];
    return latest.score !== undefined ? String(latest.score) + '%' : '';
  } catch(e) {
    logError('fetchAcademyExamGrade', e);
    return '';
  }
}

function fetchZmaExamGrade(cid, trainingType) {
  try {
    var ss = SpreadsheetApp.openById(ZMA_GRADEBOOK_ID);
    var sheets = ss.getSheets();
    for (var i = 0; i < sheets.length; i++) {
      var sheet = sheets[i];
      var data = sheet.getDataRange().getValues();
      for (var r = 1; r < data.length; r++) {
        if (String(data[r][0]) === String(cid)) {
          // Return the grade from column 2 (index 1) or whatever column has it
          return data[r][1] ? String(data[r][1]) : '';
        }
      }
    }
    return '';
  } catch(e) {
    logError('fetchZmaExamGrade', e);
    return '';
  }
}


// ─── ZMA TRAINER ROSTER ──────────────────────────────────────────────────────

function fetchZmaTrainers() {
  try {
    var url = 'https://api.vatusa.net/v2/facility/ZMA/roster/both?apikey=' + CONFIG.VATUSA_API_KEY;
    var resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    if (resp.getResponseCode() !== 200) return { success: false, trainers: [] };
    var data = JSON.parse(resp.getContentText());
    var roster = data.data || data || [];
    if (!Array.isArray(roster)) return { success: false, trainers: [] };

    var trainers = roster.filter(function(m) {
      var roles = m.roles || [];
      return roles.some(function(r) {
        return (r.role === 'INS' || r.role === 'MTR') &&
               (r.facility || '').toUpperCase() === 'ZMA';
      });
    }).map(function(m) {
      var zmRoles = (m.roles || []).filter(function(r) {
                     return (r.role === 'INS' || r.role === 'MTR') &&
                            (r.facility || '').toUpperCase() === 'ZMA';
                   }).map(function(r) { return r.role; });
      var hasIns = zmRoles.indexOf('INS') !== -1;
      var hasMtr = zmRoles.indexOf('MTR') !== -1;
      var roleLabel = (hasIns && hasMtr) ? 'Instructor / Mentor'
                    : hasIns ? 'Instructor'
                    : hasMtr ? 'Mentor'
                    : 'Staff';
      return {
        cid:       m.cid,
        firstName: m.fname || '',
        lastName:  m.lname || '',
        name:      (m.fname || '') + ' ' + (m.lname || ''),
        rating:    m.rating_short || '',
        role:      roleLabel,
        roles:     zmRoles
      };
    });

    return { success: true, trainers: trainers };
  } catch(e) {
    logError('fetchZmaTrainers', e);
    return { success: false, trainers: [], error: e.message };
  }
}


// ─── TRAINER RATINGS SHEET ───────────────────────────────────────────────────

function loadAllTrainerRatings() {
  try {
    var sheet = getSheet(TABS.TRAINERS);
    var data = sheet.getDataRange().getValues();
    var result = {};
    for (var i = 1; i < data.length; i++) {
      var cid = String(data[i][0]).trim();
      if (!cid) continue;
      var ratings = data[i][1] ? String(data[i][1]).split(',').map(function(r){return r.trim();}).filter(Boolean) : [];
      var maxStudents = data[i][2] ? String(data[i][2]).trim() : '';
      result[cid] = { ratings: ratings, maxStudents: maxStudents };
    }
    return { success: true, ratings: result };
  } catch(e) {
    logError('loadAllTrainerRatings', e);
    return { success: false, ratings: {} };
  }
}

function saveTrainerRatingsToSheet(data) {
  try {
    var sheet = getSheet(TABS.TRAINERS);
    var cid = String(data.cid || '').trim();
    var name = data.name || '';
    var ratings = Array.isArray(data.ratings) ? data.ratings.join(', ') : String(data.ratings || '');
    var maxStudents = String(data.maxStudents || '');

    // Find existing row for this CID and update it, or append new row
    var values = sheet.getDataRange().getValues();
    var found = false;
    for (var i = 1; i < values.length; i++) {
      if (String(values[i][0]).trim() === cid) {
        sheet.getRange(i + 1, 1, 1, 4).setValues([[cid, ratings, maxStudents, name]]);
        found = true;
        break;
      }
    }
    if (!found) {
      sheet.appendRow([cid, ratings, maxStudents, name]);
    }
    return { success: true };
  } catch(e) {
    logError('saveTrainerRatingsToSheet', e);
    return { success: false, error: e.message };
  }
}


// ─── TRAINER ASSIGNMENT EMAIL ─────────────────────────────────────────────────

function sendTrainerNotificationEmail(trainerCid, trainerName, studentName, studentCid, rating) {
  try {
    if (!trainerCid) return;
    // Look up trainer email via VATUSA
    var url = 'https://api.vatusa.net/v2/user/' + trainerCid + '?apikey=' + CONFIG.VATUSA_API_KEY;
    var resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    if (resp.getResponseCode() !== 200) return;
    var data = JSON.parse(resp.getContentText());
    var trainerData = data.data || data;
    var trainerEmail = trainerData.email || '';
    if (!trainerEmail) return;

    var trainerFirstName = trainerName ? trainerName.split(' ')[0] : 'Trainer';
    var subject = '[ZMA] New Student Assignment – ' + studentName;
    var bodyHtml = '<p>Hello ' + trainerFirstName + ',</p>' +
      '<p>You have been assigned <strong>' + studentName + '</strong> (CID: ' + studentCid + ') ' +
      'as your student for <strong>' + rating + '</strong>.</p>' +
      '<p>A Discord thread will be made, and your student will post their availability there. ' +
      'If you have any questions or concerns, please contact the TA.</p>' +
      '<p>ZMA Training Team</p>';
    var plainBody = 'Hello ' + trainerFirstName + ',\n\n' +
      'You have been assigned ' + studentName + ' (CID: ' + studentCid + ') ' +
      'as your student for ' + rating + '.\n\n' +
      'A Discord thread will be made, and your student will post their availability there. ' +
      'If you have any questions or concerns, please contact the TA.\n\n' +
      'ZMA Training Team';

    MailApp.sendEmail({
      to: trainerEmail,
      subject: subject,
      body: plainBody,
      htmlBody: buildEmailHtml('Training Assignment', 'STUDENT ASSIGNED', bodyHtml),
      name: 'ZMA Training Team'
    });
    logAction('TRAINER_EMAIL', 'Sent assignment email to ' + trainerName + ' (' + trainerEmail + ')');
  } catch(e) {
    logError('sendTrainerNotificationEmail', e);
  }
}

function handleSendTrainerAssignment(data) {
  try {
    var sheet = getSheet(resolveTabName(data.tabName) || (data.queueType === 'Visitor' ? TABS.VISITOR : TABS.HOME));
    var rowIndex = data.rowIndex;
    var trainerName = data.trainerName || '';
    var trainerCid  = data.trainerCid  || '';

    sheet.getRange(rowIndex, COLS.ASSIGNED_TRAINER).setValue(trainerName);
    sheet.getRange(rowIndex, COLS.STATUS).setValue('Training');
    sheet.getRange(rowIndex, COLS.TRAINER_ASSIGNED_AT).setValue(trainerName ? new Date().toISOString() : '');
    sheet.getRange(rowIndex, COLS.UPDATED_AT).setValue(new Date().toISOString());
    formatQueueRow(sheet, rowIndex, 'Training');
    logAction('SEND_TRAINER_ASSIGNMENT', 'Assigned ' + trainerName + ' row ' + rowIndex);

    // Send student email
    if (data.sendStudentEmail) {
      var rowData = sheet.getRange(rowIndex, 1, 1, 16).getValues()[0];
      var studentEmail = rowData[COLS.EMAIL - 1] || '';
      var studentFirst = rowData[COLS.FIRST_NAME - 1] || '';
      var rating = rowData[COLS.TRAINING_TYPE - 1] || '';
      if (studentEmail) {
        var subject = '[ZMA] Trainer Assigned – ' + rating;
        var bodyHtml = '<p>Hello ' + studentFirst + ',</p>' +
          '<p>You have been assigned a trainer: <strong>' + trainerName + '</strong>.</p>' +
          '<p>A Discord thread will be created. Please post your availability there.</p>';
        MailApp.sendEmail({ to: studentEmail, subject: subject,
          body: 'Hello ' + studentFirst + ',\n\nYou have been assigned a trainer: ' + trainerName + '.',
          htmlBody: buildEmailHtml('Training Update', 'TRAINER ASSIGNED', bodyHtml),
          name: 'ZMA Training Team' });
      }
    }

    // Optionally notify trainer
    if (data.notifyTrainer && trainerCid) {
      var rowData2 = sheet.getRange(rowIndex, 1, 1, 16).getValues()[0];
      var studentName = (rowData2[COLS.FIRST_NAME - 1] || '') + ' ' + (rowData2[COLS.LAST_NAME - 1] || '');
      var studentCid  = String(rowData2[COLS.CID - 1] || '');
      var rating2     = rowData2[COLS.TRAINING_TYPE - 1] || '';
      try { sendTrainerNotificationEmail(trainerCid, trainerName, studentName, studentCid, rating2); }
      catch(e) { logError('trainerEmailInSend', e); }
    }

    return jsonResponse({ success: true });
  } catch(e) {
    logError('handleSendTrainerAssignment', e);
    return jsonResponse({ success: false, error: e.message });
  }
}


// ─── QUEUE MOVE / REORDER ─────────────────────────────────────────────────────

function handleMoveEntry(data) {
  try {
    var fromSheet = getSheet(data.fromTab);
    var toSheet   = getSheet(data.toTab);
    var rowIndex  = data.rowIndex;
    var numCols   = Math.max(fromSheet.getLastColumn(), 16);
    var rowData   = fromSheet.getRange(rowIndex, 1, 1, numCols).getValues()[0];
    while (rowData.length < 16) rowData.push('');
    rowData[COLS.QUEUE_TYPE - 1] = data.toTab;
    rowData[COLS.UPDATED_AT - 1] = new Date().toISOString();
    toSheet.appendRow(rowData);
    var newRow = toSheet.getLastRow();
    rowData[COLS.POSITION - 1] = getNextPosition(toSheet) - 1; // already appended
    renumberQueue(toSheet);
    fromSheet.deleteRow(rowIndex);
    renumberQueue(fromSheet);
    logAction('MOVE_ENTRY', 'Moved row from ' + data.fromTab + ' to ' + data.toTab);
    return jsonResponse({ success: true });
  } catch(e) {
    logError('handleMoveEntry', e);
    return jsonResponse({ success: false, error: e.message });
  }
}

function handleReorderQueue(data) {
  try {
    var sheet = getSheet(resolveTabName(data.tabName));
    var rows  = sheet.getDataRange().getValues();
    var header = rows[0];
    var body   = rows.slice(1);
    // data.order is array of 0-based indices into body
    var reordered = (data.order || []).map(function(i) { return body[i]; });
    sheet.clearContents();
    sheet.appendRow(header);
    reordered.forEach(function(row) { sheet.appendRow(row); });
    renumberQueue(sheet);
    logAction('REORDER_QUEUE', 'Reordered ' + data.tabName);
    return jsonResponse({ success: true });
  } catch(e) {
    logError('handleReorderQueue', e);
    return jsonResponse({ success: false, error: e.message });
  }
}


// ─── PERMANENT DELETE (ARCHIVED / DENIED) ────────────────────────────────────

function handlePermanentDeleteArchived(data) {
  try {
    var sheet = getSheet(TABS.ARCHIVED);
    sheet.deleteRow(data.rowIndex);
    logAction('PERM_DELETE_ARCHIVED', 'Deleted archived row ' + data.rowIndex);
    return jsonResponse({ success: true });
  } catch(e) {
    logError('handlePermanentDeleteArchived', e);
    return jsonResponse({ success: false, error: e.message });
  }
}

function handlePermanentDeleteDenied(data) {
  try {
    var sheet = getSheet(TABS.DENIED);
    sheet.deleteRow(data.rowIndex);
    logAction('PERM_DELETE_DENIED', 'Deleted denied row ' + data.rowIndex);
    return jsonResponse({ success: true });
  } catch(e) {
    logError('handlePermanentDeleteDenied', e);
    return jsonResponse({ success: false, error: e.message });
  }
}


// ─── ADD COMPLETED RECORD MANUALLY ───────────────────────────────────────────

function handleAddCompletedRecord(data) {
  try {
    var sheet = getSheet(TABS.ARCHIVED);
    var row = [
      data.submitted    || new Date().toISOString(),
      data.position     || '',
      data.status       || 'Completed',
      data.firstName    || '',
      data.lastName     || '',
      data.cid          || '',
      data.trainingType || '',
      data.email        || '',
      data.notes        || '',
      new Date().toISOString(),
      data.queueType    || '',
      '',
      data.examGrade    || '',
      data.zmaExamGrade || '',
      data.assignedTrainer   || '',
      data.trainerAssignedAt || ''
    ];
    sheet.appendRow(row);
    logAction('ADD_COMPLETED', 'Manually added completed record for CID ' + data.cid);
    return jsonResponse({ success: true });
  } catch(e) {
    logError('handleAddCompletedRecord', e);
    return jsonResponse({ success: false, error: e.message });
  }
}


// ─── EMAIL HTML BUILDER ───────────────────────────────────────────────────────

function buildEmailHtml(headerTitle, headerSubtitle, bodyHtml) {
  return `<!DOCTYPE html>
<html>
<head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1.0"></head>
<body style="margin:0;padding:0;background-color:#f0f4f8;font-family:'Helvetica Neue',Arial,sans-serif;">
  <table width="100%" cellpadding="0" cellspacing="0" style="background-color:#f0f4f8;padding:40px 0;">
    <tr><td align="center">
      <table width="600" cellpadding="0" cellspacing="0" style="background:#ffffff;border-radius:8px;overflow:hidden;box-shadow:0 2px 8px rgba(0,0,0,0.08);">
        <!-- Header -->
        <tr><td style="background:#0d2340;padding:28px 40px;">
          <p style="margin:0;color:#7eb3e8;font-size:11px;font-weight:700;letter-spacing:2px;text-transform:uppercase;">${headerTitle}</p>
          <p style="margin:8px 0 0;color:#ffffff;font-size:22px;font-weight:700;">${headerSubtitle}</p>
        </td></tr>
        <!-- Body -->
        <tr><td style="padding:36px 40px;color:#2d3748;font-size:15px;line-height:1.7;">
          ${bodyHtml}
        </td></tr>
        <!-- Footer -->
        <tr><td style="background:#f7fafc;padding:20px 40px;border-top:1px solid #e2e8f0;">
          <p style="margin:0;color:#718096;font-size:12px;">Miami ARTCC Training Department &nbsp;|&nbsp; <a href="mailto:ta@zmaartcc.net" style="color:#4a90d9;">ta@zmaartcc.net</a></p>
        </td></tr>
      </table>
    </td></tr>
  </table>
</body>
</html>`;
}


// ─── TRAINING STATISTICS ──────────────────────────────────────────────────────

function getTrainingStatsPublic() {
  try {
    var allRows = [];
    [TABS.HOME, TABS.VISITOR].forEach(function(tab) {
      var d = getSheetData(tab);
      if (d.success) allRows = allRows.concat(d.rows);
    });
    var archived = getSheetData(TABS.ARCHIVED);
    var completedRows = archived.success ? archived.rows : [];

    var waitByRating = {};
    allRows.concat(completedRows).forEach(function(r) {
      if (!r.timestamp || !r.trainerAssignedAt) return;
      var joined   = new Date(r.timestamp);
      var assigned = new Date(r.trainerAssignedAt);
      var days = Math.round((assigned - joined) / 86400000);
      if (isNaN(days) || days < 0) return;
      var key = (r.rating || 'Unknown').trim();
      if (!waitByRating[key]) waitByRating[key] = [];
      waitByRating[key].push(days);
    });

    var trainByRating = {};
    var completionsByTrainer = {};
    completedRows.forEach(function(r) {
      if (r.assignedTrainer && r.assignedTrainer.trim()) {
        var t = r.assignedTrainer.trim();
        completionsByTrainer[t] = (completionsByTrainer[t] || 0) + 1;
      }
      if (r.trainerAssignedAt && r.updatedAt) {
        var start = new Date(r.trainerAssignedAt);
        var end   = new Date(r.updatedAt);
        var days  = Math.round((end - start) / 86400000);
        if (!isNaN(days) && days >= 0) {
          var key = (r.rating || 'Unknown').trim();
          if (!trainByRating[key]) trainByRating[key] = [];
          trainByRating[key].push(days);
        }
      }
    });

    function avg(arr) {
      if (!arr || !arr.length) return null;
      return Math.round(arr.reduce(function(a,b){return a+b;},0) / arr.length);
    }

    var waitStats    = Object.keys(waitByRating).sort().map(function(k){ return { rating: k, avgDays: avg(waitByRating[k]), count: waitByRating[k].length }; });
    var trainStats   = Object.keys(trainByRating).sort().map(function(k){ return { rating: k, avgDays: avg(trainByRating[k]), count: trainByRating[k].length }; });
    var trainerStats = Object.keys(completionsByTrainer).sort().map(function(k){ return { trainer: k, completions: completionsByTrainer[k] }; })
                       .sort(function(a,b){ return b.completions - a.completions; });

    return { success: true, waitStats: waitStats, trainStats: trainStats, trainerStats: trainerStats };
  } catch(e) {
    logError('getTrainingStatsPublic', e);
    return { success: false, error: e.message };
  }
}


// ─── EXAM REQUESTS ────────────────────────────────────────────────────────────

const EXAM_PASSWORDS = {
  'TPA GND/DEL': 'pNqu64',
  'FLL GND/DEL': 'poJHH4246',
  'TPA TWR':     '28dhG',
  'FLL TWR':     'qnd53K',
  'TPA TRACON':  'dg;BN43'
};

// Called from ExamRequestForm via google.script.run

function getExamRequestSheet() {
  var ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  var sheet = ss.getSheetByName('Exam Requests');
  if (!sheet) {
    sheet = ss.insertSheet('Exam Requests');
    var headers = ['Timestamp','First Name','Last Name','CID','Exams','Status','Email','Rating'];
    var hRange = sheet.getRange(1, 1, 1, headers.length);
    hRange.setValues([headers]);
    hRange.setFontWeight('bold');
    hRange.setBackground('#1565c0');
    hRange.setFontColor('#ffffff');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function handleExamRequestFromForm(data) {
  try {
    var sheet = getExamRequestSheet();

    // Look up email + rating via VATUSA at submission time — store so we never re-fetch
    var email = '';
    var ratingShort = '';
    try {
      var member = fetchVatusaMember(data.cid);
      if (member) {
        email = member.email || '';
        var ratingNum = parseInt(member.rating) || 0;
        var ratingMap = {0:'OBS',1:'S1',2:'S2',3:'S3',4:'C1',5:'C1',6:'C3',7:'C3',8:'I1',9:'I1',10:'I3',11:'I3',12:'SUP',13:'ADM'};
        ratingShort = ratingMap[ratingNum] || String(ratingNum);
      }
    } catch(e) { logError('examRequest_vatusa', e); }

    sheet.appendRow([
      new Date().toISOString(),
      data.firstName || '',
      data.lastName  || '',
      String(data.cid || ''),
      data.exams     || '',
      'Pending',
      email,
      ratingShort
    ]);

    return { success: true };
  } catch(e) {
    logError('handleExamRequestFromForm', e);
    return { success: false, error: e.message };
  }
}

function getExamRequestsPublic() {
  try {
    var sheet = getExamRequestSheet();
    var lastRow = sheet.getLastRow();
    if (lastRow <= 1) return { success: true, rows: [] };

    // Read all at once — no per-row API calls
    var values = sheet.getRange(2, 1, lastRow - 1, 8).getValues();
    var rows = [];
    for (var i = 0; i < values.length; i++) {
      var row = values[i];
      var status = String(row[5] || '').trim();
      if (status !== 'Pending') continue;
      rows.push({
        rowIndex:    i + 2,
        timestamp:   row[0] ? String(row[0]) : '',
        firstName:   String(row[1] || ''),
        lastName:    String(row[2] || ''),
        cid:         String(row[3] || ''),
        exams:       String(row[4] || ''),
        status:      status,
        email:       String(row[6] || ''),
        ratingShort: String(row[7] || '?')
      });
    }
    return { success: true, rows: rows };
  } catch(e) {
    logError('getExamRequestsPublic', e);
    return { success: false, rows: [], error: e.message };
  }
}

function handleExamRequestAction(data) {
  try {
    var sheet = getExamRequestSheet();
    var action = data.examAction || data.action; // 'approve' or 'deny'

    // Mark row as actioned
    sheet.getRange(data.rowIndex, 6).setValue(action === 'approve' ? 'Approved' : 'Denied');

    // Look up email if not provided
    var email = data.email || '';
    if (!email) {
      try {
        var member = fetchVatusaMember(data.cid);
        if (member) email = member.email || '';
      } catch(e) {}
    }
    if (!email) return { success: false, error: 'No email found for CID ' + data.cid };

    if (action === 'approve') {
      var exams = (data.exams || '').split(',').map(function(e){ return e.trim(); }).filter(Boolean);
      var codeLines = exams.map(function(exam) {
        var pw = EXAM_PASSWORDS[exam] || '(contact TA)';
        return '<tr><td style="padding:8px 16px;font-weight:600;color:#2a3a4a;">' + exam + '</td>' +
               '<td style="padding:8px 16px;font-family:monospace;font-size:16px;color:#0077b6;letter-spacing:0.05em;">' + pw + '</td></tr>';
      }).join('');
      var codePlain = exams.map(function(exam){ return exam + ': ' + (EXAM_PASSWORDS[exam]||'?'); }).join('\n');

      var bodyHtml =
        '<p style="margin:0 0 16px 0;font-size:16px;color:#1a2a3a;">Hello <strong>' + data.firstName + '</strong>,</p>' +
        '<p style="margin:0 0 16px 0;font-size:15px;color:#2a3a4a;line-height:1.7;">Your exam request has been approved. Below are your exam code(s):</p>' +
        '<table style="width:100%;border-collapse:collapse;margin:0 0 20px 0;background:#f7fafc;border-radius:8px;overflow:hidden;">' +
        '<tr style="background:#e8f0fe;"><th style="padding:10px 16px;text-align:left;color:#0a3d6b;font-size:13px;letter-spacing:0.08em;text-transform:uppercase;">Exam</th>' +
        '<th style="padding:10px 16px;text-align:left;color:#0a3d6b;font-size:13px;letter-spacing:0.08em;text-transform:uppercase;">Code</th></tr>' +
        codeLines + '</table>' +
        '<p style="margin:0 0 16px 0;font-size:15px;color:#2a3a4a;line-height:1.7;">You will have <strong>30 days</strong> from this email to complete the exam(s). If you have any questions, contact <a href="mailto:ta@zmaartcc.net" style="color:#0ea5c8;">ta@zmaartcc.net</a>.</p>' +
        '<p style="margin:0;font-size:15px;color:#2a3a4a;">Good luck!</p>';

      var plain = 'Hello ' + data.firstName + ',\n\nYour exam request has been approved.\n\n' + codePlain +
        '\n\nYou have 30 days from this email to complete the exam(s). Questions? Email ta@zmaartcc.net.\n\nGood luck!\nZMA Training Team';

      MailApp.sendEmail({
        to: email,
        subject: '[ZMA] Exam Code(s) – ' + data.exams,
        body: plain,
        htmlBody: buildEmailHtml('Exam Request', 'REQUEST APPROVED', bodyHtml),
        name: 'ZMA Training Team'
      });

    } else {
      // Denial email
      var denyBody =
        '<p style="margin:0 0 16px 0;font-size:16px;color:#1a2a3a;">Hello <strong>' + data.firstName + '</strong>,</p>' +
        '<p style="margin:0 0 16px 0;font-size:15px;color:#2a3a4a;line-height:1.7;">Thank you for submitting your exam request. Unfortunately, your request for <strong>' + data.exams + '</strong> has not been approved at this time.</p>' +
        '<p style="margin:0;font-size:15px;color:#2a3a4a;line-height:1.7;">If you believe this is in error or have questions, please contact <a href="mailto:ta@zmaartcc.net" style="color:#0ea5c8;">ta@zmaartcc.net</a>.</p>';

      var denyPlain = 'Hello ' + data.firstName + ',\n\nYour exam request for ' + data.exams +
        ' has not been approved at this time. If you have questions, contact ta@zmaartcc.net.\n\nZMA Training Team';

      MailApp.sendEmail({
        to: email,
        subject: '[ZMA] Exam Request Update',
        body: denyPlain,
        htmlBody: buildEmailHtml('Exam Request', 'REQUEST NOT APPROVED', denyBody),
        name: 'ZMA Training Team'
      });
    }

    logAction('EXAM_REQUEST_' + action.toUpperCase(), data.firstName + ' ' + data.lastName + ' CID:' + data.cid + ' Exams:' + data.exams);
    return { success: true };
  } catch(e) {
    logError('handleExamRequestAction', e);
    return { success: false, error: e.message };
  }
}


function getExamHistoryPublic() {
  try {
    var sheet = getExamRequestSheet();
    var lastRow = sheet.getLastRow();
    if (lastRow <= 1) return { success: true, rows: [] };

    var values = sheet.getRange(2, 1, lastRow - 1, 8).getValues();
    var rows = [];
    for (var i = 0; i < values.length; i++) {
      var status = String(values[i][5] || '').trim();
      if (status !== 'Approved' && status !== 'Denied') continue;
      rows.push({
        rowIndex:   i + 2,
        timestamp:  values[i][0] ? String(values[i][0]) : '',
        firstName:  String(values[i][1] || ''),
        lastName:   String(values[i][2] || ''),
        cid:        String(values[i][3] || ''),
        exams:      String(values[i][4] || ''),
        status:     status,
        email:      String(values[i][6] || ''),
        ratingShort: String(values[i][7] || '')
      });
    }
    // Most recent first
    rows.reverse();
    return { success: true, rows: rows };
  } catch(e) {
    logError('getExamHistoryPublic', e);
    return { success: false, rows: [], error: e.message };
  }
}


// ╔══════════════════════════════════════════════════════════════════════════════╗
// ║                        ZMA EXAM CENTER — BACKEND                           ║
// ║  Tabs: Exams | Exam Sections | Questions | Question Banks | Exam Attempts   ║
// ╚══════════════════════════════════════════════════════════════════════════════╝

// ─── EXAM TAB NAMES & COLUMN MAPS ────────────────────────────────────────────

const EXAM_TABS = {
  EXAMS:    'Exams',
  SECTIONS: 'Exam Sections',
  QUESTIONS:'Questions',
  BANKS:    'Question Banks',
  ATTEMPTS: 'Exam Attempts',
};

// ─── SHEET SETUP ─────────────────────────────────────────────────────────────

function setupExamSheets() {
  var ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);

  function ensureExamSheet(name, headers) {
    var sheet = ss.getSheetByName(name);
    if (!sheet) {
      sheet = ss.insertSheet(name);
      var r = sheet.getRange(1, 1, 1, headers.length);
      r.setValues([headers]);
      r.setFontWeight('bold');
      r.setBackground('#0d3b6e');
      r.setFontColor('#ffffff');
      sheet.setFrozenRows(1);
    }
    return sheet;
  }

  ensureExamSheet(EXAM_TABS.EXAMS, [
    'ExamID','Title','Rating','Status','Password','PassingPct',
    'TotalPoints','TimeLimitMins','MaxAttempts','CooldownHours',
    'ShowFeedback','ShowCorrectAnswers','Version','CreatedAt','UpdatedAt','Description'
  ]);

  ensureExamSheet(EXAM_TABS.SECTIONS, [
    'SectionID','ExamID','Title','SectionOrder','Mode',
    'BankID','QuestionCount','Description'
  ]);

  ensureExamSheet(EXAM_TABS.QUESTIONS, [
    'QuestionID','ExamID','SectionID','BankID','Type','QuestionText',
    'ImageURL','ChoicesJSON','CorrectAnswerJSON','MatchingPairsJSON',
    'Points','Explanation','OrderIndex','CreatedAt','UpdatedAt'
  ]);

  ensureExamSheet(EXAM_TABS.BANKS, [
    'BankID','Name','Rating','Description','CreatedAt','UpdatedAt'
  ]);

  ensureExamSheet(EXAM_TABS.ATTEMPTS, [
    'AttemptID','ExamID','ExamVersion','CID','FirstName','LastName',
    'StartedAt','SubmittedAt','Score','TotalPoints','Pct','Passed',
    'AnswersJSON','FeedbackJSON','IsPreview'
  ]);

  logAction('EXAM_SETUP', 'Exam Center sheets created/verified');
  return { success: true };
}

// ─── ID GENERATOR ────────────────────────────────────────────────────────────

function generateId(prefix) {
  return prefix + '_' + new Date().getTime() + '_' + Math.floor(Math.random() * 10000);
}

// ─── EXAM CRUD ────────────────────────────────────────────────────────────────

function createExam(data) {
  try {
    var sheet = SpreadsheetApp.openById(CONFIG.SHEET_ID).getSheetByName(EXAM_TABS.EXAMS);
    if (!sheet) return { success: false, error: 'Exams sheet not found. Run setupExamSheets() first.' };

    var examId = generateId('EX');
    var now = new Date().toISOString();
    sheet.appendRow([
      examId,
      data.title        || 'Untitled Exam',
      data.rating       || '',           // S1, S2, S3, C1, T1, etc.
      'Draft',                           // always start as Draft
      data.password     || '',
      data.passingPct   || 70,
      data.totalPoints  || 100,
      data.timeLimitMins|| '',           // blank = untimed
      data.maxAttempts  || 3,
      data.cooldownHours|| 24,
      data.showFeedback       !== false ? 'Y' : 'N',
      data.showCorrectAnswers !== false ? 'Y' : 'N',
      1,                                 // version
      now, now,
      data.description  || ''
    ]);
    logAction('CREATE_EXAM', examId + ' — ' + data.title);
    return { success: true, examId: examId };
  } catch(e) {
    logError('createExam', e);
    return { success: false, error: e.message };
  }
}

function updateExam(data) {
  try {
    var sheet = SpreadsheetApp.openById(CONFIG.SHEET_ID).getSheetByName(EXAM_TABS.EXAMS);
    var rows = sheet.getDataRange().getValues();
    for (var i = 1; i < rows.length; i++) {
      if (String(rows[i][0]) === String(data.examId)) {
        var now = new Date().toISOString();
        // Bump version if publishing
        var version = parseInt(rows[i][12]) || 1;
        if (data.status === 'Published' && rows[i][3] !== 'Published') version++;
        sheet.getRange(i+1, 1, 1, 16).setValues([[
          rows[i][0],                                     // ExamID unchanged
          data.title          || rows[i][1],
          data.rating         || rows[i][2],
          data.status         || rows[i][3],
          data.password       !== undefined ? data.password       : rows[i][4],
          data.passingPct     !== undefined ? data.passingPct     : rows[i][5],
          data.totalPoints    !== undefined ? data.totalPoints    : rows[i][6],
          data.timeLimitMins  !== undefined ? data.timeLimitMins  : rows[i][7],
          data.maxAttempts    !== undefined ? data.maxAttempts    : rows[i][8],
          data.cooldownHours  !== undefined ? data.cooldownHours  : rows[i][9],
          data.showFeedback       !== undefined ? (data.showFeedback       ? 'Y':'N') : rows[i][10],
          data.showCorrectAnswers !== undefined ? (data.showCorrectAnswers ? 'Y':'N') : rows[i][11],
          version,
          rows[i][13],                                    // CreatedAt unchanged
          now,
          data.description    !== undefined ? data.description    : rows[i][15],
        ]]);
        logAction('UPDATE_EXAM', data.examId);
        return { success: true, version: version };
      }
    }
    return { success: false, error: 'Exam not found: ' + data.examId };
  } catch(e) {
    logError('updateExam', e);
    return { success: false, error: e.message };
  }
}

function getExams(filterStatus) {
  try {
    var sheet = SpreadsheetApp.openById(CONFIG.SHEET_ID).getSheetByName(EXAM_TABS.EXAMS);
    if (!sheet) return { success: false, exams: [] };
    var rows = sheet.getDataRange().getValues();
    var exams = [];
    for (var i = 1; i < rows.length; i++) {
      var r = rows[i];
      if (!r[0]) continue;
      if (filterStatus && r[3] !== filterStatus) continue;
      exams.push(examRowToObj(r));
    }
    return { success: true, exams: exams };
  } catch(e) {
    logError('getExams', e);
    return { success: false, exams: [] };
  }
}

function getExamById(examId) {
  try {
    var sheet = SpreadsheetApp.openById(CONFIG.SHEET_ID).getSheetByName(EXAM_TABS.EXAMS);
    var rows = sheet.getDataRange().getValues();
    for (var i = 1; i < rows.length; i++) {
      if (String(rows[i][0]) === String(examId)) {
        return { success: true, exam: examRowToObj(rows[i]) };
      }
    }
    return { success: false, error: 'Not found' };
  } catch(e) {
    logError('getExamById', e);
    return { success: false, error: e.message };
  }
}

function examRowToObj(r) {
  return {
    examId:            String(r[0]),
    title:             String(r[1] || ''),
    rating:            String(r[2] || ''),
    status:            String(r[3] || 'Draft'),
    password:          String(r[4] || ''),
    passingPct:        Number(r[5]) || 70,
    totalPoints:       Number(r[6]) || 100,
    timeLimitMins:     r[7] !== '' ? Number(r[7]) : null,
    maxAttempts:       Number(r[8]) || 3,
    cooldownHours:     Number(r[9]) || 24,
    showFeedback:      r[10] === 'Y',
    showCorrectAnswers:r[11] === 'Y',
    version:           Number(r[12]) || 1,
    createdAt:         String(r[13] || ''),
    updatedAt:         String(r[14] || ''),
    description:       String(r[15] || ''),
  };
}

function publishExam(examId) {
  return updateExam({ examId: examId, status: 'Published' });
}

function archiveExam(examId) {
  return updateExam({ examId: examId, status: 'Archived' });
}

function deleteExam(examId) {
  try {
    var ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
    // Remove from Exams sheet
    var sheet = ss.getSheetByName(EXAM_TABS.EXAMS);
    var rows = sheet.getDataRange().getValues();
    for (var i = rows.length - 1; i >= 1; i--) {
      if (String(rows[i][0]) === String(examId)) { sheet.deleteRow(i+1); break; }
    }
    // Cascade delete sections and questions
    _deleteRowsWhere(ss.getSheetByName(EXAM_TABS.SECTIONS),  1, examId);
    _deleteRowsWhere(ss.getSheetByName(EXAM_TABS.QUESTIONS),  1, examId);
    logAction('DELETE_EXAM', examId);
    return { success: true };
  } catch(e) {
    logError('deleteExam', e);
    return { success: false, error: e.message };
  }
}

function _deleteRowsWhere(sheet, colIndex, matchValue) {
  if (!sheet) return;
  var rows = sheet.getDataRange().getValues();
  for (var i = rows.length - 1; i >= 1; i--) {
    if (String(rows[i][colIndex]) === String(matchValue)) sheet.deleteRow(i+1);
  }
}

// ─── QUESTION BANKS ───────────────────────────────────────────────────────────

function createBank(data) {
  try {
    var sheet = SpreadsheetApp.openById(CONFIG.SHEET_ID).getSheetByName(EXAM_TABS.BANKS);
    if (!sheet) return { success: false, error: 'Run setupExamSheets() first.' };
    var bankId = generateId('BK');
    var now = new Date().toISOString();
    sheet.appendRow([bankId, data.name || 'Unnamed Bank', data.rating || '', data.description || '', now, now]);
    logAction('CREATE_BANK', bankId + ' — ' + data.name);
    return { success: true, bankId: bankId };
  } catch(e) {
    logError('createBank', e);
    return { success: false, error: e.message };
  }
}

function getBanks() {
  try {
    var sheet = SpreadsheetApp.openById(CONFIG.SHEET_ID).getSheetByName(EXAM_TABS.BANKS);
    if (!sheet) return { success: true, banks: [] };
    var rows = sheet.getDataRange().getValues();
    var banks = [];
    for (var i = 1; i < rows.length; i++) {
      if (!rows[i][0]) continue;
      banks.push({ bankId: String(rows[i][0]), name: String(rows[i][1]), rating: String(rows[i][2]), description: String(rows[i][3]) });
    }
    return { success: true, banks: banks };
  } catch(e) {
    logError('getBanks', e);
    return { success: false, banks: [] };
  }
}

function deleteBank(bankId) {
  try {
    var ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
    _deleteRowsWhere(ss.getSheetByName(EXAM_TABS.BANKS), 0, bankId);
    // Remove any questions that belong to this bank only (no exam)
    var qSheet = ss.getSheetByName(EXAM_TABS.QUESTIONS);
    var qRows = qSheet.getDataRange().getValues();
    for (var i = qRows.length - 1; i >= 1; i--) {
      if (String(qRows[i][3]) === String(bankId) && !qRows[i][1]) qSheet.deleteRow(i+1);
    }
    logAction('DELETE_BANK', bankId);
    return { success: true };
  } catch(e) {
    logError('deleteBank', e);
    return { success: false, error: e.message };
  }
}

// ─── EXAM SECTIONS ────────────────────────────────────────────────────────────

function createSection(data) {
  try {
    var sheet = SpreadsheetApp.openById(CONFIG.SHEET_ID).getSheetByName(EXAM_TABS.SECTIONS);
    if (!sheet) return { success: false, error: 'Run setupExamSheets() first.' };
    var sectionId = generateId('SEC');
    sheet.appendRow([
      sectionId,
      data.examId       || '',
      data.title        || 'Section',
      data.sectionOrder || 1,
      data.mode         || 'fixed',     // 'fixed' or 'random'
      data.bankId       || '',          // used when mode='random'
      data.questionCount|| '',          // how many to pull (random mode)
      data.description  || ''
    ]);
    logAction('CREATE_SECTION', sectionId + ' for exam ' + data.examId);
    return { success: true, sectionId: sectionId };
  } catch(e) {
    logError('createSection', e);
    return { success: false, error: e.message };
  }
}

function updateSection(data) {
  try {
    var sheet = SpreadsheetApp.openById(CONFIG.SHEET_ID).getSheetByName(EXAM_TABS.SECTIONS);
    var rows = sheet.getDataRange().getValues();
    for (var i = 1; i < rows.length; i++) {
      if (String(rows[i][0]) === String(data.sectionId)) {
        sheet.getRange(i+1, 1, 1, 8).setValues([[
          rows[i][0],
          rows[i][1],
          data.title         || rows[i][2],
          data.sectionOrder  !== undefined ? data.sectionOrder  : rows[i][3],
          data.mode          || rows[i][4],
          data.bankId        !== undefined ? data.bankId        : rows[i][5],
          data.questionCount !== undefined ? data.questionCount : rows[i][6],
          data.description   !== undefined ? data.description   : rows[i][7],
        ]]);
        return { success: true };
      }
    }
    return { success: false, error: 'Section not found' };
  } catch(e) {
    logError('updateSection', e);
    return { success: false, error: e.message };
  }
}

function getSectionsByExam(examId) {
  try {
    var sheet = SpreadsheetApp.openById(CONFIG.SHEET_ID).getSheetByName(EXAM_TABS.SECTIONS);
    if (!sheet) return { success: true, sections: [] };
    var rows = sheet.getDataRange().getValues();
    var sections = [];
    for (var i = 1; i < rows.length; i++) {
      if (String(rows[i][1]) !== String(examId)) continue;
      sections.push({
        sectionId:     String(rows[i][0]),
        examId:        String(rows[i][1]),
        title:         String(rows[i][2]),
        sectionOrder:  Number(rows[i][3]),
        mode:          String(rows[i][4]),
        bankId:        String(rows[i][5] || ''),
        questionCount: rows[i][6] !== '' ? Number(rows[i][6]) : null,
        description:   String(rows[i][7] || ''),
      });
    }
    sections.sort(function(a,b){ return a.sectionOrder - b.sectionOrder; });
    return { success: true, sections: sections };
  } catch(e) {
    logError('getSectionsByExam', e);
    return { success: false, sections: [] };
  }
}

function deleteSection(sectionId) {
  try {
    var ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
    _deleteRowsWhere(ss.getSheetByName(EXAM_TABS.SECTIONS), 0, sectionId);
    _deleteRowsWhere(ss.getSheetByName(EXAM_TABS.QUESTIONS), 2, sectionId);
    logAction('DELETE_SECTION', sectionId);
    return { success: true };
  } catch(e) {
    logError('deleteSection', e);
    return { success: false, error: e.message };
  }
}

// ─── QUESTIONS ────────────────────────────────────────────────────────────────

/*
  Question types:
    mc-single   — multiple choice, one correct answer
    mc-multi    — multiple choice, multiple correct answers
    true-false  — true or false
    short-answer— free text, graded manually or by keyword match
    matching    — match left items to right items
    fill-blank  — sentence with blank(s), student fills in

  choicesJSON (mc-single / mc-multi):  ["Choice A","Choice B","Choice C","Choice D"]
  correctAnswerJSON:
    mc-single:   "Choice A"
    mc-multi:    ["Choice A","Choice C"]
    true-false:  "True" or "False"
    short-answer:["keyword1","keyword2"]   (accepted keywords)
    fill-blank:  ["answer1","answer2"]     (one per blank, in order)
  matchingPairsJSON (matching only):
    [{"left":"Item 1","right":"Match A"},{"left":"Item 2","right":"Match B"}]
*/

function createQuestion(data) {
  try {
    var sheet = SpreadsheetApp.openById(CONFIG.SHEET_ID).getSheetByName(EXAM_TABS.QUESTIONS);
    if (!sheet) return { success: false, error: 'Run setupExamSheets() first.' };
    var qId = generateId('Q');
    var now = new Date().toISOString();
    sheet.appendRow([
      qId,
      data.examId         || '',
      data.sectionId      || '',
      data.bankId         || '',
      data.type           || 'mc-single',
      data.questionText   || '',
      data.imageURL       || '',
      data.choices        ? JSON.stringify(data.choices)       : '',
      data.correctAnswer  ? JSON.stringify(data.correctAnswer) : '',
      data.matchingPairs  ? JSON.stringify(data.matchingPairs) : '',
      data.points         || 1,
      data.explanation    || '',
      data.orderIndex     || 1,
      now, now
    ]);
    logAction('CREATE_QUESTION', qId + ' type:' + data.type + ' exam:' + (data.examId||data.bankId));
    return { success: true, questionId: qId };
  } catch(e) {
    logError('createQuestion', e);
    return { success: false, error: e.message };
  }
}

function updateQuestion(data) {
  try {
    var sheet = SpreadsheetApp.openById(CONFIG.SHEET_ID).getSheetByName(EXAM_TABS.QUESTIONS);
    var rows = sheet.getDataRange().getValues();
    for (var i = 1; i < rows.length; i++) {
      if (String(rows[i][0]) !== String(data.questionId)) continue;
      var now = new Date().toISOString();
      sheet.getRange(i+1, 1, 1, 15).setValues([[
        rows[i][0],
        rows[i][1],
        rows[i][2],
        rows[i][3],
        data.type          || rows[i][4],
        data.questionText  !== undefined ? data.questionText  : rows[i][5],
        data.imageURL      !== undefined ? data.imageURL      : rows[i][6],
        data.choices       !== undefined ? JSON.stringify(data.choices)      : rows[i][7],
        data.correctAnswer !== undefined ? JSON.stringify(data.correctAnswer): rows[i][8],
        data.matchingPairs !== undefined ? JSON.stringify(data.matchingPairs): rows[i][9],
        data.points        !== undefined ? data.points        : rows[i][10],
        data.explanation   !== undefined ? data.explanation   : rows[i][11],
        data.orderIndex    !== undefined ? data.orderIndex    : rows[i][12],
        rows[i][13],
        now
      ]]);
      logAction('UPDATE_QUESTION', data.questionId);
      return { success: true };
    }
    return { success: false, error: 'Question not found' };
  } catch(e) {
    logError('updateQuestion', e);
    return { success: false, error: e.message };
  }
}

function deleteQuestion(questionId) {
  try {
    var sheet = SpreadsheetApp.openById(CONFIG.SHEET_ID).getSheetByName(EXAM_TABS.QUESTIONS);
    _deleteRowsWhere(sheet, 0, questionId);
    logAction('DELETE_QUESTION', questionId);
    return { success: true };
  } catch(e) {
    logError('deleteQuestion', e);
    return { success: false, error: e.message };
  }
}

function getQuestionsByExam(examId) {
  try {
    var sheet = SpreadsheetApp.openById(CONFIG.SHEET_ID).getSheetByName(EXAM_TABS.QUESTIONS);
    if (!sheet) return { success: true, questions: [] };
    var rows = sheet.getDataRange().getValues();
    var questions = [];
    for (var i = 1; i < rows.length; i++) {
      if (String(rows[i][1]) !== String(examId)) continue;
      questions.push(questionRowToObj(rows[i]));
    }
    questions.sort(function(a,b){ return a.orderIndex - b.orderIndex; });
    return { success: true, questions: questions };
  } catch(e) {
    logError('getQuestionsByExam', e);
    return { success: false, questions: [] };
  }
}

function getQuestionsByBank(bankId) {
  try {
    var sheet = SpreadsheetApp.openById(CONFIG.SHEET_ID).getSheetByName(EXAM_TABS.QUESTIONS);
    if (!sheet) return { success: true, questions: [] };
    var rows = sheet.getDataRange().getValues();
    var questions = [];
    for (var i = 1; i < rows.length; i++) {
      if (String(rows[i][3]) !== String(bankId)) continue;
      questions.push(questionRowToObj(rows[i]));
    }
    return { success: true, questions: questions };
  } catch(e) {
    logError('getQuestionsByBank', e);
    return { success: false, questions: [] };
  }
}

function questionRowToObj(r) {
  function safeParse(val) {
    try { return val ? JSON.parse(val) : null; } catch(e) { return null; }
  }
  return {
    questionId:    String(r[0]),
    examId:        String(r[1] || ''),
    sectionId:     String(r[2] || ''),
    bankId:        String(r[3] || ''),
    type:          String(r[4] || 'mc-single'),
    questionText:  String(r[5] || ''),
    imageURL:      String(r[6] || ''),
    choices:       safeParse(r[7]),
    correctAnswer: safeParse(r[8]),
    matchingPairs: safeParse(r[9]),
    points:        Number(r[10]) || 1,
    explanation:   String(r[11] || ''),
    orderIndex:    Number(r[12]) || 1,
    createdAt:     String(r[13] || ''),
    updatedAt:     String(r[14] || ''),
  };
}

// ─── EXAM DELIVERY (build exam for frontend) ──────────────────────────────────

/*
  buildExamForDelivery(examId, isPreview):
  - Validates password is NOT checked here (done separately via verifyExamPassword)
  - Returns exam metadata + sections with questions in correct order
  - For 'random' sections: pulls questions from the linked bank and shuffles
  - Strips correct answers from the returned payload (shown only after submit)
*/

function verifyExamPassword(examId, password) {
  try {
    var result = getExamById(examId);
    if (!result.success) return { success: false, error: 'Exam not found' };
    var exam = result.exam;
    if (exam.status !== 'Published') return { success: false, error: 'Exam is not available' };
    if (exam.password && exam.password !== password) return { success: false, error: 'Incorrect password' };
    return { success: true };
  } catch(e) {
    logError('verifyExamPassword', e);
    return { success: false, error: e.message };
  }
}

function checkAttemptEligibility(examId, cid) {
  try {
    var exam = getExamById(examId);
    if (!exam.success) return { eligible: false, reason: 'Exam not found' };
    var e = exam.exam;

    var sheet = SpreadsheetApp.openById(CONFIG.SHEET_ID).getSheetByName(EXAM_TABS.ATTEMPTS);
    if (!sheet) return { eligible: true };
    var rows = sheet.getDataRange().getValues();

    var attempts = [];
    for (var i = 1; i < rows.length; i++) {
      if (String(rows[i][1]) === String(examId) &&
          String(rows[i][3]) === String(cid)    &&
          rows[i][18] !== 'Y') {   // not a preview
        attempts.push({ submittedAt: rows[i][7], passed: rows[i][11] === 'Y' });
      }
    }

    // Check max attempts
    if (e.maxAttempts && attempts.length >= e.maxAttempts) {
      var passed = attempts.some(function(a){ return a.passed; });
      if (passed) return { eligible: false, reason: 'You have already passed this exam.' };
      return { eligible: false, reason: 'Maximum attempts (' + e.maxAttempts + ') reached.' };
    }

    // Check cooldown
    if (e.cooldownHours && attempts.length > 0) {
      var lastAttempt = attempts[attempts.length - 1];
      var lastTime = new Date(lastAttempt.submittedAt).getTime();
      var cooldownMs = e.cooldownHours * 3600000;
      var elapsed = Date.now() - lastTime;
      if (elapsed < cooldownMs) {
        var hoursLeft = Math.ceil((cooldownMs - elapsed) / 3600000);
        return { eligible: false, reason: 'Cooldown active. Try again in ' + hoursLeft + ' hour(s).' };
      }
    }

    return { eligible: true, attemptNumber: attempts.length + 1 };
  } catch(e) {
    logError('checkAttemptEligibility', e);
    return { eligible: false, reason: e.message };
  }
}

function buildExamForDelivery(examId, isPreview) {
  try {
    var examResult = getExamById(examId);
    if (!examResult.success) return { success: false, error: 'Exam not found' };
    var exam = examResult.exam;

    var sectionsResult = getSectionsByExam(examId);
    var sections = sectionsResult.sections || [];

    var deliverySections = sections.map(function(sec) {
      var questions = [];

      if (sec.mode === 'random' && sec.bankId) {
        // Pull from bank and shuffle
        var bankQ = getQuestionsByBank(sec.bankId);
        var pool = bankQ.questions || [];
        pool = _shuffle(pool);
        if (sec.questionCount) pool = pool.slice(0, sec.questionCount);
        questions = pool;
      } else {
        // Fixed — get questions assigned to this section
        var allQ = getQuestionsByExam(examId);
        questions = (allQ.questions || []).filter(function(q){ return q.sectionId === sec.sectionId; });
      }

      // Strip correct answers for delivery (unless preview with answers shown)
      var safeQ = questions.map(function(q) {
        var safe = {
          questionId:  q.questionId,
          type:        q.type,
          questionText:q.questionText,
          imageURL:    q.imageURL,
          choices:     q.choices,
          matchingPairs: q.matchingPairs ? q.matchingPairs.map(function(p){ return { left: p.left }; }) : null,
          points:      q.points,
          orderIndex:  q.orderIndex,
        };
        // In preview mode include answers + explanation
        if (isPreview) {
          safe.correctAnswer = q.correctAnswer;
          safe.explanation   = q.explanation;
        }
        return safe;
      });

      return {
        sectionId:    sec.sectionId,
        title:        sec.title,
        sectionOrder: sec.sectionOrder,
        mode:         sec.mode,
        questions:    safeQ,
      };
    });

    // Return safe exam metadata (no password in payload)
    return {
      success: true,
      exam: {
        examId:        exam.examId,
        title:         exam.title,
        rating:        exam.rating,
        description:   exam.description,
        totalPoints:   exam.totalPoints,
        passingPct:    exam.passingPct,
        timeLimitMins: exam.timeLimitMins,
        showFeedback:  exam.showFeedback,
        showCorrectAnswers: exam.showCorrectAnswers,
        version:       exam.version,
      },
      sections: deliverySections,
    };
  } catch(e) {
    logError('buildExamForDelivery', e);
    return { success: false, error: e.message };
  }
}

function _shuffle(arr) {
  var a = arr.slice();
  for (var i = a.length - 1; i > 0; i--) {
    var j = Math.floor(Math.random() * (i + 1));
    var tmp = a[i]; a[i] = a[j]; a[j] = tmp;
  }
  return a;
}

// ─── EXAM SUBMISSION & GRADING ────────────────────────────────────────────────

/*
  submitExamAttempt(data):
  data = {
    examId, cid, firstName, lastName, isPreview,
    startedAt,   (ISO string)
    answers: [   (one per question)
      { questionId, answer }  — answer matches type:
        mc-single:   "Choice A"
        mc-multi:    ["Choice A","Choice C"]
        true-false:  "True"
        short-answer:"free text"
        matching:    [{"left":"Item1","right":"Match A"}, ...]
        fill-blank:  ["word1","word2"]
    ]
  }
*/

function submitExamAttempt(data) {
  try {
    var examResult = getExamById(data.examId);
    if (!examResult.success) return { success: false, error: 'Exam not found' };
    var exam = examResult.exam;

    // Get all questions for this exam (with correct answers)
    var allQ = getQuestionsByExam(data.examId);
    // Also pull bank questions for random sections
    var sectionsResult = getSectionsByExam(data.examId);
    var bankQuestions = [];
    (sectionsResult.sections || []).forEach(function(sec) {
      if (sec.mode === 'random' && sec.bankId) {
        var bq = getQuestionsByBank(sec.bankId);
        bankQuestions = bankQuestions.concat(bq.questions || []);
      }
    });
    var questionMap = {};
    allQ.questions.forEach(function(q){ questionMap[q.questionId] = q; });
    bankQuestions.forEach(function(q){ questionMap[q.questionId] = q; });

    // Grade each answer
    var totalEarned = 0;
    var totalPossible = 0;
    var feedback = [];

    (data.answers || []).forEach(function(ans) {
      var q = questionMap[ans.questionId];
      if (!q) return;
      totalPossible += q.points;
      var correct = gradeAnswer(q, ans.answer);
      var earned  = correct ? q.points : 0;
      totalEarned += earned;
      feedback.push({
        questionId:    q.questionId,
        questionText:  q.questionText,
        type:          q.type,
        studentAnswer: ans.answer,
        correctAnswer: exam.showCorrectAnswers ? q.correctAnswer : undefined,
        explanation:   exam.showFeedback       ? q.explanation   : undefined,
        correct:       correct,
        earned:        earned,
        possible:      q.points,
      });
    });

    var pct    = totalPossible > 0 ? Math.round((totalEarned / totalPossible) * 100) : 0;
    var passed = pct >= exam.passingPct;

    // Save attempt
    if (!data.isPreview) {
      var sheet = SpreadsheetApp.openById(CONFIG.SHEET_ID).getSheetByName(EXAM_TABS.ATTEMPTS);
      if (sheet) {
        var attemptId = generateId('AT');
        sheet.appendRow([
          attemptId,
          data.examId,
          exam.version,
          String(data.cid || ''),
          data.firstName || '',
          data.lastName  || '',
          data.startedAt || '',
          new Date().toISOString(),
          totalEarned,
          totalPossible,
          pct,
          passed ? 'Y' : 'N',
          JSON.stringify(data.answers || []),
          JSON.stringify(feedback),
          data.isPreview ? 'Y' : 'N'
        ]);
        logAction('EXAM_ATTEMPT', 'CID:' + data.cid + ' exam:' + data.examId + ' score:' + pct + '%' + (passed?' PASS':' FAIL'));
      }
    }

    return {
      success:       true,
      score:         totalEarned,
      totalPoints:   totalPossible,
      pct:           pct,
      passed:        passed,
      passingPct:    exam.passingPct,
      feedback:      exam.showFeedback ? feedback : null,
    };
  } catch(e) {
    logError('submitExamAttempt', e);
    return { success: false, error: e.message };
  }
}

function gradeAnswer(question, studentAnswer) {
  var correct = question.correctAnswer;
  if (correct === null || correct === undefined) return false;

  switch(question.type) {
    case 'mc-single':
    case 'true-false':
      return String(studentAnswer).trim().toLowerCase() === String(correct).trim().toLowerCase();

    case 'mc-multi':
      if (!Array.isArray(studentAnswer) || !Array.isArray(correct)) return false;
      var sa = studentAnswer.map(function(x){ return String(x).trim().toLowerCase(); }).sort();
      var ca = correct.map(function(x){ return String(x).trim().toLowerCase(); }).sort();
      return JSON.stringify(sa) === JSON.stringify(ca);

    case 'short-answer':
      // Pass if student answer contains any accepted keyword
      var keywords = Array.isArray(correct) ? correct : [correct];
      var ans = String(studentAnswer).trim().toLowerCase();
      return keywords.some(function(k){ return ans.indexOf(String(k).trim().toLowerCase()) !== -1; });

    case 'fill-blank':
      if (!Array.isArray(correct)) return false;
      var blanks = Array.isArray(studentAnswer) ? studentAnswer : [studentAnswer];
      return correct.every(function(c, idx) {
        return blanks[idx] && String(blanks[idx]).trim().toLowerCase() === String(c).trim().toLowerCase();
      });

    case 'matching':
      if (!Array.isArray(correct) || !Array.isArray(studentAnswer)) return false;
      return correct.every(function(pair) {
        return studentAnswer.some(function(sa) {
          return String(sa.left).trim().toLowerCase()  === String(pair.left).trim().toLowerCase() &&
                 String(sa.right).trim().toLowerCase() === String(pair.right).trim().toLowerCase();
        });
      });

    default:
      return false;
  }
}

// ─── ATTEMPT HISTORY ─────────────────────────────────────────────────────────

function getAttemptsByExam(examId) {
  try {
    var sheet = SpreadsheetApp.openById(CONFIG.SHEET_ID).getSheetByName(EXAM_TABS.ATTEMPTS);
    if (!sheet) return { success: true, attempts: [] };
    var rows = sheet.getDataRange().getValues();
    var attempts = [];
    for (var i = 1; i < rows.length; i++) {
      if (String(rows[i][1]) !== String(examId)) continue;
      attempts.push(attemptRowToObj(rows[i]));
    }
    attempts.sort(function(a,b){ return new Date(b.submittedAt) - new Date(a.submittedAt); });
    return { success: true, attempts: attempts };
  } catch(e) {
    logError('getAttemptsByExam', e);
    return { success: false, attempts: [] };
  }
}

function getAttemptsByCid(cid) {
  try {
    var sheet = SpreadsheetApp.openById(CONFIG.SHEET_ID).getSheetByName(EXAM_TABS.ATTEMPTS);
    if (!sheet) return { success: true, attempts: [] };
    var rows = sheet.getDataRange().getValues();
    var attempts = [];
    for (var i = 1; i < rows.length; i++) {
      if (String(rows[i][3]) !== String(cid)) continue;
      attempts.push(attemptRowToObj(rows[i]));
    }
    return { success: true, attempts: attempts };
  } catch(e) {
    logError('getAttemptsByCid', e);
    return { success: false, attempts: [] };
  }
}

function attemptRowToObj(r) {
  function safeParse(v) { try { return v ? JSON.parse(v) : null; } catch(e) { return null; } }
  return {
    attemptId:   String(r[0]),
    examId:      String(r[1]),
    examVersion: Number(r[2]),
    cid:         String(r[3]),
    firstName:   String(r[4] || ''),
    lastName:    String(r[5] || ''),
    startedAt:   String(r[6] || ''),
    submittedAt: String(r[7] || ''),
    score:       Number(r[8]),
    totalPoints: Number(r[9]),
    pct:         Number(r[10]),
    passed:      r[11] === 'Y',
    answers:     safeParse(r[12]),
    feedback:    safeParse(r[13]),
    isPreview:   r[14] === 'Y',
  };
}

// ─── ADMIN ACTION ROUTER (exam center) ───────────────────────────────────────
// These are called via handleAdminAction({ action: 'exam_XXX', ... })

function handleExamCenterAction(data) {
  var action = data.examAction;
  switch(action) {
    case 'setup':              return setupExamSheets();
    case 'createExam':         return createExam(data);
    case 'updateExam':         return updateExam(data);
    case 'getExams':           return getExams(data.filterStatus);
    case 'getExamById':        return getExamById(data.examId);
    case 'publishExam':        return publishExam(data.examId);
    case 'archiveExam':        return archiveExam(data.examId);
    case 'deleteExam':         return deleteExam(data.examId);
    case 'createBank':         return createBank(data);
    case 'getBanks':           return getBanks();
    case 'deleteBank':         return deleteBank(data.bankId);
    case 'createSection':      return createSection(data);
    case 'updateSection':      return updateSection(data);
    case 'getSectionsByExam':  return getSectionsByExam(data.examId);
    case 'deleteSection':      return deleteSection(data.sectionId);
    case 'createQuestion':     return createQuestion(data);
    case 'updateQuestion':     return updateQuestion(data);
    case 'deleteQuestion':     return deleteQuestion(data.questionId);
    case 'getQuestionsByExam': return getQuestionsByExam(data.examId);
    case 'getQuestionsByBank': return getQuestionsByBank(data.bankId);
    case 'buildExam':          return buildExamForDelivery(data.examId, data.isPreview);
    case 'verifyPassword':     return verifyExamPassword(data.examId, data.password);
    case 'checkEligibility':   return checkAttemptEligibility(data.examId, data.cid);
    case 'submitAttempt':      return submitExamAttempt(data);
    case 'getAttemptsByExam':  return getAttemptsByExam(data.examId);
    case 'getAttemptsByCid':   return getAttemptsByCid(data.cid);
    default: return { success: false, error: 'Unknown exam action: ' + action };
  }
}


// ─── EXAM PORTAL PUBLIC FUNCTIONS ────────────────────────────────────────────

// CID lookup for exam portal — checks VATUSA roster
function lookupVatsimMember(cid) {
  try {
    var url = 'https://api.vatusa.net/v2/user/' + cid + '?apikey=' + CONFIG.VATUSA_API_KEY;
    var resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    var json = JSON.parse(resp.getContentText());
    if (!json || !json.data) return { success: false };
    var d = json.data;
    var name = (d.fname || '') + ' ' + (d.lname || '');
    return { success: true, name: name.trim(), cid: String(cid), rating: d.rating_short || '' };
  } catch(e) {
    logError('lookupVatsimMember', e);
    return { success: false };
  }
}

// Returns all published exams (no passwords in payload)
function getPublishedExams() {
  try {
    var result = getExams('Published');
    var exams = (result.exams || []).map(function(ex) {
      return {
        examId:        ex.examId,
        title:         ex.title,
        rating:        ex.rating,
        description:   ex.description,
        passingPct:    ex.passingPct,
        totalPoints:   ex.totalPoints,
        timeLimitMins: ex.timeLimitMins,
        maxAttempts:   ex.maxAttempts,
        cooldownHours: ex.cooldownHours,
        version:       ex.version,
        hasPassword:   !!ex.password,
        password:      ex.password,  // needed for front-end verify step
      };
    });
    return { success: true, exams: exams };
  } catch(e) {
    logError('getPublishedExams', e);
    return { success: false, exams: [] };
  }
}

// Public wrapper — verify password
function verifyExamPasswordPublic(examId, password) {
  return verifyExamPassword(examId, password);
}

// Public wrapper — check attempt eligibility
function checkAttemptEligibilityPublic(examId, cid) {
  return checkAttemptEligibility(examId, cid);
}

// Public wrapper — build exam for delivery (no correct answers)
function buildExamPublic(examId) {
  return buildExamForDelivery(examId, false);
}

// Public wrapper — submit attempt
function submitExamAttemptPublic(data) {
  return submitExamAttempt(data);
}

// Public wrapper — get attempts by CID
function getAttemptsByCidPublic(cid) {
  return getAttemptsByCid(cid);
}

// Recent passes — last 20 passing attempts (names anonymized to first name + last initial)
function getRecentPasses() {
  try {
    var sheet = SpreadsheetApp.openById(CONFIG.SHEET_ID).getSheetByName(EXAM_TABS.ATTEMPTS);
    if (!sheet) return { success: true, passes: [] };
    var rows = sheet.getDataRange().getValues();
    var passes = [];
    for (var i = 1; i < rows.length; i++) {
      if (rows[i][11] !== 'Y') continue;   // not passed
      if (rows[i][14] === 'Y') continue;   // preview
      var examResult = getExamById(String(rows[i][1]));
      var examTitle = examResult.success ? examResult.exam.title : String(rows[i][1]);
      var firstName = String(rows[i][4] || '');
      var lastName  = String(rows[i][5] || '');
      var displayName = firstName + (lastName ? ' ' + lastName.charAt(0) + '.' : '');
      var submitted = rows[i][7];
      var date = submitted ? new Date(submitted).toLocaleDateString('en-US', { month: 'short', day: 'numeric' }) : '';
      passes.push({ name: displayName.trim() || 'Unknown', examTitle: examTitle, date: date, submittedAt: submitted });
    }
    // Sort by most recent, take last 20
    passes.sort(function(a,b){ return new Date(b.submittedAt) - new Date(a.submittedAt); });
    return { success: true, passes: passes.slice(0, 20) };
  } catch(e) {
    logError('getRecentPasses', e);
    return { success: true, passes: [] };
  }
}
