'use strict';

const PRES_CONFIG = Object.freeze({
  maxExecutionMs: 270000,
  maxBatchMs: 15000,
  maxInsertionsPerBatch: 12,
  maxRemovalsPerBatch: 20,
  folderPropertyKey: 'PRES_SURVEY_FOLDER_ID',
  folderNamePropertyKey: 'PRES_SURVEY_FOLDER_NAME',
  sheetPropertyKey: 'PRES_SURVEY_SHEET_ID',
  sheetNamePropertyKey: 'PRES_SURVEY_SHEET_NAME',
  runStatePropertyKey: 'PRES_RUN_STATE',
  userPickerKeyPropertyKey: 'PRES_PICKER_API_KEY',
  userPickerProjectPropertyKey: 'PRES_PICKER_PROJECT_NUMBER',
  projectInfoPropertyKey: 'PRES_PROJECT_INFO',
  markerPrefix: 'PRES_SURVEY',
  coverMarkerNumber: 0,
  surveyTabName: 'Survey',
  stylesTabName: 'Styles',
  materialsTabName: 'Materials',
  findingsTabName: 'Findings',
  jurisdictionsTabName: 'Jurisdictions',
  roofFormsTabName: 'RoofForms',
  windowTypesTabName: 'WindowTypes',
  doorTypesTabName: 'DoorTypes',
  porchTypesTabName: 'PorchTypes',
  foundationsTabName: 'Foundations',
  resourceTypesTabName: 'ResourceTypes',
  useTypesTabName: 'UseTypes',
  periodsTabName: 'Periods',
  unmatchedTabName: 'Unmatched',
  maxImageWidthPx: 520,
  sidebarTitle: 'Preservation Survey',
  diagnosticDocPrefix: 'Preservation Survey Diagnostic Log',
  pickerDialogWidth: 720,
  pickerDialogHeight: 560,
  defaultPickerDeveloperKey: '',
  defaultPickerCloudProjectNumber: '',
  maxConsoleLines: 24,
  maxStoredErrors: 24,
  maxStoredErrorDetails: 12,
});

const PRES_ACTIONS = Object.freeze({
  describe: Object.freeze({
    key: 'describe',
    label: 'Generate Descriptions',
    describe: true,
    insert: false,
    rebuild: false,
    forceRegenerate: true,
  }),
  insert: Object.freeze({
    key: 'insert',
    label: 'Insert Missing Photos + Captions',
    describe: true,
    insert: true,
    rebuild: false,
    forceRegenerate: false,
  }),
  rebuild: Object.freeze({
    key: 'rebuild',
    label: 'Rebuild All Photos + Captions',
    describe: true,
    insert: true,
    rebuild: true,
    forceRegenerate: true,
  }),
});

/** Opens the Preservation Survey menu in the active Google Doc. */
function onOpen() {
  const ui = DocumentApp.getUi();
  ui.createMenu('Preservation Survey')
      .addItem('Open Sidebar', 'showSidebar')
      .addSeparator()
      .addSubMenu(ui.createMenu('Setup')
          .addItem('Link Survey Sheet', 'showSheetPicker')
          .addItem('Link Photo Folder', 'showPicker')
          .addItem('Initialize Sheet Tabs & Dropdowns', 'initializeSurveySheet')
          .addItem('Set Project Info', 'showProjectInfoSidebar'))
      .addSeparator()
      .addSubMenu(ui.createMenu('Generate')
          .addItem('Generate Descriptions from Sheet', 'showSidebarForGenerateDescriptions')
          .addItem('Insert Missing Photos + Captions', 'showSidebarForInsertMissing')
          .addItem('Rebuild All Photos + Captions', 'showSidebarForRebuild'))
      .addSeparator()
      .addSubMenu(ui.createMenu('Export')
          .addItem('Export KML to Drive', 'exportKml')
          .addItem('Export Survey Summary CSV', 'exportCsv'))
      .addToUi();
}

/** Opens the main sidebar. */
function showSidebar() {
  showSidebar_('');
}

/** Opens the sidebar and starts the description action. */
function showSidebarForGenerateDescriptions() {
  showSidebar_('describe');
}

/** Opens the sidebar and starts the missing-photo insertion action. */
function showSidebarForInsertMissing() {
  showSidebar_('insert');
}

/** Opens the sidebar and starts the rebuild action. */
function showSidebarForRebuild() {
  showSidebar_('rebuild');
}

function showSidebar_(requestedAction) {
  const template = HtmlService.createTemplateFromFile('Sidebar');
  template.bootstrapJson = JSON.stringify(buildSidebarModel_(requestedAction || ''));
  const html = template.evaluate()
      .setTitle(PRES_CONFIG.sidebarTitle)
      .setSandboxMode(HtmlService.SandboxMode.IFRAME);
  DocumentApp.getUi().showSidebar(html);
}

/** Opens the Drive folder Picker dialog for the photo folder. */
function showPicker() {
  showPickerDialog_('folder');
}

/** Opens the Drive file Picker dialog for the survey spreadsheet. */
function showSheetPicker() {
  showPickerDialog_('sheet');
}

function showPickerDialog_(mode) {
  const template = HtmlService.createTemplateFromFile('Picker');
  template.bootstrapJson = JSON.stringify(buildPickerDialogModel_(mode));
  const title = mode === 'sheet' ? 'Select Survey Sheet' : 'Select Photo Folder';
  const html = template.evaluate()
      .setWidth(PRES_CONFIG.pickerDialogWidth)
      .setHeight(PRES_CONFIG.pickerDialogHeight)
      .setSandboxMode(HtmlService.SandboxMode.IFRAME);
  DocumentApp.getUi().showModalDialog(html, title);
}

/** Opens the folder Picker from the sidebar. */
function openPickerDialog() {
  showPicker();
  return true;
}

/** Opens the sheet Picker from the sidebar. */
function openSheetPickerDialog() {
  showSheetPicker();
  return true;
}

/** Returns an OAuth token for Google Picker. */
function getOAuthToken() {
  return ScriptApp.getOAuthToken();
}

/** Returns the current sidebar model. */
function getSidebarState() {
  return buildSidebarModel_('');
}

function buildSidebarModel_(requestedAction) {
  return {
    requestedAction: requestedAction,
    folder: getConfiguredFolderInfo_(),
    sheet: getConfiguredSheetInfo_(),
    picker: buildPickerClientState_(),
    projectInfo: getProjectInfo_(),
    run: buildClientRunState_(getStoredRunState_()),
  };
}

function buildPickerDialogModel_(mode) {
  return {
    mode: mode,
    folder: getConfiguredFolderInfo_(),
    sheet: getConfiguredSheetInfo_(),
    picker: buildPickerClientState_(),
    title: PRES_CONFIG.sidebarTitle,
  };
}

/** Saves a photo folder from a pasted URL or ID. */
function saveFolderInput(input) {
  const folderId = extractDriveId_(input);
  const folder = saveConfiguredFolderById_(folderId,
      'Could not save that Drive folder. Make sure the folder URL or ID is valid and accessible.');
  return buildSidebarModelAfterConfigSave_({folder: buildFolderInfo_(folder)});
}

/** Saves a survey spreadsheet from a pasted URL or ID. */
function saveSheetInput(input) {
  const sheetId = extractDriveId_(input);
  const spreadsheet = saveConfiguredSheetById_(sheetId,
      'Could not save that Google Sheet. Make sure the Sheet URL or ID is valid and accessible.');
  return buildSidebarModelAfterConfigSave_({sheet: buildSheetInfo_(spreadsheet)});
}

/** Creates a blank survey spreadsheet and links it to the active document. */
function createAndLinkSurveySheet() {
  const documentName = DocumentApp.getActiveDocument().getName();
  const spreadsheet = SpreadsheetApp.create(documentName + ' - Preservation Survey');
  saveConfiguredSheetById_(spreadsheet.getId(), 'The new survey Sheet could not be linked.');
  return buildSidebarModelAfterConfigSave_({sheet: buildSheetInfo_(spreadsheet)});
}

/** Handles a Google Picker folder or spreadsheet selection. */
function handlePickerResponse(id, itemName, mode) {
  if (!id) {
    throw new Error('No Drive item was selected.');
  }

  if (mode === 'sheet') {
    const spreadsheet = saveConfiguredSheetById_(String(id),
        'The selected survey Sheet could not be opened.', itemName);
    return buildSidebarModelAfterConfigSave_({sheet: buildSheetInfo_(spreadsheet)});
  }

  const folder = saveConfiguredFolderById_(String(id),
      'The selected photo folder could not be opened.', itemName);
  return buildSidebarModelAfterConfigSave_({folder: buildFolderInfo_(folder)});
}

/** Saves Google Picker credentials in UserProperties. */
function savePickerCredentials(apiKey, cloudProjectNumber) {
  const normalizedApiKey = String(apiKey || '').trim();
  const normalizedProjectNumber = String(cloudProjectNumber || '').trim();
  if (!normalizedApiKey) {
    throw new Error('Enter a Google Picker API key before saving.');
  }
  if (!normalizedProjectNumber || !/^\d+$/.test(normalizedProjectNumber)) {
    throw new Error('Enter the Google Cloud project number as digits only.');
  }
  const userProperties = PropertiesService.getUserProperties();
  userProperties.setProperty(PRES_CONFIG.userPickerKeyPropertyKey, normalizedApiKey);
  userProperties.setProperty(PRES_CONFIG.userPickerProjectPropertyKey, normalizedProjectNumber);
  return buildPickerClientState_();
}

/** Clears user-scoped Google Picker credentials. */
function clearPickerCredentials() {
  const userProperties = PropertiesService.getUserProperties();
  userProperties.deleteProperty(PRES_CONFIG.userPickerKeyPropertyKey);
  userProperties.deleteProperty(PRES_CONFIG.userPickerProjectPropertyKey);
  return buildPickerClientState_();
}

/** Starts a Preservation Survey batch. */
function startPhotoReportRun(actionKey) {
  return withPhotoReportLock_(function() {
    return executePhotoReportBatch_(actionKey, true);
  });
}

/** Resumes a paused Preservation Survey batch. */
function resumePhotoReportRun() {
  return withPhotoReportLock_(function() {
    const runState = getStoredRunState_();
    if (!runState || !runState.action) {
      return buildClientRunState_(runState);
    }
    return executePhotoReportBatch_(runState.action, false);
  });
}

/** Clears stored Preservation Survey batch state. */
function clearPhotoReportRun() {
  clearStoredRunState_();
  return buildClientRunState_(null);
}

function withPhotoReportLock_(callback) {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(1000)) {
    const runState = getStoredRunState_();
    if (runState) {
      runState.status = 'running';
      runState.message = 'Another Preservation Survey batch is already running. Waiting for it to finish.';
      runState.updatedAt = new Date().toISOString();
      return buildClientRunState_(runState);
    }
    throw new Error('Another Preservation Survey batch is already running. Wait for it to finish, then try again.');
  }

  try {
    return callback();
  } finally {
    lock.releaseLock();
  }
}

function executePhotoReportBatch_(actionKey, resetRunState) {
  const action = getPreservationAction_(actionKey);
  const batchLimits = buildBatchLimits_();
  let runState = resetRunState ? null : getStoredRunState_();

  try {
    const sheet = getConfiguredSurveySheet_();
    const surveyRead = readSurveyRows_(sheet);
    if (!surveyRead.rows.length) {
      throw new Error('No survey rows were found. Fill out at least one row on the Survey tab, then run the batch again.');
    }

    const folder = action.insert ? getConfiguredFolder_() : null;
    runState = initializeRunState_(runState, action, folder, surveyRead.rows.length, resetRunState);
    runState.status = 'running';
    runState.limitReached = false;
    runState.needsContinuation = false;
    runState.updatedAt = new Date().toISOString();
    if (resetRunState) {
      appendRunLog_(runState, action.label + ' started.');
    } else {
      appendRunLog_(runState, action.label + ' resumed.');
    }
    saveProgressSnapshot_(runState);

    if (action.describe && !runState.descriptionComplete) {
      runState.phase = 'describe';
      runState.message = buildRunPhaseMessage_(runState);
      saveProgressSnapshot_(runState);
      const describeSummary = describeSurveyRowsBatch_(sheet, surveyRead, action, batchLimits, runState);
      if (!describeSummary.complete) {
        return pauseRunState_(runState, 'describe', describeSummary.limitReached,
            describeSummary.limitReached ?
              'Execution time limit reached while generating descriptions.' :
              'Generating descriptions in the next batch.');
      }
      runState.descriptionComplete = true;
      appendRunLog_(runState, 'Description generation complete.');
      saveProgressSnapshot_(runState);
    }

    if (!action.insert) {
      return completeRunState_(runState);
    }

    const body = DocumentApp.getActiveDocument().getBody();
    if (action.rebuild && !runState.removalComplete) {
      runState.phase = 'remove';
      appendRunLog_(runState, 'Clearing previously managed survey blocks before rebuild.');
      saveProgressSnapshot_(runState);
      const removalSummary = removeManagedSurveyParagraphsBatch_(body, batchLimits, runState);
      if (!removalSummary.complete) {
        return pauseRunState_(runState, 'remove', removalSummary.limitReached,
            removalSummary.limitReached ?
              'Execution time limit reached while clearing managed survey blocks.' :
              'Removing previously managed survey blocks in the next batch.');
      }
      runState.removalComplete = true;
      appendRunLog_(runState, 'Managed survey block removal complete.');
      saveProgressSnapshot_(runState);
    }

    const freshSurveyRead = readSurveyRows_(sheet);
    const folderImages = collectFolderImages_(folder, runState);
    runState.duplicateFolderNumbers = folderImages.duplicateNumbers;
    ensureReportCoverBlock_(body, runState);

    let processedThisBatch = 0;
    runState.phase = 'insert';
    for (const row of freshSurveyRead.rows) {
      if (String(row.status || '').toLowerCase() === 'skip') {
        pushUniqueNumber_(runState.completedNumbers, row.photoNumber);
        continue;
      }
      if (runState.completedNumbers.indexOf(row.photoNumber) !== -1) {
        continue;
      }
      const pauseReason = getInsertPauseReason_(processedThisBatch, batchLimits);
      if (pauseReason) {
        return pauseRunState_(runState, 'insert', pauseReason === 'limit',
            pauseReason === 'limit' ?
              'Execution time limit reached. Resume to continue processing the remaining photos.' :
              'Continuing photo insertion in the next batch.');
      }
      processSurveyEntry_(body, row, folderImages, runState, action);
      processedThisBatch += 1;
    }

    return completeRunState_(runState);
  } catch (error) {
    return finalizeRunFailure_(action, runState, error);
  }
}

function describeSurveyRowsBatch_(spreadsheet, surveyRead, action, batchLimits, runState) {
  const surveySheet = surveyRead.sheet;
  const lookups = buildPreservationLookups_(spreadsheet);
  const outputs = [];
  let processedThisBatch = 0;

  for (const row of surveyRead.rows) {
    const alreadyDone = Boolean(String(row.generatedDescription || '').trim());
    const completed = runState.describedRowNumbers.indexOf(row.rowNumber) !== -1;
    if (completed || (alreadyDone && !action.forceRegenerate)) {
      pushUniqueNumber_(runState.describedRowNumbers, row.rowNumber);
      outputs.push(null);
      continue;
    }

    const pauseReason = getDescribePauseReason_(processedThisBatch, batchLimits);
    if (pauseReason) {
      writeDescriptionOutputs_(surveySheet, surveyRead.rows, outputs);
      return {complete: false, limitReached: pauseReason === 'limit'};
    }

    const description = assembleDescription_(row, lookups);
    row.generatedDescription = description;
    outputs.push([description]);
    pushUniqueNumber_(runState.describedRowNumbers, row.rowNumber);
    runState.describedCount += 1;
    runState.message = 'Generated description for survey row ' + row.rowNumber + '.';
    appendRunLog_(runState, runState.message);
    saveProgressSnapshot_(runState);
    processedThisBatch += 1;
  }

  writeDescriptionOutputs_(surveySheet, surveyRead.rows, outputs);
  return {complete: true, limitReached: false};
}

function writeDescriptionOutputs_(sheet, rows, outputs) {
  const pending = [];
  let startRow = null;
  for (let index = 0; index < rows.length; index += 1) {
    if (!outputs[index]) {
      if (pending.length) {
        sheet.getRange(startRow, PRES_SURVEY_COLUMN.generatedDescription, pending.length, 1).setValues(pending);
        pending.length = 0;
        startRow = null;
      }
      continue;
    }
    if (startRow === null) {
      startRow = rows[index].rowNumber;
    }
    pending.push(outputs[index]);
  }
  if (pending.length) {
    sheet.getRange(startRow, PRES_SURVEY_COLUMN.generatedDescription, pending.length, 1).setValues(pending);
  }
}

function getDescribePauseReason_(processedThisBatch, batchLimits) {
  if (Date.now() > batchLimits.hardDeadlineMs) {
    return 'limit';
  }
  if (processedThisBatch >= batchLimits.maxInsertionsPerBatch || Date.now() > batchLimits.softDeadlineMs) {
    return 'batch';
  }
  return '';
}

function removeManagedSurveyParagraphsBatch_(body, batchLimits, runState) {
  let removedCount = 0;
  for (let index = body.getNumChildren() - 1; index >= 0; index -= 1) {
    const pauseReason = getRemovalPauseReason_(removedCount, batchLimits);
    if (pauseReason) {
      return {removedCount: removedCount, complete: false, limitReached: pauseReason === 'limit'};
    }
    const child = body.getChild(index);
    if (!getManagedInfoFromBodyChild_(child)) {
      continue;
    }
    body.removeChild(child);
    removedCount += 1;
    runState.removedCount += 1;
    runState.message = 'Removing managed survey content. ' + removedCount + ' paragraph(s) removed this batch.';
    saveProgressSnapshot_(runState);
  }
  return {removedCount: removedCount, complete: true, limitReached: false};
}

function processSurveyEntry_(body, row, folderImages, runState, action) {
  const imageRecord = folderImages.byNumber[row.photoNumber];
  if (!imageRecord) {
    pushUniqueNumber_(runState.missingNumbers, row.photoNumber);
    pushUniqueNumber_(runState.completedNumbers, row.photoNumber);
    runState.message = 'Photo ' + row.photoNumber + ': no matching image found.';
    appendRunLog_(runState, runState.message);
    saveProgressSnapshot_(runState);
    return;
  }

  if (!action.rebuild && hasManagedEntry_(body, row.photoNumber)) {
    runState.skippedCount += 1;
    pushUniqueNumber_(runState.completedNumbers, row.photoNumber);
    runState.message = 'Photo ' + row.photoNumber + ': existing managed entry found, skipped.';
    appendRunLog_(runState, runState.message);
    saveProgressSnapshot_(runState);
    return;
  }

  try {
    insertSurveyEntryBlock_(body, row, imageRecord.file);
    runState.insertedCount += 1;
    runState.message = 'Photo ' + row.photoNumber + ': inserted ' + imageRecord.file.getName() + '.';
    appendRunLog_(runState, runState.message);
  } catch (error) {
    const errorDetail = {
      type: 'insert',
      number: row.photoNumber,
      fileId: imageRecord.file.getId(),
      fileName: imageRecord.file.getName(),
      message: error.message,
    };
    runState.errors.push(formatDiagnosticError_(errorDetail));
    runState.errorDetails.push(errorDetail);
    runState.message = 'Photo ' + row.photoNumber + ': failed to insert ' + imageRecord.file.getName() + '.';
    appendRunLog_(runState, runState.message + ' ' + formatDiagnosticError_(errorDetail));
  }

  pushUniqueNumber_(runState.completedNumbers, row.photoNumber);
  saveProgressSnapshot_(runState);
}

function insertSurveyEntryBlock_(body, row, file) {
  const marker = buildManagedMarker_(row.photoNumber, file.getId(), 'entry');
  const heading = body.appendParagraph('Photo ' + row.photoNumber + ' - ' + (row.streetAddress || 'Unaddressed resource'));
  heading.setHeading(DocumentApp.ParagraphHeading.HEADING3);
  heading.setSpacingBefore(12);
  heading.editAsText().setBold(true).setForegroundColor('#000000').setFontSize(12);
  appendTextMarker_(heading, marker);

  const imageParagraph = body.appendParagraph('');
  imageParagraph.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  imageParagraph.setSpacingAfter(6);
  const image = imageParagraph.insertInlineImage(0, file.getBlob());
  image.setAltTitle('Preservation Survey Photo ' + row.photoNumber);
  image.setAltDescription(marker);
  resizeInlineImage_(image, PRES_CONFIG.maxImageWidthPx);

  const description = String(row.generatedDescription || '').trim() || assembleDescription_(row, buildPreservationLookups_(getConfiguredSurveySheet_()));
  const descriptionParagraph = body.appendParagraph(description);
  descriptionParagraph.setSpacingAfter(4);
  appendTextMarker_(descriptionParagraph, marker);

  const metadataParagraph = body.appendParagraph(buildMetadataLine_(row));
  metadataParagraph.setSpacingAfter(12);
  metadataParagraph.editAsText().setFontSize(9).setForegroundColor('#666666');
  appendTextMarker_(metadataParagraph, marker);
}

function buildMetadataLine_(row) {
  const parts = [];
  if (row.gpsLat && row.gpsLon) {
    parts.push('GPS: ' + row.gpsLat + ', ' + row.gpsLon);
  }
  if (row.recorder) {
    parts.push('Recorded by: ' + row.recorder);
  }
  if (row.surveyDate) {
    parts.push('Survey Date: ' + formatValueForDisplay_(row.surveyDate));
  }
  return parts.join('  |  ');
}

function ensureReportCoverBlock_(body, runState) {
  if (hasManagedEntry_(body, PRES_CONFIG.coverMarkerNumber)) {
    return;
  }
  const info = getProjectInfo_();
  const marker = buildManagedMarker_(PRES_CONFIG.coverMarkerNumber, DocumentApp.getActiveDocument().getId(), 'cover');
  const title = body.insertParagraph(0, info.projectName || 'Historic Preservation Field Survey');
  title.setHeading(DocumentApp.ParagraphHeading.TITLE);
  appendTextMarker_(title, marker);
  const lines = [
    info.projectNumber ? 'Project Number: ' + info.projectNumber : '',
    info.county || info.state ? 'Location: ' + [info.county, info.state].filter(Boolean).join(', ') : '',
    info.surveyType ? 'Survey Type: ' + info.surveyType : '',
    info.preparedBy ? 'Prepared By: ' + info.preparedBy : '',
    info.preparedFor ? 'Prepared For: ' + info.preparedFor : '',
    info.surveyDate ? 'Survey Date: ' + info.surveyDate : '',
    info.shpoFormNumber ? 'SHPO Form: ' + info.shpoFormNumber : '',
  ].filter(Boolean);
  const cover = body.insertParagraph(1, lines.join('\n'));
  cover.setSpacingAfter(18);
  appendTextMarker_(cover, marker);
  appendRunLog_(runState, 'Inserted preservation survey cover block.');
}

function hasManagedEntry_(body, number) {
  for (let index = 0; index < body.getNumChildren(); index += 1) {
    const marker = getManagedInfoFromBodyChild_(body.getChild(index));
    if (marker && marker.number === number) {
      return true;
    }
  }
  return false;
}

function getManagedInfoFromBodyChild_(child) {
  if (!child || child.getType() !== DocumentApp.ElementType.PARAGRAPH) {
    return null;
  }
  const paragraph = child.asParagraph();
  const textMarker = parseManagedMarkerFromText_(paragraph.getText());
  if (textMarker) {
    return textMarker;
  }
  let marker = null;
  for (let index = 0; index < paragraph.getNumChildren(); index += 1) {
    const paragraphChild = paragraph.getChild(index);
    if (paragraphChild.getType() === DocumentApp.ElementType.INLINE_IMAGE) {
      marker = parseManagedMarker_(paragraphChild.asInlineImage().getAltDescription()) || marker;
    }
  }
  return marker;
}

function buildManagedMarker_(number, fileId, kind) {
  return [PRES_CONFIG.markerPrefix, String(number), fileId, kind || 'entry'].join('|');
}

function parseManagedMarker_(text) {
  if (!text) {
    return null;
  }
  const parts = String(text).split('|');
  if (parts.length < 3 || parts[0] !== PRES_CONFIG.markerPrefix) {
    return null;
  }
  return {number: Number(parts[1]), fileId: parts[2], kind: parts[3] || 'entry'};
}

function parseManagedMarkerFromText_(text) {
  const match = String(text || '').match(/\[\[PRES_SURVEY\|[^\]]+\]\]/);
  if (!match) {
    return null;
  }
  return parseManagedMarker_(match[0].slice(2, -2));
}

function appendTextMarker_(paragraph, marker) {
  const text = paragraph.appendText('\n[[' + marker + ']]');
  text.setFontSize(1).setForegroundColor('#ffffff');
}

function readSurveyRows_(spreadsheet) {
  const sheet = spreadsheet.getSheetByName(PRES_CONFIG.surveyTabName);
  if (!sheet) {
    throw new Error('The linked survey Sheet does not have a Survey tab. Run Setup > Initialize Sheet Tabs & Dropdowns, then try again.');
  }
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return {sheet: sheet, rows: []};
  }
  const values = sheet.getRange(2, 1, lastRow - 1, PRES_SURVEY_HEADERS.length).getValues();
  const rows = [];
  for (let index = 0; index < values.length; index += 1) {
    const row = parseSurveyRow_(values[index], index + 2);
    if (!row.photoNumber && !row.streetAddress && !row.featureId) {
      continue;
    }
    rows.push(row);
  }
  return {sheet: sheet, rows: rows};
}

function parseSurveyRow_(values, rowNumber) {
  const get = function(column) {
    return values[column - 1];
  };
  return {
    rowNumber: rowNumber,
    photoNumber: Number(get(PRES_SURVEY_COLUMN.photoNumber)) || rowNumber - 1,
    streetAddress: cleanCell_(get(PRES_SURVEY_COLUMN.streetAddress)),
    city: cleanCell_(get(PRES_SURVEY_COLUMN.city)),
    state: cleanCell_(get(PRES_SURVEY_COLUMN.state)),
    featureId: cleanCell_(get(PRES_SURVEY_COLUMN.featureId)),
    resourceType: cleanCell_(get(PRES_SURVEY_COLUMN.resourceType)),
    useHistoric: cleanCell_(get(PRES_SURVEY_COLUMN.useHistoric)),
    useCurrent: cleanCell_(get(PRES_SURVEY_COLUMN.useCurrent)),
    stylePrimary: cleanCell_(get(PRES_SURVEY_COLUMN.stylePrimary)),
    styleSecondary: cleanCell_(get(PRES_SURVEY_COLUMN.styleSecondary)),
    stories: cleanCell_(get(PRES_SURVEY_COLUMN.stories)),
    roofForm: cleanCell_(get(PRES_SURVEY_COLUMN.roofForm)),
    roofMaterial: cleanCell_(get(PRES_SURVEY_COLUMN.roofMaterial)),
    primaryCladding: cleanCell_(get(PRES_SURVEY_COLUMN.primaryCladding)),
    secondaryCladding: cleanCell_(get(PRES_SURVEY_COLUMN.secondaryCladding)),
    foundationType: cleanCell_(get(PRES_SURVEY_COLUMN.foundationType)),
    windowType: cleanCell_(get(PRES_SURVEY_COLUMN.windowType)),
    windowMaterial: cleanCell_(get(PRES_SURVEY_COLUMN.windowMaterial)),
    doorType: cleanCell_(get(PRES_SURVEY_COLUMN.doorType)),
    porchType: cleanCell_(get(PRES_SURVEY_COLUMN.porchType)),
    findings: [
      cleanCell_(get(PRES_SURVEY_COLUMN.finding1)),
      cleanCell_(get(PRES_SURVEY_COLUMN.finding2)),
      cleanCell_(get(PRES_SURVEY_COLUMN.finding3)),
      cleanCell_(get(PRES_SURVEY_COLUMN.finding4)),
    ].filter(isRealSelection_),
    integrityOverall: cleanCell_(get(PRES_SURVEY_COLUMN.integrityOverall)),
    integrityLocation: cleanCell_(get(PRES_SURVEY_COLUMN.integrityLocation)),
    integrityDesign: cleanCell_(get(PRES_SURVEY_COLUMN.integrityDesign)),
    integritySetting: cleanCell_(get(PRES_SURVEY_COLUMN.integritySetting)),
    integrityMaterials: cleanCell_(get(PRES_SURVEY_COLUMN.integrityMaterials)),
    integrityWorkmanship: cleanCell_(get(PRES_SURVEY_COLUMN.integrityWorkmanship)),
    integrityFeeling: cleanCell_(get(PRES_SURVEY_COLUMN.integrityFeeling)),
    integrityAssociation: cleanCell_(get(PRES_SURVEY_COLUMN.integrityAssociation)),
    nrCriteria: cleanCell_(get(PRES_SURVEY_COLUMN.nrCriteria)),
    periodOfSignificance: cleanCell_(get(PRES_SURVEY_COLUMN.periodOfSignificance)),
    dateOfConstruction: cleanCell_(get(PRES_SURVEY_COLUMN.dateOfConstruction)),
    architectBuilder: cleanCell_(get(PRES_SURVEY_COLUMN.architectBuilder)),
    gpsLat: cleanCell_(get(PRES_SURVEY_COLUMN.gpsLat)),
    gpsLon: cleanCell_(get(PRES_SURVEY_COLUMN.gpsLon)),
    recorder: cleanCell_(get(PRES_SURVEY_COLUMN.recorder)),
    surveyDate: get(PRES_SURVEY_COLUMN.surveyDate),
    surveyorNotes: cleanCell_(get(PRES_SURVEY_COLUMN.surveyorNotes)),
    status: cleanCell_(get(PRES_SURVEY_COLUMN.status)),
    generatedDescription: cleanCell_(get(PRES_SURVEY_COLUMN.generatedDescription)),
  };
}

function cleanCell_(value) {
  if (value instanceof Date) {
    return value;
  }
  return String(value || '').trim();
}

function isRealSelection_(value) {
  const text = String(value || '').trim();
  return Boolean(text && text !== '(none)');
}

function collectFolderImages_(folder, runState) {
  const imageRecords = [];
  const duplicateNumbers = [];
  const iterator = folder.getFiles();
  let totalFiles = 0;
  let imageCandidates = 0;
  while (iterator.hasNext()) {
    const file = iterator.next();
    totalFiles += 1;
    const fileSupport = getSupportedImageInfo_(file);
    if (!fileSupport.supported) {
      continue;
    }
    imageCandidates += 1;
    const number = extractTrailingNumber_(file.getName());
    if (number === null) {
      continue;
    }
    imageRecords.push({
      number: number,
      file: file,
      preferredName: isPreferredImageName_(file.getName(), number),
      updatedAtMs: file.getLastUpdated().getTime(),
    });
  }
  imageRecords.sort(function(left, right) {
    return left.number - right.number ||
      Number(right.preferredName) - Number(left.preferredName) ||
      right.updatedAtMs - left.updatedAtMs ||
      left.file.getName().localeCompare(right.file.getName()) ||
      left.file.getId().localeCompare(right.file.getId());
  });
  const byNumber = {};
  for (const record of imageRecords) {
    if (byNumber[record.number]) {
      duplicateNumbers.push(record.number);
      continue;
    }
    byNumber[record.number] = record;
  }
  if (runState) {
    appendRunLog_(runState, 'Folder scan: ' + totalFiles + ' file(s), ' + imageCandidates + ' image candidate(s), ' + imageRecords.length + ' numbered image(s).');
  }
  return {byNumber: byNumber, duplicateNumbers: uniqueNumbers_(duplicateNumbers)};
}

function getSupportedImageInfo_(file) {
  if (/^image\//i.test(file.getMimeType())) {
    return {supported: true, viaFileName: false};
  }
  return {supported: isLikelyImageByFileName_(file.getName()), viaFileName: true};
}

function isLikelyImageByFileName_(fileName) {
  return /\.(avif|bmp|gif|heic|heif|jpe?g|png|tiff?|webp)$/i.test(String(fileName || ''));
}

function isPreferredImageName_(fileName, number) {
  const escapedNumber = String(number).replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  const exactNamePattern = new RegExp('^(Image|Photo)\\s*' + escapedNumber + '(\\.[^.]+)?$', 'i');
  return exactNamePattern.test(fileName);
}

function extractTrailingNumber_(fileName) {
  const baseName = String(fileName || '').replace(/\.[^.]+$/, '');
  const match = baseName.match(/(\d+)(?!.*\d)/);
  return match ? Number(match[1]) : null;
}

function resizeInlineImage_(image, maxWidthPx) {
  const currentWidth = image.getWidth();
  const currentHeight = image.getHeight();
  if (!currentWidth || !currentHeight || currentWidth <= maxWidthPx) {
    return;
  }
  const scale = maxWidthPx / currentWidth;
  image.setWidth(Math.round(currentWidth * scale));
  image.setHeight(Math.round(currentHeight * scale));
}

function buildBatchLimits_() {
  const startedAtMs = Date.now();
  return {
    startedAtMs: startedAtMs,
    hardDeadlineMs: startedAtMs + PRES_CONFIG.maxExecutionMs,
    softDeadlineMs: startedAtMs + PRES_CONFIG.maxBatchMs,
    maxInsertionsPerBatch: PRES_CONFIG.maxInsertionsPerBatch,
    maxRemovalsPerBatch: PRES_CONFIG.maxRemovalsPerBatch,
  };
}

function initializeRunState_(runState, action, folder, totalCount, resetRunState) {
  if (resetRunState || !runState || runState.action !== action.key) {
    return createRunState_(action, folder, totalCount);
  }
  const nextState = runState;
  nextState.actionLabel = action.label;
  nextState.folderId = folder ? folder.getId() : '';
  nextState.folderName = folder ? folder.getName() : '';
  nextState.totalCount = totalCount;
  nextState.completedNumbers = ensureNumberArray_(nextState.completedNumbers);
  nextState.describedRowNumbers = ensureNumberArray_(nextState.describedRowNumbers);
  nextState.missingNumbers = ensureNumberArray_(nextState.missingNumbers);
  nextState.insertedCount = Number(nextState.insertedCount || 0);
  nextState.describedCount = Number(nextState.describedCount || 0);
  nextState.removedCount = Number(nextState.removedCount || 0);
  nextState.skippedCount = Number(nextState.skippedCount || 0);
  nextState.errors = ensureStringArray_(nextState.errors);
  nextState.errorDetails = ensureObjectArray_(nextState.errorDetails);
  nextState.consoleLines = ensureRecentStringArray_(nextState.consoleLines);
  nextState.duplicateFolderNumbers = ensureNumberArray_(nextState.duplicateFolderNumbers);
  nextState.descriptionComplete = Boolean(nextState.descriptionComplete || !action.describe);
  nextState.removalComplete = Boolean(nextState.removalComplete || action.rebuild === false);
  nextState.startedAt = nextState.startedAt || new Date().toISOString();
  nextState.updatedAt = new Date().toISOString();
  nextState.diagnosticLogUrl = String(nextState.diagnosticLogUrl || '');
  return nextState;
}

function createRunState_(action, folder, totalCount) {
  return {
    action: action.key,
    actionLabel: action.label,
    status: 'running',
    phase: action.describe ? 'describe' : 'insert',
    folderId: folder ? folder.getId() : '',
    folderName: folder ? folder.getName() : '',
    totalCount: totalCount,
    completedNumbers: [],
    describedRowNumbers: [],
    missingNumbers: [],
    duplicateFolderNumbers: [],
    insertedCount: 0,
    describedCount: 0,
    removedCount: 0,
    skippedCount: 0,
    errors: [],
    errorDetails: [],
    consoleLines: [],
    limitReached: false,
    needsContinuation: false,
    descriptionComplete: !action.describe,
    removalComplete: action.rebuild === false,
    startedAt: new Date().toISOString(),
    updatedAt: new Date().toISOString(),
    message: 'Preparing preservation survey batch.',
    diagnosticLogUrl: '',
  };
}

function pauseRunState_(runState, phase, limitReached, message) {
  runState.phase = phase;
  runState.limitReached = limitReached;
  runState.needsContinuation = needsMoreWork_(runState);
  runState.status = runState.needsContinuation ? (limitReached ? 'paused' : 'running') : 'complete';
  runState.updatedAt = new Date().toISOString();
  runState.message = runState.needsContinuation ? buildPauseMessage_(runState, message) : buildCompletionMessage_(runState);
  appendRunLog_(runState, runState.message);
  if (!runState.needsContinuation) {
    runState.phase = 'complete';
    runState.limitReached = false;
  }
  saveRunState_(runState);
  return buildClientRunState_(runState);
}

function completeRunState_(runState) {
  runState.phase = 'complete';
  runState.status = 'complete';
  runState.needsContinuation = false;
  runState.limitReached = false;
  runState.updatedAt = new Date().toISOString();
  runState.message = buildCompletionMessage_(runState);
  appendRunLog_(runState, 'Batch complete. ' + runState.message);
  saveRunState_(runState);
  return buildClientRunState_(runState);
}

function needsMoreWork_(runState) {
  if (!runState.descriptionComplete) {
    return true;
  }
  if (!runState.removalComplete) {
    return true;
  }
  return runState.completedNumbers.length < runState.totalCount;
}

function getInsertPauseReason_(processedThisBatch, batchLimits) {
  if (Date.now() > batchLimits.hardDeadlineMs) {
    return 'limit';
  }
  if (processedThisBatch >= batchLimits.maxInsertionsPerBatch || Date.now() > batchLimits.softDeadlineMs) {
    return 'batch';
  }
  return '';
}

function getRemovalPauseReason_(removedCount, batchLimits) {
  if (Date.now() > batchLimits.hardDeadlineMs) {
    return 'limit';
  }
  if (removedCount >= batchLimits.maxRemovalsPerBatch || Date.now() > batchLimits.softDeadlineMs) {
    return 'batch';
  }
  return '';
}

function finalizeRunFailure_(action, runState, error) {
  const nextState = runState || createFallbackRunState_(action);
  const errorMessage = error && error.message ? error.message : String(error);
  nextState.status = 'error';
  nextState.phase = 'error';
  nextState.needsContinuation = false;
  nextState.limitReached = false;
  nextState.updatedAt = new Date().toISOString();
  nextState.errors = ensureStringArray_(nextState.errors);
  nextState.errorDetails = ensureObjectArray_(nextState.errorDetails);
  nextState.consoleLines = ensureRecentStringArray_(nextState.consoleLines);
  nextState.errors.push(errorMessage);
  nextState.errorDetails.push({type: 'fatal', number: '', fileId: '', fileName: '', message: errorMessage});
  nextState.message = 'Run failed. Review the diagnostic log for details.';
  appendRunLog_(nextState, errorMessage);
  nextState.diagnosticLogUrl = createDiagnosticLogSafely_(nextState);
  saveRunState_(nextState);
  return buildClientRunState_(nextState);
}

function createFallbackRunState_(action) {
  const activeDocument = DocumentApp.getActiveDocument();
  const folderInfo = getConfiguredFolderInfo_();
  return {
    action: action ? action.key : '',
    actionLabel: action ? action.label : 'Preservation Survey',
    status: 'error',
    phase: 'error',
    folderId: folderInfo ? folderInfo.id : '',
    folderName: folderInfo ? folderInfo.name : '',
    totalCount: 0,
    completedNumbers: [],
    describedRowNumbers: [],
    missingNumbers: [],
    duplicateFolderNumbers: [],
    insertedCount: 0,
    describedCount: 0,
    removedCount: 0,
    skippedCount: 0,
    errors: [],
    errorDetails: [],
    consoleLines: [],
    limitReached: false,
    needsContinuation: false,
    descriptionComplete: false,
    removalComplete: action ? action.rebuild === false : true,
    startedAt: new Date().toISOString(),
    updatedAt: new Date().toISOString(),
    message: '',
    diagnosticLogUrl: '',
    documentId: activeDocument.getId(),
    documentName: activeDocument.getName(),
  };
}

function createDiagnosticLogSafely_(runState) {
  try {
    return createDiagnosticLog_(runState);
  } catch (diagnosticError) {
    runState.errors.push('Could not create diagnostic log: ' + (diagnosticError.message || String(diagnosticError)));
    return '';
  }
}

function createDiagnosticLog_(runState) {
  const activeDocument = DocumentApp.getActiveDocument();
  const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  const diagnosticDocument = DocumentApp.create(PRES_CONFIG.diagnosticDocPrefix + ' - ' + timestamp);
  const diagnosticBody = diagnosticDocument.getBody();
  diagnosticBody.appendParagraph('Preservation Survey Diagnostic Log').setHeading(DocumentApp.ParagraphHeading.TITLE);
  diagnosticBody.appendParagraph('Generated: ' + timestamp);
  diagnosticBody.appendParagraph('Action: ' + (runState.actionLabel || runState.action || 'Unknown'));
  diagnosticBody.appendParagraph('Source document: ' + activeDocument.getName() + ' (' + activeDocument.getId() + ')');
  if (runState.folderName || runState.folderId) {
    diagnosticBody.appendParagraph('Photo folder: ' + (runState.folderName || 'Unknown') + ' (' + (runState.folderId || 'n/a') + ')');
  }
  diagnosticBody.appendParagraph('Progress: ' + String(runState.completedNumbers ? runState.completedNumbers.length : 0) + ' of ' + String(runState.totalCount || 0));
  if (runState.errors && runState.errors.length) {
    diagnosticBody.appendParagraph('Errors').setHeading(DocumentApp.ParagraphHeading.HEADING1);
    for (const errorMessage of runState.errors) {
      diagnosticBody.appendListItem(errorMessage);
    }
  }
  if (runState.errorDetails && runState.errorDetails.length) {
    diagnosticBody.appendParagraph('Detailed diagnostics').setHeading(DocumentApp.ParagraphHeading.HEADING1);
    for (const detail of runState.errorDetails) {
      diagnosticBody.appendListItem([
        'Type: ' + String(detail.type || 'unknown'),
        'Photo: ' + String(detail.number || 'n/a'),
        'File name: ' + String(detail.fileName || 'n/a'),
        'File ID: ' + String(detail.fileId || 'n/a'),
        'Message: ' + String(detail.message || 'n/a'),
      ].join(' | '));
    }
  }
  if (runState.duplicateFolderNumbers && runState.duplicateFolderNumbers.length) {
    diagnosticBody.appendParagraph('Duplicate image numbers in folder').setHeading(DocumentApp.ParagraphHeading.HEADING1);
    diagnosticBody.appendParagraph(runState.duplicateFolderNumbers.join(', '));
  }
  if (runState.missingNumbers && runState.missingNumbers.length) {
    diagnosticBody.appendParagraph('Missing image numbers').setHeading(DocumentApp.ParagraphHeading.HEADING1);
    diagnosticBody.appendParagraph(runState.missingNumbers.join(', '));
  }
  diagnosticDocument.saveAndClose();
  return diagnosticDocument.getUrl();
}

/** Prompts for project metadata and stores it in DocumentProperties. */
function showProjectInfoSidebar() {
  const ui = DocumentApp.getUi();
  const current = getProjectInfo_();
  const fields = [
    ['projectName', 'Project name'],
    ['projectNumber', 'Project number'],
    ['county', 'County'],
    ['state', 'State'],
    ['surveyType', 'Survey type'],
    ['preparedBy', 'Prepared by'],
    ['preparedFor', 'Prepared for'],
    ['surveyDate', 'Survey date'],
    ['shpoFormNumber', 'SHPO form number'],
  ];
  const next = {};
  for (const field of fields) {
    const response = ui.prompt('Set Project Info', field[1] + ':', ui.ButtonSet.OK_CANCEL);
    if (response.getSelectedButton() !== ui.Button.OK) {
      return;
    }
    next[field[0]] = response.getResponseText() || current[field[0]] || '';
  }
  saveProjectInfo_(next);
  ui.alert('Project information saved.');
}

function getProjectInfo_() {
  const raw = PropertiesService.getDocumentProperties().getProperty(PRES_CONFIG.projectInfoPropertyKey);
  const defaults = {
    projectName: '', projectNumber: '', county: '', state: '', surveyType: '', preparedBy: '',
    preparedFor: '', surveyDate: '', shpoFormNumber: '',
  };
  if (!raw) {
    return defaults;
  }
  try {
    return Object.assign(defaults, JSON.parse(raw));
  } catch (error) {
    return defaults;
  }
}

function saveProjectInfo_(projectInfo) {
  PropertiesService.getDocumentProperties().setProperty(PRES_CONFIG.projectInfoPropertyKey, JSON.stringify(projectInfo));
}

/** Exports survey rows as a CSV file in the photo folder. */
function exportCsv() {
  const spreadsheet = getConfiguredSurveySheet_();
  const sheet = spreadsheet.getSheetByName(PRES_CONFIG.surveyTabName);
  if (!sheet) {
    throw new Error('The linked survey Sheet does not have a Survey tab. Initialize the Sheet before exporting CSV.');
  }
  const values = sheet.getRange(1, 1, sheet.getLastRow(), PRES_SURVEY_HEADERS.length).getDisplayValues();
  const csv = values.map(function(row) {
    return row.map(escapeCsvCell_).join(',');
  }).join('\n');
  const info = getProjectInfo_();
  const folder = getConfiguredFolder_();
  const fileName = sanitizeFileName_((info.projectName || 'Preservation_Survey') + '_Survey_' + formatDateForFile_(new Date()) + '.csv');
  folder.createFile(fileName, csv, MimeType.CSV);
  DocumentApp.getActiveDocument().toast('CSV exported: ' + fileName);
}

function escapeCsvCell_(value) {
  const text = String(value || '');
  if (/[",\n\r]/.test(text)) {
    return '"' + text.replace(/"/g, '""') + '"';
  }
  return text;
}

function getConfiguredFolder_() {
  const folderId = PropertiesService.getDocumentProperties().getProperty(PRES_CONFIG.folderPropertyKey);
  if (!folderId) {
    throw new Error('No photo folder is configured. Use Preservation Survey > Setup > Link Photo Folder first.');
  }
  try {
    return DriveApp.getFolderById(folderId);
  } catch (error) {
    throw new Error('The saved photo folder could not be opened. Use Preservation Survey > Setup > Link Photo Folder again.');
  }
}

function getConfiguredSurveySheet_() {
  const sheetId = PropertiesService.getDocumentProperties().getProperty(PRES_CONFIG.sheetPropertyKey);
  if (!sheetId) {
    throw new Error('No survey Sheet is configured. Use Preservation Survey > Setup > Link Survey Sheet first.');
  }
  try {
    return SpreadsheetApp.openById(sheetId);
  } catch (error) {
    throw new Error('The saved survey Sheet could not be opened. Use Preservation Survey > Setup > Link Survey Sheet again.');
  }
}

function getConfiguredFolderInfo_() {
  const documentProperties = PropertiesService.getDocumentProperties();
  const folderId = documentProperties.getProperty(PRES_CONFIG.folderPropertyKey);
  const savedFolderName = String(documentProperties.getProperty(PRES_CONFIG.folderNamePropertyKey) || '');
  if (!folderId) {
    return null;
  }
  try {
    return buildFolderInfo_(DriveApp.getFolderById(folderId));
  } catch (error) {
    return {id: folderId, name: savedFolderName || 'Configured folder unavailable', invalid: true};
  }
}

function getConfiguredSheetInfo_() {
  const documentProperties = PropertiesService.getDocumentProperties();
  const sheetId = documentProperties.getProperty(PRES_CONFIG.sheetPropertyKey);
  const savedSheetName = String(documentProperties.getProperty(PRES_CONFIG.sheetNamePropertyKey) || '');
  if (!sheetId) {
    return null;
  }
  try {
    return buildSheetInfo_(SpreadsheetApp.openById(sheetId));
  } catch (error) {
    return {id: sheetId, name: savedSheetName || 'Configured Sheet unavailable', invalid: true};
  }
}

function buildFolderInfo_(folder) {
  return {id: folder.getId(), name: folder.getName()};
}

function buildSheetInfo_(spreadsheet) {
  return {id: spreadsheet.getId(), name: spreadsheet.getName(), url: spreadsheet.getUrl()};
}

function saveConfiguredFolderById_(folderId, errorMessage, preferredFolderName) {
  let folder;
  try {
    folder = DriveApp.getFolderById(folderId);
  } catch (error) {
    throw new Error(errorMessage);
  }
  const documentProperties = PropertiesService.getDocumentProperties();
  documentProperties.setProperty(PRES_CONFIG.folderPropertyKey, folder.getId());
  documentProperties.setProperty(PRES_CONFIG.folderNamePropertyKey, String(preferredFolderName || folder.getName()));
  clearStoredRunState_();
  return folder;
}

function saveConfiguredSheetById_(sheetId, errorMessage, preferredSheetName) {
  let spreadsheet;
  try {
    spreadsheet = SpreadsheetApp.openById(sheetId);
  } catch (error) {
    throw new Error(errorMessage);
  }
  const documentProperties = PropertiesService.getDocumentProperties();
  documentProperties.setProperty(PRES_CONFIG.sheetPropertyKey, spreadsheet.getId());
  documentProperties.setProperty(PRES_CONFIG.sheetNamePropertyKey, String(preferredSheetName || spreadsheet.getName()));
  clearStoredRunState_();
  return spreadsheet;
}

function buildSidebarModelAfterConfigSave_(overrides) {
  return Object.assign(buildSidebarModel_(''), overrides || {});
}

function extractDriveId_(value) {
  const input = String(value || '').trim();
  if (!input) {
    throw new Error('Paste a Google Drive URL or raw file ID.');
  }
  const pathMatch = input.match(/\/(?:folders|d)\/([a-zA-Z0-9_-]+)/);
  if (pathMatch) {
    return pathMatch[1];
  }
  const idParamMatch = input.match(/[?&]id=([a-zA-Z0-9_-]+)/);
  if (idParamMatch) {
    return idParamMatch[1];
  }
  if (/^[a-zA-Z0-9_-]{10,}$/.test(input)) {
    return input;
  }
  throw new Error('Could not find a Google Drive ID in that value. Paste the full URL or the raw ID.');
}

function getPickerConfig_() {
  const userProperties = PropertiesService.getUserProperties();
  const userDeveloperKey = String(userProperties.getProperty(PRES_CONFIG.userPickerKeyPropertyKey) || '').trim();
  const userCloudProjectNumber = String(userProperties.getProperty(PRES_CONFIG.userPickerProjectPropertyKey) || '').trim();
  if (userDeveloperKey && userCloudProjectNumber) {
    return {developerKey: userDeveloperKey, cloudProjectNumber: userCloudProjectNumber, source: 'user'};
  }
  const defaultDeveloperKey = String(PRES_CONFIG.defaultPickerDeveloperKey || '').trim();
  const defaultCloudProjectNumber = String(PRES_CONFIG.defaultPickerCloudProjectNumber || '').trim();
  if (defaultDeveloperKey && defaultCloudProjectNumber) {
    return {developerKey: defaultDeveloperKey, cloudProjectNumber: defaultCloudProjectNumber, source: 'repo'};
  }
  return null;
}

function buildPickerClientState_() {
  const pickerConfig = getPickerConfig_();
  return {
    configured: Boolean(pickerConfig),
    source: pickerConfig ? pickerConfig.source : '',
    developerKey: pickerConfig ? pickerConfig.developerKey : '',
    cloudProjectNumber: pickerConfig ? pickerConfig.cloudProjectNumber : '',
    hasUserCredentials: hasUserPickerCredentials_(),
  };
}

function hasUserPickerCredentials_() {
  const userProperties = PropertiesService.getUserProperties();
  return Boolean(userProperties.getProperty(PRES_CONFIG.userPickerKeyPropertyKey) && userProperties.getProperty(PRES_CONFIG.userPickerProjectPropertyKey));
}

function getPreservationAction_(actionKey) {
  if (!PRES_ACTIONS[actionKey]) {
    throw new Error('Unknown Preservation Survey action: ' + actionKey + '. Open the sidebar and choose a valid action.');
  }
  return PRES_ACTIONS[actionKey];
}

function getStoredRunState_() {
  const rawState = PropertiesService.getDocumentProperties().getProperty(PRES_CONFIG.runStatePropertyKey);
  if (!rawState) {
    return null;
  }
  try {
    return JSON.parse(rawState);
  } catch (error) {
    clearStoredRunState_();
    return null;
  }
}

function saveRunState_(runState) {
  normalizeRunStateForStorage_(runState);
  PropertiesService.getDocumentProperties().setProperty(PRES_CONFIG.runStatePropertyKey, JSON.stringify(runState));
}

function clearStoredRunState_() {
  PropertiesService.getDocumentProperties().deleteProperty(PRES_CONFIG.runStatePropertyKey);
}

function normalizeRunStateForStorage_(runState) {
  runState.consoleLines = ensureRecentStringArray_(runState.consoleLines);
  runState.errors = ensureRecentStringArray_(runState.errors, PRES_CONFIG.maxStoredErrors);
  runState.errorDetails = ensureRecentObjectArray_(runState.errorDetails, PRES_CONFIG.maxStoredErrorDetails);
}

function saveProgressSnapshot_(runState) {
  runState.updatedAt = new Date().toISOString();
  saveRunState_(runState);
}

function buildClientRunState_(runState) {
  if (!runState) {
    return {
      action: '', actionLabel: '', status: 'idle', phase: 'idle', folderName: '', totalCount: 0,
      processedCount: 0, remainingCount: 0, progressPercent: 0, insertedCount: 0, describedCount: 0,
      removedCount: 0, skippedCount: 0, missingNumbers: [], duplicateFolderNumbers: [], errors: [],
      consoleLines: [], limitReached: false, needsContinuation: false, canResume: false,
      message: 'Ready to link a survey Sheet and photo folder.', startedAt: '', updatedAt: '', diagnosticLogUrl: '',
    };
  }
  const processedCount = runState.phase === 'describe' ?
    ensureNumberArray_(runState.describedRowNumbers).length :
    ensureNumberArray_(runState.completedNumbers).length;
  const totalCount = Number(runState.totalCount || 0);
  const remainingCount = Math.max(totalCount - processedCount, 0);
  const progressPercent = totalCount ? Math.max(0, Math.min(100, Math.round((processedCount / totalCount) * 100))) : 0;
  return {
    action: String(runState.action || ''),
    actionLabel: String(runState.actionLabel || ''),
    status: String(runState.status || 'idle'),
    phase: String(runState.phase || 'idle'),
    folderName: String(runState.folderName || ''),
    totalCount: totalCount,
    processedCount: processedCount,
    remainingCount: remainingCount,
    progressPercent: progressPercent,
    insertedCount: Number(runState.insertedCount || 0),
    describedCount: Number(runState.describedCount || 0),
    removedCount: Number(runState.removedCount || 0),
    skippedCount: Number(runState.skippedCount || 0),
    missingNumbers: sortedNumberCopy_(runState.missingNumbers),
    duplicateFolderNumbers: sortedNumberCopy_(runState.duplicateFolderNumbers),
    errors: ensureStringArray_(runState.errors),
    consoleLines: ensureRecentStringArray_(runState.consoleLines),
    limitReached: Boolean(runState.limitReached),
    needsContinuation: Boolean(runState.needsContinuation),
    canResume: Boolean(runState.needsContinuation && runState.limitReached),
    message: String(runState.message || ''),
    startedAt: String(runState.startedAt || ''),
    updatedAt: String(runState.updatedAt || ''),
    diagnosticLogUrl: String(runState.diagnosticLogUrl || ''),
  };
}

function buildRunPhaseMessage_(runState) {
  if (runState.phase === 'describe') {
    return 'Generating formal descriptions from the Survey Sheet.';
  }
  if (runState.phase === 'remove') {
    return 'Removing previously managed survey blocks.';
  }
  return buildProcessingMessage_(runState);
}

function buildProcessingMessage_(runState) {
  const totalCount = Number(runState.totalCount || 0);
  if (!totalCount) {
    return 'Preparing preservation survey batch.';
  }
  const nextIndex = Math.min(runState.completedNumbers.length + 1, totalCount);
  return 'Processing photo ' + nextIndex + ' of ' + totalCount + '.';
}

function buildPauseMessage_(runState, fallbackMessage) {
  if (runState.phase === 'describe') {
    return fallbackMessage + ' Described ' + String(runState.describedCount || 0) + ' row(s) so far.';
  }
  if (runState.phase === 'remove') {
    return fallbackMessage + ' Removed ' + String(runState.removedCount || 0) + ' paragraph(s) so far.';
  }
  return fallbackMessage + ' ' + buildProcessingMessage_(runState);
}

function buildCompletionMessage_(runState) {
  const missingCount = ensureNumberArray_(runState.missingNumbers).length;
  const errorCount = ensureStringArray_(runState.errors).length;
  return [
    'Complete.',
    'Descriptions ' + String(runState.describedCount || 0) + '.',
    'Inserted ' + String(runState.insertedCount || 0) + '.',
    'Skipped ' + String(runState.skippedCount || 0) + '.',
    'Missing ' + String(missingCount) + '.',
    'Errors ' + String(errorCount) + '.',
  ].join(' ');
}

function formatDiagnosticError_(detail) {
  return ['Photo ' + String(detail.number || 'n/a'), 'file "' + String(detail.fileName || 'Unknown file') + '"', 'ID ' + String(detail.fileId || 'n/a'), String(detail.message || 'Unknown error')].join(' | ');
}

function ensureNumberArray_(values) {
  if (!Array.isArray(values)) {
    return [];
  }
  return values.map(function(value) {
    return Number(value);
  }).filter(function(value) {
    return !Number.isNaN(value);
  });
}

function ensureStringArray_(values) {
  if (!Array.isArray(values)) {
    return [];
  }
  return values.map(function(value) {
    return String(value);
  });
}

function ensureRecentStringArray_(values, maxItems) {
  const limit = Number(maxItems || PRES_CONFIG.maxConsoleLines);
  const normalized = ensureStringArray_(values);
  return normalized.slice(Math.max(normalized.length - limit, 0));
}

function ensureObjectArray_(values) {
  if (!Array.isArray(values)) {
    return [];
  }
  return values.filter(function(value) {
    return value && typeof value === 'object';
  });
}

function ensureRecentObjectArray_(values, maxItems) {
  const normalized = ensureObjectArray_(values);
  const limit = Number(maxItems || PRES_CONFIG.maxStoredErrorDetails);
  return normalized.slice(Math.max(normalized.length - limit, 0));
}

function sortedNumberCopy_(values) {
  return sortNumbers_(ensureNumberArray_(values).slice());
}

function appendRunLog_(runState, message) {
  const trimmedMessage = String(message || '').trim();
  if (!trimmedMessage) {
    return;
  }
  runState.consoleLines = ensureRecentStringArray_(runState.consoleLines);
  const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'HH:mm:ss');
  const nextLine = '[' + timestamp + '] ' + trimmedMessage;
  const lastLine = runState.consoleLines.length ? runState.consoleLines[runState.consoleLines.length - 1] : '';
  if (lastLine === nextLine) {
    return;
  }
  runState.consoleLines.push(nextLine);
  runState.consoleLines = ensureRecentStringArray_(runState.consoleLines);
}

function uniqueNumbers_(values) {
  const seen = {};
  const uniqueValues = [];
  for (const value of values) {
    if (seen[value]) {
      continue;
    }
    seen[value] = true;
    uniqueValues.push(value);
  }
  return sortNumbers_(uniqueValues);
}

function pushUniqueNumber_(values, value) {
  if (values.indexOf(value) !== -1) {
    return;
  }
  values.push(value);
}

function sortNumbers_(values) {
  return values.sort(function(left, right) {
    return left - right;
  });
}

function formatValueForDisplay_(value) {
  if (value instanceof Date) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  return String(value || '').trim();
}

function formatDateForFile_(date) {
  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyyMMdd');
}

function sanitizeFileName_(value) {
  return String(value || 'Preservation_Survey').replace(/[^a-zA-Z0-9_.-]+/g, '_');
}
