'use strict';

const PRES_SURVEY_HEADERS = Object.freeze([
  'Photo #', 'Street Address', 'City', 'State', 'Parcel ID / Feature ID', 'Resource Type',
  'Use - Historic', 'Use - Current', 'Architectural Style (Primary)',
  'Architectural Style (Secondary)', 'Stories', 'Roof Form', 'Roof Material', 'Primary Cladding',
  'Secondary Cladding', 'Foundation Type', 'Window Type', 'Window Material', 'Door Type',
  'Porch Type', 'Finding 1', 'Finding 2', 'Finding 3', 'Finding 4', 'Integrity - Overall',
  'Integrity - Location', 'Integrity - Design', 'Integrity - Setting', 'Integrity - Materials',
  'Integrity - Workmanship', 'Integrity - Feeling', 'Integrity - Association', 'NR Criteria',
  'Period of Significance', 'Date of Construction (approx)', 'Architect / Builder', 'GPS Lat',
  'GPS Lon', 'Recorder', 'Survey Date', 'Surveyor Notes', 'Status', 'Generated Description',
]);

const PRES_SURVEY_COLUMN = Object.freeze({
  photoNumber: 1,
  streetAddress: 2,
  city: 3,
  state: 4,
  featureId: 5,
  resourceType: 6,
  useHistoric: 7,
  useCurrent: 8,
  stylePrimary: 9,
  styleSecondary: 10,
  stories: 11,
  roofForm: 12,
  roofMaterial: 13,
  primaryCladding: 14,
  secondaryCladding: 15,
  foundationType: 16,
  windowType: 17,
  windowMaterial: 18,
  doorType: 19,
  porchType: 20,
  finding1: 21,
  finding2: 22,
  finding3: 23,
  finding4: 24,
  integrityOverall: 25,
  integrityLocation: 26,
  integrityDesign: 27,
  integritySetting: 28,
  integrityMaterials: 29,
  integrityWorkmanship: 30,
  integrityFeeling: 31,
  integrityAssociation: 32,
  nrCriteria: 33,
  periodOfSignificance: 34,
  dateOfConstruction: 35,
  architectBuilder: 36,
  gpsLat: 37,
  gpsLon: 38,
  recorder: 39,
  surveyDate: 40,
  surveyorNotes: 41,
  status: 42,
  generatedDescription: 43,
});

/** Initializes the linked survey Sheet tabs, reference data, and dropdown validation. */
function initializeSurveySheet() {
  const spreadsheet = getConfiguredSurveySheet_();
  initializeReferenceTabs_(spreadsheet);
  initializeSurveyTab_(spreadsheet);
  initializeUnmatchedTab_(spreadsheet);
  DocumentApp.getActiveDocument().toast('Survey Sheet initialized: ' + spreadsheet.getName());
  return true;
}

function initializeReferenceTabs_(spreadsheet) {
  writeReferenceTab_(spreadsheet, PRES_CONFIG.stylesTabName,
      ['Name', 'Alt Names', 'Date Range', 'Regions', 'Primary Characteristics', 'Common Materials', 'Description Block', 'NPS Briefs', 'NRHP Style Code'],
      PRES_STYLES.map(function(row) {
        return [row.name, row.altNames.join('; '), row.dateRange, row.regions.join('; '), row.primaryCharacteristics.join('; '), row.commonMaterials.join('; '), row.descriptionBlock, row.npsBriefs.join('; '), row.nrhpStyleCode];
      }), 'Name');
  writeReferenceTab_(spreadsheet, PRES_CONFIG.materialsTabName,
      ['Term', 'Category', 'Formal Description', 'Period', 'Integrity Notes'],
      PRES_MATERIALS.map(function(row) {
        return [row.term, row.category, row.formalDescription, row.period, row.integrityNotes];
      }), 'Term');
  writeReferenceTab_(spreadsheet, PRES_CONFIG.findingsTabName,
      ['Term', 'Code', 'Category', 'Severity', 'Integrity Aspects', 'Description Phrase', 'SOI Reference', 'NPS Brief Reference', 'Recommended Treatment'],
      PRES_FINDINGS.map(function(row) {
        return [row.term, row.code, row.category, row.severity, row.integrityAspects.join('; '), row.descriptionPhrase, row.soiReference, row.npsBriefReference, row.recommendedTreatment];
      }), 'Term');
  writeReferenceTab_(spreadsheet, PRES_CONFIG.jurisdictionsTabName,
      ['State Code', 'State Name', 'SHPO Name', 'SHPO URL', 'Survey Form Name', 'Survey Form Number', 'NRHP Office', 'Notes'],
      PRES_JURISDICTIONS.map(function(row) {
        return [row.stateCode, row.stateName, row.shpoName, row.shpoUrl, row.surveyFormName, row.surveyFormNumber, row.nrhpOffice, row.notes];
      }), 'StateCode');
  writeReferenceTab_(spreadsheet, PRES_CONFIG.roofFormsTabName,
      ['Term', 'Formal Description', 'Common Styles'],
      PRES_ROOF_FORMS.map(function(row) {
        return [row.term, row.formalDescription, row.commonStyles.join('; ')];
      }), 'Term');
  writeReferenceTab_(spreadsheet, PRES_CONFIG.windowTypesTabName,
      ['Term', 'Formal Description', 'Config Code', 'Common Period'],
      PRES_WINDOW_TYPES.map(function(row) {
        return [row.term, row.formalDescription, row.configCode, row.commonPeriod];
      }), 'Term');
  writeReferenceTab_(spreadsheet, PRES_CONFIG.doorTypesTabName,
      ['Term', 'Formal Description'],
      PRES_DOOR_TYPES.map(function(row) {
        return [row.term, row.formalDescription];
      }), 'Term');
  writeReferenceTab_(spreadsheet, PRES_CONFIG.porchTypesTabName,
      ['Term', 'Formal Description'],
      PRES_PORCH_TYPES.map(function(row) {
        return [row.term, row.formalDescription];
      }), 'Term');
  writeReferenceTab_(spreadsheet, PRES_CONFIG.foundationsTabName,
      ['Term', 'Formal Description'],
      PRES_FOUNDATIONS.map(function(row) {
        return [row.term, row.formalDescription];
      }), 'Term');
  writeReferenceTab_(spreadsheet, PRES_CONFIG.resourceTypesTabName,
      ['Term', 'NPS Definition'],
      PRES_RESOURCE_TYPES.map(function(row) {
        return [row.term, row.npsDefinition];
      }), 'Term');
  writeReferenceTab_(spreadsheet, PRES_CONFIG.useTypesTabName,
      ['Term', 'Category', 'NRHP Code'],
      PRES_USE_TYPES.map(function(row) {
        return [row.term, row.category, row.nrhpCode];
      }), 'Term');
  writeReferenceTab_(spreadsheet, PRES_CONFIG.periodsTabName,
      ['Term'], PRES_PERIODS.map(function(row) {
        return [row.term];
      }), 'Term');
}

function writeReferenceTab_(spreadsheet, name, headers, rows, rangeSuffix) {
  let sheet = spreadsheet.getSheetByName(name);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(name);
  }
  if (sheet.getLastRow() <= 1) {
    sheet.clear();
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    if (rows.length) {
      sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
      sheet.getRange(2, 1, rows.length, headers.length).sort(1);
    }
    sheet.setFrozenRows(1);
    sheet.autoResizeColumns(1, Math.min(headers.length, 8));
  }
  const lastRow = Math.max(sheet.getLastRow(), 2);
  ensureNamedRange_(spreadsheet, name + '_' + rangeSuffix, sheet.getRange(2, 1, Math.max(lastRow - 1, 1), 1));
}

function initializeSurveyTab_(spreadsheet) {
  let sheet = spreadsheet.getSheetByName(PRES_CONFIG.surveyTabName);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(PRES_CONFIG.surveyTabName, 0);
  }
  sheet.getRange(1, 1, 1, PRES_SURVEY_HEADERS.length).setValues([PRES_SURVEY_HEADERS]);
  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(2);
  sheet.getRange(1, 1, 1, PRES_SURVEY_HEADERS.length).setFontWeight('bold').setBackground('#e8eaed');
  sheet.getRange(1, 1, Math.max(sheet.getMaxRows(), 100), PRES_SURVEY_HEADERS.length).applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY);
  setSurveyColumnWidths_(sheet);
  applySurveyValidations_(spreadsheet, sheet);
  protectGeneratedDescriptionColumn_(sheet);
}

function applySurveyValidations_(spreadsheet, sheet) {
  const maxRows = Math.max(sheet.getMaxRows() - 1, 200);
  const fromRange = function(name) {
    return SpreadsheetApp.newDataValidation()
        .requireValueInRange(spreadsheet.getRangeByName(name), true)
        .setAllowInvalid(false)
        .build();
  };
  setColumnValidation_(sheet, PRES_SURVEY_COLUMN.state, maxRows, fromRange_(spreadsheet, PRES_CONFIG.jurisdictionsTabName + '_StateCode'));
  setColumnValidation_(sheet, PRES_SURVEY_COLUMN.resourceType, maxRows, fromRange_(spreadsheet, PRES_CONFIG.resourceTypesTabName + '_Term'));
  setColumnValidation_(sheet, PRES_SURVEY_COLUMN.useHistoric, maxRows, fromRange_(spreadsheet, PRES_CONFIG.useTypesTabName + '_Term'));
  setColumnValidation_(sheet, PRES_SURVEY_COLUMN.useCurrent, maxRows, fromRange_(spreadsheet, PRES_CONFIG.useTypesTabName + '_Term'));
  setColumnValidation_(sheet, PRES_SURVEY_COLUMN.stylePrimary, maxRows, fromRange_(spreadsheet, PRES_CONFIG.stylesTabName + '_Name'));
  setColumnValidation_(sheet, PRES_SURVEY_COLUMN.styleSecondary, maxRows, listValidation_(buildValuesWithNone_(PRES_STYLES, 'name')));
  setColumnValidation_(sheet, PRES_SURVEY_COLUMN.stories, maxRows, listValidation_(['1', '1.5', '2', '2.5', '3', '3+', 'Unknown']));
  setColumnValidation_(sheet, PRES_SURVEY_COLUMN.roofForm, maxRows, fromRange_(spreadsheet, PRES_CONFIG.roofFormsTabName + '_Term'));
  setColumnValidation_(sheet, PRES_SURVEY_COLUMN.roofMaterial, maxRows, listValidation_(filterMaterialTermsByCategory_('Roofing')));
  setColumnValidation_(sheet, PRES_SURVEY_COLUMN.primaryCladding, maxRows, listValidation_(filterMaterialTermsByCategory_('Cladding')));
  setColumnValidation_(sheet, PRES_SURVEY_COLUMN.secondaryCladding, maxRows, listValidation_(['(none)'].concat(filterMaterialTermsByCategory_('Cladding'))));
  setColumnValidation_(sheet, PRES_SURVEY_COLUMN.foundationType, maxRows, fromRange_(spreadsheet, PRES_CONFIG.foundationsTabName + '_Term'));
  setColumnValidation_(sheet, PRES_SURVEY_COLUMN.windowType, maxRows, fromRange_(spreadsheet, PRES_CONFIG.windowTypesTabName + '_Term'));
  setColumnValidation_(sheet, PRES_SURVEY_COLUMN.windowMaterial, maxRows, listValidation_(['Wood', 'Steel', 'Aluminum', 'Vinyl', 'Unknown']));
  setColumnValidation_(sheet, PRES_SURVEY_COLUMN.doorType, maxRows, fromRange_(spreadsheet, PRES_CONFIG.doorTypesTabName + '_Term'));
  setColumnValidation_(sheet, PRES_SURVEY_COLUMN.porchType, maxRows, listValidation_(buildValuesWithNone_(PRES_PORCH_TYPES, 'term')));
  setColumnValidation_(sheet, PRES_SURVEY_COLUMN.finding1, maxRows, fromRange_(spreadsheet, PRES_CONFIG.findingsTabName + '_Term'));
  setColumnValidation_(sheet, PRES_SURVEY_COLUMN.finding2, maxRows, listValidation_(buildValuesWithNone_(PRES_FINDINGS, 'term')));
  setColumnValidation_(sheet, PRES_SURVEY_COLUMN.finding3, maxRows, listValidation_(buildValuesWithNone_(PRES_FINDINGS, 'term')));
  setColumnValidation_(sheet, PRES_SURVEY_COLUMN.finding4, maxRows, listValidation_(buildValuesWithNone_(PRES_FINDINGS, 'term')));
  setColumnValidation_(sheet, PRES_SURVEY_COLUMN.integrityOverall, maxRows, listValidation_(['High', 'Moderate-High', 'Moderate', 'Moderate-Low', 'Low']));
  for (const column of [PRES_SURVEY_COLUMN.integrityLocation, PRES_SURVEY_COLUMN.integrityDesign, PRES_SURVEY_COLUMN.integritySetting, PRES_SURVEY_COLUMN.integrityMaterials, PRES_SURVEY_COLUMN.integrityWorkmanship, PRES_SURVEY_COLUMN.integrityFeeling, PRES_SURVEY_COLUMN.integrityAssociation]) {
    setColumnValidation_(sheet, column, maxRows, listValidation_(['Retained', 'Compromised', 'Lost']));
  }
  setColumnValidation_(sheet, PRES_SURVEY_COLUMN.nrCriteria, maxRows, SpreadsheetApp.newDataValidation()
      .requireFormulaSatisfied('=OR(AG2=\"\",REGEXMATCH(AG2,\"^(A|B|C|D)(,\\\\s*(A|B|C|D))*$\"))')
      .setAllowInvalid(false)
      .build());
  setColumnValidation_(sheet, PRES_SURVEY_COLUMN.periodOfSignificance, maxRows, fromRange_(spreadsheet, PRES_CONFIG.periodsTabName + '_Term'));
  setColumnValidation_(sheet, PRES_SURVEY_COLUMN.status, maxRows, listValidation_(['Pending', 'Complete', 'Needs Review', 'Skip']));
}

function fromRange_(spreadsheet, name) {
  return SpreadsheetApp.newDataValidation()
      .requireValueInRange(spreadsheet.getRangeByName(name), true)
      .setAllowInvalid(false)
      .build();
}

function listValidation_(values) {
  return SpreadsheetApp.newDataValidation().requireValueInList(values, true).setAllowInvalid(false).build();
}

function setColumnValidation_(sheet, column, rowCount, validation) {
  sheet.getRange(2, column, rowCount, 1).setDataValidation(validation);
}

function buildValuesWithNone_(rows, key) {
  return ['(none)'].concat(rows.map(function(row) {
    return row[key];
  }));
}

function filterMaterialTermsByCategory_(prefix) {
  return PRES_MATERIALS.filter(function(row) {
    return row.category.indexOf(prefix) === 0;
  }).map(function(row) {
    return row.term;
  });
}

function protectGeneratedDescriptionColumn_(sheet) {
  const protections = sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
  for (const protection of protections) {
    if (protection.getRange().getA1Notation() === 'AQ:AQ') {
      return;
    }
  }
  const protection = sheet.getRange('AQ:AQ').protect();
  protection.setDescription('This column is written by the script. Do not edit manually.');
  protection.setWarningOnly(true);
}

function setSurveyColumnWidths_(sheet) {
  sheet.setColumnWidth(PRES_SURVEY_COLUMN.photoNumber, 60);
  sheet.setColumnWidth(PRES_SURVEY_COLUMN.streetAddress, 200);
  sheet.setColumnWidths(3, 38, 140);
  sheet.setColumnWidth(PRES_SURVEY_COLUMN.surveyorNotes, 250);
  sheet.setColumnWidth(PRES_SURVEY_COLUMN.generatedDescription, 300);
}

function initializeUnmatchedTab_(spreadsheet) {
  let sheet = spreadsheet.getSheetByName(PRES_CONFIG.unmatchedTabName);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(PRES_CONFIG.unmatchedTabName);
  }
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, 4).setValues([['Photo #', 'Field Name', 'Value Entered', 'Timestamp']]);
  }
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, 4).setFontWeight('bold').setBackground('#fce8e6');
}

function ensureNamedRange_(spreadsheet, name, range) {
  const existing = spreadsheet.getRangeByName(name);
  if (existing) {
    spreadsheet.removeNamedRange(name);
  }
  spreadsheet.setNamedRange(name, range);
}
