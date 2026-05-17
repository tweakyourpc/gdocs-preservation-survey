'use strict';

function buildPreservationLookups_(spreadsheet) {
  return {
    styles: indexRowsByKey_(PRES_STYLES, 'name'),
    materials: indexRowsByKey_(PRES_MATERIALS, 'term'),
    roofForms: indexRowsByKey_(PRES_ROOF_FORMS, 'term'),
    windowTypes: indexRowsByKey_(PRES_WINDOW_TYPES, 'term'),
    doorTypes: indexRowsByKey_(PRES_DOOR_TYPES, 'term'),
    porchTypes: indexRowsByKey_(PRES_PORCH_TYPES, 'term'),
    foundations: indexRowsByKey_(PRES_FOUNDATIONS, 'term'),
    findings: indexRowsByKey_(PRES_FINDINGS, 'term'),
    resourceTypes: indexRowsByKey_(PRES_RESOURCE_TYPES, 'term'),
    unmatchedSheet: spreadsheet ? getOrCreateUnmatchedSheet_(spreadsheet) : null,
    unmatchedRows: [],
  };
}

function indexRowsByKey_(rows, key) {
  const index = new Map();
  for (const row of rows) {
    index.set(normalizeLookupKey_(row[key]), row);
  }
  return index;
}

function assembleDescription_(surveyRow, lookups) {
  const style = lookupReference_(lookups, 'styles', surveyRow.stylePrimary, surveyRow, 'Architectural Style (Primary)');
  const secondaryStyle = lookupReference_(lookups, 'styles', surveyRow.styleSecondary, surveyRow, 'Architectural Style (Secondary)');
  const roofForm = lookupReference_(lookups, 'roofForms', surveyRow.roofForm, surveyRow, 'Roof Form');
  const roofMaterial = lookupReference_(lookups, 'materials', surveyRow.roofMaterial, surveyRow, 'Roof Material');
  const cladding = lookupReference_(lookups, 'materials', surveyRow.primaryCladding, surveyRow, 'Primary Cladding');
  const secondaryCladding = lookupReference_(lookups, 'materials', surveyRow.secondaryCladding, surveyRow, 'Secondary Cladding');
  const windowType = lookupReference_(lookups, 'windowTypes', surveyRow.windowType, surveyRow, 'Window Type');
  const doorType = lookupReference_(lookups, 'doorTypes', surveyRow.doorType, surveyRow, 'Door Type');
  const porchType = lookupReference_(lookups, 'porchTypes', surveyRow.porchType, surveyRow, 'Porch Type');
  const resourceType = lookupReference_(lookups, 'resourceTypes', surveyRow.resourceType, surveyRow, 'Resource Type');
  const findings = [];
  for (const findingTerm of surveyRow.findings || []) {
    const finding = lookupReference_(lookups, 'findings', findingTerm, surveyRow, 'Finding');
    if (finding) {
      findings.push(finding);
    }
  }

  flushUnmatchedRows_(lookups);

  const headingParts = [];
  if (surveyRow.featureId) {
    headingParts.push(surveyRow.featureId);
  }
  if (surveyRow.streetAddress) {
    headingParts.push(surveyRow.streetAddress);
  }

  const sentences = [];
  const introParts = [];
  if (surveyRow.stories) {
    introParts.push(surveyRow.stories + '-story');
  }
  if (style) {
    introParts.push(style.name);
  }
  if (resourceType) {
    introParts.push(resourceType.term);
  } else if (surveyRow.resourceType) {
    introParts.push(surveyRow.resourceType);
  }
  let intro = introParts.join(' ');
  if (surveyRow.dateOfConstruction) {
    intro += intro ? ', ' + surveyRow.dateOfConstruction : surveyRow.dateOfConstruction;
  }
  appendSentence_(sentences, intro);
  if (style) {
    appendSentence_(sentences, style.descriptionBlock);
  }
  if (secondaryStyle) {
    appendSentence_(sentences, 'Secondary stylistic influence includes ' + secondaryStyle.name + ' features.');
  }
  appendSentence_(sentences, roofForm ? roofForm.formalDescription : '');
  appendSentence_(sentences, roofMaterial ? 'Roofing is ' + roofMaterial.formalDescription + '.' : '');

  let claddingSentence = cladding ? 'Primary wall cladding is ' + cladding.formalDescription : '';
  if (claddingSentence && secondaryCladding) {
    claddingSentence += '; secondary cladding is ' + secondaryCladding.formalDescription;
  }
  appendSentence_(sentences, claddingSentence);
  appendSentence_(sentences, windowType ? windowType.formalDescription + '.' : '');
  appendSentence_(sentences, doorType ? doorType.formalDescription + '.' : '');
  appendSentence_(sentences, porchType ? porchType.formalDescription + '.' : '');

  for (const finding of findings) {
    appendSentence_(sentences, finding.descriptionPhrase);
  }
  appendSentence_(sentences, buildIntegrityStatement_(surveyRow));
  if (surveyRow.integrityOverall) {
    appendSentence_(sentences, 'Overall integrity: ' + surveyRow.integrityOverall + '.');
  }
  if (surveyRow.surveyorNotes) {
    appendSentence_(sentences, surveyRow.surveyorNotes);
  }

  const body = sentences.join(' ').replace(/\s+/g, ' ').trim();
  return [headingParts.join(' - '), body].filter(Boolean).join('\n');
}

function lookupReference_(lookups, mapName, value, surveyRow, fieldName) {
  if (!isRealSelection_(value)) {
    return null;
  }
  const map = lookups[mapName];
  const record = map ? map.get(normalizeLookupKey_(value)) : null;
  if (record) {
    return record;
  }
  addUnmatchedLookup_(lookups, surveyRow, fieldName, value);
  return null;
}

function normalizeLookupKey_(value) {
  return String(value || '').trim().toLowerCase();
}

function addUnmatchedLookup_(lookups, surveyRow, fieldName, value) {
  if (!lookups || !lookups.unmatchedRows) {
    return;
  }
  lookups.unmatchedRows.push([
    surveyRow.photoNumber || '',
    fieldName,
    String(value || ''),
    new Date(),
  ]);
}

function flushUnmatchedRows_(lookups) {
  if (!lookups || !lookups.unmatchedSheet || !lookups.unmatchedRows.length) {
    return;
  }
  const startRow = lookups.unmatchedSheet.getLastRow() + 1;
  lookups.unmatchedSheet.getRange(startRow, 1, lookups.unmatchedRows.length, 4)
      .setValues(lookups.unmatchedRows);
  lookups.unmatchedRows = [];
}

function getOrCreateUnmatchedSheet_(spreadsheet) {
  let sheet = spreadsheet.getSheetByName(PRES_CONFIG.unmatchedTabName);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(PRES_CONFIG.unmatchedTabName);
    sheet.getRange(1, 1, 1, 4).setValues([['Photo #', 'Field Name', 'Value Entered', 'Timestamp']]);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function appendSentence_(sentences, value) {
  const text = String(value || '').trim();
  if (!text) {
    return;
  }
  const normalized = text.charAt(0).toUpperCase() + text.slice(1);
  sentences.push(/[.!?]$/.test(normalized) ? normalized : normalized + '.');
}


function buildIntegrityStatement_(row) {
  const aspects = [
    ['Location', row.integrityLocation],
    ['Design', row.integrityDesign],
    ['Setting', row.integritySetting],
    ['Materials', row.integrityMaterials],
    ['Workmanship', row.integrityWorkmanship],
    ['Feeling', row.integrityFeeling],
    ['Association', row.integrityAssociation],
  ];
  const retained = [];
  const compromised = [];
  const lost = [];
  for (const aspect of aspects) {
    if (aspect[1] === 'Retained') {
      retained.push(aspect[0]);
    } else if (aspect[1] === 'Compromised') {
      compromised.push(aspect[0]);
    } else if (aspect[1] === 'Lost') {
      lost.push(aspect[0]);
    }
  }
  const clauses = [];
  if (retained.length) {
    clauses.push('integrity of ' + formatList_(retained) + ' is retained');
  }
  if (compromised.length) {
    clauses.push('integrity of ' + formatList_(compromised) + ' is compromised');
  }
  if (lost.length) {
    clauses.push('integrity of ' + formatList_(lost) + ' has been lost');
  }
  return clauses.length ? clauses.join('; ') + '.' : '';
}

function formatList_(values) {
  if (!values.length) {
    return '';
  }
  if (values.length === 1) {
    return values[0];
  }
  if (values.length === 2) {
    return values[0] + ' and ' + values[1];
  }
  return values.slice(0, -1).join(', ') + ', and ' + values[values.length - 1];
}
