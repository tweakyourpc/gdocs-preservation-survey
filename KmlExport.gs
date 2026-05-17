'use strict';

/** Exports non-skipped survey rows with GPS coordinates to a KML file in the photo folder. */
function exportKml() {
  const spreadsheet = getConfiguredSurveySheet_();
  const surveyRead = readSurveyRows_(spreadsheet);
  const projectInfo = getProjectInfo_();
  const placemarks = [];
  for (const row of surveyRead.rows) {
    if (String(row.status || '').toLowerCase() === 'skip') {
      continue;
    }
    if (!row.gpsLat || !row.gpsLon) {
      continue;
    }
    const lat = Number(row.gpsLat);
    const lon = Number(row.gpsLon);
    if (Number.isNaN(lat) || Number.isNaN(lon)) {
      continue;
    }
    placemarks.push(buildKmlPlacemark_(row, lat, lon));
  }
  if (!placemarks.length) {
    throw new Error('No exportable GPS coordinates were found. Add decimal GPS Lat and GPS Lon values on the Survey tab, then export again.');
  }
  const kml = buildKmlDocument_(projectInfo, placemarks);
  const folder = getConfiguredFolder_();
  const fileName = sanitizeFileName_((projectInfo.projectName || 'Preservation_Survey') + '_Survey_' + formatDateForFile_(new Date()) + '.kml');
  folder.createFile(fileName, kml, 'application/vnd.google-earth.kml+xml');
  DocumentApp.getActiveDocument().toast('KML exported: ' + fileName);
  return fileName;
}

function buildKmlDocument_(projectInfo, placemarks) {
  return [
    '<?xml version="1.0" encoding="UTF-8"?>',
    '<kml xmlns="http://www.opengis.net/kml/2.2">',
    '<Document>',
    '<name>' + escapeXml_(projectInfo.projectName || 'Preservation Survey') + '</name>',
    '<Style id="preservationSurveyPin"><IconStyle><color>ff00ffff</color><scale>1.1</scale><Icon><href>http://maps.google.com/mapfiles/kml/pushpin/ylw-pushpin.png</href></Icon></IconStyle></Style>',
    placemarks.join('\n'),
    '</Document>',
    '</kml>',
  ].join('\n');
}

function buildKmlPlacemark_(row, lat, lon) {
  const name = 'Photo ' + row.photoNumber + ' - ' + (row.streetAddress || 'Unaddressed resource');
  const description = [
    row.generatedDescription || '',
    row.gpsLat && row.gpsLon ? 'GPS: ' + row.gpsLat + ', ' + row.gpsLon : '',
    row.recorder ? 'Recorded by: ' + row.recorder : '',
    row.surveyDate ? 'Survey Date: ' + formatValueForDisplay_(row.surveyDate) : '',
  ].filter(Boolean).join('<br>');
  return [
    '<Placemark>',
    '<name>' + escapeXml_(name) + '</name>',
    '<styleUrl>#preservationSurveyPin</styleUrl>',
    '<description><![CDATA[' + description.replace(/]]>/g, ']]]]><![CDATA[>') + ']]></description>',
    '<Point><coordinates>' + String(lon) + ',' + String(lat) + ',0</coordinates></Point>',
    '</Placemark>',
  ].join('');
}

function escapeXml_(value) {
  return String(value || '')
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&apos;');
}
