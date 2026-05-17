# gdocs-preservation-survey

Historic Preservation Field Survey is a Google Apps Script fork of `gdocs-photo-report` for deterministic SHPO-style field survey documentation in Google Docs. It uses a structured Google Sheet, controlled dropdown vocabularies, numbered Drive photos, and resumable Apps Script batches. No AI or LLM calls are used in v1.

## What It Does

- Creates and manages a linked Survey Sheet with reference tabs and dropdown validation.
- Generates formal resource descriptions into the Survey tab from lookup tables in `ReferenceData.gs`.
- Inserts numbered photos and caption blocks into the active Google Doc.
- Preserves resumable batching, LockService protection, document property run state, idempotency markers, and diagnostic log generation from the original project.
- Exports KML placemarks and CSV survey summaries to the linked photo folder.

## Setup

1. Open the Google Doc that will become the survey report.
2. Open `Extensions > Apps Script`.
3. Add these files to the Apps Script project: `Code.gs`, `ReferenceData.gs`, `AssemblyEngine.gs`, `SheetSetup.gs`, `KmlExport.gs`, `Sidebar.html`, `Picker.html`, and `appsscript.json`.
4. Save the script project and reload the Google Doc.
5. Open `Preservation Survey > Open Sidebar`.
6. Link or create a Survey Sheet.
7. Run `Initialize Sheet` to create the Survey tab, reference tabs, named ranges, dropdowns, and the Unmatched log.
8. Link the Drive photo folder that contains numbered files such as `Photo 1.jpg` or `Image 1.jpg`.
9. Fill out Survey rows and run `Generate Descriptions`, then `Insert Missing` or `Rebuild All`.

## Survey Workflow

The Survey tab is the working surface. Surveyors fill columns A through AP. Column AQ is script-written and protected with a warning. Most fields use dropdowns populated from the reference tabs; optional freetext is limited to notes, identifiers, names, dates, and coordinates.

`Generate Descriptions` reads Survey rows, resolves lookup values, logs unmatched values to the Unmatched tab, and writes deterministic formal descriptions to AQ. `Insert Missing` matches `Photo #` to the last number in each image filename and appends managed report blocks to the Doc. `Rebuild All` first removes prior managed survey blocks, then reinserts from the Sheet.

If Apps Script approaches execution limits, the run state is stored in DocumentProperties and the sidebar exposes Resume. Normal soft-batch pauses continue automatically from the sidebar.

## Survey Columns

A `Photo #`; B `Street Address`; C `City`; D `State`; E `Parcel ID / Feature ID`; F `Resource Type`; G `Use - Historic`; H `Use - Current`; I `Architectural Style (Primary)`; J `Architectural Style (Secondary)`; K `Stories`; L `Roof Form`; M `Roof Material`; N `Primary Cladding`; O `Secondary Cladding`; P `Foundation Type`; Q `Window Type`; R `Window Material`; S `Door Type`; T `Porch Type`; U-X `Finding 1-4`; Y `Integrity - Overall`; Z-AF seven NR integrity aspects; AG `NR Criteria`; AH `Period of Significance`; AI `Date of Construction`; AJ `Architect / Builder`; AK `GPS Lat`; AL `GPS Lon`; AM `Recorder`; AN `Survey Date`; AO `Surveyor Notes`; AP `Status`; AQ `Generated Description`.

Rows with Status `Skip` are omitted from Doc insertion and KML export.

## Project Info

Use `Preservation Survey > Setup > Set Project Info` to store project metadata in DocumentProperties under `PRES_PROJECT_INFO`. The report cover block is inserted once and tracked by a managed marker.

## KML Export

`Preservation Survey > Export > Export KML to Drive` reads rows with decimal GPS latitude and longitude, excludes skipped rows, and writes one KML placemark per resource to the linked photo folder. Each placemark includes the AQ description and metadata in a CDATA block.

## Reference Data

`ReferenceData.gs` includes pre-populated arrays for architectural styles, materials, roof forms, window types, door types, porch types, findings, integrity aspects, periods, resource types, use types, foundations, and U.S. jurisdictions. The Sheet setup process writes these into reference tabs only when a tab is blank beyond its header, so local SHPO customizations are preserved on repeated initialization.

## V1 Boundaries

This version does not implement AI generation, photo analysis, real-time GPS mapping, PDF export, email notifications, multi-user conflict resolution beyond LockService, or Gemini Vision. Those are planned v2 candidates only.

## OAuth Scopes

The manifest requests Google Docs, Drive, Spreadsheets, container UI, ScriptApp, and user email scopes because the tool edits the active document, creates and updates Sheets, reads and writes Drive exports, opens Picker UI, and uses ScriptApp OAuth tokens.
