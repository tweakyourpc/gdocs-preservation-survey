# Setup

This project deploys to a container-bound Google Apps Script project with `clasp`.

## First-Time Deployment

1. Install local tooling:

```bash
npm install
```

2. Authenticate `clasp` with the Google account that owns the target Google Doc:

```bash
npx clasp login
```

3. Create a new Google Doc.

4. In the Doc, open `Extensions > Apps Script`.

5. In the Apps Script editor, open `Project Settings` and copy the Script ID.

6. Clone that Apps Script project locally:

```bash
npx clasp clone <scriptId>
```

This creates `.clasp.json`, which points this local folder at that Apps Script project. Do not commit `.clasp.json`; it is deployment-specific and gitignored.

7. Push the local Apps Script files:

```bash
npm run push
```

8. Reload the Google Doc and confirm the `Preservation Survey` menu appears.

9. Run `Preservation Survey > Setup > Link Survey Sheet`, then pick or create the survey Sheet.

10. Run `Preservation Survey > Setup > Link Photo Folder`, then pick the Drive folder with numbered photos.

11. Run `Preservation Survey > Setup > Initialize Sheet Tabs & Dropdowns`.

12. Run `Preservation Survey > Setup > Set Project Info`.

13. Fill Survey rows, mostly from dropdowns.

14. Run `Preservation Survey > Generate > Generate Descriptions from Sheet`.

15. Run `Preservation Survey > Generate > Insert Missing Photos + Captions`.

## Iteration Workflow

Use this loop while testing:

```bash
# edit locally
npm run push
# test in the Google Doc
```

If you make changes in the Apps Script editor and need them locally:

```bash
npm run pull
```

Open the bound Apps Script project:

```bash
npm run open
```

Watch execution logs:

```bash
npm run logs
```

## GitHub Actions Deployment

The workflow in `.github/workflows/push.yml` pushes Apps Script files on changes to `.gs` or `.html` files on `main`.

Create a `CLASP_TOKEN` repository secret before using it:

1. Run `npx clasp login --no-localhost`.

2. Complete the browser authentication flow.

3. Copy the contents of `~/.clasprc.json`.

4. Paste that JSON as the `CLASP_TOKEN` secret value in GitHub.

The workflow also requires a committed or generated `.clasp.json` at runtime. Because `.clasp.json` is intentionally gitignored, adapt the workflow to write one from a repository secret if you use different Apps Script projects per environment.
