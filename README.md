# Dust Journal

Dust Journal is a single-user Google Apps Script web app backed by one Google Spreadsheet. It stores journal entries, special-date rules, and hidden metadata in the spreadsheet and renders the UI from `Index.html`.

## What’s in this repo

- `Code.gs` - server logic, sheet access, caching, version metadata, and photo handling.
- `Index.html` - the web UI.
- `appsscript.json` - Apps Script manifest and scopes.
- `camera.svg` - source artwork used for the photo button.

Legacy or experimental copies live in the subfolders and are not part of the current Apps Script build.

## Features

- Journal entry capture with date, location, and text.
- “On this day” entry lookup with special-date overrides.
- Special date rules and display modes.
- Photo attachment stored in Google Drive and linked from the sheet.
- Version metadata written to the hidden `DustMeta` sheet.

## Install / Setup

1. Create a new Google Apps Script project.
2. Copy in `Code.gs`, `Index.html`, and `appsscript.json`.
3. Make sure the manifest includes the required scopes for Sheets and Drive.
4. Open the spreadsheet you want Dust to use as the active spreadsheet.
5. Run any function once from the editor to authorize access.

## Deploy

1. In Apps Script, choose **Deploy > New deployment**.
2. Select **Web app**.
3. Set **Execute as** to the deployment owner.
4. Set access to **Only myself** for the intended single-user setup.
5. Deploy and open the web app URL.

## Usage

- Add a journal entry from the `Today` section.
- Use `On this day` to preview matching historical entries.
- Add special dates in `Settings`.
- Use the camera button to attach a photo to an entry.

## Versioning

- `Code.gs` and `Index.html` each carry their own version constant and changelog comment.
- Bump the matching version whenever that file changes.
- The app also writes the current versions into the hidden `DustMeta` sheet.

## Notes

- This project is intentionally single-user. Do not share the backing spreadsheet unless you also revisit the data access assumptions.
- Apps Script HTML Service renders modern HTML/CSS/JS, but it still runs inside Google’s web app environment and sandbox restrictions apply.

