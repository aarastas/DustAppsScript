# Dust Journal

Dust Journal is a single-user Google Apps Script web app that runs from a spreadsheet-bound Apps Script project. It stores journal entries, special-date rules, and hidden metadata in the spreadsheet and renders the UI from `Index.html`.

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

Use any Google Sheet. No special spreadsheet name is required.

1. Open the Google Sheet you want to use with Dust.
2. In that sheet, go to **Extensions > Apps Script**.
3. Add or replace the project files with `Code.gs`, `Index.html`, and `appsscript.json`.
4. Save the project.
5. Run any function once from the editor to authorize access.
6. After authorization, the script will create the required tabs automatically the first time it runs.

Required tabs and storage are created by the app itself:

- `Dust`
- `SpecialDates`
- `DustMeta`
- `Dust Photos`

If you are starting from a fresh spreadsheet, you can still use the sheet right away. Existing sheets are not renamed or required by name.

## Deploy

1. In Apps Script, choose **Deploy > New deployment**.
2. Select **Web app**.
3. Set **Execute as** to the deployment owner.
4. Set access to **Only myself** for the intended single-user setup.
5. Deploy and open the web app URL.
6. If you are installing this for another person, they should do the same steps inside their own Google Sheet so the project is bound to their spreadsheet.

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

- This project is intentionally single-user.
- There is no dependency on a particular spreadsheet filename or pre-existing tab layout.
- Do not share the backing spreadsheet unless you also revisit the data access assumptions.
- Apps Script HTML Service renders modern HTML/CSS/JS, but it still runs inside Google’s web app environment and sandbox restrictions apply.
