# ROUGE Collections Page — Separate Image Sheet

The Home page and Collections subpage now use separate image sources.

## 1. Update Apps Script

Use the included `GoogleAppsScript.gs` in the Apps Script project connected to your ROUGE spreadsheet.

## 2. Create the Collections image sheet

In Apps Script, run this function once:

`setupCollectionPageImagesSheet()`

It creates a sheet named **Collection Page Images** with:

| Key | Image | Alt | Status |
|---|---|---|---|
| dresses | Google Drive image link | Dresses collection | Active |
| clothing | Google Drive image link | Clothing collection | Active |
| accessories | Google Drive image link | Accessories collection | Active |

Replace the sample image links with your own Google Drive sharing links.

## 3. Deploy the Apps Script

Redeploy the Web App so the new `collectionImages` API action is available. Keep the same Web App URL if possible.

## 4. Website behavior

`collections.html` requests `action=collectionImages` and applies the returned images to the three collection tiles. Google Drive `/file/d/.../view` links are converted automatically by the Apps Script to a usable thumbnail URL.

The existing Home page `Homepage Images` sheet and its API action are unchanged.
