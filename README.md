# ROUGE — Luxury Women's Fashion Store

## V4.2 — Google Sheets Product Connection

This version keeps the ROUGE visual identity and connects the product catalog to the Google Sheets API.

### Product flow

Google Sheet → Google Apps Script API → `rouge-data.js` → ROUGE pages

The website uses the following API endpoint:

`https://script.google.com/macros/s/AKfycbwLnGjb7qiKsGNmkfkxhQGZ7xFC4pus8Ur2NHi9pMM-vdo8MMReg1KfwQq78ranPuUe/exec`

### Google Sheet product columns

The Apps Script API is expected to return products using these fields:

- ID
- Name
- Category
- Price
- Description
- Color
- Sizes
- Stock
- Image1
- Image2
- Image3
- Featured
- Status

Only products with `Status = Active` are returned by the API.

### Adding a real product

Add a row to the `Products` sheet. Example:

| ID | Name | Category | Price | Description | Color | Sizes | Stock | Image1 | Image2 | Image3 | Featured | Status |
|---|---|---|---:|---|---|---|---:|---|---|---|---|---|
| R001 | White Silk Dress | Dresses | 2850 | Sculpted silk silhouette | Ivory | XS,S,M,L | 10 | image URL | image URL | image URL | TRUE | Active |

If Image1/Image2/Image3 are empty, the website uses its existing demo artwork as a fallback.

### Important

The public API should contain catalog data only. Do not put passwords, payment information, private customer data, or admin credentials in the public Products endpoint.

### Local development

Open the folder in VS Code and run `index.html` with Live Server.

## Google Drive image URLs
You can paste a Google Drive sharing URL into Image1/Image2/Image3. The frontend automatically converts common Drive file-share URLs into a Drive thumbnail URL for use in `<img>` elements.

The Drive file must be shared as **Anyone with the link → Viewer** (or otherwise publicly readable). If the file is private, the browser cannot load it from the public ROUGE website.


## Google Sheets connection
This version uses JSONP for the public Google Apps Script product endpoint, which avoids browser cross-origin fetch issues. Update the Apps Script `doGet` and `jsonResponse` functions using the version supplied with this package.


### V4.9 image fitting
Product images use `object-fit: contain` and are constrained to the full product-card frame so the complete image stays inside the box without cropping.


## Homepage collection images — Google Sheets

The three homepage collection-card images are now independent from the hero image and can be changed from Google Sheets.

### One-time setup

1. Open the same Google Apps Script project connected to the ROUGE Products sheet.
2. Replace/update the Apps Script code with the included `GoogleAppsScript.gs`.
3. In Apps Script, run `setupHomepageImagesSheet()` once and authorize it.
4. A new sheet named **Homepage Images** will be created.
5. Deploy the Apps Script as a Web App again using the same deployment URL.

### Homepage Images sheet columns

| Key | Image | Alt | Status |
|---|---|---|---|
| dresses | Google Drive image link | Dresses editorial | Active |
| clothing | Google Drive image link | Clothing editorial | Active |
| accessories | Google Drive image link | Accessories editorial | Active |

You only need to change the **Image** cell. The three homepage cards are independent, so changing `dresses` will not change the hero image, clothing image, or accessories image.

Google Drive images should be shared so the public website can read them: **Anyone with the link → Viewer**.


## New In subpage
The NEW IN navigation item now opens `new-in.html`, a separate catalog page showing only products whose tag is `NEW IN`. COLLECTIONS continues to open `collections.html`.


## V5.8 — Full product layout on every page tab

The page tabs now use the **exact same headers and column order as Products**.

Canonical product headers:

| ID | Title | Name | Category | Price | Description | Color | Sizes | Stock | Image1 | Image2 | Image3 | Featured | Status |
|---|---|---|---|---:|---|---|---|---:|---|---|---|---|---|

The `Title` column is now also part of the master **Products** sheet. Existing products automatically receive their `Name` as the initial `Title` when the setup function is run.

The same full layout is used by:

- **Products**
- **New In**
- **Collections**
- **Dresses**
- **Clothing**
- **Accessories**

### How it works

Each page tab is both a **product-entry tab** and a **page assignment tab**. If you add a complete product row to any page tab, the script automatically copies that row into **Products** using the `ID`.

- New ID → appended to Products.
- Existing ID → the Products row is updated.
- Multi-row paste is supported.
- A product appears on a website page when its row exists in that page's tab.
- The same product can exist in multiple page tabs.
- The row order in a page tab controls the order on that website page.
- `Title` and `Image1` from the page row are used for that page's product card.
- Google Drive sharing links can be used in `Image1`, `Image2`, and `Image3`.

### One-time setup after upgrading

1. Replace the Apps Script code with the included `GoogleAppsScript.gs`.
2. Run **`setupProductPageSheets()`** once.
3. The existing **Products** sheet gets the new `Title` column.
4. All five page tabs are created or converted to the same full product layout.
5. Existing page-tab rows are preserved and converted where possible.
6. Redeploy the Web App as a new version.

For a large paste/import or recovery, you can run **`syncAllPageSheetsToProducts()`** manually.

After setup, you can manage products directly from **New In, Collections, Dresses, Clothing, or Accessories**, and the product will automatically be synchronized into **Products**.

### Google Drive images

Use a Google Drive sharing link such as:

`https://drive.google.com/file/d/FILE_ID/view`

The image file must be shared as **Anyone with the link → Viewer** so the public website can display it.

After changing Apps Script, redeploy using **Deploy → Manage deployments → Edit → New version → Deploy**.


## V6 — Click-to-edit visual editor

You can now click directly on headlines, paragraphs, and background images on the **live site** and edit them in place. No code, no separate admin page.

### What's editable right now

- Home, Dresses, Clothing, Accessories, New In, Collections, About: hero background image, hero kicker/title/subtitle, and the main editorial/story headings, paragraphs, and numbered list lines.
- Product cards and the 3 big Collections/Homepage image tiles are unchanged — keep managing those from the **Products**, **Homepage Images**, and **Collection Page Images** sheets as before.
- Anything else (nav links, footer, buttons) is not yet click-editable — see "Adding more editable spots" below.

### One-time setup

1. Replace the Apps Script project's code with the included `GoogleAppsScript.gs`.
2. In Apps Script, run **`setupSiteContentSheet()`** once and authorize it. This creates a **Site Content** sheet.
3. Run **`setEditPassword('yourpassword')`** once (edit the value first) — this is the password the editor will ask for. Without this step, saving from the site is locked.
4. Redeploy: **Deploy → Manage deployments → Edit → New version → Deploy**.

### How to use it

1. Open any page on the live site with `?edit=1` added to the URL, e.g. `https://yoursite.com/about.html?edit=1`.
2. Every editable block gets a subtle red dashed outline on hover.
3. Click a text block → a small panel opens with a text box → edit → **SAVE**.
4. Click an image block → a panel opens with a field for a Google Drive share link or image URL, with a live preview → **SAVE**.
5. The first time you save, it will ask for the password you set in step 3 (remembered for the rest of that browser session).
6. Changes save straight to the **Site Content** sheet and appear on the live site immediately for every visitor — no redeploy needed.

### Adding more editable spots

Every editable block is just an HTML element with a `data-ck="unique.key.name"` attribute (add `data-ck-type="image"` for images, `data-ck-type="html"` if the text contains `<br>`/`<em>`). If there's a card, word, or image anywhere on the site you want added to the click-editor, tell me which page and element and I'll tag it — it's a one-line change, no rebuild of the system required.


## V7 — Product & Image Manager web app (admin-products.html)

A separate, dedicated page for managing **products** and the **2 sets of big image cards** (Homepage + Collections) — no Google Sheet needed at all for day-to-day use. This is completely separate from the click-to-edit system above; neither one interferes with the other.

### Setup

Nothing extra — it reuses the same Apps Script deployment and the same password you set with `setEditPassword('yourpassword')`.

### How to use it

1. Open `admin-products.html` (through Live Server, or upload it alongside your other pages once hosted).
2. Enter the edit password once.
3. **Products tab:** see every product in a table. Click **+ ADD PRODUCT** to create one, or **EDIT** / **DELETE** on any row. Fields: title, category, price, stock, description, color, sizes, up to 3 images, featured (controls whether it shows on "New In"), and status (Active/Inactive).
   - Which pages a product appears on (New In / Collections / Dresses / Clothing / Accessories) is worked out automatically from **Category** and **Featured** — you don't manage that separately.
4. **Homepage Collection Cards** and **Collections Page Cards** sections: each has 3 cards (Dresses / Clothing / Accessories) with an image field (paste a Drive link or URL), title, alt text, and status — same data these two sheets already held, just with a live preview and a Save button instead of editing sheet cells.

Everything saves straight back to your existing **Products**, **Homepage Images**, and **Collection Page Images** sheets — you can still open the spreadsheet directly any time and it'll match exactly what the admin page shows.


## Orders — checkout & admin.html

`checkout.html` submits every order to the Apps Script backend, and `admin.html` (Order Management) reads and updates them from a dedicated **Orders** sheet.

### One-time setup

1. Make sure the Apps Script project's code matches the included `GoogleAppsScript.gs`.
2. Optionally run `setupOrdersSheet()` once in Apps Script to create the sheet in advance — it isn't required, since the **Orders** sheet is also created automatically the first time a customer places an order.
3. Make sure `setEditPassword('yourpassword')` has been run — `admin.html` uses the same shared edit password as the visual editor and the Product & Image Manager.
4. Redeploy: **Deploy → Manage deployments → Edit → New version → Deploy**.

### Orders sheet columns

| Order Number | Date | Email | Phone | Full Name | Country | City | Postal Code | Address | Payment | Items | Total | Status |
|---|---|---|---|---|---|---|---|---|---|---|---|---|

- Order numbers are generated automatically (`RG-0001`, `RG-0002`, ...).
- `Items` stores the order's line items as JSON (name, size, qty, price).
- `Status` starts as `Pending` and can be changed from `admin.html` (Pending / Confirmed / Shipped / Cancelled).
- The hidden `website` field on the checkout form is a spam honeypot — real customers never fill it in; submissions that do are silently ignored.
