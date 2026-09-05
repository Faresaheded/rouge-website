const CONFIG = {
  SHEET_NAME: 'Products',
  HOMEPAGE_IMAGES_SHEET_NAME: 'Homepage Images',
  COLLECTIONS_IMAGES_SHEET_NAME: 'Collection Page Images',
  CONTENT_SHEET_NAME: 'Site Content',
  ORDERS_SHEET_NAME: 'Orders',
  PAGE_SHEETS: {
    'New In': 'new-in',
    'Collections': 'collections',
    'Dresses': 'dresses',
    'Clothing': 'clothing',
    'Accessories': 'accessories'
  }
};

function doGet(e) {
  try {
    const action = (e && e.parameter && e.parameter.action) || 'products';
    const callback = sanitizeCallback((e && e.parameter && e.parameter.callback) || '');
    let result;

    if (action === 'products') result = getProducts((e && e.parameter) || {});
    else if (action === 'product') result = getProduct((e && e.parameter && e.parameter.id) || '');
    else if (action === 'categories') result = getCategories();
    else if (action === 'homepageImages') result = getHomepageImages();
    else if (action === 'collectionImages') result = getCollectionImages();
    else if (action === 'siteContent') result = getSiteContent();
    else if (action === 'updateContent') result = updateSiteContent(e.parameter);
    else if (action === 'saveProduct') result = saveProduct(e.parameter);
    else if (action === 'deleteProduct') result = deleteProduct(e.parameter);
    else if (action === 'saveHomepageImage') result = saveKeyedImageRow_(CONFIG.HOMEPAGE_IMAGES_SHEET_NAME, e.parameter);
    else if (action === 'saveCollectionImage') result = saveKeyedImageRow_(CONFIG.COLLECTIONS_IMAGES_SHEET_NAME, e.parameter);
    else if (action === 'passwordStatus') result = { success: true, passwordIsSet: !!PropertiesService.getScriptProperties().getProperty('EDIT_PASSWORD') };
    else if (action === 'createOrder') result = createOrder((e && e.parameter) || {});
    else if (action === 'getOrders') result = getOrders((e && e.parameter) || {});
    else if (action === 'updateOrderStatus') result = updateOrderStatus((e && e.parameter) || {});
    else if (action === 'submitContact') result = submitContact((e && e.parameter) || {});
    else if (action === 'subscribeNewsletter') result = subscribeNewsletter((e && e.parameter) || {});
    else result = { success: false, error: 'Unknown action' };

    return jsonResponse(result, callback);
  } catch (error) {
    return jsonResponse({ success: false, error: String(error && error.message || error) },
      sanitizeCallback((e && e.parameter && e.parameter.callback) || ''));
  }
}

/**
 * Handles saves made from the ROUGE visual editor (?edit=1 on the live site).
 * Body is sent as text/plain JSON on purpose, so the browser never sends a
 * CORS preflight (Apps Script Web Apps cannot answer OPTIONS requests).
 */
function doPost(e) {
  try {
    const body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    if (body.action === 'updateContent') return jsonResponse(updateSiteContent(body));
    if (body.action === 'saveProduct') return jsonResponse(saveProduct(body));
    if (body.action === 'deleteProduct') return jsonResponse(deleteProduct(body));
    if (body.action === 'saveHomepageImage') return jsonResponse(saveKeyedImageRow_(CONFIG.HOMEPAGE_IMAGES_SHEET_NAME, body));
    if (body.action === 'saveCollectionImage') return jsonResponse(saveKeyedImageRow_(CONFIG.COLLECTIONS_IMAGES_SHEET_NAME, body));
    if (body.action === 'createOrder') return jsonResponse(createOrder(body));
    if (body.action === 'getOrders') return jsonResponse(getOrders(body));
    if (body.action === 'updateOrderStatus') return jsonResponse(updateOrderStatus(body));
    if (body.action === 'submitContact') return jsonResponse(submitContact(body));
    if (body.action === 'subscribeNewsletter') return jsonResponse(subscribeNewsletter(body));
    return jsonResponse({ success: false, error: 'Unknown action' });
  } catch (error) {
    return jsonResponse({ success: false, error: String(error && error.message || error) });
  }
}

/**
 * Very lightweight shared-secret check. Set your password once by running
 * setEditPassword('yourpassword') from the Apps Script editor. This keeps the
 * password out of any file that ships to the browser.
 */
function checkEditPassword_(password) {
  const stored = PropertiesService.getScriptProperties().getProperty('EDIT_PASSWORD');
  if (!stored) return false; // editing is locked until a password is set
  return String(password || '').trim() === String(stored).trim();
}

function setEditPassword(newPassword) {
  PropertiesService.getScriptProperties().setProperty('EDIT_PASSWORD', String(newPassword || '').trim());
  return 'Edit password set. You can now use ?edit=1 on the live site.';
}

/**
 * Site Content sheet — every editable headline/paragraph/image on every page
 * lives here as one row per data-ck key.
 * Columns: Key | Type | Value | Page | Status
 */
function getSiteContent() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.CONTENT_SHEET_NAME);
  if (!sheet) {
    return {
      success: false,
      error: `Sheet "${CONFIG.CONTENT_SHEET_NAME}" was not found. Run setupSiteContentSheet() once.`
    };
  }

  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return { success: true, content: {} };

  const headers = values[0].map(h => String(h).trim());
  const keyIndex = headers.indexOf('Key');
  const typeIndex = headers.indexOf('Type');
  const valueIndex = headers.indexOf('Value');
  const statusIndex = headers.indexOf('Status');
  if (keyIndex === -1 || valueIndex === -1) {
    throw new Error('Site Content sheet must contain Key and Value columns.');
  }

  const content = {};
  values.slice(1).forEach(row => {
    const key = String(row[keyIndex] || '').trim();
    const status = statusIndex >= 0 ? String(row[statusIndex] || 'Active').trim().toLowerCase() : 'active';
    if (!key || status === 'inactive') return;
    const type = typeIndex >= 0 ? String(row[typeIndex] || 'text').trim().toLowerCase() : 'text';
    let value = String(row[valueIndex] || '').trim();
    if (type === 'image') value = normalizeHomepageImageUrl(value);
    content[key] = { value, type };
  });

  return { success: true, content };
}

/**
 * Upserts a single Site Content row. Called by the visual editor.
 */
function updateSiteContent(body) {
  if (!checkEditPassword_(body.password)) {
    return { success: false, error: 'Invalid password' };
  }
  const key = String(body.key || '').trim();
  if (!key) return { success: false, error: 'Missing key' };

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CONFIG.CONTENT_SHEET_NAME);
  if (!sheet) {
    setupSiteContentSheet();
    sheet = ss.getSheetByName(CONFIG.CONTENT_SHEET_NAME);
  }

  const values = sheet.getDataRange().getValues();
  const headers = values[0].map(h => String(h).trim());
  const keyIndex = headers.indexOf('Key');
  const typeIndex = headers.indexOf('Type');
  const valueIndex = headers.indexOf('Value');
  const pageIndex = headers.indexOf('Page');
  const statusIndex = headers.indexOf('Status');

  let targetRow = -1;
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][keyIndex] || '').trim() === key) { targetRow = i + 1; break; }
  }

  const rowData = [];
  rowData[keyIndex] = key;
  rowData[typeIndex] = String(body.type || 'text');
  rowData[valueIndex] = String(body.value || '');
  if (pageIndex >= 0) rowData[pageIndex] = String(body.page || '');
  if (statusIndex >= 0) rowData[statusIndex] = 'Active';

  if (targetRow === -1) {
    sheet.getRange(sheet.getLastRow() + 1, 1, 1, headers.length).setValues([rowData]);
  } else {
    sheet.getRange(targetRow, 1, 1, headers.length).setValues([rowData]);
  }

  return { success: true };
}

/**
 * Run once from the Apps Script editor to create the Site Content sheet.
 */
function setupSiteContentSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CONFIG.CONTENT_SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(CONFIG.CONTENT_SHEET_NAME);

  const headers = ['Key', 'Type', 'Value', 'Page', 'Status'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  sheet.setFrozenRows(1);
  sheet.setColumnWidth(1, 260);
  sheet.setColumnWidth(3, 420);
  return `Ready: ${CONFIG.CONTENT_SHEET_NAME}. Rows are created automatically the first time you save an edit from the site (?edit=1). Remember to also run setEditPassword('yourpassword') once.`;
}

function sanitizeCallback(name) {
  // Only allow a JavaScript identifier/dotted identifier so the JSONP endpoint
  // cannot be used to inject arbitrary script.
  return /^[A-Za-z_$][A-Za-z0-9_$]*(\.[A-Za-z_$][A-Za-z0-9_$]*)*$/.test(name) ? name : '';
}

/* =========================================================================
 * PRODUCT MANAGER — used by admin-products.html
 * ========================================================================= */

/**
 * Creates or updates one product from the admin web app, then re-syncs which
 * page tabs it belongs to (based on Category + Featured), so the admin form
 * is the single place that controls both the product data AND where it shows
 * up on the site.
 */
function saveProduct(p) {
  if (!checkEditPassword_(p.password)) return { success: false, error: 'Invalid password' };

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) throw new Error(`Sheet "${CONFIG.SHEET_NAME}" was not found.`);
  ensureProductsTitleColumn(sheet);

  let id = String(p.id || '').trim();
  if (!id) id = generateNextProductId_(sheet);

  const row = [
    id,
    String(p.title || p.name || '').trim(),
    String(p.name || p.title || '').trim(),
    String(p.category || '').trim(),
    Number(p.price || 0),
    String(p.description || '').trim(),
    String(p.color || '').trim(),
    String(p.sizes || '').trim(),
    Number(p.stock || 0),
    String(p.image1 || '').trim(),
    String(p.image2 || '').trim(),
    String(p.image3 || '').trim(),
    p.featured && String(p.featured).toLowerCase() === 'true' ? 'TRUE' : 'FALSE',
    String(p.status || 'Active').trim()
  ];

  const values = sheet.getDataRange().getValues();
  let targetRow = -1;
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][0] || '').trim() === id) { targetRow = i + 1; break; }
  }
  if (targetRow === -1) {
    sheet.getRange(sheet.getLastRow() + 1, 1, 1, PRODUCT_HEADERS.length).setValues([row]);
  } else {
    sheet.getRange(targetRow, 1, 1, PRODUCT_HEADERS.length).setValues([row]);
  }
  formatProductSheet(sheet);

  syncProductToPageTabs_(row);

  return { success: true, id };
}

/**
 * Removes a product from Products and from every page tab it appears in.
 */
function deleteProduct(body) {
  if (!checkEditPassword_(body.password)) return { success: false, error: 'Invalid password' };
  const id = String(body.id || '').trim();
  if (!id) return { success: false, error: 'Missing id' };

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const productsSheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  removeRowById_(productsSheet, id);

  Object.keys(CONFIG.PAGE_SHEETS).forEach(sheetName => {
    removeRowById_(ss.getSheetByName(sheetName), id);
  });

  return { success: true };
}

function removeRowById_(sheet, id) {
  if (!sheet || sheet.getLastRow() < 2) return;
  const values = sheet.getDataRange().getValues();
  for (let i = values.length - 1; i >= 1; i--) {
    if (String(values[i][0] || '').trim() === id) {
      sheet.deleteRow(i + 1);
    }
  }
}

/**
 * Same page-assignment rules used by setupProductPageSheets(), applied to a
 * single product row so admin-created/edited products land on the right
 * pages automatically.
 */
function syncProductToPageTabs_(row) {
  const id = String(row[0] || '').trim();
  const category = String(row[3] || '').trim().toLowerCase();
  const featured = String(row[12] || '').trim().toLowerCase() === 'true';

  const wants = {
    'New In': featured,
    'Collections': true,
    'Dresses': ['dresses', 'evening'].includes(category),
    'Clothing': category === 'clothing',
    'Accessories': category === 'accessories'
  };

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  Object.keys(CONFIG.PAGE_SHEETS).forEach(sheetName => {
    let sheet = ss.getSheetByName(sheetName);
    if (!sheet) sheet = ss.insertSheet(sheetName);
    ensurePageSheetLayout(sheet);

    const values = sheet.getDataRange().getValues();
    let targetRow = -1;
    for (let i = 1; i < values.length; i++) {
      if (String(values[i][0] || '').trim() === id) { targetRow = i + 1; break; }
    }

    if (wants[sheetName]) {
      if (targetRow === -1) {
        sheet.getRange(sheet.getLastRow() + 1, 1, 1, PRODUCT_HEADERS.length).setValues([row]);
      } else {
        sheet.getRange(targetRow, 1, 1, PRODUCT_HEADERS.length).setValues([row]);
      }
    } else if (targetRow !== -1) {
      sheet.deleteRow(targetRow);
    }
    formatProductSheet(sheet);
  });
}

function generateNextProductId_(sheet) {
  const values = sheet.getDataRange().getValues();
  let max = 0;
  values.slice(1).forEach(row => {
    const m = String(row[0] || '').match(/(\d+)/);
    if (m) max = Math.max(max, Number(m[1]));
  });
  return 'R' + String(max + 1).padStart(3, '0');
}

/* =========================================================================
 * ORDERS — used by checkout.html (createOrder) and admin.html (getOrders,
 * updateOrderStatus).
 * Sheet name: Orders
 * Columns: Order Number | Date | Email | Phone | Full Name | Country | City
 *          | Postal Code | Address | Payment | Items | Total | Status
 * ========================================================================= */

const ORDER_HEADERS = [
  'Order Number', 'Date', 'Email', 'Phone', 'Full Name', 'Country', 'City',
  'Postal Code', 'Address', 'Payment', 'Items', 'Subtotal', 'Shipping', 'Total',
  'Status', 'Order Token', 'Tracking Number', 'Tracking URL'
];

/**
 * Run once from the Apps Script editor (optional — createOrder() will also
 * create this sheet automatically the first time an order comes in).
 */
function setupOrdersSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CONFIG.ORDERS_SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(CONFIG.ORDERS_SHEET_NAME);
  ensureOrdersSheetLayout_(sheet);
  return `Ready: ${CONFIG.ORDERS_SHEET_NAME}`;
}

function ensureOrdersSheetLayout_(sheet) {
  const existingLastCol = Math.max(sheet.getLastColumn(), 1);
  const existingHeaders = sheet.getRange(1, 1, 1, existingLastCol).getValues()[0].map(h => String(h).trim());
  if (existingHeaders.every(h => !h)) {
    sheet.getRange(1, 1, 1, ORDER_HEADERS.length).setValues([ORDER_HEADERS]);
  } else {
    // Preserve existing data and add any new operational columns by name.
    ORDER_HEADERS.forEach((header, i) => {
      if (existingHeaders.indexOf(header) === -1) {
        sheet.getRange(1, sheet.getLastColumn() + 1).setValue(header);
      }
    });
    const finalHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => String(h).trim());
    // Keep the canonical columns first; move legacy columns only when necessary.
    ORDER_HEADERS.forEach((header, targetIndex) => {
      const currentIndex = finalHeaders.indexOf(header);
      if (currentIndex === -1) return;
      if (currentIndex !== targetIndex) {
        sheet.moveColumns(sheet.getRange(1, currentIndex + 1, sheet.getMaxRows(), 1), targetIndex + 1);
        finalHeaders.splice(currentIndex, 1);
        finalHeaders.splice(targetIndex, 0, header);
      }
    });
  }
  sheet.getRange(1, 1, 1, ORDER_HEADERS.length).setValues([ORDER_HEADERS]);
  sheet.getRange(1, 1, 1, ORDER_HEADERS.length).setFontWeight('bold');
  sheet.setFrozenRows(1);
  sheet.setColumnWidth(9, 260);
  sheet.setColumnWidth(11, 320);
}

function getOrdersSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CONFIG.ORDERS_SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(CONFIG.ORDERS_SHEET_NAME);
  ensureOrdersSheetLayout_(sheet);
  return sheet;
}

function generateNextOrderNumber_(sheet) {
  const values = sheet.getDataRange().getValues();
  let max = 0;
  values.slice(1).forEach(row => {
    const m = String(row[0] || '').match(/(\d+)/);
    if (m) max = Math.max(max, Number(m[1]));
  });
  return 'RG-' + String(max + 1).padStart(4, '0');
}

function getMasterProductsForOrder_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) throw new Error(`Sheet "${CONFIG.SHEET_NAME}" was not found.`);
  ensureProductsTitleColumn(sheet);
  const values = sheet.getDataRange().getValues();
  const headers = values[0].map(h => String(h).trim());
  const idx = {};
  PRODUCT_HEADERS.forEach(h => idx[h] = headers.indexOf(h));
  const map = {};
  values.slice(1).forEach(row => {
    const id = String(row[idx.ID] || '').trim();
    if (!id) return;
    const status = String(row[idx.Status] || 'Active').trim().toLowerCase();
    map[id.toLowerCase()] = {
      id,
      externalId: id,
      name: String(row[idx.Name] || row[idx.Title] || '').trim(),
      price: Number(row[idx.Price] || 0),
      stock: Number(row[idx.Stock] || 0),
      sizes: String(row[idx.Sizes] || '').split(',').map(x => x.trim()).filter(Boolean),
      status
    };
  });
  return { sheet, idx, map };
}

function sendCustomerEmail_(to, subject, htmlBody) {
  if (!to) return false;
  try {
    MailApp.sendEmail({to: String(to).trim(), subject, htmlBody, body: htmlBody.replace(/<[^>]+>/g, ' ')});
    return true;
  } catch (e) {
    console.warn('Customer email failed: ' + e);
    return false;
  }
}

function notifyOrderCreated_(order) {
  const lines = (order.items || []).map(it => `${it.name} × ${it.qty} — EGP ${Number(it.price * it.qty).toLocaleString()}`).join('<br>');
  const body = `<div style="font-family:Arial,sans-serif;line-height:1.7"><h2>ROUGE — Order ${order.orderNumber}</h2><p>Thank you, ${escapeEmailHtml_(order.fullName)}. Your order has been received and is pending confirmation.</p><p>${lines}</p><p><strong>Total: EGP ${Number(order.total).toLocaleString()}</strong></p><p>We will contact you to confirm delivery.</p></div>`;
  return sendCustomerEmail_(order.email, `ROUGE Order ${order.orderNumber} — Received`, body);
}

function notifyOrderStatus_(order, status) {
  const copy = {
    Confirmed: 'Your ROUGE order has been confirmed and is being prepared.',
    Shipped: 'Your ROUGE order has shipped.',
    Cancelled: 'Your ROUGE order has been cancelled. Please contact Client Services if you need assistance.'
  }[status];
  if (!copy) return false;
  let tracking = '';
  if (status === 'Shipped' && (order.trackingNumber || order.trackingUrl)) {
    tracking = `<p>Tracking: ${escapeEmailHtml_(order.trackingNumber || '')}${order.trackingUrl ? ` — <a href="${escapeAttribute_(order.trackingUrl)}">Track shipment</a>` : ''}</p>`;
  }
  const body = `<div style="font-family:Arial,sans-serif;line-height:1.7"><h2>ROUGE — Order ${order.orderNumber}</h2><p>${copy}</p>${tracking}<p>Thank you for choosing ROUGE.</p></div>`;
  return sendCustomerEmail_(order.email, `ROUGE Order ${order.orderNumber} — ${status}`, body);
}

function escapeEmailHtml_(v) { return String(v || '').replace(/[&<>]/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;'}[c])); }
function escapeAttribute_(v) { return String(v || '').replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c])); }

function setAdminEmail(email) {
  PropertiesService.getScriptProperties().setProperty('ADMIN_EMAIL', String(email || '').trim());
  return 'Admin notification email saved.';
}

function getAdminEmail_() {
  return String(PropertiesService.getScriptProperties().getProperty('ADMIN_EMAIL') || Session.getEffectiveUser().getEmail() || '').trim();
}

function notifyAdminNewOrder_(order) {
  const to = getAdminEmail_();
  if (!to) return false;
  const lines = (order.items || []).map(it => `${it.name} × ${it.qty} / ${it.size}`).join('<br>');
  return sendCustomerEmail_(to, `ROUGE New Order ${order.orderNumber}`, `<div style="font-family:Arial,sans-serif;line-height:1.7"><h2>New ROUGE order</h2><p><strong>${escapeEmailHtml_(order.orderNumber)}</strong> — ${escapeEmailHtml_(order.fullName)}</p><p>${escapeEmailHtml_(order.email)} / ${escapeEmailHtml_(order.phone)}</p><p>${escapeEmailHtml_(order.address)}, ${escapeEmailHtml_(order.city)}, ${escapeEmailHtml_(order.country)}</p><p>${lines}</p><p><strong>Total: EGP ${Number(order.total).toLocaleString()}</strong></p></div>`);
}

/**
 * Creates an order using server-side product prices and inventory. The browser
 * sends product IDs + quantities only; totals and stock are authoritative here.
 */
function createOrder(body) {
  if (String(body.website || '').trim()) return { success: true, orderNumber: 'RG-0000', date: new Date().toISOString() };

  const email = String(body.email || '').trim();
  const fullName = String(body.fullName || '').trim();
  const address = String(body.address || '').trim();
  const city = String(body.city || '').trim();
  const phone = String(body.phone || '').trim();
  const country = String(body.country || '').trim();
  const payment = String(body.payment || '').trim();
  const orderToken = String(body.orderToken || '').trim();
  if (!email || !fullName || !address || !city || !phone || !country) return { success: false, error: 'Please complete all required delivery details.' };
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) return { success: false, error: 'Please enter a valid email address.' };
  if (payment !== 'Cash on Delivery') return { success: false, error: 'This checkout currently supports Cash on Delivery only.' };
  if (!orderToken) return { success: false, error: 'Missing order token. Please refresh checkout and try again.' };

  let requested = [];
  try { requested = JSON.parse(body.items || '[]'); } catch (e) { requested = []; }
  if (!Array.isArray(requested) || !requested.length) return { success: false, error: 'Your bag is empty.' };

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const sheet = getOrdersSheet_();
    const values = sheet.getDataRange().getValues();
    const headers = values[0].map(h => String(h).trim());
    const orderTokenIndex = headers.indexOf('Order Token');
    if (orderTokenIndex >= 0) {
      const duplicate = values.slice(1).find(r => String(r[orderTokenIndex] || '').trim() === orderToken);
      if (duplicate) return { success: true, orderNumber: String(duplicate[0]), date: duplicate[1] instanceof Date ? duplicate[1].toISOString() : String(duplicate[1] || '') };
    }

    const catalog = getMasterProductsForOrder_();
    const validated = [];
    let subtotal = 0;
    const stockUpdates = [];
    for (const raw of requested) {
      const id = String(raw.id || raw.externalId || '').trim().toLowerCase();
      const product = catalog.map[id];
      const qty = Math.floor(Number(raw.qty || 0));
      const size = String(raw.size || '').trim();
      if (!product || product.status !== 'active') return { success: false, error: `One of the selected products is no longer available.` };
      if (!Number.isFinite(qty) || qty < 1 || qty > 50) return { success: false, error: `Invalid quantity for ${product.name}.` };
      if (product.stock < qty) return { success: false, error: `${product.name} is no longer available in the requested quantity.` };
      if (product.sizes.length && size && !product.sizes.includes(size)) return { success: false, error: `${product.name} does not have size ${size}.` };
      const lineTotal = product.price * qty;
      subtotal += lineTotal;
      validated.push({ id: product.id, name: product.name, size: size || (product.sizes[0] || 'ONE SIZE'), qty, price: product.price });
      stockUpdates.push({ row: catalog.sheet.getDataRange().getValues().findIndex(r => String(r[catalog.idx.ID] || '').trim() === product.id) + 1, newStock: product.stock - qty });
    }

    const shipping = 0; // Set the live shipping rule here before launch.
    const total = subtotal + shipping;
    const orderNumber = generateNextOrderNumber_(sheet);
    const date = new Date();
    const row = [orderNumber, date.toISOString(), email, phone, fullName, country, city, String(body.postalCode || '').trim(), address, payment, JSON.stringify(validated), subtotal, shipping, total, 'Pending', orderToken, '', ''];
    sheet.getRange(sheet.getLastRow() + 1, 1, 1, ORDER_HEADERS.length).setValues([row]);

    // Decrement inventory atomically with the order creation.
    stockUpdates.forEach(u => { if (u.row > 1) catalog.sheet.getRange(u.row, catalog.idx.Stock + 1).setValue(u.newStock); });

    const order = { orderNumber, date: date.toISOString(), email, phone, fullName, country, city, postalCode: String(body.postalCode || '').trim(), address, payment, items: validated, subtotal, shipping, total, status: 'Pending', trackingNumber: '', trackingUrl: '' };
    notifyOrderCreated_(order);
    notifyAdminNewOrder_(order);
    return { success: true, orderNumber, date: date.toISOString() };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Returns every order, newest first. Requires the shared edit password
 * (same one used by the visual editor and the admin product manager).
 */
function getOrders(params) {
  if (!checkEditPassword_(params.password)) return { success: false, error: 'Invalid password' };
  const sheet = getOrdersSheet_();
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return { success: true, orders: [] };
  const headers = values[0].map(h => String(h).trim());
  const idx = {}; headers.forEach((h, i) => idx[h] = i);
  const orders = values.slice(1).map(row => {
    let items = []; try { items = JSON.parse(row[idx['Items']] || '[]'); } catch (e) {}
    return {
      orderNumber: String(row[idx['Order Number']] || '').trim(),
      date: row[idx['Date']] instanceof Date ? row[idx['Date']].toISOString() : String(row[idx['Date']] || ''),
      email: String(row[idx['Email']] || '').trim(), phone: String(row[idx['Phone']] || '').trim(),
      fullName: String(row[idx['Full Name']] || '').trim(), country: String(row[idx['Country']] || '').trim(),
      city: String(row[idx['City']] || '').trim(), postalCode: String(row[idx['Postal Code']] || '').trim(),
      address: String(row[idx['Address']] || '').trim(), payment: String(row[idx['Payment']] || '').trim(),
      items, subtotal: Number(row[idx['Subtotal']] || row[idx['Total']] || 0), shipping: Number(row[idx['Shipping']] || 0),
      total: Number(row[idx['Total']] || 0), status: String(row[idx['Status']] || 'Pending').trim(),
      trackingNumber: String(row[idx['Tracking Number']] || '').trim(), trackingUrl: String(row[idx['Tracking URL']] || '').trim()
    };
  }).filter(o => o.orderNumber);
  orders.reverse();
  return { success: true, orders };
}

function updateOrderStatus(params) {
  if (!checkEditPassword_(params.password)) return { success: false, error: 'Invalid password' };
  const orderNumber = String(params.orderNumber || '').trim();
  const status = String(params.status || '').trim();
  const allowed = ['Pending', 'Confirmed', 'Shipped', 'Cancelled'];
  if (!orderNumber || !allowed.includes(status)) return { success: false, error: 'Invalid order status.' };

  const sheet = getOrdersSheet_();
  const values = sheet.getDataRange().getValues();
  const headers = values[0].map(h => String(h).trim());
  const idx = {}; headers.forEach((h, i) => idx[h] = i);
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][idx['Order Number']] || '').trim() === orderNumber) {
      const oldStatus = String(values[i][idx['Status']] || 'Pending').trim();
      sheet.getRange(i + 1, idx['Status'] + 1).setValue(status);
      const trackingNumber = String(params.trackingNumber || values[i][idx['Tracking Number']] || '').trim();
      const trackingUrl = String(params.trackingUrl || values[i][idx['Tracking URL']] || '').trim();
      if (idx['Tracking Number'] >= 0) sheet.getRange(i + 1, idx['Tracking Number'] + 1).setValue(trackingNumber);
      if (idx['Tracking URL'] >= 0) sheet.getRange(i + 1, idx['Tracking URL'] + 1).setValue(trackingUrl);
      const order = {
        orderNumber, email: String(values[i][idx['Email']] || '').trim(), fullName: String(values[i][idx['Full Name']] || '').trim(),
        items: (() => { try { return JSON.parse(values[i][idx['Items']] || '[]'); } catch (e) { return []; } })(),
        total: Number(values[i][idx['Total']] || 0), trackingNumber, trackingUrl
      };
      if (oldStatus !== status) notifyOrderStatus_(order, status);
      return { success: true };
    }
  }
  return { success: false, error: 'Order not found.' };
}

/* =========================================================================
 * CLIENT SERVICES + NEWSLETTER
 * ========================================================================= */

const CONTACT_HEADERS = ['Date', 'Name', 'Email', 'Subject', 'Message', 'Status'];
const NEWSLETTER_HEADERS = ['Date', 'Email', 'Source', 'Status'];

function ensureSimpleSheet_(name, headers) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  const current = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), headers.length)).getValues()[0].map(h => String(h).trim());
  headers.forEach((h, i) => { if (current.indexOf(h) === -1) sheet.getRange(1, sheet.getLastColumn() + 1).setValue(h); });
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold'); sheet.setFrozenRows(1);
  return sheet;
}

function submitContact(body) {
  const name = String(body.name || '').trim(), email = String(body.email || '').trim(), subject = String(body.subject || '').trim(), message = String(body.message || '').trim();
  if (!name || !email || !subject || !message) return { success:false, error:'Please complete all contact fields.' };
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) return { success:false, error:'Please enter a valid email address.' };
  const sheet = ensureSimpleSheet_('Messages', CONTACT_HEADERS);
  sheet.appendRow([new Date().toISOString(), name, email, subject, message, 'New']);
  const admin = getAdminEmail_();
  if (admin) sendCustomerEmail_(admin, `ROUGE Client Message — ${subject}`, `<div style="font-family:Arial,sans-serif;line-height:1.7"><h2>${escapeEmailHtml_(subject)}</h2><p><strong>${escapeEmailHtml_(name)}</strong> — ${escapeEmailHtml_(email)}</p><p>${escapeEmailHtml_(message).replace(/\n/g,'<br>')}</p></div>`);
  return { success:true };
}

function subscribeNewsletter(body) {
  const email = String(body.email || '').trim().toLowerCase();
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) return { success:false, error:'Please enter a valid email address.' };
  const sheet = ensureSimpleSheet_('Newsletter', NEWSLETTER_HEADERS);
  const values = sheet.getDataRange().getValues();
  const emailIndex = NEWSLETTER_HEADERS.indexOf('Email');
  if (values.slice(1).some(r => String(r[emailIndex] || '').trim().toLowerCase() === email)) return { success:true, alreadySubscribed:true };
  sheet.appendRow([new Date().toISOString(), email, String(body.source || 'website').trim(), 'Subscribed']);
  return { success:true };
}

function setupClientServiceSheets() {
  ensureSimpleSheet_('Messages', CONTACT_HEADERS);
  ensureSimpleSheet_('Newsletter', NEWSLETTER_HEADERS);
  return 'Ready: Messages + Newsletter';
}

/**
 * Shared upsert for the Homepage Images and Collection Page Images sheets
 * (both use Key | Image | Alt | Title | Status columns).
 */
function saveKeyedImageRow_(sheetName, body) {
  if (!checkEditPassword_(body.password)) return { success: false, error: 'Invalid password' };
  const key = String(body.key || '').trim().toLowerCase();
  if (!key) return { success: false, error: 'Missing key' };

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) sheet = ss.insertSheet(sheetName);
  ensureTitleColumn(sheet);

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => String(h).trim());
  const keyIndex = headers.indexOf('Key');
  const imageIndex = headers.indexOf('Image');
  const altIndex = headers.indexOf('Alt');
  const titleIndex = headers.indexOf('Title');
  const statusIndex = headers.indexOf('Status');

  if (keyIndex === -1 || imageIndex === -1) {
    return { success: false, error: `Sheet "${sheetName}" is missing Key/Image columns.` };
  }

  const values = sheet.getDataRange().getValues();
  let targetRow = -1;
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][keyIndex] || '').trim().toLowerCase() === key) { targetRow = i + 1; break; }
  }

  const rowData = [];
  rowData[keyIndex] = key;
  rowData[imageIndex] = String(body.image || '').trim();
  if (altIndex >= 0) rowData[altIndex] = String(body.alt || '').trim();
  if (titleIndex >= 0) rowData[titleIndex] = String(body.title || '').trim();
  if (statusIndex >= 0) rowData[statusIndex] = String(body.status || 'Active').trim();

  if (targetRow === -1) {
    sheet.getRange(sheet.getLastRow() + 1, 1, 1, headers.length).setValues([rowData]);
  } else {
    sheet.getRange(targetRow, 1, 1, headers.length).setValues([rowData]);
  }
  return { success: true };
}

function getProducts(params) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) throw new Error(`Sheet "${CONFIG.SHEET_NAME}" was not found.`);

  // Keep the master Products sheet in the same structure as every page tab.
  ensureProductsTitleColumn(sheet);

  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return { success: true, count: 0, products: [] };

  const headers = values[0].map(h => String(h).trim());
  let products = values.slice(1)
    .map(row => rowToProduct(headers, row))
    .filter(product => product.id)
    .filter(product => String(product.status).trim().toLowerCase() === 'active');

  // The presence of a product row in a page tab controls whether it appears
  // on that website page. All product data itself is kept in Products.
  const pageAssignments = getPageAssignments();
  products = products.map(product => {
    const assignment = pageAssignments[String(product.id).trim()] || { pages: [], overrides: {} };
    return {
      ...product,
      pages: assignment.pages || [],
      pageOverrides: assignment.overrides || {}
    };
  });

  if (params.category) {
    const category = String(params.category).toLowerCase();
    products = products.filter(p => p.category.toLowerCase() === category);
  }

  if (String(params.featured).toLowerCase() === 'true') {
    products = products.filter(p => p.featured === true);
  }

  if (params.search) {
    const search = String(params.search).toLowerCase();
    products = products.filter(p => (`${p.name} ${p.title} ${p.category} ${p.description} ${p.color}`).toLowerCase().includes(search));
  }

  return { success: true, count: products.length, products };
}


function getPageAssignments() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const assignments = {};

  Object.keys(CONFIG.PAGE_SHEETS).forEach(sheetName => {
    const pageKey = CONFIG.PAGE_SHEETS[sheetName];
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet || sheet.getLastRow() < 2) return;

    const values = sheet.getDataRange().getValues();
    const headers = values[0].map(h => String(h).trim());
    const lowerHeaders = headers.map(h => h.toLowerCase());
    const idIndex = lowerHeaders.indexOf('id');
    const titleIndex = lowerHeaders.indexOf('title');
    const image1Index = lowerHeaders.indexOf('image1');
    if (idIndex === -1) return;

    values.slice(1).forEach((row, rowOffset) => {
      const id = String(row[idIndex] || '').trim();
      if (!id) return;

      if (!assignments[id]) assignments[id] = { pages: [], overrides: {} };
      if (!assignments[id].pages.includes(pageKey)) assignments[id].pages.push(pageKey);

      // Keep the page override structure for frontend compatibility. Because
      // page tabs now mirror Products, these values are normally identical
      // to the master product values.
      assignments[id].overrides[pageKey] = {
        title: titleIndex >= 0 ? String(row[titleIndex] || '').trim() : '',
        image: image1Index >= 0 ? normalizeHomepageImageUrl(String(row[image1Index] || '').trim()) : '',
        sortOrder: rowOffset + 1
      };
    });
  });

  return assignments;
}

/**
 * Exact master/page-tab headers.
 * Title was intentionally added immediately after ID.
 */
const PRODUCT_HEADERS = [
  'ID', 'Title', 'Name', 'Category', 'Price', 'Description', 'Color',
  'Sizes', 'Stock', 'Image1', 'Image2', 'Image3', 'Featured', 'Status'
];

function ensureProductsTitleColumn(sheet) {
  const lastCol = Math.max(sheet.getLastColumn(), 1);
  let headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h).trim());

  if (headers.indexOf('Title') === -1) {
    // Add Title immediately after ID without disturbing the existing fields.
    const idIndex = headers.indexOf('ID');
    if (idIndex >= 0) {
      sheet.insertColumnAfter(idIndex + 1);
    } else {
      sheet.insertColumnBefore(1);
    }
  }

  // Normalize the header row to the canonical product layout.
  sheet.getRange(1, 1, 1, PRODUCT_HEADERS.length).setValues([PRODUCT_HEADERS]);
  sheet.getRange(1, 1, 1, PRODUCT_HEADERS.length).setFontWeight('bold');
  sheet.setFrozenRows(1);

  // Existing products that do not have a title use their Name as the title.
  const values = sheet.getDataRange().getValues();
  const titleIndex = PRODUCT_HEADERS.indexOf('Title');
  const nameIndex = PRODUCT_HEADERS.indexOf('Name');
  const idIndex = PRODUCT_HEADERS.indexOf('ID');
  if (values.length > 1) {
    const titleValues = values.slice(1).map(row => {
      const title = String(row[titleIndex] || '').trim();
      const name = String(row[nameIndex] || '').trim();
      return [title || name];
    });
    sheet.getRange(2, titleIndex + 1, titleValues.length, 1).setValues(titleValues);
  }

  sheet.autoResizeColumns(1, PRODUCT_HEADERS.length);
  sheet.setColumnWidth(2, 220);
  sheet.setColumnWidth(3, 220);
  sheet.setColumnWidth(6, 360);
}

/**
 * Makes every page tab use EXACTLY the same headers/order as Products.
 * Page membership is controlled by which products exist in that tab.
 */
function setupProductPageSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const productsSheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (!productsSheet) throw new Error(`Sheet "${CONFIG.SHEET_NAME}" was not found.`);

  ensureProductsTitleColumn(productsSheet);

  const productsValues = productsSheet.getDataRange().getValues();
  const headers = productsValues[0].map(h => String(h).trim());
  const idIndex = headers.indexOf('ID');
  const categoryIndex = headers.indexOf('Category');
  const featuredIndex = headers.indexOf('Featured');
  if (idIndex === -1) throw new Error('Products sheet must contain an ID column.');

  const rows = productsValues.slice(1).filter(r => String(r[idIndex] || '').trim());
  const pages = {
    'New In': rows.filter(r => featuredIndex >= 0 && String(r[featuredIndex]).trim().toLowerCase() === 'true'),
    'Collections': rows.slice(),
    'Dresses': rows.filter(r => ['dresses','evening'].includes(String(categoryIndex >= 0 ? r[categoryIndex] : '').trim().toLowerCase())),
    'Clothing': rows.filter(r => String(categoryIndex >= 0 ? r[categoryIndex] : '').trim().toLowerCase() === 'clothing'),
    'Accessories': rows.filter(r => String(categoryIndex >= 0 ? r[categoryIndex] : '').trim().toLowerCase() === 'accessories')
  };

  Object.keys(CONFIG.PAGE_SHEETS).forEach(sheetName => {
    let sheet = ss.getSheetByName(sheetName);
    if (!sheet) sheet = ss.insertSheet(sheetName);

    ensurePageSheetLayout(sheet);

    // Only seed a page tab when it is empty. Never overwrite products the user
    // has already placed there.
    if (sheet.getLastRow() < 2 && pages[sheetName].length) {
      sheet.getRange(2, 1, pages[sheetName].length, PRODUCT_HEADERS.length).setValues(pages[sheetName]);
    }

    formatProductSheet(sheet);
  });

  return 'Ready: Products + all page tabs now use the same 14-column product layout.';
}

function ensurePageSheetLayout(sheet) {
  // If the sheet contains the old ID | Title | Image | Active | Sort Order
  // layout, rebuild its rows into the new full-product layout when possible.
  const lastRow = sheet.getLastRow();
  const lastCol = Math.max(sheet.getLastColumn(), 1);
  const oldHeaders = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h).trim());

  if (oldHeaders.join('|') === PRODUCT_HEADERS.join('|')) return;

  const idIndex = oldHeaders.findIndex(h => h.toLowerCase() === 'id');
  const titleIndex = oldHeaders.findIndex(h => h.toLowerCase() === 'title');
  const imageIndex = oldHeaders.findIndex(h => h.toLowerCase() === 'image');
  const activeIndex = oldHeaders.findIndex(h => h.toLowerCase() === 'active');

  const oldRows = lastRow > 1 ? sheet.getRange(2, 1, lastRow - 1, lastCol).getValues() : [];
  const productsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_NAME);
  const masterValues = productsSheet.getDataRange().getValues();
  const masterHeaders = masterValues[0].map(h => String(h).trim());
  const masterIdIndex = masterHeaders.indexOf('ID');
  const masterMap = {};
  masterValues.slice(1).forEach(row => {
    const id = String(row[masterIdIndex] || '').trim();
    if (id) masterMap[id] = row.slice(0, PRODUCT_HEADERS.length);
  });

  const converted = [];
  oldRows.forEach(row => {
    const id = idIndex >= 0 ? String(row[idIndex] || '').trim() : '';
    if (!id) return;
    const base = masterMap[id] ? masterMap[id].slice() : new Array(PRODUCT_HEADERS.length).fill('');
    const title = titleIndex >= 0 ? String(row[titleIndex] || '').trim() : '';
    const image = imageIndex >= 0 ? String(row[imageIndex] || '').trim() : '';
    if (title) base[1] = title;
    if (image) base[9] = image;
    if (activeIndex >= 0 && String(row[activeIndex]).trim().toLowerCase() === 'false') return;
    converted.push(base);
  });

  sheet.clear();
  sheet.getRange(1, 1, 1, PRODUCT_HEADERS.length).setValues([PRODUCT_HEADERS]);
  if (converted.length) sheet.getRange(2, 1, converted.length, PRODUCT_HEADERS.length).setValues(converted);
}

function formatProductSheet(sheet) {
  sheet.getRange(1, 1, 1, PRODUCT_HEADERS.length).setValues([PRODUCT_HEADERS]);
  sheet.getRange(1, 1, 1, PRODUCT_HEADERS.length).setFontWeight('bold');
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, PRODUCT_HEADERS.length);
  sheet.setColumnWidth(2, 220);
  sheet.setColumnWidth(3, 220);
  sheet.setColumnWidth(6, 360);
}

/**
 * Automatically copies product rows entered/edited in any page tab into the
 * master Products tab. Existing IDs are updated; new IDs are appended.
 *
 * This is a spreadsheet onEdit trigger, so normal manual edits in Google
 * Sheets sync immediately without needing a web-app request.
 */
function onEdit(e) {
  if (!e || !e.range) return;
  const sheet = e.range.getSheet();
  const sheetName = sheet.getName();
  if (!Object.prototype.hasOwnProperty.call(CONFIG.PAGE_SHEETS, sheetName)) return;
  if (e.range.getRow() < 2) return;

  // Support both single-cell edits and multi-row paste operations.
  const startRow = e.range.getRow();
  const rowCount = e.range.getNumRows();
  const rows = sheet.getRange(startRow, 1, rowCount, PRODUCT_HEADERS.length).getValues();

  rows.forEach(row => {
    const id = String(row[0] || '').trim();
    if (!id) return;
    syncPageRowToProducts_(row);
  });
}

function syncPageRowToProducts_(row) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const productsSheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (!productsSheet) return;

  ensureProductsTitleColumn(productsSheet);
  const values = productsSheet.getDataRange().getValues();
  const idIndex = 0;
  let targetRow = -1;

  for (let i = 1; i < values.length; i++) {
    if (String(values[i][idIndex] || '').trim() === String(row[0]).trim()) {
      targetRow = i + 1;
      break;
    }
  }

  if (targetRow === -1) {
    productsSheet.getRange(productsSheet.getLastRow() + 1, 1, 1, PRODUCT_HEADERS.length).setValues([row]);
  } else {
    productsSheet.getRange(targetRow, 1, 1, PRODUCT_HEADERS.length).setValues([row]);
  }

  formatProductSheet(productsSheet);
}

/**
 * Manual recovery/sync helper. Use this if you paste/import many rows at once
 * or want to force all page tabs into Products.
 */
function syncAllPageSheetsToProducts() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const productsSheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (!productsSheet) throw new Error(`Sheet "${CONFIG.SHEET_NAME}" was not found.`);
  ensureProductsTitleColumn(productsSheet);

  const masterValues = productsSheet.getDataRange().getValues();
  const byId = {};
  masterValues.slice(1).forEach((row, i) => {
    const id = String(row[0] || '').trim();
    if (id) byId[id] = i + 2;
  });

  Object.keys(CONFIG.PAGE_SHEETS).forEach(sheetName => {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet || sheet.getLastRow() < 2) return;
    ensurePageSheetLayout(sheet);
    const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, PRODUCT_HEADERS.length).getValues();
    rows.forEach(row => {
      const id = String(row[0] || '').trim();
      if (!id) return;
      if (byId[id]) {
        productsSheet.getRange(byId[id], 1, 1, PRODUCT_HEADERS.length).setValues([row]);
      } else {
        productsSheet.getRange(productsSheet.getLastRow() + 1, 1, 1, PRODUCT_HEADERS.length).setValues([row]);
        byId[id] = productsSheet.getLastRow();
      }
    });
  });

  formatProductSheet(productsSheet);
  return 'All page-tab products have been synchronized into Products.';
}

function getProduct(id) {
  if (!id) return { success: false, error: 'Product ID is required.' };
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) throw new Error(`Sheet "${CONFIG.SHEET_NAME}" was not found.`);

  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return { success: false, error: 'No products found.' };
  const headers = values[0].map(h => String(h).trim());

  for (let i = 1; i < values.length; i++) {
    const product = rowToProduct(headers, values[i]);
    if (String(product.id).toLowerCase() === String(id).toLowerCase()) {
      if (String(product.status).trim().toLowerCase() !== 'active') {
        return { success: false, error: 'Product is not active.' };
      }
      return { success: true, product };
    }
  }
  return { success: false, error: 'Product not found.' };
}


/**
 * Returns the three independent homepage collection-card images.
 * Sheet name: Homepage Images
 * Columns: Key | Image | Alt | Status
 */
function getHomepageImages() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CONFIG.HOMEPAGE_IMAGES_SHEET_NAME);

  // If the sheet does not exist yet, return a clear setup error instead of
  // breaking the whole product API.
  if (!sheet) {
    return {
      success: false,
      error: `Sheet "${CONFIG.HOMEPAGE_IMAGES_SHEET_NAME}" was not found. Run setupHomepageImagesSheet() once.`
    };
  }

  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return { success: true, images: {} };

  const headers = values[0].map(h => String(h).trim());
  const keyIndex = headers.indexOf('Key');
  const imageIndex = headers.indexOf('Image');
  const altIndex = headers.indexOf('Alt');
  const titleIndex = headers.indexOf('Title');
  const statusIndex = headers.indexOf('Status');

  if (keyIndex === -1 || imageIndex === -1) {
    throw new Error('Homepage Images sheet must contain Key and Image columns.');
  }

  const images = {};
  values.slice(1).forEach(row => {
    const key = String(row[keyIndex] || '').trim().toLowerCase();
    const rawImage = String(row[imageIndex] || '').trim();
    const alt = altIndex >= 0 ? String(row[altIndex] || '').trim() : '';
    const status = statusIndex >= 0 ? String(row[statusIndex] || 'Active').trim().toLowerCase() : 'active';
    if (key && rawImage && status !== 'inactive') {
      const title = titleIndex >= 0 ? String(row[titleIndex] || '').trim() : '';
      images[key] = { image: normalizeHomepageImageUrl(rawImage), alt, title };
    }
  });

  return { success: true, images };
}

/**
 * Accepts either a normal image URL, a Google Drive sharing URL, or a
 * Google Drive file URL. Drive sharing links are converted to a thumbnail
 * URL that can be used directly by an <img> element.
 */
function normalizeHomepageImageUrl(url) {
  const s = String(url || '').trim();
  if (!s) return '';

  // Already a direct image/data URL or local asset path.
  if (!/^https?:\/\//i.test(s)) return s;
  if (/drive\.google\.com/.test(s)) {
    let id = '';
    let m = s.match(/drive\.google\.com\/file\/d\/([a-zA-Z0-9_-]+)/);
    if (m) id = m[1];
    if (!id) {
      m = s.match(/[?&]id=([a-zA-Z0-9_-]+)/);
      if (m) id = m[1];
    }
    if (id) return 'https://drive.google.com/thumbnail?id=' + encodeURIComponent(id) + '&sz=w1600';
  }
  return s;
}

/**
 * Run this function once from the Apps Script editor.
 * It creates the sheet and the three rows used by the ROUGE homepage.
 */
function setupHomepageImagesSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CONFIG.HOMEPAGE_IMAGES_SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(CONFIG.HOMEPAGE_IMAGES_SHEET_NAME);

  ensureTitleColumn(sheet);
  const existing = sheet.getLastRow();
  if (existing < 2) {
    sheet.getRange(2, 1, 3, 5).setValues([
      ['dresses', 'assets/hero.png', 'Dresses editorial', 'THE WHITE\nSILHOUETTE', 'Active'],
      ['clothing', 'assets/hero.png', 'Clothing editorial', 'AFTER\nDARK', 'Active'],
      ['accessories', 'assets/hero.png', 'Accessories editorial', 'SIGNATURE\nRED', 'Active']
    ]);
  }
  sheet.getRange(1, 1, 1, 4).setFontWeight('bold');
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, 4);
  return `Ready: ${CONFIG.HOMEPAGE_IMAGES_SHEET_NAME}`;
}



/**
 * Returns the three independent collection-page card images.
 * Sheet name: Collection Page Images
 * Columns: Key | Image | Alt | Status
 */
function getCollectionImages() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.COLLECTIONS_IMAGES_SHEET_NAME);

  if (!sheet) {
    return {
      success: false,
      error: `Sheet "${CONFIG.COLLECTIONS_IMAGES_SHEET_NAME}" was not found. Run setupCollectionPageImagesSheet() once.`
    };
  }

  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return { success: true, images: {} };

  const headers = values[0].map(h => String(h).trim());
  const keyIndex = headers.indexOf('Key');
  const imageIndex = headers.indexOf('Image');
  const altIndex = headers.indexOf('Alt');
  const titleIndex = headers.indexOf('Title');
  const statusIndex = headers.indexOf('Status');

  if (keyIndex === -1 || imageIndex === -1) {
    throw new Error('Collection Page Images sheet must contain Key and Image columns.');
  }

  const images = {};
  values.slice(1).forEach(row => {
    const key = String(row[keyIndex] || '').trim().toLowerCase();
    const rawImage = String(row[imageIndex] || '').trim();
    const alt = altIndex >= 0 ? String(row[altIndex] || '').trim() : '';
    const status = statusIndex >= 0 ? String(row[statusIndex] || 'Active').trim().toLowerCase() : 'active';
    if (key && rawImage && status !== 'inactive') {
      const title = titleIndex >= 0 ? String(row[titleIndex] || '').trim() : '';
      images[key] = { image: normalizeHomepageImageUrl(rawImage), alt, title };
    }
  });

  return { success: true, images };
}

/**
 * Run once from the Apps Script editor.
 * Creates the separate sheet used only by the Collections subpage.
 */
function setupCollectionPageImagesSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CONFIG.COLLECTIONS_IMAGES_SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(CONFIG.COLLECTIONS_IMAGES_SHEET_NAME);

  ensureTitleColumn(sheet);
  if (sheet.getLastRow() < 2) {
    sheet.getRange(2, 1, 3, 5).setValues([
      ['dresses', 'assets/hero.png', 'Dresses collection', 'The White\nSilhouette', 'Active'],
      ['clothing', 'assets/hero.png', 'Clothing collection', 'Modern\nStructure', 'Active'],
      ['accessories', 'assets/hero.png', 'Accessories collection', 'Signature\nDetails', 'Active']
    ]);
  }
  sheet.getRange(1, 1, 1, 4).setFontWeight('bold');
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, 4);
  return `Ready: ${CONFIG.COLLECTIONS_IMAGES_SHEET_NAME}`;
}

function ensureTitleColumn(sheet) {
  const lastCol = Math.max(sheet.getLastColumn(), 1);
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h).trim());
  const hasAnyHeader = headers.some(Boolean);

  if (!hasAnyHeader) {
    // Brand new / never-initialized sheet (e.g. saved to from admin-products.html
    // before the one-time setup*() function was ever run). Write the full
    // canonical header row from scratch instead of trying to "insert a Title
    // column" next to headers that don't exist yet.
    const fullHeaders = ['Key', 'Image', 'Alt', 'Title', 'Status'];
    sheet.getRange(1, 1, 1, fullHeaders.length).setValues([fullHeaders]);
    sheet.getRange(1, 1, 1, fullHeaders.length).setFontWeight('bold');
    return;
  }

  if (headers.indexOf('Title') === -1) {
    // Existing sheets originally had Key | Image | Alt | Status.
    // Insert Title before Status so existing Status values stay intact.
    const statusIndex = headers.indexOf('Status');
    if (statusIndex >= 0) {
      sheet.insertColumnBefore(statusIndex + 1);
    } else {
      sheet.insertColumnAfter(lastCol);
    }
  }
  const cols = sheet.getLastColumn();
  const finalHeaders = sheet.getRange(1, 1, 1, cols).getValues()[0].map(h => String(h).trim());
  const titleIndex = finalHeaders.indexOf('Title');
  if (titleIndex >= 0) sheet.getRange(1, titleIndex + 1).setValue('Title');
  sheet.getRange(1, 1, 1, sheet.getLastColumn()).setFontWeight('bold');
}

function getCategories() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) throw new Error(`Sheet "${CONFIG.SHEET_NAME}" was not found.`);
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return { success: true, categories: [] };
  const headers = values[0].map(h => String(h).trim());
  const categoryIndex = headers.indexOf('Category');
  if (categoryIndex === -1) throw new Error('Category column was not found.');
  const categories = [...new Set(values.slice(1).map(r => String(r[categoryIndex]).trim()).filter(Boolean))];
  return { success: true, categories };
}

function rowToProduct(headers, row) {
  const data = {};
  headers.forEach((header, index) => data[header] = row[index]);
  return {
    id: String(data.ID || '').trim(),
    title: String(data.Title || data.Name || '').trim(),
    name: String(data.Name || data.Title || '').trim(),
    category: String(data.Category || '').trim(),
    price: Number(data.Price || 0),
    description: String(data.Description || '').trim(),
    color: String(data.Color || '').trim(),
    sizes: String(data.Sizes || '').split(',').map(s => s.trim()).filter(Boolean),
    stock: Number(data.Stock || 0),
    images: [data.Image1, data.Image2, data.Image3].map(v => String(v || '').trim()).filter(Boolean),
    featured: String(data.Featured || '').trim().toLowerCase() === 'true',
    status: String(data.Status || '').trim()
  };
}

function jsonResponse(data, callback) {
  const json = JSON.stringify(data);
  if (callback) {
    return ContentService
      .createTextOutput(`${callback}(${json});`)
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService
    .createTextOutput(json)
    .setMimeType(ContentService.MimeType.JSON);
}
