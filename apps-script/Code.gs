// ── OneSignal config (fill this in) ───────────────────────────
var ONESIGNAL_APP_ID       = 'db6beb30-8bc6-44fc-8dee-833e0ace7746';
var ONESIGNAL_REST_API_KEY = 'YOUR_ONESIGNAL_REST_API_KEY_HERE'; // OneSignal dashboard → Settings → Keys & IDs

// ── Pexels config (free stock-photo search, fill this in) ─────
var PEXELS_API_KEY = 'YOUR_PEXELS_API_KEY_HERE'; // pexels.com/api

// Searches Pexels for a representative photo of an item name, downloads
// it, and saves it to Drive so the app has a stable URL to use — done
// server-side (not in the browser) purely to avoid CORS, since the
// browser can't read cross-origin image bytes from pexels.com directly.
function fetchItemImage_(query) {
  try {
    if (!PEXELS_API_KEY || PEXELS_API_KEY === 'YOUR_PEXELS_API_KEY_HERE') {
      return { success: false, error: 'PEXELS_API_KEY not set in Apps Script' };
    }
    var searchRes = UrlFetchApp.fetch(
      'https://api.pexels.com/v1/search?query=' + encodeURIComponent(query) + '&per_page=1',
      { headers: { Authorization: PEXELS_API_KEY }, muteHttpExceptions: true }
    );
    var searchData = JSON.parse(searchRes.getContentText());
    var photo = searchData.photos && searchData.photos[0];
    if (!photo) return { success: false, error: 'No matching stock photo found' };

    var imgRes = UrlFetchApp.fetch(photo.src.medium, { muteHttpExceptions: true });
    var blob = imgRes.getBlob();

    var folders = DriveApp.getFoldersByName('Procurement Item Images');
    var folder = folders.hasNext() ? folders.next() : DriveApp.createFolder('Procurement Item Images');
    var file = folder.createFile(blob).setName(query + '.jpg');
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    // drive.google.com/uc?export=view is unreliable for direct <img> embedding
    // (Google's anti-hotlinking measures often serve a scan-warning page
    // instead of the raw image). The dedicated thumbnail endpoint is built
    // for exactly this and renders reliably inline.
    return { success: true, imageUrl: 'https://drive.google.com/thumbnail?sz=w400&id=' + file.getId() };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

// ── Notification helpers ──────────────────────────────────────
function sendOneSignal_(title, msg) {
  if (!ONESIGNAL_REST_API_KEY || ONESIGNAL_REST_API_KEY === 'YOUR_ONESIGNAL_REST_API_KEY_HERE') return;
  try {
    UrlFetchApp.fetch(
      'https://onesignal.com/api/v1/notifications',
      { method: 'post', contentType: 'application/json',
        headers: { Authorization: 'Basic ' + ONESIGNAL_REST_API_KEY },
        payload: JSON.stringify({
          app_id: ONESIGNAL_APP_ID,
          included_segments: ['Subscribed Users'],
          headings: { en: title },
          contents: { en: msg }
        }),
        muteHttpExceptions: true }
    );
  } catch(e) { Logger.log('OneSignal error: ' + e); }
}

function pushNotification_(type, msg, details) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('Notifications');
  if (!sh) {
    sh = ss.insertSheet('Notifications');
    sh.appendRow(['ID','Type','Message','Details','Created At']);
  }
  var n = Math.max(sh.getLastRow(), 1);
  var id = 'NOTIF-' + ('000' + n).slice(-4);
  sh.appendRow([id, type, msg, JSON.stringify(details || {}), new Date().toISOString()]);
  sendOneSignal_('Site Procurement', msg);
}

function handleNotify_(type, d) {
  var icons = { sr_submitted:'&#x1F4CB;', sr_approved:'&#x2705;',
                po_created:'&#x1F6D2;',  mre_submitted:'&#x1F4E6;' };
  var icon = icons[type] || '&#x1F514;';
  // Message is pre-built in the app; GAS just prepends the icon
  var msg = d.message ? icon + ' ' + d.message : icon + ' ' + (d.id || type);
  pushNotification_(type, msg, d);
}

// ── doPost ────────────────────────────────────────────────────
function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    // NOTIFY action — write to Notifications sheet + send OneSignal push
    if (data.action === 'notify') {
      handleNotify_(data.type, data.data || {});
      return jsonOut_({success:true});
    }

    // FETCH ITEM IMAGE — search Pexels, save to Drive, return the URL
    if (data.action === 'fetchItemImage') {
      return jsonOut_(fetchItemImage_(data.query));
    }

    // Every write below touches a data sheet, so serialize them: without a
    // lock, two users saving at the same moment can both read the same
    // "no existing row" snapshot and end up appending duplicate rows, or
    // both look up the same row index and clobber each other's update.
    var lock = LockService.getScriptLock();
    if (!lock.tryLock(10000)) {
      return jsonOut_({success:false, error:'Sheet is busy — please retry.'});
    }
    try {
      var sheet = ss.getSheetByName(data.sheet)
                 || ss.insertSheet(data.sheet);

      // DELETE row by ID
      if (data.action === 'delete') {
        var delRow = findRowById_(sheet, data.id);
        if (delRow > -1) sheet.deleteRow(delRow);
        return jsonOut_({success:true});
      }

      // --- Ensure header row is present; add any missing columns ---
      var sheetHeaders;
      if (sheet.getLastRow() === 0) {
        sheet.appendRow(data.headers);
        sheetHeaders = data.headers.slice();
      } else {
        sheetHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]
                           .map(function(h){ return String(h); });
        // Append any column the app now sends that the sheet doesn't have yet
        data.headers.forEach(function(h) {
          if (sheetHeaders.indexOf(h) === -1) {
            sheetHeaders.push(h);
            sheet.getRange(1, sheetHeaders.length).setValue(h);
          }
        });
      }

      // Map values to the sheet's actual column order (handles reordering & new cols)
      var mappedRow = sheetHeaders.map(function(h) {
        var idx = data.headers.indexOf(h);
        return idx >= 0 ? data.row[idx] : '';
      });

      // UPDATE existing row if ID matches, else append
      var rowIdx = findRowById_(sheet, mappedRow[0]);
      if (rowIdx > -1) {
        sheet.getRange(rowIdx, 1, 1, mappedRow.length).setValues([mappedRow]);
      } else {
        sheet.appendRow(mappedRow);
      }

      return jsonOut_({success:true});
    } finally {
      lock.releaseLock();
    }
  } catch (err) {
    return jsonOut_({success:false, error: err.message});
  }
}

// Finds the 1-indexed row whose column-A value matches id. Reads only
// column A — not getDataRange(), which pulls every column (including the
// large JSON blobs in the others) of every row — so lookups stay cheap as
// a sheet grows into the thousands of rows.
function findRowById_(sheet, id) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return -1;
  var ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  for (var i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === String(id)) return i + 2;
  }
  return -1;
}

function jsonOut_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
                        .setMimeType(ContentService.MimeType.JSON);
}

// ── doGet ─────────────────────────────────────────────────────
function doGet(e) {
  try {
    var action = e.parameter.action;
    if (action === 'readAll') {
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var sheets = ['Requisitions','PurchaseOrders',
                    'SiteReceipts','Works','MasterData','Notifications'];
      var result = {};
      sheets.forEach(function(name) {
        var sheet = ss.getSheetByName(name);
        if (!sheet || sheet.getLastRow() < 2) { result[name] = []; return; }
        var rows = sheet.getDataRange().getValues();
        var headers = rows[0];
        result[name] = rows.slice(1).map(function(row) {
          var obj = {};
          headers.forEach(function(h, i) { obj[h] = row[i]; });
          return obj;
        });
      });
      return jsonOut_(result);
    }
    return jsonOut_({status:'ok'});
  } catch (err) {
    return jsonOut_({status:'error', error: err.message});
  }
}
