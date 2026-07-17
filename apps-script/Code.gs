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

// ══════════════════════════════════════════════════════════════
//  NORMALIZED SHEET STORAGE
//
//  The client still sends/receives records the same way it always has
//  (a header row plus 'Items (JSON)' / 'Labor Items (JSON)' blobs) — that
//  wire format doesn't change, so none of index.html needed to change.
//  What changes is what THIS script does with it: instead of stuffing
//  those JSON blobs into a single cell, each line item becomes its own
//  readable row in a dedicated child tab (e.g. RequisitionItems), and
//  admin master data (team, vendors, materials, labor, projects) each
//  get their own real tab instead of one opaque MasterData blob.
// ══════════════════════════════════════════════════════════════

// Maps a header-tab sheet name to how its two JSON-blob columns split
// into rows of a child "line item" tab, and back.
var ITEM_TABS = {
  Requisitions: {
    sheet: 'RequisitionItems',
    itemsHeaderName: 'Items (JSON)',
    laborHeaderName: 'Labor Items (JSON)',
    headers: ['Requisition ID','Type','Name','Qty','Unit','Work'],
    toRows: function(id, items, laborItems) {
      var rows = [];
      (items || []).forEach(function(i) { rows.push([id, 'Material', i.material, i.qty, i.unit, i.work || '']); });
      (laborItems || []).forEach(function(i) { rows.push([id, 'Labor', i.trade, i.qty, i.unit, i.work || '']); });
      return rows;
    },
    fromRows: function(rows) {
      var items = [], laborItems = [];
      rows.forEach(function(r) {
        if (r['Type'] === 'Material') items.push({ material: r['Name'], qty: r['Qty'], unit: r['Unit'], work: r['Work'] || '' });
        else laborItems.push({ trade: r['Name'], qty: r['Qty'], unit: r['Unit'], work: r['Work'] || '' });
      });
      return { items: items, laborItems: laborItems };
    }
  },
  PurchaseOrders: {
    sheet: 'POItems',
    itemsHeaderName: 'Items (JSON)',
    laborHeaderName: 'Labor Items (JSON)',
    headers: ['PO ID','Type','Name','Qty','Unit','Rate','Amount','Work'],
    toRows: function(id, items, laborItems) {
      var rows = [];
      (items || []).forEach(function(i) { rows.push([id, 'Material', i.material, i.qty, i.unit, i.rate || '', i.amount, i.work || '']); });
      (laborItems || []).forEach(function(i) { rows.push([id, 'Labor', i.trade, i.qty, i.unit, i.rate || '', i.amount, i.work || '']); });
      return rows;
    },
    fromRows: function(rows) {
      var items = [], laborItems = [];
      rows.forEach(function(r) {
        var base = { qty: r['Qty'], unit: r['Unit'], rate: r['Rate'] || '', amount: r['Amount'] || 0, work: r['Work'] || '' };
        if (r['Type'] === 'Material') { base.material = r['Name']; items.push(base); }
        else { base.trade = r['Name']; laborItems.push(base); }
      });
      return { items: items, laborItems: laborItems };
    }
  },
  SiteReceipts: {
    sheet: 'ReceiptItems',
    itemsHeaderName: 'Items (JSON)',
    laborHeaderName: 'Labor Items (JSON)',
    headers: ['Receipt ID','Type','Name','Ordered Qty','Received Qty','Unit','Condition'],
    toRows: function(id, items, laborItems) {
      var rows = [];
      (items || []).forEach(function(i) { rows.push([id, 'Material', i.material, i.orderedQty || '', i.receivedQty || '', i.unit, i.condition || '']); });
      (laborItems || []).forEach(function(i) { rows.push([id, 'Labor', i.trade, '', i.qty || '', i.unit, '']); });
      return rows;
    },
    fromRows: function(rows) {
      var items = [], laborItems = [];
      rows.forEach(function(r) {
        if (r['Type'] === 'Material') {
          items.push({ material: r['Name'], orderedQty: r['Ordered Qty'] || '', receivedQty: r['Received Qty'] || '', unit: r['Unit'], condition: r['Condition'] || '' });
        } else {
          laborItems.push({ trade: r['Name'], qty: r['Received Qty'] || '', unit: r['Unit'] });
        }
      });
      return { items: items, laborItems: laborItems };
    }
  },
  Works: {
    sheet: 'WorkEstimateItems',
    itemsHeaderName: 'Estimated Materials (JSON)',
    laborHeaderName: 'Estimated Labor (JSON)',
    headers: ['Work ID','Type','Name','Qty','Unit'],
    toRows: function(id, items, laborItems) {
      var rows = [];
      (items || []).forEach(function(i) { rows.push([id, 'Material', i.material, i.qty, i.unit]); });
      (laborItems || []).forEach(function(i) { rows.push([id, 'Labor', i.trade, i.qty, i.unit]); });
      return rows;
    },
    fromRows: function(rows) {
      var items = [], laborItems = [];
      rows.forEach(function(r) {
        if (r['Type'] === 'Material') items.push({ material: r['Name'], qty: r['Qty'], unit: r['Unit'] });
        else laborItems.push({ trade: r['Name'], qty: r['Qty'], unit: r['Unit'] });
      });
      return { items: items, laborItems: laborItems };
    }
  }
};

// Writes the child-tab rows for one parent document, replacing whatever
// was there before (the client always sends the document's full current
// item list, so "clear this parent's rows, then append the new set" is
// simpler and just as correct as a diff).
function writeChildRows_(ss, cfg, parentId, items, laborItems) {
  deleteChildRows_(ss, cfg, parentId);
  var rows = cfg.toRows(parentId, items, laborItems);
  if (!rows.length) return;
  var sheet = ss.getSheetByName(cfg.sheet) || ss.insertSheet(cfg.sheet);
  if (sheet.getLastRow() === 0) sheet.appendRow(cfg.headers);
  sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, cfg.headers.length).setValues(rows);
}

// Deletes every child-tab row belonging to parentId. Reads only column A
// (the parent-ID column), not the whole row, so this stays cheap even
// once a child tab has many thousands of line items across all documents.
function deleteChildRows_(ss, cfg, parentId) {
  var sheet = ss.getSheetByName(cfg.sheet);
  if (!sheet || sheet.getLastRow() < 2) return;
  var ids = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
  for (var i = ids.length - 1; i >= 0; i--) {
    if (String(ids[i][0]) === String(parentId)) sheet.deleteRow(i + 2);
  }
}

// Full-replace a small admin catalog tab (Projects, TeamMembers, etc.)
// with the rows given. These are edited only from the Admin Panel, low
// frequency and low volume, so a clear-and-rewrite is simplest and safe.
function writeCatalog_(ss, name, headers, rows) {
  var sheet = ss.getSheetByName(name) || ss.insertSheet(name);
  sheet.clearContents();
  sheet.appendRow(headers);
  if (rows.length) sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
}

// Reads a sheet into an array of {header: value} objects, keyed by its
// header row — the same shape doGet has always returned per sheet.
function readSheetAsObjects_(ss, name) {
  var sheet = ss.getSheetByName(name);
  if (!sheet || sheet.getLastRow() < 2) return [];
  var rows = sheet.getDataRange().getValues();
  var headers = rows[0];
  return rows.slice(1).map(function(row) {
    var obj = {};
    headers.forEach(function(h, i) { obj[h] = row[i]; });
    return obj;
  });
}

// Splits the incoming masterData JSON blob across its real catalog tabs.
function writeMasterData_(ss, md) {
  writeCatalog_(ss, 'Projects', ['ID','Name','Lat','Lng'],
    (md.sites || []).map(function(s) { return [s.id, s.name, s.lat || '', s.lng || '']; }));
  writeCatalog_(ss, 'TeamMembers', ['ID','Name','Role','Approver','PIN Hash'],
    (md.teamMembers || []).map(function(m) { return [m.id, m.name, m.role || '', !!m.approver, m.pinHash || '']; }));
  writeCatalog_(ss, 'Categories', ['ID','Name'],
    (md.categories || []).map(function(c) { return [c.id, c.name]; }));
  writeCatalog_(ss, 'Vendors', ['ID','Name','Categories'],
    (md.vendors || []).map(function(v) { return [v.id, v.name, (v.categories || []).join(', ')]; }));
  writeCatalog_(ss, 'Materials', ['ID','Name','Unit','Category','Image URL'],
    (md.materials || []).map(function(m) { return [m.id, m.name, m.unit || '', m.category || '', m.imageUrl || '']; }));
  writeCatalog_(ss, 'LaborTypes', ['ID','Name','Unit','Category','Image URL'],
    (md.laborTypes || []).map(function(l) { return [l.id, l.name, l.unit || '', l.category || '', l.imageUrl || '']; }));
}

// Rebuilds the masterData object by joining the catalog tabs — the exact
// same shape APP.masterData has always expected.
function buildMasterData_(ss) {
  var projects = readSheetAsObjects_(ss, 'Projects').map(function(r) {
    var o = { id: r['ID'], name: r['Name'] };
    if (r['Lat']) o.lat = r['Lat'];
    if (r['Lng']) o.lng = r['Lng'];
    return o;
  });
  var teamMembers = readSheetAsObjects_(ss, 'TeamMembers').map(function(r) {
    return {
      id: r['ID'], name: r['Name'], role: r['Role'] || '',
      approver: r['Approver'] === true || r['Approver'] === 'TRUE' || r['Approver'] === '1',
      pinHash: r['PIN Hash'] || ''
    };
  });
  var categories = readSheetAsObjects_(ss, 'Categories').map(function(r) {
    return { id: r['ID'], name: r['Name'] };
  });
  var vendors = readSheetAsObjects_(ss, 'Vendors').map(function(r) {
    return {
      id: r['ID'], name: r['Name'],
      categories: r['Categories'] ? String(r['Categories']).split(',').map(function(s) { return s.trim(); }).filter(Boolean) : []
    };
  });
  var materials = readSheetAsObjects_(ss, 'Materials').map(function(r) {
    var o = { id: r['ID'], name: r['Name'], unit: r['Unit'] || '', category: r['Category'] || '' };
    if (r['Image URL']) o.imageUrl = r['Image URL'];
    return o;
  });
  var laborTypes = readSheetAsObjects_(ss, 'LaborTypes').map(function(r) {
    var o = { id: r['ID'], name: r['Name'], unit: r['Unit'] || '', category: r['Category'] || '' };
    if (r['Image URL']) o.imageUrl = r['Image URL'];
    return o;
  });
  return { teamMembers: teamMembers, categories: categories, vendors: vendors, materials: materials, sites: projects, works: [], laborTypes: laborTypes };
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

    // Every write below touches one or more data sheets, so serialize
    // them: without a lock, two users saving at the same moment can both
    // read the same "no existing row" snapshot and end up appending
    // duplicate rows, or both look up the same row index and clobber
    // each other's update.
    var lock = LockService.getScriptLock();
    if (!lock.tryLock(10000)) {
      return jsonOut_({success:false, error:'Sheet is busy — please retry.'});
    }
    try {
      // MASTER DATA — decompose the blob into its real catalog tabs
      if (data.sheet === 'MasterData') {
        writeMasterData_(ss, JSON.parse(data.row[1]));
        return jsonOut_({success:true});
      }

      var sheet = ss.getSheetByName(data.sheet)
                 || ss.insertSheet(data.sheet);
      var itemCfg = ITEM_TABS[data.sheet];

      // DELETE row by ID (and any child line-item rows it owns)
      if (data.action === 'delete') {
        var delRow = findRowById_(sheet, data.id);
        if (delRow > -1) sheet.deleteRow(delRow);
        if (itemCfg) deleteChildRows_(ss, itemCfg, data.id);
        return jsonOut_({success:true});
      }

      // Pull the two item-JSON columns out before they ever touch the
      // header tab — they get their own rows in a child tab instead.
      var headers = data.headers;
      var rowVals = data.row;
      var itemsJson = null, laborJson = null;
      if (itemCfg) {
        var iIdx = headers.indexOf(itemCfg.itemsHeaderName);
        var lIdx = headers.indexOf(itemCfg.laborHeaderName);
        if (iIdx > -1) itemsJson = rowVals[iIdx];
        if (lIdx > -1) laborJson = rowVals[lIdx];
        var strippedHeaders = [], strippedRow = [];
        headers.forEach(function(h, idx) {
          if (idx === iIdx || idx === lIdx) return;
          strippedHeaders.push(h);
          strippedRow.push(rowVals[idx]);
        });
        headers = strippedHeaders;
        rowVals = strippedRow;
      }

      // --- Ensure header row is present; add any missing columns ---
      var sheetHeaders;
      if (sheet.getLastRow() === 0) {
        sheet.appendRow(headers);
        sheetHeaders = headers.slice();
      } else {
        sheetHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]
                           .map(function(h){ return String(h); });
        // Append any column the app now sends that the sheet doesn't have yet
        headers.forEach(function(h) {
          if (sheetHeaders.indexOf(h) === -1) {
            sheetHeaders.push(h);
            sheet.getRange(1, sheetHeaders.length).setValue(h);
          }
        });
      }

      // Map values to the sheet's actual column order (handles reordering & new cols)
      var mappedRow = sheetHeaders.map(function(h) {
        var idx = headers.indexOf(h);
        return idx >= 0 ? rowVals[idx] : '';
      });

      // UPDATE existing row if ID matches, else append
      var rowIdx = findRowById_(sheet, mappedRow[0]);
      if (rowIdx > -1) {
        sheet.getRange(rowIdx, 1, 1, mappedRow.length).setValues([mappedRow]);
      } else {
        sheet.appendRow(mappedRow);
      }

      if (itemCfg) {
        var items = itemsJson ? JSON.parse(itemsJson) : [];
        var laborItems = laborJson ? JSON.parse(laborJson) : [];
        writeChildRows_(ss, itemCfg, mappedRow[0], items, laborItems);
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
// column A — not getDataRange(), which pulls every column of every row —
// so lookups stay cheap as a sheet grows into the thousands of rows.
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
      var headerSheets = ['Requisitions','PurchaseOrders','SiteReceipts','Works','Notifications'];
      var result = {};
      headerSheets.forEach(function(name) {
        var rows = readSheetAsObjects_(ss, name);
        var cfg = ITEM_TABS[name];
        if (cfg) {
          var childRows = readSheetAsObjects_(ss, cfg.sheet);
          var byParent = {};
          childRows.forEach(function(r) {
            var pid = r[cfg.headers[0]]; // parent-ID column is always first
            (byParent[pid] = byParent[pid] || []).push(r);
          });
          rows.forEach(function(r) {
            var grouped = cfg.fromRows(byParent[r['ID']] || []);
            r[cfg.itemsHeaderName] = JSON.stringify(grouped.items);
            r[cfg.laborHeaderName] = JSON.stringify(grouped.laborItems);
          });
        }
        result[name] = rows;
      });
      result.MasterData = [{ 'Key': 'masterData', 'Data (JSON)': JSON.stringify(buildMasterData_(ss)) }];
      return jsonOut_(result);
    }
    return jsonOut_({status:'ok'});
  } catch (err) {
    return jsonOut_({status:'error', error: err.message});
  }
}
