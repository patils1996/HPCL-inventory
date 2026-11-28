/**
 * HPCL Active Inventory Portal - Google Apps Script (Push & Pull Sync)
 * Version 3.1
 *
 * Instructions:
 * 1. Open your Google Sheet -> Extensions -> Apps Script
 * 2. Delete any existing code and paste this entire file.
 * 3. Click "Deploy" -> "New Deployment" -> Select type "Web App".
 *    (Already deployed an older version? Use "Deploy" -> "Manage deployments" -> edit (pencil)
 *     -> Version: "New version" -> Deploy, so the /exec URL stays the same.)
 * 4. Set Description to "HPCL Inventory Sync".
 * 5. Set "Execute as": "Me".
 * 6. Set "Who has access": "Anyone" (Crucial for mobile & browser sync).
 * 7. Click "Deploy", grant permissions, and copy the Web App URL!
 * 8. Paste the Web App URL into the HPCL Inventory Portal -> Backup & Sync tab.
 *
 * What changed in 3.0
 * - New "Dispatch_Letters" sheet (office dispatch letters and their status).
 * - EMR sheet gains dl_no, dispatched_by, photos_json and other_items_json columns.
 * - New "upload_photo" action: dispatch photos are saved to Google Drive in a folder
 *   "HPCL Dispatch Photos/<EMR no>" and shared as "anyone with the link can view" so other
 *   devices can show them. Redeploying asks you to allow Drive access.
 *
 * What changed in 2.1
 * - Multi-asset EMRs keep every asset: the full asset list is stored in the EMR sheet's
 *   "items_json" column and rebuilt on pull (previously only the first row survived a round-trip).
 * - Dates are returned as plain DD/MM/YYYY in the spreadsheet's time zone (previously they came back
 *   as UTC timestamps, which showed the previous day in India).
 * - All cells are written as plain text, so serial numbers such as "072163" or "202412000642" are
 *   not turned into numbers / scientific notation by Sheets.
 * - A script lock prevents two devices from overwriting each other mid-write.
 */

var SHEETS = {
  pressure_dus: {
    name: "Pressure_DUs", aliases: ["Pressure_DUs", "Inventory", "Pressure DUs", "DUs"],
    headers: ["sno", "type", "variant", "make", "model", "serial_no", "receipt_date", "status", "remarks"]
  },
  stps: {
    name: "STPs", aliases: ["STPs", "STP"],
    headers: ["sno", "hp_type", "make", "serial_no", "model", "receipt_date", "status", "remarks"]
  },
  inword: {
    name: "Inword", aliases: ["Inword", "Inward"],
    headers: ["item", "received_from", "received_date", "qty"]
  },
  stp_outword: {
    name: "STP_Outword", aliases: ["STP_Outword", "STP_Outward", "STP OUTWORD"],
    headers: ["sno", "hp_type", "make", "model", "serial_no", "status", "receipt_date", "gate_pass_date", "gate_pass_no", "destination", "vendor", "out_date", "emr_date", "emr_no"]
  },
  du_outword: {
    name: "DU_Outword", aliases: ["DU_Outword", "DU_Outward", "DU OUTWORD"],
    headers: ["sno", "type", "variant", "make", "model", "serial_no", "status", "receipt_date", "gate_pass_date", "gate_pass_no", "destination", "vendor", "out_date", "emr_date", "emr_no"]
  },
  tank_outword: {
    name: "Tank_Outword", aliases: ["Tank_Outword", "Tank_Outward", "TANK OUTWORD"],
    headers: ["sno", "tank_capacity", "gate_pass_date", "gate_pass_no", "destination", "vendor", "emr_date", "emr_no"]
  },
  others_outword: {
    name: "Others_Outword", aliases: ["Others_Outword", "Others_Outward", "OTHERS OUTWORD"],
    headers: ["sno", "material", "qty", "gate_pass_date", "gate_pass_no", "destination", "vendor", "emr_date", "emr_no"]
  },
  gate_passes: {
    name: "EMR", aliases: ["EMR", "EMRs", "EMR_Records", "Gate_Passes", "Gate Passes"],
    headers: ["emr_no", "category", "serial_no", "make", "model", "variant", "vendor", "destination", "out_date", "gate_pass_no", "vehicle_no", "driver", "remarks", "dl_no", "dispatched_by", "created_at", "items_json", "other_items_json", "photos_json"]
  },
  dispatch_letters: {
    name: "Dispatch_Letters", aliases: ["Dispatch_Letters", "Dispatch Letters", "Letters"],
    headers: ["dl_no", "dl_date", "status", "destination", "vendor", "purpose", "issued_by", "designation", "office", "ref", "remarks", "emr_no", "dispatched_on", "dispatched_by", "vehicle_no", "driver", "gate_pass_no", "cancel_reason", "created_at", "items_json", "other_items_json"]
  }
};

var PHOTO_ROOT_NAME = "HPCL Dispatch Photos";
var META_SHEET_NAME = "_Sync_Meta";

function isValidExecUrl(u) {
  var s = String(u || "");
  return s.indexOf("https://script.google.com/") === 0 && s.indexOf("/exec") !== -1 && s.indexOf("/dev") === -1;
}

function readSyncMeta(ss) {
  var sh = ss.getSheetByName(META_SHEET_NAME);
  var meta = { revision: "0", updated_at: "", exec_url: "" };
  if (!sh) return meta;
  var vals = sh.getDataRange().getValues();
  for (var i = 1; i < vals.length; i++) {
    var k = String(vals[i][0] || "").trim();
    var v = String(vals[i][1] || "").trim();
    if (k) meta[k] = v;
  }
  return meta;
}

function writeSyncMeta(ss, meta) {
  var sh = ss.getSheetByName(META_SHEET_NAME);
  if (!sh) sh = ss.insertSheet(META_SHEET_NAME);
  var rows = [
    ["key", "value"],
    ["revision", String(meta.revision || "1")],
    ["updated_at", String(meta.updated_at || new Date().toISOString())],
    ["exec_url", String(meta.exec_url || "")]
  ];
  sh.clear();
  var range = sh.getRange(1, 1, rows.length, 2);
  range.setNumberFormat("@");
  range.setValues(rows);
  sh.getRange(1, 1, 1, 2).setFontWeight("bold");
  sh.setFrozenRows(1);
}

function getPhotoFolder(emrNo) {
  var props = PropertiesService.getScriptProperties();
  var rootId = props.getProperty("PHOTO_ROOT_ID");
  var root = null;
  if (rootId) {
    try { root = DriveApp.getFolderById(rootId); } catch (err) { root = null; }
  }
  if (!root) {
    var it = DriveApp.getFoldersByName(PHOTO_ROOT_NAME);
    root = it.hasNext() ? it.next() : DriveApp.createFolder(PHOTO_ROOT_NAME);
    props.setProperty("PHOTO_ROOT_ID", root.getId());
  }
  var name = String(emrNo || "Unsorted").replace(/[\\/:*?"<>|]/g, "-");
  var subs = root.getFoldersByName(name);
  return subs.hasNext() ? subs.next() : root.createFolder(name);
}

function uploadPhoto(data) {
  if (!data.data || !data.photo_id) {
    return jsonResponse({ status: "error", message: "Photo data missing." });
  }
  var folder = getPhotoFolder(data.emr_no);
  var fileName = data.photo_id + ".jpg";
  // Re-uploads of the same photo return the existing file
  var existing = folder.getFilesByName(fileName);
  var file = existing.hasNext() ? existing.next() : null;
  if (!file) {
    var bytes = Utilities.base64Decode(data.data);
    var blob = Utilities.newBlob(bytes, data.mime || "image/jpeg", fileName);
    file = folder.createFile(blob);
    file.setDescription("EMR " + (data.emr_no || "") + (data.dl_no ? " / " + data.dl_no : "") + " / taken " + (data.taken_at || ""));
    try {
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    } catch (err) {
      // Some Google Workspace domains block link sharing; the file is still saved.
    }
  }
  return jsonResponse({ status: "success", file_id: file.getId(), url: file.getUrl() });
}

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  try {
    var contents = e && e.postData ? e.postData.contents : "";
    var data = null;
    if (contents) {
      try { data = JSON.parse(contents); } catch (err) { data = null; }
    }
    if (!data || typeof data !== "object") {
      return jsonResponse({ status: "error", message: "Request body is not valid JSON." });
    }

    // Pull request via POST
    if (data.action === "pull") {
      return pullData();
    }
    if (data.action === "upload_photo") {
      return uploadPhoto(data);
    }

    if (!data.inventory && !data.inword && !data.gate_passes && !data.du_outword && !data.stp_outword) {
      return jsonResponse({ status: "error", message: "No inventory data found in request." });
    }

    var lock = LockService.getScriptLock();
    if (!lock.tryLock(30000)) {
      return jsonResponse({ status: "error", message: "Another device is syncing right now. Please retry in a few seconds." });
    }
    var savedMeta;
    try {
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var currMeta = readSyncMeta(ss);
      if (!data._force && data._base_revision && currMeta.revision && currMeta.revision !== "0" &&
          String(data._base_revision) !== String(currMeta.revision)) {
        return jsonResponse({
          status: "conflict",
          message: "Another user updated Google Sheets while your copy was older.",
          revision: currMeta.revision,
          updated_at: currMeta.updated_at,
          exec_url: currMeta.exec_url || "",
          data: buildPullResult(ss)
        });
      }
      savedMeta = pushData(data, currMeta);
    } finally {
      lock.releaseLock();
    }

    return jsonResponse({
      status: "success",
      message: "Data successfully synced to Google Sheets!",
      synced_at: savedMeta.updated_at,
      updated_at: savedMeta.updated_at,
      revision: savedMeta.revision,
      exec_url: savedMeta.exec_url || ""
    });
  } catch (err) {
    return jsonResponse({ status: "error", message: err.toString() });
  }
}

function doGet(e) {
  // GET request returns the full JSON database for Pull operations
  try {
    return pullData();
  } catch (err) {
    return jsonResponse({ status: "error", message: err.toString() });
  }
}

function findSheet(ss, names) {
  for (var k = 0; k < names.length; k++) {
    var sh = ss.getSheetByName(names[k]);
    if (sh) return sh;
  }
  // Case / space / underscore-insensitive match
  var all = ss.getSheets();
  for (var s = 0; s < all.length; s++) {
    var sName = all[s].getName().toLowerCase().replace(/[\s_]/g, "");
    for (var j = 0; j < names.length; j++) {
      if (sName === names[j].toLowerCase().replace(/[\s_]/g, "")) return all[s];
    }
  }
  return null;
}

function buildPullResult(ss) {
  var tz = ss.getSpreadsheetTimeZone() || "Asia/Kolkata";

  function readSheet(def) {
    var sheet = findSheet(ss, def.aliases);
    if (!sheet) return [];
    var data = sheet.getDataRange().getValues();
    if (data.length < 2) return [];
    var headers = data[0].map(function (h) { return String(h).trim(); });
    var rows = [];
    for (var i = 1; i < data.length; i++) {
      var rowObj = {};
      var hasValue = false;
      for (var j = 0; j < headers.length; j++) {
        if (!headers[j]) continue;
        var val = data[i][j];
        if (val instanceof Date) {
          val = Utilities.formatDate(val, tz, "dd/MM/yyyy");
        }
        if (val !== "" && val !== null && val !== undefined) hasValue = true;
        rowObj[headers[j]] = val;
      }
      if (hasValue) rows.push(rowObj);
    }
    return rows;
  }

  // Turn every "<name>_json" column back into its array
  function unpackJson(rows) {
    rows.forEach(function (r) {
      Object.keys(r).forEach(function (k) {
        if (k.length > 5 && k.slice(-5) === "_json") {
          if (r[k]) {
            try { r[k.slice(0, -5)] = JSON.parse(r[k]); } catch (err) { /* leave flattened */ }
          }
          delete r[k];
        }
      });
    });
    return rows;
  }
  var gatePasses = unpackJson(readSheet(SHEETS.gate_passes));
  var letters = unpackJson(readSheet(SHEETS.dispatch_letters));

  return {
    inventory: {
      pressure_dus: readSheet(SHEETS.pressure_dus),
      stps: readSheet(SHEETS.stps)
    },
    inword: readSheet(SHEETS.inword),
    stp_outword: readSheet(SHEETS.stp_outword),
    du_outword: readSheet(SHEETS.du_outword),
    tank_outword: readSheet(SHEETS.tank_outword),
    others_outword: readSheet(SHEETS.others_outword),
    gate_passes: gatePasses,
    dispatch_letters: letters
  };
}

function pullData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var result = buildPullResult(ss);
  var meta = readSyncMeta(ss);
  return jsonResponse({
    status: "success",
    revision: meta.revision || "",
    updated_at: meta.updated_at || "",
    exec_url: meta.exec_url || "",
    data: result
  });
}

function pushData(data, currMeta) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  currMeta = currMeta || readSyncMeta(ss);

  function cellValue(row, key) {
    var val = row[key];
    if (key.length > 5 && key.slice(-5) === "_json") {
      var arr = row[key.slice(0, -5)];
      return Array.isArray(arr) && arr.length ? JSON.stringify(arr) : "";
    }
    // Gate Pass vs EMR field-name compatibility
    if ((val === undefined || val === null || val === "") && key === "gate_pass_no") val = row.emr_no;
    if ((val === undefined || val === null || val === "") && key === "gate_pass_date") val = row.emr_date;
    if ((val === undefined || val === null || val === "") && key === "emr_no") val = row.gate_pass_no;
    if ((val === undefined || val === null || val === "") && key === "emr_date") val = row.gate_pass_date;
    if (val === undefined || val === null) return "";
    if (typeof val === "object") return JSON.stringify(val);
    return String(val);
  }

  function writeSheet(def, rows) {
    if (!rows || !Array.isArray(rows)) return;
    var sheet = findSheet(ss, def.aliases);
    if (!sheet) sheet = ss.insertSheet(def.name);

    var payload = [def.headers];
    for (var i = 0; i < rows.length; i++) {
      var r = rows[i] || {};
      payload.push(def.headers.map(function (key) { return cellValue(r, key); }));
    }

    sheet.clear();
    var range = sheet.getRange(1, 1, payload.length, def.headers.length);
    range.setNumberFormat("@"); // plain text: keeps serials / leading zeros / dates exactly as sent
    range.setValues(payload);
    sheet.getRange(1, 1, 1, def.headers.length).setFontWeight("bold");
    sheet.setFrozenRows(1);
  }

  if (data.inventory && data.inventory.pressure_dus) writeSheet(SHEETS.pressure_dus, data.inventory.pressure_dus);
  if (data.inventory && data.inventory.stps) writeSheet(SHEETS.stps, data.inventory.stps);
  if (data.inword) writeSheet(SHEETS.inword, data.inword);
  if (data.stp_outword) writeSheet(SHEETS.stp_outword, data.stp_outword);
  if (data.du_outword) writeSheet(SHEETS.du_outword, data.du_outword);
  if (data.tank_outword) writeSheet(SHEETS.tank_outword, data.tank_outword);
  if (data.others_outword) writeSheet(SHEETS.others_outword, data.others_outword);
  if (data.gate_passes) writeSheet(SHEETS.gate_passes, data.gate_passes);
  if (data.dispatch_letters) writeSheet(SHEETS.dispatch_letters, data.dispatch_letters);

  var resolvedExec = "";
  if (isValidExecUrl(data._exec_url)) resolvedExec = String(data._exec_url).trim();
  else {
    try {
      var svcUrl = ScriptApp.getService().getUrl();
      if (isValidExecUrl(svcUrl)) resolvedExec = svcUrl;
    } catch (err) {}
  }
  if (!resolvedExec && isValidExecUrl(currMeta.exec_url)) resolvedExec = currMeta.exec_url;

  var nextMeta = {
    revision: String((parseInt(currMeta.revision || "0", 10) || 0) + 1),
    updated_at: new Date().toISOString(),
    exec_url: resolvedExec
  };
  writeSyncMeta(ss, nextMeta);
  return nextMeta;
}
