/**
 * HPCL Active Inventory Portal - Google Apps Script (Push & Pull Sync)
 * 
 * Instructions:
 * 1. Open your Google Sheet -> Extensions -> Apps Script
 * 2. Delete any existing code and paste this entire file.
 * 3. Click "Deploy" -> "New Deployment" -> Select type "Web App".
 * 4. Set Description to "HPCL Inventory Sync".
 * 5. Set "Execute as": "Me".
 * 6. Set "Who has access": "Anyone" (Crucial for mobile & browser sync).
 * 7. Click "Deploy", grant permissions, and copy the Web App URL!
 * 8. Paste the Web App URL into the HPCL Inventory Portal -> Backup & Sync tab.
 */

function doPost(e) {
  try {
    var contents = e.postData ? e.postData.contents : "";
    var data = {};
    if (contents) {
      try { data = JSON.parse(contents); } catch(err) {}
    }
    
    // Check if this is a Pull request via POST
    if (data.action === "pull") {
      return pullData();
    }
    
    // Otherwise it is a Push request: write incoming inventory data to sheets
    pushData(data);
    return ContentService.createTextOutput(JSON.stringify({
      status: "success",
      message: "Data successfully synced to Google Sheets!"
    })).setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({
      status: "error",
      message: err.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  // GET request returns the full JSON database for Pull operations
  return pullData();
}

function pullData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var result = {
    inventory: { pressure_dus: [], stps: [] },
    inword: [], stp_outword: [], du_outword: [], tank_outword: [], others_outword: [], gate_passes: []
  };
  
  function readSheet(sheetNames) {
    if (!Array.isArray(sheetNames)) sheetNames = [sheetNames];
    var sheet = null;
    for (var k = 0; k < sheetNames.length; k++) {
      sheet = ss.getSheetByName(sheetNames[k]);
      if (sheet) break;
    }
    if (!sheet) {
      // Case-insensitive match
      var allSheets = ss.getSheets();
      for (var s = 0; s < allSheets.length; s++) {
        var sName = allSheets[s].getName().toLowerCase().replace(/[\s_]/g, '');
        for (var k = 0; k < sheetNames.length; k++) {
          if (sName === sheetNames[k].toLowerCase().replace(/[\s_]/g, '')) {
            sheet = allSheets[s];
            break;
          }
        }
        if (sheet) break;
      }
    }
    if (!sheet) return [];
    var data = sheet.getDataRange().getValues();
    if (data.length < 2) return [];
    var headers = data[0];
    var rows = [];
    for (var i = 1; i < data.length; i++) {
      var rowObj = {};
      var hasValue = false;
      for (var j = 0; j < headers.length; j++) {
        var val = data[i][j];
        if (val !== "" && val !== null && val !== undefined) hasValue = true;
        rowObj[headers[j]] = val;
      }
      if (hasValue) rows.push(rowObj);
    }
    return rows;
  }
  
  result.inventory.pressure_dus = readSheet(["Pressure_DUs", "Inventory", "Pressure DUs", "DUs"]);
  result.inventory.stps = readSheet(["STPs", "STP"]);
  result.inword = readSheet(["Inword", "Inward"]);
  result.stp_outword = readSheet(["STP_Outword", "STP_Outward", "STP OUTWORD"]);
  result.du_outword = readSheet(["DU_Outword", "DU_Outward", "DU OUTWORD"]);
  result.tank_outword = readSheet(["Tank_Outword", "Tank_Outward", "TANK OUTWORD"]);
  result.others_outword = readSheet(["Others_Outword", "Others_Outward", "OTHERS OUTWORD"]);
  result.gate_passes = readSheet(["EMR", "EMRs", "EMR_Records", "Gate_Passes", "Gate Passes"]);
  
  return ContentService.createTextOutput(JSON.stringify({
    status: "success",
    data: result
  })).setMimeType(ContentService.MimeType.JSON);
}

function pushData(data) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  function writeSheet(preferredName, possibleNames, headers, rows) {
    if (!rows || !Array.isArray(rows)) return;
    if (!Array.isArray(possibleNames)) possibleNames = [possibleNames];
    var sheet = null;
    for (var k = 0; k < possibleNames.length; k++) {
      sheet = ss.getSheetByName(possibleNames[k]);
      if (sheet) break;
    }
    if (!sheet) {
      sheet = ss.insertSheet(preferredName);
    }
    sheet.clear();
    var payload = [headers];
    for (var i = 0; i < rows.length; i++) {
      var row = [];
      for (var j = 0; j < headers.length; j++) {
        var key = headers[j];
        var val = rows[i][key] !== undefined && rows[i][key] !== null ? rows[i][key] : "";
        // Fallback checks for Gate Pass vs EMR compatibility
        if (val === "" && key === "gate_pass_no" && rows[i]["emr_no"] !== undefined) val = rows[i]["emr_no"];
        if (val === "" && key === "gate_pass_date" && rows[i]["emr_date"] !== undefined) val = rows[i]["emr_date"];
        if (val === "" && key === "emr_no" && rows[i]["gate_pass_no"] !== undefined) val = rows[i]["gate_pass_no"];
        if (val === "" && key === "emr_date" && rows[i]["gate_pass_date"] !== undefined) val = rows[i]["gate_pass_date"];
        row.push(val);
      }
      payload.push(row);
    }
    if (payload.length > 0) {
      sheet.getRange(1, 1, payload.length, headers.length).setValues(payload);
    }
  }
  
  if (data.inventory && data.inventory.pressure_dus) {
    writeSheet("Pressure_DUs", ["Pressure_DUs", "Inventory", "Pressure DUs"], ["sno", "type", "variant", "make", "model", "serial_no", "receipt_date", "status", "remarks"], data.inventory.pressure_dus);
  }
  if (data.inventory && data.inventory.stps) {
    writeSheet("STPs", ["STPs", "STP"], ["sno", "hp_type", "make", "serial_no", "model", "receipt_date", "status", "remarks"], data.inventory.stps);
  }
  if (data.inword) {
    writeSheet("Inword", ["Inword", "Inward"], ["item", "received_from", "received_date", "qty"], data.inword);
  }
  if (data.stp_outword) {
    writeSheet("STP_Outword", ["STP_Outword", "STP_Outward", "STP OUTWORD"], ["sno", "hp_type", "make", "serial_no", "gate_pass_date", "gate_pass_no", "destination", "vendor", "out_date", "emr_date", "emr_no"], data.stp_outword);
  }
  if (data.du_outword) {
    writeSheet("DU_Outword", ["DU_Outword", "DU_Outward", "DU OUTWORD"], ["sno", "type", "variant", "make", "model", "serial_no", "gate_pass_date", "gate_pass_no", "destination", "vendor", "out_date", "emr_date", "emr_no"], data.du_outword);
  }
  if (data.tank_outword) {
    writeSheet("Tank_Outword", ["Tank_Outword", "Tank_Outward", "TANK OUTWORD"], ["sno", "tank_capacity", "gate_pass_date", "gate_pass_no", "destination", "vendor", "emr_date", "emr_no"], data.tank_outword);
  }
  if (data.others_outword) {
    writeSheet("Others_Outword", ["Others_Outword", "Others_Outward", "OTHERS OUTWORD"], ["sno", "material", "qty", "gate_pass_date", "gate_pass_no", "destination", "vendor", "emr_date", "emr_no"], data.others_outword);
  }
  if (data.gate_passes) {
    writeSheet("EMR", ["EMR", "EMRs", "EMR_Records", "Gate_Passes", "Gate Passes"], ["emr_no", "category", "serial_no", "make", "model", "vendor", "destination", "out_date", "gate_pass_no", "vehicle_no", "driver", "remarks"], data.gate_passes);
  }
}
