// Instructions tab on open
function onOpen() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("INSTRUCTIONS"); 
  if (sheet) {
    ss.setActiveSheet(sheet);
  }
}

function onEdit(e) {
  var range = e.range;
  var sheet = range.getSheet();
  var sheetName = "2026 FUND MONITORING"; //Update to reflect the new sheet name
  
  // Only run on the specified sheet
  if (sheet.getName() !== sheetName) return; 

  var row = range.getRow();
  var col = range.getColumn();
  var val = range.getValue();

  // Skip header rows (1-2)
  if (row <= 2) return;

//1. SAA Dropdown*
if (col === 1) {
    var cellB = sheet.getRange(row, 2);
    
    if (val === "SAA") {
      var options = ["DO 2026-0062","DO 2026-0144", "DO 2026-0194","DO 2026-0209","DO 2026-0221","Others"]; //Update to reflect the new DOs
      var rule = SpreadsheetApp.newDataValidation()
        .requireValueInList(options)
        .setAllowInvalid(false)
        .build();
      cellB.setDataValidation(rule);
    } else {
      cellB.clear({contentsOnly: true, validationsOnly: true}); 
      cellB.clearDataValidation(); 
      SpreadsheetApp.flush(); 
    }
    return; 
  }
  
  //2. Duplicate Checker*
  if (col === 2 || col === 3 || col === 4) { 
    if (checkDuplicates(sheet, row)) {
      range.clearContent(); 
      SpreadsheetApp.getUi().alert("This record already exists. Please review the other rows for any duplicates.");
      return;
    }
  }

function normalizeValue(val) {
  if (val === null || val === undefined) return "";
  return String(val).toLowerCase().replace(/[^a-z0-9]/g, "");
}

function toNumber(value) {
  if (value === "" || value === null) return 0;
  return Number(String(value).replace(/,/g, ""));
}

function checkDuplicates(sheet, currentRow) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 3) return false;
  var curIssuanceRaw = sheet.getRange(currentRow, 2).getValue();
  var curActivityRaw = sheet.getRange(currentRow, 3).getValue();
  var curAllotment = sheet.getRange(currentRow, 4).getValue();

  if (curIssuanceRaw === "" || curActivityRaw === "" || curAllotment === "") return false;

  var curIssuanceNorm = normalizeValue(curIssuanceRaw);
  var curActivityNorm = normalizeValue(curActivityRaw);
  var data = sheet.getRange(3, 2, lastRow - 2, 3).getValues(); 

  for (var i = 0; i < data.length; i++) {
    if (i + 3 === currentRow) continue;
    if (normalizeValue(data[i][0]) === curIssuanceNorm && 
        normalizeValue(data[i][1]) === curActivityNorm && 
        data[i][2] === curAllotment) {
      return true; 
    }
  }
  return false; 
}

  //3. Sequential Encoding Validation*
  var skipCols = [2, 8, 10, 11, 12, 13, 14, 15, 17, 19, 20, 21, 22, 23, 24, 26, 28, 29, 30, 31, 32, 33, 35, 37, 38, 39, 40, 41, 42];
  if (col > 1 && !skipCols.includes(col)) {
    var prevCol = col - 1;
    while (skipCols.includes(prevCol) && prevCol > 1) {
      prevCol--;
    }

    var prevCell = sheet.getRange(row, prevCol).getValue();
    if (prevCell === "" || prevCell === null) {
      range.clearContent();
      SpreadsheetApp.getUi().alert("Please complete the previous mandatory column first.");
      sheet.getRange(row, prevCol).activate();
      return;
    }
  }

  //4. SAA Data Validation*
  if (col === 3) {
    var colAValue = sheet.getRange(row, 1).getValue(); 
    var colBValue = sheet.getRange(row, 2).getValue(); 
    if (colAValue === "SAA" && (colBValue === "" || colBValue === null)) {
      range.clearContent();
      SpreadsheetApp.getUi().alert("Column B cannot be blank if Funding Source is SAA.");
      sheet.getRange(row, 2).activate();
      return;
    }
  }

  //5. Numbers Only Validation*
  var numericCols = [4, 7, 9, 16, 18, 25, 27, 34, 36]; 
  if (numericCols.includes(col)) {
    if (val !== "" && isNaN(Number(String(val).replace(/,/g, "")))) {
      range.clearContent();
      SpreadsheetApp.getUi().alert("This column accepts numbers only.");
      range.activate();
      return;
    }
  }

  //6. Amount Obligated Validation*
  var obligatedCols = [7, 16, 25, 34]; 
  if (obligatedCols.includes(col)) {
    var allotment = toNumber(sheet.getRange(row, 4).getValue());
    var amountEntered = toNumber(val);
    if (amountEntered > allotment) {
      range.clearContent();
      SpreadsheetApp.getUi().alert("AMOUNT OBLIGATED exceeds the ALLOTMENT AMOUNT.");
      return;
    }
  }

  //7. Amount Disbursed Validation*
  var disbursedCols = [9, 18, 27, 36]; 
  if (disbursedCols.includes(col)) {
    var values = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];
    var totalDisbursed, totalObligated;

    switch(col) {
      case 9:
        totalDisbursed = toNumber(values[8]);
        totalObligated = toNumber(values[6]);
        break;
      case 18:
        totalDisbursed = toNumber(values[8]) + toNumber(values[17]);
        totalObligated = toNumber(values[6]) + toNumber(values[15]);
        break;
      case 27:
        totalDisbursed = toNumber(values[8]) + toNumber(values[17]) + toNumber(values[26]);
        totalObligated = toNumber(values[6]) + toNumber(values[15]) + toNumber(values[24]);
        break;
      case 36:
        totalDisbursed = toNumber(values[8]) + toNumber(values[17]) + toNumber(values[26]) + toNumber(values[35]);
        totalObligated = toNumber(values[6]) + toNumber(values[15]) + toNumber(values[24]) + toNumber(values[33]);
        break;
    }

    if (totalDisbursed > totalObligated) {
      range.clearContent();
      SpreadsheetApp.getUi().alert("AMOUNT DISBURSED exceeds AMOUNT OBLIGATED.");
      return;
    }
  }

  //8. Unmerge Merged Cells*
  var mergedRanges = range.getMergedRanges();
  if (mergedRanges.length > 0) {
    for (var i = 0; i < mergedRanges.length; i++) {
      mergedRanges[i].breakApart();
    }
    SpreadsheetApp.getUi().alert("Merged cells are not allowed.");
  }
} 
