/*jslint browser, white*/
/*global SpreadsheetApp*/

function getAuth() {
  "use strict";
  return;
}

function setCellFocus(mySheet, myRange, taskNum) {
  "use strict";
  var rowData = myRange.getDisplayValues();
  var newRange = {};
  rowData.some(
    function (curVal, index) {
      if (curVal[3] === taskNum) {
        newRange = mySheet.getRange("E" + (index + 3));
        mySheet.setActiveRange(newRange);
        return true;
      }
    });
}

function taskUrl(taskStr) {
  "use strict";
  return "=HYPERLINK(" +
    "\"http://clientservices/amsweb/TaskView.aspx?TaskNumber=" +
      taskStr + "\"" + ", \"" + taskStr + "\")";
}

function onEdit() {
  "use strict";
  var mySheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var mySheetName = mySheet.getName();
  var sheetArray = ["Current", "FS Oncall"];
  var myRange;
  var vals;
  var fRowDate;
  var fRowTime;
  var fRowSite;
  var fRowTask;
  if (sheetArray.indexOf(mySheetName) >= 0) {
    // get populated rows and columns
    myRange = mySheet.getDataRange().offset(2, 0);
    // get values from columns A (date), B (time),
    // C (site), and D (task) in first row
    vals = mySheet.getRange("A3:D3").getValues();
    fRowDate = vals[0][0];
    fRowTime = vals[0][1];
    fRowSite = vals[0][2];
    fRowTask = vals[0][3];
    // start sort if first row has values for date, time, site, and task
    if (fRowDate && fRowTime && fRowSite && fRowTask) {
      // create url from task number
      if (/\d{7,}/.test(Number(fRowTask))) {
        myRange.getCell(1, 4)
        .setValue(taskUrl(fRowTask));
      }
      // Sort descending by column A then column B
      myRange.sort(
        [{
          column: 1,
          ascending: false
        },
         {
           column: 2,
           ascending: false
         }
        ]
      );
      /* *
      * insertRows does not add rows to the sheet. A "Service error"
      * will occur if there are no blank rows in the sheet.
      */
      mySheet.insertRows(3);
      // after sort find new row and move focus to that row
      setCellFocus(mySheet, myRange, fRowTask);
    }
  }
}

function onOpen() {
  "use strict";
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var menuEntries = [{
    name: "Current Date and Time",
    functionName: "enterNewTask"
  },
                     {
                       name: "Re-sort sheet",
                       functionName: "mySort"
                     }
                    ];
  ss.addMenu("MEDITECH Tools", menuEntries);
  /*
  ss.toast(
  "Spreadsheet will auto-sort after Date, Time, " +
  "Site, and Task have been entered in row 2.",
  "Use the MEDITECH Tools menu to enter a new issue.",
  -1
  );
  */
}

function enterNewTask() {
  "use strict";
  var ss = SpreadsheetApp.getActive();
  var currentSheet = ss.getActiveSheet();
  var mySheetName = currentSheet.getName();
  var sheetArray = ["Current", "FS Oncall"];
  var tempValues;
  var newRange;
  var tempDate;
  if (sheetArray.indexOf(mySheetName) >= 0) {
    tempValues = [];
    tempValues[0] = [];
    tempDate = new Date();
    tempValues[0][0] = tempDate.toLocaleDateString();
    tempValues[0][1] = tempDate.toTimeString().split(" ")[0];
    currentSheet.getRange(3, 1, 1, 2).setValues(tempValues);
    newRange = currentSheet.getRange("C3");
    currentSheet.setActiveRange(newRange);
  }
}

function mySort() {
  "use strict";
  var mySheet = SpreadsheetApp.getActiveSpreadsheet()
  .getActiveSheet();
  var mySheetName = mySheet.getName();
  var sheetArray = ["Current", "FS Oncall"];
  var myRange;
  if (sheetArray.indexOf(mySheetName) >= 0) {
    // get populated rows and columns
    myRange = mySheet.getDataRange().offset(2, 0);
    // Sort descending by column A then column B
    myRange.sort(
      [{
        column: 1,
        ascending: false
      },
       {
         column: 2,
         ascending: false
       }
      ]
    );
    /* *
    * insertRows does not add rows to the sheet. A "Service error"
    * will occur if there are no blank rows in the sheet.
    */
    mySheet.insertRows(3);
  }
}
