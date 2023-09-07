/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    // Determine if the user's version of Office supports all the Office.js APIs that are used in the tutorial.
    if (!Office.context.requirements.isSetSupported('ExcelApi', '1.7')) {
      console.log('Sorry. The tutorial add-in uses Excel.js APIs that are not available in your version of Office.');
    }
    // document.getElementById("run").onclick = run;
    document.getElementById("create-table").onclick = createTable;
    document.getElementById("filter-table").onclick = filterTable;
    document.getElementById("sort-asc").onclick = () => sortTable();
    document.getElementById("sort-dsc").onclick = () => sortTable(ascending = false);
    document.getElementById("change-fill").onclick = changeFillColor;
    document.getElementById("change-sheet").onclick = activeWorksheet;
    document.getElementById("create-chart").onclick = createChart;
    document.getElementById("rename-column").onclick = renameColumn;
    document.getElementById("get-data").onclick = getTableData;
    document.getElementById("newSheet").onclick = createSheet;
    document.getElementById("create-table-with-formula").onclick = createTableWithCalculation;
  }
});


// changing selected worksheet fill color 

async function changeFillColor() {
  await Excel.run(async (context) => {
    const color = document.getElementById("color").value;
    const color_code = document.getElementById("color_code").value;
    console.log(color, color_code);

    const range = context.workbook.getSelectedRange();
    range.load("address");
    if (color_code) {
      range.format.fill.color = color_code;
    }
    else {
      range.format.fill.color = color;

    }
    await context.sync();
  })
    .catch(function (error) {
      console.log("Invalid color");
    });
}


// creating a table
async function createTable() {
  await Excel.run(async (context) => {

    // TODO1: Queue table creation logic here.
    const currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
    const expensesTable = currentWorksheet.tables.add("A1:E1", true /*hasHeaders*/);
    const studentTable = currentWorksheet.tables.add("G9:J9", true /*hasHeaders*/);
    expensesTable.name = "ExpensesTable";
    studentTable.name = "StudentTable"

    // TODO2: Queue commands to populate the table with data.
    expensesTable.getHeaderRowRange().values =
      [["Date", "Merchant", "Category", "Amount", "Positive"]];
    studentTable.getHeaderRowRange().values =
      [["ID", "Name", "Class", "Result"]];

    expensesTable.rows.add(null /*add at the end*/, [
      ["1/1/2017", "The Phone Company", "Communications", "420", "0"],
      ["1/2/2017", "Northwind Electric Cars", "Transportation", "142.33", "0"],
      ["1/5/2017", "Best For You Organics Company", "Groceries", "27.9", "1"],
      ["1/10/2017", "Coho Vineyard", "Restaurant", "33", "0"],
      ["1/11/2017", "Bellows College", "Education", "350.1", "0"],
      ["1/15/2017", "Trey Research", "Other", "135", "0"],
      ["1/15/2017", "Best For You Organics Company", "Groceries", "97.88", "0"]
    ]);


    studentTable.rows.add(null, [
      ["01", "Ayat", "9", "3.5"],
      ["02", "Rahat", "2", "1.5"],
      ["09", "Fayed", "1", "4.5"],
    ])


    // TODO3: Queue commands to format the table.
    expensesTable.columns.getItemAt(3).getRange().numberFormat = [['\u20AC#,##0.00']];
    expensesTable.columns.getItemAt(4).getRange().numberFormat = [['General']];
    expensesTable.getRange().format.autofitColumns();
    expensesTable.getRange().format.autofitRows();

    studentTable.columns.getItemAt(0).getRange().numberFormat = [['General']];
    studentTable.columns.getItemAt(2).getRange().numberFormat = [['General']];
    studentTable.columns.getItemAt(3).getRange().numberFormat = [['General']];
    studentTable.getRange().format.autofitColumns();
    studentTable.getRange().format.autofitRows();


    await context.sync();
  })
    .catch(function (error) {
      console.log("Error: " + error);
      if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
      }
    });
}

// filtering a table
async function filterTable() {
  await Excel.run(async (context) => {

    const currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
    const expensesTable = currentWorksheet.tables.getItem('ExpensesTable');
    const categoryFilter = expensesTable.columns.getItem('Category').filter;
    categoryFilter.applyValuesFilter(['Education', 'Groceries']);

    await context.sync();
  })
    .catch(function (error) {
      console.log("Error: " + error);
      if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
      }
    });
}

// sorting a table 

async function sortTable(ascending = true) {
  await Excel.run(async (context) => {

    const currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
    const expensesTable = currentWorksheet.tables.getItem('ExpensesTable');
    const sortFields = [
      {
        key: 0,            // Merchant column
        ascending: ascending,
      }
    ];

    expensesTable.sort.apply(sortFields);

    await context.sync();
  })
    .catch(function (error) {
      console.log("Error: " + error);
      if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
      }
    });
}


// run formula

async function createTableWithCalculation() {
  await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getActiveWorksheet();

    // Create the headers and format them to stand out.
    let headers = [
      ["Product", "Quantity", "Unit Price", "Totals"]
    ];
    let headerRange = sheet.getRange("B2:E2");
    headerRange.values = headers;
    headerRange.format.fill.color = "#4472C4";
    headerRange.format.font.color = "black";

    // Create the product data rows.
    let productData = [
      ["Almonds", 6, 7.5],
      ["Coffee", 20, 34.5],
      ["Chocolate", 10, 9.56],
    ];
    let dataRange = sheet.getRange("B3:D5");
    dataRange.values = productData;

    // Create the formulas to total the amounts sold.
    let totalFormulas = [
      ["=C3 * D3"],
      ["=C4 * D4"],
      ["=C5 * D5"],
      ["=SUM(E3:E5)"]
    ];
    let totalRange = sheet.getRange("E3:E6");
    // dataRange.format.font.color = "black";
    totalRange.format.font.color = "black";
    totalRange.formulas = totalFormulas;
    totalRange.format.font.bold = true;

    // Display the totals as US dollar amounts.
    totalRange.numberFormat = [["$0.00"]];
    // creating a table
    sheet.tables.add("B2:E6", true);
    await context.sync();
  })
    .catch(function (error) {
      console.log("Error: " + error);
      if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
      }
    });
}

async function createChart() {
  await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getActiveWorksheet();
    let chart = sheet.charts.add(Excel.ChartType.columnStacked, sheet.getRange("C3:D5"));
    // chart.setPosition("A9", "D10");
    chart.setPosition("A9");


    chart.title.text = "Expenses";
    chart.legend.position = "top"
    chart.legend.format.fill.setSolidColor("white");
    chart.dataLabels.format.font.size = 10;
    chart.dataLabels.format.font.color = "white";
    chart.series.getItemAt(0).name = 'Value in euro;';


    // chart.left = 0; // Set the left position to 0 (column A)
    // chart.top = 150;
    await context.sync();
  })
}


// change active worksheet 

async function activeWorksheet() {
  await Excel.run(async (context) => {
    naame = document.getElementById("sheetName").value;
    // console.log(naame);
    context.workbook.worksheets.getItem(`${naame}`).activate();
    await context.sync();
  })
    .catch(function (error) {
      console.log("Error: " + error);

      if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
      }
    });
}




async function createSheet() {
  await Excel.run(async (context) => {
    naame = document.getElementById("newSheetName").value;
    // console.log(naame);
    context.workbook.worksheets.add(naame);
    // context.workbook.worksheets.getItem(naame).freezePanes.freezeRows(1)
    // context.workbook.worksheets.getItem(naame).freezePanes.unfreeze();
    // context.workbook.worksheets.getItem(naame).delete();


    await context.sync();

    //toggle worksheet protection
    // const currWorksheet = context.workbook.worksheets.getActiveWorksheet().load("protection/protected");
    // return context.sync().then(() => {
    //   if (currWorksheet.protection.protected)
    //     currWorksheet.protection.unprotect();
    //   else
    //     currWorksheet.protection.protect();
    // }).then(context.sync);

  })
    .catch(function (error) {
      console.log("Error: " + error);

      if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
      }
    });
}


// rename column
async function renameColumn() {
  await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getItem("Sheet1");

    let expensesTable = sheet.tables.getItem("ExpensesTable");
    expensesTable.columns.load("items");

    await context.sync();

    expensesTable.columns.items[0].name = "Purchase date";

    sheet.getUsedRange().format.autofitColumns();
    sheet.getUsedRange().format.autofitRows();

    await context.sync();
  })
    .catch(function (error) {
      console.log("Error: " + error);

      if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
      }
    });
}

// add table new column
// await Excel.run(async (context) => {
//   let sheet = context.workbook.worksheets.getItem("Sample");
//   let expensesTable = sheet.tables.getItem("ExpensesTable");

//   expensesTable.columns.add(null /*add columns to the end of the table*/, [
//       ["Day of the Week"],
//       ["Saturday"],
//       ["Friday"],
//       ["Monday"],
//       ["Thursday"],
//       ["Sunday"],
//       ["Saturday"],
//       ["Monday"]
//   ]);

//   sheet.getUsedRange().format.autofitColumns();
//   sheet.getUsedRange().format.autofitRows();

//   await context.sync();
// });



// add a column that adds formulas
// await Excel.run(async (context) => {
//   let sheet = context.workbook.worksheets.getItem("Sample");
//   let expensesTable = sheet.tables.getItem("ExpensesTable");

//   expensesTable.columns.add(null /*add columns to the end of the table*/, [
//       ["Type of the Day"],
//       ['=IF(OR((TEXT([DATE], "dddd") = "Saturday"), (TEXT([DATE], "dddd") = "Sunday")), "Weekend", "Weekday")'],
//       ['=IF(OR((TEXT([DATE], "dddd") = "Saturday"), (TEXT([DATE], "dddd") = "Sunday")), "Weekend", "Weekday")'],
//       ['=IF(OR((TEXT([DATE], "dddd") = "Saturday"), (TEXT([DATE], "dddd") = "Sunday")), "Weekend", "Weekday")'],
//       ['=IF(OR((TEXT([DATE], "dddd") = "Saturday"), (TEXT([DATE], "dddd") = "Sunday")), "Weekend", "Weekday")'],
//       ['=IF(OR((TEXT([DATE], "dddd") = "Saturday"), (TEXT([DATE], "dddd") = "Sunday")), "Weekend", "Weekday")'],
//       ['=IF(OR((TEXT([DATE], "dddd") = "Saturday"), (TEXT([DATE], "dddd") = "Sunday")), "Weekend", "Weekday")'],
//       ['=IF(OR((TEXT([DATE], "dddd") = "Saturday"), (TEXT([DATE], "dddd") = "Sunday")), "Weekend", "Weekday")']
//   ]);

//   sheet.getUsedRange().format.autofitColumns();
//   sheet.getUsedRange().format.autofitRows();

//   await context.sync();
// });



async function getTableData () {
  await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getItem("Sheet1");
    let expensesTable = sheet.tables.getItem("ExpensesTable");

    // Get data from the header row.
    let headerRange = expensesTable.getHeaderRowRange().load("values");

    // Get data from the table.
    let bodyRange = expensesTable.getDataBodyRange().load("values");

    // Get data from a single column.
    let columnRange = expensesTable.columns.getItem("Merchant").getDataBodyRange().load("values");

    // Get data from a single row.
    let rowRange = expensesTable.rows.getItemAt(1).load("values");

    // Sync to populate proxy objects with data from Excel.
    await context.sync();

    let headerValues = headerRange.values;
    let bodyValues = bodyRange.values;
    let merchantColumnValues = columnRange.values;
    let secondRowValues = rowRange.values;

    // Write data from table back to the sheet
    sheet.getRange("A11:A11").values = [["Results"]];
    sheet.getRange("A13:E13").values = headerValues;
    sheet.getRange("A14:E20").values = bodyValues;
    sheet.getRange("B23:B29").values = merchantColumnValues;
    sheet.getRange("A32:E32").values = secondRowValues;

    // Sync to update the sheet in Excel.
    await context.sync();
});
}