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
    document.getElementById("change-fill").onclick = changeFillColor;
  }
});



// async function run() {
//   await Excel.run(async (context) => {
//     const range = context.workbook.getSelectedRange();
//     console.log(context);
//     // Read the range address
//     range.load("address");

//     // Update the fill color
//     range.format.fill.color = "Pink";


//   await context.sync();
//   //   return await context.sync();
//   // })
//   // .catch(function (error) {
//   //   console.log("Error: " + error);
//   //   if (error instanceof OfficeExtension.Error) {
//   //     console.log("Debug info: " + JSON.stringify(error.debugInfo));
//   //   }
//   });
// }


// await Excel.run(async (context) => {


//   await context.sync();
// });


// changing selected worksheet fill color 

async function changeFillColor() {
  await Excel.run(async (context) => {
    const color = document.getElementById("color").value;
    const color_code = document.getElementById("color_code").value;
    console.log(color,color_code);
    const range = context.workbook.getSelectedRange();
    range.load("address");
    if(color_code){
      range.format.fill.color = color_code;
    }
    else{
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
    const studentTable = currentWorksheet.tables.add("G1:J1", true /*hasHeaders*/);
    expensesTable.name = "ExpensesTable";
    studentTable.name = "StudentTable"

    // TODO2: Queue commands to populate the table with data.
    expensesTable.getHeaderRowRange().values =
      [["Date", "Merchant", "Category", "Amount","Positive"]];
    studentTable.getHeaderRowRange().values =
      [["ID", "Name", "Class", "Result"]];

    expensesTable.rows.add(null /*add at the end*/, [
      ["1/1/2017", "The Phone Company", "Communications", "420","0"],
      ["1/2/2017", "Northwind Electric Cars", "Transportation", "142.33","0"],
      ["1/5/2017", "Best For You Organics Company", "Groceries", "27.9","1"],
      ["1/10/2017", "Coho Vineyard", "Restaurant", "33","0"],
      ["1/11/2017", "Bellows College", "Education", "350.1","0"],
      ["1/15/2017", "Trey Research", "Other", "135","0"],
      ["1/15/2017", "Best For You Organics Company", "Groceries", "97.88","0"]
    ]);


    studentTable.rows.add(null, [
      ["01","Ayat","9","3.5"],
      ["02","Rahat","2","1.5"],
      ["09","Fayed","1","4.5"],
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