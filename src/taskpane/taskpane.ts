/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";

    document.getElementById("webscrape").onclick = webscrape;
  }
});

export async function webscrape() {
  try {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      
      /*
      const data = [
          ["Name", "Age", "City"],
          ["Alice", 25, "New York"],
          ["Bob", 30, "Los Angeles"],
          ["Charlie", 35, "Chicago"]
      ];

      let selectedRange = context.workbook.getSelectedRange();
      selectedRange.load("address");

      await context.sync();

      let firstCellAddress = selectedRange.address.split(":")[0]; 
      let startCell = sheet.getRange(firstCellAddress);

      let range = startCell.getResizedRange(data.length - 1, data[0].length - 1);
      range.values = data;

      const table = sheet.tables.add(range, true);
      table.name = "PeopleTable";

      table.getHeaderRowRange().format.fill.color = "lightgray";
      table.getRange().format.autofitColumns();
      table.getRange().format.autofitRows();

      await context.sync();
      */
    });
  } catch (error) {
    console.error(error);
  }
}
