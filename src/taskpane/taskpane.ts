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
    document.getElementById("convert").onclick = convertToExcel;
  }
});

export async function webscrape() {
  try {
    await Excel.run(async (context) => {
      /*
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      
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

export async function convertToExcel() {
  try {
    await Excel.run(async (context) => {
      const fileInput = document.getElementById("file-input") as HTMLInputElement;

      if (fileInput && fileInput.files.length > 0) {
        const file = fileInput.files[0];

        if (file.type === "application/pdf") {
          const formData = new FormData();
          formData.append("file", file);

          const blobUrl = URL.createObjectURL(file);

          displayMessage(blobUrl);

          setTimeout(() => {
            URL.revokeObjectURL(blobUrl);
          }, 60 * 1000);

          const response = await fetch("https://enterprise.factful.io/api/convert-to-excel", {
            method: "POST",
            body: formData
          });

          if (response.ok) {
            const responseData = await response.json();

          } else {
            displayMessage("Failed to convert the file");
          }
        } else {
          displayMessage("Please select a PDF file.");
        }
      } else {
        displayMessage("No file selected");
      }

      await context.sync();
    });
  } catch (error) {
    console.error(error);
  }
}

export async function displayMessage(message: string) {
  try {
    await Excel.run(async (context) => {
      const messageDisplay = document.getElementById("message");

      messageDisplay.innerText = message;

      await context.sync();
    });
  } catch (error) {
    console.error(error);
  } 
}