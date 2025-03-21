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
          displayMessage("Uploading PDF for conversion...");
          
          const sasToken = "sv=2024-11-04&ss=bfqt&srt=o&sp=rwdlacupiytfx&se=2025-04-20T11:40:31Z&st=2025-03-20T03:40:31Z&spr=https&sig=6o41aND51GX%2B0Q9r2ke49PHBV3CWlFDpIdQtqZUWX9w%3D";
          const accountName = "b2bmvpstorage";
          const containerName = "b2bvmp-container";

          const blobName = `${Date.now()}-${file.name}`;
          const blobUrl = `https://${accountName}.blob.core.windows.net/${containerName}/${encodeURIComponent(blobName)}`;

          try {
            const uploadResponse = await fetch(`${blobUrl}?${sasToken}`, {
              method: "PUT",
              headers: {
                "x-ms-blob-type": "BlockBlob",
                "Content-Type": file.type
              },
              body: file
            });
            
            if (uploadResponse.ok) {
              displayMessage("PDF uploaded successfully, converting to Excel...");

              const promptInput = document.getElementById("name") as HTMLInputElement;
              const prompt = promptInput && promptInput.value ? promptInput.value : "Extract all tables from the PDF and create an Excel spreadsheet";

              const apiResponse = await fetch("https://enterprise.factful.io/api/convert-to-excel", {
                method: "POST",
                headers: {
                  "Content-Type": "application/json"
                },
                body: JSON.stringify({
                  blob_url: blobUrl,
                  prompt: prompt
                })
              });

              if (apiResponse.ok) {
                const responseData = await apiResponse.json();
                
                if (responseData.data) {
                  await processTablesAndInsert(context, responseData.data);
                  displayMessage("PDF converted successfully and data inserted into Excel!");
                } else {
                  displayMessage("Failed to extract structured data from the PDF");
                }
              } else {
                displayMessage("Failed to convert PDF to Excel");
              }
            } else {
              displayMessage(`Upload failed: ${uploadResponse.status}`);
            }
          } catch (error) {
            displayMessage(`Error: ${error.message}`);
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
    displayMessage(`Error: ${error.message}`);
  }
}

async function processTablesAndInsert(context: Excel.RequestContext, data: any) {
  try {

    const sheet = context.workbook.worksheets.getActiveWorksheet();
    
    const formattedData = Array.isArray(data) ? data : [];
    
    if (formattedData.length === 0) {
      displayMessage("No data found in the API response");
      return;
    }

    const range = sheet.getRange("A1").getResizedRange(formattedData.length - 1, formattedData[0].length - 1);
    range.values = formattedData;
    
    const table = sheet.tables.add(range, true);
    table.name = "DataTable";
    
    table.getHeaderRowRange().format.fill.color = "#4472C4";
    table.getHeaderRowRange().format.font.color = "white";
    table.getHeaderRowRange().format.font.bold = true;
    
    range.format.autofitColumns();
    range.format.autofitRows();

    await tryCreateChart(sheet, formattedData);
    
    displayMessage("Data successfully inserted into Excel!");
    
    await context.sync();
  } catch (error) {
    console.error("Error processing data:", error);
    displayMessage(`Error processing data: ${error.message}`);
  }
}

async function tryCreateChart(sheet: Excel.Worksheet, data: any[][]) {
  try {
    if (data.length <= 1) {
      return; 
    }

    const numericColumns = [];
    
    for (let colIndex = 0; colIndex < data[0].length; colIndex++) {
      let isNumeric = true;
      let numericCount = 0;

      for (let rowIndex = 1; rowIndex < data.length; rowIndex++) {
        const cellValue = data[rowIndex][colIndex];

        if (cellValue !== "" && !isNaN(Number(cellValue))) {
          numericCount++;
        }
      }
      isNumeric = numericCount >= (data.length - 1) * 0.7;
      
      if (isNumeric) {
        numericColumns.push(colIndex);
      }
    }
  
    if (numericColumns.length > 0) {
      let categoryColumn = -1;
      for (let colIndex = 0; colIndex < data[0].length; colIndex++) {
        if (!numericColumns.includes(colIndex)) {
          categoryColumn = colIndex;
          break;
        }
      }
      
      if (categoryColumn === -1) {
        categoryColumn = 0;
      }

      const chartBodyRange = sheet.getRange("A2").getResizedRange(data.length - 2, data[0].length - 1);

      let chartType: Excel.ChartType;
      if (numericColumns.length === 1) {
        chartType = Excel.ChartType.columnClustered;
      } else {
        chartType = Excel.ChartType.line;
      }
      
      const chart = sheet.charts.add(chartType, chartBodyRange, Excel.ChartSeriesBy.auto);

      chart.setPosition("A" + (data.length + 2), "H" + (data.length + 20));
      
      chart.title.text = "Data Visualization";
      chart.legend.position = Excel.ChartLegendPosition.right;
      chart.axes.categoryAxis.title.text = data[0][categoryColumn];
    }
  } catch (error) {
    console.error("Error creating chart:", error);
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

