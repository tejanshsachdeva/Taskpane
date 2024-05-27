/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("searchBox").addEventListener("input", filterColumns);
    loadColumns();
  }
});

async function loadColumns() {
  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const range = sheet.getUsedRange();
    range.load("values, address");

    await context.sync();

    const headers = range.values[0];
    const columnList = document.getElementById("columnList");
    columnList.innerHTML = "";

    headers.forEach((header, index) => {
      const columnDiv = document.createElement("div");
      columnDiv.textContent = header;
      columnDiv.classList.add("column-item");
      columnDiv.addEventListener("click", () => selectColumn(index));
      columnList.appendChild(columnDiv);
    });
  });
}




function filterColumns(event) {
  const query = event.target.value.toLowerCase();
  const items = document.getElementsByClassName("column-item");
  
  Array.from(items).forEach((item: HTMLElement) => {
    if (item.textContent.toLowerCase().includes(query)) {
      item.style.display = "block";
    } else {
      item.style.display = "none";
    }
  });
}
async function selectColumn(index: number) {
  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const range = sheet.getUsedRange();
    const column = range.getColumn(index);

    column.select();
    await context.sync();
  });
}
