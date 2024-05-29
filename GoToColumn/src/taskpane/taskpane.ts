let isAscending = true; // Ensure this line is at the top level of your script
let selectedColumnIndex: number = -1; // Store the selected column index

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("searchBox").addEventListener("input", filterColumns);
    document.getElementById("sortButton").addEventListener("click", sortColumns);
    document.getElementById("sheetDropdown").addEventListener("change", loadColumns);

    // Attach event listener for the hide/unhide button
    document.getElementById("toggleHideButton").addEventListener("click", toggleHideUnhide);

    loadSheets();
  }
});

async function loadSheets() {
  await Excel.run(async (context) => {
    const sheets = context.workbook.worksheets;
    sheets.load("items/name");
    await context.sync();

    const sheetDropdown = document.getElementById("sheetDropdown");
    sheetDropdown.innerHTML = "";

    sheets.items.forEach((sheet) => {
      const option = document.createElement("option");
      option.text = sheet.name;
      option.value = sheet.name;
      sheetDropdown.appendChild(option);
    });

    // Load columns for the initially selected sheet
    loadColumns();
  });
}

async function loadColumns() {
  await Excel.run(async (context) => {
      const sheetName = (document.getElementById("sheetDropdown") as HTMLSelectElement).value;
      const sheet = context.workbook.worksheets.getItem(sheetName);
      const range = sheet.getUsedRange();
      range.load("values, address");

      await context.sync();

      const headers = range.values[0];
      const columnList = document.getElementById("columnList");
      columnList.innerHTML = "";

      // Array to hold column objects for loading columnHidden property
      const columns = [];

      headers.forEach((header, index) => {
          const columnDiv = document.createElement("div");
          columnDiv.textContent = header || "<missing name>";
          columnDiv.classList.add("column-item");

          // Load column hidden state
          const column = range.getColumn(index);
          column.load("columnHidden");
          columns.push(column);

          columnDiv.addEventListener("click", () => selectColumn(index));
          columnList.appendChild(columnDiv);
      });

      await context.sync();

      // Apply the hidden-column class based on the hidden state
      columns.forEach((column, index) => {
          const columnDiv = columnList.children[index] as HTMLElement;
          if (column.columnHidden) {
              columnDiv.classList.add("hidden-column");
          }
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

function sortColumns() {
  const columnList = document.getElementById("columnList");
  const items = Array.from(columnList.getElementsByClassName("column-item"));
  const sortButton = document.getElementById("sortButton");

  items.sort((a: HTMLElement, b: HTMLElement) => {
    return isAscending
      ? a.textContent.localeCompare(b.textContent, undefined, { caseFirst: 'lower' })
      : b.textContent.localeCompare(a.textContent, undefined, { caseFirst: 'lower' });
  });

  columnList.innerHTML = "";
  items.forEach((item) => {
    columnList.appendChild(item);
  });

  // Update the button text
  sortButton.textContent = isAscending ? 'Sort (Z-A)' : 'Sort (A-Z)';

  isAscending = !isAscending; // Toggle the order for the next click
}

async function selectColumn(index: number) {
  selectedColumnIndex = index; // Update the selected column index

  // Show the toggle button
  const toggleButton = document.getElementById("toggleHideButton");
  toggleButton.style.display = "block";

  await Excel.run(async (context) => {
      const sheetName = (document.getElementById("sheetDropdown") as HTMLSelectElement).value;
      const sheet = context.workbook.worksheets.getItem(sheetName);
      const range = sheet.getUsedRange();
      const column = range.getColumn(index);

      column.load("columnHidden");
      await context.sync();

      // Update button text based on column's current state
      toggleButton.textContent = column.columnHidden ? "Unhide Column" : "Hide Column";

      column.select();
      await context.sync();
  });
}

async function toggleHideUnhide() {
  await Excel.run(async (context) => {
      const sheetName = (document.getElementById("sheetDropdown") as HTMLSelectElement).value;
      const sheet = context.workbook.worksheets.getItem(sheetName);
      const range = sheet.getUsedRange();
      const column = range.getColumn(selectedColumnIndex); // Use the selected column index

      column.load("columnHidden");
      await context.sync();

      // Toggle the hidden state of the column
      column.columnHidden = !column.columnHidden;
      await context.sync();

      // Update button text based on the new state
      const toggleButton = document.getElementById("toggleHideButton");
      toggleButton.textContent = column.columnHidden ? "Unhide Column" : "Hide Column";

      // Update the CSS class of the column item based on the new state
      const columnDiv = document.getElementById("columnList").children[selectedColumnIndex] as HTMLElement;
      if (column.columnHidden) {
          columnDiv.classList.add("hidden-column");
      } else {
          columnDiv.classList.remove("hidden-column");
      }
  });
}

