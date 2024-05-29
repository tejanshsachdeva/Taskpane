let isAscending = true; // Ensure this line is at the top level of your script

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("searchBox").addEventListener("input", filterColumns);
    document.getElementById("sortButton").addEventListener("click", sortColumns);
    document.getElementById("sheetDropdown").addEventListener("change", loadColumns);

    // Attach event listeners for hide/unhide buttons
    document.getElementById("hideButton").addEventListener("click", hideColumn);
    document.getElementById("unhideButton").addEventListener("click", unhideColumn);

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

    headers.forEach((header, index) => {
      const columnDiv = document.createElement("div");
      columnDiv.textContent = header || "<missing name>";
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

let selectedColumnIndex: number = -1; // Store the selected column index

async function selectColumn(index: number) {
  selectedColumnIndex = index; // Update the selected column index

  document.getElementById("hideButton").style.display = "block";
  document.getElementById("unhideButton").style.display = "block";

  await Excel.run(async (context) => {
    const sheetName = (document.getElementById("sheetDropdown") as HTMLSelectElement).value;
    const sheet = context.workbook.worksheets.getItem(sheetName);
    const range = sheet.getUsedRange();
    const column = range.getColumn(index);

    column.select();
    await context.sync();
  });
}


async function hideColumn() {
  await Excel.run(async (context) => {
    const sheetName = (document.getElementById("sheetDropdown") as HTMLSelectElement).value;
    const sheet = context.workbook.worksheets.getItem(sheetName);
    const range = sheet.getUsedRange();
    const column = range.getColumn(selectedColumnIndex); // Use the selected column index

    column.columnHidden = true;
    await context.sync();
  });
}

async function unhideColumn() {
  await Excel.run(async (context) => {
    const sheetName = (document.getElementById("sheetDropdown") as HTMLSelectElement).value;
    const sheet = context.workbook.worksheets.getItem(sheetName);
    const range = sheet.getUsedRange();
    const column = range.getColumn(selectedColumnIndex); // Use the selected column index

    column.columnHidden = false;
    await context.sync();
  });
}
