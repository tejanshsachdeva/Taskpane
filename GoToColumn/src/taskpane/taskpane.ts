let isAscending = true; // Track ascending/descending sort state
let selectedColumnIndex: number = -1; // Store the selected column index
let originalOrder: string[] = []; // Store the original column order
let sortState = 0; // 0: default, 1: ascending, 2: descending

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("searchBox").addEventListener("input", filterColumns);
    document.getElementById("sortButton").addEventListener("click", sortColumns);
    document.getElementById("sheetDropdown").addEventListener("change", loadColumns);
    document.getElementById("toggleHideButton").addEventListener("click", toggleHideUnhide);
    document.getElementById("showProfileCheckbox").addEventListener("change", toggleProfileVisibility);
    // document.getElementById("refreshButton").addEventListener("click", refreshAddIn); // Add this line

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

    originalOrder = headers.map(header => `${header}`); // Store the original column order

    // Array to hold column objects for loading columnHidden property
    const columns = [];

    headers.forEach((header, index) => {
      const columnDiv = document.createElement("div");
      const columnLetter = getColumnLetter(index + 1); // Get column letter
      columnDiv.textContent = `${header} (${columnLetter})` || "<missing name>";
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



// Helper function to convert column index to letter
function getColumnLetter(index: number): string {
  let letter = '';
  while (index > 0) {
    const mod = (index - 1) % 26;
    letter = String.fromCharCode(65 + mod) + letter;
    index = Math.floor((index - mod) / 26);
  }
  return letter;
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

  if (sortState === 0) {
    // Reset to default order
    items.sort((a, b) => {
      const aName = a.textContent.split(" (")[0]; // Extract the column name
      const bName = b.textContent.split(" (")[0]; // Extract the column name
      return originalOrder.indexOf(aName) - originalOrder.indexOf(bName);
    });
    sortButton.textContent = "Sort (A-Z)";
    sortState = 1;
  } else if (sortState === 1) {
    // Sort in ascending order
    items.sort((a: HTMLElement, b: HTMLElement) => a.textContent.localeCompare(b.textContent, undefined, { caseFirst: 'lower' }));
    sortButton.textContent = "Sort (Z-A)";
    sortState = 2;
  } else {
    // Sort in descending order
    items.sort((a: HTMLElement, b: HTMLElement) => b.textContent.localeCompare(a.textContent, undefined, { caseFirst: 'lower' }));
    sortButton.textContent = "Reset to Default";
    sortState = 0;
  }

  columnList.innerHTML = "";
  items.forEach((item) => {
    columnList.appendChild(item);
  });
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

    column.load("values, columnHidden");
    await context.sync();

    // Update button text based on column's current state
    toggleButton.textContent = column.columnHidden ? "Unhide Column" : "Hide Column";

    column.select();
    await context.sync();

    // Calculate and display column profile if checkbox is checked
    const showProfile = (document.getElementById("showProfileCheckbox") as HTMLInputElement).checked;
    if (showProfile) {
      displayColumnProfile(column.values);
    } else {
      hideColumnProfile();
    }
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
function displayColumnProfile(values: any[][]) {
  const totalCount = values.length;
  const errorCount = values.filter(row => row[0] instanceof Error).length;
  const emptyCount = values.filter(row => row[0] === null || row[0] === '').length;
  const distinctValues = [...new Set(values.map(row => row[0]))];
  const distinctCount = distinctValues.length;
  const uniqueCount = distinctValues.filter(value => values.filter(row => row[0] === value).length === 1).length;
  const nanCount = values.filter(row => typeof row[0] === 'number' && isNaN(row[0])).length;

  let numericValues = values.filter(row => typeof row[0] === 'number' && !isNaN(row[0])).map(row => row[0]);
  let minValue, maxValue, averageValue;

  if (numericValues.length > 0) {
    minValue = Math.min(...numericValues);
    maxValue = Math.max(...numericValues);
    averageValue = numericValues.reduce((acc, val) => acc + val, 0) / numericValues.length;
  } else {
    minValue = maxValue = averageValue = "N/A";
  }

  // Update the UI
  document.getElementById("totalCount").textContent = totalCount.toString();
  document.getElementById("errorCount").textContent = errorCount.toString();
  document.getElementById("emptyCount").textContent = emptyCount.toString();
  document.getElementById("distinctCount").textContent = distinctCount.toString();
  document.getElementById("uniqueCount").textContent = uniqueCount.toString();
  document.getElementById("nanCount").textContent = nanCount.toString();
  document.getElementById("minValue").textContent = minValue.toString();
  document.getElementById("maxValue").textContent = maxValue.toString();
  document.getElementById("averageValue").textContent = averageValue.toString();

  // Show the column profile section
  document.getElementById("columnProfile").style.display = "block";
}

function hideColumnProfile() {
  // Hide the column profile section
  document.getElementById("columnProfile").style.display = "none";
}

function toggleProfileVisibility(event) {
  const checkbox = event.target as HTMLInputElement;
  if (!checkbox.checked) {
    hideColumnProfile();
  }
}