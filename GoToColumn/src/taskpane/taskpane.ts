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
    document.getElementById("toggleLockButton").addEventListener("click", toggleLockSheet); // Always show the lock sheet button
    document.getElementById("showProfileCheckbox").addEventListener("change", toggleProfileVisibility);
    Excel.run(async (context) => {
      context.workbook.worksheets.onAdded.add(loadSheets);
      context.workbook.worksheets.onDeleted.add(loadSheets);
      context.workbook.worksheets.onChanged.add(loadSheets);

      await context.sync();
    }).catch(error => {
      console.error(error);
    });

    loadSheets();
  }
});

async function loadSheets() {
  await Excel.run(async (context) => {
    const sheets = context.workbook.worksheets;
    sheets.load("items/name");
    await context.sync();

    const sheetDropdown = document.getElementById("sheetDropdown") as HTMLSelectElement;
    const currentSheetName = sheetDropdown.value; // Store the currently selected sheet name

    sheetDropdown.innerHTML = "";

    sheets.items.forEach((sheet) => {
      const option = document.createElement("option");
      option.text = sheet.name;
      option.value = sheet.name;
      sheetDropdown.appendChild(option);

      // Register the onChanged event handler for each sheet
      sheet.onChanged.add(loadColumns);
    });

    // Restore the previously selected sheet, or select the first sheet if none was previously selected
    if (currentSheetName && sheets.items.some(sheet => sheet.name === currentSheetName)) {
      sheetDropdown.value = currentSheetName;
    } else {
      sheetDropdown.selectedIndex = 0;
    }

    // Load columns for the currently selected sheet
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
      columnDiv.textContent = `${header || '<missing name>'} (${columnLetter}-${index + 1})`; // Handle missing header
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

    // Update lock button text based on the sheet's protection state
    const sheetProtection = sheet.protection;
    sheetProtection.load("protected");
    await context.sync();

    const toggleLockButton = document.getElementById("toggleLockButton");
    toggleLockButton.textContent = sheetProtection.protected ? "Unlock Sheet" : "Lock Sheet";
  }).catch(error => {
    console.error(error);
  });
}


async function selectColumn(index: number) {
  selectedColumnIndex = index; // Update the selected column index

  // Show the toggle button for hide/unhide
  const toggleHideButton = document.getElementById("toggleHideButton");
  toggleHideButton.style.display = "block";

  await Excel.run(async (context) => {
    const sheetName = (document.getElementById("sheetDropdown") as HTMLSelectElement).value;
    if (!sheetName) {
      console.error("No sheet selected.");
      return;
    }

    const sheet = context.workbook.worksheets.getItem(sheetName);
    const range = sheet.getUsedRange();
    range.load("values, address");

    await context.sync();

    const column = range.getColumn(index);
    column.load("values, columnHidden");

    await context.sync();

    // Update button text based on column's current state
    toggleHideButton.textContent = column.columnHidden ? "Unhide Column" : "Hide Column";

    column.select();
    await context.sync();

    // Calculate and display column profile if checkbox is checked
    const showProfile = (document.getElementById("showProfileCheckbox") as HTMLInputElement).checked;
    if (showProfile) {
      displayColumnProfile(column.values);
    } else {
      hideColumnProfile();
    }
  }).catch(error => {
    console.error(error);
  });
}

async function toggleHideUnhide() {
  await Excel.run(async (context) => {
    const sheetName = (document.getElementById("sheetDropdown") as HTMLSelectElement).value;
    if (!sheetName) {
      console.error("No sheet selected.");
      return;
    }

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
  }).catch(error => {
    console.error(error);
  });
}

async function toggleLockSheet() {
  await Excel.run(async (context) => {
    const sheetName = (document.getElementById("sheetDropdown") as HTMLSelectElement).value;
    if (!sheetName) {
      console.error("No sheet selected.");
      return;
    }

    const sheet = context.workbook.worksheets.getItem(sheetName);

    // Load the sheet protection state
    const sheetProtection = sheet.protection;
    sheetProtection.load("protected");
    await context.sync();

    // Toggle the protection state of the sheet
    if (sheetProtection.protected) {
      await sheet.protection.unprotect();
    } else {
      await sheet.protection.protect();
    }

    const toggleButton = document.getElementById("toggleLockButton");
    if (sheetProtection.protected) {
      toggleButton.textContent = "Lock Sheet";
    } else {
      toggleButton.textContent = "Unlock Sheet";
    }
  }).catch(error => {
    console.error(error);
  });
}

function filterColumns() {
  const searchTerm = (document.getElementById("searchBox") as HTMLInputElement).value.toLowerCase();
  const columnItems = document.getElementsByClassName("column-item");

  for (let i = 0; i < columnItems.length; i++) {
    const columnItem = columnItems[i] as HTMLElement;
    const columnName = columnItem.textContent.toLowerCase();
    if (columnName.includes(searchTerm)) {
      columnItem.style.display = "block";
    } else {
      columnItem.style.display = "none";
    }
  }
}

//c1
function sortColumns() {
  const columnList = document.getElementById("columnList");
  const items = Array.from(columnList.getElementsByClassName("column-item"));
  const sortButton = document.getElementById("sortButton");

  if (sortState === 0) {
    // Reset to default order based on the number in parentheses (second last character)
    items.sort((a, b) => {
      const aText = a.textContent;
      const bText = b.textContent;

      // Extract the number in parentheses (format: 'name (X-Y)')
      const aNumber = parseInt(aText.match(/\(([^-]+)-(\d+)\)/)[2], 10);
      const bNumber = parseInt(bText.match(/\(([^-]+)-(\d+)\)/)[2], 10);

      return aNumber - bNumber;
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



function getColumnLetter(columnNumber: number): string {
  let columnLetter = "";
  while (columnNumber > 0) {
    const modulo = (columnNumber - 1) % 26;
    columnLetter = String.fromCharCode(65 + modulo) + columnLetter;
    columnNumber = Math.floor((columnNumber - modulo) / 26);
  }
  return columnLetter;
}

function displayColumnProfile(values: any[][]) {
  const dataRows = values.slice(1);
  const totalCount = dataRows.length;

  let errorCount = 0;
  let emptyCount = 0;
  let uniqueCount = 0;
  let nanCount = 0;
  let numericValues: number[] = [];
  let dateValues: Date[] = [];
  let distinctMap = new Map();

  dataRows.forEach(row => {
    const value = row[0];
    if (value === null || value === '') {
      emptyCount++;
    } else if (isErrorValue(value)) {
      errorCount++;
    } else if (typeof value === 'number') {
      if (isNaN(value)) {
        nanCount++;
      } else {
        numericValues.push(value);
      }
    } else if (typeof value === 'string' && isValidDate(value)) {
      dateValues.push(parseDateString(value));
    }

    distinctMap.set(value, (distinctMap.get(value) || 0) + 1);
  });

  distinctMap.forEach((count, value) => {
    if (count === 1) uniqueCount++;
  });

  const distinctCount = distinctMap.size;

  let minValue: string | number, maxValue: string | number, sumValue: number | string = "N/A", averageValue: number | string = "N/A";

  if (numericValues.length > 0) {
    minValue = Math.min(...numericValues).toFixed(2);
    maxValue = Math.max(...numericValues).toFixed(2);
    sumValue = numericValues.reduce((acc, val) => acc + val, 0);
    averageValue = (sumValue / numericValues.length).toFixed(2);
    sumValue = sumValue.toFixed(2);
  } else if (dateValues.length > 0) { 
    minValue = new Date(Math.min(...dateValues.map(date => date.getTime()))).toLocaleDateString();
    maxValue = new Date(Math.max(...dateValues.map(date => date.getTime()))).toLocaleDateString();
  } else {
    minValue = maxValue = "N/A";
  }

  document.getElementById("totalCount").textContent = totalCount.toString();
  document.getElementById("errorCount").textContent = errorCount.toString();
  document.getElementById("emptyCount").textContent = emptyCount.toString();
  document.getElementById("distinctCount").textContent = distinctCount.toString();
  document.getElementById("uniqueCount").textContent = uniqueCount.toString();
  document.getElementById("nanCount").textContent = nanCount.toString();
  document.getElementById("minValue").textContent = minValue.toString();
  document.getElementById("maxValue").textContent = maxValue.toString();
  document.getElementById("averageValue").textContent = averageValue.toString();
  document.getElementById("sumValue").textContent = sumValue.toString();

  document.getElementById("columnProfile").style.display = "block";
}





function hideColumnProfile() {
  // Hide the column profile section
  document.getElementById("columnProfile").style.display = "none";
}

function isValidDate(dateString: string): boolean {
  return !isNaN(Date.parse(dateString));
}
function toggleProfileVisibility(event) {
  const checkbox = event.target as HTMLInputElement;
  if (!checkbox.checked) {
    hideColumnProfile();
  }
}

function parseDateString(dateString: string): Date {
  return new Date(dateString);
}

function isErrorValue(value: any): boolean {
  return value instanceof Error || (typeof value === 'string' && value.startsWith("#"));
}
