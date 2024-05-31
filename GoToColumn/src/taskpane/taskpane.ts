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
    if (!sheetName) {
      console.error("No sheet selected.");
      return;
    }

    const sheet = context.workbook.worksheets.getItem(sheetName);
    const range = sheet.getUsedRange();
    range.load("values, address");

    await context.sync();

    const headers = range.values[0];
    const columnList = document.getElementById("columnList");
    columnList.innerHTML = "";

    originalOrder = [...headers]; // Store the original column order

    const columns: Excel.Range[] = [];

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

    // Update button text based on the new state
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

function sortColumns() {
  const columnList = document.getElementById("columnList");
  const columns = Array.from(columnList.children);

  columns.sort((a, b) => {
    const nameA = a.textContent.toLowerCase();
    const nameB = b.textContent.toLowerCase();
    if (isAscending) {
      return nameA.localeCompare(nameB);
    } else {
      return nameB.localeCompare(nameA);
    }
  });

  isAscending = !isAscending; // Toggle the sort order

  columns.forEach((column) => columnList.appendChild(column));
}

function toggleProfileVisibility() {
  const showProfile = (document.getElementById("showProfileCheckbox") as HTMLInputElement).checked;
  if (showProfile) {
    displayColumnProfile();
  } else {
    hideColumnProfile();
  }
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

function displayColumnProfile(values: any) {
  // Implement logic to display column profile
}

function hideColumnProfile() {
  // Implement logic to hide column profile
}
