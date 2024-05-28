let isAscending = true; // Add this line at the top of your script

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("searchBox").addEventListener("input", filterColumns);
    document.getElementById("sortButton").addEventListener("click", sortColumns);
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
    return isAscending ? a.textContent.localeCompare(b.textContent) : b.textContent.localeCompare(a.textContent);
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
  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const range = sheet.getUsedRange();
    const column = range.getColumn(index);

    column.select();
    await context.sync();
  });
}
