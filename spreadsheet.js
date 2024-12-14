// Define default values for rows and columns
let numRows = 25; // Set initial number of rows
let numCols = 25; // Set initial number of columns

const table = document.getElementById("spreadsheet");
const contextMenu = document.getElementById('contextMenu');
let selectedRow = null;
let selectedCol = null;

// Function to get the Excel-style column name (A, B, C, ..., Z, AA, AB, ..., AZ, BA, ...)
function getColumnName(colIndex) {
  let columnName = '';
  while (colIndex > 0) {
    const remainder = (colIndex - 1) % 26;  // Get remainder (0 = A, 1 = B, ..., 25 = Z)
    let letter = String.fromCharCode(65 + remainder);  // Directly get the character
    columnName = letter + columnName;  // Prepend the character to the column name
    colIndex = Math.floor((colIndex - 1) / 26);  // Update colIndex for the next significant letter
  }
  return columnName;
}

// Modify the table creation to use Excel-style column naming
function createTable() {
  table.innerHTML = ""; // Clear any existing table content

  // Create column headers (A, B, C, D, ..., Z, AA, AB, ...)
  const headerRow = document.createElement("tr");
  headerRow.appendChild(document.createElement("th")); // Top-left corner empty cell
  for (let col = 1; col <= numCols; col++) {
    const th = document.createElement("th");
    th.textContent = getColumnName(col); // Use Excel-style column names
    headerRow.appendChild(th);
  }
  table.appendChild(headerRow);

  // Generate the table rows with row numbers (1, 2, 3, 4, ...)
  for (let row = 0; row < numRows; row++) {
    const tr = document.createElement("tr");

    // Row header (1, 2, 3, 4, ...)
    const rowHeader = document.createElement("th");
    rowHeader.textContent = row + 1; // Row numbers as digits (1, 2, 3, ...)
    tr.appendChild(rowHeader);

    // Create data cells for each column
    for (let col = 0; col < numCols; col++) {
      const td = document.createElement("td");

      // Create input for cell content
      const input = document.createElement("input");
      input.type = "text";
      td.appendChild(input);

      // Create resizer for columns
      const resizerCol = document.createElement("div");
      resizerCol.className = "resizer";
      td.appendChild(resizerCol);

      // Create resizer for rows
      const resizerRow = document.createElement("div");
      resizerRow.className = "resizer-row";
      td.appendChild(resizerRow);

      // Append cell to row
      tr.appendChild(td);

      // Handle column resizing
      resizerCol.addEventListener("mousedown", function (e) {
        e.preventDefault();
        const startX = e.clientX;
        const startWidth = td.offsetWidth;

        function onMouseMove(e) {
          const newWidth = startWidth + (e.clientX - startX);
          td.style.width = `${newWidth}px`;
        }

        function onMouseUp() {
          document.removeEventListener("mousemove", onMouseMove);
          document.removeEventListener("mouseup", onMouseUp);
        }

        document.addEventListener("mousemove", onMouseMove);
        document.addEventListener("mouseup", onMouseUp);
      });

      // Handle row resizing
      resizerRow.addEventListener("mousedown", function (e) {
        e.preventDefault();
        const startY = e.clientY;
        const startHeight = td.offsetHeight;

        function onMouseMove(e) {
          const newHeight = startHeight + (e.clientY - startY);
          td.style.height = `${newHeight}px`;
        }

        function onMouseUp() {
          document.removeEventListener("mousemove", onMouseMove);
          document.removeEventListener("mouseup", onMouseUp);
        }

        document.addEventListener("mousemove", onMouseMove);
        document.addEventListener("mouseup", onMouseUp);
      });

      // Add event listener for cell click to highlight row and column
      td.addEventListener("click", function () {
        // Remove highlights from previous selected row and column
        document.querySelectorAll(".highlight-row").forEach(row => row.classList.remove("highlight-row"));
        document.querySelectorAll(".highlight-col").forEach(col => col.classList.remove("highlight-col"));

        // Highlight the corresponding row and column
        highlightRowAndColumn(row, col);
      });
    }
    table.appendChild(tr);
  }
}

// Function to highlight the row and column
function highlightRowAndColumn(row, col) {
  // Highlight the selected row
  const rowHeader = table.querySelectorAll("tr")[row + 1].querySelector("th");
  rowHeader.classList.add("highlight-row");

  // Highlight the selected column
  const headerCells = table.querySelectorAll("tr:first-child th");
  headerCells[col + 1].classList.add("highlight-col");
}

// Initial table creation
createTable();

// Show context menu
function showContextMenu(event) {
  const cell = event.target;

  // Determine whether the clicked element is a row header or column header
  if (cell.tagName === 'TH') {
    // Check if the clicked element is a row header (first column)
    if (cell.parentElement === table.querySelector("tr:first-child")) {
      // Right-clicked on the column header, show Add Column
      document.getElementById("addRowBtn").style.display = 'none'; // Hide Add Row
      document.getElementById("addColBtn").style.display = 'block'; // Show Add Column
    } else {
      // Right-clicked on the row header, show Add Row
      document.getElementById("addRowBtn").style.display = 'block'; // Show Add Row
      document.getElementById("addColBtn").style.display = 'none'; // Hide Add Column
    }

    // Show context menu
    contextMenu.style.display = 'block';
    contextMenu.style.left = `${event.pageX}px`;
    contextMenu.style.top = `${event.pageY}px`;

    // Get the row and column index
    if (cell.parentElement === table.querySelector("tr:first-child")) {
      // Column header clicked
      const colIndex = Array.from(cell.parentElement.children).indexOf(cell);
      selectedRow = null; // Reset selected row
      selectedCol = colIndex;
    } else {
      // Row header clicked
      const rowIndex = Array.from(cell.parentElement.children).indexOf(cell);
      selectedRow = rowIndex;
      selectedCol = null; // Reset selected column
    }
  } else {
    // If the right-click is on a table cell, reset the context menu
    contextMenu.style.display = 'none';
  }
}

// Event listener to show context menu on right click
table.addEventListener('contextmenu', function (event) {
  event.preventDefault();
  showContextMenu(event);
});

// Hide context menu on click anywhere
document.addEventListener('click', function () {
  contextMenu.style.display = 'none';
});

// Function to add a new row
document.getElementById("addRowBtn").addEventListener("click", function () {
  numRows++;  // Increase the number of rows
  const tr = document.createElement("tr");

  // Create new row header (row number)
  const rowHeader = document.createElement("th");
  rowHeader.textContent = numRows; // Update row number dynamically
  tr.appendChild(rowHeader);

  // Add new cells for each column
  for (let col = 0; col < numCols; col++) {
    const td = document.createElement("td");
    const input = document.createElement("input");
    input.type = "text";
    td.appendChild(input);

    // Add resizer divs
    const resizerCol = document.createElement("div");
    resizerCol.className = "resizer";
    td.appendChild(resizerCol);
    const resizerRow = document.createElement("div");
    resizerRow.className = "resizer-row";
    td.appendChild(resizerRow);

    tr.appendChild(td);

    // Re-apply event listeners for resizing
    resizerCol.addEventListener("mousedown", function (e) {
      e.preventDefault();
      const startX = e.clientX;
      const startWidth = td.offsetWidth;

      function onMouseMove(e) {
        const newWidth = startWidth + (e.clientX - startX);
        td.style.width = `${newWidth}px`;
      }

      function onMouseUp() {
        document.removeEventListener("mousemove", onMouseMove);
        document.removeEventListener("mouseup", onMouseUp);
      }

      document.addEventListener("mousemove", onMouseMove);
      document.addEventListener("mouseup", onMouseUp);
    });

    resizerRow.addEventListener("mousedown", function (e) {
      e.preventDefault();
      const startY = e.clientY;
      const startHeight = td.offsetHeight;

      function onMouseMove(e) {
        const newHeight = startHeight + (e.clientY - startY);
        td.style.height = `${newHeight}px`;
      }

      function onMouseUp() {
        document.removeEventListener("mousemove", onMouseMove);
        document.removeEventListener("mouseup", onMouseUp);
      }

      document.addEventListener("mousemove", onMouseMove);
      document.addEventListener("mouseup", onMouseUp);
    });
  }

  table.appendChild(tr);
  document.getElementById("contextMenu").style.display = 'none';
});

// Function to add a new column
document.getElementById("addColBtn").addEventListener("click", function () {
  numCols++;  // Increase the number of columns

  // Update the column headers
  const headerRow = table.querySelector("tr:first-child");
  const th = document.createElement("th");
  th.textContent = getColumnName(numCols); // Add new column header dynamically
  headerRow.appendChild(th);

  // Add a new cell in each existing row
  for (let row = 1; row <= numRows; row++) {
    const tr = table.querySelectorAll("tr")[row];
    const td = document.createElement("td");
    const input = document.createElement("input");
    input.type = "text";
    td.appendChild(input);

    // Add resizer divs
    const resizerCol = document.createElement("div");
    resizerCol.className = "resizer";
    td.appendChild(resizerCol);
    const resizerRow = document.createElement("div");
    resizerRow.className = "resizer-row";
    td.appendChild(resizerRow);

    tr.appendChild(td);

    // Re-apply event listeners for resizing
    resizerCol.addEventListener("mousedown", function (e) {
      e.preventDefault();
      const startX = e.clientX;
      const startWidth = td.offsetWidth;

      function onMouseMove(e) {
        const newWidth = startWidth + (e.clientX - startX);
        td.style.width = `${newWidth}px`;
      }

      function onMouseUp() {
        document.removeEventListener("mousemove", onMouseMove);
        document.removeEventListener("mouseup", onMouseUp);
      }

      document.addEventListener("mousemove", onMouseMove);
      document.addEventListener("mouseup", onMouseUp);
    });

    resizerRow.addEventListener("mousedown", function (e) {
      e.preventDefault();
      const startY = e.clientY;
      const startHeight = td.offsetHeight;

      function onMouseMove(e) {
        const newHeight = startHeight + (e.clientY - startY);
        td.style.height = `${newHeight}px`;
      }

      function onMouseUp() {
        document.removeEventListener("mousemove", onMouseMove);
        document.removeEventListener("mouseup", onMouseUp);
      }

      document.addEventListener("mousemove", onMouseMove);
      document.addEventListener("mouseup", onMouseUp);
    });
  }
  document.getElementById("contextMenu").style.display = 'none';
});

