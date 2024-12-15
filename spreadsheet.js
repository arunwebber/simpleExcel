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
// Function to create a table header row
function createHeaderRow() {
  const headerRow = document.createElement("tr");
  headerRow.appendChild(document.createElement("th")); // Top-left corner empty cell
  
  for (let col = 1; col <= numCols; col++) {
    const th = document.createElement("th");
    th.textContent = getColumnName(col); // Excel-style column names
    headerRow.appendChild(th);
  }
  return headerRow;
}

// Function to create a data cell with input and resizers
function createDataCell() {
  const td = document.createElement("td");

  // Input field for cell content
  const input = document.createElement("input");
  input.type = "text";
  td.appendChild(input);

  // Resizers for columns and rows
  td.appendChild(createResizer("resizer", "width", td));  // For column resizing
  td.appendChild(createResizer("resizer-row", "height", td)); // For row resizing

  return td;
}

// Function to create a resizer for column or row
function createResizer(className, dimension, td) {
  const resizer = document.createElement("div");
  resizer.className = className;

  resizer.addEventListener("mousedown", function (e) {
    e.preventDefault();
    const start = dimension === "width" ? e.clientX : e.clientY;
    const startSize = dimension === "width" ? td.offsetWidth : td.offsetHeight;

    function onMouseMove(e) {
      const newSize = dimension === "width"
        ? startSize + (e.clientX - start)
        : startSize + (e.clientY - start);

      // Apply the new size to the corresponding dimension (width or height)
      if (dimension === "width") {
        td.style.width = `${newSize}px`;
      } else {
        td.style.height = `${newSize}px`;
      }
    }

    function onMouseUp() {
      document.removeEventListener("mousemove", onMouseMove);
      document.removeEventListener("mouseup", onMouseUp);
    }

    document.addEventListener("mousemove", onMouseMove);
    document.addEventListener("mouseup", onMouseUp);
  });

  return resizer;
}

// Function to create a table row with data cells
function createTableRow(row) {
  const tr = document.createElement("tr");

  // Row header (row number)
  const rowHeader = document.createElement("th");
  rowHeader.textContent = row + 1; // Row number
  tr.appendChild(rowHeader);

  // Create data cells for each column
  for (let col = 0; col < numCols; col++) {
    const td = createDataCell();

    // Add event listener to highlight row and column
    td.addEventListener("click", function () {
      document.querySelectorAll(".highlight-row").forEach(row => row.classList.remove("highlight-row"));
      document.querySelectorAll(".highlight-col").forEach(col => col.classList.remove("highlight-col"));
      highlightRowAndColumn(row, col);
    });

    tr.appendChild(td);
  }

  return tr;
}

// Main function to create the table
function createTable() {
  table.innerHTML = ""; // Clear existing content

  // Create and append header row
  const headerRow = createHeaderRow();
  table.appendChild(headerRow);

  // Create and append table body rows
  for (let row = 0; row < numRows; row++) {
    const tableRow = createTableRow(row);
    table.appendChild(tableRow);
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

  if (cell.tagName === 'TH') {
    // Show Add Row or Add Column based on the clicked header
    const isColumnHeader = cell.parentElement === table.querySelector("tr:first-child");
    document.getElementById("addRowBtn").style.display = isColumnHeader ? 'none' : 'block';
    document.getElementById("addColBtn").style.display = isColumnHeader ? 'block' : 'none';

    // Show context menu
    contextMenu.style.display = 'block';
    contextMenu.style.left = `${event.pageX}px`;
    contextMenu.style.top = `${event.pageY}px`;

    // Set the selected row/column
    if (isColumnHeader) {
      selectedRow = null;
      selectedCol = Array.from(cell.parentElement.children).indexOf(cell);
    } else {
      selectedRow = Array.from(cell.parentElement.children).indexOf(cell);
      selectedCol = null;
    }
  } else {
    contextMenu.style.display = 'none'; // Hide context menu if right-clicked on a table cell
  }
}

// Event listener for context menu on right-click
table.addEventListener('contextmenu', (event) => {
  event.preventDefault();
  showContextMenu(event);
});

// Hide context menu on click anywhere
document.addEventListener('click', () => contextMenu.style.display = 'none');

// Function to handle adding a row or column
function addElement(isRow) {
  const tr = document.createElement("tr");
  
  // Add header or cells
  if (isRow) {
    const rowHeader = document.createElement("th");
    rowHeader.textContent = numRows + 1;
    tr.appendChild(rowHeader);

    for (let col = 0; col < numCols; col++) {
      tr.appendChild(createCell());
    }
    table.appendChild(tr);
    numRows++;
  } else {
    const headerRow = table.querySelector("tr:first-child");
    const th = document.createElement("th");
    th.textContent = getColumnName(numCols + 1);
    headerRow.appendChild(th);

    for (let row = 1; row <= numRows; row++) {
      const td = createCell();
      table.querySelectorAll("tr")[row].appendChild(td);
    }
    numCols++;
  }

  document.getElementById("contextMenu").style.display = 'none';
}

// Create a new cell with input and resizer functionality
function createCell() {
  const td = document.createElement("td");
  const input = document.createElement("input");
  input.type = "text";
  td.appendChild(input);

  // Add resizer divs
  ['resizer', 'resizer-row'].forEach(className => {
    const resizer = document.createElement("div");
    resizer.className = className;
    td.appendChild(resizer);

    resizer.addEventListener("mousedown", (e) => resizeCell(e, td, className));
  });

  return td;
}

// Handle resizing logic
function resizeCell(e, td, direction) {
  e.preventDefault();
  const startPos = direction === 'resizer' ? e.clientX : e.clientY;
  const startSize = direction === 'resizer' ? td.offsetWidth : td.offsetHeight;

  const onMouseMove = (e) => {
    const newSize = direction === 'resizer' ? startSize + (e.clientX - startPos) : startSize + (e.clientY - startPos);
    direction === 'resizer' ? td.style.width = `${newSize}px` : td.style.height = `${newSize}px`;
  };

  const onMouseUp = () => {
    document.removeEventListener("mousemove", onMouseMove);
    document.removeEventListener("mouseup", onMouseUp);
  };

  document.addEventListener("mousemove", onMouseMove);
  document.addEventListener("mouseup", onMouseUp);
}

// Add row and column button event listeners
document.getElementById("addRowBtn").addEventListener("click", () => addElement(true));
document.getElementById("addColBtn").addEventListener("click", () => addElement(false));


let history = []; // Stack for undo/redo history
let historyIndex = -1; // Keeps track of the current position in the history stack
let isUndoingOrRedoing = false; // Flag to avoid redundant changes during undo/redo

// Save the current state of the table (called after every modification)
function saveState() {
  if (isUndoingOrRedoing) return; // Avoid saving state during undo/redo

  // Get the current state of the table (e.g., cells and their values)
  const currentState = [];
  const rows = table.querySelectorAll('tr');
  rows.forEach((row, rowIndex) => {
    const cells = row.querySelectorAll('td input');
    const rowData = [];
    cells.forEach(cell => rowData.push(cell.value));
    currentState.push(rowData);
  });

  // Check if we are undoing and avoid adding to history
  if (historyIndex < history.length - 1) {
    // Slice off the "redo" stack after the current position
    history = history.slice(0, historyIndex + 1);
  }

  // Only add to history if there has been a change (non-initial state)
  if (historyIndex === -1 || JSON.stringify(currentState) !== JSON.stringify(history[historyIndex])) {
    history.push(currentState);
    historyIndex++;
  }

  // Optionally, limit the history length (e.g., max 50 states)
  if (history.length > 50) history.shift(); // Limit the history length
}

// Undo the last change
function undo() {
  if (historyIndex > 0) {
    isUndoingOrRedoing = true;
    historyIndex--; // Move the index backward
    loadState(history[historyIndex]);
    isUndoingOrRedoing = false;
  }
}

// Redo the undone change
function redo() {
  if (historyIndex < history.length - 1) {
    isUndoingOrRedoing = true;
    historyIndex++; // Move the index forward
    loadState(history[historyIndex]);
    isUndoingOrRedoing = false;
  }
}

// Load a given state into the table
function loadState(state) {
  const rows = table.querySelectorAll('tr');
  rows.forEach((row, rowIndex) => {
    const cells = row.querySelectorAll('td input');
    state[rowIndex].forEach((value, colIndex) => {
      cells[colIndex].value = value;
    });
  });
}

// Listen for keyboard shortcuts (Ctrl+Z for Undo, Ctrl+Y for Redo)
document.addEventListener("keydown", function(event) {
  if (event.ctrlKey || event.metaKey) {
    if (event.key === "z" || event.key === "Z") {
      undo();  // Trigger undo on Ctrl+Z
    } else if (event.key === "y" || event.key === "Y") {
      redo();  // Trigger redo on Ctrl+Y
    }
  }
});

// Call saveState after any modification in the table (e.g., after cell editing)
table.addEventListener("input", saveState);

// Save initial state after creating the table
createTable();
saveState();


// Function to save the current state to localStorage
function saveToLocalStorage() {
  const currentState = [];
  const rows = table.querySelectorAll('tr');
  rows.forEach((row, rowIndex) => {
    const cells = row.querySelectorAll('td input');
    const rowData = [];
    cells.forEach(cell => rowData.push(cell.value));
    currentState.push(rowData);
  });
  
  localStorage.setItem("tableState", JSON.stringify(currentState));
}

// Function to load the saved state from localStorage
function loadFromLocalStorage() {
  const savedState = localStorage.getItem("tableState");
  if (savedState) {
    const state = JSON.parse(savedState);
    loadState(state);
  }
}

// Save the table state on input change
table.addEventListener("input", saveToLocalStorage);

// Load the saved state when the page loads
document.addEventListener("DOMContentLoaded", loadFromLocalStorage);

document.addEventListener('DOMContentLoaded', function() {
  const clearButton = document.getElementById('clearBtn');

  clearButton.addEventListener('click', function() {
      console.log('Clearing localStorage...');
      localStorage.clear();  // Clear the localStorage
      location.reload();     // Reload the page
  });
});
