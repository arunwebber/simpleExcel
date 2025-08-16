// A simple class to handle localStorage operations
class StorageManager {
  static saveToLocalStorage(key, value) {
    try {
      localStorage.setItem(key, JSON.stringify(value));
    } catch (e) {
      console.error("Failed to save to localStorage", e);
    }
  }

  static getFromLocalStorage(key) {
    try {
      const value = localStorage.getItem(key);
      return value ? JSON.parse(value) : null;
    } catch (e) {
      console.error("Failed to get from localStorage", e);
      return null;
    }
  }
}

// Manages the state history for undo/redo functionality
class StateManager {
  constructor(spreadsheet) {
    this.spreadsheet = spreadsheet;
    this.history = [];
    this.historyIndex = -1;
  }

  // Save the current state of the spreadsheet to the history stack
  saveState() {
    // Get the current data from the table
    const currentState = this.spreadsheet.tableManager.getCurrentData();
    const lastState = this.history[this.historyIndex];

    // Don't save if the state hasn't changed
    if (JSON.stringify(currentState) === JSON.stringify(lastState)) {
      return;
    }

    // Trim the history if we've undone actions
    if (this.historyIndex < this.history.length - 1) {
      this.history = this.history.slice(0, this.historyIndex + 1);
    }

    // Push the new state and update the index
    this.history.push(currentState);
    this.historyIndex++;

    // Keep the history stack from getting too large
    if (this.history.length > 50) {
      this.history.shift();
      this.historyIndex--;
    }

    // Save the state to localStorage
    this.spreadsheet.saveToLocalStorage();
  }

  // Undo the last action
  undo() {
    if (this.historyIndex > 0) {
      this.historyIndex--;
      this.spreadsheet.tableManager.loadData(this.history[this.historyIndex]);
      this.spreadsheet.saveToLocalStorage();
    }
  }

  // Redo the last undone action
  redo() {
    if (this.historyIndex < this.history.length - 1) {
      this.historyIndex++;
      this.spreadsheet.tableManager.loadData(this.history[this.historyIndex]);
      this.spreadsheet.saveToLocalStorage();
    }
  }

  // Handle keyboard shortcuts (Ctrl+Z and Ctrl+Y)
  handleKeyboardShortcuts(event) {
    if (event.ctrlKey) {
      if (event.key === 'z' || event.key === 'Z') {
        event.preventDefault();
        this.undo();
      } else if (event.key === 'y' || event.key === 'Y') {
        event.preventDefault();
        this.redo();
      }
    }
  }
}

// Manages highlighting of rows and columns
class LintingManager {
  constructor(spreadsheet) {
    this.spreadsheet = spreadsheet;
  }

  highlightRowAndColumn(row, col) {
    this.clearHighlighting();
    const tableRows = this.spreadsheet.table.querySelectorAll("tr");
    
    // Highlight the row header
    const rowHeader = tableRows[row + 1]?.querySelector("th");
    if (rowHeader) rowHeader.classList.add("highlight-row");

    // Highlight the column header
    const headerCells = tableRows[0]?.querySelectorAll("th");
    if (headerCells && headerCells[col + 1]) headerCells[col + 1].classList.add("highlight-col");
  }

  clearHighlighting() {
    document.querySelectorAll(".highlight-row").forEach(el => el.classList.remove("highlight-row"));
    document.querySelectorAll(".highlight-col").forEach(el => el.classList.remove("highlight-col"));
  }
}

// Manages the right-click context menu
class ContextMenuManager {
  constructor(spreadsheet) {
    this.spreadsheet = spreadsheet;
    this.contextMenu = document.getElementById('contextMenu');
    this.selectedRow = null;
    this.selectedCol = null;
    this.addEventListeners();
  }

  addEventListeners() {
    this.spreadsheet.table.addEventListener("contextmenu", this.showContextMenu.bind(this));
    document.addEventListener("click", () => this.contextMenu.style.display = 'none');
    
    // Attach event listeners to all context menu buttons
    document.getElementById("addRowAboveBtn").addEventListener("click", () => this.spreadsheet.tableManager.addElement(true, 'above', this.selectedRow));
    document.getElementById("addRowBelowBtn").addEventListener("click", () => this.spreadsheet.tableManager.addElement(true, 'below', this.selectedRow));
    document.getElementById("addColLeftBtn").addEventListener("click", () => this.spreadsheet.tableManager.addElement(false, 'left', this.selectedCol));
    document.getElementById("addColRightBtn").addEventListener("click", () => this.spreadsheet.tableManager.addElement(false, 'right', this.selectedCol));
    document.getElementById("deleteRowBtn").addEventListener("click", () => this.spreadsheet.tableManager.deleteElement(true, this.selectedRow));
    document.getElementById("deleteColBtn").addEventListener("click", () => this.spreadsheet.tableManager.deleteElement(false, this.selectedCol));
  }

  showContextMenu(event) {
    event.preventDefault();
    const cell = event.target.closest('th, td');

    if (!cell) {
        this.contextMenu.style.display = 'none';
        return;
    }

    const parentRow = cell.closest('tr');
    const isHeaderRow = parentRow.rowIndex === 0;
    const isRowHeader = cell.tagName === 'TH' && !isHeaderRow;
    const isColHeader = cell.tagName === 'TH' && isHeaderRow;

    // Show/hide buttons based on what was clicked
    document.getElementById("addRowAboveBtn").style.display = isRowHeader ? 'block' : 'none';
    document.getElementById("addRowBelowBtn").style.display = isRowHeader ? 'block' : 'none';
    document.getElementById("deleteRowBtn").style.display = isRowHeader ? 'block' : 'none';
    document.getElementById("addColLeftBtn").style.display = isColHeader ? 'block' : 'none';
    document.getElementById("addColRightBtn").style.display = isColHeader ? 'block' : 'none';
    document.getElementById("deleteColBtn").style.display = isColHeader ? 'block' : 'none';

    // Update selected row and column
    if (isRowHeader) {
      this.selectedRow = parentRow.rowIndex;
      this.selectedCol = null;
    } else if (isColHeader) {
      this.selectedRow = null;
      this.selectedCol = cell.cellIndex;
    } else {
      // Don't show the context menu for data cells
      this.contextMenu.style.display = 'none';
      return;
    }

    // Position and display the menu
    this.contextMenu.style.display = 'block';
    this.contextMenu.style.left = `${event.pageX}px`;
    this.contextMenu.style.top = `${event.pageY}px`;
  }
}

// Manages all table-related DOM manipulation
class TableManager {
  constructor(spreadsheet) {
    this.spreadsheet = spreadsheet;
    this.table = spreadsheet.table;
    this.addTableListeners();
  }

  // New method to centralize all cell event listeners
  addTableListeners() {
    this.table.addEventListener("click", (event) => {
      const cell = event.target.closest('td');
      if (cell) {
        // Find the current row and column index of the clicked cell
        const row = cell.parentElement.rowIndex - 1; // Subtract 1 for header row
        const col = cell.cellIndex - 1; // Subtract 1 for row header column
        
        // Correctly highlight the current row and column
        this.spreadsheet.lintingManager.highlightRowAndColumn(row, col);
      }
    });

    this.table.addEventListener("input", () => {
        this.spreadsheet.stateManager.saveState();
    });
  }

  // The rest of the methods remain the same as the previous correct code
  createHeaderRow() {
    const headerRow = document.createElement("tr");
    headerRow.appendChild(document.createElement("th"));
    for (let col = 1; col <= this.spreadsheet.numCols; col++) {
      const th = document.createElement("th");
      th.textContent = this.spreadsheet.getColumnName(col);
      headerRow.appendChild(th);
    }
    return headerRow;
  }

  createDataCell() {
    const td = document.createElement("td");
    const input = document.createElement("input");
    input.type = "text";
    
    td.appendChild(input);
    td.appendChild(this.createResizer("resizer", "width", td));
    td.appendChild(this.createResizer("resizer-row", "height", td));
    
    return td;
  }

  createResizer(className, dimension, td) {
    const resizer = document.createElement("div");
    resizer.className = className;
    resizer.addEventListener("mousedown", (e) => {
        e.preventDefault();
        const start = dimension === "width" ? e.clientX : e.clientY;
        const startSize = dimension === "width" ? td.offsetWidth : td.offsetHeight;

        const onMouseMove = (e) => {
            const newSize = dimension === "width" ?
                startSize + (e.clientX - start) :
                startSize + (e.clientY - start);
            
            if (newSize > 20) { // Prevents the cell from becoming too small
                if (dimension === "width") {
                    td.style.width = `${newSize}px`;
                } else {
                    td.style.height = `${newSize}px`;
                }
            }
        };

        const onMouseUp = () => {
            document.removeEventListener("mousemove", onMouseMove);
            document.removeEventListener("mouseup", onMouseUp);
        };

        document.addEventListener("mousemove", onMouseMove);
        document.addEventListener("mouseup", onMouseUp);
    });
    return resizer;
  }

  createTableRow() {
    const tr = document.createElement("tr");
    const rowHeader = document.createElement("th");
    tr.appendChild(rowHeader);

    for (let col = 0; col < this.spreadsheet.numCols; col++) {
      tr.appendChild(this.createDataCell());
    }
    return tr;
  }

  createTable() {
    this.table.innerHTML = "";
    this.table.appendChild(this.createHeaderRow());
    for (let row = 0; row < this.spreadsheet.numRows; row++) {
      this.table.appendChild(this.createTableRow());
    }
    this.updateRowHeaders();
  }

  addElement(isRow, position, index) {
    if (isRow) {
        const referenceIndex = (position === 'above') ? index : index + 1;
        const newRow = this.createTableRow();
        const referenceRow = this.table.rows[referenceIndex];
        
        if (referenceRow) {
            this.table.insertBefore(newRow, referenceRow);
        } else {
            this.table.appendChild(newRow);
        }
        this.spreadsheet.numRows++;
        this.updateRowHeaders();
    } else {
        const referenceIndex = (position === 'left') ? index : index + 1;
        const rows = this.table.querySelectorAll('tr');
        
        rows.forEach((row, rowIndex) => {
            if (rowIndex === 0) {
                const th = document.createElement('th');
                th.textContent = this.spreadsheet.getColumnName(this.spreadsheet.numCols + 1);
                row.insertBefore(th, row.cells[referenceIndex]);
            } else {
                const td = this.createDataCell();
                row.insertBefore(td, row.cells[referenceIndex]);
            }
        });
        this.spreadsheet.numCols++;
        this.updateColumnHeaders();
    }
    this.spreadsheet.stateManager.saveState();
  }

  deleteElement(isRow, index) {
    if (isRow) {
        if (this.spreadsheet.numRows <= 1) return;
        this.table.deleteRow(index);
        this.spreadsheet.numRows--;
        this.updateRowHeaders();
    } else {
        if (this.spreadsheet.numCols <= 1) return;
        const rows = this.table.querySelectorAll("tr");
        rows.forEach(row => row.deleteCell(index));
        this.spreadsheet.numCols--;
        this.updateColumnHeaders();
    }
    this.spreadsheet.stateManager.saveState();
  }

  updateRowHeaders() {
    const rows = this.table.querySelectorAll("tr");
    for (let i = 1; i < rows.length; i++) {
        const rowHeader = rows[i].querySelector("th");
        if (rowHeader) rowHeader.textContent = i;
    }
  }

  updateColumnHeaders() {
    const headerRow = this.table.querySelector("tr:first-child");
    const headers = headerRow.querySelectorAll("th");
    for (let i = 1; i < headers.length; i++) {
        headers[i].textContent = this.spreadsheet.getColumnName(i);
    }
  }
  
  clearTable() {
    this.table.querySelectorAll("td input").forEach(input => input.value = '');
    this.spreadsheet.stateManager.saveState();
  }
  
  getCurrentData() {
    const data = [];
    const rows = this.table.querySelectorAll('tr');
    for (let i = 1; i < rows.length; i++) {
        const row = rows[i];
        const rowData = [];
        const cells = row.querySelectorAll('td input');
        cells.forEach(cell => rowData.push(cell.value));
        data.push(rowData);
    }
    return data;
  }
  
  loadData(data) {
    const rows = this.table.querySelectorAll('tr');
    for (let i = 1; i < rows.length; i++) {
        const cells = rows[i].querySelectorAll('td input');
        if (data && data[i - 1]) {
            cells.forEach((cell, colIndex) => {
                const savedValue = data[i - 1][colIndex];
                cell.value = savedValue || '';
            });
        }
    }
  }
}

// The main class that ties everything together
class Spreadsheet {
  constructor(numRows = 25, numCols = 25) {
    this.numRows = numRows;
    this.numCols = numCols;
    this.table = document.getElementById("spreadsheet");

    // Initialize managers
    this.stateManager = new StateManager(this);
    this.lintingManager = new LintingManager(this);
    this.contextMenuManager = new ContextMenuManager(this);
    this.tableManager = new TableManager(this);
    
    // Load existing data or create a new table
    this.loadFromLocalStorage();
    this.addEventListeners();
  }

  getColumnName(colIndex) {
    let columnName = '';
    while (colIndex > 0) {
      columnName = String.fromCharCode(65 + ((colIndex - 1) % 26)) + columnName;
      colIndex = Math.floor((colIndex - 1) / 26);
    }
    return columnName;
  }

  addEventListeners() {
    // Add clear button listener
    document.getElementById("clearBtn")?.addEventListener("click", () => {
      this.tableManager.clearTable();
    });

    // Save state on every input change
    this.table.addEventListener("input", () => this.stateManager.saveState());

    // Handle keyboard shortcuts
    document.addEventListener("keydown", (e) => this.stateManager.handleKeyboardShortcuts(e));
  }

  saveToLocalStorage() {
    const saveObj = {
      numRows: this.numRows,
      numCols: this.numCols,
      data: this.tableManager.getCurrentData()
    };
    StorageManager.saveToLocalStorage('spreadsheetData', saveObj);
  }

  loadFromLocalStorage() {
    const saved = StorageManager.getFromLocalStorage('spreadsheetData');
    if (saved) {
      this.numRows = saved.numRows;
      this.numCols = saved.numCols;
      this.tableManager.createTable();
      this.tableManager.loadData(saved.data);
      this.stateManager.history.push(saved.data);
      this.stateManager.historyIndex = 0;
    } else {
      this.tableManager.createTable();
      this.stateManager.saveState(); // Save the initial empty state
    }
  }
}

// Initialize the spreadsheet
const spreadsheet = new Spreadsheet();