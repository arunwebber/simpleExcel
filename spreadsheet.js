class StorageManager {
  // Save data to localStorage
  static saveToLocalStorage(key, value) {
      localStorage.setItem(key, value);
  }

  // Retrieve data from localStorage
  static getFromLocalStorage(key, defaultValue = '') {
      return localStorage.getItem(key) || defaultValue;
  }
}

class InputManager {
  constructor(spreadsheet) {
    this.spreadsheet = spreadsheet;
  }

  handleCellInput(event) {
    const input = event.target;
    const td = input.parentElement;
    const row = td.parentElement;
    const col = Array.from(row.children).indexOf(td) - 1; // Index of column in the row
    const rowIndex = Array.from(this.spreadsheet.table.children).indexOf(row) - 1; // Index of row in the table

    // Save the current input state in the history
    this.spreadsheet.saveState();
  }
}

class KeyboardShortcutManager {
  constructor(noteElement) {
    this.noteElement = noteElement;
    this.undoStack = [];
    this.redoStack = [];
  }

  // Save the current state to undo stack
  saveState() {
    this.undoStack.push(this.noteElement.innerText);
    if (this.undoStack.length > 50) {
      this.undoStack.shift(); // Keep the stack size manageable
    }
    this.redoStack = []; // Clear redo stack after a new action
  }

  // Undo the last action
  undo() {
    if (this.undoStack.length > 0) {
      const lastState = this.undoStack.pop();
      this.redoStack.push(this.noteElement.innerText);
      this.noteElement.innerText = lastState;
      StorageManager.saveToLocalStorage('savedNote', this.noteElement.innerText);
    }
  }

  // Redo the last undone action
  redo() {
    if (this.redoStack.length > 0) {
      const lastState = this.redoStack.pop();
      this.undoStack.push(this.noteElement.innerText);
      this.noteElement.innerText = lastState;
      StorageManager.saveToLocalStorage('savedNote', this.noteElement.innerText);
    }
  }

  // Handle keyboard shortcuts for undo and redo
  handleKeyboardShortcuts(event) {
    if (event.ctrlKey) {
      if (event.key === 'z' || event.key === 'Z') {
        event.preventDefault();
        this.undo(); // Perform undo action
      } else if (event.key === 'y' || event.key === 'Y') {
        event.preventDefault();
        this.redo(); // Perform redo action
      }
    }
  }
}

class LintingManager {
  constructor(spreadsheet) {
    this.spreadsheet = spreadsheet;
  }

  highlightRowAndColumn(row, col) {
    // Highlight the row header
    const rowHeader = this.spreadsheet.table.querySelectorAll("tr")[row + 1].querySelector("th");
    rowHeader.classList.add("highlight-row");

    // Highlight the column header
    const headerCells = this.spreadsheet.table.querySelectorAll("tr:first-child th");
    headerCells[col + 1].classList.add("highlight-col");
  }

  clearHighlighting() {
    // Remove highlighting from all rows and columns
    document.querySelectorAll(".highlight-row").forEach((row) => row.classList.remove("highlight-row"));
    document.querySelectorAll(".highlight-col").forEach((col) => col.classList.remove("highlight-col"));
  }
}

class ContextMenuManager {
  constructor(spreadsheet) {
    this.spreadsheet = spreadsheet;
    this.contextMenu = document.getElementById('contextMenu');
    this.selectedRow = null;
    this.selectedCol = null;
    this.addEventListeners();
  }

  addEventListeners() {
    this.spreadsheet.table.addEventListener("contextmenu", (event) => {
      event.preventDefault();
      this.showContextMenu(event);
    });

    // New Event Listeners for Delete Buttons
    document.getElementById("deleteRowBtn").addEventListener("click", () => {
        if (this.selectedRow !== null) {
            this.spreadsheet.tableManager.deleteElement(true, this.selectedRow);
        }
    });

    document.getElementById("deleteColBtn").addEventListener("click", () => {
        if (this.selectedCol !== null) {
            this.spreadsheet.tableManager.deleteElement(false, this.selectedCol);
        }
    });

    document.addEventListener("click", () => this.contextMenu.style.display = 'none');
  }

  showContextMenu(event) {
    const cell = event.target;

    if (cell.tagName === 'TH') {
      const isColumnHeader = cell.parentElement === this.spreadsheet.table.querySelector("tr:first-child");
      
      // Updated logic to show/hide all buttons
      document.getElementById("addRowBtn").style.display = isColumnHeader ? 'none' : 'block';
      document.getElementById("deleteRowBtn").style.display = isColumnHeader ? 'none' : 'block';
      document.getElementById("addColBtn").style.display = isColumnHeader ? 'block' : 'none';
      document.getElementById("deleteColBtn").style.display = isColumnHeader ? 'block' : 'none';

      this.contextMenu.style.display = 'block';
      this.contextMenu.style.left = `${event.pageX}px`;
      this.contextMenu.style.top = `${event.pageY}px`;

      if (isColumnHeader) {
        this.selectedRow = null;
        this.selectedCol = Array.from(cell.parentElement.children).indexOf(cell);
      } else {
        // Correctly identify the row index by its position in the table body
        this.selectedRow = Array.from(cell.parentElement.parentElement.children).indexOf(cell.parentElement);
        this.selectedCol = null;
      }
    } else {
      this.contextMenu.style.display = 'none';
    }
  }
}

class TableManager {
  constructor(spreadsheet) {
    this.spreadsheet = spreadsheet;
    this.table = spreadsheet.table;
  }

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

    resizer.addEventListener("mousedown", (e) => this.handleResize(e, dimension, td));
    return resizer;
  }

  handleResize(e, dimension, td) {
    e.preventDefault();
    const start = dimension === "width" ? e.clientX : e.clientY;
    const startSize = dimension === "width" ? td.offsetWidth : td.offsetHeight;

    const onMouseMove = (e) => {
      const newSize = dimension === "width"
        ? startSize + (e.clientX - start)
        : startSize + (e.clientY - start);

      if (dimension === "width") {
        td.style.width = `${newSize}px`;
      } else {
        td.style.height = `${newSize}px`;
      }
    };

    const onMouseUp = () => {
      document.removeEventListener("mousemove", onMouseMove);
      document.removeEventListener("mouseup", onMouseUp);
    };

    document.addEventListener("mousemove", onMouseMove);
    document.addEventListener("mouseup", onMouseUp);
  }

  createTableRow(row) {
    const tr = document.createElement("tr");
    const rowHeader = document.createElement("th");
    rowHeader.textContent = row + 1;
    tr.appendChild(rowHeader);

    for (let col = 0; col < this.spreadsheet.numCols; col++) {
      const td = this.createDataCell();
      td.addEventListener("click", () => {
        this.spreadsheet.lintingManager.clearHighlighting();
        this.spreadsheet.lintingManager.highlightRowAndColumn(row, col);
      });

      td.querySelector("input").addEventListener("input", (e) => this.spreadsheet.inputManager.handleCellInput(e));
      tr.appendChild(td);
    }

    return tr;
  }

  createTable() {
    this.table.innerHTML = "";
    const headerRow = this.createHeaderRow();
    this.table.appendChild(headerRow);

    for (let row = 0; row < this.spreadsheet.numRows; row++) {
      const tableRow = this.createTableRow(row);
      this.table.appendChild(tableRow);
    }
  }

  addElement(isRow) {
    if (isRow) {
      const tr = this.createTableRow(this.spreadsheet.numRows);
      this.table.appendChild(tr);
      this.spreadsheet.numRows++;
    } else {
      this.spreadsheet.numCols++;
      const headerRow = this.table.querySelector("tr:first-child");
      const th = document.createElement("th");
      th.textContent = this.spreadsheet.getColumnName(this.spreadsheet.numCols);
      headerRow.appendChild(th);

      const rows = this.table.querySelectorAll("tr");
      for (let i = 1; i < rows.length; i++) {
        const td = this.createDataCell();
        rows[i].appendChild(td);
      }
    }
    this.spreadsheet.contextMenuManager.contextMenu.style.display = 'none';
    this.spreadsheet.saveState();
  }

  // --- NEW: Method to delete a row or column ---
  deleteElement(isRow, index) {
      if (isRow) {
          // Prevent deleting all rows
          if (this.spreadsheet.numRows <= 1) return;
          this.table.deleteRow(index);
          this.spreadsheet.numRows--;
          this.updateRowHeaders();
      } else {
          // Prevent deleting all columns
          if (this.spreadsheet.numCols <= 1) return;
          const rows = this.table.querySelectorAll("tr");
          rows.forEach(row => {
              row.deleteCell(index);
          });
          this.spreadsheet.numCols--;
          this.updateColumnHeaders();
      }
      this.spreadsheet.contextMenuManager.contextMenu.style.display = 'none';
      this.spreadsheet.saveState(); // Save state after deletion
  }

  // --- NEW: Helper to re-number row headers after deletion ---
  updateRowHeaders() {
      const rows = this.table.querySelectorAll("tr");
      // Start from 1 to skip the column header row
      for (let i = 1; i < rows.length; i++) {
          const rowHeader = rows[i].querySelector("th");
          if (rowHeader) {
              rowHeader.textContent = i;
          }
      }
  }

  // --- NEW: Helper to re-letter column headers after deletion ---
  updateColumnHeaders() {
      const headerRow = this.table.querySelector("tr:first-child");
      const headers = headerRow.querySelectorAll("th");
      // Start from 1 to skip the empty top-left corner
      for (let i = 1; i < headers.length; i++) {
          headers[i].textContent = this.spreadsheet.getColumnName(i);
      }
  }

  clearTable() {
    const inputs = this.table.querySelectorAll("td input");
    inputs.forEach(input => input.value = '');
  }
}

class Spreadsheet {
  constructor(numRows = 25, numCols = 25) {
    this.numRows = numRows;
    this.numCols = numCols;
    this.table = document.getElementById("spreadsheet");
    this.history = [];
    this.historyIndex = -1;
    this.isUndoingOrRedoing = false;

    this.inputManager = new InputManager(this);
    this.keyboardShortcutManager = new KeyboardShortcutManager(this.table);
    this.lintingManager = new LintingManager(this);
    this.contextMenuManager = new ContextMenuManager(this);
    this.tableManager = new TableManager(this);

    this.tableManager.createTable();
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

  saveState() {
    if (this.isUndoingOrRedoing) return;

    const currentState = Array.from(this.table.querySelectorAll('tr')).map(row =>
      Array.from(row.querySelectorAll('td input')).map(cell => cell ? cell.value : null)
    ).filter(row => row.length > 0);

    if (this.historyIndex < this.history.length - 1) {
      this.history = this.history.slice(0, this.historyIndex + 1);
    }

    if (this.historyIndex === -1 || JSON.stringify(currentState) !== JSON.stringify(this.history[this.historyIndex])) {
      this.history.push(currentState);
      this.historyIndex++;
    }

    if (this.history.length > 50) this.history.shift();

    this.saveToLocalStorage(currentState);
  }

  undo() { this.keyboardShortcutManager.undo(); }

  redo() { this.keyboardShortcutManager.redo(); }

  loadState(state) {
    const rows = this.table.querySelectorAll('tr');
    for (let i = 1; i < rows.length; i++) {
        const cells = rows[i].querySelectorAll('td input');
        if (state && state[i-1]) {
            cells.forEach((cell, colIndex) => {
                const savedValue = state[i-1][colIndex];
                cell.value = savedValue || ''; 
            });
        }
    }
  }

  addEventListeners() {
    const addRowBtn = document.getElementById("addRowBtn");
    const addColBtn = document.getElementById("addColBtn");
    const clearBtn = document.getElementById("clearBtn");

    addRowBtn?.addEventListener("click", () => this.tableManager.addElement(true));
    addColBtn?.addEventListener("click", () => this.tableManager.addElement(false));
    clearBtn?.addEventListener("click", () => {
      this.tableManager.clearTable();
      this.saveState();
    });

    this.table.addEventListener("input", () => this.saveState());
  }

  // 🔹 Save both table structure & cell data
  saveToLocalStorage(state) {
    const saveObj = {
      numRows: this.numRows,
      numCols: this.numCols,
      data: state
    };
    localStorage.setItem('spreadsheetData', JSON.stringify(saveObj));
  }

  // 🔹 Load both table structure & cell data
  loadFromLocalStorage() {
    const saved = JSON.parse(localStorage.getItem('spreadsheetData'));
    if (saved) {
      this.numRows = saved.numRows || this.numRows;
      this.numCols = saved.numCols || this.numCols;
      this.tableManager.createTable(); // rebuild table with saved structure
      this.loadState(saved.data); // load saved cell values
    }
  }
}


// Create a new spreadsheet
const spreadsheet = new Spreadsheet();