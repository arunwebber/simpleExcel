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

class StateManager {
  constructor(spreadsheet) {
    this.spreadsheet = spreadsheet;
    this.history = [];
    this.historyIndex = -1;
  }

  saveState() {
    const currentSheetData = this.spreadsheet.tableManager.getCurrentData();
    const currentSheetName = this.spreadsheet.sheetManager.activeSheetName;
    
    if (!this.history[this.historyIndex] || JSON.stringify(currentSheetData) !== JSON.stringify(this.history[this.historyIndex].data)) {
        if (this.historyIndex < this.history.length - 1) {
            this.history = this.history.slice(0, this.historyIndex + 1);
        }

        this.history.push({ sheet: currentSheetName, data: currentSheetData });
        this.historyIndex++;
        
        if (this.history.length > 50) {
            this.history.shift();
            this.historyIndex--;
        }
    }
    this.spreadsheet.saveToLocalStorage();
  }

  undo() {
    if (this.historyIndex > 0) {
      this.historyIndex--;
      const prevState = this.history[this.historyIndex];
      this.spreadsheet.sheetManager.switchSheet(prevState.sheet);
      this.spreadsheet.tableManager.loadData(prevState.data);
      this.spreadsheet.saveToLocalStorage();
    }
  }

  redo() {
    if (this.historyIndex < this.history.length - 1) {
      this.historyIndex++;
      const nextState = this.history[this.historyIndex];
      this.spreadsheet.sheetManager.switchSheet(nextState.sheet);
      this.spreadsheet.tableManager.loadData(nextState.data);
      this.spreadsheet.saveToLocalStorage();
    }
  }

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

class LintingManager {
  constructor(spreadsheet) {
    this.spreadsheet = spreadsheet;
  }

  highlightRowAndColumn(row, col) {
    this.clearHighlighting();
    const tableRows = this.spreadsheet.table.querySelectorAll("tr");

    // ✅ Highlight the entire row (row + 1 to skip column headers)
    const targetRow = tableRows[row + 1];
    if (targetRow) {
      targetRow.querySelectorAll("th, td").forEach(cell =>
        cell.classList.add("highlight-row")
      );
    }

    // ✅ Highlight the entire column (col + 1 to skip row headers)
    tableRows.forEach((tr) => {
      const cells = tr.querySelectorAll("th, td");
      if (cells[col + 1]) {
        cells[col + 1].classList.add("highlight-col");
      }
    });
  }

  clearHighlighting() {
    document.querySelectorAll(".highlight-row").forEach(el =>
      el.classList.remove("highlight-row")
    );
    document.querySelectorAll(".highlight-col").forEach(el =>
      el.classList.remove("highlight-col")
    );
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
    this.spreadsheet.table.addEventListener("contextmenu", this.showContextMenu.bind(this));
    document.addEventListener("click", () => this.contextMenu.style.display = 'none');
    
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

    document.getElementById("addRowAboveBtn").style.display = isRowHeader ? 'block' : 'none';
    document.getElementById("addRowBelowBtn").style.display = isRowHeader ? 'block' : 'none';
    document.getElementById("deleteRowBtn").style.display = isRowHeader ? 'block' : 'none';
    document.getElementById("addColLeftBtn").style.display = isColHeader ? 'block' : 'none';
    document.getElementById("addColRightBtn").style.display = isColHeader ? 'block' : 'none';
    document.getElementById("deleteColBtn").style.display = isColHeader ? 'block' : 'none';

    if (isRowHeader) {
      this.selectedRow = parentRow.rowIndex;
      this.selectedCol = null;
    } else if (isColHeader) {
      this.selectedRow = null;
      this.selectedCol = cell.cellIndex;
    } else {
      this.contextMenu.style.display = 'none';
      return;
    }

    this.contextMenu.style.display = 'block';
    this.contextMenu.style.left = `${event.pageX}px`;
    this.contextMenu.style.top = `${event.pageY}px`;
  }
}

class TableManager {
    constructor(spreadsheet) {
        this.spreadsheet = spreadsheet;
        this.table = spreadsheet.table;
        this.currentSortCol = null;
        this.currentSortDir = 'none';
        this.addTableListeners();
    }

    addTableListeners() {
        this.table.addEventListener("click", (event) => {
            const cell = event.target.closest('td');
            if (cell) {
                const row = cell.parentElement.rowIndex - 1;
                const col = cell.cellIndex - 1;
                this.spreadsheet.lintingManager.highlightRowAndColumn(row, col);
            }
        });

        this.table.addEventListener("click", (e) => {
            const sortContainer = e.target.closest('.sort-container');
            if (sortContainer) {
                const headerCell = sortContainer.closest('th');
                const colIndex = headerCell.cellIndex;
                this.sort(colIndex);
            }
        });

        this.table.addEventListener("input", () => this.spreadsheet.stateManager.saveState());

        document.addEventListener("mouseup", (event) => {
            const isResizer = event.target.className.includes('resizer');
            if (isResizer) {
                this.spreadsheet.stateManager.saveState();
            }
        });
    }

    createHeaderRow() {
        const headerRow = document.createElement("tr");
        headerRow.appendChild(document.createElement("th"));
        for (let col = 1; col <= this.spreadsheet.numCols; col++) {
            const th = document.createElement("th");
            
            const headerContent = document.createElement('div');
            headerContent.className = 'header-content';

            const headerText = document.createElement('span');
            headerText.className = 'header-text';
            headerText.textContent = this.spreadsheet.getColumnName(col);
            headerContent.appendChild(headerText);

            const sortContainer = document.createElement('div');
            sortContainer.className = 'sort-container';
            
            const upArrow = document.createElement('div');
            upArrow.className = 'sort-arrow up';
            sortContainer.appendChild(upArrow);
            
            const downArrow = document.createElement('div');
            downArrow.className = 'sort-arrow down';
            sortContainer.appendChild(downArrow);
            
            headerContent.appendChild(sortContainer);
            th.appendChild(headerContent);
            
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
                const newSize = dimension === "width"
                    ? startSize + (e.clientX - start)
                    : startSize + (e.clientY - start);

                if (newSize > 20) {
                    if (dimension === "width") {
                        const colIndex = td.cellIndex;
                        const headerCell = this.spreadsheet.table.querySelector(`tr:first-child th:nth-child(${colIndex + 1})`);
                        if (headerCell) headerCell.style.width = `${newSize}px`;

                        this.spreadsheet.table.querySelectorAll(`tr td:nth-child(${colIndex + 1})`)
                            .forEach(cell => cell.style.width = `${newSize}px`);
                    } else {
                        td.style.height = `${newSize}px`;
                    }
                }
            };

            const onMouseUp = () => {
                document.removeEventListener("mousemove", onMouseMove);
                document.removeEventListener("mouseup", onMouseUp);
                this.spreadsheet.stateManager.saveState();
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
                    
                    const headerContent = document.createElement('div');
                    headerContent.className = 'header-content';

                    const headerText = document.createElement('span');
                    headerText.className = 'header-text';
                    headerText.textContent = this.spreadsheet.getColumnName(this.spreadsheet.numCols + 1);
                    headerContent.appendChild(headerText);

                    const sortContainer = document.createElement('div');
                    sortContainer.className = 'sort-container';
                    
                    const upArrow = document.createElement('div');
                    upArrow.className = 'sort-arrow up';
                    sortContainer.appendChild(upArrow);
                    
                    const downArrow = document.createElement('div');
                    downArrow.className = 'sort-arrow down';
                    sortContainer.appendChild(downArrow);
                    
                    headerContent.appendChild(sortContainer);
                    th.appendChild(headerContent);
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
            const headerText = headers[i].querySelector('.header-text');
            if (headerText) {
                headerText.textContent = this.spreadsheet.getColumnName(i);
            }
        }
    }

    clearTable() {
        this.table.querySelectorAll("td input").forEach(input => input.value = '');
        this.table.querySelectorAll("td").forEach(td => {
            td.style.width = '';
            td.style.height = '';
        });
        this.table.querySelectorAll("tr:first-child th").forEach((th, i) => {
            if (i > 0) th.style.width = '';
        });
        this.spreadsheet.stateManager.saveState();
        this.spreadsheet.saveToLocalStorage();
    }

    getCurrentData() {
        const data = [];
        const rows = this.table.querySelectorAll('tr');
        for (let i = 1; i < rows.length; i++) {
            const row = rows[i];
            const rowData = [];
            const cells = row.querySelectorAll('td');
            cells.forEach(cell => {
                const input = cell.querySelector('input');
                rowData.push({
                    value: input.value,
                    width: cell.style.width,
                    height: cell.style.height
                });
            });
            data.push(rowData);
        }
        return data;
    }

    loadData(data) {
        const rows = this.table.querySelectorAll('tr');
        for (let i = 1; i < rows.length; i++) {
            const cells = rows[i].querySelectorAll('td');
            if (data && data[i - 1]) {
                cells.forEach((cell, colIndex) => {
                    const savedCell = data[i - 1][colIndex];
                    if (savedCell) {
                        cell.querySelector('input').value = savedCell.value || '';
                        if (savedCell.width) cell.style.width = savedCell.width;
                        if (savedCell.height) cell.style.height = savedCell.height;
                    }
                });
            }
        }
    }
    
    // The correct sort method that keeps adjacent rows together and re-renders the table
    sort(colIndex) {
        const headerCells = this.table.querySelector('tr').querySelectorAll('th');
        
        // Reset all sort indicators
        headerCells.forEach(th => {
            const upArrow = th.querySelector('.sort-arrow.up');
            const downArrow = th.querySelector('.sort-arrow.down');
            if (upArrow) upArrow.classList.remove('active');
            if (downArrow) downArrow.classList.remove('active');
        });

        let newSortDir;
        if (this.currentSortCol === colIndex) {
            newSortDir = this.currentSortDir === 'asc' ? 'desc' : 'asc';
        } else {
            this.currentSortCol = colIndex;
            newSortDir = 'asc';
        }

        this.currentSortDir = newSortDir;
        
        const currentHeader = headerCells[colIndex];
        const upArrow = currentHeader.querySelector('.sort-arrow.up');
        const downArrow = currentHeader.querySelector('.sort-arrow.down');
        
        if (newSortDir === 'asc') {
            if (upArrow) upArrow.classList.add('active');
        } else {
            if (downArrow) downArrow.classList.add('active');
        }

        // Get all data from the table into a temporary array
        const tableData = this.getCurrentData();
        
        // Sort the temporary data array
        tableData.sort((rowA, rowB) => {
            const cellA = rowA[colIndex - 1].value;
            const cellB = rowB[colIndex - 1].value;

            const isNumeric = !isNaN(parseFloat(cellA)) && isFinite(cellA) && !isNaN(parseFloat(cellB)) && isFinite(cellB);
            
            let valA = isNumeric ? parseFloat(cellA) : cellA.toLowerCase();
            let valB = isNumeric ? parseFloat(cellB) : cellB.toLowerCase();

            if (valA < valB) return newSortDir === 'asc' ? -1 : 1;
            if (valA > valB) return newSortDir === 'asc' ? 1 : -1;
            return 0;
        });
        
        // Clear and re-render the table with the sorted data
        this.clearTable();
        this.loadData(tableData);
        
        this.spreadsheet.stateManager.saveState();
    }
}
class SheetManager {
  constructor(spreadsheet) {
    this.spreadsheet = spreadsheet;
    this.sheets = {};
    this.activeSheetName = null;
    this.sheetCount = 0;
    this.tabContainer = document.getElementById('tabContainer');
    this.addTabButton = document.getElementById('addTabBtn');

    if (this.addTabButton) {
      this.addTabButton.addEventListener('click', () => this.addSheet());
    }
  }

  addSheet(sheetName = null) {
    this.sheetCount++;
    const name = sheetName || `Sheet ${this.sheetCount}`;
    this.sheets[name] = {
      numRows: 25,
      numCols: 25,
      data: [],
      columnWidths: []
    };
    this.createTab(name);
    this.switchSheet(name);
    this.spreadsheet.saveToLocalStorage();
  }

  createTab(name) {
    const tab = document.createElement('div');
    tab.className = 'tab';
    tab.dataset.sheetName = name;

    const titleSpan = document.createElement('span');
    titleSpan.className = 'tab-title';
    titleSpan.textContent = name;
    
    const closeBtn = document.createElement('span');
    closeBtn.className = 'close-btn';
    closeBtn.textContent = 'x';
    closeBtn.addEventListener('click', (e) => {
      e.stopPropagation();
      this.deleteSheet(name);
    });

    tab.appendChild(titleSpan);
    tab.appendChild(closeBtn);
    
    this.tabContainer.insertBefore(tab, this.addTabButton);
    tab.addEventListener('click', () => this.switchSheet(name));

    tab.addEventListener('dblclick', () => this.handleTabTitleDoubleClick(tab));
  }

  handleTabTitleDoubleClick(tab) {
    const titleSpan = tab.querySelector('.tab-title');
    const originalTitle = titleSpan.textContent;
    
    const input = document.createElement('input');
    input.type = 'text';
    input.value = originalTitle;
    input.className = 'edit-title'; 
    input.style.width = `${titleSpan.offsetWidth}px`;

    titleSpan.style.display = 'none';
    tab.insertBefore(input, titleSpan);
    
    input.focus();
    input.select();

    const saveTitle = () => {
        const newTitle = input.value.trim();
        if (newTitle && newTitle !== originalTitle) {
            this.updateSheetName(originalTitle, newTitle);
        }
        input.remove();
        titleSpan.style.display = 'block';
    };

    input.addEventListener('blur', saveTitle);
    input.addEventListener('keydown', (e) => {
        if (e.key === 'Enter') {
            saveTitle();
        }
    });
  }

  updateSheetName(oldName, newName) {
      if (this.sheets[newName]) {
          alert("A sheet with this name already exists. Please choose a different name.");
          return;
      }

      const orderedSheets = Object.keys(this.sheets);
      const newSheets = {};
      
      orderedSheets.forEach(sheetName => {
          if (sheetName === oldName) {
              newSheets[newName] = this.sheets[oldName];
          } else {
              newSheets[sheetName] = this.sheets[sheetName];
          }
      });
      this.sheets = newSheets;

      if (this.activeSheetName === oldName) {
          this.activeSheetName = newName;
      }

      this.spreadsheet.stateManager.history.forEach(state => {
          if (state.sheet === oldName) {
              state.sheet = newName;
          }
      });

      const oldTab = this.tabContainer.querySelector(`[data-sheet-name="${oldName}"]`);
      if (oldTab) {
          oldTab.dataset.sheetName = newName;
          oldTab.querySelector('.tab-title').textContent = newName;
          oldTab.removeEventListener('click', this.switchSheet);
          oldTab.addEventListener('click', () => this.switchSheet(newName));
          oldTab.querySelector('.close-btn').removeEventListener('click', this.deleteSheet);
          oldTab.querySelector('.close-btn').addEventListener('click', (e) => {
              e.stopPropagation();
              this.deleteSheet(newName);
          });
      }
      this.spreadsheet.saveToLocalStorage();
  }

  deleteSheet(name) {
    if (Object.keys(this.sheets).length <= 1) {
      alert("Cannot delete the last sheet.");
      return;
    }

    const sheetNames = Object.keys(this.sheets);
    const deletedIndex = sheetNames.indexOf(name);
    
    delete this.sheets[name];
    const tabToDelete = this.tabContainer.querySelector(`[data-sheet-name="${name}"]`);
    if (tabToDelete) {
      tabToDelete.remove();
    }

    const newActiveSheetNames = Object.keys(this.sheets);
    const newActiveSheetIndex = (deletedIndex === sheetNames.length - 1) ? deletedIndex - 1 : deletedIndex;
    this.switchSheet(newActiveSheetNames[newActiveSheetIndex]);
    
    this.spreadsheet.saveToLocalStorage();
  }

  switchSheet(name) {
    if (this.activeSheetName) {
      this.saveActiveSheetState();
      this.tabContainer.querySelector(`[data-sheet-name="${this.activeSheetName}"]`)?.classList.remove('active');
    }

    this.activeSheetName = name;
    if (this.sheets[name]) {
        this.spreadsheet.numRows = this.sheets[name].numRows;
        this.spreadsheet.numCols = this.sheets[name].numCols;

        this.spreadsheet.tableManager.createTable();
        this.spreadsheet.tableManager.loadData(this.sheets[name].data);
        this.spreadsheet.applyColumnWidths(this.sheets[name].columnWidths);
    }
    
    this.tabContainer.querySelector(`[data-sheet-name="${name}"]`)?.classList.add('active');
    this.spreadsheet.saveToLocalStorage();
  }

  saveActiveSheetState() {
    if (!this.activeSheetName || !this.sheets[this.activeSheetName]) return;
    this.sheets[this.activeSheetName].numRows = this.spreadsheet.numRows;
    this.sheets[this.activeSheetName].numCols = this.spreadsheet.numCols;
    this.sheets[this.activeSheetName].data = this.spreadsheet.tableManager.getCurrentData();
    this.sheets[this.activeSheetName].columnWidths = this.spreadsheet.getColumnWidths();
  }

  loadSheets(sheetsData, activeSheet) {
    this.sheets = sheetsData;
    this.sheetCount = Object.keys(sheetsData).length;

    const tabs = this.tabContainer.querySelectorAll('.tab');
    tabs.forEach(tab => tab.remove());
    
    Object.keys(this.sheets).forEach(name => this.createTab(name));
    
    this.switchSheet(activeSheet);
  }
}

class DownloadPrintManager {
  constructor(spreadsheet) {
    this.spreadsheet = spreadsheet;
  }

  downloadAsCSV() {
    let csv = [];
    const rows = this.spreadsheet.table.querySelectorAll("tr");
    
    for (let i = 0; i < rows.length; i++) {
        let row = [], cols = rows[i].querySelectorAll("td, th");
        for (let j = 0; j < cols.length; j++) {
            let cellValue = "";
            if (cols[j].tagName === "TD") {
                cellValue = cols[j].querySelector("input").value;
            } else if (cols[j].tagName === "TH") {
                cellValue = cols[j].innerText.trim();
                cellValue = cellValue.replace(/\s*$/, ''); // Remove trailing space from the sort indicator
            }
            row.push(`"${cellValue.replace(/"/g, '""')}"`); // Handle commas and quotes
        }
        csv.push(row.join(","));
    }

    const csvFile = new Blob([csv.join("\n")], { type: "text/csv" });
    const downloadLink = document.createElement("a");
    downloadLink.download = "spreadsheet.csv";
    downloadLink.href = window.URL.createObjectURL(csvFile);
    downloadLink.style.display = "none";
    document.body.appendChild(downloadLink);
    downloadLink.click();
    document.body.removeChild(downloadLink);
  }

  printTable() {
    const tableToPrint = this.spreadsheet.table.outerHTML;
    const printWindow = window.open('', '', 'height=600,width=800');
    printWindow.document.write('<html><head><title>Print Spreadsheet</title>');
    printWindow.document.write('<style>');
    printWindow.document.write('body { font-family: Arial, sans-serif; }');
    printWindow.document.write('table { border-collapse: collapse; width: 100%; }');
    printWindow.document.write('th, td { border: 1px solid #000; padding: 8px; text-align: left; }');
    printWindow.document.write('th { background-color: #f2f2f2; }');
    printWindow.document.write('</style>');
    printWindow.document.write('</head><body>');
    printWindow.document.write(tableToPrint);
    printWindow.document.write('</body></html>');
    printWindow.document.close();
    printWindow.focus();
    printWindow.print();
  }
}

class Spreadsheet {
  constructor(numRows, numCols, tableId = 'spreadsheet') {
    this.numRows = numRows;
    this.numCols = numCols;
    this.table = document.getElementById(tableId);

    this.tableManager = new TableManager(this);
    this.stateManager = new StateManager(this);
    this.lintingManager = new LintingManager(this);
    this.contextMenuManager = new ContextMenuManager(this);
    this.sheetManager = new SheetManager(this);
    this.downloadPrintManager = new DownloadPrintManager(this);

    this.loadState();
    this.addGlobalListeners();

    const storedState = StorageManager.getFromLocalStorage('spreadsheetState');
    if (!storedState || Object.keys(storedState.sheets).length === 0) {
      this.sheetManager.addSheet();
    } else {
      this.sheetManager.loadSheets(storedState.sheets, storedState.activeSheet);
    }
  }

  addGlobalListeners() {
    document.getElementById("clearBtn").addEventListener("click", () => this.tableManager.clearTable());
    document.getElementById("downloadBtn").addEventListener("click", () => this.downloadPrintManager.downloadAsCSV());
    document.getElementById("printBtn").addEventListener("click", () => this.downloadPrintManager.printTable());
    document.getElementById("darkModeToggle").addEventListener("change", (e) => this.toggleDarkMode(e.target.checked));
    document.addEventListener("keydown", (e) => this.stateManager.handleKeyboardShortcuts(e));
  }

  saveToLocalStorage() {
    this.sheetManager.saveActiveSheetState();
    const state = {
      sheets: this.sheetManager.sheets,
      activeSheet: this.sheetManager.activeSheetName,
      darkMode: document.body.classList.contains('dark-mode')
    };
    StorageManager.saveToLocalStorage('spreadsheetState', state);
  }

  loadState() {
    const storedState = StorageManager.getFromLocalStorage('spreadsheetState');
    if (storedState) {
      if (storedState.darkMode) {
        document.body.classList.add('dark-mode');
        document.getElementById("darkModeToggle").checked = true;
      }
    }
  }

  toggleDarkMode(isDarkMode) {
    if (isDarkMode) {
      document.body.classList.add('dark-mode');
    } else {
      document.body.classList.remove('dark-mode');
    }
    this.saveToLocalStorage();
  }

  getColumnName(index) {
    let temp,
      letter = '';
    while (index > 0) {
      temp = (index - 1) % 26;
      letter = String.fromCharCode(temp + 65) + letter;
      index = (index - temp - 1) / 26;
    }
    return letter;
  }

  getColumnWidths() {
    const widths = [];
    const headerCells = this.table.querySelectorAll('tr:first-child th');
    for (let i = 1; i < headerCells.length; i++) {
      widths.push(headerCells[i].style.width);
    }
    return widths;
  }

  applyColumnWidths(widths) {
    const headerCells = this.table.querySelectorAll('tr:first-child th');
    for (let i = 1; i < headerCells.length; i++) {
      if (widths[i - 1]) {
        headerCells[i].style.width = widths[i - 1];
      }
    }
  }
}

document.addEventListener('DOMContentLoaded', () => {
  new Spreadsheet(25, 25);
});