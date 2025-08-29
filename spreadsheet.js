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

class ApiKeyManager {
  constructor() {
    this.modal = document.getElementById('apiKeyModal');
    this.apiKeyInput = document.getElementById('apiKeyInput');
    this.saveBtn = document.getElementById('saveApiKeyBtn');
    this.closeBtn = document.getElementById('closeModalBtn');
    this.settingsBtn = document.getElementById('settingsButton');

    this.bindEvents();
  }

  bindEvents() {
    this.settingsBtn.addEventListener('click', () => {
      this.showModal();
    });
    this.saveBtn.addEventListener('click', () => {
      this.saveApiKey();
    });
    this.closeBtn.addEventListener('click', () => {
      this.hideModal();
    });
    this.modal.addEventListener('click', (event) => {
      if (event.target === this.modal) {
        this.hideModal();
      }
    });
  }

  showModal() {
    const existingKey = StorageManager.getFromLocalStorage('apiKey');
    if (existingKey) {
      this.apiKeyInput.value = existingKey;
    } else {
      this.apiKeyInput.value = '';
    }
    this.modal.style.display = 'flex';
  }

  hideModal() {
    this.modal.style.display = 'none';
  }

  saveApiKey() {
    const key = this.apiKeyInput.value.trim();
    if (key) {
      StorageManager.saveToLocalStorage('apiKey', key);
      alert('API Key saved successfully!');
      this.hideModal();
    } else {
      alert('Please enter a valid API key.');
    }
  }
}

class SpreadsheetAIIntegration {
    constructor(spreadsheet) {
        this.spreadsheet = spreadsheet;
        this.aiPanel = document.getElementById('rightPanel');
        this.aiNoteElement = document.querySelector('.ai-note');
        this.aiWriteButton = document.getElementById('aiWriteButton');
        this.apiKeyManager = new ApiKeyManager();
        this.pollingTimeout = null;
        this.apiCache = {};
        this.bindEvents();
    }

    bindEvents() {
        this.aiWriteButton.addEventListener('click', () => this.handleAiWrite());
    }

    async handleAiWrite() {
        const apiKey = StorageManager.getFromLocalStorage('apiKey');
        if (!apiKey) {
            this.apiKeyManager.showModal();
            return;
        }

        const dataToSummarize = this.extractSheetData();
        if (!dataToSummarize) {
            this.aiNoteElement.textContent = 'No data to summarize.';
            return;
        }

        const cacheKey = await this.generateCacheKey(dataToSummarize);

        if (this.apiCache[cacheKey]) {
            this.aiNoteElement.textContent = this.apiCache[cacheKey];
            return;
        }

        this.aiNoteElement.textContent = 'Loading...';

        try {
            const requestBody = { content: dataToSummarize };
            const response = await this.callSharpApi('POST', '/v1/content/summarize', apiKey, requestBody);

            if (response.status_url) {
                this.aiNoteElement.textContent = 'Job accepted. Waiting for result...';
                this.pollForStatus(response.status_url, apiKey, cacheKey);
            }

        } catch (error) {
            this.aiNoteElement.textContent = `Error: ${error.message}`;
        }
    }

    extractSheetData() {
        const data = this.spreadsheet.tableManager.getCurrentData();
        if (!data || data.length === 0) return null;

        let extractedText = '';
        const headerCells = this.spreadsheet.table.querySelectorAll('tr:first-child th');
        const headers = Array.from(headerCells).slice(1).map(th => th.querySelector('.header-text').textContent);

        extractedText += headers.join('\t') + '\n';
        data.forEach(row => {
            const rowValues = row.map(cell => cell.value);
            extractedText += rowValues.join('\t') + '\n';
        });

        return extractedText.trim();
    }

    async generateCacheKey(content) {
        const encoder = new TextEncoder();
        const data = encoder.encode(content);
        const hashBuffer = await crypto.subtle.digest('SHA-256', data);
        const hashArray = Array.from(new Uint8Array(hashBuffer));
        const hashHex = hashArray.map(b => b.toString(16).padStart(2, '0')).join('');
        return `summary-${hashHex.substring(0, 16)}`;
    }

    async pollForStatus(statusUrl, apiKey, cacheKey, retries = 0) {
        const maxRetries = 10;
        const pollInterval = 3000;

        if (retries >= maxRetries) {
            this.aiNoteElement.textContent = 'Job timed out. Please try again.';
            return;
        }

        try {
            const response = await this.callSharpApi('GET', statusUrl, apiKey);

            if (response.data.attributes.status === 'completed' || response.data.attributes.status === 'success') {
                const result = JSON.parse(response.data.attributes.result);
                const finalResult = result.summary || 'No summary found.';
                this.aiNoteElement.textContent = finalResult;
                this.apiCache[cacheKey] = finalResult;
                return;
            } else if (response.data.attributes.status === 'failed') {
                this.aiNoteElement.textContent = `Job failed: ${response.data.attributes.message || 'Unknown error'}`;
                return;
            }

            this.pollingTimeout = setTimeout(() => this.pollForStatus(statusUrl, apiKey, cacheKey, retries + 1), pollInterval);

        } catch (error) {
            this.aiNoteElement.textContent = `Polling failed: ${error.message}`;
        }
    }

    async callSharpApi(method, path, apiKey, body = null) {
        const url = path.startsWith('http') ? path : `https://sharpapi.com/api${path}`;

        const headers = new Headers();
        headers.append("Accept", "application/json");
        headers.append("Authorization", `Bearer ${apiKey}`);

        if (method === 'POST') {
            headers.append("Content-Type", "application/json");
        }

        const requestOptions = {
            method,
            headers,
            body: body ? JSON.stringify(body) : null,
            redirect: "follow"
        };

        const response = await fetch(url, requestOptions);
        const result = await response.json();

        if (!response.ok && response.status !== 202) {
            throw new Error(result.message || `API call failed with status ${response.status}`);
        }

        return result;
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

class FormulaManager {
    constructor(spreadsheet) {
        this.spreadsheet = spreadsheet;
        this.cache = new Map();
        this.precedence = {
            '+': 1,
            '-': 1,
            '*': 2,
            '/': 2,
        };
    }

    isOperator(token) {
        return token in this.precedence;
    }

    toPostfix(expression) {
        const outputQueue = [];
        const operatorStack = [];
        const tokens = expression.match(/([A-Z]+[0-9]+|\d+(\.\d+)?|[+\-*/()]| )/g).filter(t => t.trim() !== '');

        for (const token of tokens) {
            if (!isNaN(parseFloat(token))) {
                outputQueue.push(token);
            } else if (token.match(/[A-Z]+[0-9]+/)) {
                outputQueue.push(token);
            } else if (this.isOperator(token)) {
                while (
                    operatorStack.length > 0 &&
                    this.isOperator(operatorStack[operatorStack.length - 1]) &&
                    this.precedence[operatorStack[operatorStack.length - 1]] >= this.precedence[token]
                ) {
                    outputQueue.push(operatorStack.pop());
                }
                operatorStack.push(token);
            } else if (token === '(') {
                operatorStack.push(token);
            } else if (token === ')') {
                while (operatorStack.length > 0 && operatorStack[operatorStack.length - 1] !== '(') {
                    outputQueue.push(operatorStack.pop());
                }
                if (operatorStack[operatorStack.length - 1] === '(') {
                    operatorStack.pop();
                }
            }
        }
        while (operatorStack.length > 0) {
            outputQueue.push(operatorStack.pop());
        }
        return outputQueue;
    }

    evaluatePostfix(postfix) {
        const stack = [];
        for (const token of postfix) {
            if (!isNaN(parseFloat(token))) {
                stack.push(parseFloat(token));
            } else if (this.isOperator(token)) {
                if (stack.length < 2) {
                    return '#ERROR!';
                }
                const operand2 = stack.pop();
                const operand1 = stack.pop();
                let result;
                switch (token) {
                    case '+':
                        result = operand1 + operand2;
                        break;
                    case '-':
                        result = operand1 - operand2;
                        break;
                    case '*':
                        result = operand1 * operand2;
                        break;
                    case '/':
                        if (operand2 === 0) return '#DIV/0!';
                        result = operand1 / operand2;
                        break;
                }
                stack.push(result);
            }
        }

        if (stack.length !== 1) {
            return '#ERROR!';
        }

        return stack.pop();
    }

    evaluateFormula(formula) {
        if (this.cache.has(formula)) {
            return this.cache.get(formula);
        }

        try {
            let expression = formula.substring(1).toUpperCase();
            
            if (expression.trim() === '') {
                return '';
            }
            
            // Replace cell references with their values
            expression = expression.replace(/[A-Z]+[0-9]+/g, (match) => {
                const colName = match.match(/[A-Z]+/)[0];
                const rowNum = parseInt(match.match(/[0-9]+/)[0], 10);
                const colIndex = this.spreadsheet.columnNameToIndex(colName);
                const cell = this.spreadsheet.table.querySelector(`tr:nth-child(${rowNum + 1}) td:nth-child(${colIndex + 2})`);
                
                if (cell) {
                    const value = cell.querySelector('input').value;
                    const parsedValue = parseFloat(value);
                    return isNaN(parsedValue) ? 0 : parsedValue;
                }
                return 0;
            });
            
            const postfix = this.toPostfix(expression);
            const result = this.evaluatePostfix(postfix);

            if (result === Infinity || result === -Infinity) {
                return '#DIV/0!';
            }
            
            this.cache.set(formula, result);
            return result;
        } catch (e) {
            console.error("Formula evaluation error:", e);
            return '#ERROR!';
        }
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

        this.table.addEventListener("input", (e) => {
            const inputElement = e.target;
            const cell = inputElement.closest('td');
            if (cell) {
                // If user deletes everything, also clear the formula
                if (inputElement.value.trim() === '') {
                    inputElement.dataset.formula = '';
                }
                const rowIndex = cell.parentElement.rowIndex - 1;
                const colIndex = cell.cellIndex - 1;
                this.spreadsheet.formulaManager.cache.clear();
                this.spreadsheet.recalculateAllFormulas();
            }
            this.spreadsheet.stateManager.saveState();
        });

        
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
        
        // Add the autofill handle
        const autofillHandle = document.createElement("div");
        autofillHandle.className = "autofill-handle";
        td.appendChild(autofillHandle);
        
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
        this.table.querySelectorAll("td input").forEach(input => {
          input.value = '';
          input.dataset.formula = '';
        });
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
                    formula: input.dataset.formula || '',
                    width: cell.style.width,
                    height: cell.style.height
                });
            });
            data.push(rowData);
        }
        return data;
    }
    
    // Original function for loading from saved state
    loadData(data) {
        if (!data || data.length === 0) return;

        // Ensure table has enough rows
        const requiredRows = data.length;
        while (this.spreadsheet.numRows < requiredRows) {
            this.addElement(true, 'below', this.spreadsheet.numRows);
        }

        // Ensure table has enough columns
        const requiredCols = data[0].length;
        while (this.spreadsheet.numCols < requiredCols) {
            this.addElement(false, 'right', this.spreadsheet.numCols);
        }

        const rows = this.table.querySelectorAll('tr');
        for (let i = 1; i < rows.length; i++) {
            const cells = rows[i].querySelectorAll('td');
            const rowData = data[i - 1] || [];
            cells.forEach((cell, colIndex) => {
                const savedCell = rowData[colIndex];
                if (savedCell) {
                    const input = cell.querySelector('input');
                    input.value = savedCell.value || '';
                    input.dataset.formula = savedCell.formula || '';
                    if (savedCell.width) cell.style.width = savedCell.width;
                    if (savedCell.height) cell.style.height = savedCell.height;
                }
            });
        }
        this.spreadsheet.recalculateAllFormulas();
    }
    
    // NEW function for loading data from CSV file
    loadDataFromCSV(data) {
        if (!data || data.length === 0) return;

        // Populate cells, starting from the first data row (second HTML row)
        const rows = this.table.querySelectorAll('tr');
        for (let i = 1; i < rows.length; i++) {
            const cells = rows[i].querySelectorAll('td');
            // The CSV data's index corresponds directly to the HTML table's data row index
            const rowData = data[i-1] || []; 
            cells.forEach((cell, colIndex) => {
                const savedCell = rowData[colIndex];
                if (savedCell) {
                    const input = cell.querySelector('input');
                    input.value = savedCell.value || '';
                    input.dataset.formula = savedCell.formula || '';
                }
            });
        }
        this.spreadsheet.recalculateAllFormulas();
    }

    sort(colIndex) {
        const headerCells = this.table.querySelector('tr').querySelectorAll('th');

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

        const tableData = this.getCurrentData();

        tableData.sort((rowA, rowB) => {
            const cellA = rowA[colIndex - 1].value;
            const cellB = rowB[colIndex - 1].value;

            const aIsEmpty = cellA === null || cellA.trim() === '';
            const bIsEmpty = cellB === null || cellB.trim() === '';

            if (aIsEmpty && !bIsEmpty) return 1;
            if (!aIsEmpty && bIsEmpty) return -1;
            if (aIsEmpty && bIsEmpty) return 0;

            const isNumeric = !isNaN(parseFloat(cellA)) && isFinite(cellA) && !isNaN(parseFloat(cellB)) && isFinite(cellB);

            let valA = isNumeric ? parseFloat(cellA) : cellA.toLowerCase();
            let valB = isNumeric ? parseFloat(cellB) : cellB.toLowerCase();

            if (valA < valB) return newSortDir === 'asc' ? -1 : 1;
            if (valA > valB) return newSortDir === 'asc' ? 1 : -1;
            return 0;
        });

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
    let name = sheetName;
    if (!name) {
      const existingNumbers = Object.keys(this.sheets).map(s => {
        const match = s.match(/^Sheet(\d+)$/);
        return match ? parseInt(match[1]) : 0;
      });
      const maxNumber = existingNumbers.length > 0 ? Math.max(...existingNumbers) : 0;
      name = `Sheet${maxNumber + 1}`;
    }

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
    
    // Start the loop from the second row (index 1) to skip the column headers row
    for (let i = 1; i < rows.length; i++) {
        let row = [], cols = rows[i].querySelectorAll("td, th");
        // Start the loop from the second cell (index 1) to skip the row header
        for (let j = 1; j < cols.length; j++) {
            let cellValue = "";
            if (cols[j].tagName === "TD") {
                cellValue = cols[j].querySelector("input").dataset.formula || cols[j].querySelector("input").value;
            } else if (cols[j].tagName === "TH") {
                cellValue = cols[j].innerText.trim();
                cellValue = cellValue.replace(/\s*$/, '');
            }
            row.push(`"${cellValue.replace(/"/g, '""')}"`);
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

  uploadCSV() {
    const fileInput = document.getElementById("fileInput");
    fileInput.click();

    fileInput.addEventListener("change", (event) => {
        const file = event.target.files[0];
        if (!file) return;

        const reader = new FileReader();
        reader.onload = (e) => {
            const fileContent = e.target.result;
            const parsedData = this.parseCSV(fileContent);

            if (parsedData.length > 0) {
                // Adjust table dimensions based on parsed data
                this.spreadsheet.numRows = parsedData.length;
                this.spreadsheet.numCols = parsedData[0].length;

                // Re-create the table with new dimensions
                this.spreadsheet.tableManager.createTable();

                // Load the entire parsed data into the table cells
                this.spreadsheet.tableManager.loadDataFromCSV(parsedData);
                
                this.spreadsheet.saveToLocalStorage();
            }
        };
        reader.readAsText(file);
    });
  }

  parseCSV(text) {
    const rows = text.split(/\r?\n/).filter(line => line.trim() !== "");
    const data = [];
    rows.forEach(row => {
        let cells = [];
        let inQuote = false;
        let currentCell = '';
        for (let i = 0; i < row.length; i++) {
            const char = row[i];
            const nextChar = row[i + 1];
            const previousChar = row[i - 1];

            if (char === '"') {
                if (inQuote && nextChar === '"') {
                    // Escaped quote
                    currentCell += '"';
                    i++;
                } else {
                    // Start or end of a quote
                    inQuote = !inQuote;
                }
            } else if (char === ',' && !inQuote) {
                cells.push({
                    value: currentCell,
                    width: "",
                    height: "",
                    formula: currentCell.startsWith('=') ? currentCell : ''
                });
                currentCell = '';
            } else {
                currentCell += char;
            }
        }
        cells.push({
            value: currentCell,
            width: "",
            height: "",
            formula: currentCell.startsWith('=') ? currentCell : ''
        });
        data.push(cells);
    });
    return data;
  }
}

class Spreadsheet {
  constructor(numRows, numCols, tableId = 'spreadsheet') {
    this.numRows = numRows;
    this.numCols = numCols;
    this.table = document.getElementById(tableId);
    this.isFormulaEditMode = false;
    this.activeFormulaCell = null;
    this.isDragging = false;
    this.dragStartCell = null;

    this.tableManager = new TableManager(this);
    this.stateManager = new StateManager(this);
    this.lintingManager = new LintingManager(this);
    this.contextMenuManager = new ContextMenuManager(this);
    this.sheetManager = new SheetManager(this);
    this.downloadPrintManager = new DownloadPrintManager(this);
    this.formulaManager = new FormulaManager(this);
    this.apiKeyManager = new ApiKeyManager(); // Add this line
    this.aiIntegration = new SpreadsheetAIIntegration(this); // Add this line

    this.addGlobalListeners();
    this.addCellListeners();
    this.addAutofillListeners();

    const storedState = StorageManager.getFromLocalStorage('spreadsheetState');
    if (storedState && storedState.darkMode) {
      document.body.classList.add('dark-mode');
      document.getElementById("darkModeToggle").checked = true;
    }

    if (!storedState || Object.keys(storedState.sheets).length === 0) {
      this.sheetManager.addSheet();
    } else {
      this.sheetManager.loadSheets(storedState.sheets, storedState.activeSheet);
    }
    this.recalculateAllFormulas();
  }

  addGlobalListeners() {
    document.getElementById("clearBtn").addEventListener("click", () => this.tableManager.clearTable());
    document.getElementById("downloadBtn").addEventListener("click", () => this.downloadPrintManager.downloadAsCSV());
    document.getElementById("printBtn").addEventListener("click", () => this.downloadPrintManager.printTable());
    document.getElementById("uploadBtn").addEventListener("click", () => this.downloadPrintManager.uploadCSV());
    document.getElementById("darkModeToggle").addEventListener("change", (e) => this.toggleDarkMode(e.target.checked));
    document.addEventListener("keydown", (e) => this.stateManager.handleKeyboardShortcuts(e));

    // Listen for mousedown on the entire document to finalize formula when clicking outside the table
    document.addEventListener("mousedown", (e) => {
      const isClickInsideTable = this.table.contains(e.target);
      if (this.isFormulaEditMode && !isClickInsideTable) {
        this.finalizeFormula();
      }
    });

    // New mousedown listener to handle formula insertion. This fires before the blur/focusin events.
    this.table.addEventListener("mousedown", (event) => {
        const cell = event.target.closest('td');
        if (this.isFormulaEditMode && cell && cell !== this.activeFormulaCell) {
            console.log("Mousedown on cell while in formula mode. Inserting reference...");
            const row = cell.parentElement.rowIndex - 1;
            const col = cell.cellIndex - 1;
            const cellId = this.getColumnName(col + 1) + (row + 1);
            console.log("Cell ID to be inserted:", cellId);

            const activeInput = this.activeFormulaCell.querySelector('input');
            const start = activeInput.selectionStart;
            const end = activeInput.selectionEnd;
            activeInput.value = activeInput.value.substring(0, start) + cellId + activeInput.value.substring(end);
            activeInput.focus();
            activeInput.setSelectionRange(start + cellId.length, start + cellId.length);
            this.highlightFormulaReferences();
        }
    });
  }

  addCellListeners() {
    this.table.addEventListener("input", (event) => {
        const input = event.target;
        if (input.tagName === 'INPUT') {
            if (input.value.startsWith('=')) {
                this.isFormulaEditMode = true;
                this.activeFormulaCell = input.closest('td');
                this.highlightFormulaReferences();
            } else {
                this.isFormulaEditMode = false;
                this.activeFormulaCell = null;
                this.clearFormulaHighlights();
            }
        }
    });

    this.table.addEventListener("focusin", (event) => {
        const input = event.target;
        if (input.tagName === 'INPUT') {
            if (input.dataset.formula) {
                input.value = input.dataset.formula;
            }
            if (input.value.startsWith('=')) {
                this.isFormulaEditMode = true;
                this.activeFormulaCell = input.closest('td');
                this.highlightFormulaReferences();
            } else {
                 this.isFormulaEditMode = false;
                 this.activeFormulaCell = null;
            }
        }
    });

    this.table.addEventListener("keydown", (event) => {
        const input = event.target;
        if (input.tagName === 'INPUT' && event.key === 'Enter') {
            this.finalizeFormula();
            input.blur();
        } else if (this.isFormulaEditMode && ['ArrowLeft', 'ArrowRight', 'ArrowUp', 'ArrowDown'].includes(event.key)) {
            event.stopPropagation();
        }
    }, true);
  }
  
  addAutofillListeners() {
    this.table.addEventListener("mousedown", (event) => {
        const handle = event.target.closest(".autofill-handle");
        if (handle) {
            event.preventDefault();
            this.isDragging = true;
            this.dragStartCell = handle.closest("td");
            document.body.classList.add("dragging");
        }
    });

    this.table.addEventListener("mouseover", (event) => {
        const cell = event.target.closest("td");
        if (this.isDragging && cell) {
            const startRow = this.dragStartCell.parentElement.rowIndex;
            const startCol = this.dragStartCell.cellIndex;
            const endRow = cell.parentElement.rowIndex;
            const endCol = cell.cellIndex;

            this.highlightAutofillRange(startRow, startCol, endRow, endCol);
        }
    });

    document.addEventListener("mouseup", (event) => {
        if (this.isDragging) {
            this.isDragging = false;
            document.body.classList.remove("dragging");
            const startCell = this.dragStartCell;
            const endCell = document.elementFromPoint(event.clientX, event.clientY).closest('td');

            if (startCell && endCell) {
                this.autofill(startCell, endCell);
            }
            this.dragStartCell = null;
            this.clearAutofillHighlights();
        }
    });
  }

  highlightAutofillRange(startRow, startCol, endRow, endCol) {
    this.clearAutofillHighlights();
    const minRow = Math.min(startRow, endRow);
    const maxRow = Math.max(startRow, endRow);
    const minCol = Math.min(startCol, endCol);
    const maxCol = Math.max(startCol, endCol);

    for (let r = minRow; r <= maxRow; r++) {
      const rowEl = this.table.rows[r];
      if (rowEl) {
        for (let c = minCol; c <= maxCol; c++) {
          const cellEl = rowEl.cells[c];
          if (cellEl) {
            cellEl.classList.add('autofill-highlight');
          }
        }
      }
    }
  }

  clearAutofillHighlights() {
    this.table.querySelectorAll('.autofill-highlight').forEach(cell => {
      cell.classList.remove('autofill-highlight');
    });
  }

  autofill(startCell, endCell) {
    const startRow = startCell.parentElement.rowIndex;
    const startCol = startCell.cellIndex;
    const endRow = endCell.parentElement.rowIndex;
    const endCol = endCell.cellIndex;

    const isHorizontal = startRow === endRow;

    const sourceInput = startCell.querySelector('input');
    const sourceFormula = sourceInput.dataset.formula;
    const sourceValue = sourceInput.value;

    const minRow = Math.min(startRow, endRow);
    const maxRow = Math.max(startRow, endRow);
    const minCol = Math.min(startCol, endCol);
    const maxCol = Math.max(startCol, endCol);

    let startVal, nextVal, step = 1;
    let textPrefix = '';
    let isNumberSeries = false;
    let isAlphaSeries = false;

    if (isHorizontal && minCol !== maxCol) {
        const nextCell = this.table.rows[startRow].cells[startCol + 1];
        startVal = sourceInput.value;
        nextVal = nextCell?.querySelector('input')?.value;
    } else if (!isHorizontal && minRow !== maxRow) {
        const nextCell = this.table.rows[startRow + 1].cells[startCol];
        startVal = sourceInput.value;
        nextVal = nextCell?.querySelector('input')?.value;
    } else {
        startVal = sourceInput.value;
        nextVal = null;
    }
    
    // Check for number series
    const startNumber = parseFloat(startVal);
    const nextNumber = parseFloat(nextVal);
    
    if (!isNaN(startNumber)) {
        if (!isNaN(nextNumber)) {
            step = nextNumber - startNumber;
        }
        isNumberSeries = true;
    } else {
        const match = sourceValue.match(/^(.*?)(\d+)$/);
        if (match) {
            textPrefix = match[1];
            startVal = parseInt(match[2], 10);
            
            const nextMatch = nextVal?.match(/^(.*?)(\d+)$/);
            if (nextMatch && nextMatch[1] === textPrefix) {
                nextVal = parseInt(nextMatch[2], 10);
                step = nextVal - startVal;
            } else {
                step = 1;
            }
            isNumberSeries = true;
        } else if (sourceValue.length === 1 && sourceValue.match(/[a-zA-Z]/)) {
            isAlphaSeries = true;
        }
    }
    
    for (let r = minRow; r <= maxRow; r++) {
        for (let c = minCol; c <= maxCol; c++) {
            if (r === startRow && c === startCol) continue;

            const targetCell = this.table.rows[r].cells[c];
            if (!targetCell) continue;

            const targetInput = targetCell.querySelector('input');
            const rowDelta = r - startRow;
            const colDelta = c - startCol;

            // Handle formulas
            if (sourceFormula) {
                const newFormula = sourceFormula.replace(/[A-Z]+[0-9]+/g, (match) => {
                    const colName = match.match(/[A-Z]+/)[0];
                    const rowNum = parseInt(match.match(/[0-9]+/)[0], 10);

                    const originalColIndex = this.columnNameToIndex(colName);
                    const newColIndex = originalColIndex + colDelta;
                    const newColName = this.getColumnName(newColIndex + 1);

                    const newRowNum = rowNum + rowDelta;

                    return newColName + newRowNum;
                });
                targetInput.dataset.formula = newFormula;
                targetInput.value = this.formulaManager.evaluateFormula(newFormula);
            } else if (isNumberSeries) {
                const currentDelta = isHorizontal ? colDelta : rowDelta;
                const newValue = (parseFloat(startNumber) || startVal) + (currentDelta * step);
                if (textPrefix) {
                    targetInput.value = textPrefix + newValue;
                } else {
                    targetInput.value = newValue;
                }
                targetInput.dataset.formula = '';
            } else if (isAlphaSeries) {
                const currentDelta = isHorizontal ? colDelta : rowDelta;
                const charCode = sourceValue.charCodeAt(0);
                const newCharCode = charCode + currentDelta;
                const newValue = String.fromCharCode(newCharCode);
                targetInput.value = newValue;
                targetInput.dataset.formula = '';
            } else {
                // Handle non-formulas and non-numbers (text)
                targetInput.value = sourceValue;
                targetInput.dataset.formula = '';
            }
        }
    }
    this.recalculateAllFormulas();
    this.stateManager.saveState();
  }

  finalizeFormula() {
    if (this.isFormulaEditMode && this.activeFormulaCell) {
        console.log("Finalizing formula...");
        const input = this.activeFormulaCell.querySelector('input');
        input.dataset.formula = input.value;
        this.recalculateAllFormulas();
        this.isFormulaEditMode = false;
        this.activeFormulaCell = null;
        this.clearFormulaHighlights();
    }
  }

  recalculateAllFormulas() {
      const allInputs = this.table.querySelectorAll('input');
      allInputs.forEach(input => {
          if (input.dataset.formula) {
              const result = this.formulaManager.evaluateFormula(input.dataset.formula);
              input.value = result;
          }
      });
  }

  highlightFormulaReferences() {
      this.clearFormulaHighlights();
      if (!this.activeFormulaCell || !this.isFormulaEditMode) return;

      const formula = this.activeFormulaCell.querySelector('input').value;
      const cellReferences = [...new Set(formula.match(/[A-Z]+[0-9]+/g))];

      if (cellReferences) {
          cellReferences.forEach((ref, index) => {
              const colName = ref.match(/[A-Z]+/)[0];
              const rowNum = parseInt(ref.match(/[0-9]+/)[0], 10);
              const colIndex = this.columnNameToIndex(colName);
              const cell = this.table.querySelector(`tr:nth-child(${rowNum + 1}) td:nth-child(${colIndex + 2})`);
              if (cell) {
                  cell.classList.add('highlight-formula-ref', `color-${index % 7}`);
              }
          });
      }
  }

  clearFormulaHighlights() {
      document.querySelectorAll('.highlight-formula-ref').forEach(el => {
          el.classList.remove('highlight-formula-ref');
          el.classList.remove(...Array.from({length: 7}, (_, i) => `color-${i}`));
      });
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

  columnNameToIndex(columnName) {
    let index = 0;
    for (let i = 0; i < columnName.length; i++) {
        index = index * 26 + (columnName.charCodeAt(i) - 64);
    }
    return index - 1;
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