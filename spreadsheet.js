// Define default values for rows and columns
let numRows = 25; // Set initial number of rows
let numCols = 25; // Set initial number of columns

const table = document.getElementById("spreadsheet");

// Function to create the table
function createTable() {
  table.innerHTML = ""; // Clear any existing table content

  // Create column headers (A, B, C, D, ...)
  const headerRow = document.createElement("tr");
  headerRow.appendChild(document.createElement("th")); // Top-left corner empty cell
  for (let col = 0; col < numCols; col++) {
    const th = document.createElement("th");
    th.textContent = String.fromCharCode(65 + col); // ASCII for A, B, C, ...
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
});

// Function to add a new column
document.getElementById("addColBtn").addEventListener("click", function () {
  numCols++;  // Increase the number of columns

  // Update the column headers
  const headerRow = table.querySelector("tr");
  const th = document.createElement("th");
  th.textContent = String.fromCharCode(65 + numCols - 1); // Update column letter
  headerRow.appendChild(th);

  // Add new cells for each row
  const rows = table.querySelectorAll("tr");
  rows.forEach((row, rowIndex) => {
    if (rowIndex === 0) return;  // Skip the header row
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

    row.appendChild(td);

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
  });
});

