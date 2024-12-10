<p align="center">
  <img src="https://raw.githubusercontent.com/arunwebber/simpleExcel/refs/heads/main/images/icon_128.png" alt="Logo">
</p>
# Simple Browser-Based Excel Sheet Extension

This is a simple Chrome extension that mimics basic functionality of an Excel sheet. It allows users to interact with a table and input data, providing a lightweight, browser-based spreadsheet experience.

## Features

- A basic table grid layout for data input.
- Fully responsive and works across modern browsers.
- Designed with simplicity in mind for quick data entry and management.

## Installation Instructions

### 1. Clone or Download the Repository

Clone or download this repository to your local machine.

### 2. Prepare the Extension Files

Ensure the following files are present in the project directory:

- `index.html`: The main HTML structure of the application.
- `style.css`: The CSS file for styling the layout and appearance.
- `spreadsheet.js`: The JavaScript file containing the logic for creating and managing the spreadsheet.

### 3. Load the Extension in Chrome

To install the extension in Chrome, follow these steps:

1. Open **Chrome** and go to `chrome://extensions/`.
2. Turn on **Developer Mode** by toggling the switch at the top-right of the Extensions page.
3. Click the **Load unpacked** button.
4. Navigate to the directory where you've saved the extension files (containing `index.html`, `style.css`, and `spreadsheet.js`) and select it.
5. The extension should now be installed and visible on your Chrome Extensions page.

### 4. Use the Extension

Once the extension is installed:

1. You can access it from the Chrome Extensions menu or pin it to the toolbar.
2. Click the extension icon to open the browser-based Excel sheet in a new tab.
3. You can start typing directly into the table cells, and the table is dynamic based on your input.

### Structure

- **index.html**: Contains the HTML structure for the application.
- **style.css**: Contains the styles for layout and visual appearance.
- **spreadsheet.js**: Implements the logic for rendering the table and managing cell data.


## Future Enhancements

- Adding support for saving and loading files.
- Improving the user interface with more features such as cell formatting, formulas, etc.



## License

This project is open-source and available under the MIT License. See the [LICENSE](LICENSE) file for more information.

## Customization for Firefox

To make the extension compatible with Firefox, replace the contents of your `manifest.json` file with the following:

```json
{
  "manifest_version": 3,
  "name": "Simple Spreadsheet Extension",
  "version": "1.1",
  "description": "Opens a new tab with a simple spreadsheet interface.",
  "action": {
    "default_title": "Open Spreadsheet",
    "default_icon": {
      "16": "images/icon_16.png",
      "48": "images/icon_48.png",
      "128": "images/icon_128.png"
    }
  },
  "background": {
    "service_worker": "background.js"
  },
  "icons": {
    "16": "images/icon_16.png",
    "48": "images/icon_48.png",
    "128": "images/icon_128.png"
  },
  "homepage_url": "https://www.arunsyoga.in",
  "author": "HashPalLabs"
}
