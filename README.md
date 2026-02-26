# Atoss Auto Read

Automatic import of **ATOSS data from an Excel file** using a browser userscript.

---

# Installation

## 1. Install Tampermonkey

Install the browser extension **Tampermonkey** from your browser’s extension store.

Tested browsers:

- Microsoft Edge  
- Firefox  
- Vivaldi  

Tampermonkey allows custom scripts to modify and automate the ATOSS web interface directly in the browser.

---

## 2. Install the User Scripts

1. Open the **Tampermonkey Dashboard**.
2. Navigate to **Utilities**.
3. Import the following scripts from the GitHub repository:

   - `atoss-auto-read-main.user.js`
   - `atoss-auto-read-sub.user.js`

---

# Scripts

## Main Script

`atoss-auto-read-main.user.js`

This script provides the following functionality:

- Prevents the scrollbar from jumping when the buttons **Korrektur** or **Verteilen** are clicked.
- Adds an **Import button** to the header of the ATOSS user interface.
- After selecting a file, iterates over all rows of the dataset.
- Passes each row to the `atoss-auto-read-sub.user.js` script for processing.

---

## Sub Script

`atoss-auto-read-sub.user.js`

This script:

- Receives row data from the main script.
- Automatically fills in the relevant fields in the ATOSS interface.
- Simulates user interaction (mouse clicks).
- Saves the entry.

---

# Data Format

The import file must be an **Excel sheet** containing the following columns:

- **Date**  
  - Tested formats: `dd.mm.yyyy` and `yyyy.mm.dd`
- **Leistungsart**
- **Empfaenger**
- **Vorgang** (text)
- **Arbeitsplatz**
- **Stunden** (text)
- **Kommentar**

---

# Usage

1. Click the **Import button** in the ATOSS interface.
2. Select the Excel file containing the data.
3. The script will iterate over each row of the file.

For each row:

- A new window is opened.
- The ATOSS daily view is loaded.
- The required fields are filled automatically.
- The **Save** button is triggered programmatically.
- The window closes automatically.

After the import process is complete, you may refresh the ATOSS main page to view the updated entries.

---

# Troubleshooting

- Some browsers block the `atoss-auto-read-sub.user.js` script. The user has to allow the external pop-ups.
