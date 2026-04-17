// ==UserScript==
// @name         atoss main: read from excel
// @namespace    http://tampermonkey.net/
// @version      1.03
// @description  try to take over the world!
// @author       You
// @match        https://he-atoss.horiba.eu:5000/?date=*
// @icon         data:image/gif;base64,R0lGODlhAQABAAAAACH5BAEKAAEALAAAAAABAAEAAAICTAEAOw==
// @require      https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js
// @grant        none
// ==/UserScript==


// change the name of the site in the head
const navbarBrand = document.querySelector('.navbar-brand');
if (navbarBrand) {
    navbarBrand.textContent = "HORIBA Stundenverteilung (Script Modification)";
}


// save position of scroll bar
function saveScrollPosition() {
    // Select the first element with the class 'preserveDiv'
    const preserveDiv = document.querySelector('.tableFixHeadIndex');
    if (preserveDiv) {
        localStorage.setItem('preservedScrollPosition', preserveDiv.scrollTop);
    }
}

// Restore the scroll position after page load
const preserveDiv = document.querySelector('.tableFixHeadIndex');
const savedScrollPosition = localStorage.getItem('preservedScrollPosition');
preserveDiv.scrollTop = parseInt(savedScrollPosition, 10);

// Add event listener to all "Korrektur" buttons
document.querySelectorAll('a.btn.btn-light, a.btn.btn-outline-secondary').forEach(button => {
    button.addEventListener('click', function () {
        // Save the scroll position before the page redirects
        saveScrollPosition();
    });
});


const sheetData = JSON.parse(localStorage.getItem('excelData'));

if (sheetData) {
    console.log("Storage data exists!")
    let numOfEntries = sheetData.length;
    console.log("Length of sheetData <<", numOfEntries);
    let i = parseFloat(localStorage.getItem('iterator'));
    console.log("iterator <<", i)
    let ii = numOfEntries - i
    i = i + 1
    localStorage.setItem('iterator', i);
    if (ii >= 0) {
        //localStorage.removeItem('excelData'); // removes only 'excelData'
        if (window.opener) {
            window.close(); // only works if opened via script
        }
    } else {
        localStorage.removeItem('excelData');
        localStorage.removeItem('iterator');
    }
}

function excelDateToJSDate(excelDate) {
    const excelEpoch = new Date(1900, 0, 1);  // Excel starts counting from Jan 1, 1900
    excelEpoch.setDate(excelEpoch.getDate() + excelDate - 2);  // Adjust for Excel date system
    const day = String(excelEpoch.getDate()).padStart(2, '0'); // Get the day, padded to two digits
    const month = String(excelEpoch.getMonth() + 1).padStart(2, '0'); // Get the month (0-indexed, so add 1), padded to two digits
    const year = excelEpoch.getFullYear(); // Get the full year
    return `${day}.${month}.${year}`;
}

function collectDataFromOverviewTable() {
    const result = [];
    const rows = document.querySelectorAll('table tbody tr');

    rows.forEach((row) => {
        const tds = row.querySelectorAll('td');
        if (!tds.length) return;

        // Date cells contain a day prefix like "Mi, 15.04.2026" — extract dd.mm.yyyy
        const rawDateText = (tds[0].textContent || '').trim();
        const dateMatch = rawDateText.match(/(\d{2}\.\d{2}\.\d{4})/);
        if (!dateMatch) return;
        const dateText = dateMatch[1];

        // td[3] is always "Verteilbare Stunden" regardless of which tds have .text-center
        if (tds.length < 4) return;
        const timeText = (tds[3].textContent || '').trim().replace(/\s+/g, ' ');
        if (!timeText || timeText === "0,00") return;

        // Keep 0,00 as well, as requested
        result.push({
            Date: dateText,
            Stunden: timeText,
            Project_name: "Gemeinkosten",
            Empfaenger: "8000002561",
            Leistungsart: "114012",
            Vorgang: "0040",
            Arbeitsplatz: "D330000",
            Kommentar: "leftover"
        });
    });


    const preview = ['Date; Stunden; Project_name; Empfaenger; Leistungsart; Vorgang; Arbeitsplatz, Kommentar', ...result.map(r => `${r.Date}; ${r.Stunden}; ${r.Project_name}; ${r.Empfaenger}; ${r.Leistungsart}; ${r.Vorgang}; ${r.Arbeitsplatz}; ${r.Kommentar}`)].join('\n');
    console.log(preview);

    return result;

}

function showEntriesPreview(entries) {
    return new Promise((resolve) => {
        const overlay = document.createElement('div');
        overlay.style.position = 'fixed';
        overlay.style.inset = '0';
        overlay.style.backgroundColor = 'rgba(0, 0, 0, 0.45)';
        overlay.style.zIndex = '999999';
        overlay.style.display = 'flex';
        overlay.style.alignItems = 'center';
        overlay.style.justifyContent = 'center';

        const dialog = document.createElement('div');
        dialog.style.background = '#fff';
        dialog.style.width = 'min(1100px, 92vw)';
        dialog.style.maxHeight = '85vh';
        dialog.style.display = 'flex';
        dialog.style.flexDirection = 'column';
        dialog.style.borderRadius = '8px';
        dialog.style.boxShadow = '0 12px 30px rgba(0, 0, 0, 0.35)';
        dialog.style.overflow = 'hidden';

        const header = document.createElement('div');
        header.style.padding = '14px 16px';
        header.style.borderBottom = '1px solid #ddd';
        header.style.fontWeight = '600';
        header.textContent = `Preview import (${entries.length} entries)`;

        const body = document.createElement('div');
        body.style.padding = '12px 16px';
        body.style.overflow = 'auto';

        const table = document.createElement('table');
        table.style.width = '100%';
        table.style.borderCollapse = 'collapse';
        table.style.fontSize = '13px';

        const columns = ['Date', 'Stunden', 'Project_name', 'Empfaenger', 'Leistungsart', 'Vorgang', 'Arbeitsplatz', 'Kommentar'];

        const thead = document.createElement('thead');
        const headRow = document.createElement('tr');
        columns.forEach((columnName) => {
            const th = document.createElement('th');
            th.textContent = columnName;
            th.style.position = 'sticky';
            th.style.top = '0';
            th.style.background = '#f4f4f4';
            th.style.borderBottom = '1px solid #ddd';
            th.style.padding = '7px 8px';
            th.style.textAlign = 'left';
            headRow.appendChild(th);
        });
        thead.appendChild(headRow);

        const tbody = document.createElement('tbody');
        entries.forEach((entry) => {
            const row = document.createElement('tr');
            columns.forEach((columnName) => {
                const td = document.createElement('td');
                td.textContent = (entry[columnName] || '').toString();
                td.style.borderBottom = '1px solid #eee';
                td.style.padding = '6px 8px';
                row.appendChild(td);
            });
            tbody.appendChild(row);
        });

        table.appendChild(thead);
        table.appendChild(tbody);
        body.appendChild(table);

        const footer = document.createElement('div');
        footer.style.display = 'flex';
        footer.style.justifyContent = 'flex-end';
        footer.style.gap = '8px';
        footer.style.padding = '12px 16px';
        footer.style.borderTop = '1px solid #ddd';

        const cancelButton = document.createElement('button');
        cancelButton.type = 'button';
        cancelButton.textContent = 'Cancel';
        cancelButton.style.padding = '6px 12px';
        cancelButton.style.cursor = 'pointer';

        const okButton = document.createElement('button');
        okButton.type = 'button';
        okButton.textContent = 'OK';
        okButton.style.padding = '6px 12px';
        okButton.style.cursor = 'pointer';

        const closeDialog = (result) => {
            document.removeEventListener('keydown', onKeyDown);
            overlay.remove();
            resolve(result);
        };

        const onKeyDown = (event) => {
            if (event.key === 'Escape') {
                closeDialog(false);
            }
        };

        cancelButton.addEventListener('click', () => closeDialog(false));
        okButton.addEventListener('click', () => closeDialog(true));
        document.addEventListener('keydown', onKeyDown);

        footer.appendChild(cancelButton);
        footer.appendChild(okButton);

        dialog.appendChild(header);
        dialog.appendChild(body);
        dialog.appendChild(footer);
        overlay.appendChild(dialog);
        document.body.appendChild(overlay);
    });
}


// Create a file input dynamically
const fileInput = document.createElement('input');
fileInput.type = 'file';
fileInput.accept = '.xlsx';
fileInput.style.position = 'fixed';
fileInput.style.top = '10px';
fileInput.style.left = '600px';
document.body.appendChild(fileInput);

// New button: import directly from visible table
const tableImportButton = document.createElement('button');
tableImportButton.type = 'button';
tableImportButton.textContent = 'Import leftover hours';
tableImportButton.style.position = 'fixed';
tableImportButton.style.top = '10px';
tableImportButton.style.left = '760px';
tableImportButton.style.zIndex = '99999';
tableImportButton.style.padding = '3px 8px';
tableImportButton.style.cursor = 'pointer';
document.body.appendChild(tableImportButton);

tableImportButton.addEventListener('click', async () => {
    const tableData = collectDataFromOverviewTable();
    if (!tableData.length) {
        alert("No valid rows found in table.");
        return;
    }

    const confirmed = await showEntriesPreview(tableData);
    if (!confirmed) {
        console.log('Import canceled by user.');
        return;
    }

    // dummy for-loop
    let i = 0;
    const interval = setInterval(() => {
        // retrieve old values
        localStorage.setItem('excelData', JSON.stringify(tableData));
        localStorage.setItem('index', JSON.stringify(i));
        localStorage.setItem('iterator', 1);
        // console.log(excelDateToJSDate(tableData[i].Date))

        // open each entry in new window — Date is already dd.mm.yyyy from the table
        window.open('https://he-atoss.horiba.eu:5000/Home/AddOrEdit?date=' + tableData[i].Date, '_blank');
        i++;

        // if all entries are done, clear the intervall / dummy for-loop
        if (i > tableData.length) {
            clearInterval(interval);
            // Clean up localStorage so old data doesn't persist when manually opening pages
            setTimeout(() => {
                localStorage.removeItem('excelData');
                localStorage.removeItem('iterator');
                localStorage.removeItem('index');
            }, 6000); // Wait for 1000ms (1 seconds) before cleanup
        }
    }, 4000); // Wait for 1000ms (1 seconds) between each iteration

});

// Handle file selection
fileInput.addEventListener('change', (event) => {
    const file = event.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (e) => {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });

        // Parse the first sheet
        const sheetName = workbook.SheetNames[0];
        const sheetData = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName]);
        console.log('Sheet Data:', sheetData);

        // dummy for-loop
        let i = 0;
        const interval = setInterval(() => {
            // retrieve old values
            localStorage.setItem('excelData', JSON.stringify(sheetData));
            localStorage.setItem('index', JSON.stringify(i));
            localStorage.setItem('iterator', 1);
            // console.log(excelDateToJSDate(sheetData[i].Date))

            // open each entry in new window
            window.open('https://he-atoss.horiba.eu:5000/Home/AddOrEdit?date=' + excelDateToJSDate(sheetData[i].Date), '_blank');
            i++;

            // if all entries are done, clear the intervall / dummy for-loop
            if (i > sheetData.length) {
                clearInterval(interval);
                // Clean up localStorage so old data doesn't persist when manually opening pages
                setTimeout(() => {
                    localStorage.removeItem('excelData');
                    localStorage.removeItem('iterator');
                    localStorage.removeItem('index');
                }, 6000); // Wait for 1000ms (1 seconds) before cleanup
            }
        }, 4000); // Wait for 1000ms (1 seconds) between each iteration



        // Store the data in localStorage
        //localStorage.setItem('excelData', JSON.stringify(sheetData[0]));
        //console.log(excelDateToJSDate(sheetData[0].Date))
        //window.open('https://he-atoss.horiba.eu:5000/Home/AddOrEdit?date=' + excelDateToJSDate(sheetData[0].Date), '_blank');
    };

    reader.readAsArrayBuffer(file);
});
