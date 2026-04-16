// ==UserScript==
// @name atoss main: read from excel
// @namespace http://tampermonkey.net/
// @version 1.04
// @description Import from Excel and directly from table data
// @author marius.teppe@horiba.com
// @match https://he-atoss.horiba.eu:5000/?date=*
// @icon data:image/gif;base64,R0lGODlhAQABAAAAACH5BAEKAAEALAAAAAABAAEAAAICTAEAOw==
// @require https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js
// @grant none
// ==/UserScript==

// change the name of the site in the head
const navbarBrand = document.querySelector('.navbar-brand');
if (navbarBrand) {
        navbarBrand.textContent = "HORIBA Stundenverteilung (Script Modification)";
}

// save position of scroll bar
function saveScrollPosition() {
        const preserveDiv = document.querySelector('.tableFixHeadIndex');
        if (preserveDiv) {
                localStorage.setItem('preservedScrollPosition', preserveDiv.scrollTop);
        }
}

// Restore the scroll position after page load
const preserveDiv = document.querySelector('.tableFixHeadIndex');
const savedScrollPosition = localStorage.getItem('preservedScrollPosition');
if (preserveDiv && savedScrollPosition !== null) {
        preserveDiv.scrollTop = parseInt(savedScrollPosition, 10) || 0;
}

// Add event listener to all "Korrektur" buttons
document.querySelectorAll('a.btn.btn-light, a.btn.btn-outline-secondary').forEach(button => {
        button.addEventListener('click', function () {
                saveScrollPosition();
        });
});

const sheetData = JSON.parse(localStorage.getItem('excelData'));

if (sheetData) {
        console.log("Storage data exists!");
        let numOfEntries = sheetData.length;
        console.log("Length of sheetData <<", numOfEntries);
        let i = parseFloat(localStorage.getItem('iterator'));
        console.log("iterator <<", i);
        let ii = numOfEntries - i;
        i = i + 1;
        localStorage.setItem('iterator', i);
        if (ii >= 0) {
                if (window.opener) {
                        window.close();
                }
        } else {
                localStorage.removeItem('excelData');
                localStorage.removeItem('iterator');
        }
}

function excelDateToJSDate(excelDate) {
        const excelEpoch = new Date(1900, 0, 1);
        excelEpoch.setDate(excelEpoch.getDate() + excelDate - 2);
        const day = String(excelEpoch.getDate()).padStart(2, '0');
        const month = String(excelEpoch.getMonth() + 1).padStart(2, '0');
        const year = excelEpoch.getFullYear();
        return `${day}.${month}.${year}`;
}

function normalizeDateForUrl(dateValue) {
        if (typeof dateValue === 'number') {
                return excelDateToJSDate(dateValue);
        }

        if (dateValue instanceof Date && !isNaN(dateValue)) {
                const day = String(dateValue.getDate()).padStart(2, '0');
                const month = String(dateValue.getMonth() + 1).padStart(2, '0');
                const year = dateValue.getFullYear();
                return `${day}.${month}.${year}`;
        }

        if (typeof dateValue === 'string') {
                const value = dateValue.trim();

                // already in dd.mm.yyyy
                if (/^\d{2}\.\d{2}\.\d{4}$/.test(value)) {
                        return value;
                }

                // convert yyyy-mm-dd to dd.mm.yyyy
                const isoMatch = value.match(/^(\d{4})-(\d{2})-(\d{2})$/);
                if (isoMatch) {
                        return `${isoMatch[3]}.${isoMatch[2]}.${isoMatch[1]}`;
                }
        }

        return "";
}

function startAtossImport(rows) {
        if (!Array.isArray(rows) || rows.length === 0) {
                alert("No rows to import.");
                return;
        }
        localStorage.setItem('excelData', JSON.stringify(rows));
        localStorage.setItem('iterator', 1);

        let i = 0;
        const interval = setInterval(() => {
                const row = rows[i];
                const dateForUrl = normalizeDateForUrl(row.Date);

                if (!dateForUrl) {
                        console.warn("Skipping row with invalid date:", row);
                        i++;
                        if (i >= rows.length) clearInterval(interval);
                        return;
                }

                localStorage.setItem('index', JSON.stringify(i));
                window.open(
                        'https://he-atoss.horiba.eu:5000/Home/AddOrEdit?date=' + encodeURIComponent(dateForUrl),
                        '_blank'
                );

                i++;
                if (i >= rows.length) {
                        clearInterval(interval);
                }
        }, 4000);

}

function collectDataFromOverviewTable() {
        const result = [];
        const rows = document.querySelectorAll('table tbody tr');

        rows.forEach((row) => {
                const tds = row.querySelectorAll('td');
                if (!tds.length) return;

                const dateText = (tds[0].textContent || '').trim();
                if (!/^\d{2}\.\d{2}\.\d{4}$/.test(dateText)) return;

                const centers = row.querySelectorAll('.text-center');
                const thirdCenter = centers[2];
                if (!thirdCenter) return;

                const timeText = (thirdCenter.textContent || '').trim().replace(/\s+/g, ' ');
                if (!timeText) return;

                // Keep 0,00 as well, as requested
                result.push({
                        Date: dateText,
                        Stunden: timeText,
                        Project_name: "Gemeinkosten",
                        Empfaenger: "8001233873",
                        Leistungsart: "114012",
                        Vorgang: "0130",
                        Arbeitsplatz: "DD330020"
                });
        });

        const preview = ['Date; Stunden; Project_name; Empfaenger; Leistungsart; Vorgang; Arbeitsplatz', ...result.map(r => `${r.Date}; ${r.Stunden}; ${r.Project_name}; ${r.Empfaenger}; ${r.Leistungsart}; ${r.Vorgang}; ${r.Arbeitsplatz}`)].join('\n');
        console.log(preview);

        return result;

}

// Create Excel file input dynamically
const fileInput = document.createElement('input');
fileInput.type = 'file';
fileInput.accept = '.xlsx';
fileInput.style.position = 'fixed';
fileInput.style.top = '10px';
fileInput.style.left = '600px';
fileInput.style.zIndex = '99999';
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

tableImportButton.addEventListener('click', () => {
        const tableData = collectDataFromOverviewTable();
        if (!tableData.length) {
                alert("No valid rows found in table.");
                return;
        }
        startAtossImport(tableData);
});

// Handle Excel file selection
fileInput.addEventListener('change', (event) => {
        const file = event.target.files[0];
        if (!file) return;

        const reader = new FileReader();
        reader.onload = (e) => {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, { type: 'array' });

                const sheetName = workbook.SheetNames[0];
                const parsedSheetData = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName]);
                console.log('Sheet Data:', parsedSheetData);

                startAtossImport(parsedSheetData);
        };

        reader.readAsArrayBuffer(file);

});