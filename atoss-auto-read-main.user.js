// ==UserScript==
// @name         atoss main: read from excel
// @namespace    http://tampermonkey.net/
// @version      2024-11-28
// @description  try to take over the world!
// @author       You
// @match        https://he-atoss.horiba.eu:5000/?date=*
// @icon         data:image/gif;base64,R0lGODlhAQABAAAAACH5BAEKAAEALAAAAAABAAEAAAICTAEAOw==
// @require      https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js
// @grant        none
// ==/UserScript==


const sheetData = JSON.parse(localStorage.getItem('excelData'));

if(sheetData) {
    console.log("Storage data exists!")
    let numOfEntries = sheetData.length;
    console.log("Length of sheetData <<", numOfEntries);
    let i = parseFloat(localStorage.getItem('iterator'));
    console.log("iterator <<", i)
    let ii = numOfEntries - i
    i = i+1
    localStorage.setItem('iterator', i);
    if(ii >= 0){
        //localStorage.removeItem('excelData'); // removes only 'excelData'
        if (window.opener) {
            window.close(); // only works if opened via script
        }
    } else{
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

function delay(ms) {
  return new Promise(resolve => setTimeout(resolve, ms));
}


// Create a file input dynamically
const fileInput = document.createElement('input');
fileInput.type = 'file';
fileInput.accept = '.xlsx';
fileInput.style.position = 'fixed';
fileInput.style.top = '10px';
fileInput.style.left = '600px';
document.body.appendChild(fileInput);

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
            if (i > sheetData.length-1) {
                clearInterval(interval);
            }
        }, 500); // Wait for 1000ms (1 seconds) between each iteration



        // Store the data in localStorage
        //localStorage.setItem('excelData', JSON.stringify(sheetData[0]));
        //console.log(excelDateToJSDate(sheetData[0].Date))
        //window.open('https://he-atoss.horiba.eu:5000/Home/AddOrEdit?date=' + excelDateToJSDate(sheetData[0].Date), '_blank');
    };

    reader.readAsArrayBuffer(file);
});
