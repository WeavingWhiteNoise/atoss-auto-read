// ==UserScript==
// @name         atoss sub: read from excel
// @namespace    http://tampermonkey.net/
// @version      1.03
// @description  read Excel file in Tampermonkey
// @author       You
// @match        https://he-atoss.horiba.eu:5000/Home/AddOrEdit?date=*
// @icon         data:image/gif;base64,R0lGODlhAQABAAAAACH5BAEKAAEALAAAAAABAAEAAAICTAEAOw==
// @require      https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js
// @grant        none
// ==/UserScript==

// Retrieve data from localStorage
const sheetData = JSON.parse(localStorage.getItem('excelData'));
const index = JSON.parse(localStorage.getItem('index'));


function selectByMouseEvents(matchSubstr){

    // get pop-up menue
    const $menu = $('.ui-autocomplete:visible');
    if (!$menu.length) { console.warn('no visible menu'); return false; }

    // find correct "Vorgang"
    const $item = $menu.find('li.ui-menu-item').filter(function(){
        return $(this).text().includes(matchSubstr);
    }).first();

    if (!$item.length) { console.warn('item not found'); return false; }

    // emulate what a user does: mouseenter -> mousedown -> mouseup -> click
    $item.trigger('mouseenter');
    $item.trigger('mousedown');
    $item.trigger('mouseup');
    $item.trigger('click');

    return true;
}


function selectAutocomplete(inputElement, searchTerm1, searchTerm2) {
    let $input = $(inputElement);

    // set search term into input
    $input.val(searchTerm1);

    // open autocomplete menu
    $input.autocomplete("search", searchTerm1);

    // simulate mouse click
    setTimeout(()=> selectByMouseEvents(searchTerm2), 2000);
}


// Ensure data exists
if (sheetData) {
    console.log('Retrieved Data:', sheetData[index]);


    const popup = document.createElement('div');
    popup.innerHTML = `
        <div style="
            position: fixed;
            top: 50%;
            left: 50%;
            transform: translate(-50%, -50%);
            background: white;
            padding: 25px;
            border-radius: 8px;
            box-shadow: 0 5px 25px rgba(0,0,0,0.3);
            z-index: 10000;
            min-width: 300px;
            text-align: center;
            font-family: sans-serif;
        ">
            <p style="margin: 0 0 20px 0; font-size: 16px; color: #333;">
                Loading. Please wait. ${sheetData[index]}
            </p>
            <button id="closePopupBtn" style="
                padding: 8px 20px;
                background: #007bff;
                color: white;
                border: none;
                border-radius: 4px;
                cursor: pointer;
                font-size: 14px;
            ">Close</button>
        </div>
    `;
    
    const overlay = document.createElement('div');
    overlay.style.cssText = `
        position: fixed;
        top: 0;
        left: 0;
        width: 100%;
        height: 100%;
        background: rgba(0,0,0,0.5);
        z-index: 9999;
    `;
    
    document.body.appendChild(overlay);
    document.body.appendChild(popup);
    
    // Close button functionality
    document.getElementById('closePopupBtn').addEventListener('click', () => {
        popup.remove();
        overlay.remove();
    });


    // Leistungsart
    let BOX_actionType = document.getElementById("actionType");
    BOX_actionType.value = sheetData[index].Leistungsart;
    BOX_actionType.classList.add("valid");

    // Auftrag
    let BOX_Auftrag = document.getElementById("Auftrag");
    selectAutocomplete(BOX_Auftrag, sheetData[index].Empfaenger, sheetData[index].Vorgang);

    // Arbeitsplatz
    let BOX_alternativWorkCenter = document.getElementById("alternativWorkCenter");
    BOX_alternativWorkCenter.value = sheetData[index].Arbeitsplatz;
    BOX_alternativWorkCenter.classList.add("valid");

    // Stunden
    let BOX_catsHours = document.getElementById("catsHours");
    let str = sheetData[index].Stunden
    let num = parseFloat(str.replace(',', '.'));
    num = Math.round(num * 100) / 100
    BOX_catsHours.value = num.toString().replace('.', ',');

    // Kommentar
    let BOX_Text = document.getElementById("Text");
    BOX_Text.value = sheetData[index].Kommentar;

    // Find the button by value attribute
    let saveButton = document.querySelector('input[type="submit"][value="Speichern"]');

    // Click the button
    if (saveButton) {
        setTimeout(()=> saveButton.click(), 4000);
    } else {
        console.error("Save button not found!");
    }

} else {
    console.error('No data found in localStorage!');
}
