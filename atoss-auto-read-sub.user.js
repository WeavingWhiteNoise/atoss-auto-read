// ==UserScript==
// @name         atoss sub: read from excel
// @namespace    http://tampermonkey.net/
// @version      2024-11-28
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

    let BOX_actionType = document.getElementById("actionType");
    BOX_actionType.value = sheetData[index].Value1;
    BOX_actionType.classList.add("valid");

    let BOX_Auftrag = document.getElementById("Auftrag");
    selectAutocomplete(BOX_Auftrag, sheetData[index].Value2, sheetData[index].Value3);

    let BOX_alternativWorkCenter = document.getElementById("alternativWorkCenter");
    BOX_alternativWorkCenter.value = sheetData[index].Value4;
    BOX_alternativWorkCenter.classList.add("valid");

    let BOX_catsHours = document.getElementById("catsHours");
    BOX_catsHours.value = sheetData[index].Value5;

    let BOX_Text = document.getElementById("Text");
    BOX_Text.value = sheetData[index].Value6;

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
