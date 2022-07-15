// VARIABLES GLOBALES
const thisBook = SpreadsheetApp.getActiveSpreadsheet();
const hojaActiva = thisBook.getActiveSheet();


function salir(){

   const inicio = thisBook.setActiveSheet(thisBook.getSheetByName("TABLERO HAWEI"), true);
                  hojaActiva.hideSheet();


};