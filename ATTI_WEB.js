const thisBook = SpreadsheetApp.getActiveSpreadsheet();
const hojaActiva = thisBook.getActiveSheet();


function doGet(e){

    const hojaDatos = thisBook.getSheetByName('DATA')
    const saldoCash = hojaDatos.getRange('d2').getDisplayValue();
    const saldoBbva = hojaDatos.getRange('d3').getDisplayValue();
    const saldoAzteca = hojaDatos.getRange('d4').getDisplayValue();

    const listas = getDatos()


    const html = HtmlService.createTemplateFromFile('index_atti');
    //la data va sobre el html
    html.saldoCash = saldoCash
    html.saldoBbva = saldoBbva
    html.saldoAzteca = saldoAzteca
    html.listas = listas
    html.getCategorias = getCategorias()

    const pagina = html.evaluate();
          pagina.addMetaTag('viewport','width=device-width, initial-scale=1,maximum-scale=1.0, user-scalable=no')
    
    return pagina



};

//FUNCION PARA  OBTENER LA DATA(SALDOS,TABLAS CASH BBVA AZTECA)

function getDatos(){
    const hojaDatos = thisBook.getSheetByName('DATA')
    const rangoData = hojaDatos.getRange('e1:r11').getDisplayValues();
          rangoData.shift()

  
    return rangoData;


};


function getCategorias(){

    const hojaDatos = thisBook.getSheetByName('DATA')
    const rangoData = hojaDatos.getRange('s1:t6').getValues();
          rangoData.shift()

  
    return rangoData;


};
