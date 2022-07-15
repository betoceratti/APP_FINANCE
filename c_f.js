
/** @OnlyCurrentDoc */

//LA FUNCION NATIVA ONOPEN EJECUTA LA FUNCION CUAN ABRE EL LIBRO


function onOpen(){

 crearMenu();


};


//FUNCION ONEDIT PARA QUE CUANDO SE DIETE SEGUB LAS CONDICIONES DE DIVERSAS FUNCIONES SE EJECCUTEN

function onEdit(e){
 
  //formatear(); 
  timeStamp();
  //copyPaste(); 
  alertas(); 
  //edit(e);

};





function edit(e){

  var range = e.range;

  var user = e.user;

  range.setNote('Last modified: ' + new Date() + "Usuario :" + user);

};




//CREAR UN MENU DONDE PODAMOS EJECUTAR NUESTRAS FUNCIONES 
function crearMenu(){
    var myMenu = SpreadsheetApp.getUi().createMenu('Ceratti_Menu');
    /*myMenu.addItem('Proteger hoja activa', 'ProtectActiva').addToUi();
    myMenu.addSeparator();
    myMenu.addItem('Proteger Ingresos', 'protectRangoFechaIngresos').addToUi();
    myMenu.addSeparator();
    myMenu.addItem('Proteger Gastos', 'protectRangoFechaGastos').addToUi();
    myMenu.addSeparator();
    myMenu.addItem('Proteger menos celdas', 'protegertodomenosceldas').addToUi();*/
    myMenu.addItem('Hide Records', 'ocultar').addToUi();
    myMenu.addSeparator();
    myMenu.addItem('Show Records', 'display').addToUi(); 
    myMenu.addSeparator();    
   

};





//Funcion para emitir alertas al realizar cierta funcion o cumplir una condicion.

function alertas(){
  
  var gastos = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('CONTROL GASTOS');
  var hojaactiva = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var celdaActiva = gastos.getActiveCell();
  var valorCondicion = celdaActiva.getValue();
  var filaActiva = celdaActiva.getRow();
  var columnActiva= celdaActiva.getColumn();
  var ultimaFila = gastos.getRange('d1').getValue();
  //Logger.log(ultimaFila);
  var Monto = gastos.getRange(ultimaFila, 3).getValue();
  var banco = gastos.getRange(filaActiva, 6).getValue();
  var descripcion = gastos.getRange(filaActiva, 4).getValue();
  
  
  //HOJA TABLERO PARA BUSCAR LOS SALDOS
  var  Saldos = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('TABLERO HAWEI');
  var  saldoBancomer = Saldos.getRange(3, 3,2,2).getValue();
  var saldoBancoazteca = Saldos.getRange(3, 7, 2,2).getValue();

 //DATOS DE HOJA LOGICA PARA BUSCAR SALDO POR CONCEPTO

  const cf = SpreadsheetApp.getActiveSpreadsheet();
  const logica = cf.getSheetByName('LOGICA');
  var areaBusquedabancomer = logica.getRange('a50:d68').getValues();
  var areaBusquedaazteca = logica.getRange('f50:i68').getValues();
  var conceptosBancomer = areaBusquedabancomer.map(descripcion => descripcion[0]);
  var conceptosBazteca = areaBusquedaazteca.map(descripcion => descripcion[0]);
  var valorBuscado = descripcion;
  var indiceAzteca = conceptosBazteca.indexOf(valorBuscado);
  var indiceBancomer = conceptosBancomer.indexOf(valorBuscado);

  var saldoxConceptoazteca = areaBusquedaazteca[indiceAzteca][3];
  var saldoxConceptobancomer = areaBusquedabancomer[indiceBancomer][3];

 // console.log(saldoxConceptoazteca); 



  
  
  if(filaActiva==ultimaFila && columnActiva >5 && banco != "" && hojaactiva.getName()=='CONTROL GASTOS'){
     
    Browser.msgBox("El Cargo por la cantidad de : $"+ Monto + "  " +"por concepto :" + descripcion +" " +" ha sido registrado con exito  ala cuenta de "+ " "+ banco);
   
    if (banco == "BANCOMER"){
    //var  sb = gastos.getRange(1,7).getValue();
    Browser.msgBox("El gasto total del mes por concepto de : " + descripcion + " es de $ : " + saldoxConceptobancomer + " " + "Tu saldo actual es de : $ " + saldoBancomer);
      
    var libro = SpreadsheetApp.getUi();

    var respuesta = libro.alert('Deseas agregar mas gastos ?',libro.ButtonSet.YES_NO); 
      if (respuesta == 'YES'){
        gastos.getRange(ultimaFila + 1, 1).activate();
        
      }else{
        gastos.hideSheet();
        var  Saldos = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('TABLERO HAWEI');
        Saldos.getRange('D3').activate();
        
        }
   }else if(banco == "BANCO AZTECA"){
  
    Browser.msgBox("El gasto total del mes por concepto de : " + descripcion + " es de $ : " + saldoxConceptoazteca + " " + "Tu saldo actual es de : $ " + saldoBancoazteca );

 
    var libro = SpreadsheetApp.getUi();
    var respuesta = libro.alert('Deseas agregar mas gastos ?',libro.ButtonSet.YES_NO); 
      if (respuesta == 'YES'){
        gastos.getRange(ultimaFila + 1, 1).activate();
        
      }else{
        gastos.hideSheet();
        
        var  Saldos = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('TABLERO HAWEI');
        Saldos.getRange('H3').activate();
        
        }
    
  }   
  }
  
   

};


function buscarDato(){

  //DATOS DE HOJA LOGICA PARA BUSCAR SALDO POR CONCEPTO

  const libro = SpreadsheetApp.getActiveSpreadsheet();
  var gastos = libro.getSheetByName('CONTROL GASTOS');
  const logica = libro.getSheetByName('LOGICA');
  var areaBusquedabancomer = logica.getRange('a50:d67').getValues();
  var areaBusquedaazteca = logica.getRange('f50:i67').getValues();
  var conceptosBancomer = areaBusquedabancomer.map(descripcion => descripcion[0]);
  var conceptosBazteca = areaBusquedaazteca.map(descripcion => descripcion[0]);
  var valorBuscado = "WALTMART";
  var indiceAzteca = conceptosBazteca.indexOf(valorBuscado);
  var indiceBancomer = conceptosBancomer.indexOf(valorBuscado);

  var saldoxConceptoazteca = areaBusquedaazteca[indiceAzteca][3];
  var saldoxConceptobancomer = areaBusquedabancomer[indiceBancomer][3];

  console.log(saldoxConceptoazteca);

};




//ALERTAS PARA CAPTURA DIRECTA EN LA HOJA DE GASTOS
function formulario(){
    var data = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('TABLERO HAWEI');
    var gastos = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('CONTROL GASTOS');
    var ultimaFila = gastos.getRange('d1').getValue();

    var fecha = data.getRange('A3').getValue();
    var monto = data.getRange('A5').getValue();
    var concepto = data.getRange('A7').getValue();
    var banco = data.getRange('A9').getValue();
    var nota = data.getRange('A11').getValue();  
      
    var gasto=[[fecha,nota,monto,concepto,banco]];  
    //Logger.log(gasto); 
      
    var rangoDestino = gastos.getRange(ultimaFila + 1, 1, 1, 5);
      
    var  data = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('TABLERO HAWEI');
    var  SaldoBancomer = data.getRange(3, 4,2,2).getValue();
    var SaldoBancoazteca = data.getRange(3, 8, 2,2).getValue();
        
      
      if(fecha != "" && monto != "" && concepto != "" && banco != ""){
        
        rangoDestino.setValues(gasto);
        Browser.msgBox("El Cargo por la cantidad de : $"+ monto + "  " +"por concepto :" + concepto +" " +" ha sido registrado con exito  ala cuenta de "+ " "+ banco);
        if (banco == "BANCOMER"){
        //var  sb = gastos.getRange(1,7).getValue();
        Browser.msgBox("Tu saldo actual es de : $ " + SaldoBancomer);  
        var libro = SpreadsheetApp.getUi();
        var respuesta = libro.alert('Deseas agregar mas gastos ?',libro.ButtonSet.YES_NO); 
          if (respuesta == 'YES'){
            var limpiar = data.getRange('A3:A13').clearContent();
            data.getRange('A3').activate();
            
          }else{
            var limpiar = data.getRange('A3:A13').clearContent();
            data.getRange('D3');
            }
      }else if(banco == "BANCO AZTECA"){
          
        Browser.msgBox("Tu saldo actual es de : $ " + SaldoBancoazteca);
      
        var libro = SpreadsheetApp.getUi();
        var respuesta = libro.alert('Deseas agregar mas gastos ?',libro.ButtonSet.YES_NO); 
          if (respuesta == 'YES'){
            var limpiar = data.getRange('A3:A13').clearContent();
            data.getRange('A3').activate();
            
          }else{
            var limpiar = data.getRange('A3:A13').clearContent();
            data.getRange('H3');
            }
        
      } 
        
        
        
        
        
        } else{
            Browser.msgBox("La data esta incompleta favor de revisar tu captura "); 
        
        
          }
       
  
};



//REGISTRAS INGRESOS MENDIANTE FORMULARIO BANCOS

function newIngresos(){
  let = plantilla = HtmlService.createTemplateFromFile('formWebIngresos');
  let page = plantilla.evaluate()

  page.setHeight(400).setWidth(600)

  const modal = SpreadsheetApp.getUi()
  modal.showModalDialog(page,"A T T I")


};

function getIngresos(){

   var hojaData = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('DATA');

  var datos = hojaData.getDataRange().getDisplayValues();
      datos.shift();

  //console.log(datos)
  
  return datos;


};



function saveEntradas(data){
     const hojaPrueba = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('test');

    var gastos = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('CONTROL INGRESOS');
    var ultimaFila = gastos.getRange('d1').getValue();
    

    const fila = [[data.fecha ,data.nota,data.monto ,data.descripcion," ",data.banco ]]
   
    
    const rangoDestino = gastos.getRange(ultimaFila +1, 1,1,6);
   
          rangoDestino.setValues(fila)
    //hojaPrueba.appendRow(fila)
 
   
    //notify(fila)



};



//FUNCION PARA EL BOOTSTARP HTML GASTOS BANCOS

function newForm(){
  let = plantilla = HtmlService.createTemplateFromFile('newHtmlBancos');
  let page = plantilla.evaluate()

  page.setHeight(400).setWidth(600)

  const modal = SpreadsheetApp.getUi()
  
  modal.showModalDialog(page,"A T T I")


};




//FUNCION PARA OBTENER LA DATA DE LOS SELECT DEL FORMULARIO

function getData(){

  var hojaData = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('DATA');

  var datos = hojaData.getDataRange().getDisplayValues();
      datos.shift();

  //console.log(datos)
  
  return datos;



};






//FUNCION PARA FORMULARIO DE GASTOS DE BANCOS QUE USBA ALERTAS BANCOS

function saveBancos(data){

    const hojaPrueba = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('test');

    var gastos = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('CONTROL GASTOS');
    var ultimaFila = gastos.getRange('d1').getValue();
    

    const fila = [[data.fecha ,data.nota,data.monto ,data.descripcion," ",data.banco ]]
   
    
    const rangoDestino = gastos.getRange(ultimaFila +1, 1,1,6);
   
          rangoDestino.setValues(fila)
    //hojaPrueba.appendRow(fila)
 
   
    //notify(fila)


};


//funcion para delte ultima fila de bancos gastos

function deleteRegister(){


    var libro = SpreadsheetApp.getActiveSpreadsheet();

    var tablero = libro.getSheetByName("TABLERO HAWEI");

    var gastos = libro.getSheetByName('CONTROL GASTOS');

    var ultimaFila = gastos.getRange('d1').getValue();


     var title = gastos.getRange(ultimaFila,2,1,6).getValues();

      

    var  pregunta = SpreadsheetApp.getUi();
    var respuesta = pregunta.alert('Estas seguro de eliminar el ultimo registro capturado con la data : '+ title + '?', pregunta.ButtonSet.YES_NO);
    
      if(respuesta == 'YES'){
    
         gastos.getRange(ultimaFila,1,1,6).clearContent();
      
        notify("Registro con descripcion :"  + title + "eliminado con exito!")
        tablero.getRange('C3').activate();
    
    
     }else{
       
       Browser.msgBox('Ningun registro fue eliminado !');
       tablero.getRange('C3').activate();
    
    } 

  
};



//FUNCION PARA MODAL HTML GASTOS CASH

function form(){
  
    //CONECTAMOS CON EL ARCHIVO HTML
    var modal = HtmlService.createTemplateFromFile("FormCash");
      
    // CREAMOS EL MODAL DONDE DESPELGAREMOS EL ARCHIVO HTML  
    var pagina = modal.evaluate();
        pagina.setHeight(500);
    var formulario = SpreadsheetApp.getUi() ; 
    formulario.showModalDialog(pagina, "A T T I");
    //formulario.showSidebar(pagina);  


};

//FUNCION PARA INGRESAR INGRESOS CASH

function saveIngresos(){

  const libro = SpreadsheetApp.getActiveSpreadsheet();
  const hojaIngresos = libro.getSheetByName('cash_control')
  let celdaActiva = hojaIngresos.getRange('t1');

  const validacion = SpreadsheetApp.getUi();

  const respuesta = validacion.alert("Deseas capturar un ingreso ?", validacion.ButtonSet.YES_NO)

  if(respuesta == "YES"){

      const monto = parseInt( Browser.inputBox("Captura una cantidad $"));

      const ultimaFila = hojaIngresos.getRange('q1').getValue();
    

      const rangoDestino = hojaIngresos.getRange(ultimaFila + 1,17).setValue(monto);
    
      notify("El ingreso por un monto de : " + monto + " fue registrado con exito !")

     celdaActiva.activate();

  } else{

    celdaActiva.activate();

  }

 

};



//FUNCION PARA INSERTAR LA DATA DE CASH 
function saveGastos(data){

    var hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('TABLERO HAWEI');
    var fecha = hoja.getRange(22, 4); 
    var monto = hoja.getRange(22, 5);
    var categoria = hoja.getRange(22, 6);
    var concepto = hoja.getRange(22, 7);  
    var nota = hoja.getRange(22,8)

    //Browser.msgBox(data.fecha);
    fecha.setValue(data.fecha);  
    monto.setValue(data.monto);
    categoria.setValue(data.category);
    concepto.setValue(data.descripcion); 
    nota.setValue(data.nota)
      
    var detalle =SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Cash_Control');
    var ultimafila = detalle.getRange('H1').getValue();
      var rangoOrigen = hoja.getRange('D22:H22').getValues();
    var destinofecha = detalle.getRange(ultimafila + 1, 1, 1, 5);
      //rangodestino.setValues(data.fecha,data.monto,data.concepto,data.banco);
      destinofecha.setValues(rangoOrigen);
      
    //detalle.appendRow([data.fecha,data.monto,data.concepto,data.banco]);
    //return true;  

};


function eliminarUltimacaptura(){

    var hojaDatos =SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Cash_Control');
    var ultimafila = hojaDatos.getRange('H1').getValue();

    var title = hojaDatos.getRange(ultimafila,2,1,3).getValues();

    var  pregunta = SpreadsheetApp.getUi();
    var respuesta = pregunta.alert('Estas seguro de eliminar el ultmo registro capturado con descripcion : '+ title + '?', pregunta.ButtonSet.YES_NO);
    
      if(respuesta == 'YES'){
    
         hojaDatos.getRange(ultimafila,1,1,5).clearContent();
      
        notify("Registro con descripcion :"  + title + "eliminado con exito!")

    
     }else{
       
       Browser.msgBox('Ningun registro fue eliminado !');
       hojaDatos.getRange('t1').activate();
    
    } 
   
   

};


//FUNCION PARA EMITIR NOTIDICACIONES TOAST 
function notify(mensaje) {
   
  SpreadsheetApp.getActive ().toast(mensaje,"A T T I SYSTEM");
  
};


//Funcion para ocultar columnas


function ocultar() {
  
  var hoja = SpreadsheetApp.getActive();
  
  var ultimaFila = hoja.getRange('d1').getValue();
  
  //Logger.log(ultimaFila);  
  
  hoja.getActiveSheet().hideRows(3, ultimaFila-10);
  
  hoja.getRange('A' + ultimaFila).activate();
};


function display(){
  var hoja = SpreadsheetApp.getActive();
  var ultimaFila = hoja.getRange('d1').getValue();
  
  hoja.getActiveSheet().showRows(3, ultimaFila);
  hoja.getRange('A' + ultimaFila).activate();
  
};


//Funcion para protejer hoja activa


function protectActiva(){
  var hojaActiva = SpreadsheetApp.getActiveSheet();
  
  hojaActiva.protect().setDescription('no tocar').removeEditors(['ceratti.web@gmail.com']);
};



function protegerHojaespecifica(){
  
  var hojaEspecifica = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('DASHBOARD MOVIL');
  var proteccion = hojaEspecifica.protect().setDescription('no tocar').removeEditors(['ceratti.web@gmail.com']);

};


function protectRango(){
  var hojaEspecifica = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('CONTROL GASTOS');
  var rango = hojaEspecifica.getRange(1, 6);
  var proteccion = rango.protect().setDescription('no tocar').removeEditors(['ceratti.web@gmail.com']);
};



function protegertodomenosceldas(){
  
  var hojaEspecifica = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('DASHBOARD MOVIL');
  var rango = hojaEspecifica.getRange(28, 1, 1, 2);
  var proteccion = hojaEspecifica.protect().setUnprotectedRanges([rango])
                                           .setDescription('no tocar')
                                           .removeEditors(['ceratti.web@gmail.com']);

};


function protectRangoFechaGastos(){
  var hojaEspecifica = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('CONTROL GASTOS');
  var celdaActiva = hojaEspecifica.getActiveCell();
  var filaActiva = celdaActiva.getRow();
  var columActiva = celdaActiva.getColumn();
  var valor = celdaActiva.getValue();
  var contador = hojaEspecifica.getRange('d1').getValue();
  var ultimaFila = hojaEspecifica.getLastRow();
  
  /*Logger.log(ultimaFila);*/
  
  
  var rango = hojaEspecifica.getRange(1, 1, contador, 10);
  var proteccion = rango.protect().setDescription('no tocar').removeEditors(['ceratti.web@gmail.com']);
};


function protectRangoFechaIngresos(){
  var hojaEspecifica = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('CONTROL INGRESOS');
  var celdaActiva = hojaEspecifica.getActiveCell();
  var filaActiva = celdaActiva.getRow();
  var columActiva = celdaActiva.getColumn();
  var valor = celdaActiva.getValue();
  var contador = hojaEspecifica.getRange('d1').getValue();
  var ultimaFila = hojaEspecifica.getLastRow();
  
  /*Logger.log(ultimaFila);*/
  
  
  var rango = hojaEspecifica.getRange(1, 1, contador, 10);
  var proteccion = rango.protect().setDescription('no tocar').removeEditors(['ceratti.web@gmail.com']);
};








//PARA PROGRAMAR FUNCIONES O EVENTOS EN EL TIEMPO, DEBEMOS IR A ACTIVADORES Y AHI CONFIGURAMOS 
function envioGmail(){
  var libro = SpreadsheetApp.getActiveSpreadsheet();
  var hoja = libro.getActiveSheet();
  var dashboard = libro.getSheetByName('Dashboard x ceratti');  
  
  var destinatario = dashboard.getRange(372,1).getValue();
  var nombre = dashboard.getRange(372,2).getValue();
  var monto = dashboard.getRange(372, 3).getValue();
  var alta = dashboard.getRange(372,4).getValue();
  var comentarios = dashboard.getRange(372,5).getValue();
  var plantilla = dashboard.getRange(374, 1).getValue();
  var contactos = dashboard.getRange(372, 1, 2, 5).getValues();
  var asunto = "Alta de caso a nombre de " + nombre;
  
  //PARA ENVIO DE MAILS A VARIOS DESTINATARIOS AL MISMO TIEMPO USAMOS CLICLO FOREACH
  //Logger.log(contactos);
  /*contactos.forEach(function(fila){
    Logger.log(fila[0]);
    GmailApp.sendEmail(fila[0], 'prueba','body')
  })*/
  
  
  //Con este manera el cuerpo del correo se va sin formato , todo pegado asi que haremos una plantilla, lo mismo podriamos hacer coin asunto 
  //var body = "Alta de credito a nombre de : " + nombre + "con fecha " + alta +"por un monto de :  "+ monto + " comentarios : " + comentarios;
  //plantilla con replace
  var body = plantilla.replace('{nombre}', nombre).replace('{alta}', alta).replace('{monto}', monto).replace('{comentarios}', comentarios);
 
  //Logger.log(body);
  if(hoja.getName()== dashboard.getName()){
     //var email = GmailApp.sendEmail(destinatario, asunto,body)
     Browser.msgBox('Email enviado con exito"');
  
  
  }else{
   Browser.msgBox('El mail no se ha enviado,hubo un error.');
  
  }
  
  
  //var email = GmailApp.sendEmail(destinatario, asunto,body)
  //var email = GmailApp.sendEmail("robertoceratti@gmail.com", "prueba","Hola probando una macro programada para tal monento especifico.")
};







/*function contador(){
var hoja = SpreadsheetApp.getActiveSheet();
var i;
  if(hoja.getLastRow()==361){ 
  i = 1;    
  i= hoja.getRange(hoja.getLastRow()+1,1).setValue(i); 
    hoja.getRange(hoja.getLastRow(),2).setValue(new Date());
  } else{ 
  
    i= hoja.getRange(hoja.getLastRow(),1).getValue();
    hoja.getRange(hoja.getLastRow()+1,1).setValue(i + 1);
    hoja.getRange(hoja.getLastRow(),2).setValue(new Date());
  }
  
};*/







//CON ESTA FUNCION TRAEMOS TODAS LAS HOJAS DE EL LIBRO Y NEDIANTE UN CICLO DESPLEGAMOS LOS NOMBRES INSERTADAS EN UN ENLACE 

function menu(){

  var misHojas = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  //Logger.log(misHojas);
  const nombreHoja = "Dashboard x ceratti";
  var hojaPrincipal = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Dashboard x ceratti');
  //hoja.getName();
  //Logger.log(hoja);
  var i = 239;
  
  var datos = hojaPrincipal.getRange(239, 1, 5,3).getValue();
 
  misHojas.forEach(function(hoja){
  var formula = 'HYPERLINK("#gid='+hoja.getSheetId()+'";"'+hoja.getName()+'")'  
   hojaPrincipal.getRange(i,1).setFormula(formula);
   //hojaPrincipal.getRange(i, 1).setValue(hoja.getName());
   
  i++;  
  //Logger.log(hoja.getSheetId());
  //Logger.log(datos);
   
  })
};



/*FUNCION PARA FILTAR LISTAS DESPLEGABLES */
/*para ver los mensahes de la consola log es con ctrl + enter*/
function filtro(){
 
  var libro = SpreadsheetApp.getActiveSpreadsheet();
  var hojaActiva = libro.getActiveSheet();
  var rangoActivo = hojaActiva.getActiveCell();
  var filaActiva = rangoActivo.getRow();
  var columnActiva = rangoActivo.getColumn();
  var valorActive = rangoActivo.getValue();
  var nombreHoja = "Dashboard x ceratti";
  var encabezados = 132;
  var listaDesplegable = 8;
  var listaDesplegable2= 9;
  
  if (hojaActiva.getName()==nombreHoja && filaActiva==encabezados && columnActiva == listaDesplegable){
    hojaActiva.getRange(filaActiva, listaDesplegable2).clearContent();
    if (valorActive == ''){
      hojaActiva.getRange(filaActiva, listaDesplegable2).clearDataValidations().clearContent();
    }
    
  else {
     var datos = libro.getSheetByName(nombreHoja);
     var rango = datos.getRange(130, 1, datos.getLastRow()-encabezados, 2).getValues();
     
      //Logger.log(rango);
    
     var filtrado = rango.filter(function(fila){
      return fila[0]== valorActive;  
      
     });
    
     //Logger.log(filtrado);  
    
     var desplegable = filtrado.map(function (fila){ 
      return fila[1];  
      
     });
    
    //Logger.log(desplegable);  
    
     var validation = SpreadsheetApp.newDataValidation().requireValueInList(desplegable).build();
     hojaActiva.getRange(filaActiva,listaDesplegable2).setDataValidation(validation); 
    
    
    }
    
   }
  
  
};



/* para igual son == para diferente es !=  y para o es || */

function formatear(){
  var hojaActual = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var celdaActiva = SpreadsheetApp.getActiveSheet().getActiveCell();
  var filaActiva = celdaActiva.getRow();
  var columActiva = celdaActiva.getColumn();
  var valor = celdaActiva.getValue();
  var ultimaFila = SpreadsheetApp.getActiveSheet().getLastRow();
  var nombre = hojaActual.getRange(150,1);
  
  
 
  if (filaActiva > 144 && columActiva > 5  && hojaActual.getName() == 'Dashboard x ceratti' ){
  
    celdaActiva.setFontColor('#9900ff').setFontFamily('Comfortaa').setFontSize(10);
    nombre.setValue(valor).setFontFamily('Comfortaa').setBackground('blue').setFontColor('white');
   
    //SpreadsheetApp.getUi().alert("El valor editado fue  : " + valor )
    
   
    }
    
    Logger.log(hojaActual.getName()); 
  //Dashboard x ceratti
  
};


  
//FUNCION PARA INSERTAR LA FECHA DE EDICION DE CIERTAS CELDAS 

function timeStamp(){
  var hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Dashboard x ceratti');
  var hojaActiva= SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var celdaActiva = hojaActiva.getActiveCell();
  var valor = celdaActiva.getValue();
  var filaActiva = celdaActiva.getRow();
  var columnActiva = celdaActiva.getColumn();
    
    if((filaActiva>=297 && filaActiva<310) && columnActiva==1 && hojaActiva.getName()==hoja.getName()){
      celdaActiva.offset(0,1).setValue(new Date());
      //Logger.log(hoja.getName());
    
    
    }


};


//LECTURA DE DATOS Y DISPLAY DE ESA DATA EN DETERMINADA HOJA, RANGO O CELDA.

function lectorData(){
    var libro = SpreadsheetApp.getActiveSpreadsheet();
    var hoja = libro.getActiveSheet();
    var hojaEspecifica = libro.getSheetByName('BALANCE CERATTI 2020');  
      
      
    //conociendo la celda especifica   
    var dato = hoja.getRange('A311').getValue();
      
    //ahora con fila y columna
    var dato1 = hoja.getRange(311,1).getValue();
      
    //ahora un rango con anotacion directa
    var rango = hoja.getRange('a311:b313').getValues();
      
    //ahora un rango con 4 args n fila n colu , cuantas filas y cuantas column
    var rango1 = hoja.getRange(311, 1, 3,2).getValues(); 
    /*Logger.log(dato); 
    Logger.log(dato1);  
    Logger.log(rango);  
    Logger.log(rango[0]);*/
    //Logger.log(rango1); 
    //Logger.log(rango1[1][0]);   
      
    //para accesar a elemetos de un arreglo en el primer caso seria a la primera fila asi Logger.log(rango[0])   
    //para accesra a un elemento de otro elemento seria asi Logger.log(rango[0][0])  

    //para leer datos de una hoja especifica que no sea la activa es asi 
      var rango3 = hojaEspecifica.getRange('K50:L60').getValues();
      var ingresos =hojaEspecifica.getRange('O18:R31').getValues();
      var ingresos1 = hojaEspecifica.getRange('O18:R18').getValues();
      
    //Logger.log(rango3[0]);  
      
      hoja.getRange('d313:e323').setValues(rango3);
      hoja.getRange('G313:J326').setValues(ingresos);
      
      var ultimaFila = hoja.getLastRow();
      hoja.getRange((ultimaFila + 1),1, 1,4).setValues(ingresos1);
      


};


function copyPaste(){
  
    var libro = SpreadsheetApp.getActiveSpreadsheet();
    var hoja = libro.getActiveSheet();
    var hojaEspecifica = libro.getSheetByName('BALANCE CERATTI 2020');  
    var celdaActiva = hoja.getActiveCell();
    var valor = celdaActiva.getValue();
    var filaActiva = celdaActiva.getRow();
    var columnActiva = celdaActiva.getColumn();  
      
      if(filaActiva >341 && columnActiva == 6 && valor == "OK" && hoja.getName()== "Dashboard x ceratti"){
      var rangoOrigen =hoja.getRange(filaActiva,1, 1, 5);
      //use en rangodestino filaactiva porqe esta en la misma hoja para prueba prodebemos byscar laultima fila + 1  
      var rangoDestino = hoja.getRange(filaActiva, 7, 1, 5);
      rangoOrigen.copyTo(rangoDestino);  
      //para mover la fila se usa moveTo
      // rangoOrigen.moveTo(rangoDestino);
      //en estos casos se debe eliminar la fila que quedaria vacia con 
      //hoja.deleteRow(filaActiva);  
      }

};












//FUNCION CALCULADORA DE PAGOS 


function calcular(){
  var sgcc = SpreadsheetApp.getActive();
  var hoja =sgcc.getSheetByName('LOGIN');
  var monto =parseFloat(Browser.inputBox('CAPTURA EL MONTO'));
  var plazo =parseFloat(Browser.inputBox('CAPTURA EL PLAZO (en meses)'));
  var tasa =parseFloat(Browser.inputBox('CAPTURA LA TASA'));
  var frecuencia =Browser.inputBox('CAPTURA LA FRECUENCIA DE PAGO'+ " " + "s=semanal,q=quincenal");
   	if (frecuencia == "s"){
	frecuencia = 4 * plazo;
	
	} else {
	
	frecuencia = 2 * plazo;
	}
	
  
  var capital = monto /  frecuencia ;
  var intereses = monto * tasa/100 * plazo;
  var interes = intereses / frecuencia;
  var iva =(interes * 16)/100; 
    
  var pago = capital + interes + iva ;
  var total =pago * frecuencia;
  
  hoja.getRange(12,12).setValue(monto);
  hoja.getRange(13,12).setValue(plazo);
  hoja.getRange(14,12).setValue(tasa/100);
  hoja.getRange(15,13).setValue(frecuencia);
  hoja.getRange(18,12).setValue(pago);
   hoja.getRange(20,12).setValue(total);
   hoja.getRange('k20').activate();
  Browser.msgBox("TU PAGO DE ACUERDO A LOS DATOS CAPTURADOS ES DE : " + " " + "$"+ pago + " " + "(tasa promedio prodemex 8%)" );
  
  
  
  
};

//CALCULADORA NORMAL

function calculadora(){
  var hojadecalculo =SpreadsheetApp.getActive();
  var hojaActiva =hojadecalculo.getSheetByName('prueba');
  var num1 = parseFloat(Browser.inputBox('captura el monto'));
  var num2 = parseFloat(Browser.inputBox('captura el monto'));
  var suma = num1 + num2 ;
  var resta =num1 - num2;
  var division = num1 / num2;
  var multiplicacion =num1 * num2;
  var porcentajes = num1 * num2  /100 ;
  hojaActiva.getRange(4, 11).setValue(suma);
  hojaActiva.getRange(5, 11).setValue(resta);
  hojaActiva.getRange(6, 11).setValue(multiplicacion);
  hojaActiva.getRange(7, 11).setValue(division);
  hojaActiva.getRange(8, 11).setValue(porcentajes);
  hojaActiva.getRange(2,11).setValue(num1);
  hojaActiva.getRange(2,12).setValue(num2);
  /*Browser.msgBox("el resultado de tu operacion es" + " " + multiplicacion);*/
  
   
  
}

//DAR FORMATO A CELDAS O RANGOS PARA NUMERO PESO Y TIPOGTAFIA 

function formato(){
  var spreadsheet =SpreadsheetApp.getActiveSheet();
  spreadsheet.getActiveRangeList().setFontFamily('Comfortaa').setFontSize(14);
  spreadsheet.getActiveRangeList().setFontColor('#0000ff');
    
   
}

function pesomexicano() {
  var spreadsheet =SpreadsheetApp.getActiveSheet();
  spreadsheet.getActiveRangeList().setNumberFormat('$ 0.00');
  
}



//PARA LLAMAR A UN AHOJA EN ESPECIFICO POR EL NOMBRE

function VENCIMIENTOS() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('B24').activate();
  spreadsheet.setActiveSheet(spreadsheetget.SheetByName('TABLADINAMICAVENCIMIENTOS'), true);
  spreadsheet.getRange('A1').activate();
};
  

//PARA SALIR Y OCULTA UNA HOJA 

function saliredoglobal() {
  Browser.msgBox("Favor de verficar la edicion de celdas,mencionar en las notas los movimientos realizados,Recuerda registrar los pagos, dando click y capturando el importe en el control de pagos del mes vigente!");
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('A1').activate();
  spreadsheet.getActiveSheet().hideSheet();
  
};

//salir y al mismo tiempo llamar una hoja en especifico
/*function salir(){
 
  var gla =SpreadsheetApp.getActive();
  gla.getRange('p4').clear();
  hojaactiva =gla.getActiveSheet().hideSheet(); 
  hojaactiva =gla.setActiveSheet(gla.getSheetByName('LOGIN'),true);
  

  
};*/


//LOGIN CON OTRAS FUNCIONES COMO SALI Y OCULTAR DATOS MEDIANTE CAMBIO DE FONSO Y COLORES

function login(){
  
 var usuario = Browser.inputBox('captura tu usuario');
 var password =Browser.inputBox('captura tu contraseña');
   
 
  if ( usuario == 'CERATTI' & password == 2702) {
       Browser.msgBox('hola' + " " + usuario + " " + "BIENVENIDO Recuerda utiliar tus claves de acceso  ");
       var libro =SpreadsheetApp.getActive();
       var hoja = libro.setActiveSheet(libro.getSheetByName('TABLERO CONSULTA'),true);
       hoja.getRange(4,16).setValue(usuario);
       WHITCOLOR();
       hoja.getRange('p4').activate().setFontFamily('Comfortaa').setFontSize(14).setFontColor('red');
       /*MAIL = GmailApp.sendEmail('robertoceratti@gmail.com', "Alerta seguriy CERATTI-PYTHON ","EL USUARIO : DARIO BARRIOS HA ACCESADO ALA SISTEMA")*/   
      
    
      
  
  
  } else if(usuario == 'Adbr' & password == 1306){
      Browser.msgBox('hola' + " " + "DARIO BARRIOS" + " " + "");
      var libro =SpreadsheetApp.getActive();
      var hoja = libro.setActiveSheet(libro.getSheetByName('TABLERO CONSULTA'),true);       
      WHITCOLOR();
      hoja.getRange(4,16).setValue(usuario);       
      hoja.getRange('p4').activate().setFontFamily('Comfortaa').setFontSize(14).setFontColor('red');
      Browser.msgBox("CERATTI TEC" + " " + " Apreciado Dario, favor de insertar las notas respectivas, cuando se realizen pagos extraordinarios, para su correcta aplicacion, Gracias ");
      MAIL = GmailApp.sendEmail('robertoceratti@gmail.com', "Alerta seguriy CERATTI-PYTHON ","EL USUARIO : DARIO BARRIOS HA ACCESADO AL SISTEMA")    
      
      
  
  }   else {
             Browser.msgBox('datos incorrectos,si olvidaste tus claves contacta al admin del sistema o verifica tu captura (sensible a mayusculas y minusculas)');
             var libro =SpreadsheetApp.getActive();
             var hoja =libro.setActiveSheet(libro.getSheetByName('login'),true);
                 hoja.getRange(1,12).setValue(usuario);
                 hoja.getRange('h28').activate();
           }       
  
  
};
  


function NOTCOLOR() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('K13:L14').activate();
  spreadsheet.getActiveRangeList().setFontColor('#ffffff');
  spreadsheet.getRange('K9:L9').activate();
  spreadsheet.getActiveRangeList().setFontColor('#9900ff');
  spreadsheet.getRange('E9:F9').activate();
  spreadsheet.getActiveRangeList().setFontColor('#9900ff');
  spreadsheet.getRange('E13:F13').activate();
  spreadsheet.getActiveRangeList().setFontColor('#9900ff');
  spreadsheet.getRange('E18:F18').activate();
  spreadsheet.getActiveRangeList().setFontColor('#9900ff');
  spreadsheet.getRange('H9:I9').activate();
  spreadsheet.getActiveRangeList().setFontColor('#9900ff');
  spreadsheet.getRange('L26:N26').activate();
  spreadsheet.getActiveRangeList().setFontColor('red');
  


};

function WHITCOLOR() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('K13:L14').activate();
  spreadsheet.getActiveRangeList().setFontColor('#0000ff');
  spreadsheet.getRange('K9:L9').activate();
  spreadsheet.getActiveRangeList().setFontColor('#ffffff');
  spreadsheet.getRange('E9:F9').activate();
  spreadsheet.getActiveRangeList().setFontColor('#ffffff');
  spreadsheet.getRange('E13:F13').activate();
  spreadsheet.getActiveRangeList().setFontColor('#ffffff');
  spreadsheet.getRange('E18:F18').activate();
  spreadsheet.getActiveRangeList().setFontColor('#ffffff');
  spreadsheet.getRange('H9:I9').activate();
  spreadsheet.getActiveRangeList().setFontColor('#ffffff');
  spreadsheet.getRange('l26:n26').activate();
  spreadsheet.getActiveRangeList().setFontColor('#ffffff');
};


/*function salir(){
 
  var gla =SpreadsheetApp.getActive();
  gla.getRange('p4').clear();
  hojaactiva =gla.getActiveSheet().hideSheet(); 
  hojaactiva =gla.setActiveSheet(gla.getSheetByName('LOGIN'),true);
  

  
};*/

//USANDO DOS FUNCIONES PARA SALIR DE UNA HOJA QUE DESEAS TENGA PROTECCION A LA VISTA

function KEY(){  
  NOTCOLOR();
  //salir();
  
};

























