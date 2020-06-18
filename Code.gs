function onOpen(e) {
  
  SpreadsheetApp.getUi()
  .createMenu('duban suarez')
  .addSubMenu(SpreadsheetApp.getUi().createMenu('nombre del archivo')
              .addItem('cambiar', 'cambiandoNombres2')
              .addItem('validar', 'verificando')
             )
  .addSubMenu(SpreadsheetApp.getUi().createMenu('llamadas')
              .addItem('extraer y consolidar', 'listarLlamada')
              
             )
  .addSubMenu(SpreadsheetApp.getUi().createMenu('Navegacion')
              .addItem(' URL xls para convertir', 'elegirCelda')
              .addItem(' incrementar Celda', 'incrementarCelda')
             )
  
  .addToUi();
}


function cambiandoNombres() {
  
  var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var hoja = activeSpreadsheet.getSheetByName("Respuestas de formulario 1");
  
  
  for(var iter =5; iter<394;iter++){
    
    var url = hoja.getRange(iter, 239).getValue();
    var url_ID = url.substring(33,200);    //   "https://drive.google.com/open?id=);
    var file = DriveApp.getFileById(url_ID);
    file.setName(hoja.getRange(iter, 238).getValue()); //file.setName("2020-05-05__1957300078743_Los_Muñequitos_FLORIANA_MARITZA_ASPRILLA_ARBOLEDA")
    
  }
  
}

function listarLlamada() {
  
  var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var hoja = activeSpreadsheet.getSheetByName("Respuestas de formulario 1");
  var hojaGuzman = activeSpreadsheet.getSheetByName("ListadoLLAMADAS_guzman");        Logger.log("ººhoja="+hoja);
  var hojaCupa = activeSpreadsheet.getSheetByName("ListadoLLAMADAS_cupa");
  
  
  var sheet = activeSpreadsheet.getSheets()[0];
  
  
  var PosIterGuzman =  Number(sheet.getRange(1, 10).getValue());  //Number(Browser.inputBox('Proximo de Guzman - el vacio'));
  var PosIterCupa = Number(sheet.getRange(1, 11).getValue());        //Number(Browser.inputBox('Proximo de Fundecupa - el vacio'));                     Logger.log("ººPosIterGuzman="+PosIterGuzman);
  var PosIni =Number(sheet.getRange(1, 12).getValue());  //Number( Browser.inputBox('desde donde vamos a revisar'));
  var PosFinLIM = Number(sheet.getRange(1, 13).getValue()); //Number( Browser.inputBox('hasta donde'));
  var  PosFin =  PosIni;  
  
  if(PosIni > PosFinLIM){
    
    Browser.msgBox("Error");
  }
  
  var adicional="";
  var opcion="";
  var url="";
  var operador="";
  for(var iter = PosIni; iter<=PosFin;iter++){
    opcion = hoja.getRange(iter, 7).getValue();
    url = hoja.getRange(iter, 240).getValue();
    operador= hoja.getRange(iter, 3).getValue();
    Logger.log("fila ="+iter);  
    
    
    Logger.log("ººopcion="+opcion);
    Logger.log("ººurl="+url);
    Logger.log("ººoperador="+operador);
    
    if (opcion=="Entrega de Racion"){
      continue;
    }
    if (opcion=="Avance de planeacion"){
      continue;
    }
    if (opcion=="Con un archivo de EXCEL"){
      if(operador=="FUN. Cuenca del Pacifico"){
        PosIterCupa =listarDesdeExcel(hojaCupa,PosIterCupa,url,hoja,iter);
        
      }
      else{
        PosIterGuzman =listarDesdeExcel(hojaGuzman,PosIterGuzman,url,hoja,iter);
        
      }
      hoja.getRange(iter, 241).setValue(iter);
    }
    if (opcion=="Con una foto al formato impreso"){
      if(operador=="FUN. Cuenca del Pacifico"){
        PosIterCupa =listarDesdeImpreso(hoja,iter,hojaCupa,PosIterCupa);
        hoja.getRange(iter, 241).setValue(iter);
      }
      else{
        PosIterGuzman =listarDesdeImpreso(hoja,iter,hojaGuzman,PosIterGuzman);
        hoja.getRange(iter, 241).setValue(iter);
      }
      
    }
    
    sheet.getRange(1, 10).setValue(PosIterGuzman);
    sheet.getRange(1, 11).setValue(PosIterCupa);
    sheet.getRange(1, 12).setValue(iter+1);
    
    
    
    //hoja.getRange(iter, 241).setValue(iter);
    
  }
  
  
  
  
}



function listarDesdeExcel(hojaDes,PosRegistro,url, hojaForm, iterMacro){ // https://docs.google.com/spreadsheets/d/
  // var url_ID = url.substring(33,200); //   "https://drive.google.com/open?id=
  var url_ID = extraerID_url(url);
  Logger.log("ºexcelºurl_ID="+url_ID);
  
  var file = DriveApp.getFileById(url_ID);
  
  
  
  var hojaOrigen = SpreadsheetApp.openById(url_ID).getSheetByName("FORMATO");
  
  
  
  // var hojaOrigen = activeSpreadsheet.getSheetByName("FORMATO");
  if(hojaOrigen==null){
    Browser.msgBox("se complico, el arhivo Origen")
  }
  var iter =12;
  
  var dtGeneral =[];
  var d  =[];
  
  
  d.push(hojaForm.getRange(iterMacro, 1).getValue()); //trae la fecha registro google
  d.push(hojaForm.getRange(iterMacro, 6).getValue());//trae la fecha del formato
  d.push(hojaForm.getRange(iterMacro, 4).getValue()+hojaForm.getRange(iterMacro, 5).getValue());//trae la uds sea cual fuere el operador
  
  dtGeneral.push(d);
  
  
  
  
  
  var ind =1;
  for(iter=12;iter<=22;iter++){
    if(hojaOrigen.getRange(iter,3).getValue() != ""){
      Logger.log("registrando");
      
      if(hojaOrigen.getRange(iter, 3).getValue() != ""){
        hojaDes.getRange(PosRegistro, 1,1,3).setValues(dtGeneral);//////esta es la buena linea   
        
        var  infoCall =   hojaOrigen.getRange(iter, 1,1, 30).getValues();
        hojaDes.getRange(PosRegistro, 9).setValue(ind);
        hojaDes.getRange(PosRegistro, 10,1,30).setValues(infoCall);
        ind++;
        PosRegistro ++;
        
      }
    }
    
  }
  
  
  return PosRegistro; // devolverl la proxima posicion vacia del listado
  
  
  
  
}

function listarDesdeImpreso(hojaForm,iterMacro,hojaDes,PosRegistro){
  Logger.log("ºregistroºiterMacro="+iterMacro);
  
  
  //   iterMacro=770;
  //PosRegistro=2356;
  //  hojaForm=SpreadsheetApp.getActive().getSheetByName("Respuestas de formulario 1");
  //hojaDes=SpreadsheetApp.getActive().getSheetByName("ListadoLLAMADAS_cupa");    //1131
  //  hojaDes=SpreadsheetApp.getActive().getSheetByName("ListadoLLAMADAS_guzman");//2356
  
  
  var condicion =false;//true;
  var llamada=0;
  var colOrigenLlamada =0;
  var colDestinoLlamada =0;
  Logger.log("hojaDes="+hojaDes);
  
  do{
    
    
    var dtGeneral =[];
    var d  =[];
    
    
    d.push(hojaForm.getRange(iterMacro, 1).getValue()); //trae la fecha registro google
    d.push(hojaForm.getRange(iterMacro, 6).getValue());//trae la fecha del formato
    d.push(hojaForm.getRange(iterMacro, 4).getValue()+hojaForm.getRange(iterMacro, 5).getValue());//trae la uds sea cual fuere el operador
    
    dtGeneral.push(d);
    
    hojaDes.getRange(PosRegistro, 1,1,3).setValues(dtGeneral);//////esta es la buena linea
    
    
    Logger.log("call1");
    llamada=1; //llamada 1
    hojaDes.getRange(PosRegistro,9).setValue(llamada);
    
    colOrigenLlamada=30;
    colDestinoLlamada =10;
    
    
    var infoCall =   hojaForm.getRange(iterMacro, colOrigenLlamada,1, 25).getValues();
    hojaDes.getRange(PosRegistro, colDestinoLlamada,1,25).setValues(infoCall);
    
    var infoAcu =[];
    var dtAcu  =[];
    dtAcu.push(hojaForm.getRange(iterMacro, 208).getValue());
    dtAcu.push(hojaForm.getRange(iterMacro, 209).getValue());
    dtAcu.push(hojaForm.getRange(iterMacro, 229).getValue());
    infoAcu.push(dtAcu);
    hojaDes.getRange(PosRegistro, 35,1,3).setValues(infoAcu);
    
    
    
    PosRegistro++; //fin llamada 1
    
    //llamada 2
    llamada=2; 
    colOrigenLlamada=55;
    
    if(noHayLlamada(hojaForm.getRange(iterMacro,colOrigenLlamada+1).getValue())){
      condicion =false;
      break;
    }
    
    hojaDes.getRange(PosRegistro, 1,1,3).setValues(dtGeneral);//////esta es la buena linea
    hojaDes.getRange(PosRegistro,9).setValue(llamada);
    
    infoCall =   hojaForm.getRange(iterMacro, colOrigenLlamada,1, 25).getValues();
    hojaDes.getRange(PosRegistro, colDestinoLlamada,1,25).setValues(infoCall);
    
    
    infoAcu =[];
    dtAcu  =[];
    dtAcu.push(hojaForm.getRange(iterMacro, 208).getValue());
    dtAcu.push(hojaForm.getRange(iterMacro, 209).getValue());
    dtAcu.push(hojaForm.getRange(iterMacro, 229).getValue());
    infoAcu.push(dtAcu);
    hojaDes.getRange(PosRegistro, 35,1,3).setValues(infoAcu);
    
    PosRegistro++; //fin llamada 2
    
    //llamada 3
    llamada=3; 
    colOrigenLlamada=80;
    
    if(noHayLlamada(hojaForm.getRange(iterMacro,colOrigenLlamada+1).getValue())){
      condicion =false;
      break;
    }
    
    hojaDes.getRange(PosRegistro, 1,1,3).setValues(dtGeneral);//////esta es la buena linea
    hojaDes.getRange(PosRegistro,9).setValue(llamada);
    
    
    
    infoCall =   hojaForm.getRange(iterMacro, colOrigenLlamada,1, 25).getValues();
    hojaDes.getRange(PosRegistro, colDestinoLlamada,1,25).setValues(infoCall);
    
    
    infoAcu =[];
    dtAcu  =[];
    dtAcu.push(hojaForm.getRange(iterMacro, 208).getValue());
    dtAcu.push(hojaForm.getRange(iterMacro, 209).getValue());
    dtAcu.push(hojaForm.getRange(iterMacro, 229).getValue());
    infoAcu.push(dtAcu);
    hojaDes.getRange(PosRegistro, 35,1,3).setValues(infoAcu);
    
    PosRegistro++; //fin llamada 3
    
    //llamada 4
    llamada=4; 
    colOrigenLlamada=105;
    
    if(noHayLlamada(hojaForm.getRange(iterMacro,colOrigenLlamada+1).getValue())){
      condicion =false;
      break;
    }
    
    hojaDes.getRange(PosRegistro, 1,1,3).setValues(dtGeneral);//////esta es la buena linea
    hojaDes.getRange(PosRegistro,9).setValue(llamada);
    
    
    
    infoCall =   hojaForm.getRange(iterMacro, colOrigenLlamada,1, 25).getValues();
    hojaDes.getRange(PosRegistro, colDestinoLlamada,1,25).setValues(infoCall);
    
    var infoAcu =[];
    var dtAcu  =[];
    dtAcu.push(hojaForm.getRange(iterMacro, 208).getValue());
    dtAcu.push(hojaForm.getRange(iterMacro, 209).getValue());
    dtAcu.push(hojaForm.getRange(iterMacro, 229).getValue());
    infoAcu.push(dtAcu);
    hojaDes.getRange(PosRegistro, 35,1,3).setValues(infoAcu);
    PosRegistro++; //fin llamada 4
    
    //llamada 5
    llamada=5; 
    colOrigenLlamada=130;
    
    if(noHayLlamada(hojaForm.getRange(iterMacro,colOrigenLlamada+1).getValue())){
      condicion =false;
      break;
    }
    
    hojaDes.getRange(PosRegistro, 1,1,3).setValues(dtGeneral);//////esta es la buena linea
    hojaDes.getRange(PosRegistro,9).setValue(llamada);
    
    
    
    infoCall =   hojaForm.getRange(iterMacro, colOrigenLlamada,1, 25).getValues();
    hojaDes.getRange(PosRegistro, colDestinoLlamada,1,25).setValues(infoCall);
    
    var infoAcu =[];
    var dtAcu  =[];
    dtAcu.push(hojaForm.getRange(iterMacro, 208).getValue());
    dtAcu.push(hojaForm.getRange(iterMacro, 209).getValue());
    dtAcu.push(hojaForm.getRange(iterMacro, 229).getValue());
    infoAcu.push(dtAcu);
    hojaDes.getRange(PosRegistro, 35,1,3).setValues(infoAcu);
    PosRegistro++; //fin llamada 5
    
    //llamada 6
    llamada=6; 
    colOrigenLlamada=155;
    
    if(noHayLlamada(hojaForm.getRange(iterMacro,colOrigenLlamada+1).getValue())){
      condicion =false;
      break;
    }
    
    hojaDes.getRange(PosRegistro, 1,1,3).setValues(dtGeneral);//////esta es la buena linea
    hojaDes.getRange(PosRegistro,9).setValue(llamada);
    
    
    
    infoCall =   hojaForm.getRange(iterMacro, colOrigenLlamada,1, 25).getValues();
    hojaDes.getRange(PosRegistro, colDestinoLlamada,1,25).setValues(infoCall);
    
    var infoAcu =[];
    var dtAcu  =[];
    dtAcu.push(hojaForm.getRange(iterMacro, 208).getValue());
    dtAcu.push(hojaForm.getRange(iterMacro, 209).getValue());
    dtAcu.push(hojaForm.getRange(iterMacro, 229).getValue());
    infoAcu.push(dtAcu);
    hojaDes.getRange(PosRegistro, 35,1,3).setValues(infoAcu);
    PosRegistro++; //fin llamada 6
    
    //llamada 7
    llamada=7; 
    colOrigenLlamada=180;
    
    if(noHayLlamada(hojaForm.getRange(iterMacro,colOrigenLlamada+1).getValue())){
      condicion =false;
      break;
    }
    
    hojaDes.getRange(PosRegistro, 1,1,3).setValues(dtGeneral);//////esta es la buena linea
    hojaDes.getRange(PosRegistro,9).setValue(llamada);
    
    
    
    infoCall =   hojaForm.getRange(iterMacro, colOrigenLlamada,1, 25).getValues();
    hojaDes.getRange(PosRegistro, colDestinoLlamada,1,25).setValues(infoCall);
    
    var infoAcu =[];
    var dtAcu  =[];
    dtAcu.push(hojaForm.getRange(iterMacro, 208).getValue());
    dtAcu.push(hojaForm.getRange(iterMacro, 209).getValue());
    dtAcu.push(hojaForm.getRange(iterMacro, 229).getValue());
    infoAcu.push(dtAcu);
    hojaDes.getRange(PosRegistro, 35,1,3).setValues(infoAcu);
    PosRegistro++; //fin llamada 7
    
    
    
    
  }while(condicion);
  
  
  
  
  
  
  return PosRegistro;// debo devolverl la proxima posicion vacia del listado
}

function noHayLlamada(nombreBene){
  if(nombreBene=="")return true;
  
  
  return false;
}


function cambiandoNombres2() {
  
  var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var hoja = activeSpreadsheet.getSheetByName("Respuestas de formulario 1");
  var PosIni = Browser.inputBox('desde que fila');
  var PosFin = Browser.inputBox('hasta cual fila');
  var adicional="";
  var opcion="";
  for(var iter =PosIni; iter<=PosFin;iter++){
    opcion = hoja.getRange(iter, 7).getValue();
    if (opcion=="Entrega de Racion"){
      adicional ="_racion"
    }
    if (opcion=="Avance de planeacion"){
      adicional ="_planeacion"
    }
    
    var url = hoja.getRange(iter, 240).getValue();
    var url_ID = extraerID_url(url);                 //---.substring(33,200);    //   "https://drive.google.com/open?id=);
    var file = DriveApp.getFileById(url_ID);
    
    
    
    file.setName(hoja.getRange(iter, 239).getValue()+adicional); //file.setName("2020-05-05__1957300078743_Los_Muñequitos_FLORIANA_MARITZA_ASPRILLA_ARBOLEDA")
    hoja.getRange(iter, 241).setValue(file.getName());
    hoja.getRange(iter, 242).setValue(file.getName());
    adicional=""
  }
  
}
function cambiandoTipoDoc() {
  
  var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var hoja = activeSpreadsheet.getSheetByName("Respuestas de formulario 1");
  var PosIni = Browser.inputBox('desde que fila');
  var PosFin = Browser.inputBox('hasta cual fila');
  var adicional="";
  var opcion="";
  for(var iter =PosIni; iter<=PosFin;iter++){
    opcion = hoja.getRange(iter, 7).getValue();
    if (opcion=="Con un archivo de EXCEL"){
      
      
      var url = hoja.getRange(iter, 239).getValue();
      var url_ID = url.substring(33,200);    //   "https://drive.google.com/open?id=);
      var file = DriveApp.getFileById(url_ID);
      
      
      
      file.setName(hoja.getRange(iter, 238).getValue()+adicional); //file.setName("2020-05-05__1957300078743_Los_Muñequitos_FLORIANA_MARITZA_ASPRILLA_ARBOLEDA")
      hoja.getRange(iter, 240).setValue(file.getName());
      adicional=""
      
      
      
    }
  }
  
}





function verificando() {
  var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var hoja = activeSpreadsheet.getSheetByName("Respuestas de formulario 1");
  var PosIni = Number(Browser.inputBox('desde que fila'));
  var PosFin = Number(Browser.inputBox('hasta cual fila'));
  
  for(var iter =PosIni; iter<PosFin;iter++){
    
    var url = hoja.getRange(iter, 239).getValue();
    var url_ID = url.substring(33,200);    //   "https://drive.google.com/open?id=);
    var file = DriveApp.getFileById(url_ID);
    
    
    hoja.getRange(iter, 240).setValue(file.getName());
    
    
  }
}



var sheetP;


function nombrePropio(name){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[0];
  
  ;
  var cell = sheet.getRange("b1");
  cell.setFormula("=PROPER(\""+name+"\")");
  
  return cell.getValue();
}


/**
* EXTRAE ID DE UNA URL
* Aprovecha  unas celdas de la primera hoja para hacer un split 
*   //"https://drive.google.com/open?id="
* @param {String} Url del archivo; Required
* @return {String} Numero Id de la Url leida de la celda O1 primera hoja 
**/
function dividiendo(dato){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[0];
  
  
  var cell = sheet.getRange("f1");
  var cell2 = sheet.getRange("c1");
  cell.setValue( dato);
  cell2.setValue( dato);
  
  return sheet.getRange("o1").getValue();
}

function elegirCelda(){
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[0];
  var Poscelda = Number(sheet.getRange(1, 14).getValue());
  
  
  var sheetFormu = ss.getSheets()[1];
  sheetFormu.activate();
  sheetFormu.setActiveRange(sheetFormu.getRange(Poscelda, 239));
  
  
  
  
}
function incrementarCelda(){
  
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[0];
  var Poscelda = Number(sheet.getRange(1, 14).getValue());
  
  Poscelda++;
  sheet.getRange(1, 14).setValue(Poscelda);
  
  
  
  
  
}





