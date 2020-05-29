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
  var  PosFin =  PosIni+3;  
  
  

  var adicional="";
  var opcion="";
  var url="";
  var operador="";
  for(var iter = PosIni; iter<=PosFin;iter++){
    opcion = hoja.getRange(iter, 7).getValue();
    url = hoja.getRange(iter, 239).getValue();
    operador= hoja.getRange(iter, 3).getValue();
    
    
    
    Logger.log("ººopcion="+opcion);
    Logger.log("ººurl="+url);
    Logger.log("ººoperador="+operador);
    
    if (opcion=="Entrega de Racion"){
      break;
    }
    if (opcion=="Avance de planeacion"){
      break;
    }
    if (opcion=="Con un archivo de EXCEL"){
      if(operador=="FUN. Cuenca del Pacifico"){
        PosIterCupa =listarDesdeExcel(hojaCupa,PosIterCupa,url,hoja,iter);
        
      }
      else{
        PosIterGuzman =listarDesdeExcel(hojaGuzman,PosIterGuzman,url,hoja,iter);
        
      }
    }
    if (opcion=="Con una foto al formato impreso"){
      if(operador=="FUN. Cuenca del Pacifico"){
        PosIterCupa =listarDesdeImpreso(hoja,iter,hojaCupa,PosIterCupa);
        
      }
      else{
        PosIterGuzman =listarDesdeImpreso(hoja,iter,hojaGuzman,PosIterGuzman);
        
      }
      
    }
    
   sheet.getRange(1, 10).setValue(PosIterGuzman);
    sheet.getRange(1, 11).setValue(PosIterCupa);
    sheet.getRange(1, 12).setValue(iter+1);
    
    // var url_ID = url.substring(33,200);    //   "https://drive.google.com/open?id=);
    // var file = DriveApp.getFileById(url_ID);
    //file.setName(hoja.getRange(iter, 238).getValue()+adicional); //file.setName("2020-05-05__1957300078743_Los_Muñequitos_FLORIANA_MARITZA_ASPRILLA_ARBOLEDA")
    
    
    hoja.getRange(iter, 240).setValue(iter);
    
  }
    

      
  
}



function listarDesdeExcel(hojaDes,PosRegistro,url, hojaForm, iterMacro){ // https://docs.google.com/spreadsheets/d/
  // var url_ID = url.substring(33,200); //   "https://drive.google.com/open?id=
  var url_ID = dividiendo(url);
  Logger.log("ºexcelºurl_ID="+url_ID);
  
  var file = DriveApp.getFileById(url_ID);
  
  
  var hojaOrigen = SpreadsheetApp.openById(url_ID).getSheetByName("FORMATO");
  
  
  
  // var hojaOrigen = activeSpreadsheet.getSheetByName("FORMATO");
  if(hojaOrigen==null){
    Browser.msgBox("se complico, el arhivo Origen")
  }
  var iter =12;
  do{
    
    hojaDes.getRange(PosRegistro,1).setValue(hojaForm.getRange(iterMacro, 1).getValue()); //trae la fecha registro google
    hojaDes.getRange(PosRegistro,2).setValue(hojaForm.getRange(iterMacro, 6).getValue());//trae la fecha del formato
    hojaDes.getRange(PosRegistro,3).setValue(hojaForm.getRange(iterMacro, 4).getValue()+hojaForm.getRange(iterMacro, 5).getValue());//trae la uds
    
    for(var carrier=9;carrier<=37;carrier++){
      
      if(carrier ==11 || carrier ==37){
        hojaDes.getRange(PosRegistro,carrier).setValue(nombrePropio(hojaOrigen.getRange(iter,carrier-8).getValue()));
      }
      else{
        hojaDes.getRange(PosRegistro,carrier).setValue(hojaOrigen.getRange(iter,carrier-8).getValue());
      }
      
    }
    
    PosRegistro ++;
    iter++;
    
  }while (hojaOrigen.getRange(iter,3).getValue() != "");
  
  
  return PosRegistro; // devolverl la proxima posicion vacia del listado
  
  
  
  
}

function listarDesdeImpreso(hojaForm,iterMacro,hojaDes,PosRegistro){
   Logger.log("ºregistroºiterMacro="+iterMacro);
  
  var condicion =true;
  var llamada=0;
  var colOrigenLlamada =0;
  var colDestinoLlamada =0;
                                                       Logger.log("hojaDes="+hojaDes);
  
  do{
    hojaDes.getRange(PosRegistro,1).setValue(hojaForm.getRange(iterMacro, 1).getValue()); //trae la fecha registro google
    hojaDes.getRange(PosRegistro,2).setValue(hojaForm.getRange(iterMacro, 6).getValue());//trae la fecha del formato
    hojaDes.getRange(PosRegistro,3).setValue(hojaForm.getRange(iterMacro, 4).getValue()+hojaForm.getRange(iterMacro, 5).getValue());//trae la uds
    
    llamada=1; //llamada 1
    hojaDes.getRange(PosRegistro,9).setValue(llamada);
    
    colOrigenLlamada=30;
    colDestinoLlamada =10;
    
    for(var carrier=0;carrier<=24;carrier++){
      if(carrier==1){
        hojaDes.getRange(PosRegistro,carrier+colDestinoLlamada).setValue(nombrePropio(hojaForm.getRange(iterMacro,carrier+colOrigenLlamada).getValue()));
      }
      else{
        hojaDes.getRange(PosRegistro,carrier+colDestinoLlamada).setValue(hojaForm.getRange(iterMacro,carrier+colOrigenLlamada).getValue());
      }
    }
    hojaDes.getRange(PosRegistro,35).setValue(hojaForm.getRange(iterMacro, 208).getValue());
    hojaDes.getRange(PosRegistro,36).setValue(hojaForm.getRange(iterMacro, 209).getValue());
    hojaDes.getRange(PosRegistro,37).setValue(nombrePropio( hojaForm.getRange(iterMacro, 229).getValue()));
    PosRegistro++; //fin llamada 1
    
    //llamada 2
    llamada=2; 
    colOrigenLlamada=55;
    
    if(noHayLlamada(hojaForm.getRange(iterMacro,colOrigenLlamada+1).getValue())){
      condicion =false;
      break;
    }
    hojaDes.getRange(PosRegistro,1).setValue(hojaForm.getRange(iterMacro, 1).getValue()); //trae la fecha registro google
    hojaDes.getRange(PosRegistro,2).setValue(hojaForm.getRange(iterMacro, 6).getValue());//trae la fecha del formato
    hojaDes.getRange(PosRegistro,3).setValue(hojaForm.getRange(iterMacro, 4).getValue()+hojaForm.getRange(iterMacro, 5).getValue());//trae la uds
    hojaDes.getRange(PosRegistro,9).setValue(llamada);
    
    for(var carrier=0;carrier<=24;carrier++){
      
      if(carrier==1){
        hojaDes.getRange(PosRegistro,carrier+colDestinoLlamada).setValue(nombrePropio(hojaForm.getRange(iterMacro,carrier+colOrigenLlamada).getValue()));
      }
      else{
        hojaDes.getRange(PosRegistro,carrier+colDestinoLlamada).setValue(hojaForm.getRange(iterMacro,carrier+colOrigenLlamada).getValue());
      }
    }
    hojaDes.getRange(PosRegistro,35).setValue(hojaForm.getRange(iterMacro, 211).getValue());
    hojaDes.getRange(PosRegistro,36).setValue(hojaForm.getRange(iterMacro, 212).getValue());
    hojaDes.getRange(PosRegistro,37).setValue(nombrePropio(hojaForm.getRange(iterMacro, 210).getValue()));
    PosRegistro++; //fin llamada 2
    
    //llamada 3
    llamada=3; 
    colOrigenLlamada=80;
    
    if(noHayLlamada(hojaForm.getRange(iterMacro,colOrigenLlamada+1).getValue())){
      condicion =false;
      break;
    }
    hojaDes.getRange(PosRegistro,1).setValue(hojaForm.getRange(iterMacro, 1).getValue()); //trae la fecha registro google
    hojaDes.getRange(PosRegistro,2).setValue(hojaForm.getRange(iterMacro, 6).getValue());//trae la fecha del formato
    hojaDes.getRange(PosRegistro,3).setValue(hojaForm.getRange(iterMacro, 4).getValue()+hojaForm.getRange(iterMacro, 5).getValue());//trae la uds
    hojaDes.getRange(PosRegistro,9).setValue(llamada);
    
    for(var carrier=0;carrier<=24;carrier++){
      if(carrier==1){
        hojaDes.getRange(PosRegistro,carrier+colDestinoLlamada).setValue(nombrePropio(hojaForm.getRange(iterMacro,carrier+colOrigenLlamada).getValue()));
      }
      else{
        hojaDes.getRange(PosRegistro,carrier+colDestinoLlamada).setValue(hojaForm.getRange(iterMacro,carrier+colOrigenLlamada).getValue());
      }
    }
    hojaDes.getRange(PosRegistro,35).setValue(hojaForm.getRange(iterMacro, 213).getValue());
    hojaDes.getRange(PosRegistro,36).setValue(hojaForm.getRange(iterMacro, 214).getValue());
    hojaDes.getRange(PosRegistro,37).setValue(nombrePropio(hojaForm.getRange(iterMacro, 215).getValue()));
    PosRegistro++; //fin llamada 3
    
    //llamada 4
    llamada=4; 
    colOrigenLlamada=105;
    
    if(noHayLlamada(hojaForm.getRange(iterMacro,colOrigenLlamada+1).getValue())){
      condicion =false;
      break;
    }
    hojaDes.getRange(PosRegistro,1).setValue(hojaForm.getRange(iterMacro, 1).getValue()); //trae la fecha registro google
    hojaDes.getRange(PosRegistro,2).setValue(hojaForm.getRange(iterMacro, 6).getValue());//trae la fecha del formato
    hojaDes.getRange(PosRegistro,3).setValue(hojaForm.getRange(iterMacro, 4).getValue()+hojaForm.getRange(iterMacro, 5).getValue());//trae la uds
    hojaDes.getRange(PosRegistro,9).setValue(llamada);
    
    for(var carrier=0;carrier<=24;carrier++){
      if(carrier==1){
        hojaDes.getRange(PosRegistro,carrier+colDestinoLlamada).setValue(nombrePropio(hojaForm.getRange(iterMacro,carrier+colOrigenLlamada).getValue()));
      }
      else{
        hojaDes.getRange(PosRegistro,carrier+colDestinoLlamada).setValue(hojaForm.getRange(iterMacro,carrier+colOrigenLlamada).getValue());
      }
    }
    hojaDes.getRange(PosRegistro,35).setValue(hojaForm.getRange(iterMacro, 216).getValue());
    hojaDes.getRange(PosRegistro,36).setValue(hojaForm.getRange(iterMacro, 217).getValue());
    hojaDes.getRange(PosRegistro,37).setValue(nombrePropio(hojaForm.getRange(iterMacro, 218).getValue()));
    PosRegistro++; //fin llamada 4
    
    //llamada 5
    llamada=5; 
    colOrigenLlamada=130;
    
    if(noHayLlamada(hojaForm.getRange(iterMacro,colOrigenLlamada+1).getValue())){
      condicion =false;
      break;
    }
    hojaDes.getRange(PosRegistro,1).setValue(hojaForm.getRange(iterMacro, 1).getValue()); //trae la fecha registro google
    hojaDes.getRange(PosRegistro,2).setValue(hojaForm.getRange(iterMacro, 6).getValue());//trae la fecha del formato
    hojaDes.getRange(PosRegistro,3).setValue(hojaForm.getRange(iterMacro, 4).getValue()+hojaForm.getRange(iterMacro, 5).getValue());//trae la uds
    hojaDes.getRange(PosRegistro,9).setValue(llamada);
    
    for(var carrier=0;carrier<=24;carrier++){
      if(carrier==1){
        hojaDes.getRange(PosRegistro,carrier+colDestinoLlamada).setValue(nombrePropio(hojaForm.getRange(iterMacro,carrier+colOrigenLlamada).getValue()));
      }
      else{
        hojaDes.getRange(PosRegistro,carrier+colDestinoLlamada).setValue(hojaForm.getRange(iterMacro,carrier+colOrigenLlamada).getValue());
      }
    }
    hojaDes.getRange(PosRegistro,35).setValue(hojaForm.getRange(iterMacro, 219).getValue());
    hojaDes.getRange(PosRegistro,36).setValue(hojaForm.getRange(iterMacro, 220).getValue());
    hojaDes.getRange(PosRegistro,37).setValue(nombrePropio(hojaForm.getRange(iterMacro, 221).getValue()));
    PosRegistro++; //fin llamada 5
    
    //llamada 6
    llamada=6; 
    colOrigenLlamada=155;
    
    if(noHayLlamada(hojaForm.getRange(iterMacro,colOrigenLlamada+1).getValue())){
      condicion =false;
      break;
    }
    hojaDes.getRange(PosRegistro,1).setValue(hojaForm.getRange(iterMacro, 1).getValue()); //trae la fecha registro google
    hojaDes.getRange(PosRegistro,2).setValue(hojaForm.getRange(iterMacro, 6).getValue());//trae la fecha del formato
    hojaDes.getRange(PosRegistro,3).setValue(hojaForm.getRange(iterMacro, 4).getValue()+hojaForm.getRange(iterMacro, 5).getValue());//trae la uds
    hojaDes.getRange(PosRegistro,9).setValue(llamada);
    
    for(var carrier=0;carrier<=24;carrier++){
      if(carrier==1){
        hojaDes.getRange(PosRegistro,carrier+colDestinoLlamada).setValue(nombrePropio(hojaForm.getRange(iterMacro,carrier+colOrigenLlamada).getValue()));
      }
      else{
        hojaDes.getRange(PosRegistro,carrier+colDestinoLlamada).setValue(hojaForm.getRange(iterMacro,carrier+colOrigenLlamada).getValue());
      }
    }
    hojaDes.getRange(PosRegistro,35).setValue(hojaForm.getRange(iterMacro, 222).getValue());
    hojaDes.getRange(PosRegistro,36).setValue(hojaForm.getRange(iterMacro, 223).getValue());
    hojaDes.getRange(PosRegistro,37).setValue(nombrePropio(hojaForm.getRange(iterMacro, 225).getValue()));
    PosRegistro++; //fin llamada 6
    
    //llamada 7
    llamada=7; 
    colOrigenLlamada=180;
    
    if(noHayLlamada(hojaForm.getRange(iterMacro,colOrigenLlamada+1).getValue())){
      condicion =false;
      break;
    }
    hojaDes.getRange(PosRegistro,1).setValue(hojaForm.getRange(iterMacro, 1).getValue()); //trae la fecha registro google
    hojaDes.getRange(PosRegistro,2).setValue(hojaForm.getRange(iterMacro, 6).getValue());//trae la fecha del formato
    hojaDes.getRange(PosRegistro,3).setValue(hojaForm.getRange(iterMacro, 4).getValue()+hojaForm.getRange(iterMacro, 5).getValue());//trae la uds
    hojaDes.getRange(PosRegistro,9).setValue(llamada);
    
    for(var carrier=0;carrier<=24;carrier++){
      if(carrier==1){
        hojaDes.getRange(PosRegistro,carrier+colDestinoLlamada).setValue(nombrePropio(hojaForm.getRange(iterMacro,carrier+colOrigenLlamada).getValue()));
      }
      else{
        hojaDes.getRange(PosRegistro,carrier+colDestinoLlamada).setValue(hojaForm.getRange(iterMacro,carrier+colOrigenLlamada).getValue());
      }
    }
    hojaDes.getRange(PosRegistro,35).setValue(hojaForm.getRange(iterMacro, 226).getValue());
    hojaDes.getRange(PosRegistro,36).setValue(hojaForm.getRange(iterMacro, 227).getValue());
    hojaDes.getRange(PosRegistro,37).setValue(nombrePropio(hojaForm.getRange(iterMacro, 228).getValue()));
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
    
    var url = hoja.getRange(iter, 239).getValue();
    var url_ID = url.substring(33,200);    //   "https://drive.google.com/open?id=);
    var file = DriveApp.getFileById(url_ID);
    

    
    file.setName(hoja.getRange(iter, 238).getValue()+adicional); //file.setName("2020-05-05__1957300078743_Los_Muñequitos_FLORIANA_MARITZA_ASPRILLA_ARBOLEDA")
    hoja.getRange(iter, 240).setValue(file.getName());
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



function dividiendo(dato){//"https://drive.google.com/open?id="
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



 

