function myFunction() {
   var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Respuestas de formulario 1");
  

  var PosIni = Browser.inputBox('desde que fila');
  var PosFin = Browser.inputBox('hasta cual fila');
    var adicional="";
  var opcion="";
  var column ;
  for(var iter =PosIni; iter<=PosFin;iter++){
    
    column=-1;
    
    opcion = hoja.getRange(iter, 7).getValue();
    if (opcion=="Con un archivo de EXCEL"){
      column= 26;
    }
        if (opcion=="Con una foto al formato impreso"){
      column= 205;
    }

    if (column>0){
      
     var url = hoja.getRange(iter, 239).getValue();
    var url_ID = url.substring(33,200);    //   "https://drive.google.com/open?id=);
    var file = DriveApp.getFileById(url_ID);
      
      
    
    
    file.setName(hoja.getRange(iter, 238).getValue()+adicional); //file.setName("2020-05-05__1957300078743_Los_Muñequitos_FLORIANA_MARITZA_ASPRILLA_ARBOLEDA")
    hoja.getRange(iter, 240).setValue(file.getName());
    adicional="";
    
    
    
    }

    
  }
  
  
}





function cambiandoTipoDoc_V2() {
  
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
    var url_ID = url.substring(33,200);    //   "https://drive.google.com/open?id=)
      var d= SPLIT(url,"https://drive.google.com/open?id=");
      
    var file = DriveApp.getFileById(url_ID);

   
    file.setName(hoja.getRange(iter, 238).getValue()+adicional); //file.setName("2020-05-05__1957300078743_Los_Muñequitos_FLORIANA_MARITZA_ASPRILLA_ARBOLEDA")
    hoja.getRange(iter, 240).setValue(file.getName());
    adicional="";
    
    
    
    }
  }
}





/**
 * CAMBIA XLS TO SHEET.GOOGLE, COPIA EL ARCHIVO A OTRA CUENTA Y LO ACTUALIZA EN EL REGISTRO 
 * @  Le solicita al usuario desde que fila hasta caul realizar dicha operacion
 * @   ////application/vnd.google-apps.spreadsheet      ///////application/vnd.openxmlformats-officedocument.spreadsheetml.sheet
 * @return {void} Numero Id de la Url leida de la celda O1 primera hoja 
 **/
function xlsToSheet(){
  var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var hoja = activeSpreadsheet.getSheetByName("Respuestas de formulario 1");
  var posIter = Number(Browser.inputBox("desde que fila va a realizar la operacion"));
  var posFin = Number(Browser.inputBox("hasta cual fila "));
 
  while (posIter<=posFin){
    var url = hoja.getRange(posIter, 8).getValue();
    var id= extraerID_url(url)
    
    var excelFile = DriveApp.getFileById(id);
    var tipoFile = excelFile.getMimeType();
    
    if(tipoFile=="application/vnd.google-apps.spreadsheet"){
      hoja.getRange(posIter, 239).setValue(url);
      
    }
    if(tipoFile=="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"){
       var blob = excelFile.getBlob();
    var resource = {
      title: excelFile.getName(),
      mimeType: MimeType.GOOGLE_SHEETS,
      parents: [{id:"1Y7dUwmdAdiPzYMp8XVztUaDLC0MoPa1E"}],
    };
    
   var newFile = Drive.Files.insert(resource, blob);
   hoja.getRange(posIter, 8).setValue("https://docs.google.com/spreadsheets/d/"+newFile.id);
   hoja.getRange(posIter, 240).setValue("https://docs.google.com/spreadsheets/d/"+newFile.id);
   hoja.getRange(posIter, 245).setValue(url);   

      DriveApp.removeFile(excelFile);
     

      
      
    }

    posIter++;
  }
  


}


/**
 * EXTRAE ID DE UNA URL 
 * @  
 * @param {String} cad es la url a consultar 
 * @return {String} Numero Id de la Url leida de la celda O1 primera hoja 
 **/
function extraerID_url(cad) {
  
  var array = [{}];
  var separador = "https://drive.google.com/open?id=";
  var separador2 = "https://docs.google.com/spreadsheets/d/"

  
  array = cad.split(separador);
  if(array.length==2){
    Logger.log(array.length);
    return array[1];
  }
  
    array = cad.split(separador2);
  if(array.length==2){
    Logger.log(array.length);
    return array[1];
  }
  
  
 
  return null;
}

