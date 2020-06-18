function mimacro1() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('D6').activate();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Respuestas de formulario 1'), true);
  spreadsheet.getRange('D6').activate();
  elegirCelda();
  spreadsheet.getRange('IE321').activate();
}


function test(){

var d="https://docs.google.com/spreadsheets/d/1DpCXk-D5FHB5Zjx-9l_mUMU4NzYP4cM-9DVr7katiEs/edit#gid=1355821151";
  var url_id= dividiendo(d);
  
      var file = DriveApp.getFileById(url_id);
  Browser.msgBox(file.getName());
}


function test1(){
 var  hojaForm=SpreadsheetApp.getActive().getSheetByName("Respuestas de formulario 1");
  var hojaDes=SpreadsheetApp.getActive().getSheetByName("test");    //1131

  var dt =   hojaForm.getRange(770, 30,1, 24).getValues();
  
  
    
   hojaDes.getRange(3, 5,1,24).setValues(dt);
  
  
  var dt2 =[];
  var d  =[];
  
  
  d.push("maria ");
    d.push("mela ");
    d.push("suda ");
dt2.push(d);
  dt2.
  
  
     hojaDes.getRange(4, 5,1,3).setValues(dt2);
  Logger.log(dt);
     Logger.log(dt2);

}