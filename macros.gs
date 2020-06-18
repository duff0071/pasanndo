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


  Browser.msgBox(NOMPROPIO("era mas facil de lo evidente"));

}