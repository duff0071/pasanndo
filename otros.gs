function numeroDeCeldasOcupadas(){
   var formatThousandsNoRounding = function(n, dp){
       var e = '', s = e+n, l = s.length, b = n < 0 ? 1 : 0,
           i = s.lastIndexOf('.'), j = i == -1 ? l : i,
           r = e, d = s.substr(j+1, dp);
       while ( (j-=3) > b ) { r = ',' + s.substr(j, 3) + r; }
       return s.substr(0, j + 3) + r +
           (dp ? '.' + d + ( d.length < dp ?
                   ('00000').substr(0, dp - d.length):e):e);
   };
   var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets()
   var cells_count = 0;
   for (var i in sheets){
       cells_count += (sheets[i].getMaxColumns() * sheets[i].getMaxRows());
   }
   var percentageCells = Math.round((100*(cells_count/2000000)));
   var completeStringInMsg = (formatThousandsNoRounding(cells_count) + " CELDAS (" + percentageCells + " % DEL ESPACIO DISPONIBLE)");
   Logger.log(formatThousandsNoRounding(cells_count))
   Browser.msgBox("Currently using:", completeStringInMsg, Browser.Buttons.OK);
}

function openTab() {
  var selection = SpreadsheetApp.getActiveSheet().getActiveCell().getValue();
  
  var html = "<script>window.open('" + selection + "');google.script.host.close();</script>";
  
  var userInterface = HtmlService.createHtmlOutput(html);
  
  SpreadsheetApp.getUi().showModalDialog(userInterface, 'Open Tab');
}

function numeroDeCeldasOcupadas(){
  var formatThousandsNoRounding = function(n, dp){
    var e = '', s = e+n, l = s.length, b = n < 0 ? 1 : 0,
      i = s.lastIndexOf('.'), j = i == -1 ? l : i,
        r = e, d = s.substr(j+1, dp);
    while ( (j-=3) > b ) { r = ',' + s.substr(j, 3) + r; }
    return s.substr(0, j + 3) + r +
      (dp ? '.' + d + ( d.length < dp ?
       ('00000').substr(0, dp - d.length):e):e);
  };
  var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets()
  var cells_count = 0;
  for (var i in sheets){
    cells_count += (sheets[i].getMaxColumns() * sheets[i].getMaxRows());
  }
  var percentageCells = Math.round((100*(cells_count/2000000)));
  var completeStringInMsg = (formatThousandsNoRounding(cells_count) + " CELDAS (" + percentageCells + " % DEL ESPACIO DISPONIBLE)");
  Logger.log(formatThousandsNoRounding(cells_count))
  Browser.msgBox("Currently using:", completeStringInMsg, Browser.Buttons.OK);
}