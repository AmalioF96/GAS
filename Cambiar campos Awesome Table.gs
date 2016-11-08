function onEdit(e) {

    var ssArticles = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Articles');
    var ssList = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('List');
    var dataArticle = ssArticles.getDataRange().getValues();
    var dataList = ssList.getDataRange().getValues();
    //Variables formulario
    var i = 2;
    
    var ultimafila = dataArticle.length+1;
   
    for (i = 3; i < ultimafila; i++) {
      var j = 1;
      var asignatura = ssArticles.getRange(i, 1).getValue();
      var elementoLista =ssList.getRange(j, 1).getValue();
      
        //var asignatura = dataArticle[i][2];
        //elementoLista=dataList[j][0];
      
        while (asignatura != elementoLista && j<dataList.length) {
            j++;
          elementoLista=ssList.getRange(j, 1).getValue();
        }
      ssArticles.getRange(i, 3).setValue(ssList.getRange(j, 2).getValue());
    }
    
}
