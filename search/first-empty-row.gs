/*
Funkce onOpen() zajisti spusteni vlozeneho kodu pri kazdem otevreni tabulky uzivatelem s opravnenim upravovat
Vytvorime dve menu - sloupcove a celkove vyhledavani
*/
function onOpen() {
  SpreadsheetApp.getUi() 
      .createMenu('Table tools')
      .addItem('Locate first empty row in column', 'firstEmptyRowInColumn')
      .addItem('Locate first empty row in table', 'firstEmptyRowInTable')
      .addToUi();
}

// Vyhledavani prazdneho radku ve sloupci
function firstEmptyRowInColumn() {
  var sheet = SpreadsheetApp.getActive();
  var column = sheet.getRange('A:A');
  var values = column.getValues();
  
  var emptyRow = 0;
  
  while ( values[emptyRow][0] != "" ) {
    emptyRow++;
  }
  
  SpreadsheetApp.getUi().alert(emptyRow+1);
}

// Vyhledavani prazdneho radku v cele tabulce
function firstEmptyRowInTable() {
  var sheet = SpreadsheetApp.getActive();
  var emptyRow = sheet.getLastRow();
  
  SpreadsheetApp.getUi().alert(emptyRow+1);
}