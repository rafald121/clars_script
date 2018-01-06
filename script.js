//TODO dodać podsumowanie do każdego typu transakcji

COLUMNS = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z']

COLOR_DOSTAWA = '#C5E89B'
COLOR_ZWROT = '#57a4f2' 
COLOR_WPLATA = '#D46A6A'

COLOR_SUMA = '#f99090'

COL_LP = 'A';COL_DATE = 'B';COL_MODEL = 'C';COL_SZTUKI = 'D';COL_CENA_SZT = 'E';COL_SUMA = 'F';INNE = 'G';
COL_SUMA_WPLATA = 'F'

COLUMNS_INDEX = {
  COL_LP:0,  COL_DATE:1,  COL_MODEL:2,  COL_SZTUKI:3,  COL_CENA_SZT:4,  COL_SUMA:5,  INNE:6,
}
COLUMNS__WPLATA_INDEX = {
  COL_LP:0,  COL_DATE:1,  COL_MODEL:2,  COL_SZTUKI:3,  COL_CENA_SZT:4,  COL_CENA_SZT_SPRZEDAZ:5,  COL_SUMA:6, INNE:7
}
DOSTAWA = 'dostawa';ZWROT = 'zwrot';WPLATA = 'wplata';SUMA = 'suma';DATA = 'Data';PODSUMOWANIE='Podsumowanie'
DOSTAWA_NAME = 'Dostawa' ; ZWROT_NAME = 'Zwrot'; WPLATA_NAME = 'Wplata' ; WSZYSTKO_NAME = 'Wszystko';PODSUMOWANIE_NAME ='Podsumowanie';
SUMY = 'SumyCzastkowe';

LEFT_CELL = 'A'
RIGHT_CELL = 'G'

var ss = SpreadsheetApp.getActiveSpreadsheet();

var sheetDostawa = ss.getSheets()[0];
var sheetZwrot = ss.getSheets()[1];
var sheetWplata = ss.getSheets()[2];
var sheetSuma = ss.getSheets()[3];
var sheetPodsumowanie = ss.getSheets()[4];
var sheetSumy = ss.getSheets()[5];


var sheetDostawa1 = ss.getSheetByName(DOSTAWA_NAME)
var sheetZwrot1 = ss.getSheetByName(ZWROT_NAME)
var sheetWplata1 = ss.getSheetByName(WPLATA_NAME)
var sheetSuma1 = ss.getSheetByName(WSZYSTKO_NAME)
var sheetPodsumowanie1 = ss.getSheetByName(PODSUMOWANIE_NAME)
var sheetSumy1 = ss.getSheetByName(SUMY)

sheetDict = {
  DOSTAWA:sheetDostawa,
  ZWROT:sheetZwrot,
  WPLATA:sheetWplata,
  SUMA:sheetSuma,
  PODSUMOWANIE:sheetPodsumowanie,
  SUMY:sheetSumy,
}

dataList = {
  DOSTAWA: sheetDostawa.getDataRange().getValues(),
  ZWROT: sheetZwrot.getDataRange().getValues(),
  WPLATA: sheetWplata.getDataRange().getValues(),
  SUMA: sheetSuma.getDataRange().getValues(),
  PODSUMOWANIE: sheetPodsumowanie.getDataRange().getValues(),
  SUMY: sheetSumy.getDataRange().getValues(),
}

//dataList123 = {
// DOSTAWA: sheetDostawa1.getDataRange().getValues(),
//  ZWROT: sheetZwrot1.getDataRange().getValues(),
 // WPLATA: sheetWplata1.getDataRange().getValues(),
//  SUMA: sheetSuma1.getDataRange().getValues(),
//  PODSUMOWANIE: sheetPodsumowanie1.getDataRange().getValues(),
//}
///////////////////////////////// FUNCTION START /////////////////////////////////

function main(){
  countDostawa()
  countZwrot()
  countWplata()
}
////////////////////////////////////////////////////
/////////////   POLICZ TABELE DOSTAWA //////////////
////////////////////////////////////////////////////
////////////////////////////////////////////////////
////////////////////////////////////////////////////
////////////////////////////////////////////////////
////////////////////////////////////////////////////
////////////////////////////////////////////////////
////////////////////////////////////////////////////
////////////////////////////////////////////////////
////////////////////////////////////////////////////

function countDostawa() {
  maxRow = dataList.DOSTAWA.length
  sumaIndicatorIndexes = getArrayOfSumaIndicators(DOSTAWA)
  dictFirstLastToSum = getFirstLastRowToSum(sumaIndicatorIndexes)
  
  for ( var sumaNr = 1 ; sumaNr<sumaIndicatorIndexes.length ; sumaNr++){ //zsumuj po kolei dla kazdej daty    // sumaNr = 1 bo pomijamy headera
    row = sumaIndicatorIndexes[sumaNr]
    row_val = dictFirstLastToSum[row] // first last 
    countSumaSztuki(row, row_val, DOSTAWA)
    countSumaSuma(row, row_val, DOSTAWA)
  }
  countSummary(sumaIndicatorIndexes)  
}
////////////////////////////////////////////////////
/////////////   POLICZ TABELE ZWROT //////////////
////////////////////////////////////////////////////
function countZwrot() {
  maxRow = dataList.ZWROT.length
  sumaIndicatorIndexes = getArrayOfSumaIndicators(ZWROT)
  dictFirstLastToSum = getFirstLastRowToSum(sumaIndicatorIndexes)
  
  for ( var sumaNr = 1 ; sumaNr<sumaIndicatorIndexes.length ; sumaNr++){ //zsumuj po kolei dla kazdej daty    // sumaNr = 1 bo pomijamy headera
    row = sumaIndicatorIndexes[sumaNr]
    row_val = dictFirstLastToSum[row] // first last 
    countSumaSztuki(row, row_val, ZWROT)
    countSumaSuma(row, row_val, ZWROT)
  }
  countSummary()
}
////////////////////////////////////////////////////
/////////////   POLICZ TABELE WPLATA ///////////////
////////////////////////////////////////////////////
function countWplata(){
  maxRow = dataList.WPLATA.length // liczba wierszy
  sumaIndicatorIndexes = getArrayOfSumaIndicators(WPLATA) // indexy wierszy które sumujemy
  dictFirstLastToSum = getFirstLastRowToSum(sumaIndicatorIndexes) // słownik pierwszego i ostatniego wiersza z kazdej sumy
  
  for ( var sumaNr = 1 ; sumaNr<sumaIndicatorIndexes.length ; sumaNr++){ //zsumuj po kolei dla kazdej daty    // sumaNr = 1 bo pomijamy headera
    row = sumaIndicatorIndexes[sumaNr]  // wiersz z sumą na daną iterację // 6, 12, 18, 23//
    row_val = dictFirstLastToSum[row] // dla 6: [2,5], dla 11: [7,10] 
    countSumaSztuki(row, row_val, WPLATA)
    countSumaSumaWplata(row, row_val, WPLATA)
  }
  
  countSummary()
}
////////////////////////////////////////////////////
/////////////   POLICZ TABELE WSZYSTKO /////////////
////////////////////////////////////////////////////
function countWszystko(){
  //countDostawa()
  //countZwrot()
  //countWplata()
  cellLp = joinCell('A',1)
  cellTyp = joinCell('B',1)
  cellData = joinCell('C',1)
  cellModel = joinCell('D',1)
  cellSztuki = joinCell('E',1)
  cellCenaNetto = joinCell('F',1)
  cellSuma = joinCell('G',1)
  cellCenaBrutto = joinCell('H',1)

  getSheetType(SUMA).getRange(cellLp).setValue('Lp')
  getSheetType(SUMA).getRange(cellTyp).setValue('Typ')
  getSheetType(SUMA).getRange(cellData).setValue('Data')
  getSheetType(SUMA).getRange(cellModel).setValue('Model')
  getSheetType(SUMA).getRange(cellSztuki).setValue('Sztuki')
  getSheetType(SUMA).getRange(cellCenaNetto).setValue('Cena/Sztuka netto')
  getSheetType(SUMA).getRange(cellSuma).setValue('Suma')
  getSheetType(SUMA).getRange(cellCenaBrutto).setValue('Cena/sztuka brutto')

  firstDateDostawa = getFirstDate(DOSTAWA)
  firstDateZwrot = getFirstDate(ZWROT)
  firstDateWplata = getFirstDate(WPLATA);
  firstDate = firstDateDostawa
  firstDates = [firstDateDostawa, firstDateZwrot, firstDateWplata]
  
  lastDateDostawa = getLastDate(DOSTAWA);
  lastDateZwrot = getLastDate(ZWROT); 
  lastDateWplata = getLastDate(WPLATA);
  lastDate = lastDateDostawa
  lastDates = [lastDateDostawa, lastDateZwrot, lastDateWplata];
  for(var i = 0 ; i<3 ; i++){
   if(firstDates[i] < firstDate){
     firstDate = firstDates[i]
   }
   if(lastDates[i] > lastDate){
     lastDate = lastDates[i]
   }                   
 }
 
  maxRowWplata = dataList.WPLATA.length // liczba wierszy
  sumaIndicatorIndexesWplata = getArrayOfSumaIndicators(WPLATA) // indexy wierszy które sumujemy
  dictFirstLastToSumWplata = getFirstLastRowToSum(sumaIndicatorIndexesWplata) // słownik pierwszego i ostatniego wiersza z kazdej sumy
  
  maxRowZwrot = dataList.ZWROT.length
  sumaIndicatorIndexesZwrot = getArrayOfSumaIndicators(ZWROT)
  dictFirstLastToSumZwrot = getFirstLastRowToSum(sumaIndicatorIndexesZwrot)
  
  maxRowDostawa = dataList.DOSTAWA.length
  sumaIndicatorIndexesDostawa = getArrayOfSumaIndicators(DOSTAWA)  // pokazuje liczba wierszy z sumą +1 bo header
  dictFirstLastToSumDostawa = getFirstLastRowToSum(sumaIndicatorIndexesDostawa)
  // daty z danej tabeli
  datesDostawa = fillListWithDatesToConsider(dictFirstLastToSumDostawa, DOSTAWA) 
  datesZwrot   = fillListWithDatesToConsider(dictFirstLastToSumZwrot, ZWROT)
  datesWplata  = fillListWithDatesToConsider(dictFirstLastToSumWplata, WPLATA)
  //
  datesDictDostawa = fillDictWithDatesToConsider(dictFirstLastToSumDostawa, DOSTAWA) 
  datesDictZwrot   = fillDictWithDatesToConsider(dictFirstLastToSumZwrot, ZWROT)
  datesDictWplata  = fillDictWithDatesToConsider(dictFirstLastToSumWplata, WPLATA)
  
  tableTypeDataFirstLast = [] //tablica tablic o formie: typ, data, first, last
  for (var key in datesDictDostawa){
    detail = [DOSTAWA,key,datesDictDostawa[key][0],datesDictDostawa[key][1]]
    tableTypeDataFirstLast.push(detail)
  }
  for (var key in datesDictZwrot){
    detail = [ZWROT,key,datesDictZwrot[key][0],datesDictZwrot[key][1]]
    tableTypeDataFirstLast.push(detail)
  }
  for (var key in datesDictWplata){
    detail = [WPLATA,key,datesDictWplata[key][0],datesDictWplata[key][1]]
    tableTypeDataFirstLast.push(detail)
  }  
  
  sortedTableTypeDataFirstLast = sortDates(tableTypeDataFirstLast, firstDate)
  
  rowCounter = 2 // rowCounter==2, bo row=1 to header
  counter = 1
  for(var l = 0 ; l < sortedTableTypeDataFirstLast.length ; l++){
    last = sortedTableTypeDataFirstLast[l][3];
    first =  sortedTableTypeDataFirstLast[l][2];
    rowAmountToAdd = last-first+1
    unpackAndPaste(sortedTableTypeDataFirstLast[l], rowCounter, counter);//dodaj do większego od jeden niż ostatni zapełniony 
    rowCounter += rowAmountToAdd;
    counter+=rowAmountToAdd
  }
  
  var lastRow = counter+1
  allSumaCell = joinCell('G',lastRow)
  allSztukiCell = joinCell('E',lastRow)
  
  sumaFirst = joinCell('G', 2)
  sumaLast = joinCell('G', lastRow-1)
  sztukiFirst = joinCell('E', 2)
  sztukiLast = joinCell('E', lastRow-1)
  getSheetType(SUMA).getRange(allSumaCell).setValue(sumFormula(sumaFirst,sumaLast))
  getSheetType(SUMA).getRange(allSztukiCell).setValue(sumFormula(sztukiFirst, sztukiLast))
  
  
  var a = 1
  
}


function countPodsumowanie(){
  //DOSTAWA//
  var dostawaData = dataList.DOSTAWA;
  var dostawaSumaArray = getArrayOfSumaIndicators(DOSTAWA)
  var lastRowNumber = dostawaSumaArray[dostawaSumaArray.length-1]//dokąd liczymy
  var modelsList = {}
  var modelsPriceList = {}
  for(var i = 1 ; i<=lastRowNumber ; i++){ // i =2 bo 1 = header
    var modelName = dostawaData[i][2];
    var sztuki = dostawaData[i][3];
    var cenaSztuka = dostawaData[i][4]
    
    if (modelName != '') { //jeśli sie rowna '' to pomiń bo to wiersz sumy
    
      if (modelName in modelsList){ // jesli jest juz taki model to dodaj sztuki
        var currentAmount = modelsList[modelName][0]
        if (currentAmount != undefined){
          modelsList[modelName][0] = currentAmount+sztuki
        }else{
          modelsList[modelName][0] = sztuki
        }
      } else { // jeśli nie ma jeszcze
        modelsList[modelName] = new Array(3)
        modelsList[modelName][0] = sztuki
        modelsPriceList[modelName] = cenaSztuka
      }
    }
  }
  //ZWROT//
  var zwrotData = dataList.ZWROT;
  var zwrotSumaArray = getArrayOfSumaIndicators(ZWROT)
  var lastRowNumber = zwrotSumaArray[zwrotSumaArray.length-1]
//  var modelsList = {}
  for(var i = 1 ; i<=lastRowNumber ; i++){ // i =2 bo 1 = header
    var modelName = zwrotData[i][2];
    var sztuki = zwrotData[i][3];
    
    if (modelName != '') { //jeśli sie rowna '' to pomiń
    
      if (modelName in modelsList){ // jesli jest juz taki model to dodaj sztuki
        var currentAmount = modelsList[modelName][1]
        if (currentAmount != undefined){
          modelsList[modelName][1] = currentAmount+sztuki
        }else{
          modelsList[modelName][1] = sztuki
        }
      } else { // jeśli nie ma jeszcze
        modelsList[modelName] = new Array(3);
        modelsList[modelName][1] = sztuki
      }
    }
  }
  // WPLATA //
  
  var wplataData = dataList.WPLATA;
  var wplataSumaArray = getArrayOfSumaIndicators(WPLATA)
  var lastRowNumber = wplataSumaArray[wplataSumaArray.length-1]
  //  var modelsList = {}
  for(var i = 1 ; i<=lastRowNumber ; i++){ // i =2 bo 1 = header
    var modelName = wplataData[i][2];
    var sztuki = wplataData[i][3];
    
    if (modelName != '') { //jeśli sie rowna '' to pomiń
    
      if (modelName in modelsList){ // jesli jest juz taki model to dodaj sztuki
        var currentAmount = modelsList[modelName][2]
        if (currentAmount != undefined){
          modelsList[modelName][2] = currentAmount+sztuki
        }else{
          modelsList[modelName][2] = sztuki
        }
      } else { // jeśli nie ma jeszcze
        modelsList[modelName] = new Array(3);
        modelsList[modelName][2] = sztuki
      }
    }
  }
  iterator = 2 // bo =1 to header
  cellLp1 = joinCell('A',1)
  cellModel1 = joinCell('B',1)
  cellDostawa1 = joinCell('C',1)
  cellZwrot1 = joinCell('D',1)
  cellWplata1 = joinCell('E',1)
  cellSuma1 = joinCell('F',1)
  cellCenaSukienki1 = joinCell('G',1)
  cellSumaNaSklepie1 = joinCell('H',1)
  getSheetType(PODSUMOWANIE).getRange(cellLp1).setValue('Lp')
  getSheetType(PODSUMOWANIE).getRange(cellModel1).setValue('Model')
  getSheetType(PODSUMOWANIE).getRange(cellDostawa1).setValue('Dostawa')
  getSheetType(PODSUMOWANIE).getRange(cellZwrot1).setValue('Zwrot')
  getSheetType(PODSUMOWANIE).getRange(cellWplata1).setValue('Wplata')
  getSheetType(PODSUMOWANIE).getRange(cellSuma1).setValue('Zostało sztuk')    
  getSheetType(PODSUMOWANIE).getRange(cellCenaSukienki1).setValue('Cena za sztuke')   
  getSheetType(PODSUMOWANIE).getRange(cellSumaNaSklepie1).setValue('Wartość sukienek w sklepie')   
  
  for (var model in modelsList){
  
    cellLp = joinCell('A',iterator)
    cellModel = joinCell('B',iterator)
    cellDostawa = joinCell('C',iterator)
    cellZwrot = joinCell('D',iterator)
    cellWplata = joinCell('E',iterator)
    cellSuma = joinCell('F',iterator)
    cellCenaSztuka = joinCell('G',iterator)
    cellSumaNaSklepie = joinCell('H',iterator)
    
    getSheetType(PODSUMOWANIE).getRange(cellLp).setValue(iterator-1)
    getSheetType(PODSUMOWANIE).getRange(cellModel).setValue(model)
    if (modelsList[model][0] != undefined)
      getSheetType(PODSUMOWANIE).getRange(cellDostawa).setValue(modelsList[model][0])
    if (modelsList[model][1] != undefined)
      getSheetType(PODSUMOWANIE).getRange(cellZwrot).setValue(modelsList[model][1])
    if (modelsList[model][2] != undefined)
      getSheetType(PODSUMOWANIE).getRange(cellWplata).setValue(modelsList[model][2])
//    if (modelsList[model][3] != undefined)
    
    getSheetType(PODSUMOWANIE).getRange(cellCenaSztuka).setValue(modelsPriceList[model])
    
    suma = 0 
    for (var i = 0 ; i < 3 ; i ++){
      value = modelsList[model][i]
      if (value != undefined){
        suma += value;
      }
    }
    getSheetType(PODSUMOWANIE).getRange(cellSuma).setValue(suma)
    getSheetType(PODSUMOWANIE).getRange(cellSumaNaSklepie).setValue(suma*modelsPriceList[model])
    
    iterator +=1
  }
  
  cellSumaSztukiFirst = joinCell('F', 2)
  cellSumaSztukiLast = joinCell('F', iterator-1)
  cellSumaSumaFirst = joinCell('H', 2)
  cellSumaSumaLast = joinCell('H', iterator-1)
  cellSumaSztuki = joinCell('F', iterator)
  cellSumaSuma = joinCell('H', iterator)
  
  var sumaSztukiFormula = sumFormula(cellSumaSztukiFirst, cellSumaSztukiLast)
  var sumaSumaFormula = sumFormula(cellSumaSumaFirst, cellSumaSumaLast)
  
  getSheetType(PODSUMOWANIE).getRange(cellSumaSztuki).setValue(sumaSztukiFormula)
  getSheetType(PODSUMOWANIE).getRange(cellSumaSuma).setValue(sumaSumaFormula)
}


function countSumy(){
    
    var sumaDostawaIndicatorIndexes = getArrayOfSumaIndicators(DOSTAWA);
    var sumaDostawaSztuki = 0; 
    var sumaDostawaSuma = 0;
    for(var i = 1 ; i < sumaDostawaIndicatorIndexes.length ; i++){//i=1 bo pomijamy "Data" header
      var cellSztuki = joinCell('D',sumaDostawaIndicatorIndexes[i]);
      var cellSuma = joinCell('F',sumaDostawaIndicatorIndexes[i]);
      var sztuka = getSheetType(DOSTAWA).getRange(cellSztuki).getValue();
      var suma = getSheetType(DOSTAWA).getRange(cellSuma).getValue();
      sumaDostawaSztuki = sumaDostawaSztuki + sztuka;
      sumaDostawaSuma = sumaDostawaSuma + suma;
    }
    var sumaZwrotIndicatorIndexes = getArrayOfSumaIndicators(ZWROT);
    var sumaZwrotSztuki = 0;
    var sumaZwrotSuma = 0;
    for(var i = 1 ; i < sumaZwrotIndicatorIndexes.length ; i++){
      var cellSztuki = joinCell('D',sumaZwrotIndicatorIndexes[i]);
      var cellSuma = joinCell('F',sumaZwrotIndicatorIndexes[i]);
      var sztuka = getSheetType(ZWROT).getRange(cellSztuki).getValue();
      var suma = getSheetType(ZWROT).getRange(cellSuma).getValue();
      sumaZwrotSztuki = sumaZwrotSztuki + sztuka;
      sumaZwrotSuma = sumaZwrotSuma + suma;
    }
    var sumaWplataIndicatorIndexes = getArrayOfSumaIndicators(WPLATA);
    var sumaWplataSztuki = 0;
    var sumaWplataSuma = 0;
    for(var i = 1 ; i < sumaWplataIndicatorIndexes.length ; i++){
      var cellSztuki = joinCell('D',sumaWplataIndicatorIndexes[i]);
      var cellSuma = joinCell('F',sumaWplataIndicatorIndexes[i]);
      var sztuka = getSheetType(WPLATA).getRange(cellSztuki).getValue();
      var suma = getSheetType(WPLATA).getRange(cellSuma).getValue();
      sumaWplataSztuki = sumaWplataSztuki + sztuka;
      sumaWplataSuma = sumaWplataSuma + suma;
    }
    
    getSheetType(SUMY).getRange('B1').setValue('Sztuki')
    getSheetType(SUMY).getRange('C1').setValue('Suma')
    getSheetType(SUMY).getRange('A2').setValue('Dostawa')
    getSheetType(SUMY).getRange('A3').setValue('Zwrot')
    getSheetType(SUMY).getRange('A4').setValue('Wplata netto')
    getSheetType(SUMY).getRange('A5').setValue('Wplata brutto')
    
    cellDostawaSztuki = joinCell('B','2')
    cellDostawaSuma = joinCell('C','2')
    cellZwrotSztuki = joinCell('B','3')
    cellZwrotSuma = joinCell('C','3')  
    cellWplataNettoSztuki = joinCell('B', '4')
    cellWplataNettoSuma = joinCell('C', '4')
    cellWplataBruttoSztuki = joinCell('B','5')  
    cellWplataBruttoSuma = joinCell('C','5')  

    
    procent = getSheetType(SUMY).getRange('E1').getValue()
    
    getSheetType(SUMY).getRange(cellDostawaSztuki).setValue(sumaDostawaSztuki)
    getSheetType(SUMY).getRange(cellDostawaSuma).setValue(sumaDostawaSuma)
    getSheetType(SUMY).getRange(cellZwrotSztuki).setValue(-sumaZwrotSztuki)
    getSheetType(SUMY).getRange(cellZwrotSuma).setValue(-sumaZwrotSuma)
   
    getSheetType(SUMY).getRange(cellWplataNettoSztuki).setValue(-sumaWplataSztuki)
    getSheetType(SUMY).getRange(cellWplataNettoSuma).setValue(-sumaWplataSuma) 
    
    getSheetType(SUMY).getRange(cellWplataBruttoSztuki).setValue(-sumaWplataSztuki)
    getSheetType(SUMY).getRange(cellWplataBruttoSuma).setValue(-sumaWplataSuma/(1-procent)) 
    

//    getSheetType(SUMY).getRange(wplataSuma).setValue(-sumaWplataSuma) 
    
}
////////////////////////////////////////////////////
////////////////////////////////////////////////////
////////////////////////////////////////////////////
////////////////////////////////////////////////////
////////////////////////////////////////////////////
/////////////// UTILS LICZENIE WSZYSTKIEGO /////////
////////////////////////////////////////////////////
////////////////////////////////////////////////////
////////////////////////////////////////////////////
////////////////////////////////////////////////////
////////////////////////////////////////////////////


function unpackAndPaste(row, rowCounter, counter){//row
// first/last CellTo - pierwsza/ostatnia komórka do której będziemy doklejać 
// first/last CellFrom - pierwsza/ostatnia komórka z której będziemy doklejać  

  type = row[0]
  if(type == DOSTAWA){
    color=COLOR_DOSTAWA
  } else if(type == ZWROT){
    color=COLOR_ZWROT
  } else if(type == WPLATA){
    color=COLOR_WPLATA
  }
  
  // KOPIOWANIE
  counterRow = 0 //liczy ile wierszy wkleja
  for(var i = row[2] ; i<=row[3] ; i++){
    firstCellFrom = joinCell(LEFT_CELL, i)
    lastCellFrom = joinCell(RIGHT_CELL, i)
    rangeFrom = rangeFormula(firstCellFrom, lastCellFrom)
    
    firstCellTo = joinCell(LEFT_CELL, rowCounter+counterRow)
    lastCellTo = joinCell(RIGHT_CELL, rowCounter+counterRow)
    rangeTo = rangeFormula(firstCellTo, lastCellTo)
    
    lpCell = joinCell('A',counter+1)
    typeCell = joinCell('B', rowCounter+counterRow)
    
    var sourceRowFrom = getSheetType(type).getRange(i, 1, 1, 8)
    var targetRowTo = getSheetType(SUMA).getRange(rowCounter+counterRow, 2, 1, 8)

    sourceRowFrom.copyTo(targetRowTo)    

//KOLORKI do sumy
    getSheetType(SUMA).getRange(lpCell).setValue(counter)
    getSheetType(SUMA).getRange(typeCell).setValue(type)
    getSheetType(SUMA).getRange(rowCounter+counterRow, 1, 1, 8).setBackground(color);
    
    counterRow++
    counter++
  }

}

function sortDates(_tableTypeDataFirstLast, _firstDate){
  tableTypeDataFirstLast = []
  for (var k = 0 ; k<_tableTypeDataFirstLast.length; k++){
    tableTypeDataFirstLast.push(_tableTypeDataFirstLast[k])
  }
  len = tableTypeDataFirstLast.length
  var table = []
  
  for(var i = 0 ; i < len ; i++){ // pętla do całej tablicy z danymi
  
    date = tableTypeDataFirstLast[0][1]
    earliest_date = [new Date(date), 0]  //data i numer indeksu w oryginale
    var j = 0
    for (j = 0 ; j < tableTypeDataFirstLast.length ; j++){

      date = tableTypeDataFirstLast[j][1]
      currently_checking_date = new Date(date)       

      if (currently_checking_date.getTime() < earliest_date[0].getTime()){
        earliest_date = [currently_checking_date, j]
      }
     
    }
    
    table.push(tableTypeDataFirstLast[earliest_date[1]])
    
    tableTypeDataFirstLast.splice(earliest_date[1],1)

    
  }
  return table
}

//lista dat, z każdego typu tabelu (dostawa,zwrto,wplata), którą pozniej uwzgledniamy w tabeli wszystko
function fillListWithDatesToConsider(dict, sheetType){
  dateList = []
  for (var key in dict){
    targetCell = joinCell(COL_DATE, dict[key][0])
    date = getSheetType(sheetType).getRange(targetCell).getValue()
    dateList.push(date)
  }
  return dateList //zwraca wszystkie daty 
}
// zwraca slownik  Date : [first_row,last_row] <- now we know which row copy to "wszystko" table 
function fillDictWithDatesToConsider(dict, sheetType){
  dateDict = {}
  for (var key in dict){
    targetCell = joinCell(COL_DATE, dict[key][0])
    date = getSheetType(sheetType).getRange(targetCell).getValue()
    dateDict[date] = dict[key]
  }
  return dateDict  
}

function getFirstDate(){
  
}


//////////////WEZ INDEKSY GRANICZNE OD DO KTORYCH TRZEBA ZSUMOWAC//////////////////////////////////////
function getFirstLastRowToSum(sumaIndicatorIndexes){
  var i;
  dict = {}
  
  for (i=1;i<sumaIndicatorIndexes.length;i++){ // nie uwzgledniamy '1' bo to headery
    first = sumaIndicatorIndexes[i-1]+1 
    last = sumaIndicatorIndexes[i]-1
    dict[sumaIndicatorIndexes[i]] = [first,last]
  }
  return dict
}
/////SUMY,AVGY//////////////////

function setAmountIndex(indexes){

}
function countSumaSztuki(row, firstLast, sheetType){
  firstCell = joinCell(COL_SZTUKI,firstLast[0])
  lastCell = joinCell(COL_SZTUKI,firstLast[1])
  targetCell = joinCell(COL_SZTUKI, row)
  getSheetType(sheetType).getRange(targetCell).setValue(sumFormula(firstCell,lastCell));
  getSheetType(sheetType).getRange(targetCell).setBackground(COLOR_SUMA)  
}
function countSumaSuma(row, firstLast, sheetType){
  firstCell = joinCell(COL_SUMA,firstLast[0])
  lastCell = joinCell(COL_SUMA,firstLast[1])
  targetCell = joinCell(COL_SUMA, row)
  getSheetType(sheetType).getRange(targetCell).setValue(sumFormula(firstCell,lastCell));
  getSheetType(sheetType).getRange(targetCell).setBackground(COLOR_SUMA);
}
function countSumaSumaWplata(row, firstLast, sheetType){
  firstCell = joinCell(COL_SUMA_WPLATA,firstLast[0])
  lastCell = joinCell(COL_SUMA_WPLATA,firstLast[1])
  targetCell = joinCell(COL_SUMA_WPLATA, row)
  getSheetType(sheetType).getRange(targetCell).setValue(sumFormula(firstCell,lastCell));
  getSheetType(sheetType).getRange(targetCell).setBackground(COLOR_SUMA);
}

function countSummary(){
}

function countAvgCenaSzt(indexes){ 
}

/////NUMERY WIERSZY W KTORYCH SUMUJEMY//////////////////

function getArrayOfSumaIndicators(sheetType) {//sheetType = DOSTAWA,WPLATA,...
  data = getDataType(sheetType)
  var row;
  sumaIndicateIndexes = [];
  for(row=0;row<data.length; row++){
    var varsuma = data[row][COLUMNS_INDEX.COL_DATE];
    if(varsuma===SUMA || varsuma===DATA)  
      sumaIndicateIndexes.push(row+1);//+1 bo spreadsheet ma indexy od 1, a nie od 0
  }
  return sumaIndicateIndexes;
}

////UTILS ////////////////////////////////////////////

function getFirstDate(sheetType){
  return getSheetType(sheetType).getRange('B2').getValue();
}

function getLastDate(sheetType){
  data = getDataType(sheetType)
  lastDateCell = joinCell('B', data.length-2)
  return getSheetType(sheetType).getRange(lastDateCell).getValue();
}

function fillArray(from,to){
  arraj = []
  for(i=from;i<to+1;i++)
    arraj.push(i);
  return arraj;
}

function sumFormula(first,last){
  return "=SUM("+first+":"+last+")";
}

function rangeFormula(first, last){
  return first+":"+last
}

function joinCell(col,row){ // np 6 + F = F6
  return col + row.toString()
}

///GETTERY /////////////////////////////////////////

function getSheetType(sheetType){
  data = null;
  if (sheetType == DOSTAWA)
    data = sheetDict.DOSTAWA
  else if(sheetType == ZWROT)
    data = sheetDict.ZWROT
  else if(sheetType == WPLATA)
    data = sheetDict.WPLATA
  else if(sheetType == SUMA)
    data = sheetDict.SUMA
  else if(sheetType == PODSUMOWANIE)
    data = sheetDict.PODSUMOWANIE
  else if(sheetType == SUMY)
    data = sheetDict.SUMY
  else
    data = 'blad'
    
   return data
}

function getDataType(sheetType){
  data = null
  if (sheetType == DOSTAWA)
    data = dataList.DOSTAWA
  else if(sheetType == ZWROT)
    data = dataList.ZWROT
  else if(sheetType == WPLATA)
    data = dataList.WPLATA
  else if(sheetType == SUMA)
    data = dataList.SUMA
  else if(sheetType == PODSUMOWANIE)
    data = dataList.PODSUMOWANIE
  else if(sheetType == SUMY)
    data = sheetDict.SUMY
  else
    data = 'blad'
    
   return data
}
