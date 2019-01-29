//Select the mouse ID and genotypes to transfer to the litter sheet



LITTERSPREADSHEET_ID = "id_goes_here";

function onOpen() {
SpreadsheetApp.getUi().createMenu('Transfer Genotypes')
.addItem('Transfer Current Sheet', "transferSheet")
.addToUi()
}




function transferSheet() {
  var geneSheet = SpreadsheetApp.getActiveSheet();
  var litterSpread = SpreadsheetApp.openById(LITTERSPREADSHEET_ID);
  var data = geneSheet.getActiveRange().getValues();
  
  var dataWidth = data[0].length;
  var dataHeight = data.length;
  
  
  if(dataWidth > 7){
    Browser.msgBox("Your selection is too wide, it will write into the 'Comments' and/or 'Age' columns in the litter sheet.");
    return;
  }
  
  //regex to break apart the mouse ID into [1] = litter, [2] = litter number, [3] = mouse ID
  var regEx = /^([A-za-z]+).(\d+)\.(\d+)$/

  var currentLitter;
  var lastLitter;
  
  for(i=0; i<dataHeight; i++) {
    if(data[i][0] == "") {
      //breaks the loop if there isn't a litter in the first colomn
      break;
    }
    
    var littermatch = regEx.exec(data[i][0]);
           
    currentLitter = littermatch[1];

    

      
    var litterSheet = litterSpread.getSheetByName(currentLitter.toUpperCase());
    
    
    var litterData = litterSheet.getRange(1, 2, litterSheet.getLastRow()).getValues();
    Logger.log(litterData.length);
    var litterStart = 0;
    var litterSize = 0;
    
    //loops through the first line of the litter spreadsheet, looks for the current litter number
    //counts the number of mice in that litter, and records the start index
    for(j=0; j<litterData.length; j++) {
      if(litterData[j][0] == ""){
        break;
      }
      if(i+litterSize >= dataHeight){
        break;
      }
      Logger.log(litterData[j][0]);
      Logger.log(litterSize);
      if(litterData[j][0] == data[i+litterSize][0]){
        Logger.log(data[i+litterSize][0]);

        Logger.log(litterStart);
        litterSize += 1;
        if(litterStart == 0){
          litterStart = j+1;
        }
      }
    }

    SpreadsheetApp.setActiveSpreadsheet(litterSpread);
    var toWrite = data.slice(i, i+litterSize);
    var writeValues = [];
    
    for(q=0; q<toWrite.length; q++){
      writeValues.push(toWrite[q].slice(1, dataWidth));
    }
    
    litterSheet.setActiveRange(litterSheet.getRange(litterStart, 14, litterSize, dataWidth-1)).setValues(writeValues);
    
    i += litterSize-1;

      
    }
    
     
}


