//trigger for writeToSheets is a form submission
//litterNames() is triggered by a form submission and form open


//Global variables -- Change these if a new spreadsheet is added or if this moves to another spreadsheet
//The ID is found in the URL, after /d/ and before /edit (https://docs.google.com/spreadsheets/d/1to6a4qqZlSkt5iW0PTBZNDjVH14GFdIjA5QHl9bXi70/edit#gid=129785279)

LITTER_SPREADSHEET_ID = "id_goes_here";
LITTER_LIST_SHEET_NAME = "List_of_all";
LITTER_TEMPLATE_SHEET_NAME = "Template";
GENOTYPING_SPREADSHEET_ID = "another_id_here";
GENOTYPING_TEMPLATE_SHEET_NAME = "Template";
FORM_RESPONSE_ID = "another_id_here";
FORM_RESPONSE_SHEET_NAME = "Form Responses 1";
GENOTYPING_PLATE_PREFIX = "WR "; //add a space at the end if you want "WR 12" instead of "WR12"
LITTER_NAME_ON_FORM = "Litter Name"; //title of the multiple choice item for litters on the Form 





function litterNames(){
  //Checks the litter spreadsheet and populates the form with mouse litter names and current number as found in the litter spreadsheet "List_of_all"
  
  var plateForm = FormApp.getActiveForm();
  var items = plateForm.getItems();
  var litterSpread = SpreadsheetApp.openById(LITTER_SPREADSHEET_ID);
  var choiceArray = litterSpread.getSheetByName(LITTER_LIST_SHEET_NAME).getDataRange().getValues();  
  var newArray = [];

  for(i=1; i<choiceArray.length; i++) {
    newArray.push(choiceArray[i][0] + " " + (choiceArray[i][1] + 1));
  }
  

  
  for (var i = 0; i < items.length; i += 1){
    var item = items[i]
      if (item.getTitle() == LITTER_NAME_ON_FORM){
        item.asMultipleChoiceItem().setChoiceValues(newArray);
        break;
    }
  }
}




function writeToSheets() {
  //This function is triggered when the form is submitted
  //It reads the most recently submitted response on the response sheet and updates the Litter Sheet and Genotyping spreadsheet with that info
  //It gets a little complicated when it adds the location of the toes to the litter sheet. Has to toggle back and forth 
  
  
  
  //set the spreadsheet variables
  var genotypingSpread = SpreadsheetApp.openById(GENOTYPING_SPREADSHEET_ID);
  var litterSpread = SpreadsheetApp.openById(LITTER_SPREADSHEET_ID);
  var responseSpread = SpreadsheetApp.openById(FORM_RESPONSE_ID);
  var responseSheet = responseSpread.getSheetByName(FORM_RESPONSE_SHEET_NAME);
  var list_of_all = litterSpread.getSheetByName(LITTER_LIST_SHEET_NAME);
  
  
  //values from recently completed form. These are dependent on position (i.e. the date of birth is in column 8 (so 7 in the zero-based array))
  var responseValues = responseSheet.getRange(responseSheet.getLastRow(), 1, 1, 26).getValues();
  var plateNum = parseInt(responseValues[0][1]);
  var plate = GENOTYPING_PLATE_PREFIX + plateNum;
  var litter = responseValues[0][2];
  var father = responseValues[0][3];
  var mother1 = responseValues[0][4];
  var mother2 = responseValues[0][5];
  var pups = parseInt(responseValues[0][6]);
  var dateOfBirth = responseValues[0][7];

  //parse the litter name with regex to exclued the number, which would mess things up when looking for the correct sheet
  var litterRegEx = /^[A-Za-z]+/;
  litter = litterRegEx.exec(litter)[0].toUpperCase();
  
  
  

  //create an array of the sex and color of the pups, the same length as the size of the litter
  //the first value in the Sex/Color column will be assigned to the sex variable, and the second to the color variable
  //if the sex variable is actually a color it is assigned to the color, and if the color variable is actually sex (i.e. two sexes were entered)
  //the sex is sex to ??
  var pupSex = [];
  var pupColor = [];
  var sex;
  var color;
  
  for(i=0; i<pups; i++) {
    if(responseValues[0][i+8] == undefined){
      sex = "";
      color = "";
      pupSex.push(sex);
      pupColor.push(color);
    }else{
      var sexAndColor = responseValues[0][i+8].split(", ");
      sex = sexAndColor[0];
      color = sexAndColor[1];
      if(sex != "F" && sex != "M"){
        color = sex;
        sex = "";
      }
      if(color == undefined){
        color = "";
      }
      if(color == "M" || color == "F"){
        sex = "??";
        color = "";
      }
      pupSex.push(sex);
      pupColor.push(color);
    }
  }
 
  //variables to write to the litter sheet are now set
  

  //check to see if a sheet with the same name as the line exists in the litter spreadsheet
  //if it doesn't, create new sheet and update list_of_all
  if(litterSpread.getSheetByName(litter) == null){
    litterSpread.insertSheet(litter, 0, {template: litterSpread.getSheetByName(LITTER_TEMPLATE_SHEET_NAME)});
    SpreadsheetApp.setActiveSpreadsheet(litterSpread);
    list_of_all.setActiveRange(list_of_all.getRange(list_of_all.getLastRow()+1, 1, 1, 2)).setValues([[litter, 1]]);
  }
  
  //check if genotyping sheet exists, create new one if it doesn't
  if(genotypingSpread.getSheetByName(plate) == null){
    genotypingSpread.insertSheet(plate, genotypingSpread.getNumSheets(), {template: genotypingSpread.getSheetByName(GENOTYPING_TEMPLATE_SHEET_NAME)});
  }
  
  //now that those sheets for sure exist, set litterSheet and plateSheet variables
  var litterSheet = litterSpread.getSheetByName(litter);
  var plateSheet = genotypingSpread.getSheetByName(plate);
  
  //find the value of the last litter on the litter sheet
  var lastRowIndex = litterSheet.getRange("A1:A").getValues().filter(String).length;
  var lastLitterRow = litterSheet.getRange(lastRowIndex, 1, 1, 1)
  var lastLitter = parseInt(lastLitterRow.getValues()[0][0]);
  var litternum = lastLitter + 1;
  
  //sets litternum = 1 if this is the first litter
  if(lastLitterRow.getA1Notation() == "A1"){
    litternum = 1;
  }
  
  //create list of the current litter's pup name
  var pupname;
  var pupnames = [];
  for(i=0; i<pups; i++){
    pupname = litter + " " + (litternum) + "." + (i+1);
    pupnames.push([pupname]);
  }
  
  
  
  //Jump over to the genotyping sheet to write the pupnames and get the plate location for writing to the litterSheet
  var lastPup = plateSheet.getRange("K1:K").getValues().filter(String).length;
  
  //initialize these variables. If there are too many pups they will come into play
  var overflow = 0;
  var tooManyPups = false;
  
  //check to see if there are too many pups. Create new sheet, change variables if there are
  if(lastPup + pups > 97){
    tooManyPups = true;
    overflow = lastPup + pups - 97;
    var newPlate = GENOTYPING_PLATE_PREFIX + (parseInt(plateNum) + 1);
    genotypingSpread.insertSheet(newPlate, genotypingSpread.getNumSheets(), {template: genotypingSpread.getSheetByName(GENOTYPING_TEMPLATE_SHEET_NAME)});
    var newPlateSheet = genotypingSpread.getSheetByName(newPlate);
  }
  
  var plateWriteRange = plateSheet.getRange((lastPup + 1), 11, pups - overflow, 1);
  
  var toWrite = [];
  for(i=0; i<(pups - overflow); i++){
    toWrite.push([pupnames[i]]);
  }
  
  //write values to the genotyping plate sheet
  plateSheet.setActiveRange(plateWriteRange).setValues(toWrite);
  
  //get the plate location to write to the litter sheet below
  var plateLocation = plateWriteRange.offset(0, -3).getValues();
  var location = [];
  for(i=0; i<plateLocation.length; i++){
    location.push(plate + " " + plateLocation[i][0]);
  }
  
  if(tooManyPups){
    //write the rest of the pups to the new plate sheet and add the location to the location array
    var newPlateWrite = newPlateSheet.getRange(2, 11, overflow, 1);
    var newPlateLocation = newPlateWrite.offset(0, -3).getValues();
    
    toWrite = [];
    for(i=pups-overflow; i<pups; i++){
      toWrite.push(pupnames[i]);
    }
    newPlateSheet.setActiveRange(newPlateWrite).setValues(toWrite);

    for(i=0; i<overflow; i++){
      location.push(newPlate + " " + newPlateLocation[i][0]);
    }
  }
  
  
  //create array to write into the litter spreadsheet
  //this is very positional based. If a column gets inserted into the littersheet this will get thrown off.
  //the width of this can be changed by adding elements into the array to space things out correctly
  var valuearray = [];
  for(i=0; i<pups; i++){
    if(i == 0) {
      valuearray.push([litternum, pupnames[i], location[i], litter + " " + litternum, father, mother1, mother2, dateOfBirth, "Y", i+1, i+1, pupColor[i], pupSex[i]]);
    } else {
      valuearray.push([litternum, pupnames[i], location[i], "", "", "", "", "", "Y", i+1, i+1, pupColor[i], pupSex[i]]);
    }
    
  }
   
  //write the array to the litter sheet. It will write one row below the last row
  litterSheet.setActiveRange(lastLitterRow.offset(1, 0, pups, valuearray[0].length)).setValues(valuearray);
  
  //update the List_of_litters sheet
  var litterOptions = list_of_all.getDataRange().getValues();
  var litterPosition;
  for(i=1; i<litterOptions.length; i++) {
    if(litterOptions[i][0] == litter){
      litterPosition = i+1;
      break;
    }
  }  
  list_of_all.setActiveRange(list_of_all.getRange(litterPosition, 1, 1, 2)).setValues([[litter, litternum]]);
  

  //run this function to update the form with info from list_of_all sheet
  litterNames();

}







