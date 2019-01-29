LITTER_SPREADSHEET_ID = "id_goes_here";
LITTER_LIST_SHEET_NAME = "List_of_all";
LITTER_TEMPLATE_SHEET_NAME = "Template";
GENOTYPING_SPREADSHEET_ID = "another_id_here";
GENOTYPING_TEMPLATE_SHEET_NAME = "Template";
FORM_RESPONSE_ID = "another_id";
FORM_RESPONSE_SHEET_NAME = "Form Responses 1"
GENOTYPING_PLATE_PREFIX = "WR "; //add a space at the end if you want "WR 12" instead of "WR12"
IGNORESHEETS = {"Template":"", "List_of_all":"", "Ignore this sheet":"", "Ignore this sheet as well":""};
LITTER_NAME_ON_FORM = "Litter Name"; //title of the multiple choice item for litters on the Form 




function sheetName() {
  return SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getName();
}


function onOpen() {
SpreadsheetApp.getUi().createMenu('Update Current Litters').addItem('Update Form', "updateForm").addItem("Update List_of_all", "updateLitters").addToUi();
}
                                  
                                  

function updateForm(){
  //Checks the litter spreadsheet and populates the form with mouse litter names and current number as found in the litter spreadsheet "List_of_all"
  //almost the same as 'litterNames()' in the Plate Form script
  //litterNames() runs every time the form is submitted or opened. 
  //This form can be triggered with a menu button to push any updates from List_of_all to the form
  
  var plateForm = FormApp.openById('160yxoV8RLm4wCCs2Ryb9cby6wg-JZgK_7ErawhdAbPs');
  var items = plateForm.getItems();
  var litterSpread = SpreadsheetApp.openById(LITTER_SPREADSHEET_ID);
  var choiceArray = litterSpread.getSheetByName(LITTER_LIST_SHEET_NAME).getDataRange().getValues();  
  var newArray = [];

  for(i=1; i<choiceArray.length; i++) {
    newArray.push(choiceArray[i][0] + " " + (choiceArray[i][1] + 1));
  }
  Logger.log(newArray);

  
  for (var i = 0; i < items.length; i += 1){
    var item = items[i]
      if (item.getTitle() == LITTER_NAME_ON_FORM){
        item.asMultipleChoiceItem().setChoiceValues(newArray);
        break;
    }
  }
}



function updateLitters() {
  //Function to update List_of_all
  //This function is triggered once a day, or with a menu button press
  //This ensures that if new lines or litters are added to spreadsheet manually, the List_of_all will reflect that
  //and the form will be valid
  var litterSheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  var listofall = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(LITTER_LIST_SHEET_NAME);
  
  var titleArray = [];
  
  for(i=0; i<litterSheets.length; i++) {
    var currSheet = litterSheets[i];
    
    if(currSheet.getSheetName() in IGNORESHEETS) {
      //do nothing
    }else {
      var title = currSheet.getSheetName();
      var currLitter = currSheet.getRange(currSheet.getRange("A1:A").getValues().filter(String).length, 1, 1, 1).getValues()[0][0];
      Logger.log(typeof(currLitter));
      titleArray.push([title, currLitter]);
    }
  }
  listofall.setActiveRange(listofall.getRange(2, 1, titleArray.length, 2)).setValues(titleArray);
}
  
