var ssID = "";
var formID = "";

var rawData = SpreadsheetApp.openById(ssID).getSheetByName("Form Responses 1");
var wsData = SpreadsheetApp.openById(ssID).getSheetByName("Limit");
var organizedData = SpreadsheetApp.openById(ssID).getSheetByName("Organized");
var form = FormApp.openById(formID);

var DEFAULT_STUDENT_LIMIT = "25";
var MENU_NAME = "Programming Exam Options";
var ORGANIZED_ROW_DEFAULT_POSITION = 3;

// initialize custom menu for spreadsheet when the excel sheet is opened
function onOpen(){
  initMenu();

}

function onEdit(){
  updateQuestions();
}

// Initializes the custom menu options
function initMenu(){
  var ui = SpreadsheetApp.getUi();
  var menu = ui.createMenu(MENU_NAME);
  menu.addItem("Update Form", "updateQuestions");
  menu.addItem("Add a Time Slot", "addQuestion");
  menu.addSeparator();
  menu.addItem("Clear All Data", "clearAllData");
  menu.addToUi();

}

//deletes all data from Form Response 1 page and the form responses
function clearAllData(){
  var response = SpreadsheetApp.getUi().prompt("Verify you want to delete all Data by typing \"confirm\"").getResponseText();
  if(response == "confirm"){
    form.deleteAllResponses();
    var numberOfRows = rawData.getLastRow();
    if(numberOfRows >1){
      rawData.deleteRows(2, rawData.getLastRow()-1)
    }
    SpreadsheetApp.getActiveSpreadsheet().toast("All Data has been cleared")
  }
  updateQuestions();
}

// Adds a new timeslot to the google form
function addQuestion(){
  var questionToAdd = addQuestionPrompt();
  addQuestionToLimitsTab(questionToAdd);
  addNewOrganizedColumnForQuestion(questionToAdd);
  updateQuestions();
}

//Prompts the user for a new timeslot
function addQuestionPrompt(){
  return SpreadsheetApp.getUi().prompt("Enter the Time Slot you'd like to add").getResponseText();
}

//Adds the new Timeslot from user to the Limits page
function addQuestionToLimitsTab(question){
  // gets all of the questions from the Limits Pages
  var questions = wsData
                   .getRange(2,1,wsData.getLastRow()-1,1)
                   .getValues()
                   .map(function(q){return q[0]})
                   .filter(function(q){return q !== ""});

  //Gets the row below the last question on the limits page
  var newQuestionsRowNumber = questions.length + 2;

  //sets the new values with the new Row number
  var columnA = "=if(C" + newQuestionsRowNumber + "<D" + newQuestionsRowNumber + ",B" + newQuestionsRowNumber+ ",\"This session is full\")" ;
  var columnB = question;
  var columnC = "=countif(\'Form Responses 1\'!E:E,B" + newQuestionsRowNumber + ")";
  var columnD = DEFAULT_STUDENT_LIMIT;

  //Set the new values in the excel sheet
  wsData.getRange(newQuestionsRowNumber, 1).setValue(columnA)
  wsData.getRange(newQuestionsRowNumber, 2).setValue(columnB)
  wsData.getRange(newQuestionsRowNumber, 3).setValue(columnC)
  wsData.getRange(newQuestionsRowNumber, 4).setValue(columnD)
}


//adds a new column that querys for the new timeslot from the forms response 1 page
function addNewOrganizedColumnForQuestion(question){
  Logger.log("adding a new organized column")
  var lastColumn = organizedData.getLastColumn();
  var newColumnLocation = lastColumn +3;
  Logger.log("new column location = " + newColumnLocation);

  var titleRow = question;
  var queryRow = "=query(\'Form Responses 1\'!D2:E,\"Select * Where E=\'" + question + "\'\")";

  Logger.log("setting query to " + queryRow);
  organizedData.getRange(ORGANIZED_ROW_DEFAULT_POSITION, newColumnLocation).setValue(titleRow)
  organizedData.getRange(ORGANIZED_ROW_DEFAULT_POSITION + 1, newColumnLocation).setValue(queryRow)

  Logger.log(lastColumn);
}


//Updates the google form with the new Timeslot
function updateQuestions(){
  var labels = wsData.getRange(1,1,1,wsData.getLastColumn()).getValues();
  var questions = wsData
                   .getRange(2,1,wsData.getLastRow()-1,1)
                   .getValues()
                   .map(function(q){return q[0]})
                   .filter(function(q){return q !== ""});
  Logger.log(questions);

  var filtered = questions.filter(function(value, index, arr){ return value !== "This session is full";});
  Logger.log(filtered);

  if(filtered.length !== 0){
    updateDropDownQuestionUsingTitle("Select a Programming Session",filtered)
  }

}


function updateDropDownQuestionUsingTitle(title, values) {
  var itemID = findIDValue(title)
  updateDropDownQuestions(itemID,values);
}


function updateDropDownQuestions(id,values) {
  var item = form.getItemById(id);
  item.asListItem().setChoiceValues(values);
}


function findIDValue(itemTitle) {
  var allItems = form.getItems();

  var i ;
  for(i = 0; i < allItems.length;i++){
    if(allItems[i].getTitle() == itemTitle){
      return allItems[i].getId();
    }
  }

  return 0;
}
