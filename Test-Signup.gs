/*
When a user presses Submit on a google form it creates an onSubmit event object.
This event object contains the information submitted by the user on the form
The onSubmit method is asynchronous to the updating of the excel sheet containing the responses
*/


// The ID for the Google Sheets, found in the URL
var ssID = "";

// The ID of the Google Form, found in the URL
var formID = "";

// a reference to the spreadsheet tab named Limit and Tab named Form Responses 1
var wsData = SpreadsheetApp.openById(ssID).getSheetByName("Limit");
var formResponses1Data = SpreadsheetApp.openById(ssID).getSheetByName("Form Responses 1");

var form = FormApp.openById(formID);


// is a reserved function for the onSubmit trigger when a form is submitted
function onSubmit(e) {

  // gets the email and submission String from the submission event object
  var email = geValueFromResponse(e, "Email Address");
  var submissionTime = geValueFromResponse(e, "Select a Programming Session");


  //Updates the form to reflect the new time options available to students
  main();


  // get all rows with the same email and a boolean saying if the limit was exceeded
  var numberOfResponses = getAllRowsWithThisEmail(email);
  var isLimitExceeded = wasLimitExceeded(submissionTime);
//  var isEmailValid = validateEmail(email);
  Logger.log(isLimitExceeded);


  // Sends an email to the given email, confirming the successfulness of the programming exam
  if(verifySuccess(numberOfResponses, isLimitExceeded)){
    MailApp.sendEmail( email,
                    "Programming Exam Successfully Scheduled",
                    "You have been scheduled to take your exam at "+ submissionTime + "\n\n" +
                    "If you need to reschedule your exam, send an email to CS240TA@cs.byu.edu\n",
                    {name:"CS240 TA"});
  }
  else{
    var message;
    removeThisEntryFromSheets(numberOfResponses[numberOfResponses.length-1]);
    if(isLimitExceeded){
      message = "Too many people have signed up for this time, Please signup for a different time\n" +
                      "Please resubmit a programming exam request here: ";
    }
    else if(numberOfResponses.length >1){
      message = "You may only Sign-up for one time. To change your time, you must contact the CS-240 TA\n"
      + "CS240TA@cs.byu.edu";
    }
    else{
      message = "For some unknown reason, your request failed.\n" +
                      "Please resubmit a programming exam request here: ";
    }


    MailApp.sendEmail( email,
                      "FAILED TO SUBMIT: Programming Exam Schedule ",
                      message,
                      {name:"CS240 TA"});




  }
}


function main() {

  // gets all of the questions from the Limits Pages
  var questions = wsData
                   .getRange(2,1,wsData.getLastRow()-1,1)
                   .getValues()
                   .map(function(q){return q[0]})
                   .filter(function(q){return q !== ""});


  // The spreadsheet automatically changes the question to "This session is full" when the limit is exeeded.
  // This function removes all questions that are called "This session is full"
  var filtered = questions.filter(function(value, index, arr){ return value !== "This session is full";});

  // Sets the Dropdown question options of the given Question title to be filtered questions.
  if(filtered.length !== 0){
    updateDropDownQuestionUsingTitle("Select a Programming Session",filtered)
  }
}

// this function takes a title for a dropdown google form question and a list of strings and sets the dropdown questions to  the list found in the values variable
function updateDropDownQuestionUsingTitle(title, values) {
  var itemID = findIDValue(title)
  updateDropDownQuestions(itemID,values);
}

//This function takes the ID of the question being updated and sets the questions to be the values found in the STring list : values
function updateDropDownQuestions(id,values) {
  var item = form.getItemById(id);
  item.asListItem().setChoiceValues(values);
}


function verifySuccess(numberOfResponses, isLimitExceeded){

  if((numberOfResponses.length == 1) && (isLimitExceeded == false)){
    return true
  }
  else{
    return false;
  }
}

function validateEmail(mail) {
 if (/^\w+([\.-]?\w+)*@\w+([\.-]?\w+)*(\.\w{2,4})+$/.test(mail)){
   Logger.log("true");
    return (true)
  }
  Logger.log("false");
    return (false)
}

function getAllRowsWithThisEmail(email){

  // Gets all emails Submitted
  var emailsSubmitted = formResponses1Data
                   .getRange(2,6,formResponses1Data.getLastRow()-1,1)
                   .getValues()
                   .map(function(q){return q[0]})
                   .filter(function(q){return q !== ""});

  // gets All Row numbers with the same Email
  var allRowsWithEmail = [];
  var currentIndex;
  for(currentIndex = 0; currentIndex < emailsSubmitted.length; currentIndex++){
    if(emailsSubmitted[currentIndex] == email){
      allRowsWithEmail.push(currentIndex+2);
    }
  }

  return allRowsWithEmail;
}
function wasLimitExceeded(submissionTime){

  // gets All of the available submission options from the Limits Page
  var options = wsData
                   .getRange(2,2,wsData.getLastRow()-1,1)
                   .getValues()
                   .map(function(q){return q[0]})
                   .filter(function(q){return q !== ""});

  //Gets looks through all of the submission options and gets the row number of the submission
  var indexNumber;
  for(indexNumber = 0; indexNumber < options.length; indexNumber++){
    if(submissionTime == options[indexNumber]){
      break;
    }
  }

  //gets the limit number for the submission
  var optionLimit = wsData
                   .getRange(indexNumber+2,4,1,1)
                   .getValues()
                   .map(function(q){return q[0]})
                   .filter(function(q){return q !== ""});

  //gets the count recorded in the options Count of the Limits page
  var optionCount = wsData
                   .getRange(indexNumber+2,3,1,1)
                   .getValues()
                   .map(function(q){return q[0]})
                   .filter(function(q){return q !== ""});

  Logger.log("optionCount = " + optionCount);
  Logger.log("optionLimit = " + optionLimit);
  // if the option count is higher than the option limit, there is an excessive number of responses. and the response is not valid
  if(optionCount > optionLimit){
    return true;
  }
  return false;
}

function removeThisEntryFromSheets(rowNumber){
  if(rowNumber != 1){
    formResponses1Data.deleteRow(rowNumber);
  }

}


//This function returns the value of the given form submission based on the title of the questiosn title
//e is the submission event, target is the string title of the question
function geValueFromResponse(e,target){
 var items = e.response.getItemResponses();
  for (i in items){
    if(items[i].getItem().getTitle() == target){
      return items[i].getResponse();
    }
  }
  return null;
}


// This Function takes all of the questions and items in the google form defined at the top the script and finds the ID value of the given itemTitle String
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
