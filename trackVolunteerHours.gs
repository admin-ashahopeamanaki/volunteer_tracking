/**
 * @author Aryanw
 */

/**
 * create menu Items
 * Input: an occured event
 */
function onOpen(e) {
  // create a new menu item
  var menu = SpreadsheetApp.getUi().createMenu("⚙️ Automation");
  var ui = SpreadsheetApp.getUi();
  menu.addItem("Update Sheet","readGivePulseEmails");
  menu.addItem("Calculate Hours","calcHours");
  menu.addItem("Initialize Sheets ","initializeSheets");
  menu.addToUi();
  
  console.log("starting execution");

}

/**
 * Initialize sheets if not already created
 */
function initializeSheets(){

  var spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
  
  //----------init VolunteerData sheet---------------------
  var dataSheet = spreadSheet.getSheetByName('VolunteerData');
  if (!dataSheet){
    console.log("initialize dataSheet")
    SpreadsheetApp.getActiveSpreadsheet().insertSheet('VolunteerData');
    dataSheet = spreadSheet.getSheetByName('VolunteerData');
    dataSheet.getRange(1,1).setBackground("yellow");
    dataSheet.getRange(1,2).setBackground("yellow");
    dataSheet.getRange(1,3).setBackground("yellow");
    dataSheet.getRange(1,4).setBackground("yellow");
    // dataSheet.getRange(5,1).setValue("Calculated Hours").setBackground("orange");
    dataSheet.appendRow(["Full Name","Event Name","Date Time","Hours"]);
  }
  else {
    console.log("VolunteerData sheet already exists")
  }

  //----------init MainSheet (CalcHours) sheet---------------
  var mainSheet = spreadSheet.getSheetByName('CalcHours');
  if (!mainSheet) {
    console.log("initialize mainSheet")
    SpreadsheetApp.getActiveSpreadsheet().insertSheet('CalcHours');
    mainSheet = spreadSheet.getSheetByName('CalcHours');
    mainSheet.getRange(1,1).setValue("Select Name").setBackground("yellow");
    mainSheet.getRange(2,1).setValue("Select Event").setBackground("yellow");
    mainSheet.getRange(3,1).setValue("From").setBackground("yellow");
    mainSheet.getRange(4,1).setValue("To").setBackground("yellow");
    mainSheet.getRange(5,1).setValue("Calculated Hours").setBackground("orange");

    // ----set date validations----
    console.log("init validations");
    var dateValidation = SpreadsheetApp.newDataValidation().requireDate().build();
    var fromCell = mainSheet.getRange(3,2);
    var toCell = mainSheet.getRange(4,2);
    fromCell.setDataValidation(dateValidation);
    toCell.setDataValidation(dateValidation);
    console.log("date validation setup complete");
  }
  else {
    console.log("CalcHours sheet already exists")
  }

/**
 * dynamically generate a drop down list from a static list
 * Input: an occured event
 */
}
function onEdit(e) {
  
  var range = e.range;
  var spreadSheet = e.source;
  var sheetName = spreadSheet.getActiveSheet().getName();
  var column = range.getColumn();
  var row = range.getRow();
  var selectedName = e.value;
  var returnValues = [];

  if(sheetName == 'CalcHours' && column == 2 && row == 1) {

    var ss= SpreadsheetApp.getActiveSpreadsheet();
    var mainSheet = ss.getSheetByName("CalcHours");
    var dataSheet = ss.getSheetByName("VolunteerData");


    //GET SQL STATEMENT
    var lastRowData = dataSheet.getLastRow();    
    for(var i = 1; i <= lastRowData; i++) {

      if(selectedName == dataSheet.getRange(i, 1).getValue()) {
        returnValues.push(dataSheet.getRange(i, 2).getValue());      
      }
    }
    mainSheet.getRange('B2').clear();
    var dropdown = mainSheet.getRange('B2');
    var rule = SpreadsheetApp.newDataValidation().requireValueInList(returnValues).build();
    dropdown.setDataValidation(rule);
  }

}


/**
 * read all email threads with "GivePulse" tag.
 */
function readGivePulseEmails(){
  console.log("starting parse")
  var label = GmailApp.getUserLabelByName("GivePulse");
  var threads = label.getThreads();

  // starting from last to first index
  // scan through every email in each thread tagged with label: "GivePulse"
  for ( var i = threads.length - 1; i >= 0; i--) {
    // all messages in a thread identified by label "automation"
    
    var messages = threads[i].getMessages();
    
    // 
    for(var j = 0;j<messages.length;j++){
      var message = messages[j];
      extractDetails(message);
    }
  }
  // primary dropdown validation
  // -- primary dropdown is validated here to adjust the number of rows --
  var spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
  var dataSheet = spreadSheet.getSheetByName('VolunteerData');
  var lastRow = dataSheet.getLastRow();
  var partRange = dataSheet.getRange("A2:A" + lastRow);
  var rule = SpreadsheetApp.newDataValidation().requireValueInRange(partRange).build();
  var cell = spreadSheet.getSheetByName("calcHours").getRange('B1');
  cell.setDataValidation(rule);
  console.log("Primary dropdown validation setup complete");
}

/**
 * parse relevant information from valid emails
 * Input: an email Object
 */

function extractDetails(message) {
  var dateTime = message.getDate();
  var subjectText = message.getSubject();
  var senderDetails = message.getFrom();

  // confirm sender email or break
  var email = new RegExp("<(.*)>");
  var temp = email.exec(senderDetails)[0];
  if (temp != "<notification@giveplus.com>") {  // use <notification@givepulse.com> in live
    console.log("email doesn't match");
    return
  }
 
  // use regex to extract name,event,hours
  var t1 = new RegExp("not(.*)gave");
  var temp1 = t1.exec(subjectText)[0].split(' ');
  var fullName = temp1[1];
  for(var index1 = 2;index1<temp1.length-1;index1++){
    fullName = fullName + " " + temp1[index1]
  }
  
  var t2 = new RegExp("gave(.*)");
  var temp3 = t2.exec(subjectText)[0].split(' ');
  var hours = temp3[1];
  
  var t3 = new RegExp("@(.*).");
  var temp3 = t3.exec(subjectText)[0].split(' ');
  var eventName = temp3[1];
  for(var index3 = 2;index3<temp3.length;index3++){
    if (index3 == temp3.length-1 && temp3[index3].charAt(temp3[index3].length - 1) == ".") {
      eventName = eventName + " " + temp3[index3].slice(0,-1);
    }
    else {
      eventName = eventName + " " + temp3[index3];
    }
  }
  
  // append in VolunteerData
  appendDetails(dateTime,subjectText,senderDetails,fullName,hours,eventName)
} 

/**
 * append relevant info in VolunteerData sheet
 * Input: info extracted from email object to be appended into VolunteerData sheet
 */
function appendDetails(dateTime,subjectText,senderDetails,fullName,hours,eventName) {

  
  var spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
  var activeSheet = spreadSheet.getSheetByName('VolunteerData');
  var dataRange = activeSheet.getDataRange();
  var data = dataRange.getValues();
  var entry_exists = false;

  // omit the entry IF name,event,dateTime matched in VolunteerData ELSE add row
  for (var i = 1; i<data.length;i++){
    if (data[i][0] == fullName && data[i][1] == eventName && data[i][2].toISOString() == dateTime.toISOString() ) {
      console.log("name, event, dateTime match")
      entry_exists = true
    }
  }
  if (entry_exists == false) {
    console.log("No match")
    activeSheet.appendRow([fullName,eventName,dateTime,hours]) // subjectText,senderDetails,dateTime
  }
}

/**
 * Calculate total hours from parameters given in CalcHours sheet and update hours cell.
 */
function calcHours(){
  // setup cells to extract from CalcHours sheet
  var mainSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('CalcHours');
  var n = mainSheet.getRange(1,2).getValue();
  var e = mainSheet.getRange(2,2).getValue();
  var fromDate = mainSheet.getRange(3,2).getValue();
  var toDate = mainSheet.getRange(4,2).getValue();

  // setup rows to loop through in VolunteerData sheet
  var dataSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("VolunteerData");
  var lastRow = dataSheet.getLastRow()
  var name = dataSheet.getRange(1,1,lastRow).getValues();
  var event = dataSheet.getRange(1,2,lastRow).getValues();
  var dateTime = dataSheet.getRange(1,3,lastRow).getValues();
  var hours = dataSheet.getRange(1,4,lastRow).getValues();
  
  // calculate hours 
  total_hours = 0
  const dates = new Set()
  for (var i=0;i<name.length;i++) {
    console.log(name[i][0],event[i][0], total_hours)
    console.log(name[i][0],event[i][0], total_hours)
    
    if (name[i][0] == n && event[i][0] == e) {
        if (fromDate<=dateTime[i][0] && dateTime[i][0]<=toDate && !dates.has(dateTime[i][0].toISOString())) {
            total_hours+= Number(hours[i])
            dates.add(dateTime[i][0].toISOString())
        }
        console.log(fromDate)
        console.log(dateTime[i][0])
        console.log(toDate)
    } 
  }
  console.log(total_hours)
  mainSheet.getRange(5,2).setValue(total_hours);

  /** For time and date acurate testing
    * const MILLIS_PER_WEEK = 1000 * 60 * 60 * 24 * 7; 
    * const fromDate = new Date(toDate.getTime() - MILLIS_PER_WEEK);
    */
}

 




