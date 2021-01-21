//clean this up
var activeWorkbook = SpreadsheetApp.getActiveSpreadsheet(); //this is actually the workbook, not the sheet
var activeSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet(); //this is a tab in the workbook
var activeRange = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getActiveRange();
var ui = SpreadsheetApp.getUi();
var date = new Date().toDateString(); //yes, please, re-initialize me
//clean this up
var documentProperties = PropertiesService.getDocumentProperties();
var store = documentProperties.getProperties();

//*********
//Let the functions begin!!! 

function onOpen(e) {
//let people know there is some loading going on
activeWorkbook.getSheetByName('scan').getRange('C4').clear()
  .setHorizontalAlignment('center')
  .setVerticalAlignment('middle')
  .setFontFamily('Arial')
  .setBackgroundColor('#fff')
  .setBorder(true, true, true, true, false, false, 'red', SpreadsheetApp.BorderStyle.DASHED)
  .setFontColor('red')
  .setFontSize(18)
  .setValue('LOADING ... ');

//create the custom menu 
SpreadsheetApp.getUi()
.createMenu('RMA DASHBOARD')
.addItem('RMA Job Dashboard', 'loadSidebar')
.addItem('Check-in New Meter', 'meterCheckIn')
.addToUi();
 
//assign column numbers to each header and store in the documentProperties

  var headerTemplate = activeWorkbook.getSheetByName('rmaSheetTemplate');
  var lastColumn = headerTemplate.getLastColumn();
  var headers = headerTemplate.getRange(1, 1, 1, lastColumn).getValues();
  for (var row in headers) {
  for (var i = 0; i < lastColumn; i++) {
  documentProperties.setProperty(headers[row][i], i+1);
    }
  }
scanBarcode();       
 }

function doGet(){
var tmp = HtmlService.createTemplateFromFile('meterCheckInForm');
return tmp.evaluate().addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function include(filename){
return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function loadSidebar() {
var htmlOutput = HtmlService
    .createHtmlOutputFromFile('buttons')
    .setTitle('RMA Dashboard');
SpreadsheetApp.getUi().showSidebar(htmlOutput); 

}

function clearReadyCellAtClose(){
var scanSheet = activeWorkbook.getSheetByName('scan');
scanSheet.getRange('D3').clear()
  .setBackgroundColor('#fff')
  .setFontColor('#fff');
 scanSheet.getRange('C4').clear();
  return true;
}

function workingOnIt(status){
var workingTmp = HtmlService.createTemplateFromFile('loader_HTML');
workingTmp.workingContent = "";
var workingDiv = "<div id='loading' class='center-align loading'><a class='btn-floating btn-large red darken-2 pulse'></a><div class='working'>working ...</div></div>";
var readyDiv = "<div id='ready' class='center-align ready'><a class='btn-floating btn-large green'></a><div class='working'>ready</div></div>";
var loggedItemDiv = "<div id='ready' class='center-align ready'><a class='btn-floating btn-large blue'></a><div class='working'>item logged</div></div>";
var noRMADiv = "<div id='noRMA' class='center-align ready'><a class='btn-floating btn-large grey pulse'></a><div class='working'>RMA ID not found</div></div>";
var scriptTag = "<script>window.close = function(){window.setTimeout(function(){google.script.host.close()},2500)};close();</script>";
var dialogTitle = " ";

if(!status){
workingTmp.workingContent = workingDiv;
workingTmp.scriptTag = " ";
}

if(status == 'AOK'){
workingTmp.workingContent = readyDiv;
workingTmp.scriptTag = scriptTag;
}

if(status=="noRMALocation"){
workingTmp.workingContent = noRMADiv;
workingTmp.scriptTag = scriptTag;
};

if(status=="Item Logged"){
workingTmp.workingContent = loggedItemDiv;
workingTmp.scriptTag = scriptTag;
};

workingTmp = workingTmp.evaluate();
workingTmp.setWidth(300).setHeight(150);
SpreadsheetApp.getUi().showModalDialog(workingTmp, dialogTitle);
}
 
//check-in new meter
//future: move this to the sidebar
function meterCheckIn() {
  let meterCheckInFormLink = 'https://script.google.com/macros/s/AKfycbzAvFhcBGDts17a_dtj39fZsn7VyOyEZ6eBo_voOB-0_-9gNSYO8_EZbQ/exec';
   var html = HtmlService.createHtmlOutput('<html><script>'
  +'window.close = function(){window.setTimeout(function(){google.script.host.close()},9)};'
  +'var a = document.createElement("a"); a.href="'+meterCheckInFormLink+'"; a.target="_blank";'
  +'if(document.createEvent){'
  +'  var event=document.createEvent("MouseEvents");'
  +'  if(navigator.userAgent.toLowerCase().indexOf("firefox")>-1){window.document.body.append(a)}'                          
  +'  event.initEvent("click",true,true); a.dispatchEvent(event);'
  +'}else{ a.click() }'
  +'close();'
  +'</script>'
  // Offer URL as clickable link in case above code fails
  +'<body style="word-break:break-word;font-family:sans-serif;">Failed to open automatically. <a href="'+meterCheckInFormLink+'" target="_blank" onclick="window.close()">Click here to proceed</a>.</body>'
  +'<script>google.script.host.setHeight(40);google.script.host.setWidth(410)</script>'
  +'</html>')
  .setWidth( 90 ).setHeight( 1 );
  SpreadsheetApp.getUi().showModalDialog(html, "Opening ..." );
}


function scanBarcode() {

//click heels together three times
var scanSheet = activeWorkbook.getSheetByName('scan').activate();
  scanSheet.getRange('C3').clear()
  .setHorizontalAlignment('center')
  .setVerticalAlignment('middle')
  .setFontFamily('Arial')
  .setBackgroundColor('#e8f5e9')
  .setFontSize(14);
var scanCellReady = scanSheet.getRange('C3').activate();

scanSheet.getRange('D3').clear()
  .setHorizontalAlignment('center')
  .setVerticalAlignment('middle')
  .setFontFamily('Arial')
  .setBackgroundColor('#388e3c')
  .setFontColor('#fff')
  .setFontSize(14)
  .setValue('READY');
  
  scanSheet.getRange('C4').clear()
  .setHorizontalAlignment('center')
  .setVerticalAlignment('middle')
  .setFontFamily('Arial')
  .setBackgroundColor('#fff')
  .setBorder(null, true, true, true, false, false, '#388e3c', SpreadsheetApp.BorderStyle.DASHED)
  .setFontColor('#388e3c')
  .setFontSize(14)
  .setValue('Open RMA Dashboard and scan barcode to begin.');

return 'AOK';
}

//this is an onEdit trigger that writes the rmaID and companyCode to the property store
function onScan(e)  {

if(activeSheet.getSheetName() == 'scan'){
      workingOnIt();
   //grab the RMA ID
      var rmaID = activeWorkbook.getActiveSheet().getActiveRange().getValue();
      documentProperties.setProperty('rmaID', rmaID);
      documentProperties.setProperty('companyCode', rmaID.slice(0, 3));
//find the RMA ID in the workbook
findRMARow(rmaID);
}   
}

//Chapter One: In which we prepare for the journey and find the all-important rmaRow
function findRMARow(rmaID) {
  var companyCode = documentProperties.getProperty('companyCode');
  var companyTab = activeWorkbook.getSheetByName(companyCode);
  if(!companyTab){workingOnIt('noRMALocation'); clearRedirect(); return;};
  //companyTab.activate();
  var rmaLocation = companyTab.createTextFinder(rmaID).findNext();
  if(!rmaLocation){workingOnIt('noRMALocation'); clearRedirect(); return;};

  var rmaRow = rmaLocation.getRow();
  
//store the row number
  documentProperties.setProperty('rmaRow', rmaRow);
  
//******WARNING: HARD-CODED VALUE BELOW ********  
  var svcOrderTotalCols = 5;
  documentProperties.setProperty('svcOrderTotalCols', svcOrderTotalCols);
  
//store the customer-ordered activities aka the service order
  var svcOrderFirstCol = documentProperties.getProperty('activity[0]');
  svcOrderFirstCol = Number(svcOrderFirstCol);

  var svcOrderValues = companyTab.getRange(rmaRow, svcOrderFirstCol, 1, svcOrderTotalCols).getValues();
  for (var col in svcOrderValues[0]) {
  documentProperties.setProperty('svcOrder['+ col +']',svcOrderValues[0][col]);
    }
  
//get the meter type and serial number from customer's side of the spreadsheet  
  var meterModelCol = Number(documentProperties.getProperty('meterModel'));
  var serialNumberCol = Number(documentProperties.getProperty('serialNumber'));
  var meterModel = companyTab.getRange(rmaRow, meterModelCol).getValue();
  var serialNumber = companyTab.getRange(rmaRow, serialNumberCol).getValue().toString();
  
//create the PROVEit meterID
  var meterID = "";
  var space = / /g;
  var prefix = /sn/i;
  if(space.test(serialNumber)) {
    meterID = serialNumber.replace(space, "_");
  } else{
    meterID = serialNumber;
  }
    if(meterID.length >= 16) {
    meterID = meterID.slice(-16);
  }
  if(!prefix.test(meterID.slice(0, 2))){
     if(meterID.length >= 14){  
    meterID = 'SN'+ meterID.slice(-meterID.length + 2);
  } else {
    meterID = 'SN' + meterID;
       }
}
//create the PROVEit meterName
  var meterName = companyCode +" "+ serialNumber +" "+ meterModel;
  documentProperties.setProperties({"meterID": meterID, "meterName": meterName});
  
//select the rmaID in the companyTab 
  companyTab.setActiveRange(rmaLocation);
  
//is the RMA closed or completed? 
 var jobCompleteCol = Number(documentProperties.getProperty('JOB COMPLETE'));
 var jobClosedCol =  Number(documentProperties.getProperty('JOB CLOSED'));
 var reasonCol =  Number(documentProperties.getProperty('REASON'));

  var isComplete = companyTab.getRange(rmaRow,jobCompleteCol).getValue();
  var isClosed = companyTab.getRange(rmaRow,jobClosedCol).getValue();
  var isClosedReason = companyTab.getRange(rmaRow,reasonCol).getValue();
  if(isComplete){
    ui.alert("Job Complete", "This job was logged as COMPLETED on " + isComplete, ui.ButtonSet.OK);
    clearRedirect();
  } else if(isClosed){
    var openRMA = ui.alert("Job Closed: " +isClosedReason,
                           "This job was logged as CLOSED on " + isClosed + " for " + isClosedReason + "."
                           + "\n\nDo you need to re-open this RMA? ", ui.ButtonSet.YES_NO_CANCEL);
                           
    if(openRMA == ui.Button.YES) {
      reopenRMA(rmaID, companyTab);
    }
    clearRedirect();
  } else {
  
  workingOnIt('AOK');
  
  }
}

function reopenRMA(rmaID, companyTab) {
      var rmaRow = Number(store['rmaRow']);
      var companyName = companyTab.getRange(rmaRow, Number(store['companyName'])).getValue();
      var contactName = companyTab.getRange(rmaRow, Number(store['contactFirstName'])).getValue();
      var contactCell = companyTab.getRange(rmaRow, Number(store['contactCell'])).getValue();
      var contactEmail = companyTab.getRange(rmaRow, Number(store['contactEmail'])).getValue();
      var reopenRMA = ui.prompt("Re-open RMA " + rmaID, "To re-open this RMA, please input your PIN", ui.ButtonSet.OK_CANCEL);
      if(reopenRMA.getSelectedButton() == ui.Button.CANCEL || reopenRMA.getSelectedButton() == ui.Button.CLOSE) {
      ui.alert("Action Canceled", "RMA " +  rmaID + " not re-opened.", ui.ButtonSet.OK);
        clearRedirect();
      } else { 
        var pinNumber = reopenRMA.getResponseText();
        if(validatePINFormat(pinNumber)){
          var noteText = companyTab.getRange(rmaRow, Number(store['JOB CLOSED'])).getNote();
          noteText += '\n\n ' + date + ' : RMA Re-opened by ' + store['pinName'] + '\n\n';
          
          companyTab.getRange(rmaRow, Number(store['JOB CLOSED'])).setNote(noteText)
            .setValue('');
          ui.alert('RMA ' + rmaID + ' has been re-opened.', ui.ButtonSet.OK);
          //need an email address in here
          GmailApp.sendEmail('', 'RMA: ' + rmaID + ' Re-opened', 'RMA: ' + rmaID + '\nRe-opened by: ' + store['pinName'] + 
          '\nDate Re-opened: ' + date + 
          '\nCompany Name: ' + companyName + 
          '\nContact Name: ' + contactName + 
          '\nContact Phone: ' + contactCell +
          '\nContact Email: ' + contactEmail +
          '\n\nPlease confirm the details and follow up with the customer if necessary.');
          
          }
        clearRedirect();
      }
}

function openRMAForm() {
  var rmaID = store['rmaID'];
  var companyTab = activeWorkbook.getSheetByName(store['companyCode']);
  var companyName = companyTab.getRange(Number(store['rmaRow']), Number(store['Company Name'])).getValue();
  var timeCode = rmaID.slice(3);
  var fileName = companyName + ' RMA ' + timeCode + 'PDF';
  var rmaDocId = DriveApp.getFilesByName(fileName).next().getId();
  var rmaDocUrl = DriveApp.getFilesByName(fileName).next().getUrl();
  //var rmaPDF = DocumentApp.openByUrl(rmaDocUrl);
  var rmaDoc = DriveApp.getFileById(rmaDocId);
  
  //rmaDoc = DocumentApp.openById(rmaDoc);
  //all that code was total shit 
  
  var html = HtmlService.createTemplateFromFile('viewRMA');
  
  //snakes and sparklers, right here 
  html.rmaDoc = rmaDocId;
  html.fileName = fileName;
 
  var dialogTitle = "RMA " + rmaID; 
  var dialogBox = html.evaluate();
  ui.showModalDialog(dialogBox.setWidth(800).setHeight(800), dialogTitle);
}

//what if the tech selects the wrong activity button?
function checkServiceOrder(activityName) {
  var ttlCols = Number(store['svcOrderTotalCols']);
  for (var i = 0; i < ttlCols; i++) {
        if(store['svcOrder[' + i + ']'].toUpperCase() == activityName) {
          return true;
        }
  }
  return false;
}

//true or false: has everything been completed? don't bullshit me
function checkJobsLogged() {
 var ttlCols = Number(store['svcOrderTotalCols']);
 var jobsLogged = [];
 var svcOrderItems = []; 
 
//build the svcOrderValues array  
  for (var i = 0; i < ttlCols; i++) {
    if(store['svcOrder[' + i + ']'] != ' ' && store['svcOrder['+i+']'] != '') {
        svcOrderItems.push(store['svcOrder[' + i + ']']);
      }
   }
//get the jobsLogged array  
  var jobs = activeWorkbook.getActiveSheet().getRange(Number(store['rmaRow']), Number(store['AS-FOUND BENCH']), 1, ttlCols).getValues()[0];
  for (var i = 0; i < ttlCols; i++) {
    if(jobs[i]) {
        jobsLogged.push(jobs[i]);
      }
   }
   
//take back one kadam and dig here -- this is why that manual entry thing happened. check lesser than /greater than relationshipss, too
   if(svcOrderItems.length == jobsLogged.length) {
     return true;
   }
  return false;
}
//make this gorgeous someday
function findActivityCell(activityName){
//store the activityName
  documentProperties.setProperty('currentActivity', activityName);
  var pinNumber;
  var enterPIN;
//get the rmaRow
  var rmaRow = Number(store['rmaRow']);
  var reasonCol =  Number(store["REASON"]);
  var activityCol = Number(store[activityName]);
  var activityCell = activeSheet.getRange(rmaRow, activityCol);
  var reasonCell = activeSheet.getRange(rmaRow,reasonCol);
  
  activeSheet.setActiveRange(activityCell);
  
//quick check on activityCell contents
  var valueCheck = activeSheet.getActiveRange().getValue();

  if(valueCheck){
    var valueCheckAlert = ui.alert(activityName + " Logged.","This activity was logged as completed on " + valueCheck + ". \n\nDo you need to re-log this activity?", ui.ButtonSet.YES_NO);
   
   //if no, then same as always, clearRedirect();
      if(valueCheckAlert == ui.Button.NO || valueCheckAlert == ui.Button.CLOSE) {
        clearRedirect();
        
        } else if(valueCheckAlert == ui.Button.YES){
           activeSheet.getActiveRange().setValue("");
          findActivityCell(activityName);
    }
    
  } else if(!valueCheck) {
  switch(activityName) {
  
    case "RECEIVE":
      onReceive();
      break;
   
    case "JOB CLOSED":
//********** think about moving this whole Job Closed thang to its own function    
      enterPIN =  ui.prompt("JOB CLOSED: " + store['reason'], "Please enter your PIN", ui.ButtonSet.OK_CANCEL);
      pinNumber = enterPIN.getResponseText();
      if(enterPIN.getSelectedButton() == ui.Button.CANCEL || enterPIN.getSelectedButton() == ui.Button.CLOSE) {
        ui.alert("Logging canceled.", ui.ButtonSet.OK);
      }
      if(validatePINFormat(pinNumber)){
        var noteText = activeSheet.getActiveRange().getNote();
        noteText += date + " : JOB CLOSED : " + store['reason'] + " : Logged by " + store['pinName'];
        activeSheet.getActiveRange()
          .setHorizontalAlignment('center')
          .setValue(date)
          .setNote(noteText);
          //.protect();
        reasonCell.setValue(store['reason']);
          //.protect();
      }
 //send mail: recipient, subject, body, options
 //need an email address in here
      GmailApp.sendEmail('', 
      'RMA: ' + store['rmaID'] + ' Closed || ' + store['reason'], 
      'RMA: ' + store['rmaID'] + 
      '\nClosed by: ' + store['pinName'] + 
      '\nReason: ' + store['reason'] + 
      '\nCompany Name: ' + activeSheet.getRange(rmaRow, Number(store['companyName'])).getValue() + 
      '\nContact Name: ' + activeSheet.getRange(rmaRow, Number(store['contactFirstName'])).getValue() + " " + activeSheet.getRange(rmaRow, Number(store['contactLastName'])).getValue() + 
      '\nContact Phone: ' + activeSheet.getRange(rmaRow, Number(store['contactCell'])).getValue() +
      '\nContact Email: ' + activeSheet.getRange(rmaRow, Number(store['contactEmail'])).getValue() +
      '\n\nPlease confirm the details and follow up with the customer if necessary.', {
      'name': 'RMA - Job Closed',
      'from': '',
      'replyTo': ''
});
      clearRedirect();
      break;
    case "AS-FOUND BENCH":
    case "AS-LEFT BENCH":
    case "METER CALIBRATION":
    case "METER REPAIR":
    case "STEAM CLEAN":
    case "JOB COMPLETE":
 
      //has it been received?      
      var isReceived = activeSheet.getRange(rmaRow, Number(store['RECEIVE'])).getValue();
      
      if(!isReceived) {
      onReceive();
        
      } else if(isReceived) {
      
        if(checkServiceOrder(activityName) == false && activityName != "JOB COMPLETE"){
          ui.alert(activityName + " was not ordered by the customer" 
                   + '\n\n Note: To update the job scope, re-scan the barcode and select, "Update RMA"', ui.ButtonSet.OK);
          clearRedirect();
            
          } else if (checkJobsLogged() == false && activityName == "JOB COMPLETE") {
                ui.alert("RMA Incomplete", "Not all required activities have been logged.", ui.ButtonSet.OK);
               clearRedirect();
            } else {
                enterPIN =  ui.prompt("Logging: " + activityName, "Please enter your PIN", ui.ButtonSet.OK_CANCEL);
                
                   if(enterPIN.getSelectedButton() == ui.Button.CANCEL || enterPIN.getSelectedButton() == ui.Button.CLOSE){
                   ui.alert("Logging canceled.", ui.ButtonSet.OK);
                   clearRedirect();
                   
                 } else {
                     pinNumber = enterPIN.getResponseText();
                     
                     //*** take away one kadam and dig here 

                     if(validatePINFormat(pinNumber)){
                     //and finally, we're down to it ...  
                        if(activityName == 'JOB COMPLETE'){
                     //show the modal dialogue that holds the 'get more info @ jobComplete' form
                      var jcHtml = HtmlService.createTemplateFromFile('jobCompleteForm');
                      var jcDialogTitle = ' Job Completed || ' + store['rmaID']; 

                   //is there any chance at all that I wrote a simple version of this to the store? 
                   //please, let me have been overthinking it, but it turns out I really need it ... deeeeeeeNIED. effFFFF.
                   var xportMethodArr = activeSheet.getRange(rmaRow, Number(store['returnShipToCustomer']), 1, 3).getValues()[0];
                   var lenX = xportMethodArr.length;
                   
                   //what 
                    if(xportMethodArr[0]){
                      jcHtml.xportMethod = '<h3>RETURN SHIPPING</h3>' +
                                           '<div id="jcTrackingNumberDiv">' +
                                           '<p class="bigBold">Return-to Name:</p>' +
                                           '<p id="jcShipName"></p>' +
                                           '</div>' +
                                           '<div>' +
                                           '<p class="bigBold">Phone:</p>' +
                                           '<p id="jcShipPhone"></p>' +
                                           '</div>' +
                                           '<div class="divTracking">' +
                                           '<p class="bigBold">Enter Tracking #: </p>' +
                                           '<input id="jcTrackingNumber" type = "text">' +
                                           '</div>';
                };
                    if(xportMethodArr[1]){
                      jcHtml.xportMethod = '<h3>CUSTOMER PICK-UP</h3>' +
                                           '<p>Call the customer to schedule a time and day for him to pick up his completed meter. </p>' +
                                           '<div>' +
                                           '<p class="bigBold">Contact Name:</p>' +
                                           '<p id="jcPickupName"></p>' +
                                           '</div>' +
                                           '<div>' +
                                           '<p class="bigBold">Phone:</p>' +
                                           '<p id="jcPickupPhone"></p>' +
                                           '</div>';
                    };
                    if(xportMethodArr[2]){
                      jcHtml.xportMethod = '<h3>COMPANY DELIVERY</h3>' +
                                           '<p>Call the customer to schedule a time and day for Company to deliver his meter. </p>' +
                                           '<div>' +
                                           '<p class="bigBold">Deliver-to Name:</p>' +
                                           '<p id="jcDeliverName"></p>' +
                                           '</div>' +
                                           '<div>' +
                                           '<p class="bigBold">Phone:</p>' +
                                           '<p id="jcDeliverPhone"></p>' +
                                           '</div>';
                    };
                   jcHtml.jcDate = new Date().toDateString();
                   jcHtml.pinName = documentProperties.getProperty('pinName');
                    var jcDialogBox = jcHtml.evaluate();
                    jcDialogBox.setWidth(725).setHeight(650);
                    SpreadsheetApp.getUi().showModalDialog(jcDialogBox, jcDialogTitle);
                    
              } else {
              
              //just set the note
                var noteText1 = activeSheet.getActiveRange().getNote();
                noteText1 += date + " : Completed : Logged by " + documentProperties.getProperty('pinName') +"\n";
                activeSheet.getActiveRange()
                   .setHorizontalAlignment('center')
                   .setValue(date)
                   .setNote(noteText1);
                 //.protect();
                  clearRedirect();
              
              }
               }
                }
              }
       }
    }
  }
}

function jobComplete(trackingNumber) {
var rmaRow = Number(store.getProperty("rmaRow"));

//ah, Little River Band -- thank you :) 
//look around you / look up here / take time to make time / make time to be there 

//iterate through the property store and see wtfranchtoast is in there ... yes, I said FRANCH
//var things = store.getProperties();
//var noteText2 = activeSheet.getActiveRange().getNote();
//for (var k in things) {
//  noteText2 += k + " : " + things[k] +" ... ";
//}
//... don't be thinking that I don't want you ... 
  
  //set the note
  var noteText1 = activeSheet.getActiveRange().getNote();
  noteText1 += date + " : Completed : Logged by " + store.getProperty('pinName') +"\n" + "Tracking #:  " + trackingNumber;
  activeSheet.getActiveRange()
     .setHorizontalAlignment('center')
     .setValue(date)
     .setNote(noteText1);
     //.protect();

//send the email -- put this in it's own function so it can return a success value 
 sendSummaryEmail(trackingNumber);
 clearRedirect();

//SUCCESS! smells like chocolate oranges 
}

function sendSummaryEmail(trackingNumber){

var rmaRow = Number(store.getProperty("rmaRow"));
var completedActivities = activeSheet.getRange(rmaRow,Number(store.getProperty('activity[0]')), 1, 5).getValues()[0];
var confEmail = activeSheet.getRange(rmaRow, Number(store.getProperty('confirmationEmail'))).getValue();
var len = completedActivities.length;
var list = '\n';

for (var i = 0; i < len; i++) {
   if (!completedActivities[i]){
   break;
}
    list += '\n - ' + completedActivities[i];
}            
if(trackingNumber == 'noTrack'){
     list += '\n\nWe will be contacting you to schedule the return of your equipment.' 
     } else {
        list += '\n\nYour equipment is being shipped back to you. The tracking number is ' + trackingNumber;
     }

 GmailApp.sendEmail(confEmail, 
      'RMA: ' + store.getProperty('rmaID') + ' Job Complete', 
      'Greetings!' + 
      '\n\nWe\'re writing to let you know that your service order for RMA ' +' ' + store.getProperty('rmaID') + ' ' + ' is completed.' +
      '\n\nWe performed the following services: ' + list + 
      '\n\nDon\'t hesitate to contact us with any questions. ' +
      '\nWe can be reached at {{a phone number}} or by replying to this email.' +
      '\n\nSincerely,' +
      '\nThe Company Team',{
      'name': 'Company RMA - Job Complete',
      'from': '',
      'replyTo': ''
  }); 
}

function onReceive() {
  var enterPIN;
  var pinNumber;
  var receiveCol = Number(store['RECEIVE']);
  var rmaRow = Number(documentProperties.getProperty('rmaRow'));
  var proveitID = Number(store['PROVEit Meter ID']);
  var proveitName = Number(store['PROVEit Meter Name']);
  var leadTimeCol = Number(store['Lead Time']);
  var deadlineCol = Number(store['DEADLINE']);
  
  var receivedCell = activeSheet.getRange(rmaRow, receiveCol).activate();
  var isReceived = receivedCell.getValue();

  //has it already been logged as RECEIVED?
  if(isReceived){
   var isReceivedAlert = ui.alert("RECEIVE Logged.","This equipment was logged as received on " + isReceived.toDateString() + ".", ui.ButtonSet.OK);
    clearRedirect();
    
    } else {

  var recvQ1 = ui.alert("RECEIVING", "Does the equipment match what the customer has described on his RMA form?", ui.ButtonSet.YES_NO_CANCEL);
  
  if(recvQ1 == ui.Button.CANCEL || recvQ1 == ui.Button.CLOSE){
  ui.alert("Logging canceled.", ui.ButtonSet.OK);
    clearRedirect();
    
    } else if (recvQ1 == ui.Button.NO){ 
          documentProperties.setProperty('reason', 'Equipment mismatch');
          findActivityCell("JOB CLOSED");
      
  } else if(recvQ1 == ui.Button.YES){
     var recvQ2 = ui.alert("RECEIVING", "Is the serial number on the equipment an exact match to the serial number on the customer RMA form?", ui.ButtonSet.YES_NO_CANCEL);
     
     if(recvQ2 == ui.Button.CANCEL || recvQ2 == ui.Button.CLOSE){
     ui.alert("Logging canceled.", ui.ButtonSet.OK);
     clearRedirect();
       
     } else if(recvQ2 == ui.Button.NO){
          documentProperties.setProperty('reason', 'Serial number mismatch');
          findActivityCell("JOB CLOSED");
       
     } else if(recvQ2 == ui.Button.YES){
           //open prompt for PIN   
          enterPIN = ui.prompt("JOB RECEIVED", "Please enter your PIN", ui.ButtonSet.OK_CANCEL);
    
        if(enterPIN.getSelectedButton() == ui.Button.CANCEL || enterPIN.getSelectedButton() == ui.Button.CLOSE){
            ui.alert("Logging canceled.", ui.ButtonSet.OK);
            clearRedirect();
                       
        } else {
          pinNumber = enterPIN.getResponseText();
          if(validatePINFormat(pinNumber)){
          
          activeSheet.getRange(rmaRow,proveitID).setValue(store['meterID']); 
          activeSheet.getRange(rmaRow,proveitName).setValue(store['meterName']);

            activeSheet.getActiveRange()
              .setHorizontalAlignment('center')
              .setValue(date)
              .setNote(date + " : RECEIVED : Logged by " + documentProperties.getProperty('pinName'))
          }
          var rcvDate = activeSheet.getRange(rmaRow,receiveCol).getValue();
          var leadTime = activeSheet.getRange(rmaRow,leadTimeCol).getValue();
          var deadlinear = new Date(rcvDate.getTime() + leadTime*3600000*24);
          activeSheet.getRange(rmaRow,deadlineCol).setValue(deadlinear);
          clearRedirect();
        }
     }
  }
 }
}

function validatePINFormat(pinNumber) {
  ui.alert(pinNumber, ui.ButtonSet.OK);
//workingOnIt();

   //was anything entered in the field? is it four digits? is it a number?
    var regex = /^\d{4}$/;
    var testRegex = regex.test(pinNumber)
  if (!pinNumber || pinNumber.length != 4 ||  testRegex == 'false') {
    var enterPIN = ui.prompt("Invalid PIN x1","Please re-enter your PIN", ui.ButtonSet.OK_CANCEL);
    pinNumber = enterPIN.getResponseText();
    if (enterPIN.getSelectedButton() == ui.Button.CANCEL || enterPIN.getSelectedButton() == ui.Button.CLOSE){
      ui.alert("CANCELED", "Logging canceled", ui.ButtonSet.OK);
      return false;
    } else if(!pinNumber || pinNumber.length != 4 || testRegex == 'false') {
      ui.alert("Invalid PIN x2","Logging canceled", ui.ButtonSet.OK);
      return false;
    } else {
       return findPIN(pinNumber);
    }
  } else {
   return findPIN(pinNumber);
      }
}
  
//and quit passing reasonCell around like a bottle of cheap vodka 
function findPIN(pinNumber){

  ui.alert(pinNumber, ui.ButtonSet.OK);

  //what if PIN is not found in the pinSheet?
    var pinSheet = SpreadsheetApp.openById('1aIUuWvcfm6CW0eG3i3lpPQulqrj782AK9nTBzWYDD84').getSheetByName("Sheet1");
    var pinCell = pinSheet.createTextFinder(pinNumber).findNext();
    //yes, yesss, I see it ... 
  
    if (!pinCell){
      if(documentProperties.getProperty('pinTry') == '1'){
        ui.alert('PIN not found x2', 'Logging canceled.', ui.ButtonSet.OK);
        return false;
        
      } else {
        var pinReenter = ui.prompt('PIN not found', 'Try entering your PIN again:', ui.ButtonSet.OK_CANCEL);
        if (pinReenter.getSelectedButton() == ui.Button.CANCEL || pinReenter.getSelectedButton() == ui.Button.CLOSE){
          ui.alert("CANCELED", "Logging canceled", ui.ButtonSet.OK) 
          return false;
          
        } else {
          pinNumber = pinReenter.getResponseText();
          documentProperties.setProperty('pinTry', '1');
          validatePINFormat(pinNumber);
          }
       }
    } else if(pinCell) {
      var pinColNumber = pinCell.getColumn();
      var pinRowNumber = pinCell.getRow();
      var pinName = pinSheet.getActiveSheet().getRange(pinRowNumber, pinColNumber + 1).getValue();
      documentProperties.setProperty('pinName', pinName);
      workingOnIt('Item Logged');
      return true;
    //jumpForJoy(); 
      
   }         
}

function openUpdateForm(){
var rmaRow = Number(store['rmaRow']);
var ttlCols = Number(store['svcOrderTotalCols']);

var isReceived = activeSheet.getRange(rmaRow, Number(store['RECEIVE'])).getValue();
   if(!isReceived) {
      onReceive();      
   } else if(isReceived) {
  var html = HtmlService.createTemplateFromFile('htmlForm');
  
  //this is the reason I can't have nice things ...
  var currentRMAItems = activeWorkbook.getSheetByName(store['companyCode']).getRange(rmaRow, Number(store['activity[0]']), 1, ttlCols).getValues()[0];//this is what is on the RMA, what the customer originally ordered
  var allPossActivities = activeWorkbook.getSheetByName('rmaSheetTemplate').getRange(1, Number(store['AS-FOUND BENCH']), 1, ttlCols).getValues()[0];//this is every service that *can* be ordered
  
  //come back and check if this still needs to happen
    allPossActivities = allPossActivities.map(function(activity){
    return activity.toLowerCase(); 
    })
    
  var loggedActivities = activeWorkbook.getSheetByName(store['companyCode']).getRange(rmaRow, Number(store['AS-FOUND BENCH']), 1, ttlCols).getValues()[0];//these are the services that have been already worked
  
  var completedItems = [];
  var updateItems = [];
  
 //attaching the loggedActivities to the name of the item
    for (var i = 0; i < ttlCols; i++) {
       if (loggedActivities[i].length < 3) {
        updateItems.push(allPossActivities[i])
      }
       else {
          completedItems.push(allPossActivities[i]);
          completedItems.sort();
       } 
      }    

//build the uncheckedItems html string  
  var uncheckedLen = updateItems.length;
  var uncheckedItems = "";
  for (var i = 0; i < uncheckedLen; i++){
  uncheckedItems += "<div class='choice'><input type='checkbox' id='" + updateItems[i] + "' value='" + updateItems[i] + "'><label for='" + updateItems[i] + "'>" + updateItems[i] + "</label></div>";
}

//build the currentRMAItems html string
  var currentLen = completedItems.length;
  var currentRMAList = "";
  for (var i = 0; i < currentLen; i++){
  if(completedItems[i] == ''){
  break;
  }
  currentRMAList += "<li>" + completedItems[i] +"</li>";
}

 let completedItemsJoined = completedItems.join();
 documentProperties.setProperty('completedItems', completedItemsJoined);
 
 currentRMAItems = currentRMAItems.join();
 documentProperties.setProperty('currentRMAItems', currentRMAItems);
 
 html.uncheckedItems = uncheckedItems;
 html.currentRMAList = currentRMAList;
 
  var dialogTitle = ' Update RMA || ' + store['rmaID']; 
  var dialogBox = html.evaluate();
  dialogBox.setWidth(750).setHeight(625);
  SpreadsheetApp.getUi().showModalDialog(dialogBox, dialogTitle);
}

}

//pass checkedValues to this function from the frontend
function submitUpdate(checkedValues, customerRep){
//yes, yes, I know, the variables, yessss
var ttlCols = Number(documentProperties.getProperty('svcOrderTotalCols'));
var completedItems = documentProperties.getProperty('completedItems');


//****************
//**************
//***********
//********
//******
//****
//here

//omg, this is that weird thing that happens when there are no completed items lol! this is why I wrote that code ... damn it. 
if(completedItems){
completedItems = completedItems.split(",").sort();
checkedValues = checkedValues.concat(completedItems);
}

//okay, right here take back one kadam and dig
//checkedValues = checkedValues.sort();

var enterPIN =  ui.prompt('Update RMA || '+ store['rmaID'], 'Please enter your PIN', ui.ButtonSet.OK_CANCEL);
var pinNumber = enterPIN.getResponseText();
   if(enterPIN.getSelectedButton() == ui.Button.CANCEL || enterPIN.getSelectedButton() == ui.Button.CLOSE) {
   ui.alert("Update RMA canceled", ui.ButtonSet.OK);
   clearRedirect();
} else {
   var wtf = validatePINFormat(pinNumber);
   if(wtf){
   
   //I really just need to lock in these particular values 
   
   var pinName = documentProperties.getProperty('pinName');
   var rmaID = documentProperties.getProperty('rmaID');
   var companyTab = activeWorkbook.getSheetByName(store['companyCode']);
   var rmaLocation = companyTab.createTextFinder(store['rmaID']).findNext();
   var rmaRow = Number(store['rmaRow']);
   var activity0 = Number(documentProperties.getProperty('activity[0]'));
   var companyName = activeSheet.getRange(rmaRow,Number(store['companyName'])).getValue();
   var contactName = activeSheet.getRange(rmaRow, Number(store['contactFirstName'])).getValue() + " ";
   contactName += activeSheet.getRange(rmaRow, Number(store['contactLastName'])).getValue();
   var contactCell = activeSheet.getRange(rmaRow, Number(store['contactCell'])).getValue();
   var contactEmail = activeSheet.getRange(rmaRow, Number(store['contactEmail'])).getValue();
   
   var noteText = companyTab.getRange(rmaRow, rmaLocation.getColumn()).getNote();
   noteText += '\n ' + date + '\nRMA UPDATED : per '+ customerRep + '\nLogged by ' + pinName + '\n';
   companyTab.getRange(rmaLocation.getRow(), rmaLocation.getColumn()).setNote(noteText);
   
   var checkedRange = companyTab.getRange(rmaRow,activity0, 1, checkedValues.length);
  //  var clearedRange = companyTab.getRange(rmaRow,activity0, 1, ttlCols).clear();
   checkedValues = [checkedValues];
   checkedRange.setValues(checkedValues); //checkedValues = checkedValues.concat(completedItems)

    var emailObj = {rmaID,pinName,companyName,contactName,contactCell,contactEmail};
     sendEmail(emailObj);
       clearRedirect();
     
     } 
   }
}
//will need email addresses in here
function sendEmail(emailObj){
   GmailApp.sendEmail('', 
      'RMA: ' + emailObj.rmaId + ' Updated', 
      'RMA: ' + emailObj.rmaId + 
      '\nUpdated by: ' + emailObj.pinName + 
      '\nCompany Name: ' + emailObj.companyName + 
      '\nContact Name: ' + emailObj.contactName + 
      '\nContact Phone: ' + emailObj.contactCell +
      '\nContact Email: ' + emailObj.contactEmail +
      '\n\nPlease confirm the details and follow up with the customer if necessary.', {
      'name': 'RMA Updated',
      'from': '',
      'replyTo': ''
  });
}

function include(fileName){
return HtmlService.createHtmlOutputFromFile(fileName).getContent();
}

function clearRedirect(){
var scanSheet = activeWorkbook.getSheetByName('scan');

scanSheet.getRange('D3').clear()
  .setBackgroundColor('#fff');
  
scanSheet.getRange('C4').clear()
  .setBackgroundColor('#fff');

var documentProperties = PropertiesService.getDocumentProperties();
    documentProperties.deleteProperty('currentActivity');
    documentProperties.deleteProperty('companyCode');
    documentProperties.deleteProperty('rmaRow');
    documentProperties.deleteProperty('rmaID');
    documentProperties.deleteProperty('meterID');
    documentProperties.deleteProperty('meterName');
    documentProperties.deleteProperty('reason');
    documentProperties.deleteProperty('pinTry');
    documentProperties.deleteProperty('pinName');
    documentProperties.deleteProperty('currentRMAItems');
    documentProperties.deleteProperty('completedItems');    
    
   var svcOrderTotalCols = Number(documentProperties.getProperty('svcOrderTotalCols'));
   
 //delete however many svcOrder items have been stored 
  for (var i = 0; i < svcOrderTotalCols; i++) {
  documentProperties.deleteProperty('svcOrder['+ i +']');
    }
return scanBarcode();
}
