var ssId = '';//id of the rma app workbook

function doGet() {
  var template = HtmlService.createTemplateFromFile('vmetricsRMAForm');
  template.companyNameList = getList('companyName', 'lists');
  template.meterBrands = getList('meterBrand', 'lists');
  //do I need companyName(s) and companyCode(s) in one place? hello companyObj
  //var companyObj = {};
  return template.evaluate();
};


function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function getList(listName, sheetName) {
  var ss = SpreadsheetApp.openById(ssId);
  sheetName = ss.getSheetByName(sheetName);
  var listLoc = sheetName.createTextFinder(listName).findNext();
  var listRow = listLoc.getRow();
  var listCol = listLoc.getColumn();
  var theListArr = sheetName.getRange(listRow, listCol).getDataRegion().getValues();
  var len = theListArr.length;
  var theListHtml = '';

  //head's up on the cheap hack right here 
  for (var i = 1; i < len; i++) {
    theListHtml += '<option>' + theListArr[i] + '</option>';
  }
  return theListHtml;
};

function prevEntry(companyCode) {
  var ss = SpreadsheetApp.openById(ssId);

  var companyTab = ss.getSheetByName(companyCode);
  var lastRow = companyTab.getLastRow();
  //because the arrays keep fuckin gup
  var confirmationEmailCol = companyTab.createTextFinder('confirmationEmail').findNext().getColumn();
  var contactFirstNameCol = companyTab.createTextFinder('contactFirstName').findNext().getColumn();
  var contactLastNameCol = companyTab.createTextFinder('contactLastName').findNext().getColumn();
  var contactEmailCol = companyTab.createTextFinder('contactEmail').findNext().getColumn();
  var contactCellCol = companyTab.createTextFinder('contactCell').findNext().getColumn();

  var returnFirstNameCol = companyTab.createTextFinder('returnFirstName').findNext().getColumn();
  var returnLastNameCol = companyTab.createTextFinder('returnLastName').findNext().getColumn();
  var returnAddressCol = companyTab.createTextFinder('returnAddress').findNext().getColumn();
  var returnSuiteCol = companyTab.createTextFinder('returnSuite').findNext().getColumn();
  var returnCityCol = companyTab.createTextFinder('returnCity').findNext().getColumn();
  var returnStateCol = companyTab.createTextFinder('returnState').findNext().getColumn();
  var returnZipCol = companyTab.createTextFinder('returnZip').findNext().getColumn();

  var confirmationEmailVal = companyTab.getRange(lastRow, confirmationEmailCol).getValue();
  var contactFirstNameVal = companyTab.getRange(lastRow, contactFirstNameCol).getValue();
  var contactLastNameVal = companyTab.getRange(lastRow, contactLastNameCol).getValue();
  var contactEmailVal = companyTab.getRange(lastRow, contactEmailCol).getValue();
  var contactCellVal = companyTab.getRange(lastRow, contactCellCol).getValue();

  var returnFirstNameVal = companyTab.getRange(lastRow, returnFirstNameCol).getValue();
  var returnLastNameVal = companyTab.getRange(lastRow, returnLastNameCol).getValue();
  var returnAddressVal = companyTab.getRange(lastRow, returnAddressCol).getValue();
  var returnSuiteVal = companyTab.getRange(lastRow, returnSuiteCol).getValue();
  var returnCityVal = companyTab.getRange(lastRow, returnCityCol).getValue();
  var returnStateVal = companyTab.getRange(lastRow, returnStateCol).getValue();
  var returnZipVal = companyTab.getRange(lastRow, returnZipCol).getValue();

  var prevEntryObj = {
    confirmationEmail: confirmationEmailVal,
    contactFirstName: contactFirstNameVal,
    contactLastName: contactLastNameVal,
    contactEmail: contactEmailVal,
    contactCell: contactCellVal,
    returnFirstName: returnFirstNameVal,
    returnLastName: returnLastNameVal,
    returnAddress: returnAddressVal,
    returnSuite: returnSuiteVal,
    returnCity: returnCityVal,
    returnState: returnStateVal,
    returnZip: returnZipVal
  }


  return prevEntryObj;
}

function scrubForm(formObj) {
  //this can't fire until the formObj is totally complete
  //so -- validation? 
  //Validation. 

  //all I need to know is if all the required fields are filled, that's it
  //I'm too tired to make this pretty ... here comes ugly
  if (!formObj.contactFirstName || !formObj.contactLastName || !formObj.contactCell || !formObj.contactEmail || !formObj.confirmationEmail || !formObj.companyName) {
    return 8
  };
  if (!formObj.meterBrand || !formObj.meterType || !formObj.meterModel || !formObj.serialNumber || !formObj.flowRate) {
    return 2
  };
  if (formObj.returnShipToCustomer) {
    if (!formObj.returnCompanyName || !formObj.returnFirstName || !formObj.returnLastName || !formObj.returnAddress || !formObj.returnCity || !formObj.returnState || !formObj.returnZip) {
      return 3
    }
  };

  if (!formObj.asFoundBench && !formObj.asLeftBench && !formObj.meterCal && !formObj.meterRepair && !formObj.steamClean) {
    return 4
  } else {
    var AOK = submitForm(formObj);
    return AOK;
  }
} //scrubForm has served its purpose -- why is it still here? 

function submitForm(formObj) {
  var companyCode = formObj.companyCode;
  //var companyTab = ss.getSheetByName(companyCode);
  var date = new Date();

  //scrub leadTime for the numbers only -- and keep it simple, there are only 4 options, ffs
  var deadline = '';
  if (formObj.lead1 == true) {
    deadline = 1;
  }
  if (formObj.lead3 == true) {
    deadline = 3;
  }
  if (formObj.lead7 == true) {
    deadline = 7;
  }
  if (formObj.lead14 == true) {
    deadline = 14;
  }

  var provingType = '';
  if (formObj.allocationProving == true) {
    provingType = 'Allocation'
  }
  if (formObj.custodyProving == true) {
    provingType = 'Custody Transfer'
  }
  if (formObj.noneProving == true) {
    provingType = 'none'
  }

  var activityNamesArr = ['As-found Bench', 'As-left Bench', 'Meter Calibration', 'Meter Repair', 'Steam Clean'];
  var truthArr = [formObj.asFoundBench, formObj.asLeftBench, formObj.meterCal, formObj.meterRepair, formObj.steamClean];
  var activities = [];

  for (var i = 0; i < truthArr.length; i++) {

    if (truthArr[i]) {
      activities[i] = activityNamesArr[i];
    } else {
      activities[i] = '';
    }
    activities.sort();
    activities.reverse();
  }

  var companyName = formObj.companyName;
  var confirmationEmail = formObj.confirmationEmail;
  //create the barCode
  var timeCode = Date.parse(date) / 1000;
  var barCode = '*' + companyCode + timeCode + '*';
  var rmaID = companyCode + timeCode;

  //send the barcode to the RMAs spreadsheet from here, 
  var rmaSS = SpreadsheetApp.openById('18TE9PvbCdGxYSnOcRsiVMOcFValZ4seeH3w7bdrhqGg');
  var companyTab = rmaSS.getSheetByName(companyCode);

  //yay it's working, but massage the formObj before sending it over, to match up with what's already going on
  companyTab.appendRow([date, companyName, formObj.confirmationEmail, formObj.contactFirstName, formObj.contactLastName, formObj.contactEmail, formObj.contactCell, formObj.returnCompanyName, formObj.returnFirstName, formObj.returnLastName, formObj.returnAddress, formObj.returnSuite, formObj.returnCity, formObj.returnState, formObj.returnZip, formObj.meterBrand, formObj.meterType, formObj.meterModel, formObj.flowRate, formObj.shipToVmetrics, formObj.dropOffToVmetrics, formObj.requestVmetricsPickup, formObj.returnShipToCustomer, formObj.returnCustPickup, formObj.returnHandDelivery, formObj.allocationProving, formObj.custodyProving, formObj.noneProving, formObj.lead1, formObj.lead3, formObj.lead7, formObj.lead14, formObj.asFoundBench, formObj.asLeftBench, formObj.meterCal, formObj.meterRepair, formObj.steamClean, deadline, provingType, activities[0].toLowerCase(), activities[1].toLowerCase(), activities[2].toLowerCase(), activities[3].toLowerCase(), activities[4].toLowerCase(), formObj.serialNumber, rmaID]);
  companyTab.createTextFinder(rmaID).findNext().setNote('Submitted online: ' + date + '\nContact Name: ' + formObj.contactFirstName + formObj.contactLastName + '\nContact Phone: ' + formObj.contactCell + '\n');

  //create the attachment 
  var rmaTemplate = DriveApp.getFileById('1W_iuDhUg1UMXiHvqiy-rlLhJkJwRKx16_befFSISFYk');
  var rmaFolder = DriveApp.getFolderById('1Y6Lv5eAzjUg56fHHEj8hfSsNwssWjYk1')
  var rmaTemplateCopy = rmaTemplate.makeCopy(companyName + ' RMA: ' + timeCode + " SN: " + formObj.serialNumber, rmaFolder);

  //  open the new file created by making a copy of the template
  var rmaDoc = DocumentApp.openById(rmaTemplateCopy.getId());

  //get the doc body in order to make changes to template contents
  var rmaDocBody = rmaDoc.getBody();
  var contactName = formObj.contactFirstName + ' ' + formObj.contactLastName;
  var shiptoName = formObj.returnFirstName + ' ' + formObj.returnLastName;

  //replaceText methods
  rmaDocBody.replaceText('{{companyName}}', companyName);
  rmaDocBody.replaceText('{{shiptoName}}', shiptoName);
  rmaDocBody.replaceText('{{address1}}', formObj.returnAddress);
  rmaDocBody.replaceText('{{address2}}', formObj.returnSuite);
  rmaDocBody.replaceText('{{city}}', formObj.returnCity);
  rmaDocBody.replaceText('{{state}}', formObj.returnState);
  rmaDocBody.replaceText('{{zip}}', formObj.returnZip);
  rmaDocBody.replaceText('{{contactName}}', contactName);
  rmaDocBody.replaceText('{{contactPhone}}', formObj.contactCell);
  rmaDocBody.replaceText('{{contactEmail}}', formObj.contactEmail);
  rmaDocBody.replaceText('{{brand}}', formObj.meterBrand);
  rmaDocBody.replaceText('{{type}}', formObj.meterType);
  rmaDocBody.replaceText('{{model}}', formObj.meterModel);
  rmaDocBody.replaceText('{{serialNumber}}', formObj.serialNumber);
  rmaDocBody.replaceText('{{serviceOrder1}}', activities[0]);
  rmaDocBody.replaceText('{{serviceOrder2}}', activities[1]);
  rmaDocBody.replaceText('{{serviceOrder3}}', activities[2]);
  rmaDocBody.replaceText('{{serviceOrder4}}', activities[3]);
  rmaDocBody.replaceText('{{serviceOrder5}}', activities[4]);
  rmaDocBody.replaceText('{{barCode}}', barCode);
  rmaDocBody.replaceText('{{provingType}}', provingType);
  rmaDocBody.replaceText('{{flowRate}}', formObj.flowRate);
  rmaDocBody.replaceText('{{leadTime}}', deadline);

  //save and close the document
  rmaDoc.saveAndClose();
  var rmaDocID = DriveApp.getFileById(rmaDoc.getId());
  var rmaPDF = rmaDocID.getAs(MimeType.PDF);
  var rmaBlob = rmaPDF.copyBlob();
  var fileName = companyName + ' RMA: ' + timeCode + ' SN: ' + formObj.serialNumber;


  //send mail: recipient, subject, body, options
  GmailApp.sendEmail(confirmationEmail, 'RMA Form Attached SN: ' + formObj.serialNumber, 'Your RMA has been received. Please print out the attached file and include it with your equipment when you ship it or bring it to Company.', {
    'name': 'Company RMA for SN ' + formObj.serialNumber,
    'attachments': [rmaPDF]
  });

  return 'AOK';
  //at the very end, return a switch to trigger the form to reset
}