function onOpen() 
{
  let ui = SpreadsheetApp.getUi();
  ui.createMenu('Mail')
      .addItem('Send Mail to Selection', 'HandleApproval')
      .addToUi();
}

function HandleApproval(){
  let ui = SpreadsheetApp.getUi();
  let response = ui.prompt('Enter Folder ID','Want to continue?', ui.ButtonSet.YES_NO);
  if (response.getSelectedButton() == ui.Button.YES) {
    let id = response.getResponseText();
//    Logger.log('Folder ID is %s.', id);
    let folder = DriveApp.getFolderById(id);
    response = ui.alert(folder.getName(),'Want to continue?', ui.ButtonSet.YES_NO);
    if(response == ui.Button.YES){
      GetCells(folder,ui);
    }
    else{
      ui.alert("aborted");
    }
  } else if (response.getSelectedButton() == ui.Button.NO) {
  Logger.log('The user didn\'t want to provide id.');
  } else {
    Logger.log('The user clicked the close button in the dialog\'s title bar.');
  }
}



function GetCells(folder,ui)
{
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getActiveSheet();
  let selection = sheet.getSelection();
  let range = selection.getActiveRange();
  
  let rows = range.getNumRows();
  let initial_row = range.getRow();
  for(let i=0;i<rows;i++){
    let name = sheet.getRange(initial_row + i, 1);
    let email = sheet.getRange(initial_row + i, 3);
    let status = sheet.getRange(initial_row + i, 4);
    if(status.getValue() !== ""){
      ui.alert(name.getValue() + " has already been messaged");
      continue;
    }
    let filename = sheet.getRange(initial_row + i, 5);
    let id = sheet.getRange(initial_row + i, 6);
    let files = folder.getFilesByName(filename.getValue() + ".jpg");
    if(files.hasNext()){
      let file = files.next();
      SendMail(name.getValue(),email.getValue(),file,status,id);
    }    
  }
}


function SendMail(name,email,file,status,id){
//  let ui = SpreadsheetApp.getUi();
//  ui.alert(name + " " + email + " " + file.getName());
  let templ = HtmlService
      .createTemplateFromFile('EmailTemplate');
  
  templ.candidate = name;
  let message = templ.evaluate().getContent();
  
  MailApp.sendEmail({
    to: email,
    subject: "Certificate",
    htmlBody: message,
    attachments: [file.getAs(MimeType.JPEG)]
  });
 	id.setValue(file.getId());
  	status.setValue("send");
  
  SpreadsheetApp.flush();
}
