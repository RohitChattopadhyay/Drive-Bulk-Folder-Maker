function onOpen(e) {
  SpreadsheetApp.getUi().createMenu("Folder Maker").addItem("Generate","main").addToUi(); 
}

function onInstall(e){
  onOpen(e);
}

function main() {
  var active_spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = active_spreadsheet.getSheets()[0];

  var last_row = sheet.getLastRow();
  
  var ui = SpreadsheetApp.getUi();

  if(last_row<2){
    ui.alert("Please add student rows and try again.")  ;
    return;
  }

  // Master Folder
  var master_folder,change_owner,new_owner;
  var resp = ui.alert('Do you have existing master folder? \nStudent folders will go inside this folder.', ui.ButtonSet.YES_NO);
  if (resp == ui.Button.NO) {
    var folder_resp = ui.prompt('Enter new folder name to be created:');
    if (folder_resp.getSelectedButton() == ui.Button.OK) {
      var new_folder_name = folder_resp.getResponseText();
      master_folder =  DriveApp.createFolder(new_folder_name);
    } else {
      return;
    }
  } else {    
    var folder_resp = ui.prompt('Enter target folder link:');
    if (folder_resp.getSelectedButton() == ui.Button.OK) {
      var new_folder_link = folder_resp.getResponseText();
      var master_folder_id = getIdFrom(new_folder_link);
      master_folder = DriveApp.getFolderById(master_folder_id);
    } else {
      return;
    }
  }

  resp = ui.alert('Do you want to transfer ownership of student folders?\n Default: You', ui.ButtonSet.YES_NO_CANCEL);

  if (resp == ui.Button.YES){
    change_owner =  true;
    var owner_resp = ui.prompt('Enter new owner email:');
    if (owner_resp.getSelectedButton() == ui.Button.OK) {
      var new_owner = owner_resp.getResponseText().replace(/\s/g, "");
      if ( new_owner == Session.getEffectiveUser().getEmail() )
        change_owner = false;
    } else {
      return;
    }
  } else if(resp == ui.Button.NO) {
    change_owner =  false;
  } else{
    return;
  }

  // Folder creation loop

  // confirmation

  resp = ui.alert('Task Information:\n\tCreate ' + (last_row-1) + ' folders\nClick Yes to Continue', ui.ButtonSet.YES_NO_CANCEL);
  if(resp == ui.Button.YES){
    createFolders(sheet,master_folder, change_owner, new_owner);
    sheet.getRange(1,3).setFormula('=HYPERLINK("'+master_folder.getUrl()+'", "Master Folder")');
  } else 
    return; 
}


function createFolders(sheet, master, change_owner = false, new_owner=""){
  var last_row = sheet.getLastRow();

  for(var i=2;i<last_row+1;i++){

    sheet.getRange(i,3).setValue("Creating Folder");
    var folder_name = sheet.getRange(i, 1).getValue();
    var student_email = sheet.getRange(i,2).getValue();    
    var folder = DriveApp.createFolder(folder_name);
    sheet.getRange(i,3).setValue("Folder Created");
    
    folder.addEditor(student_email);
    sheet.getRange(i,3).setValue("Permission Granted");
    
    if(change_owner){
      folder.setOwner(new_owner)
    }

    sheet.getRange(i,3).setValue("Moving Folder to Master");
    master.addFolder(folder);//put student folder in teacher folder
    
    sheet.getRange(i,3).setValue(folder.getUrl());
  }
}

function getIdFrom(url) {
  var id = "";
  var parts = url.split(/^(([^:\/?#]+):)?(\/\/([^\/?#]*))?([^?#]*)(\?([^#]*))?(#(.*))?/);
  if (url.indexOf('?id=') >= 0){
     id = (parts[6].split("=")[1]).replace("&usp","");
     return id;
   } else {
   id = parts[5].split("/");
   //Using sort to get the id as it is the longest element. 
   var sortArr = id.sort(function(a,b){return b.length - a.length});
   id = sortArr[0];
   return id;
   }
 }
