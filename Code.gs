function onOpen() {
  const ui = SpreadsheetApp.getUi();
  const menu1 = ui.createMenu('AutoFill Docs');
  menu1.addItem('Create New Docs', 'createNewGoogleDocs')
  menu1.addToUi();

  const menu2 = ui.createMenu('Merge Docs');
  menu2.addItem('Merge Docs', 'mergeGoogleDocs')
  menu2.addToUi();
}

function test(){
  var fieldNumber = 8;
  const folder_list = getFolderIDList();
  const folderID = matchingFolderID(fieldNumber,folder_list);
}

function getFolderIDList(){
  const destinationFolder = DriveApp.getFolderById('1XBYwb_EchO03pQeOQfE-PBdMF6ypkK6j'); // TBAA/2024 folder
  var field_folders = destinationFolder.getFolders();
  var folder_list = [];
  while (field_folders.hasNext()){
    var field_folder = field_folders.next();
    folder_list.push([field_folder.getName(), field_folder.getId()]);
  }

  return folder_list;
}

function matchingFolderID(fieldNumber, folder_list) {
  const folder_dict = new Map(folder_list.map(([v, k]) => [v, k]));
  var folderID;
  switch (fieldNumber){
    case "1":
      folderID = folder_dict.get("Field1");
      break;
    case "2":
      folderID = folder_dict.get("Field2");
      break;
    case "3B":
      folderID = folder_dict.get("Field3B");
      break;
    case "3C":
      folderID = folder_dict.get("Field3C");
      break;
    case "4":
      folderID = folder_dict.get("Field4");
      break;      
    case "5": 
      folderID = folder_dict.get("Field5");
      break;
    case "6":
      folderID = folder_dict.get("Field6");
      break;
    case "7":
      folderID = folder_dict.get("Field7");
      break;
  }
  return folderID
}


const documentProperties = PropertiesService.getDocumentProperties();

function createNewGoogleDocs() {
  //This value should be the id of your document template that we created in the last step

    const googleDocTemplate = DriveApp.getFileById('1c6GKGb2tyLVQVhoPWVYHN3IqQAwszTRYM_YDso3rdw0'); //the doc template
    Logger.log(googleDocTemplate.getName());
    Logger.log(googleDocTemplate.getBlob().isGoogleType());
  
  //This value should be the id of the folder where you want your completed documents stored

  var folder_dict = getFolderIDList();
  Logger.log(folder_dict);
  //Here we store the sheet as a variable
  
  var ss = SpreadsheetApp.openById('11t7kwG9K0z0Q12SvPTNRHA-zbBQuolKmO'); //the game list sheet
  const sheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName('Sheet1')

  const [header, ...values] = sheet.getDataRange().getValues();
  const statusIndex = header.indexOf("status");
  const totalValues = values.length;
  
  //Now we get all of the values as a 2D array
  const rows = sheet.getDataRange().getValues();
  //docfinal = DocumentApp.create("All Game Cards");
  //var bodyfinal = docfinal.getBody();
  //var textfinal = bodyfinal.editAsText();

  var fieldNumber;
  var folderID;
  var numfiles = 18;
  //Start processing each spreadsheet row
  //

  const startTime = new Date();
  const triggerId = documentProperties.getProperty("timeOutTriggerId");

  for (let i = 0; i < totalValues; i++) {
      const row = values[i];
      const status = row[statusIndex];

      const editabletemp = googleDocTemplate;

      //if (index === 0) return; //do nothing function if at the header row
      //if (index > numfiles) return; //keep iterating until you hit row 143
      //Using the row data in a template literal, we make a copy of our template document in our destinationFolder
        
      if (status !== "DONE"){
        Logger.log("i type:" + typeof(i));
        Logger.log("statusIndex type:" + typeof(statusIndex));
        const row1 = i + 1 + 1;
        const col1 = statusIndex + 1;
        sheet.getRange(row1, col1).setValue("Processing...");
        SpreadsheetApp.flush();

        fieldNumber = row[7]; //field number of the game we're on in the list

        folderID=matchingFolderID(fieldNumber, folder_dict);
        
        const copy = editabletemp.makeCopy(`${row[1]} Game ${row[0]} Field ${row[7]}`, DriveApp.getFolderById(folderID)); //creates doc i.e. "8/28/22 Game 100 Field 9" in the matching field folder

        //Once we have the copy, we then open it using the DocumentApp
        const doc = DocumentApp.openById(copy.getId());

        //All of the content lives in the body, so we get that for editing
        const body = doc.getBody();
        Logger.log("writing values of game " + row[0])
        //In these lines, we replace our replacement tokens with values from our spreadsheet row
        body.replaceText('{{Game Number}}', row[0]);
        body.replaceText('{{Game Date}}', row[1]);
        body.replaceText('{{Game Start Time}}', row[2]);
        body.replaceText('{{Gender}}', row[7]);
        body.replaceText('{{Age}}', row[6]);
        body.replaceText('{{Field Number}}', row[3]);
        body.replaceText('{{Home Club}}', row[4]);
        body.replaceText('{{Away Club}}', row[5]);

        //We make our changes permanent by saving and closing the document
        //textfinal.appendText(body);
        doc.saveAndClose();
        sheet.getRange((i + 1) + 1, statusIndex + 1).setValue("DONE");
        SpreadsheetApp.flush();
        Logger.log("game " + row[0] + " complete");
        //Store the url of our new document in a variable
        //Write that value back to the 'Document Link' column in the spreadsheet. 
      }
    
  }//docfinal.saveAndClose();
}


const isTimeUp = (startTime) => new Date().getTime() - startTime.getTime() > 3000;

function mergeGoogleDocs() {
  var folder_list = getFolderIDList();

  for (var m = 0; m < folder_list.length; m++){
    Logger.log(folder_list[m][0]);
    Logger.log(folder_list[m][1]);

    var folder = DriveApp.getFolderById(folder_list[m][1]);
    var files = folder.getFiles();
    
    var idList = [];
    while (files.hasNext()) {
      var file = files.next();
      idList.push(file.getId());
    }
    var base = DriveApp.getFileById(idList[0]).makeCopy(`${folder.getName()} Game Cards`); // duplicate the first (0th) file in the folder to make it the base file where the rest will be appended
    var baseID = base.getId(); //get doc id of the new file
    var baseDoc = DocumentApp.openById(baseID); //open the base doc
    var body = baseDoc.getActiveSection(); //make the body of the document a variable

    for (var i = 1; i < idList.length; ++i ) {
      var otherDoc = DocumentApp.openById(idList[i])
      var otherBody = otherDoc.getActiveSection(); //make the body of the 2nd document in your folder a variable
      var totalElements = otherBody.getNumChildren(); //grab the "children" of the body of your secondary file

      for( var j = 0; j < totalElements; ++j ) { //for each element in the secondary file
        var element = otherBody.getChild(j).copy(); //grab the element from the secondary file
        var type = element.getType(); //note the type of the element
        try{
          if( type == DocumentApp.ElementType.PARAGRAPH )
            body.appendParagraph(element); //add element if paragraph
          else if( type == DocumentApp.ElementType.TABLE )
            body.appendTable(element); //add element if table
          else if( type == DocumentApp.ElementType.LIST_ITEM )
            body.appendListItem(element); //add element if list
          else
            throw new Error("Unknown element type: "+type); //throw an error if the type of the element is not captured in a condition above
        }catch (e){
          Utilities.sleep(10); //sleep for a sec
        }
      }

      otherDoc.saveAndClose(); //close the other doc to prep for opening the next
    }
  }

  Logger.log(idList[0]);

  var docIDs = files;

  baseDoc.saveAndClose();
}
