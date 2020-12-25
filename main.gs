function go(){
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  // get value of the top right cell, which holds the name of the json file
  var jsonName = sheet.getRange('A1').getValue();

  // fetch file from google drive and parse it
  var persons = parseJSONfromDrive(jsonName);

  // remove nonmember objects in the array of people
  persons = shaveNonmembers(persons);

  // sort the objects in persons[] by Last Name, then secondly by age
  persons = sortByAtt(persons,"lastName","ageCategoryId");

  // compress people objects into households
  var data = householdSquish(persons);

  // orient spreadsheet based on selected cell
  const topRowNum = sheet.getActiveCell().getRow(); // user can select what row to begin the table at, if they want some empty header rows

  // header row
  var headers = ["Last Name","First Name","Status","Last Contact","Last Dinner","Received Ward Plan?","Commitments made","Kept commitments?","References"];
  sheet.getRange(topRowNum,1,1,headers[0].length).setValues([headers]);

  // copy household array to the spreadsheet
  sheet.getRange(topRowNum+1,1,data.length,data[0].length).setValues(data);

  // apply formatting rules!!!
  formatSheet(sheet,topRowNum);
}


// fetch file from google drive and parse it
function parseJSONfromDrive(fileName){
  // search Google Drive for the file
  var jsonFiles = DriveApp.getFilesByName(fileName);
  // Drive doesn't understand how to fetch text from a file, so pull the contents as binary (blob) first, and then interpret as a string
  var jsonTxt = jsonFiles.next().getBlob().getDataAsString("UTF-8");
  
  // JSON.parse() tries to digest \n escape characters, so find-replace with \\n so they reduce to a single backslash
  jsonTxt = jsonTxt.replace(/\n/g, '\\n');
  // the downloaded file has lots of weird zero-width characters, so delete them
  jsonTxt = jsonTxt.replace(/[\u200B-\u200D\uFEFF]/g, '');

  // parse jsonTxt as an object, which can be laid out into a spreadsheet
  var flData = JSON.parse(jsonTxt);
  // the data from area book comes in an array called {areabookData}.persons
  return flData.persons;
}

// remove nonmember objects in the array of people
function shaveNonmembers(persons){
  // "missionName" is an easy column to distinguish members from nonmembers- all members will have an empty cell here, but nonmembers will have the name of the mission
  // delete all person rows with something in missionName
  var p=0;
  do{
    if(!persons[p]["missionName"]){ // check for an empty string under the key missionName
      p++; // only increment if a person wasn't deleted
    }else{
      persons.splice(p,1); // delete 1 entry at location p- they're a non-member
    }
  }while(p<persons.length); // check at the end of each loop because the length is shrinking

  return persons;
}

// order an array of objects or an array of arrays by a specific attribute/index
function sortByAtt(arr, key, key2=null){ // arr must contain objects or arrays
  return arr.sort(
    (a,b) => // describe how to compare the objects
    ( a[key]>b[key] )? 1:
      (a[key] === b[key])? ((a[key2] < b[key2]) ? 1 : -1)
    :-1
  );
} ////TODO////


// compress people objects into households
function householdSquish(persons){
  // we want an array of household IDs so we can search them
  var houses = [];
  var data = [];

  // accumulate people into their homes by household ID
  for(var p=0;p<persons.length;p++){
    // the current person's household has either been created already or must be created now
    var hloc = houses.indexOf( persons[p]["householdGuid"] );

    if( hloc == -1 ){ // -1 means it wasn't found in houses[]
      hloc = data.length; // make a new household on the end
      data.push([]); // add empty array to the end
      data[hloc].push( persons[p]["lastName"] );
      data[hloc].push( persons[p]["firstName"] );

      houses.push( persons[p]["householdGuid"] ); // add the new household to the list

    }else{ // found the household!!
      data[hloc][1] += ", " + persons[p]["firstName"]; // append the person's name to the list of first names in the household
    }
  }
  // underscorejs.org/#sortBy

  return data;
}

// formatting to make everything pretty
function formatSheet(sheet,topRowNum){
  ////TODO////
};


