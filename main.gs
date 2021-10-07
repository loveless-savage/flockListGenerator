// main function
function go(){
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  // get value of the top right cell, which holds the name of the json file
  var jsonName = sheet.getRange('A1').getValue();
 
  // fetch file from google drive and parse it
  var persons = parseJSONfromDrive(jsonName);
 
  // remove nonmember objects in the array of people
  persons = shaveNonmembers(persons);
 
  // sort the objects in persons[] by household, then secondly by age
  persons = sortByAtt(persons,"householdGuid",true,"ageCategoryId",false);
  // compress people objects into households
  var data = householdSquish(persons);
  // now that the households have been condensed, we can sort them by last name without worrying about multiple last names in one household
  data = sortByAtt(data,0,true);
 
  // orient spreadsheet based on selected cell
  const topRowNum = sheet.getActiveCell().getRow(); // user can select what row to begin the table at, if they want some empty header rows
 
  // header row
  var headers = ["Last Name","First Name","Status","Last Contact","Last Dinner","Received Ward Plan?","Commitments made","Kept commitments?","References"];
  sheet.getRange(topRowNum,1,1,headers[0].length).setValues([headers]);
 
  // copy household array to the spreadsheet
  console.log("Entering data into the spreadsheet on row " + topRowNum);
  sheet.getRange(topRowNum+1,1,data.length,data[0].length).setValues(data);
 
  // apply formatting rules!!!
  var dataRange = sheet.getRange(topRowNum,1,data.length+1,headers[0].length);
  formatSheet(sheet,dataRange,topRowNum,headers);
}
 
 
// print an array of all data from the sheet, for diagnostic purposes
function dumpRawData(){
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  // get value of the top right cell, which holds the name of the json file
  var jsonName = sheet.getRange('A1').getValue();
 
  // fetch file from google drive and parse it
  var persons = parseJSONfromDrive(jsonName);
 
  // prepare a 2D array to fill with object data
  var data = [[]];
  for(var i=1;i<persons.length;i++) data.push([]); // javascript wants empty arrays to push to
  // headers are extracted in parallel, but separately
  var headers = [];
 
  // read each person's attribute for key 0, & then for key 1, etc.
  for(key in persons[0]){
    // collect key as header
    headers.push(key);
    // now loop through all person objects and read key X from each one
    for(var i=0;i<persons.length;i++){
      data[i].push(persons[i][key]);
    }
  }
}
 
// fetch file from google drive and parse it
function parseJSONfromDrive(fileName){
  console.log("Fetching data file from google drive...");
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
  console.log("Extracting member data from file...");
  // all members will have an empty cell under the key "missionName", but nonmembers will have the name of the mission
  // online people will have empty cells for "missionName" like members, but they have an ID number in "stewardCmisId"- all other dots have an empty cell here
  // delete all person rows with something in missionName
  var p=0;
  do{
    if((!persons[p]["missionName"]) && (!persons[p]["stewardCmisId"])){ // check that both keys are empty
      p++; // only increment if a person wasn't deleted
    }else{
      persons.splice(p,1); // delete 1 entry at location p- they're a non-member
    }
  }while(p<persons.length); // check at the end of each loop because the length is shrinking
 
  return persons;
}
 
// order an array of objects or an array of arrays by a specific attribute/index
function sortByAtt(arr, key, ascending, key2, ascending2){ // arr must contain objects or arrays
  // add default values
  ascending2 = ascending2 && true;
 
  // we want to be able to compare ascending or descending
  function dComp(val1,val2,isGreater){return (isGreater? val1>val2 : val2>val1);}
  function customSort(a,b){
    if(dComp(a[key],b[key],ascending)){ // a[key] is greater than b[key]
      return 1;
    }else if(a[key]===b[key]){ // fall to next comparator
      return ( dComp(a[key2],b[key2],ascending2)?1:-1 ); // compare secondary variables
    }else{ // a[key] is less than b[key]
      return -1;
    }
  }
  return arr.sort( customSort );
}
 
 
// compress people objects into households
function householdSquish(persons){
  console.log("Compressing individuals into households...");
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
function formatSheet(sheet,dataRange,topRowNum,headers){
  console.log("Applying formatting rules:");
  // alternating colors
  console.log(" '-> Alternating colors");
  if(dataRange.getBandings()==="") dataRange.applyRowBanding(SpreadsheetApp.BandingTheme.BLUE); // don't apply if bandings are already present!
  // center align & allow text wrapping in all cells
  console.log(" '-> Text centered & wrapping");
  dataRange.setHorizontalAlignment("center");
  dataRange.setVerticalAlignment("middle");
  dataRange.setWrap(true);
  // freeze last name column & header row
  console.log(" '-> Frozen header row & column");
  sheet.setFrozenColumns(1);
  sheet.setFrozenRows(topRowNum);
 
 // now we need a quick way to extract a single column from the given range
  function extractColumn(headerName){
    return sheet.getRange( // getRange(row,column,numRows) returns one column
      dataRange.getRow()+1, // getRow()+1 to skip headers
      headers.indexOf(headerName)+1, // +1 to convert array numbering [0,1..n-1] to sheet numbering [1,2..n]
      dataRange.getLastRow() // end range at bottom of data
    );
  }
 
 // Status dropdown
  console.log(" '-> Activity level dropdowns for status column");
  // these are the typical options for member statuses
  var statusList = ["unknown","active","returning","less active","inactive","hostile","moved","other"];
  // give the array of member statuses to a dropdown data structure, which is called data validation
  var statusRule = 
    SpreadsheetApp.newDataValidation()
    .requireValueInList(statusList)
    .build();
  // locate the "Status" column among the headers and apply the dropdown rules
  extractColumn("Status").setDataValidation(statusRule);

 // conditional formatting
  console.log(" '-> Automatic date coloring in Last Contact / Dinner columns");
  // the conditional formatting rules are stored in a global stack, which we will copy
  var ruleMasterList = sheet.getConditionalFormatRules(); // once we're done modifying this copy, we will use it to update the original

  // colorRule is a temporary variable which we will use to pass the constructed rule to our stack
  var colorRule = SpreadsheetApp.newConditionalFormatRule() // turn the date cells dark red w/ white text if they say NONE
    .whenTextEqualTo("NONE")
    .setBackground("#8A3535").setFontColor("#FFFFFF")
    .setRanges( [extractColumn("Last Contact"),extractColumn("Last Dinner")] ) // this applies to both date columns
    .build();
  ruleMasterList.push(colorRule);

  // now we're going to loop through numbers of days associated with colors
  var daysForLastContact = [ // one array for each column, since they're tuned slightly different
    {days:7,color:"#77C58C"},
    {days:15,color:"#B7E1CD"},
    {days:30,color:"#FFD966"},
    {days:60,color:"#F6B26B"}
  ];
  var daysForLastDinner = [ // days for last dinner are tuned bigger, because it's not as pressing as having a recent contact
    {days:7,color:"#77C58C"},
    //{days:15,color:"#B7E1CD"}, // no light green!
    {days:42,color:"#FFD966"},
    {days:90,color:"#F6B26B"}
  ];

  // now loop through both columns separately
  daysForLastContact.forEach(function(step){ // step through each mini-object in the array (Last Contact column)
    colorRule = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(
        "=" + // TODO
        extractColumn("Last Contact").getA1Notation().split(":")[0] + // this complicated thing here just pulls out the address of the top cell
        ">TODAY()-" +
        step.days // step is a copy of the current element of daysForLastContact[]
      )
      .setBackground(step.color) // step contains .days and .color
      .setRanges([ extractColumn("Last Contact") ])
      .build();
    ruleMasterList.push(colorRule);
  });
  daysForLastDinner.forEach(function(step){ // step through each mini-object in the array (Last Dinner column)
    colorRule = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(
        "=" + // TODO
        extractColumn("Last Dinner").getA1Notation().split(":")[0] + // this complicated thing here just pulls out the address of the top cell
        ">TODAY()-" +
        step.days // step is a copy of the current element of daysForLastContact[]
      )
      .setBackground(step.color) // step contains .days and .color
      .setRanges([ extractColumn("Last Dinner") ])
      .build();
    ruleMasterList.push(colorRule);
  });

  colorRule = SpreadsheetApp.newConditionalFormatRule() // turn the date cells gray if they aren't empty
    .whenCellNotEmpty()
    .setBackground("#B7B7B7")
    .setRanges( [extractColumn("Last Contact"),extractColumn("Last Dinner")] ) // this applies to both date columns
    .build();
  ruleMasterList.push(colorRule);

  sheet.setConditionalFormatRules(ruleMasterList);

};
 
/* TODO
--> allow user to select which formatting rules / headers to implement (advanced)
*/
