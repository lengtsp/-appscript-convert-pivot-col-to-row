function transformData(inputSheetName, outputSheetName, baseColumnsCount, typeCount, totalTypes) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var inputSheet = ss.getSheetByName(inputSheetName);
  var outputSheet = ss.getSheetByName(outputSheetName);
  
  // If the output sheet doesn't exist, create it
  if (!outputSheet) {
    outputSheet = ss.insertSheet(outputSheetName);
  }
  
  // Read the data from the input sheet
  var data = inputSheet.getDataRange().getValues();
  
  // Separate the headers and the data
  var headers = data[0];
  var rows = data.slice(1);
  
  // Prepare the transformed data
  var transformedData = [];
  
  // Define the new headers
  var newHeaders = headers.slice(0, baseColumnsCount).concat(['type']);
  for (var i = 0; i < typeCount; i++) {
    newHeaders.push('type1_' + (i + 1));
  }
  
  // For each row, split the data into groups of columns for each type and reformat
  rows.forEach(function(row) {
    var baseData = row.slice(0, baseColumnsCount);
    
    // Process each type
    for (var i = 0; i < totalTypes; i++) {
      var type = 'type' + (i + 1);
      var start = baseColumnsCount + (i * typeCount);
      var typeData = row.slice(start, start + typeCount);
      var newRow = baseData.concat([type]).concat(typeData);
      transformedData.push(newRow);
    }
  });
  
  // Clear any existing data in the output sheet and write the transformed data
  outputSheet.clear();
  outputSheet.getRange(1, 1, 1, newHeaders.length).setValues([newHeaders]);
  outputSheet.getRange(2, 1, transformedData.length, newHeaders.length).setValues(transformedData);
}

function executeTransform() {
  transformData("Sheet1", "Sheet2", 7, 3, 4);
}




function onOpen() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var menuItems = [
    {name: 'Run executeTransform', functionName: 'executeTransform'}
  ];
  spreadsheet.addMenu('>> Custom Menu <<', menuItems);
}



