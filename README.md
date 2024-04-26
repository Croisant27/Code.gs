# Code.gs
//First '.gs' file - formatting data, and generating line chart

//setLimits function stores each generator's limit in each round
function setLimits(round) {
  var generatorLimits = {};

  switch (round) {
    case "r1":
      generatorLimits = { //window 5 hours
        "Coal": 5400*5,
        "Gas A": 11000*5,
        "Gas B": 11000*5,
        "Gas C": 11000*5,
        "Hydro": 2000*5,
        "Pumped Storage": 2800*5,
        "Onshore Wind" : 13000*5*0.2,
        "Offshore Wind" : 11000*5*0.2,
        "Nuclear" : 8000*5,
        "Solar PV" : 14000*5*0.1,
      };
      break;
    case "r2":
      generatorLimits = { //window 5 hours
        "Coal": 5400*5,
        "Gas A": 11000*5,
        "Gas B": 11000*5,
        "Gas C": 11000*5,
        "Hydro": 2000*5,
        "Pumped Storage": 2800*5,
        "Onshore Wind" : 13000*5*0.5,
        "Offshore Wind" : 11000*5*0.5,
        "Nuclear" : 8000*5,
        "Solar PV" : 14000*5*0.9,
      };
      break;
    case "r3":
      generatorLimits = { //window 4 hours
        "Coal": 5400*4,
        "Gas A": 11000*4,
        "Gas B": 11000*4,
        "Gas C": 11000*4,
        "Hydro": 2000*4,
        "Pumped Storage": 2800*4,
        "Onshore Wind" : 13000*4*0.1,
        "Offshore Wind" : 11000*4*0.1,
        "Nuclear" : 8000*4,
        "Solar PV" : 14000*4*0.4,
      };
      break;
    case "r4":
      generatorLimits = { //window 6 hours
        "Coal": 5400*6,
        "Gas A": 11000*6,
        "Gas B": 11000*6,
        "Gas C": 11000*6,
        "Hydro": 2000*6,
        "Pumped Storage": 2800*6,
        "Onshore Wind" : 13000*6*0.8,
        "Offshore Wind" : 11000*6*0.8,
        "Nuclear" : 8000*6,
      };
      break;
    case "r5": //window 4 hours
      generatorLimits = {
        "Coal": 5400*4,
        "Gas A": 11000*4,
        "Gas B": 11000*4,
        "Gas C": 11000*4,
        "Hydro": 2000*4,
        "Pumped Storage": 2800*4,
        "Onshore Wind" : 13000*4*0.6,
        "Offshore Wind" : 11000*4*0.6,
        "Nuclear" : 8000*4,
        "Solar PV" : 14000*4*0,
      };
      break;
    default:
      // Handle the case when the round is not recognized
      break;
  }

  return generatorLimits;
}

//determine the generation limit for each energy source
function limitCapacity(data) {
  var dataSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Form responses 1");
  var round = dataSheet.getRange(2, 9).getValue(); // Assuming round information is in column I (9th column)

  var generatorLimit = setLimits(round);

  for (var i = 1; i < data.length; i++) {
    var generatorType = data[i][1];
    var groupSellingPower = data[i][2] + data[i][5];

    Logger.log("Processing row: " + i);
    Logger.log("Generator Type: " + generatorType);
    Logger.log("Group Selling Power: " + groupSellingPower);
    
    if (generatorType in generatorLimit) {
      if (groupSellingPower > generatorLimit[generatorType]) {
        Logger.log("Exceeds limit! Applying changes.");

        dataSheet.getRange(i + 1, 2).setFontColor("red");

        // Calculate the remaining compound value
        var remainingCompound = generatorLimit[generatorType] - data[i][2];

        // Update the compound value (Column F)
        data[i][5] = Math.max(0, remainingCompound); // Ensure compound value is not negative

        // Update the first bid value (Column C)
        data[i][2] += Math.min(0, remainingCompound); // Increment first bid if compound value is negative

        // Set the values in the spreadsheet
        dataSheet.getRange(i + 1, 6).setValue(data[i][5]);
        dataSheet.getRange(i + 1, 3).setValue(data[i][2]);

        Logger.log("Updated values - First Bid: " + data[i][2] + ", Compound: " + data[i][5]);
      }
    } else {
      Logger.log("Generator type not found in generatorLimit.");
      // Handle the case when generatorType is not in generatorLimit
    }
  }
}

function updateTable()
{
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dataSheet = ss.getSheetByName("Form responses 1");

  // Check if a sheet named "strikePrice" already exists. If yes, delete it.
  var oldSheet = ss.getSheetByName("strikePrice");
  if (oldSheet != null) {
    ss.deleteSheet(oldSheet);
  }

  // Create a new sheet named "strikePrice" and set it as the targetSheet.
  var targetSheet = ss.insertSheet("strikePrice");

  var data = dataSheet.getDataRange().getValues();

  var outputDataFirst = [];
  var totalSupplyFirst = 0;
  var priceDataFirst = {};

  var outputDataCompound = [];
  var totalSupplyCompound = 0;
  var priceDataCompound = {};

  var generatorLimit = {}

  limitCapacity(data, generatorLimit);

  for(var i = 1; i < data.length; i++)
  {
    var priceFirst = data[i][3];
    var studentBoughtFirst = data[i][2];
    var priceCompound = data[i][6];
    var studentBoughtCompound = data[i][5];

    priceDataFirst[priceFirst] = (priceDataFirst[priceFirst] || 0) + studentBoughtFirst;
    priceDataCompound[priceCompound] = (priceDataCompound[priceCompound] || 0) + studentBoughtCompound;
  }

  var sortedPricesFirst = Object.keys(priceDataFirst).sort(function(a,b)
  {
    return a-b;
  });

  var sortedPricesCompound = Object.keys(priceDataCompound).sort(function(a,b)
  {
    return a-b;
  });

  for(var i = 0; i < sortedPricesFirst.length; i++)
  {
    var priceFirst = sortedPricesFirst[i];
    totalSupplyFirst += priceDataFirst[priceFirst];
    outputDataFirst.push([priceFirst, priceDataFirst[priceFirst], totalSupplyFirst]);
  }

  for(var i = 0; i < sortedPricesCompound.length; i++)
  {
    var priceCompound = sortedPricesCompound[i];
    totalSupplyCompound += priceDataCompound[priceCompound];
    outputDataCompound.push([priceCompound, priceDataCompound[priceCompound], totalSupplyCompound]);
  }

  var combinedData = outputDataFirst.concat(outputDataCompound);

  //create map to combined the same data
  var combinedMap = {};

  for(var i = 0; i < combinedData.length;i++)
  {
    var price = combinedData[i][0];
    var powerSold = combinedData[i][1];

    if(combinedMap[price]){
      combinedMap[price] += powerSold;
    }else{
      combinedMap[price] = powerSold;
    }
  }

  var combinedTable = [];
  var totalStudentBuy = 0;
  
  for (var price in combinedMap) {
    totalStudentBuy += combinedMap[price];
    combinedTable.push([price, combinedMap[price], totalStudentBuy]);
  }

  combinedTable.sort(function(a, b) {
    return a[0] - b[0];
  });

  targetSheet.getRange(1,1,1,3).setValues([["£/MWh Combined", "Consumers (MWh)", "TOTAL (MWh)"]]);
  targetSheet.getRange(2, 1, combinedTable.length, 3).setValues(combinedTable);

  runAllFunctions();
}


// Generate table2 based on the provided logic
function generateTable2() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("strikePrice");

  // table1 is in columns A, B, C, create table2 in columns D and E
  var table1Data = sheet.getRange("A2:C" + sheet.getLastRow()).getValues();

  // Initialize variables for table2
  var table2Data = [["Energy (GWh)", "Price (£/MWh)"]];
  var startingPoint = 0;
  var stoppingPoint = 0;

  // Loop through table1 data
  for (var i = 0; i < table1Data.length; i++) {
    // Calculate starting and stopping points
    startingPoint = stoppingPoint;
    stoppingPoint = (table1Data[i][2] / 1000).toFixed(1);

    // Add data to table2
    table2Data.push([startingPoint, table1Data[i][0]]);
    table2Data.push([stoppingPoint, table1Data[i][0]]);

    // If it's the last row of table1, exit the loop
    if (i === table1Data.length - 1) {
      table2Data.push([stoppingPoint, 180]);
    }
  }

  // Write table2Data to columns D and E
  sheet.getRange(1, 4, table2Data.length, 2).setValues(table2Data);
}

function moveTableData() {
  // Define the data for each table
  var tables = {
    'r1': [
      [0, 150],
      [115, 150],
      [115, 120],
      [125, 120],
      [125, 50],
      [132.5, 50],
      [132.5, 20],
      [132.5, 20],
      [132.5, 0],
      [133, 0],
      [500, 0]
    ],
    'r2': [
      [0, 150],
      [125, 150],
      [125, 120],
      [195, 120],
      [195, 50],
      [195, 50],
      [195, 20],
      [220, 20],
      [220, 0],
      [220, 0],
      [500, 0]
    ],
    'r3': [
      [0, 150],
      [125, 150],
      [125, 120],
      [175, 120],
      [175, 50],
      [175, 50],
      [175, 20],
      [205, 20],
      [205, 0],
      [205, 0],
      [500, 0]
    ],
    'r4': [
      [0, 150],
      [145, 150],
      [145, 120],
      [155, 120],
      [155, 50],
      [155, 50],
      [155, 20],
      [180, 20],
      [180, 0],
      [180, 0],
      [500, 0]
    ],
    'r5': [
      [0, 150],
      [40, 150],
      [40, 120],
      [45, 120],
      [45, 50],
      [45, 50],
      [45, 20],
      [75, 20],
      [75, 0],
      [75, 0],
      [500, 0]
    ]
  };

  // Get the active spreadsheet
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  
  // Get the sheet where the movement will occur
  var destinationSheet = spreadsheet.getSheetByName('strikePrice');
  
  // Get the table name from cell I2 in "Form responses 1"
  var sourceSheet = spreadsheet.getSheetByName('Form responses 1');
  var tableName = sourceSheet.getRange('I2').getValue();
  
  // Get the data for the chosen table
  var tableData = tables[tableName];
  
  if (tableData) { 
    // Move the first column to D27 and the second column to F27 in the destination sheet
    destinationSheet.getRange('D27').offset(0, 0, tableData.length, 1).setValues(tableData.map(row => [row[0]]));
    destinationSheet.getRange('F27').offset(0, 0, tableData.length, 1).setValues(tableData.map(row => [row[1]]));
  } else {
    // Handle the case where an invalid table name is entered in I2
    Logger.log('Invalid table name');
  }
}

function createLineChart() {
  // Get the active sheet
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("strikePrice");
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sourceSheet = spreadsheet.getSheetByName('Form responses 1');
  var tableName = sourceSheet.getRange('I2').getValue();

  // Create a new chart
  var chart = sheet.newChart()
    .asLineChart()
    .addRange(sheet.getRange("D1:E" + sheet.getLastRow())) // First series from D and E columns
    .setOption('series', {
      0: { // First series customization
        labelInLegend: 'Aggregate Supply'
      },
      1: { // Second series customization
        labelInLegend: 'Aggregate Demand'
      }
    })
    .addRange(sheet.getRange("F1:G" + sheet.getLastRow())) // Second series from F and G columns
    .setPosition(5, 5, 0, 0)
    .setTitle('Aggregate Supply VS Aggregate Demand ' + tableName.toUpperCase())
    .setXAxisTitle('Energy (GWh)')
    .setYAxisTitle('Price (£/MWh)')
    .setOption('vAxes', {
      0: {
        viewWindow: {
          max: 180
        },
        minorGridlines: {
          count: 4
        }
      }
    })
    .build();

  // Insert the chart into the sheet
  sheet.insertChart(chart);
}

function runAllFunctions() {
  generateTable2();
  moveTableData();
  createLineChart();
}
