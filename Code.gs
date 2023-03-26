// TODO: Add multiline capability for order and spendings input
const Spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

// Orders Definitions
const ORDERS_SHEET_NAME = "Orders";
const ORDERS_DATA_RANGE = "A1:M";
const ORDERS_ID_COLNAME = "Order No.";
const ORDERS_PRICE_COLNAME = "Value of Order";
const ORDERS_PAYMENT_REC_COLNAME = "Payment Received?";
const ORDERS_DATE_COLNAME = "Date Payment Received";
const ORDERS_MONEY_WITH_COLNAME = "Account Paid To";
const ORDERS_TYPE_COLNAME = "Nature of Invoice";
const ORDERS_DESC_COLNAME = "Description of Order";

// Cashflow Definitions
const CASHFLOW_SHEET_NAME = "Cashflow";
const CASHFLOW_DATA_RANGE = "B2:E";
const CASHFLOW_ID_COLNAME = "Cashflow Id";
const CASHFLOW_START_ROW = 724; // Row no. at which Cashflow Id = 100001
const CASHFLOW_START_COL = 2; // Col no. of "CashflowId", index starts from 1

// Spendings Definitions
const SPENDINGS_SHEET_NAME = "Spendings";
const SPENDINGS_DATA_RANGE = "A1:F";
const SPENDINGS_ID_COLNAME = "SN";
const SPENDINGS_DATE_COLNAME = "Date";
const SPENDINGS_DESC_COLNAME = "Name Of Spending";
const SPENDINGS_TYPE_COLNAME = "Type of Spending";
const SPENDINGS_COST_COLNAME = "Cost Incurred";
const SPENDINGS_REIMBURSED_COLNAME = "Reimbursed";

// Inventory Definitions
const INVENTORY_SHEET_NAME = "Inventory";
const INVENTORY_DATA_RANGE = "A1:E";
const INVENTORY_START_ROW = 2; // Row no. at which Cashflow Id = 100001
const INVENTORY_START_COL = 1; // Col no. of "CashflowId", index starts from 1

// Inventory By Date Definitions
const INVENTORYDATE_SHEET_NAME = "Inventory By Date";
const INVENTORYDATE_DATA_RANGE = "A1:F";
const INVENTORYDATE_START_ROW = 2; // Row no. at which Cashflow Id = 100001
const INVENTORYDATE_START_COL = 1; // Col no. of "CashflowId", index starts from 1

// Profits Definitions
const PROFITS_SHEET_NAME = "Profits";
const PROFITS_DATA_RANGE = "A1:G";
const PROFITS_START_COL = 1;
const PROFITS_START_ROW = 2;

// Imitates Python's Pandas DataFrame
class DataFrame {
  constructor(colNames, data) {
    this.colNames = colNames;
    this.data = data;
  }

  loc(colNames) {
    let newData = this.data.map(        
      row => colNames.map(
        colName => row[this.colNames.indexOf(colName)]
      )
    );
    return new DataFrame(this.colNames, newData);
  }

  T() {
    return new DataFrame(this.colNames, this.data[0].map((_, colIndex) => this.data.map(row => row[colIndex])));
  }

  filterDF(filterConds) {
    let colNames = this.colNames;
    return new DataFrame(
      this.colNames,
      this.data.filter(
        function(row) {  
          let bool = true;
          for (let i = 0; i < filterConds.length; i++) {
            let [filterColName, filterFunc, filterValue] = filterConds[i];
            let filterColIdx = colNames.indexOf(filterColName);
            bool = bool && filterFunc(row[filterColIdx], filterValue);
          }
          return bool;
        }
      )
    );
  }
}

class Sheet {
  constructor(name, dataRange) {
    let sheet = Spreadsheet.getSheetByName(name);
    let sheetValues = sheet.getRange(dataRange).getValues();
    this.df = new DataFrame(sheetValues[0], sheetValues.slice(1));
  }
}

let SpendingsDF = new Sheet(SPENDINGS_SHEET_NAME, SPENDINGS_DATA_RANGE).df;
let OrdersDF = new Sheet(ORDERS_SHEET_NAME, ORDERS_DATA_RANGE).df;
let InventoryDF = new Sheet(INVENTORY_SHEET_NAME, INVENTORY_DATA_RANGE).df;
let InventoryDateDF = new Sheet(INVENTORYDATE_SHEET_NAME, INVENTORYDATE_DATA_RANGE).df;

const arrayColumn = (arr, n) => arr.map(x => x[n]);
const getColData = (data, colNames, colName) => arrayColumn(data, colNames.indexOf(colName));

const formEntryDict = [
  {
    entryName: "Order No.",
    cellId: "D3",
  },
  {
    entryName: "Nature of Invoice",
    cellId: "D5",
  },
  {
    entryName: "Carousell/Facebook/Friend",
    cellId: "D7",
  },
  {
    entryName: "Order Description",
    cellId: "D9",
  },
  {
    entryName: "Selling Price",
    cellId: "D11",
  },
  {
    entryName: "Payment Received?",
    cellId: "D13",
  },
  {
    entryName: "Money With",
    cellId: "D15",
  },
  {
    entryName: "Remarks",
    cellId: "D17",
  }
];

function getFirstColFromA1Notation(a1Notation) {
  let [startCol, endCol] = a1Notation.replace(/[^A-Z:]+/gi, '').split(":");
  let [startRow, _] = a1Notation.replace(/[^0-9:]+/gi, '').split(":");  
  return startCol + startRow + ":" + endCol + startRow;
}

// Return A if col=1, B if col=2...
function getColNotation(col) {
  let character = "A".charCodeAt(0);
  return String.fromCharCode(character + col - 1);
}

// Helper function for arr.sort
function sortByDateFunc(row1, row2, dateCol) {
  let date1 = new Date(row1[dateCol]);
  let date2 = new Date (row2[dateCol]);
  if (date1.getTime() < date2.getTime()) {
    return -1;
  } 
  if (date1.getTime() > date2.getTime()) {
    return 1;
  }
  return 0;
}

// Clear out existing data and then set it
function clearAndSet(sheetName, startCol, startRow, data) {
  /* Clearing out existing data is important as it accounts for
   the case where the length of data decreases due to deletion) */
  let clearContentRangeStart = getColNotation(startCol) + String(startRow);
  let clearContentRangeEnd = getColNotation(startCol + data[0].length-1);
  let clearContentRange = clearContentRangeStart + ":" + clearContentRangeEnd;
  // Logger.log(clearContentRange);

  let sheet = Spreadsheet.getSheetByName(sheetName)
  sheet.getRange(clearContentRange).clearContent();
  sheet.getRange(
    startRow, startCol, data.length, data[0].length
  ).setValues(data);
}

// Extract orders which have payment received and money with Eltelierworks
function getReceivedPayments() {
  const filterConds = [
    [ORDERS_ID_COLNAME, (x,y) => x>=y, 100001],
    [ORDERS_PAYMENT_REC_COLNAME, (x,y) => x===y, "Yes"],
    [ORDERS_MONEY_WITH_COLNAME, (x,y) => x===y, "Eltelierworks"]
  ];

  let filteredValues = OrdersDF.filterDF(filterConds);
  // Logger.log(filteredValues);
  return filteredValues;
}

// Extract spendings which "Reimbursed"="Yes"
function getReimbursedSpendings() {
  const filterConds = [
    [SPENDINGS_REIMBURSED_COLNAME, (x,y) => x===y, "Yes"]
  ];

  let filteredValues = SpendingsDF.filterDF(filterConds);
  // Logger.log(filteredValues);
  return filteredValues;
}

// Updates "Cashflow" with payments from orders and outflows from spendings
function updateCashflow() {  
  const paymentsDF = getReceivedPayments();
  const spendingsDF = getReimbursedSpendings();
  let cashFlowData = [];

  // Logger.log(payments);
  // Logger.log(outflows);

  // Orders (payments) data processing
  let [ordersDateData, ordersIdsData, ordersPriceData] = paymentsDF.loc([ORDERS_DATE_COLNAME, ORDERS_ID_COLNAME, ORDERS_PRICE_COLNAME]).T().data;

  for(let i = 0; i < ordersDateData.length; i++) {
    descString = "Payment for Order Id = " + String(ordersIdsData[i])
    cashFlowData.push([ordersDateData[i], descString, ordersPriceData[i], ""]);
  }

  // Spendings data processing
  let [spendingsDateData, spendingsIdsData, spendingsCostData] = spendingsDF.loc([SPENDINGS_DATE_COLNAME, SPENDINGS_ID_COLNAME, SPENDINGS_COST_COLNAME]).T().data;
  for(let i = 0; i < spendingsDateData.length; i++) {
    let descString = "Expense Id = " + String(spendingsIdsData[i]);

    cashFlowData.push([spendingsDateData[i], descString, "", spendingsCostData[i]]);
  }

  // Sort Cashflow by date
  cashFlowData.sort((row1, row2) => sortByDateFunc(row1, row2, 0));

  // Logger.log(cashFlowData);

  clearAndSet(CASHFLOW_SHEET_NAME, CASHFLOW_START_COL, CASHFLOW_START_ROW, cashFlowData);
}

// Updates "Inventory By Date"
function updateInventoryByDate() {
  let inventoryData = []; // [Item No., Bought No., Sold No., Left, Bought Price, Date]
  let inventorySpendingsLog = []; // [itemName, itemCost, numBought, numSold, Date]
  let inventoryOrdersLog = []; // [itemName, itemCost, numBought, numSold, Date]

  // Process spendings
  let spendingsDF = getReimbursedSpendings();
  const spendingsConds = [
    [SPENDINGS_TYPE_COLNAME, (x,y) => x===y, "Cost of Goods"]
  ];
  spendingsDF = spendingsDF.filterDF(spendingsConds);
  let [spendingsNameData, spendingsDateData] = spendingsDF.loc([SPENDINGS_DESC_COLNAME, SPENDINGS_DATE_COLNAME]).T().data;
  for(let i = 0; i < spendingsNameData.length; i++) {
    let [spendingName, spendingDate] = [spendingsNameData[i], new Date(spendingsDateData[i])];
    let [numBought, itemName, itemCost] = spendingName.split(",");
    itemCost = Number(itemCost.split("$")[1]);
    
    spendingDate = [spendingDate.getMonth()+1, spendingDate.getDate(), spendingDate.getFullYear()].join("/");
    // Logger.log(spendingDate);
    inventorySpendingsLog.push([itemName, numBought, 0, itemCost, spendingDate]);
  }

  // Process orders
  let ordersDF = getReceivedPayments();
  const ordersConds = [
    [ORDERS_TYPE_COLNAME, (x,y) => x===y, "FDM"]
  ];
  ordersDF = ordersDF.filterDF(ordersConds);
  let [ordersNameData, ordersDateData] = ordersDF.loc([ORDERS_DESC_COLNAME, ORDERS_DATE_COLNAME]).T().data
  for(let i = 0; i < ordersNameData.length; i++) {
    let [orderName, orderDate] = [ordersNameData[i], ordersDateData[i]];
    let [numSold, itemName] = orderName.split(",");
    inventoryOrdersLog.push([itemName, 0, numSold, "-", orderDate]);
  }

  // Aggregate spendings with the same date -> inventoryData
  let dateCol = 4;
  inventorySpendingsLog = inventorySpendingsLog.reduce((x, y) => {
    (x[y[0]] = x[y[0]] || []).push(y);
    return x;
  }, {}); // Group by itemName
  for (let name in inventorySpendingsLog) {
    inventorySpendingsLog[name] = inventorySpendingsLog[name].reduce((x, y) => {
      (x[y[dateCol]] = x[y[dateCol]] || []).push(y);
      return x;
    }, {}); // For each itemName group, group by date
    
    // Logger.log(inventorySpendingsLog[name]);

    for (let date in inventorySpendingsLog[name]) {
      let numBought = 0;
      let numSold = 0;

      for (let i = 0; i < inventorySpendingsLog[name][date].length; i++) {
        let row = inventorySpendingsLog[name][date][i];
        numBought += Number(row[1]);
        numSold += Number(row[2]);
      }
      
      let numLeft = numBought-numSold;
      // Asssumption: items of the same name bought on the same date have the same cost.
      let cost = inventorySpendingsLog[name][date][0][3]; 

      inventoryData.push([name, numBought, numSold, numLeft, cost, date]);
    }
  }

  // Sort from earliest to latest, so that later on we can remove earlier dated item first.
  inventoryData = inventoryData.sort((row1, row2) => sortByDateFunc(row1, row2, 6));

  // Remove items in orders from inventory, earliest dated first.
  inventoryOrdersLog = inventoryOrdersLog.reduce((x, y) => {
    (x[y[0]] = x[y[0]] || []).push(y);
    return x;
  }, {}); // Group by itemName
  for (name in inventoryOrdersLog) {
    let totalNumSold = 0;
    for (let i = 0; i < inventoryOrdersLog[name].length; i++) {
      // Logger.log(inventoryOrdersLog[name][i][2]);
      totalNumSold += Number(inventoryOrdersLog[name][i][2]);
    }

    // Logger.log(name);
    // Logger.log(totalNumSold);

    for (let i = 0; i < inventoryData.length; i++) {
      if (totalNumSold == 0) break;
      else {
        let [itemName, numBought, numSold, numLeft, cost, date] = inventoryData[i];
        if (itemName != name) continue;
        let addNumSold = Math.min(numLeft, totalNumSold)
        let newNumSold = addNumSold + numSold;
        totalNumSold -= addNumSold;
        inventoryData[i] = [itemName, numBought, newNumSold, numBought-newNumSold, cost, date];
      }
    }
  }

  clearAndSet(INVENTORYDATE_SHEET_NAME, INVENTORYDATE_START_COL, INVENTORYDATE_START_ROW, inventoryData);

  return inventoryData;
}

// Updates "Inventory"
function updateInventory() {
  let inventoryByDateData = updateInventoryByDate(); // [Item No., Bought No., Sold No., Left, Bought Price]
  let inventoryData = [];

  // Aggregates items with same name and same cost.
  inventoryByDateData = inventoryByDateData.reduce((x, y) => {
    (x[y[0]] = x[y[0]] || []).push(y);
    return x;
  }, {}); // Group by itemName
  for (let name in inventoryByDateData) {
    inventoryByDateData[name] = inventoryByDateData[name].reduce((x, y) => {
      (x[y[4]] = x[y[4]] || []).push(y);
      return x;
    }, {}); // Group by itemName then by cost
    
    // Logger.log(inventoryByDateData[name]);

    for (let cost in inventoryByDateData[name]) {
      let numBought = 0;
      let numSold = 0;

      for (let i = 0; i < inventoryByDateData[name][cost].length; i++) {
        let row = inventoryByDateData[name][cost][i];
        numBought += Number(row[1]);
        numSold += Number(row[2]);
      }
      
      let numLeft = numBought-numSold;
      // Asssumption: items of the same name bought on the same date have the same cost.

      inventoryData.push([name, numBought, numSold, numLeft, cost]);
    }
  }

  clearAndSet(INVENTORY_SHEET_NAME, INVENTORY_START_COL, INVENTORY_START_ROW, inventoryData);
}

// Updates "Profits"
function updateProfits() {
  let profitsData = [];

  // Process orders
  let ordersDF = getReceivedPayments();
  const ordersConds = [
    [ORDERS_TYPE_COLNAME, (x,y) => x===y, "FDM"]
  ];
  ordersDF = ordersDF.filterDF(ordersConds);
  let [orderIds, orderNames, orderPrices, orderDates] = ordersDF.loc([ORDERS_ID_COLNAME, ORDERS_DESC_COLNAME, ORDERS_PRICE_COLNAME, ORDERS_DATE_COLNAME]).T().data
  for(let i = 0; i < orderNames.length; i++) {
    let [orderName, orderDate] = [orderNames[i], orderDates[i]];
    let [itemCount, itemName] = orderName.split(",");
    profitsData.push([orderIds[i], itemCount, itemName, -1, orderPrices[i], -1, -1, orderDate]);
  }
  
  // Sort orders (stored as profits) by date
  profitsData.sort((row1, row2) => sortByDateFunc(row1, row2, 6));

  // Update profits with orders
  let inventoryByDateData = updateInventoryByDate(); // [itemName, numBought, numSold, numLeft, cost, date]
  /* Create a copy of inventoryByDateData with only what we need and for keeping track of items left after
  fulfilling orders. Note that we take numBought as count because we want the number before fulfilling any
  orders. */ // [itemName, count, cost, date]
  let runningInventory = (new DataFrame([0,1,2,3,4,5], inventoryByDateData)).loc([0,1,4,5]).data; 

  for (let i = 0; i < profitsData.length; i++) {
    let [orderId, itemCount, itemName, _0, soldPrice, _1, _2, soldDate] = profitsData[i];    
    let itemCountLeft = itemCount;
    let boughtPrice = 0;
    let boughtDates = "";

    for (let j = 0; j < runningInventory.length; j++) {
      let [invName, invCount, invCost, invDate] = runningInventory[j];
      if (invName != itemName || invCount === 0) continue;

      itemCountLeft = Math.max(itemCount - invCount, 0);
      let countTakenFromThisRow = (itemCount - itemCountLeft);
      
      boughtPrice += countTakenFromThisRow * invCost;
      boughtDates += String(invDate);
      
      invCount -= countTakenFromThisRow;
      runningInventory[j] = [invName, invCount, invCost, invDate];

      if (itemCountLeft === 0) break;
      else boughtDates += ", ";
    }

    profitsData[i] = [orderId, itemCount, itemName, boughtPrice, soldPrice, soldPrice-boughtPrice, boughtDates, soldDate];
  }

  clearAndSet(PROFITS_SHEET_NAME, PROFITS_START_COL, PROFITS_START_ROW, profitsData);
}

// Clears Order Entry Form
function clearForm() {
  let ui = SpreadsheetApp.getUi();
  let response = ui.alert("Reset Confirmation", 'Do you want to reset this form?',ui.ButtonSet.YES_NO);
 
  if (response == ui.Button.YES) 
  {
    formEntryDict.forEach(function(item, index) {
      ORDERS_ENTRY_SHEET.getRange(item.cellId).clear();
      ORDERS_ENTRY_SHEET.getRange(item.cellId).setBackground("#FFFFFF");
    });
  }

  return true ;
}
