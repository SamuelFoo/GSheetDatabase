// TODO: Add multiline capability for order and spendings input
const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

// Order Entry Definitions
const ORDERS_ENTRY_SHEET = spreadsheet.getSheetByName("Order Entry");

// Orders Definitions
const ORDERS_SHEET = spreadsheet.getSheetByName("Orders");
const ORDERS_DATA_RANGE = "A1:M";
const ORDERS_ID_COLNAME = "Order No.";
const ORDERS_PRICE_COLNAME = "Value of Order";
const ORDERS_PAYMENT_REC_COLNAME = "Payment Received?";
const ORDERS_DATE_COLNAME = "Date Payment Received";
const ORDERS_MONEY_WITH_COLNAME = "Account Paid To";
const ORDERS_TYPE_COLNAME = "Nature of Invoice";
const ORDERS_DESC_COLNAME = "Description of Order";

// Cashflow Definitions
const CASHFLOW_SHEET = spreadsheet.getSheetByName("Cashflow");
const CASHFLOW_DATA_RANGE = "B2:E";
const CASHFLOW_ID_COLNAME = "Cashflow Id";
const CASHFLOW_START_ROW = 724; // Row no. at which Cashflow Id = 100001
const CASHFLOW_START_COL = 2; // Col no. of "CashflowId", index starts from 1

// Spendings Definitions
const SPENDINGS_SHEET = spreadsheet.getSheetByName("Spendings");
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

class Sheet {
  constructor(name, dataRange) {
    this.name = name;
    this.dataRange = dataRange;
    this.sheet = spreadsheet.getSheetByName(name);

    let colNamesRange = getFirstColFromA1Notation(dataRange);
    this.colNames = this.sheet.getRange(colNamesRange).getValues()[0];
  }
}

let SpendingsSheet = new Sheet("Spendings", SPENDINGS_DATA_RANGE);
let OrdersSheet = new Sheet("Orders", ORDERS_DATA_RANGE);
let InventorySheet = new Sheet(INVENTORY_SHEET_NAME, INVENTORY_DATA_RANGE);
let InventoryDateSheet = new Sheet(INVENTORYDATE_SHEET_NAME, INVENTORYDATE_DATA_RANGE);

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

// Extract orders which have payment received and money with Eltelierworks
function findPaymentsFromOrders() {
  const values = ORDERS_SHEET.getRange(ORDERS_DATA_RANGE).getValues();
  const orderIdsCol = values[0].indexOf(ORDERS_ID_COLNAME);
  const paymentRecCol = values[0].indexOf(ORDERS_PAYMENT_REC_COLNAME);
  const moneyWithCol = values[0].indexOf(ORDERS_MONEY_WITH_COLNAME);

  let filteredValues = values.filter(function(row) {  
    return (
      row[orderIdsCol] >= 100001 && 
      row[paymentRecCol] === "Yes"&&
      row[moneyWithCol] === "Eltelierworks"
    );
  });
  
  // Logger.log(filteredValues);
  return filteredValues;
}

// Extract spendings which "With"="Etelierworks"
function findOutflowsFromSpendings() {
  const values = SPENDINGS_SHEET.getRange(SPENDINGS_DATA_RANGE).getValues();
  const moneyWithCol = values[0].indexOf(SPENDINGS_REIMBURSED_COLNAME);  

  let filteredValues = values.filter(function(row) {  
    return (row[moneyWithCol] === "Yes");
  });
  
  // Logger.log(filteredValues);
  return filteredValues;
}

// Updates "Cashflow" with payments from orders and outflows from spendings
function updateCashflow() {
  // Helper functions
  function sortByDateFunc(row1, row2, dateCol) {
    let date1 = new Date(row1[dateCol]);
    let date2 = new Date(row2[dateCol]);
    if (date1.getTime() < date2.getTime()) {
      return -1;
    } 
    if (date1.getTime() > date2.getTime()) {
      return 1;
    }
    return 0;
  
  }  
  
  const payments = findPaymentsFromOrders();
  const outflows = findOutflowsFromSpendings();
  var cashFlowData = [];

  // Logger.log(payments);
  // Logger.log(outflows);

  // Orders (Payments) data processing
  let orderColNamesRange = getFirstColFromA1Notation(ORDERS_DATA_RANGE);
  let orderColNames = ORDERS_SHEET.getRange(orderColNamesRange).getValues()[0];
  
  const orderDateData = getColData(payments, orderColNames, ORDERS_DATE_COLNAME);
  const orderIdsData = getColData(payments, orderColNames, ORDERS_ID_COLNAME);
  const orderPriceData = getColData(payments, orderColNames, ORDERS_PRICE_COLNAME);

  for(var i = 0; i < payments.length; i++) {
    descString = "Payment for Order Id = " + String(orderIdsData[i])
    cashFlowData.push([orderDateData[i], descString, orderPriceData[i], ""]);
  }

  // Spendings (Outflows) data processing
  let spendingsColNamesRange = getFirstColFromA1Notation(SPENDINGS_DATA_RANGE);
  let spendingsColNames = SPENDINGS_SHEET.getRange(spendingsColNamesRange).getValues()[0];

  const spendingsDateData = getColData(outflows, spendingsColNames, SPENDINGS_DATE_COLNAME);
  const spendingsIdsData = getColData(outflows, spendingsColNames, SPENDINGS_ID_COLNAME);
  const spendingsCostData = getColData(outflows, spendingsColNames, SPENDINGS_COST_COLNAME);

  for(var i = 0; i < outflows.length; i++) {
    descString = "Expense Id = " + String(spendingsIdsData[i])
    cashFlowData.push([spendingsDateData[i], descString, "", spendingsCostData[i]]);
  }

  // Sort Cashflow by date
  cashFlowData.sort((row1, row2) => sortByDateFunc(row1, row2, 0));

  // Logger.log(cashFlowData);

  // Clear out existing data and then set it
  /* Clearing out existing data is important as it accounts for
   the case where the length of data decreases due to deletion) */
  let clearContentRangeStart = getColNotation(CASHFLOW_START_COL) + String(CASHFLOW_START_ROW);
  let clearContentRangeEnd = getColNotation(CASHFLOW_START_COL + cashFlowData[0].length-1);
  let clearContentRange = clearContentRangeStart + ":" + clearContentRangeEnd;
  // Logger.log(clearContentRange);

  CASHFLOW_SHEET.getRange(clearContentRange).clearContent();
  CASHFLOW_SHEET.getRange(
    CASHFLOW_START_ROW, CASHFLOW_START_COL, cashFlowData.length, cashFlowData[0].length
  ).setValues(cashFlowData);
}

// Updates "Inventory By Date"
function updateInventoryDate() {
  let inventoryData = []; // [Item No., Bought No., Sold No., Left, Bought Price, Date]
  let inventoryChangeLog = []; // [itemName, itemCost, numBought, numSold]

  // Process spendings
  let spendings = findOutflowsFromSpendings();
  let spendingsTypeCol = SpendingsSheet.colNames.indexOf(SPENDINGS_TYPE_COLNAME);
  spendings = spendings.filter(function(row) {  
    return (row[spendingsTypeCol] === "Cost of Goods");
  });
  let spendingsNameData = getColData(spendings, SpendingsSheet.colNames, SPENDINGS_DESC_COLNAME);
  for(var i = 0; i < spendings.length; i++) {
    let spendingName = spendingsNameData[i];
    let [numBought, itemName, itemCost] = spendingName.split(",");
    itemCost =  Number(itemCost.split("$")[1]);
    inventoryChangeLog.push([itemName, numBought, 0, itemCost]);
  }

  // Group by itemName
  inventoryChangeLog = inventoryChangeLog.reduce((x, y) => {
    (x[y[0]] = x[y[0]] || []).push(y);
    return x;
  }, {});

  // Group by itemName then by itemCost
  for (let name in inventoryChangeLog) {
    inventoryChangeLog[name] = inventoryChangeLog[name].reduce((x, y) => {
      (x[y[3]] = x[y[3]] || []).push(y);
      return x;
    }, {});
    
    // Logger.log(inventoryChangeLog[name]);

    for (let cost in inventoryChangeLog[name]) {
      let numBought = 0;
      let numSold = 0;

      for (let i = 0; i < inventoryChangeLog[name][cost].length; i++) {
        let row = inventoryChangeLog[name][cost][i];
        numBought += Number(row[1]);
        numSold += Number(row[2]);
      }

      inventoryData.push([name, numBought, numSold, numBought-numSold, cost]);
    }
  }

  function sortByCostFunc(row1, row2, costCol) {
    let cost1 = row1[costCol];
    let cost2 = row2[costCol];
    if (cost1 < cost2) {
      return -1;
    } 
    if (cost1 > cost2) {
      return 1;
    }
    return 0;
  }  
  inventoryData = inventoryData.sort(sortByCostFunc);

  // Process orders
  let inventoryOrderLog = [];
  let orders = findPaymentsFromOrders();
  let ordersTypeCol = OrdersSheet.colNames.indexOf(ORDERS_TYPE_COLNAME);
  orders = orders.filter(function(row) {  
    return (row[ordersTypeCol] === "FDM");
  });
  let ordersNameData = getColData(orders, OrdersSheet.colNames, ORDERS_DESC_COLNAME);
  for(var i = 0; i < orders.length; i++) {
    let orderName = ordersNameData[i];
    let [numSold, itemName] = orderName.split(",");
    inventoryOrderLog.push([itemName, numSold]);
  }

  // Group by itemName
  inventoryOrderLog = inventoryOrderLog.reduce((x, y) => {
    (x[y[0]] = x[y[0]] || []).push(y);
    return x;
  }, {});

  for (name in inventoryOrderLog) {
    let totalNumSold = 0;
    for (let i = 0; i < inventoryOrderLog[name].length; i++) {
      Logger.log(inventoryOrderLog[name][i][1]);
      totalNumSold += Number(inventoryOrderLog[name][i][1]);
    }

    // Logger.log(name);
    // Logger.log(totalNumSold);

    for (let i = 0; i < inventoryData.length; i++) {
      if (totalNumSold == 0) break;
      else {
        let [itemName, numBought, numSold, numLeft, cost] = inventoryData[i];
        if (itemName != name) continue;
        let addNumSold = Math.min(numLeft, totalNumSold)
        let newNumSold = addNumSold + numSold;
        totalNumSold -= addNumSold;
        inventoryData[i] = [itemName, numBought, newNumSold, numBought-newNumSold, cost];
      }
    }
  }

  // Logger.log(inventoryData);

  // Clear out existing data and then set it
  /* Clearing out existing data is important as it accounts for
   the case where the length of data decreases due to deletion) */
  let clearContentRangeStart = getColNotation(INVENTORY_START_COL) + String(INVENTORY_START_ROW);
  let clearContentRangeEnd = getColNotation(INVENTORY_START_COL + inventoryData[0].length-1);
  let clearContentRange = clearContentRangeStart + ":" + clearContentRangeEnd;
  // Logger.log(clearContentRange);

  InventorySheet.sheet.getRange(clearContentRange).clearContent();
  InventorySheet.sheet.getRange(
    INVENTORY_START_ROW, INVENTORY_START_COL, inventoryData.length, inventoryData[0].length
  ).setValues(inventoryData);
}

// Updates "Inventory"
function updateInventory() {
  updateInventoryDate();
}

// Updates "Profits"
function updateProfits() {
  let orders = findPaymentsFromOrders();
  let spendings = findOutflowsFromSpendings();
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
