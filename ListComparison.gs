function processOrderingList() {
  let orderingListSpreadsheet = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/19rV2sw-JVJaYHYmG0xC5FrpqE9uGjKYKOdBOiwvBxI4/edit?gid=389276120#gid=389276120");
  let orderingSheet = orderingListSpreadsheet.getSheetByName("Ordering List");
  let range = orderingSheet.getRange(2, 1, orderingSheet.getLastRow() - 1, 9);
  let data = range.getValues();
  let vList = [];
  let npList = [];

  data.forEach(row => {
    row = row.map((value, index) => (index < 5 ? String(value).toLowerCase().trim() : value));
    if (row[7]) {
      let date = new Date(row[7]);
      row[7] = Utilities.formatDate(date, Session.getScriptTimeZone(), 'MM/dd/yy');
    };

    let listItem = {
      itemName: row[4],
      brand: row[2],
      subCat: row[3],
      category: row[1],
      stock: row[5],
      sold: row[6],
      listDate: row[7],
      orderDate: row[8]
    };

    if (row[0] === 'v' || row[0] === 'V') {
      vList.push(listItem);
    } else if (row[0] === 'np' || row[0] === 'NP') {
      npList.push(listItem);
    } else {
      vList.push(listItem);
      npList.push(listItem);
    };
  });

  let transfersSheet = orderingListSpreadsheet.getSheetByName("Transfers");
  range = transfersSheet.getRange(2, 1, transfersSheet.getLastRow() - 1, 9);
  data = range.getValues();
  data.forEach(row => {
    row = row.map((value, index) => (index < 5 ? String(value).toLowerCase() : value));
    if (row[7]) {
      let date = new Date(row[7]);
      row[7] = Utilities.formatDate(date, Session.getScriptTimeZone(), 'MM/dd/yy');
    };
    npList.push({
      itemName: row[4],
      brand: row[2],
      subCat: row[3],
      category: row[1],
      stock: row[5],
      sold: row[6],
      listDate: row[7],
      orderDate: row[8]
    });
  });

  return { vList, npList };
};

function processLocationSheet(sheet) {
  let range = sheet.getRange(2, 2, sheet.getLastRow() - 1, 4);
  let data = range.getValues();

  let processedData = data.map(row => {
    return row.map((value, index) => (index < 4 ? String(value).toLowerCase() : value));
  });

  return processedData;
};

function compareLists() {
  let { vList, npList } = processOrderingList();
  let veniceSheet = ss.getSheetByName("Alt Venice List");
  let npSheet = ss.getSheetByName("Alt North Port List");
  let veniceData = processLocationSheet(veniceSheet);
  let npData = processLocationSheet(npSheet);

  veniceData.forEach((row, index) => {
    let matchFound = vList.some(item =>
      row[1] === item.itemName && (
        row[0] === item.brand || row[0] === item.subCat
      )
    );
    
    if (matchFound) {
      veniceSheet.getRange(index + 2, 1).setValue(true);
    } else {
      veniceSheet.getRange(index + 2, 1).setValue(false);
    }
  });

  npData.forEach((row, index) => {
    let matchFound = npList.some(item =>
      row[1] === item.itemName && (
        row[0] === item.brand || row[0] === item.subCat
      )
    );
    
    if (matchFound) {
      npSheet.getRange(index + 2, 1).setValue(true);
    } else {
      npSheet.getRange(index + 2, 1).setValue(false);
    }
  });
};
