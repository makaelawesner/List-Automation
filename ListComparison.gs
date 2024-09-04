function toTitleCase(str) {
    return str.toLowerCase().split(' ').map(function(word) {
        return word.charAt(0).toUpperCase() + word.slice(1);
    }).join(' ');
};

function processOrderingList() {
  let orderingListSpreadsheet = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/19rV2sw-JVJaYHYmG0xC5FrpqE9uGjKYKOdBOiwvBxI4/edit?gid=389276120#gid=389276120");
  let orderingSheet = orderingListSpreadsheet.getSheetByName("Ordering List");
  let range = orderingSheet.getRange(2, 1, orderingSheet.getLastRow() - 1, 9);
  let data = range.getValues();
  let vList = [];
  let npList = [];

  data.forEach(row => {
    row = row.map((value, index) => (index < 5 ? String(value).toLowerCase() : value));
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

    if (row[0] === 'v') {
      vList.push(listItem);
    } else if (row[0] === 'np') {
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
  let range = sheet.getRange(2, 1, sheet.getLastRow() - 1, 4);
  let data = range.getValues();

  let processedData = data.map(row => {
    return row.map((value, index) => (index < 4 ? String(value).toLowerCase() : value));
  });

  return processedData;
};

function writeMatchingResultsToSheet(sheet, data) {
  sheet.getRange(1, 1, 1, 7).setValues([['Category', 'Brand', 'SubCat', 'Item Name', 'Stock', 'Sold L.M.', 'Listed Date']]);
  
  if (data.length > 0) {
    let values = data.map(item => [
        toTitleCase(item.category), 
        toTitleCase(item.brand), 
        toTitleCase(item.subCat), 
        toTitleCase(item.itemName), 
        item.stock, 
        item.sold, 
        item.listDate
    ]);
    sheet.getRange(2, 1, values.length, 7).setValues(values);
  };

  sheet.getDataRange().applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY);
  sheet.setFrozenRows(1);
  sheet.getDataRange().createFilter();
  sheet.getRange('D1').activate().getFilter().sort(2, true);
  sheet.getRange('C1').activate().getFilter().sort(3, true);
  sheet.getRange('B1').activate().getFilter().sort(2, true);
  sheet.getRange('A1').activate().getFilter().sort(1, true);
  sheet.setColumnWidth(1, 100);
  sheet.setColumnWidth(2, 150);
  sheet.setColumnWidth(3, 150);
  sheet.setColumnWidth(4, 200);
  sheet.autoResizeColumn(5);
  sheet.autoResizeColumn(6);
  sheet.autoResizeColumn(7);
};

function writeNonMatchingResultsToSheet(sheet, data) {
    sheet.getRange(1, 1, 1, 5).setValues([['Category', 'Brand', 'SubCat', 'Item', 'Sold L.M.']]);

    if (data.length > 0) {
        let values = data.map(item => {
            let brand, subCat, itemName;

            if (item.category !== "disposables" && item.itemName.includes('-')) {
                [brand, subCat] = item.itemName.split('-').map(str => str.trim());
                itemName = subCat; 
            } else {
                brand = item.itemName;
                subCat = 'n/a';
                itemName = brand;
            };

            if ((item.category === "sstash" && subCat !== "cones" && subCat !== "fluid") || item.category === "accessories" || item.category === "vending") {
                [brand, subCat] = [subCat, brand];
            } else if (item.category === "sstash" && subCat === "cones") {
                item.variationName = brand + " - " + item.variationName;
                brand = "n/a";
            } else if (item.category === "sstash" && subCat === "fluid") {
                item.variationName = brand + " - " + item.variationName;
                subCat = brand;
                brand = "n/a";
            };

            if (item.category === "cbd") {
                if (item.variationName.includes('-')) {
                    subCat = item.variationName.split('-')[0].trim();
                    item.variationName = item.variationName.split('-')[1].trim();
                };
            };

            if (item.category === "coils & pods") {
                let subCatValue = subCat;
                subCat = brand.split(' ')[1].trim() + " - " + subCatValue;
                brand = brand.split(' ')[0].trim();
            };

            if (item.category === "e-liquid") {
                const variation = item.variationName.toLowerCase();

                if (variation.includes("3mg") || variation.includes("6mg")) {
                    subCat = "freebase";
                } else {
                    subCat = "salt";
                };
            };

            if ((item.category === "kits" || item.category === "mods" || item.category === "mech" || item.category === "tanks") && !brand.startsWith("x ")) {
                brand = brand.split(' ')[0].trim();
            };

            return [
                toTitleCase(item.category), 
                toTitleCase(brand), 
                toTitleCase(subCat), 
                toTitleCase(item.variationName), 
                item.qty
            ];
        });

        sheet.getRange(2, 1, values.length, 5).setValues(values);
    };

    sheet.getDataRange().applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, sheet.getLastColumn()).setHorizontalAlignment('center');
    sheet.getRange(1, 4, sheet.getLastRow(), 2).setHorizontalAlignment('center');
    sheet.getDataRange().createFilter();
    sheet.getRange('D1').activate().getFilter().sort(2, true);
    sheet.getRange('C1').activate().getFilter().sort(3, true);
    sheet.getRange('B1').activate().getFilter().sort(2, true);
    sheet.getRange('A1').activate().getFilter().sort(1, true);
    sheet.setColumnWidth(5, 50);
    sheet.setColumnWidth(4, 300);
    sheet.setColumnWidth(3, 150);
    sheet.setColumnWidth(2, 150);
    sheet.setColumnWidth(1, 100);
    hideUnsortedNomoItems(sheet);
    // findAndReplace(sheet, 3, "Coils", "Coil");
    // findAndReplace(sheet, 3, "Pods", "Pod");
};

function hideNPVending(sheet) {
  const lastRow = sheet.getLastRow();
  const valuesA = sheet.getRange(`A2:A${lastRow}`).getValues();
  const rowsToHide = new Set();

  valuesA.forEach((valueA, i) => {
    const valueAString = valueA[0].toString().toLowerCase();
    if (valueAString.startsWith("vending")) {
      rowsToHide.add(i + 2);
    };
  });

  rowsToHide.forEach(row => sheet.hideRows(row));
};

function compareLists() {
  let { vList, npList } = processOrderingList();
  let veniceData = processLocationSheet(ss.getSheetByName("Alt Venice List"));
  let npData = processLocationSheet(ss.getSheetByName("Alt North Port List"));

  let veniceItemNames = vList.map(item => item.itemName);
  let npItemNames = npList.map(item => item.itemName);

  let veniceMatches = vList.filter(item => 
    veniceData.some(row => 
      row[1] === item.itemName && (
        row[0] === item.brand || row[0].includes(item.subCat)
      )
    )
  );

  let npMatches = npList.filter(item => 
    npData.some(row => 
      row[1] === item.itemName && (
        row[0] === item.brand || row[0].includes(item.subCat)
      )
    )
  );

  let veniceNotInOrdering = veniceData.filter(row => 
    !veniceItemNames.includes(row[1])
  ).map(row => ({
    category: row[2],
    itemName: row[0],
    variationName: row[1],
    qty: row[3]
  }));

  let npNotInOrdering = npData.filter(row => 
    !npItemNames.includes(row[1])
  ).map(row => ({
    category: row[2],
    itemName: row[0],
    variationName: row[1],
    qty: row[3]
  }));

  createOrReplaceSheet(" | ");

  createOrReplaceSheet("Venice - Listed");
  const veniceListedSheet = ss.getSheetByName("Venice - Listed");
  veniceListedSheet.setTabColor(veniceColor);
  writeMatchingResultsToSheet(veniceListedSheet, veniceMatches);

  createOrReplaceSheet("Venice - Unlisted");
  const veniceUnlistedSheet = ss.getSheetByName("Venice - Unlisted");
  veniceUnlistedSheet.setTabColor(veniceColor);
  writeNonMatchingResultsToSheet(veniceUnlistedSheet, veniceNotInOrdering);

  createOrReplaceSheet("North Port - Listed");
  const npListedSheet = ss.getSheetByName("North Port - Listed");
  npListedSheet.setTabColor(npColor);
  writeMatchingResultsToSheet(npListedSheet, npMatches);

  createOrReplaceSheet("North Port - Unlisted");
  const npUnlistedSheet = ss.getSheetByName("North Port - Unlisted");
  npUnlistedSheet.setTabColor(npColor);
  writeNonMatchingResultsToSheet(npUnlistedSheet, npNotInOrdering);
  hideNPVending(npUnlistedSheet);
};

// function processOrderingList() {
//   let vList = [];
//   let npList = [];
//   let orderingListSpreadsheet = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/19rV2sw-JVJaYHYmG0xC5FrpqE9uGjKYKOdBOiwvBxI4/edit?gid=389276120#gid=389276120");
//   let orderingSheet = orderingListSpreadsheet.getSheetByName("Ordering List");
//   let range = orderingSheet.getRange(2, 1, orderingSheet.getLastRow(), 9);
//   let data = range.getValues();
//   data.forEach(row => {
//     if (row[7]) {
//       let date = new Date(row[7]);
//       row[7] = Utilities.formatDate(date, Session.getScriptTimeZone(), 'MM/dd/yy');
//     };

//     let listItem = {
//       itemName: String(row[4]).toLowerCase().trim(),
//       brand: String(row[2]).toLowerCase().trim(),
//       subCat: String(row[3]).toLowerCase().trim(),
//       category: String(row[1]).toLowerCase().trim(),
//       stock: Number(row[5]),
//       sold: Number(row[6]),
//       listDate: row[7],
//       orderDate: row[8],
//       productName: `${String(row[1]).toLowerCase().trim()} ${String(row[2]).toLowerCase().trim()} ${String(row[3]).toLowerCase().trim()} ${String(row[4]).toLowerCase().trim()}`
//     };

//     if (row[0] === 'v') {
//       vList.push(listItem);
//     } else if (row[0] === 'np') {
//       npList.push(listItem);
//     } else {
//       vList.push(listItem);
//       npList.push(listItem);
//     };
//   });
  
//   let transfersSheet = orderingListSpreadsheet.getSheetByName("Transfers");
//   range = transfersSheet.getRange(2, 1, transfersSheet.getLastRow(), 9);
//   data = range.getValues(); 
//   data.forEach(row => {
//     row = row.map((value, index) => (index < 5 ? String(value).toLowerCase() : value));
//     if (row[7]) {
//       let date = new Date(row[7]);
//       row[7] = Utilities.formatDate(date, Session.getScriptTimeZone(), 'MM/dd/yy');
//     };
//     let transferItem = {
//       itemName: String(row[4]).toLowerCase().trim(),
//       brand: String(row[2]).toLowerCase().trim(),
//       subCat: String(row[3]).toLowerCase().trim(),
//       category: String(row[1]).toLowerCase().trim(),
//       stock: Number(row[5]),
//       sold: Number(row[6]),
//       listDate: row[7],
//       orderDate: row[8],
//       productName: `${String(row[1]).toLowerCase().trim()} ${String(row[2]).toLowerCase().trim()} ${String(row[3]).toLowerCase().trim()} ${String(row[4]).toLowerCase().trim()}`
//     };
//     npList.push(transferItem);
//   });

//   return { vList, npList };
// };

// function processLocationSheet(sheet) {
//   let processedData = [];
//   let range = sheet.getRange(2, 1, sheet.getLastRow() - 1, 4);
//   let data = range.getValues();

//   data.forEach(row => {
//     let listItem = {
//       itemName: String(row[0]).toLowerCase().trim(),
//       variation: String(row[1]).toLowerCase().trim(),
//       category: String(row[2]).toLowerCase().trim(),
//       qty: Number(row[3]),
//       productName: `${String(row[2]).toLowerCase().trim()} ${String(row[0]).toLowerCase().trim()} ${String(row[1]).toLowerCase().trim()}`
//     };
//     processedData.push(listItem);
//   });
//   return processedData;
// };

// function compareLists() {
//   let { vList, npList } = processOrderingList();
//   let veniceData = processLocationSheet(ss.getSheetByName("Alt Venice List"));
//   let npData = processLocationSheet(ss.getSheetByName("Alt North Port List"));

//   let veniceProducts = new Set(vList.map(item => item.productName));
//   let npProducts = new Set(npList.map(item => item.productName));

//   let commonValues = [];
//   let uniqueValues = [];

//   veniceData.forEach(item => {
//     if (veniceProducts.has(item.productName)) {
//       commonValues.push(item);
//     } else {
//       uniqueValues.push(item);
//     }
//   });

//   console.log("Common Values in Both Lists:");
//   commonValues.forEach(value => console.log(value.productName));

//   console.log("\nUnique Values in Data But Not in Lists:");
//   uniqueValues.forEach(value => console.log(value.productName));
// };
