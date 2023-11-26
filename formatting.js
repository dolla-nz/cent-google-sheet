// ---------------------------------------
// Formatting and styling functions
// ---------------------------------------

/**
 * Creates a header row with basic formatting for the given sheet name
 * @param {string} sheetName
 * @returns void
 */
function createHeaderRow(sheetName) {
  console.info("fn.createHeaderRow", sheetName);
  const ss = sheetFromName(sheetName);

  // id, date, description, amount, balance, accountName, accountNumber, type, merchantname, category, nzfcc, otherAccount code, ref, particulars, insertionDate
  const headers = {
    CentTransactions: [
      `=image("${CENT_ICON}")`,
      "Date",
      "Description",
      "Amount",
      "Account Name",
      "Account Number",
      "Balance",
      "Type",
      "Merchant Name",
      "Category",
      "NZFCC.org",
      "Other Account",
      "Particulars",
      "Code",
      "Reference",
      "Date Added",
    ],
    CentAccounts: [
      `=image("${CENT_ICON}")`,
      "Institution Name",
      "Account Name",
      "Account Number",
      "Type",
      "Current Balance",
      "Available Balance",
      "Status",
      "Date Refreshed",
      "Date Added",
    ],
    CentBalanceHistory: [
      `=image("${CENT_ICON}")`,
      "Institution Name",
      "Account Name",
      "Account Number",
      "Type",
      "Current Balance",
      "Available Balance",
      "Date",
    ],
    CentCustomCategories: [
      "Set Category",
      "Description",
      "NZFCC.org",
      "Minimum",
      "Maximum",
      "Overwrite Existing",
    ],
  };

  const headerRow = headers[sheetName];

  if (!headerRow) {
    console.error("fn.createHeaderRow", `No headers for ${sheetName} found`);
    return;
  }

  const moneyCols = [
    "Amount",
    "Balance",
    "Current Balance",
    "Available Balance",
    "Minimum",
    "Maximum",
  ];

  headerRow.forEach((header, i) => {
    if (header === `=image("${CENT_ICON}")`) {
      formatIdColumn(i + 1);
    } else if (header === "NZFCC.org") {
      createDropdown(i + 1, keys.nzfcc_categories);
    } else if (header === "Category") {
      createDropdown(i + 1, keys.pfm_categories);
    } else if (header === "Overwrite Existing") {
      createDropdown(i + 1, ["", "Yes"]);
    } else if (moneyCols.includes(header)) {
      formatMoneyColumn(i + 1);
      if (header === "Amount") {
        formatAmountColumn(i + 1);
      }
    } else if (header.includes("Date")) {
      formatDateColumn(i + 1);
    }
    console.info("fn.createHeaderRow.success");
  });

  const h = ss.getRange("1:1");
  h.setFontColor("#ffffff").setBackground(CENT_PINK);
  ss.getRange(1, 1, 1, headerRow.length).setValues([headerRow]);

  // Set the header row height
  ss.setRowHeight(1, 50);

  if (headerRow[0] === `=image("${CENT_ICON}")`) {
    // Set the id column width
    ss.setColumnWidth(1, 50);
  }

  ss.setFrozenRows(1);
}

/**
 * Creates a comment for the given sheet name
 */
function createComment(sheetName) {
  console.info("fn.createComment", sheetName);
  const comments = {
    CentTransactions:
      "This sheet contains all the transactions from every account connected to Cent",
    CentAccounts: "This sheet contains all the accounts connected to Cent",
    CentBalanceHistory:
      "This sheet adds a new row for each account when a sync is completed, which allows you to track the balance of your accounts over time.",
  };

  if (comments[sheetName]) {
    const ss = sheetFromName(sheetName);
    ss.getRange(1, 1).setNote(comments[sheetName]);
    console.info("fn.createComment", "note added", comments[sheetName]);
  }
  console.info("fn.createComment.success");
}

/**
 * Creates a dropdown for the given column index and type. The type should be either pfm_categories or nzfcc_categories
 * @param {number} colIndex
 * @param {string|string[]} type - the type of dropdown to create or an array of options
 * @returns void
 */
function createDropdown(colIndex, type) {
  console.info("fn.createDropdown", colIndex, type);
  if (!colIndex || !type) {
    throw new Error("Need both column index and dropdown type");
  }

  let cat = type;

  if (type === keys.pfm_categories || type === keys.nzfcc_categories) {
    cat = [
      "",
      ...JSON.parse(PropertiesService.getScriptProperties().getProperty(type)),
    ];
  }
  console.info("fn.createDropdown", "cat", cat);

  const ss = SpreadsheetApp.getActiveSheet();

  const range = ss.getRange(2, colIndex, 1000);

  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList([...cat])
    .build();

  range.setDataValidation(rule);
  console.info("fn.createDropdown.success");
}

/**
 * Sets the ids column to be white text on a white background
 * @param {number} colIndex
 */
function formatIdColumn(colIndex) {
  console.info("fn.formatIdColumn", colIndex);
  const ss = SpreadsheetApp.getActiveSheet();

  const range = ss.getRange(2, colIndex, 1000);
  range.setFontColor("#FFFFFF");
  console.info("fn.formatIdColumn.success");
}

/**
 * Sets the number format for the given column index to nicely display NZD amounts
 * @param {number} colIndex
 */
function formatMoneyColumn(colIndex) {
  console.info("fn.formatMoneyColumn", colIndex);
  const ss = SpreadsheetApp.getActiveSheet();

  const range = ss.getRange(2, colIndex, 1000);
  range.setNumberFormat("$#,##0.00");
  console.info("fn.formatMoneyColumn.success");
}

/**
 * Sets the number format for the given column index to nicely display NZD amounts
 * @param {number} colIndex
 */
function formatDateColumn(colIndex) {
  console.info("fn.formatDateColumn", colIndex);
  const ss = SpreadsheetApp.getActiveSheet();

  const range = ss.getRange(2, colIndex, 1000);
  range.setNumberFormat("dd/mm/yyyy");
  console.info("fn.formatDateColumn.success");
}

/**
 * Sets the font colour to green for positive values
 * @param {number} colIndex
 */
function formatAmountColumn(colIndex) {
  try {
    console.info("fn.formatAmountColumn", colIndex);
    const ss = SpreadsheetApp.getActiveSheet();
    const range = ss.getRange(2, colIndex, 1000);

    const rule = SpreadsheetApp.newConditionalFormatRule()
      .whenNumberGreaterThan(0)
      .setFontColor(DOLLA_GREEN)
      .setRanges([range])
      .build();
    const rules = ss.getConditionalFormatRules();
    rules.push(rule);
    ss.setConditionalFormatRules(rules);
    console.info("fn.formatAmountColumn.success");
  } catch (error) {
    console.error("fn.formatAmountColumn", error);
  }
}
