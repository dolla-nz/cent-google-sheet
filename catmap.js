// ---------------------------------------
// CatMap - Categorize transactions based on rules
// Shown in the UI as Custom Categories
// ---------------------------------------

function catmap() {
  console.info("fn.catmap");

  const catmapSheet = sheetFromName("CentCustomCategories");
  const transactionSheet = sheetFromName("CentTransactions");

  const catHeaders = getHeaders("CentCustomCategories");
  const transactionHeaders = getHeaders("CentTransactions");

  const [clr, clc] = [catmapSheet.getLastRow(), catmapSheet.getLastColumn()];
  const [tlr, tlc] = [
    transactionSheet.getLastRow(),
    transactionSheet.getLastColumn(),
  ];

  if (clr <= 1 || tlr <= 1) {
    console.info("fn.catmap.success", "No rules or transactions");
    return;
  }

  const catRulesArray = catmapSheet.getRange(2, 1, clr - 1, clc).getValues();

  catRulesArray.forEach((ruleRow, ruleIndex) => {
    const transactionArray = transactionSheet
      .getRange(2, 1, tlr - 1, tlc)
      .getValues();

    const ruleObj = createMatchingObject(ruleRow, catHeaders);
    const setObj = createSetObject(ruleRow, catHeaders);

    transactionArray.forEach((transactionRow, transactionIndex) => {
      const transactionObject = createRowObject(
        transactionRow,
        transactionHeaders
      );

      const match = matchRow(transactionObject, ruleObj);

      if (match) {
        const modifiedObj = newRowObject(transactionObject, setObj);
        const modifiedRow = createRowFromObj(modifiedObj, transactionHeaders);

        transactionSheet
          .getRange(transactionIndex + 2, 1, 1, modifiedRow[0].length)
          .setValues(modifiedRow);
        console.info("fn.catmap", "writing to sheet", transactionIndex + 2);
      } else {
        console.info("fn.catmap", "No match");
      }
    });
    console.info("fn.catmap", "flushing", ruleIndex + 2);
    SpreadsheetApp.flush();
  });
  console.info("fn.catmap.success");
}

function pressCatmap() {
  console.info("fn.pressCatmap");
  catmap();
  console.info("fn.pressCatmap.success");
  return;
}

/**
 * Helper functions for catmap
 */

function getHeaders(sheetName) {
  console.info("fn.getHeaders", sheetName);
  let ss;
  if (sheetName) {
    ss = sheetFromName(sheetName);
  } else {
    ss = SpreadsheetApp.getActiveSheet();
  }

  const lastCol = ss.getLastColumn();
  const array = ss.getRange(1, 1, 1, lastCol).getValues()[0];

  console.info("fn.getHeaders.success", array.length);
  return array.map((v) => v.toString().toLowerCase());
}

function criteriaIndexes(header) {
  return header
    .map((v, i) => (v.startsWith("set ") ? false : i))
    .filter((v) => v !== false);
}

/** Creates object with the form
 * {"header": "value"} */
function createRowObject(row, h) {
  const o = h.reduce(
    (prev, v, i) => ({
      ...prev,
      [v]: row[i],
    }),
    {}
  );
  return o;
}

function createMatchingObject(row, h) {
  console.info("fn.createMatchingObject", row, h);
  if (!h) {
    h = getHeaders("CentCustomCategories");
  }

  const o = h.reduce(
    (prev, v, i) => {
      if (h[i].startsWith("set ") || h[i] === "overwrite existing") {
        return prev;
      }

      return row[i] === ""
        ? prev
        : {
            ...prev,
            [v]: typeof row[i] === "string" ? row[i].toLowerCase() : row[i],
          };
    },

    {}
  );

  console.info("fn.createMatchingObject.success", o);
  return o;
}

function createSetObject(row, h) {
  console.info("fn.createSetObject", row, h);
  if (!h) {
    h = getHeaders("CentCustomCategories");
  }
  const o = h.reduce(
    (prev, v, i) => {
      if (!h[i].startsWith("set ") && h[i] !== "overwrite existing") {
        return prev;
      }

      let key = v;
      if (h[i].startsWith("set ")) {
        key = v.slice(4);
      }
      return row[i] === "" ? prev : { ...prev, [key]: row[i] };
    },

    {}
  );

  console.info("fn.createSetObject.success", o);
  return o;
}

// Returns true if the rowObject matches ALL of the matchingObject rules
function matchRow(rowObj, matchingObj) {
  const matches = Object.keys(matchingObj).map((k) => {
    // special cases
    if (k === "minimum") {
      return rowObj["amount"] >= matchingObj[k];
    }

    if (k === "maximum") {
      return rowObj["amount"] <= matchingObj[k];
    }

    if (k === "before") {
      return rowObj["date"] <= matchingObj[k];
    }

    if (k === "after") {
      return rowObj["date"] >= matchingObj[k];
    }

    // generic string matching
    if (typeof rowObj[k] === "string") {
      // case insensitive match
      return rowObj[k].toLowerCase().includes(matchingObj[k]);
    }
    console.error("fn.matchRow.error", k, rowObj[k], matchingObj[k]);
    return true;
  });
  // console.info(matches, !matches.includes(false) )
  return !matches.includes(false);
}

function newRowObject(rowObj, ruleObj) {
  if (ruleObj["overwrite existing"]) {
    return { ...rowObj, ...ruleObj };
  } else {
    Object.keys(ruleObj).forEach(
      (k) => (rowObj[k] = rowObj[k] === "" ? ruleObj[k] : rowObj[k])
    );
    return rowObj;
  }
}

/**
 * Creates a 2D array ready to insert from an array of headers and rowObject
 * @param {string[]} h - Array of header titles
 * @param {Object} rowObj - Object with header titles as keys
 */
function createRowFromObj(rowObj, h) {
  const row = h.map((headerTitle) => rowObj[headerTitle]);
  // return as 2D array for insertion
  return [row];
}
