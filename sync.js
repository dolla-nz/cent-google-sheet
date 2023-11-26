// ---------------------------------------
// Syncing data with the Cent back end
// (which is just a wafer thin proxy for Akahu)
// ---------------------------------------

/**
 * Run through each of our akahu APIs and sync the data.
 * If there's a new account, we'll sync 90 days back, otherwise we'll sync 30 days back.
 * This reduces the load on akahu's servers and makes the sync faster.
 */
function cron() {
  try {
    // Accounts
    console.info("fn.cron", "Syncing Accounts");
    const newAccounts = _syncAccounts();
    //Balances
    console.info("fn.cron", "Syncing Balances");
    _syncBalances();
    // Sync 90 days back if there are new accounts, otherwise 30 days
    console.info("fn.cron", "Syncing Transactions");
    _syncTransactions(newAccounts ? 90 : 30);

    console.info("fn.cron", "Syncing Transactions");
    catmap();

    console.info("fn.cron.success");
  } catch (error) {
    console.error("fn.cron", error);
    throw error;
  }
}

function _syncTransactions(days = 90) {
  console.info("fn._syncTransactions", days);

  const ss = sheetFromName("CentTransactions");
  const userToken = _getUserToken();

  const ids = getIds("CentTransactions");

  // Check if this is a new sheet (ie lr = 1)
  const lr = ss.getLastRow();
  console.info("fn._syncTransactions", "Last row", lr);

  try {
    // flashImage();

    const now = new Date();
    const startDate = new Date(now.getTime() - days * MS_PER_DAY);

    const start = `&start=${startDate.toISOString()}`;
    const end = ""; //`&end=${now.toISOString()}`
    let cursor = "";

    console.info("fn._syncTransactions", "Start", start, "End", end);

    try {
      // Get the accounts
      const res = UrlFetchApp.fetch("https://api.cent.nz/v1/sync/accounts", {
        headers: {
          Authorization: "Bearer " + userToken,
        },
      });

      // Insert the transactions
      const accountsResult = JSON.parse(res);
      const accounts = accountsResult.items;

      do {
        // Get the transactions
        console.info("fn._syncTransactions", "Call Akahu");
        const res = UrlFetchApp.fetch(
          "https://api.cent.nz/v1/sync/transactions?" + start + end + cursor,
          {
            headers: {
              Authorization: "Bearer " + userToken,
            },
          }
        );

        // Insert the transactions
        const transactionResult = JSON.parse(res);
        const transactions = transactionResult.items;

        console.info(
          "fn._syncTransactions",
          "Got",
          transactions.length,
          "transactions"
        );

        // Parse the transactions into a 2d array, ready to plop into the spreadsheet.
        // This is much faster than inserting one row at a time.
        const transactionRows = transactions
          .map((transaction, i) => {
            const {
              _id,
              date,
              description,
              amount,
              balance,
              type,
              merchant,
              category,
              meta,
              _account,
            } = transaction;
            if (_id) {
              if (ids.includes(_id)) {
                return false;
              } else {
                const account = accounts.find((a) => a._id === _account);

                // id, date, description, amount,  accountName, accountNumber,balance, type, merchantname, category, nzfcc, otherAccount particulars, code, ref, insertionDate
                return [
                  _id,
                  new Date(date),
                  description,
                  amount,
                  account?.name,
                  account?.formatted_account,
                  balance,
                  type,
                  merchant?.name,
                  category?.groups?.personal_finance?.name,
                  category?.name,
                  meta?.other_account,
                  meta?.particulars,
                  meta?.code,
                  meta?.reference,
                  now,
                ];
              }
            }
            return false;
          })
          .filter((x) => Boolean(x));

        console.info(
          "fn._syncTransactions",
          "Filtered down to",
          transactionRows.length,
          "transactions"
        );

        if (transactionRows.length) {
          const lastRow = ss.getLastRow();
          const range = ss.getRange(
            lastRow + 1,
            1,
            transactionRows.length,
            transactionRows[0].length
          );
          range.setValues(transactionRows);
          console.info(
            "fn._syncTransactions",
            "Inserted",
            transactionRows.length,
            "transactions",
            `from row ${lastRow + 1} to ${lastRow + transactionRows.length}`
          );
        }

        const cursorRes = transactionResult?.cursor?.next || "null";
        cursor = cursorRes === "null" ? cursorRes : `&cursor=${cursorRes}`;
      } while (cursor !== "null");
    } catch (e) {
      console.error("fn._syncTransactions", e);
      SpreadsheetApp.getActive().toast("⚠️ Error fetching transactions");
    }

    // resize if this is a new sheet
    if (lr <= 1) {
      console.info("fn._syncTransactions", "Resizing Column width");
      const lc = ss.getLastColumn();
      // resize columns excluding the id column
      ss.autoResizeColumns(2, lc - 1);
    }

    const [nlr, nlc] = [ss.getLastRow(), ss.getLastColumn()];
    console.info("new lastrow", nlr, "new last column", nlc);
    // Sort by date
    if (nlr - lr > 0) {
      console.info(
        "fn._syncTransactions",
        "Sorting by date in range",
        lr + 1,
        1,
        nlr - lr,
        nlc
      );
      ss.getRange(lr + 1, 1, nlr - lr, nlc).sort({
        column: 2,
        ascending: true,
      });
    }

    console.info("fn._syncTransactions", "flush sheet");
    SpreadsheetApp.flush();
    console.info("fn._syncTransactions.success");
  } catch (e) {
    console.error("fn._syncTransactions", e);
    SpreadsheetApp.getActive().toast(f.message, "⚠️ Error");
  }
}

function _syncBalances() {
  console.info("fn.syncBalances_");

  const ss = sheetFromName("CentBalanceHistory");
  const userToken = _getUserToken();

  // Check if this is a new sheet
  const lr = ss.getLastRow();
  if (lr <= 1) {
    console.info("fn.syncBalances_", "Resizing Column width");
    const lc = ss.getLastColumn();
    // resize columns excluding the id column
    ss.autoResizeColumns(2, lc - 1);
  }

  try {
    // Get the accounts
    console.info("fn.syncBalances_", "Calling Akahu");

    const res = UrlFetchApp.fetch("https://api.cent.nz/v1/sync/accounts", {
      headers: {
        Authorization: "Bearer " + userToken,
      },
    });

    const accountsResult = JSON.parse(res);
    const accounts = accountsResult.items;
    console.info("fn.syncBalances_", "Got", accounts.length, "accounts");

    const now = new Date();

    // Parse the accounts into a 2d array, ready to plop into the spreadsheet.
    const balanceRows = accounts
      // Remove inactive accounts
      .filter((account) => account.status === "ACTIVE")
      .map((account) => {
        const { _id, connection, name, formatted_account, balance, type } =
          account;
        if (_id) {
          return [
            _id,
            connection?.name,
            name,
            formatted_account,
            type,
            balance?.current,
            balance?.available,
            now,
          ];
        }
        return false;
      })
      .filter((x) => Boolean(x));

    console.info("fn.syncBalances_", "Got", balanceRows.length, "balance rows");

    // Insert new rows at the bottom of the spreadsheet
    if (balanceRows.length) {
      console.info("fn.syncBalances_", "Inserting Rows");
      const lastRow = ss.getLastRow();
      const range = ss.getRange(
        lastRow + 1,
        1,
        balanceRows.length,
        balanceRows[0].length
      );
      range.setValues(balanceRows);
    }

    console.info("fn.syncBalances_", "flush sheet");
    SpreadsheetApp.flush();
  } catch (e) {
    console.error("fn.syncBalances_", e);
    SpreadsheetApp.getActive().toast(e.message, "⚠️ Error");
  }

  // Check if this is a new sheet
  if (lr <= 1) {
    console.info("fn.syncBalances_", "Resizing Column width");
    const lc = ss.getLastColumn();
    // resize columns excluding the id column
    ss.autoResizeColumns(2, lc - 1);
  }
  console.info("fn.syncBalances_.success");
}

function _syncAccounts() {
  console.info("fn._syncAccounts");
  const ss = sheetFromName("CentAccounts");
  const userToken = _getUserToken();

  // Check if this is a new sheet
  const lr = ss.getLastRow();
  if (lr <= 1) {
    console.info("fn._syncAccounts", "Resizing Column width");
    const lc = ss.getLastColumn();
    // resize columns excluding the id column
    ss.autoResizeColumns(2, lc - 1);
  }

  try {
    console.info("fn._syncAccounts", "Fetching Accounts");
    // Get the accounts
    const res = UrlFetchApp.fetch("https://api.cent.nz/v1/sync/accounts", {
      headers: {
        Authorization: "Bearer " + userToken,
      },
    });

    // Insert the transactions
    const accountsResult = JSON.parse(res);
    const accounts = accountsResult.items;
    const now = new Date();

    const ids = getIds("CentAccounts");

    // Mark accounts not in the response as DELETED
    const fetchedIds = accounts.map((account) => account._id);

    // Check and mark rows as DELETED if _id not found in fetchedIds
    for (let i = 2; i <= lr; i++) {
      const rowId = ss.getRange(i, 1).getValue();
      if (!fetchedIds.includes(rowId)) {
        ss.getRange(i, 8).setValue("DELETED");
      }
    }

    // Parse the accounts into a 2d array, ready to plop into the spreadsheet.
    // This is much faster than inserting one row at a time.
    const accountRows = accounts
      .map((account, i) => {
        const {
          _id,
          connection,
          name,
          status,
          formatted_account,
          balance,
          refreshed,
          type,
        } = account;
        if (_id) {
          if (ids.includes(_id)) {
            // replace row
            const rowIndex = ids.indexOf(_id) + 1;
            console.info(
              "fn._syncAccounts",
              "replacing row",
              _id,
              "rowIndex",
              rowIndex
            );
            const newRow = [
              _id,
              connection?.name,
              name,
              formatted_account,
              type,
              balance?.current,
              balance?.available,
              status,
              new Date(refreshed?.balance),
            ];
            ss.getRange(rowIndex, 1, 1, newRow.length).setValues([newRow]);
            return false;
          }
          return [
            _id,
            connection?.name,
            name,
            formatted_account,
            type,
            balance?.current,
            balance?.available,
            status,
            new Date(refreshed?.balance),
            now,
          ];
        }
        return false;
      })
      .filter((x) => Boolean(x));
    if (accountRows.length) {
      console.info("fn._syncAccounts", `Adding ${accountRows.length} new rows`);
      const lr = ss.getLastRow();
      const range = ss.getRange(
        lr + 1,
        1,
        accountRows.length,
        accountRows[0].length
      );
      range.setValues(accountRows);
    }

    SpreadsheetApp.flush();

    console.info("fn._syncAccounts.success", accountRows.length || 0);
    // Return the number of new accounts
    return accountRows.length || 0;
  } catch (f) {
    console.error("fn._syncAccounts", f);
    SpreadsheetApp.getActive().toast(f.message, "⚠️ Error");
    return 0;
  }
}

/** A function to sync categories in the script properties with those from akahu/nzfcc */
function syncCategories() {
  try {
    console.info("fn.syncCategories");

    const res = UrlFetchApp.fetch(
      "https://nzfcc.org/downloads/categories.json"
    );
    const json = JSON.parse(res.getContentText());
    const nzfcc = json.map((c) => c?.name || false).filter((x) => Boolean(x));
    const pfm = json
      .map((c) => c?.groups?.personal_finance?.name || false)
      .filter((x) => Boolean(x));

    const dedupeNzfcc = [...new Set(nzfcc)];
    const dedupePfm = [...new Set([...pfm, "Other", "Income"])];

    const ps = PropertiesService.getScriptProperties();
    ps.setProperty(keys.nzfcc_categories, JSON.stringify(dedupeNzfcc));
    ps.setProperty(keys.pfm_categories, JSON.stringify(dedupePfm));

    console.info("fn.syncCategories.success");
    return;
  } catch (error) {
    console.error("fn.syncCategories", error);
  }
}
