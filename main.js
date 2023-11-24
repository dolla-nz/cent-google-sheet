/** Some important constants */
const DOLLA_BLUE = "#336ce7";
const DOLLA_RED = "#cb2431";
const DOLLA_GREEN = "#34C759";
const DOLLA_LOGO = "https://static.dolla.nz/images/logo.png";
const CENT_ICON = "https://static.dolla.nz/images/cent-icon.png";
const CENT_LOGO = "https://static.dolla.nz/images/cent-logo.png";
const GET_QR_CODE = "https://static.dolla.nz/images/get-qr.png";
const CENT_PINK = "#EE3C80";

const MS_PER_DAY = 1000 * 60 * 60 * 24; // for date maths

// Less error prone than using the strings directly
const keys = {
  user_token: "cent_user_token",
  app_token: "cent_app_token",
  daily_cron: "cent_cron_enabled",
  nzfcc_categories: "nzfcc_categories",
  pfm_categories: "pfm_categories",
  action: "cent_action",
};

function onLaunch() {
  console.info("fn.onLaunch");
  const readOnly = isSheetReadOnly();

  if (readOnly) {
    PropertiesService.getUserProperties().setProperty(keys.action, "READ_ONLY");
    return createErrorCard(
      "This sheet is read-only. Please make a copy of this sheet or create a new sheet to use Cent."
    );
  }

  initialiseSheet();

  const cronEnabled = PropertiesService.getDocumentProperties().getProperty(
    keys.daily_cron
  );
  // ie not set to "true" or "false"
  if (!cronEnabled) {
    console.info("fn.onLaunch", "Creating cron trigger for first launch");
    createCronTrigger();
    PropertiesService.getDocumentProperties().setProperty(
      keys.daily_cron,
      "true"
    );
  }

  console.info("fn.onLaunch.success");
  return createMainCard();
}

/**
 * Creates the main UI screen for the sidebar. If there isn't a user token to call akahu with,
 * it will render the sign up card instead.
 */
function createMainCard() {
  console.info("fn.createMainCard");
  nav = CardService.newNavigation();
  SpreadsheetApp.getActiveSheet();
  const documentProperties = PropertiesService.getDocumentProperties();

  const cronEnabled = documentProperties.getProperty(keys.daily_cron) ?? false;

  const userToken = _getUserToken();

  if (!userToken) {
    return createSigninCard();
  }

  const connectionText = CardService.newTextParagraph().setText(
    "✅ Successfully connected"
  );

  // The sync button
  const action = CardService.newAction().setFunctionName("pressSyncButton");
  const syncButton = CardService.newTextButton()
    .setText("Sync")
    .setOnClickAction(action)
    .setTextButtonStyle(CardService.TextButtonStyle.FILLED);
  const syncText = CardService.newDecoratedText()
    .setText("Manually sync all data now")
    .setButton(syncButton);

  const catmapAction = CardService.newAction().setFunctionName("pressCatmap");
  const catmapButton = CardService.newTextButton()
    .setText("Categorise")
    .setOnClickAction(catmapAction)
    .setTextButtonStyle(CardService.TextButtonStyle.FILLED);
  const catmapText = CardService.newDecoratedText()
    .setText("Run custom rules")
    .setButton(catmapButton);

  const enableSyncing = CardService.newDecoratedText()
    .setTopLabel("Enable Automatic Syncing")
    .setText("Cent will automatically add the latest data around 1AM each day")
    .setWrapText(true)
    .setSwitchControl(
      CardService.newSwitch()
        .setFieldName("cronEnabled")
        .setValue(Boolean(true))
        .setSelected(cronEnabled === "true" ? true : false)
        .setOnChangeAction(
          CardService.newAction().setFunctionName("handleSwitchChange")
        )
    );

  // The footer buttons
  const authAction = CardService.newAction().setFunctionName("pressLogin");
  const revokeAction = CardService.newAction().setFunctionName("pressRevoke");
  const footer = CardService.newFixedFooter()
    .setPrimaryButton(
      CardService.newTextButton()
        .setText("Add Accounts")
        .setBackgroundColor(DOLLA_BLUE)
        .setOnClickAction(authAction)
    )
    .setSecondaryButton(
      CardService.newTextButton()
        .setText("Revoke Accounts")
        .setBackgroundColor(DOLLA_RED)
        .setOnClickAction(revokeAction)
    );

  // Assemble the widgets and return the card.
  const section = CardService.newCardSection()
    .addWidget(connectionText)
    .addWidget(syncText)
    .addWidget(enableSyncing)
    .addWidget(catmapText);
  const card = CardService.newCardBuilder()
    .addSection(section)
    .setFixedFooter(footer);

  console.info("fn.createMainCard.success");

  return card.build();
}

function createSigninCard() {
  console.info("fn.createSigninCard");
  const logoGI = CardService.newGridItem()
    .setTextAlignment(CardService.HorizontalAlignment.CENTER)
    .setLayout(CardService.GridItemLayout.TEXT_BELOW)
    .setImage(CardService.newImageComponent().setImageUrl(CENT_LOGO));

  const cardSectionGrid = CardService.newGrid()
    .setNumColumns(1)
    .addItem(logoGI);

  const signupText = CardService.newTextParagraph().setText(
    "Cent allows you to automatically sync information from your bank to Google Sheets"
  );

  const authAction = CardService.newAction().setFunctionName("pressLogin");
  const authButton = CardService.newTextButton()
    .setText("Sign in with Akahu")
    .setOnClickAction(authAction)
    .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
    .setBackgroundColor(DOLLA_BLUE);

  const cardSection = CardService.newCardSection()
    .addWidget(cardSectionGrid)
    .addWidget(signupText)
    .addWidget(authButton);

  const card = CardService.newCardBuilder().addSection(cardSection);

  console.info("fn.createSigninCard.success");

  return card.build();
}

function pressLogin() {
  console.info("fn.pressLogin");
  var nav = CardService.newNavigation().updateCard(createOAuthWaitingCard());
  console.info("fn.pressLogin.success");
  return CardService.newActionResponseBuilder()
    .setStateChanged(true)
    .setNavigation(nav)
    .setOpenLink(
      CardService.newOpenLink()
        .setUrl(getAkahuOauthUrl())
        // Something is going wrong here and it's VERY ANNOYING
        // Won't automatically refresh card, so a user has to press 'done'
        .setOnClose(CardService.OnClose.NOTHING)
        .setOpenAs(CardService.OpenAs.OVERLAY)
    )
    .build();
}

function pressRevoke() {
  console.info("fn.pressRevoke");

  const userToken = _getUserToken();
  if (userToken) {
    const res = UrlFetchApp.fetch("https://api.cent.nz/v1/auth", {
      headers: {
        Authorization: "Bearer " + userToken,
      },
      method: "delete",
    });
  }

  const userProperties = PropertiesService.getUserProperties();
  userProperties.setProperty(keys.user_token, "");
  userProperties.setProperty(keys.action, "");
  var nav = CardService.newNavigation().updateCard(createSigninCard());
  console.info("fn.pressRevoke.success");
  return CardService.newActionResponseBuilder()
    .setStateChanged(true)
    .setNavigation(nav)
    .build();
}

function createOAuthWaitingCard() {
  console.info("fn.createOAuthWaitingCard");
  const logoGI = CardService.newGridItem()
    .setTextAlignment(CardService.HorizontalAlignment.CENTER)
    .setLayout(CardService.GridItemLayout.TEXT_BELOW)
    .setImage(CardService.newImageComponent().setImageUrl(CENT_LOGO));

  const cardSectionGrid = CardService.newGrid()
    .setNumColumns(1)
    .addItem(logoGI);

  const signupText = CardService.newTextParagraph().setText(
    "Loading in a new tab"
  );

  const doneAction = CardService.newAction().setFunctionName("onDone");
  const doneButton = CardService.newTextButton()
    .setText("Done")
    .setOnClickAction(doneAction)
    .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
    .setBackgroundColor(DOLLA_BLUE);
  const cancelButton = CardService.newTextButton()
    .setText("Cancel")
    .setOnClickAction(doneAction)
    .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
    .setBackgroundColor(DOLLA_RED);

  const buttonSet = CardService.newButtonSet()
    .addButton(cancelButton)
    .addButton(doneButton);

  const cardSection = CardService.newCardSection()
    .addWidget(cardSectionGrid)
    .addWidget(signupText)
    .addWidget(buttonSet);

  const card = CardService.newCardBuilder().addSection(cardSection);

  console.info("fn.createOAuthWaitingCard.success");

  return card.build();
}

function onDone() {
  console.info("fn.onDone");
  var nav = CardService.newNavigation().updateCard(createMainCard());
  console.info("fn.onDone.success");
  return CardService.newActionResponseBuilder()
    .setStateChanged(true)
    .setNavigation(nav)
    .build();
}

function createErrorCard(
  message = "An unexpected error has occurred - please try again."
) {
  console.info("fn.createErrorCard");
  const logoGI = CardService.newGridItem()
    .setTextAlignment(CardService.HorizontalAlignment.CENTER)
    .setLayout(CardService.GridItemLayout.TEXT_BELOW)
    .setImage(CardService.newImageComponent().setImageUrl(CENT_LOGO));

  const cardSectionGrid = CardService.newGrid()
    .setNumColumns(1)
    .addItem(logoGI);

  const errorText = CardService.newTextParagraph().setText(
    `❌ Error: ${message}`
  );

  const cardSection = CardService.newCardSection()
    .addWidget(cardSectionGrid)
    .addWidget(errorText);

  const card = CardService.newCardBuilder().addSection(cardSection);

  console.info("fn.createErrorCard.success");
  return card.build();
}

function handleSwitchChange(e) {
  console.info("fn.handleSwitchChange", e);

  const cronEnabled = e?.formInput?.cronEnabled ?? false;
  if (cronEnabled) {
    createCronTrigger();
  } else {
    disableCronTrigger();
  }
  PropertiesService.getDocumentProperties().setProperty(
    keys.daily_cron,
    cronEnabled
  );
  console.info(
    "fn.handleSwitchChange",
    "Set cronEnabled property to",
    cronEnabled
  );
  // Create a new card with the same text.
  const card = createMainCard();

  // Create an action response that instructs the add-on to replace
  // the current card with the new one.
  const navigation = CardService.newNavigation().updateCard(card);
  const actionResponse =
    CardService.newActionResponseBuilder().setNavigation(navigation);

  console.info("fn.handleSwitchChange.success");
  return actionResponse.build();
}

function pressSyncButton() {
  console.info("fn.pressSyncButton");
  cron();

  const card = createMainCard();

  // Create an action response that instructs the add-on to replace
  // the current card with the new one. This prevents it building a
  // navigation tree with back buttons to the same card
  const navigation = CardService.newNavigation().updateCard(card);
  const actionResponse =
    CardService.newActionResponseBuilder().setNavigation(navigation);

  console.info("fn.pressSyncButton.success");
  return actionResponse.build();
}

function pressSyncTransactions() {
  console.info("fn.pressSyncTransactions");
  _syncTransactions();
  console.info("fn.pressSyncTransactions.success");
}

function pressSyncAccounts() {
  console.info("fn.pressSyncAccounts");
  _syncAccounts();
  console.info("fn.pressSyncAccounts.success");
  return;
}

/**
 * Creates a new trigger to run the cron function every day.
 * A time is selected between 4 and 5 AM randomly by Google, it's then run at that time every time.
 */
function createCronTrigger() {
  try {
    console.info("fn.createCronTrigger");
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    const triggers = ScriptApp.getUserTriggers(ss);
    const trigger = triggers.find((t) => t.getHandlerFunction() === "cron");
    if (!trigger) {
      console.info("fn.createCronTrigger", "creating trigger");
      ScriptApp.newTrigger("cron").timeBased().atHour(4).everyDays(1).create();
      // ScriptApp.newTrigger("cron").timeBased().everyHours(1).create();
    } else {
      console.info("fn.createCronTrigger", "trigger already exists");
    }
    console.info("fn.createCronTrigger.success");
  } catch (error) {
    console.error("fn.createCronTrigger", error);
  }
}

/**
 * Deletes the cron trigger used for syncing data.
 * There's a limit on the number of triggers you can have per user and per sheet, so we delete ALL of them
 * (even though there should only be one) so we don't ever run into the limits.
 */
function disableCronTrigger() {
  try {
    console.info("fn.disableCronTrigger");
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    const triggers = ScriptApp.getUserTriggers(ss);
    triggers.forEach((t) => {
      const tid = t.getHandlerFunction();
      if (tid === "cron") {
        console.info("fn.disableCronTrigger", "deleting trigger", t);
        ScriptApp.deleteTrigger(t);
      }
    });
    console.info("fn.disableCronTrigger.success");
  } catch (error) {
    console.error("fn.disableCronTrigger", error);
  }
}

// Legacy
function dollaCron() {
  try {
    console.info("fn.disableDollaCronTrigger");
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    const triggers = ScriptApp.getUserTriggers(ss);
    triggers.forEach((t) => {
      const tid = t.getHandlerFunction();
      if (tid === "dollaCron") {
        console.info("fn.disableDollaCronTrigger", "deleting trigger", t);
        ScriptApp.deleteTrigger(t);
      }
    });
    console.info("fn.disableDollaCronTrigger.success");
  } catch (error) {
    console.error("fn.disableDollaCronTrigger", error);
  }
}
