// ---------------------------------------
// OAuth related functions
// ---------------------------------------

function getAkahuOauthUrl() {
  console.info("fn.getAkahuOauthUrl");
  // Create a state token, which comes with a listener for the `/usercallback` endpoint.
  const _app_token = _getAppToken();
  const stateToken = ScriptApp.newStateToken()
    .withMethod("onAuth")
    .withTimeout(3600)
    .createToken();
  const centRedirect = "https://api.cent.nz/v1/auth";
  const url = `https://oauth.akahu.io?client_id=${_app_token}&response_type=code&redirect_uri=${encodeURIComponent(
    centRedirect
  )}&scope=ENDURING_CONSENT&state=${stateToken}`;

  console.info("fn.getAkahuOauthUrl.success");
  return url;
}

/**
 * This function is called when the user is redirected back to the sheet
 *
 * It should return some html to display in the OAuth pop up
 */
function onAuth(request) {
  console.info("fn.onAuth");
  const { error, error_description, token } = request.parameter;

  const userProperties = PropertiesService.getUserProperties();
  // send error to the sidebar
  if (error) {
    console.error(error);

    return HtmlService.createHtmlOutput(
      `<p>⚠️ ${error} ${error_description}</p>`
    );
  }
  console.info("fn.onAuth", "Setting properties");
  userProperties.setProperty(keys.action, "");
  userProperties.setProperty(keys.user_token, token);

  console.info("fn.onAuth.success");
  return HtmlService.createHtmlOutput(
    `<p>Thanks! You can <a onclick="window.close()">close this page</a> and go back to your spreadsheet now.</p>`
  );
}
