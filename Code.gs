function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Email")
    .addItem("Send Email", "sendEmails")
    .addToUi();
}

function sendEmails() {
  const configurationData = getConfigurationObject();
  for (let config of configurationData) {
    configureAndSendEmail(config);
  }
}

function configureAndSendEmail(config) {
  try {
    const {
      [CONSTANTS.KEYS.CONTENT]: content,
      [CONSTANTS.KEYS.SUBJECT]: subject,
      [CONSTANTS.KEYS.EMAIL]: emails,
      [CONSTANTS.KEYS.ATTACHMENTS]: attachmentIds,
      [CONSTANTS.KEYS.CC]: cc,
      [CONSTANTS.KEYS.BCC]: bcc,
      [CONSTANTS.KEYS.HTML_CONTENT]: htmlBody,
    } = config;
    const options = { cc, bcc, htmlBody };
    if (attachmentIds) {
      const attachments = attachmentIds
        .split(",")
        .map((id) => DriveApp.getFileById(id).getBlob());
      options.attachments = attachments;
    }
    const draft = GmailApp.createDraft(emails, subject, content, options);
    const sleepTimeout = getSleepTimeout();
    Utilities.sleep(sleepTimeout);
    draft.send();
  } catch (err) {
    Logger.log(err);
    logError(config, err.message);
  }
}

function getSleepTimeout() {
  const num = generateRandomNumber(CONSTANTS.RANDOM_TIMEOUT_RANGE_IN_S * 1000);
  const randomCoeff = generateRandomNumber();
  return randomCoeff * num;
}

function generateRandomNumber(range = 1) {
  const random = Math.random();
  return range * random;
}

function logError(errorConfig, error) {
  const activeSheet = SpreadsheetApp.getActive();
  let errorSheet = activeSheet.getSheetByName("Errors");
  if (!errorSheet) {
    errorSheet = activeSheet.insertSheet("Errors");
    const configurationSheet = activeSheet.getSheetByName(
      CONSTANTS.EMAIL_CONFIGURATION_SHEETNAME
    );
    const configurationSheetData = configurationSheet
      .getDataRange()
      .getValues();
    const headers = configurationSheetData[0];
    errorSheet.appendRow(headers);
  }
  errorSheet.appendRow([...errorConfig.row, error]);
}

function getConfigurationObject() {
  const configurationObject = [];
  const activeSheet = SpreadsheetApp.getActive();
  const configurationSheet = activeSheet.getSheetByName(
    CONSTANTS.EMAIL_CONFIGURATION_SHEETNAME
  );
  const configurationSheetData = configurationSheet.getDataRange().getValues();
  const headers = configurationSheetData[0];
  for (let idx = 1; idx < configurationSheetData.length; idx += 1) {
    const obj = {};
    const row = configurationSheetData[idx];
    for (const rowIdx in row) {
      obj[headers[rowIdx]] = row[rowIdx].length ? row[rowIdx] : undefined;
    }
    obj.row = row;
    configurationObject.push(obj);
  }

  return configurationObject;
}
