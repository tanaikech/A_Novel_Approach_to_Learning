/**
 * ### Description
 * Create a custom menu when Spreadsheet is opened.
 *
 * @return {void}
 */
function onOpen() {
  SpreadsheetApp.getUi().createMenu("ManagementLearning")
    .addItem("Run", "main_ManagementLearning")
    .addItem("Reset", "reset_ManagementLearning")
    .addToUi();
}


/**
 * ### Description
 * Main method.
 *
 * @param {Object} object Event object of installable OnSubmit trigger on Google Forms.
 * @param {SpreadsheetApp.Range} object.range
 * @param {FormApp.Form} object.source
 * @param {FormApp.FormResponse} object.response
 * @param {String} object.triggerUid
 * @param {String} object.authMode
 * 
 * @return {void}
 *
 */
function main_ManagementLearning(e) {
  const m = new ManagementLearning();
  try {
    m.run(e);
  } catch (err) {
    console.log(err.stack);
  }
}

/**
 * ### Description
 * Reset.
 *
 * @return {void}
 */
function reset_ManagementLearning() {
  const m = new ManagementLearning();
  try {
    m.reset();
  } catch (err) {
    console.log(err.stack);
  }
}
