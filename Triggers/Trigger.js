
/**
 * Yearly Trigger
 * Checks for trigger. Only creates trigger with unique name
 * Uses Installable trigger to automate the trigger
 * Click here to learn more https://developers.google.com/apps-script/guides/triggers/installable
 * @params {string} triggerName : Name of the trigger
 */
function yearlyTrigger(triggerName) {
  var scriptArray = ScriptApp.getProjectTriggers();
  if (scriptArray.find(script => script.getHandlerFunction() === triggerName))
    return;
  let d = dateHelper.getYearEnd();
  ScriptApp
    .newTrigger(triggerName)
    .timeBased()
    .at(d)
    .create();
}


/**
 * Check for particular spreadsheet and constantly listen when spreadsheet is open
 * Checks for trigger. Only creates trigger with unique name
 * Uses Installable trigger to automate the trigger
 * Click here to learn more https://developers.google.com/apps-script/guides/triggers/installable
 * @params {Spreadsheet} ss : Targeted Spreadsheet
 * @params {string} triggerName : Name of the trigger
 */
function onOpenSpreadsheetTrigger(ss,triggerName) {
  var triggers = ScriptApp.getProjectTriggers();
  if (!triggers.find(trigger => (trigger.getTriggerSourceId() == ss.getId()) && trigger.getHandlerFunction() === triggerName)) {
  console.log("Inserting Trigger");
  ScriptApp.newTrigger(triggerName)
    .forSpreadsheet(ss)
    .onOpen()
    .create();
  }
}

