/* Commands for ribbon buttons */

Office.onReady(() => {});

/**
 * Opens the taskpane (called from ribbon button).
 * @param {Office.AddinCommands.Event} event
 */
function showTaskpane(event) {
  Office.addin.showAsTaskpane();
  event.completed();
}

// Register with Office
Office.actions.associate("showTaskpane", showTaskpane);
