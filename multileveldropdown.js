function dropdown_battery() {
  var activeCell = SpreadsheetApp.getActiveSpreadsheet().getRange("Project Configuration!C11");
  var activeValue = activeCell.getValue();

  // Set the output range in C15, C16, C17, C18
  var outputRange = activeCell.offset(4, 0, 4, 1);

  // Check if the active sheet is "Project Configuration"
  if (activeCell.getSheet().getName() === "Project Configuration") {
    var configSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Config");
    var configDataRange = configSheet.getRange(2, 16, configSheet.getLastRow() - 1, 2);
    var configData = configDataRange.getValues();

    // Find the corresponding UAV model in the Config sheet
    var activeUAV = activeValue;
    var batteryOptions = [];

    for (var i = 0; i < configData.length; i++) {
      if (configData[i][0] === activeUAV) {
        batteryOptions.push(configData[i][1]);
      }
    }

    // Set the data validation rule for the dependent dropdown in C15, C16, C17, C18
    if (activeValue !== "") {
      if (batteryOptions.length > 0) {
        var rule = SpreadsheetApp.newDataValidation().requireValueInList(batteryOptions).build();
        outputRange.setDataValidation(rule);
      } else {
        // If no matching UAV found, clear the dependent dropdown in C15, C16, C17, C18
        outputRange.clearContent().clearDataValidations();
      }
    } else {
      // If input cell is blank, clear the dependent dropdown in C15, C16, C17, C18
      outputRange.clearContent().clearDataValidations();
    }
  }
}
