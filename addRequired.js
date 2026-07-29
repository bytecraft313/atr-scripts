function addRequired() {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("survey");
  
    if (!sheet) {
      SpreadsheetApp.getUi().alert("Sheet 'survey' not found.");
      return;
    }
  
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
  
    const typeCol = headers.indexOf("type");
    const requiredCol = headers.indexOf("required");
    const requiredMsgCol = headers.indexOf("required message");
    const requiredMsgDariCol = headers.indexOf("required message:Dari");
    const requiredMsgPashtoCol = headers.indexOf("required message:Pashto");
  
    if (
      typeCol === -1 ||
      requiredCol === -1 ||
      requiredMsgCol === -1 ||
      requiredMsgDariCol === -1 ||
      requiredMsgPashtoCol === -1
    ) {
      SpreadsheetApp.getUi().alert(
        "Could not find one or more required columns:\n" +
        "type, required, required message, required message:Dari, required message:Pashto"
      );
      return;
    }
  
    for (let i = 1; i < data.length; i++) {
      const type = String(data[i][typeCol]).trim();
  
      const shouldRequire =
        type === "text" ||
        type === "decimal" ||
        type === "integer" ||
        type === "date" ||
        type === "audio" ||
        type === "image" ||
        type === "geopoint" ||
        type.startsWith("select_one") ||
        type.startsWith("select_multiple");
  
      if (shouldRequire) {
        if (!data[i][requiredCol]) {
          data[i][requiredCol] = "TRUE";
        }
        if (!data[i][requiredMsgCol]) {
          data[i][requiredMsgCol] = "The response is required";
        }
        if (!data[i][requiredMsgDariCol]) {
          data[i][requiredMsgDariCol] = "جواب حتمی است";
        }
        if (!data[i][requiredMsgPashtoCol]) {
          data[i][requiredMsgPashtoCol] = "ځواب حتمي دی";
        }
      }
    }
  
    sheet.getDataRange().setValues(data);
  
    SpreadsheetApp.getUi().alert("Required settings added successfully.");
  }