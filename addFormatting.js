function addFormatting() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("survey"); // Change if needed

  const data = sheet.getDataRange().getValues();

  if (data.length < 2) return;

  // Find column indexes
  const headers = data[0].map(h => String(h).trim());

Logger.log(headers);

const typeCol = headers.indexOf("type");
const dariCol = headers.indexOf("label:Dari");
const pashtoCol = headers.indexOf("label:Pashto");
const englishCol = headers.indexOf("label:English");

  if (typeCol === -1 || dariCol === -1 || pashtoCol === -1 || englishCol === -1) {
    SpreadsheetApp.getUi().alert(
      "Could not find one or more required columns:\n" +
      "type, label:Dari, label:Pashto, label:English"
    );
    return;
  }

  const prefix = "<p dir='rtl'><b>";
  const englishPrefix = "<b>";

  for (let i = 1; i < data.length; i++) {
    const type = String(data[i][typeCol]).trim();

    const matches =
      type === "text" ||
      type === "decimal" ||
      type === "integer" ||
      type === "date" ||
      type === "note" ||
      type === "audio" ||
      type === "image" ||
      type === "geopoint" ||
      type.startsWith("select_one") ||
      type.startsWith("select_multiple");

    if (!matches) continue;

    // Dari
    let dari = String(data[i][dariCol] || "");
    if (dari && !dari.trim().startsWith(prefix)) {
      data[i][dariCol] = prefix + dari;
    }

    // Pashto
    let pashto = String(data[i][pashtoCol] || "");
    if (pashto && !pashto.trim().startsWith(prefix)) {
      data[i][pashtoCol] = prefix + pashto;
    }

    // English
    let english = String(data[i][englishCol] || "");
    if (english && !english.trim().startsWith(englishPrefix)) {
      data[i][englishCol] = englishPrefix + english;
    }
  }



  // Write everything back in one operation
  sheet.getDataRange().setValues(data);

  SpreadsheetApp.getUi().alert("RTL formatting added where needed.");
}