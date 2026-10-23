function myFunction() {
  SpreadsheetApp.getActive().getSheets().forEach(s =>
    s.getDataRange().getMergedRanges().forEach(r => r.breakApart())
  );
}