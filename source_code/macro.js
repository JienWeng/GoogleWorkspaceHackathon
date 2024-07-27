function formatData() {
    var sheet = SpreadsheetApp.getActive();
    rangeF = sheet.getRange('F:F');
    rangeF.setNumberFormat('[$RM]* #,##0.00');
    rangeI = sheet.getRange('I:I');
    rangeI.setNumberFormat('d"-"mmm"-"yyyy');
    rangeJ = sheet.getRange('J:J');
    rangeJ.setNumberFormat('d"-"mmm"-"yyyy');
    rangeM = sheet.getRange('M:M');
    rangeM.setNumberFormat('d"-"mmm"-"yyyy');
  };