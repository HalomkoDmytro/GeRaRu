async function downloadFileFromDrive(fileId, fileName) {

  const exportUrl = "https://docs.google.com/document/d/" + fileId + "/export?format=docx";

    const htmlOutput = HtmlService.createHtmlOutput(
    '<div style="text-align: center; font-family: Arial, sans-serif;">' +
      '<p>Файл створено</p>' +      
      '<a href="' +
      exportUrl +
      '" target="_blank" download ' +
      'style="display: inline-block; padding: 10px 20px; font-size: 14px; ' +
      "color: white; background-color: #34A853; border-radius: 5px; " +
      'text-decoration: none;">' +
      "Завантажити DOCX 😎" +
      "</a>" +
      "</div>",
  )
    .setWidth(300)
    .setHeight(200);

  SpreadsheetApp.getUi().showModalDialog(htmlOutput, ``);
}
