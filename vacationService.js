const VACATION_TABLE_ID = "1-mjP12xbjhKow9YGciKn3y8CN5IkCMmFrNfJ6z2EPpY";

// отримати номер відпускного з таблиці відпустко по імені
function getVacationNumber(pip) {
  const vacationSheet = SpreadsheetApp.openById(VACATION_TABLE_ID);
  const data = vacationSheet.getSheetByName("vocation").getRange('A1:N').getValues();

  for(let i = data.length-1; i >= 0; i--) {

    if(data[i][4] === pip) {
      return data[i][13];
    }
  }
}


