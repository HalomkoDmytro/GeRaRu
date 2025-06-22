function onOpen() {
  SpreadsheetApp.getUi().createMenu('🍄 Магія 🍄')
    .addItem('Очисти що наклацав 🧼🧽✨', 'clearSelected')
    .addItem('Створи 🧙‍♂️ мені рапорт 🙏🥺', 'generateRaport')
    .addToUi();
}


