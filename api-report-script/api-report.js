function generateReport() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  // Получаем данные из API
  var response = UrlFetchApp.fetch("https://api.publicapis.org/entries");
  var data = JSON.parse(response.getContentText()).entries;

  // Фильтруем данные, исключая объекты с HTTPS: false
  data = data.filter((entry) => entry.HTTPS);

  // Сортируем данные по имени API
  data.sort((a, b) => a.API.localeCompare(b.API));

  // Заполняем таблицу данными
  var headerRow = ["API", "Description", "Auth", "HTTPS", "Link", "Category"];
  sheet.appendRow(headerRow);
  var range = sheet.getRange(2, 1, data.length, headerRow.length);
  range.setValues(
    data.map((entry) => [
      entry.API,
      entry.Description,
      entry.Auth,
      entry.HTTPS,
      entry.Link,
      entry.Category,
    ])
  );

  // Добавляем форматирование таблицы
  range.setBorder(true, true, true, true, true, true);
  range.setBackground("#f0f8ff");
  range.setFontFamily("Arial");
  range.setFontSize(10);
  range.setHorizontalAlignment("left");
  range.setVerticalAlignment("middle");
  sheet.getRange("A1:F1").setBackground("#add8e6");
  sheet.getRange("A1:F1").setFontWeight("bold");
  sheet.autoResizeColumns(1, 6);
}
