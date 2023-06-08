const { google } = require("googleapis");
const keys = require("./keys.json");

// Подключение к Google Sheets API с помощью сервисного аккаунта
const client = new google.auth.JWT(keys.client_email, null, keys.private_key, [
  "https://www.googleapis.com/auth/spreadsheets",
]);

client.authorize((err, tokens) => {
  if (err) {
    console.error(err);
    return;
  }
  console.log("Successfully connected to Google Sheets API!");
});

// Функция для создания отчета в Google Sheets
async function createSheetReport() {
  const sheets = google.sheets({ version: "v4", auth: client });

  try {
    // Создаем новый лист с названием "API Report"
    const res = await sheets.spreadsheets.create({
      resource: {
        properties: {
          title: "API Report",
        },
      },
    });
    const spreadsheetId = res.data.spreadsheetId;

    // Получаем список листов в созданной таблице
    const sheetRes = await sheets.spreadsheets.get({
      spreadsheetId,
      includeGridData: true,
    });
    const sheetId = sheetRes.data.sheets[0].properties.sheetId;

    // Заполняем заголовок таблицы
    await sheets.spreadsheets.batchUpdate({
      spreadsheetId,
      resource: {
        requests: [
          // Обновляем свойства листа
          {
            updateSheetProperties: {
              properties: {
                title: "API Report",
                sheetId,
                gridProperties: {
                  rowCount: 1,
                  columnCount: 7,
                  frozenRowCount: 1,
                },
              },
              fields:
                "title,gridProperties(rowCount,columnCount,frozenRowCount)",
            },
          },
          // Устанавливаем форматирование для шапки таблицы
          {
            repeatCell: {
              range: {
                sheetId,
                startRowIndex: 0,
                endRowIndex: 1,
              },
              cell: {
                userEnteredFormat: {
                  textFormat: { bold: true },
                  backgroundColor: { red: 0.8, green: 0.8, blue: 0.8 },
                },
              },
              fields: "userEnteredFormat(textFormat,backgroundColor)",
            },
          },
          // Заполняем шапку таблицы значениями
          {
            updateCells: {
              rows: [
                {
                  values: [
                    { userEnteredValue: { stringValue: "API" } },
                    { userEnteredValue: { stringValue: "Description" } },
                    { userEnteredValue: { stringValue: "Auth" } },
                    { userEnteredValue: { stringValue: "HTTPS" } },
                    { userEnteredValue: { stringValue: "Cors" } },
                    { userEnteredValue: { stringValue: "Link" } },
                    { userEnteredValue: { stringValue: "Category" } },
                  ],
                },
              ],
              start: { sheetId, rowIndex: 0, columnIndex: 0 },
              fields: "*",
            },
          },
        ],
      },
    });

    // Запрашиваем данные из API
    const axios = require("axios");
    const entries = (await axios.get("https://api.publicapis.org/entries")).data
      .entries;

    // Фильтруем записи с HTTPS = false и сортируем по названию API
    const filteredEntries = entries.filter((entry) => entry.HTTPS !== false);
    const sortedEntries = filteredEntries.sort((a, b) =>
      a.API.localeCompare(b.API)
    );

    // Преобразуем записи в формат таблицы Google Sheets
    const data = sortedEntries.map((entry) => [
      { userEnteredValue: { stringValue: entry.API } },
      { userEnteredValue: { stringValue: entry.Description } },
      { userEnteredValue: { stringValue: entry.Auth } },
      { userEnteredValue: { boolValue: entry.HTTPS } },
      { userEnteredValue: { stringValue: entry.Cors } },
      {
        userEnteredValue: {
          hyperlink: entry.Link,
          formulaValue: `"${entry.Link}"`,
        },
      },
      { userEnteredValue: { stringValue: entry.Category } },
    ]);

    // Заполняем таблицу данными из API
    await sheets.spreadsheets.values.update({
      spreadsheetId,
      range: "A2:G",
      valueInputOption: "USER_ENTERED",
      resource: { values: data },
    });

    console.log("Successfully created and populated the Google Sheets report!");
  } catch (error) {
    console.error(error);
  }
}

createSheetReport();
