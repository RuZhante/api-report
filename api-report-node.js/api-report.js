import axios from "axios";
import ExcelJS from "exceljs";

async function fetchData() {
  try {
    const response = await axios.get("https://api.publicapis.org/entries");
    const data = response.data.entries.filter((entry) => entry.HTTPS === true);
    return data;
  } catch (error) {
    throw new Error(`Ошибка при запросе API: ${error.message}`);
  }
}

function formatHeaderRow(row) {
  row.fill = {
    type: "pattern",
    pattern: "solid",
    fgColor: { argb: "FFC4D79B" },
    bgColor: { argb: "FFC4D79B" },
  };
  row.border = {
    top: { style: "thin" },
    left: { style: "thin" },
    bottom: { style: "thin" },
    right: { style: "thin" },
  };
}

function formatDataRow(row) {
  row.eachCell((cell) => {
    cell.border = {
      top: { style: "thin" },
      left: { style: "thin" },
      bottom: { style: "thin" },
      right: { style: "thin" },
    };
  });
}

async function generateReport() {
  const data = await fetchData();

  // Создаем новую книгу Excel и лист с названием APIs
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet("APIs");

  // Добавляем заголовки в лист
  sheet.addRow([
    "API",
    "Description",
    "Auth",
    "HTTPS",
    "Cors",
    "Link",
    "Category",
  ]);
  formatHeaderRow(sheet.getRow(1));

  // Добавляем данные в лист, отсортированные по алфавиту по полю API
  data
    .sort((a, b) => a.API.localeCompare(b.API))
    .forEach((entry) => {
      sheet.addRow([
        entry.API,
        entry.Description,
        entry.Auth,
        entry.HTTPS.toString(),
        entry.Cors,
        entry.Link,
        entry.Category,
      ]);
    });

  // Делаем ссылки кликабельными
  sheet.getColumn("F").eachCell({ includeEmpty: true }, (cell) => {
    if (cell.value) {
      cell.value = { text: cell.value, hyperlink: cell.value };
    }
  });

  sheet.eachRow((row, rowNumber) => {
    if (rowNumber === 1) {
      formatHeaderRow(row);
    } else {
      formatDataRow(row);
    }
  });

  // Записываем книгу в файл Excel
  await workbook.xlsx.writeFile("api-report.xlsx");
  console.log("Отчет успешно сгенерирован");
}

generateReport();
