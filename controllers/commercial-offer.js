const ExcelJS = require('exceljs');
const path = require('path');
const fs = require('fs').promises;
const { fanDataDb, priceDb } = require('../utils/db');

const { NotFoundError } = require('../utils/errors/not-found-error');

const { MYSQL_FAN_DATABASE } = process.env;

// Экспортируем функцию обработки запроса

module.exports.getCommercialOffer = async (req, res, next) => {
  const templatePath = path.join(__dirname, '../template/commercial-offer.xlsx');
  const selectedData = req.body.historyItem;

  let outputPath;

  try {
    // Загружаем шаблон Excel-файла

    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(templatePath);

    // Получаем лист Excel
    const worksheet = workbook.getWorksheet('ТКП');

    // Начало вставки данных
    const startRow = 26;
    let item = 0; // Порядковый номер вставляемой системы

    // Считаем необходимое кол-во строк
    let currentRow = startRow;

    // Получаем данные из базы вентиляторов mySQL (по названиям опций)
    await Promise.all(selectedData.map(async (data) => {
      const optionsQuery = await fanDataDb.query(`
        SELECT ZRS, ZRSI, ZRN, ZRF, ZRC, ZRD, Regulator
        FROM ${MYSQL_FAN_DATABASE}.zfr_options
        WHERE model = ?
      `, [data.fanName]);

      const [priceDbData] = await priceDb.query(`
      SELECT *
      FROM Price
      WHERE Model IN (?, ?, ?, ?, ?, ?, ?, ?);
    `, [
        data.fanName,
        optionsQuery[0][0].ZRS,
        optionsQuery[0][0].ZRSI,
        optionsQuery[0][0].ZRN,
        optionsQuery[0][0].ZRF,
        optionsQuery[0][0].ZRC,
        optionsQuery[0][0].ZRD,
        optionsQuery[0][0].Regulator,
      ]);

      if (priceDbData.length === 0 || optionsQuery.length === 0) {
        throw new NotFoundError({ message: 'Не удалось найти данные в базе' });
      }

      // Функция добавления заголовка с названием системы и порядковым номером

      item += 1;

      const addHeader = () => {
        // Вставляем строки для одной единицы

        currentRow += 1;

        // worksheet.duplicateRow(currentRow, 1, true);

        worksheet.mergeCells(`B${currentRow}:E${currentRow}`);
        worksheet.getCell(`A${currentRow}`).value = {
          richText: [
            { text: `${item}`, font: { name: 'Arial', size: 11, bold: true } },
          ],
        };
        worksheet.getCell(`A${currentRow}`).alignment = { vertical: 'middle', horizontal: 'center' };

        worksheet.getCell(`B${currentRow}`).value = {
          richText: [
            { text: `${data.systemNameValue} (L=${data.flowRateValue}м3/ч, Рс=${data.staticPressureValue}Па)`, font: { name: 'Arial', size: 11, bold: true } },
          ],
        };
        worksheet.getRow(currentRow).height = 47;
      };

      addHeader();

      // функция добавления данных для элементов

      const addData = (priceData) => {
        // Объединяем ячейки для названия - пока что ломает код, если в ТКП больше одной системы
        // worksheet.mergeCells(`B${currentRow + 1}:E${currentRow + 1}`);

        // Вставляем следующие данные на строке
        worksheet.getCell(`B${currentRow + 1}`).value = priceData.ModelTKP;

        worksheet.getCell(`H${currentRow + 1}`).value = priceData.Price;
        worksheet.getCell(`H${currentRow + 1}`).numFmt = '#,##0.00';

        worksheet.getCell(`I${currentRow + 1}`).value = 0;
        worksheet.getCell(`I${currentRow + 1}`).numFmt = '#,#0.0%';

        worksheet.getCell(`L${currentRow + 1}`).value = {
          formula: `H${currentRow + 1}*(1-I${currentRow + 1})`,
          result: (priceData.Price * (1 - 0)),
        };

        worksheet.getCell(`M${currentRow + 1}`).value = 1;
        worksheet.getCell(`M${currentRow + 1}`).style.font = { name: 'Arial', size: 11, color: { argb: '0000FF' } };

        worksheet.getCell(`P${currentRow + 1}`).value = {
          formula: `L${currentRow + 1}*M${currentRow + 1}`,
          result: (priceData.Price),
        };

        worksheet.getCell(`Q${currentRow + 1}`).value = priceData.NSKod;

        // Центрирование контента для остальных ячеек на строке
        ['H', 'I', 'M', 'Q'].forEach((column) => {
          worksheet.getCell(`${column}${currentRow + 1}`).alignment = { vertical: 'middle', horizontal: 'center' };
        });

        currentRow += 1;
      };

      addData(priceDbData[0]);

      if (data.selectedOptions.selectFlatRoofSocket && priceDbData[1]) {
        // worksheet.duplicateRow(currentRow, 1, false);
        addData(priceDbData[1]);
        console.log(currentRow);
      }

      if (data.selectedOptions.selectFlatRoofSocketSilencer && priceDbData[2]) {
        // worksheet.duplicateRow(currentRow, 1, false);
        addData(priceDbData[2]);
        console.log(currentRow);
      }

      if (data.selectedOptions.selectSlantRoofSocketSilencer && priceDbData[3]) {
        // worksheet.duplicateRow(currentRow, 1, false);
        addData(priceDbData[3]);
        console.log(currentRow);
      }

      if (data.selectedOptions.selectBackDraftDamper && priceDbData[4]) {
        // worksheet.duplicateRow(currentRow, 1, false);
        addData(priceDbData[4]);
        console.log(currentRow);
      }

      if (data.selectedOptions.selectFlexibleConnector && priceDbData[5]) {
        // worksheet.duplicateRow(currentRow, 1, false);
        // console.log(priceDbData[5]);
        // console.log(`Гибкая вставка подобрана для ${data.systemNameValue}`);
        addData(priceDbData[5]);
        console.log(currentRow);
      }

      if (data.selectedOptions.selectFlange && priceDbData[6]) {
        // worksheet.duplicateRow(currentRow, 1, false);
        addData(priceDbData[6]);
        console.log(currentRow);
      }

      if (data.selectedOptions.selectRegulator && priceDbData[7]) {
        // worksheet.duplicateRow(currentRow, 1, false);
        // console.log(priceDbData[7]);
        // console.log(`Регулятор скорости подобран для ${data.systemNameValue}`);
        addData(priceDbData[7]);
        console.log(currentRow);
      }
    }));

    // Генерация уникальных имён файлов для предотвращения конфликтов при удалении

    const generateUniqueFileName = () => {
      const timestamp = new Date().getTime();
      return `newDataSheet_${timestamp}.xlsx`;
    };

    // Сохраняем результат в новый файл Excel

    outputPath = path.join(__dirname, `../uploads/${generateUniqueFileName()}`);
    await workbook.xlsx.writeFile(outputPath);
    // Читаем содержимое файла в бинарном формате
    const fileContent = await fs.readFile(outputPath, 'binary');
    // Устанавливаем заголовки Content-Type и Content-Disposition
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', 'attachment; filename=ZFR-Commercial.xlsx');
    // Передаем содержимое файла в ответ
    res.write(fileContent, 'binary');
    res.end();
  } catch (e) {
    next(e);
  } finally {
    // Перемещаем код удаления файла за пределы блока catch
    try {
      if (outputPath) {
        await fs.unlink(outputPath);
        console.log('Файл успешно удален');
      }
    } catch (unlinkError) {
      console.error('Ошибка удаления файла:', unlinkError);
    }
  }
};

// .finally(() => {
// // действия после завершения итрации
// worksheet.spliceRows(currentRow, 0, []);
// worksheet.getCell(`I${currentRow}`).fill = {
//   type: 'pattern',
//   pattern: 'solid',
//   fgColor: { argb: 'C5D9F1' },
//   bgColor: { argb: 'C5D9F1' },
// };
// worksheet.getCell(`M${currentRow}`).fill = {
//   type: 'pattern',
//   pattern: 'solid',
//   fgColor: { argb: 'C5D9F1' },
//   bgColor: { argb: 'C5D9F1' },
// };
// currentRow += 1;
// });
