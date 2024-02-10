const ExcelJS = require('exceljs');
const path = require('path');
const fs = require('fs').promises;
const { fanDataDb, priceDb } = require('../utils/db');

const { NotFoundError } = require('../utils/errors/not-found-error');

const { MYSQL_FAN_DATABASE } = process.env;

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
    const startRow = 24;
    let item = 0; // Порядковый номер вставляемой системы

    // Считаем необходимое кол-во строк
    let currentRow = startRow;

    // Массив для хранения результатов SQL запросов
    const queryResults = await Promise.all(selectedData.map(async (data) => {
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

      return { data, priceDbData, optionsQuery };
    }));

    // Теперь, когда у нас есть все результаты запросов, мы можем использовать их в цикле map

    queryResults.forEach(({ data, priceDbData, optionsQuery }) => {
      // Увеличиваем порядковый номер вставляемой системы
      item += 1;

      // Функция добавления заголовка с названием системы и порядковым номером
      const addHeader = () => {
        currentRow += 2;

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
        worksheet.mergeCells(`B${currentRow + 1}:E${currentRow + 1}`);

        // Функция, которая форматирует текст и возвращает его в формате rich text
        const formatText = (text) => {
          // Регулярное выражение для поиска ключевых слов и всего текста после них,
          // не включая кириллические символы
          const regex = /(Zilon|MTY|IDS)[^\u0400-\u04FF]*/g;
          const formattedText = [];
          let match;
          let lastIndex = 0;

          // Поиск всех совпадений в тексте
          while ((match = regex.exec(text)) !== null) {
            const startIndex = match.index;
            const matchedText = match[0];

            // Добавление нежирной части текста перед найденным текстом
            if (startIndex > lastIndex) {
              formattedText.push({
                text: text.substring(lastIndex, startIndex),
                font: { name: 'Arial', size: 11, bold: false },
              });
            }

            // Добавление жирной части текста после найденных ключевых слов
            formattedText.push({
              text: matchedText,
              font: { name: 'Arial', size: 11, bold: true },
            });

            lastIndex = startIndex + matchedText.length; // Обновление индекса последнего символа
          }

          // Добавление оставшейся нежирной части текста после последнего найденного текста
          if (lastIndex < text.length) {
            formattedText.push({
              text: text.substring(lastIndex),
              font: { name: 'Arial', size: 11, bold: false },
            });
          }

          return { richText: formattedText };
        };

        const formattedText = formatText(priceData.ModelTKP);
        worksheet.getCell(`B${currentRow + 1}`).value = formattedText;

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
        ['H', 'I', 'M', 'Q'].forEach((column) => {
          worksheet.getCell(`${column}${currentRow + 1}`).alignment = { vertical: 'middle', horizontal: 'center' };
        });
        worksheet.getCell(`B${currentRow}`).alignment = { vertical: 'middle', wrapText: true };
        currentRow += 1;
      };

      const zfrData = priceDbData.find((i) => i.Model === data.fanName);

      if (zfrData) {
        addData(zfrData);
        worksheet.getCell(`B${currentRow - 1}`).style = { alignment: { horizontal: 'center', vertical: 'middle' } };
      }

      if (data.selectedOptions.selectFlatRoofSocket) {
        const { ZRS } = optionsQuery[0][0];
        const flatRoofSocketData = priceDbData.find((i) => i.Model === ZRS);
        if (flatRoofSocketData) {
          addData(flatRoofSocketData);
        }
      }

      if (data.selectedOptions.selectFlatRoofSocketSilencer) {
        const { ZRSI } = optionsQuery[0][0];
        const flatRoofSocketSilencerData = priceDbData.find((i) => i.Model === ZRSI);
        if (flatRoofSocketSilencerData) {
          addData(flatRoofSocketSilencerData);
        }
      }

      if (data.selectedOptions.selectSlantRoofSocketSilencer) {
        worksheet.getRow(currentRow + 1).height = 30;
        const { ZRN } = optionsQuery[0][0];
        const slantRoofSocketSilencerData = priceDbData.find((i) => i.Model === ZRN);
        if (slantRoofSocketSilencerData) {
          addData(slantRoofSocketSilencerData);
        }
      }

      if (data.selectedOptions.selectBackDraftDamper) {
        const { ZRD } = optionsQuery[0][0];
        const backDraftDamperData = priceDbData.find((i) => i.Model === ZRD);
        if (backDraftDamperData) {
          addData(backDraftDamperData);
        }
      }

      if (data.selectedOptions.selectFlexibleConnector) {
        const { ZRC } = optionsQuery[0][0];
        const flexibleConnectorData = priceDbData.find((i) => i.Model === ZRC);
        if (flexibleConnectorData) {
          addData(flexibleConnectorData);
        }
      }

      if (data.selectedOptions.selectFlange) {
        const { ZRF } = optionsQuery[0][0];
        const flangeData = priceDbData.find((i) => i.Model === ZRF);
        if (flangeData) {
          addData(flangeData);
        }
      }

      if (data.selectedOptions.selectRegulator) {
        currentRow += 1;
        worksheet.mergeCells(`B${currentRow}:E${currentRow}`);
        worksheet.getCell(`B${currentRow}`).value = {
          richText: [
            { text: 'Комплект обвязки и автоматики в составе', font: { name: 'Arial', size: 11, bold: true } },
          ],
        };
        const { Regulator } = optionsQuery[0][0];
        const regulatorData = priceDbData.find((i) => i.Model === Regulator);
        if (regulatorData) {
          addData(regulatorData);
        }

        // работает почему-то только если делать currentRow - 1, просто с currentRow не срабатывает
        // скорее всего не применяется, если там ещё нет контента
        worksheet.getCell(`B${currentRow - 1}`).style = { alignment: { horizontal: 'center', vertical: 'middle' } };
      }
    });

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
