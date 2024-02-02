const ExcelJS = require('exceljs');
const path = require('path');
const fs = require('fs').promises;
const { NotFoundError } = require('../utils/errors/not-found-error');

const { MYSQL_FAN_DATABASE } = process.env;

// Экспортируем функцию обработки запроса

module.exports.getCommercialOffer = async (req, res, next) => {
  const templatePath = path.join(__dirname, '../template/commercial-offer.xlsx');
  const { priceDb, fanDataDb } = req;
  const selectedData = req.body.historyItem;

  let outputPath;

  try {
    // Загружаем шаблон Excel-файла

    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(templatePath);

    // Получаем лист Excel
    const worksheet = workbook.getWorksheet('ТКП');

    // Получаем данные из базы вентиляторов mySQL (по названиям опций)
    await Promise.all(selectedData.map(async (data) => {
      const optionsQuery = await fanDataDb.promise().query(`
        SELECT ZRS, ZRSI, ZRN, ZRF, ZRC, ZRD, Regulator
        FROM ${MYSQL_FAN_DATABASE}.zfr_options
        WHERE model = ?
      `, [data.fanName]);

      const [priceDbData] = await priceDb.promise().query(`
        SELECT *
        FROM Price
        WHERE Model IN (?, ?, ?, ?, ?, ?, ?, ?)
      `, [
        data.fanName,
        optionsQuery.ZRS,
        optionsQuery.ZRSI,
        optionsQuery.ZRN,
        optionsQuery.ZRF,
        optionsQuery.ZRC,
        optionsQuery.ZRD,
        optionsQuery.Regulator,
      ]);

      if (priceDbData.length === 0 || optionsQuery.length === 0) {
        throw new NotFoundError({ message: 'Не удалось найти данные в базе' });
      }

      worksheet.addRow([
        data.systemNameValue || 0,
        data.fanName || 0,
        data.flowRateValue || 0,
        data.staticPressureValue || 0,
        priceDbData[0].ModelTKP || 0, // использование 0, если значение null или undefined
        priceDbData[0].Price || 0,
        priceDbData[0].NSKod || 0,
      ]);
    }));

    // Заполняем данные с фронтенда в ячейки
    // worksheet.getCell('E11').value = ;
    // Генерация уникальных имён файлов для предотвращения конфликтов при удалении

    const generateUniqueFileName = () => {
      const timestamp = new Date().getTime();
      return `newDataSheet_${timestamp}.xlsx`;
    };
    // let startRow = 57; // Начальная строка
    // if (selectedData.selectedOptions.selectFlatRoofSocket) {
    // }
    // if (selectedData.selectedOptions.selectFlatRoofSocketSilencer) {
    // }
    // if (selectedData.selectedOptions.selectSlantRoofSocketSilencer) {
    // }
    // if (selectedData.selectedOptions.selectFlexibleConnector) {
    // }
    // if (selectedData.selectedOptions.selectFlange) {
    // };
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
