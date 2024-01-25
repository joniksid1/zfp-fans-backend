const ExcelJS = require('exceljs');
const path = require('path');
const fs = require('fs').promises;
const Jimp = require('jimp');
const NotFoundError = require('../utils/errors/not-found-error');

// Экспортируем функцию обработки запроса

module.exports.getDataSheet = async (req, res) => {
  const templatePath = path.join(__dirname, '../template/data-worksheet.xlsx');
  const { db } = req;
  const selectedData = req.body.historyItem;

  // Создаем изображение из base64

  const buffer = Buffer.from(selectedData.plotImage.split(',')[1], 'base64');
  const jimpImage = await Jimp.read(buffer);

  // Получаем буфер изображения

  const imageBuffer = await jimpImage.getBufferAsync(Jimp.MIME_PNG);

  // Генерация уникального имени для изображения

  const generateUniqueImageFileName = () => {
    const timestamp = new Date().getTime();
    return `image_${timestamp}.png`;
  };

  // Формируем путь для сохранения изображения

  const imageOutputPath = path.join(__dirname, `../uploads/${generateUniqueImageFileName()}`);
  await jimpImage.writeAsync(imageOutputPath);

  let outputPath;

  try {
    // Получаем данные из базы mySQL

    const [techDataQuery] = await db.promise().query(`
      SELECT
        id,
        model,
        max_airflow_m3h,
        max_static_pressure_pa,
        voltage_V,
        power_consumption_kW,
        max_operating_current_A,
        rotation_frequency_rpm,
        sound_power_level_dBA,
        airflow_temperature_range,
        capacitor_mF,
        electrical_connections_scheme
      FROM zfrfans.zfr_data
      WHERE model = ?
    `, [selectedData.fanName]);

    const [dimensionsQuery] = await db.promise().query(`
      SELECT
        id,
        model,
        l,
        l1,
        l2,
        h,
        d,
        l3,
        kg
      FROM zfrfans.zfr_dimensions
      WHERE model = ?
    `, [selectedData.fanName]);

    // Проверка наличия данных в ответе

    if (techDataQuery.length === 0 || dimensionsQuery.length === 0) {
      throw new NotFoundError({ message: 'Не удалось найти данные в базе данных' });
    }

    // Извлекаем данные из ответов на SQL-запросы

    const mySqlTechData = techDataQuery[0];
    const mySqlDimensionsData = dimensionsQuery[0];

    // Загружаем шаблон Excel-файла

    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(templatePath);

    // Получаем лист Excel

    const worksheet = workbook.getWorksheet('Техника');

    // Заполняем данные с фронтенда в ячейки

    worksheet.getCell('E11').value = selectedData.systemNameValue;
    worksheet.getCell('E10').value = selectedData.staticPressureValue;
    worksheet.getCell('E9').value = selectedData.flowRateValue;

    // Заполняем данные из таблицы zfr_data

    worksheet.getCell('E12').value = mySqlTechData.model;
    worksheet.getCell('G19').value = mySqlTechData.voltage_V;
    worksheet.getCell('G20').value = mySqlTechData.power_consumption_kW;
    worksheet.getCell('G21').value = mySqlTechData.max_operating_current_A;
    worksheet.getCell('G22').value = mySqlTechData.rotation_frequency_rpm;
    worksheet.getCell('G23').value = mySqlTechData.sound_power_level_dBA;

    // Заполняем данные из таблицы zfr_dimensions

    worksheet.getCell('G24').value = mySqlDimensionsData.kg;
    worksheet.getCell('B55').value = mySqlDimensionsData.l;
    worksheet.getCell('C55').value = mySqlDimensionsData.l1;
    worksheet.getCell('D55').value = mySqlDimensionsData.l2;
    worksheet.getCell('E55').value = mySqlDimensionsData.h;
    worksheet.getCell('F55').value = mySqlDimensionsData.d;
    worksheet.getCell('J10').value = mySqlDimensionsData.h;
    worksheet.getCell('J9').value = mySqlDimensionsData.l1;
    worksheet.getCell('J11').value = mySqlDimensionsData.l1;
    worksheet.getCell('G55').value = mySqlDimensionsData.l3;

    // Заполняем остальные поля

    const currentDate = new Date();
    worksheet.getCell('J5').value = currentDate;

    // Вставляем изображение в ячейку B28

    const imageId = workbook.addImage({
      buffer: imageBuffer,
      extension: 'png',
    });

    worksheet.addImage(imageId, {
      tl: { col: 1, row: 25 },
      br: { col: 10, row: 42 },
      editAs: 'oneCell',
    });

    // Генерация уникальных имён файлов для предотвращения конфликтов при удалении

    const generateUniqueFileName = () => {
      const timestamp = new Date().getTime();
      return `newDataSheet_${timestamp}.xlsx`;
    };

    let startRow = 59; // Начальная строка для дополнительных опций
    const maxSecondSheetRows = 99; // Максимальное кол-во строк для вставки опции на 2-м листе
    const thirdSheetStartRows = 112; // Строка для начала контента третьего листа

    if (selectedData.selectedOptions.selectFlatRoofSocket) {
      // Добавляем данные изображений из относительного пути ../images/flat-roof-socket/...

      const imagePath1 = path.join(__dirname, '../images/flat-roof-socket/flat-roof-socket.jpg');
      const image1Buffer = await fs.readFile(imagePath1);
      const imageId1 = workbook.addImage({
        buffer: image1Buffer,
        extension: 'jpeg',
      });
      worksheet.addImage(imageId1, {
        tl: { col: 1, row: startRow + 2 },
        br: { col: 3, row: startRow + 8 },
        editAs: 'oneCell',
      });

      const imagePath2 = path.join(__dirname, '../images/flat-roof-socket/flat-roof-socket-dimensions.jpg');
      const image2Buffer = await fs.readFile(imagePath2);
      const imageId2 = workbook.addImage({
        buffer: image2Buffer,
        extension: 'jpeg',
      });
      worksheet.addImage(imageId2, {
        tl: { col: 4, row: startRow + 2 },
        br: { col: 8, row: startRow + 8 },
        editAs: 'oneCell',
      });
      startRow += 12;
      if (startRow >= maxSecondSheetRows && startRow <= thirdSheetStartRows) {
        worksheet.spliceRows(startRow, 0, [], [], [], [], [], []);
        startRow += 6;
      }
      console.log(`${startRow} после flat-roof-socket`);
    } else {
      // Удаляем строки, если нет опции

      for (let i = 0; i < 12; i += 1) {
        worksheet.spliceRows(startRow, 1);
      }
    }

    if (selectedData.selectedOptions.selectFlatRoofSocketSilencer) {
      // Добавляем данные изображений из относительного пути ../images/flat-roof-socket-silencer/...

      const imagePath1 = path.join(__dirname, '../images/flat-roof-socket-silencer/flat-roof-socket-silencer.jpg');
      const image1Buffer = await fs.readFile(imagePath1);
      const imageId1 = workbook.addImage({
        buffer: image1Buffer,
        extension: 'jpeg',
      });
      worksheet.addImage(imageId1, {
        tl: { col: 1, row: startRow + 2 },
        br: { col: 3, row: startRow + 8 },
        editAs: 'oneCell',
      });

      const imagePath2 = path.join(__dirname, '../images/flat-roof-socket-silencer/flat-roof-socket-silencer-dimensions.gif');
      const image2Buffer = await fs.readFile(imagePath2);
      const imageId2 = workbook.addImage({
        buffer: image2Buffer,
        extension: 'gif',
      });
      worksheet.addImage(imageId2, {
        tl: { col: 4, row: startRow + 2 },
        br: { col: 8, row: startRow + 8 },
        editAs: 'oneCell',
      });
      startRow += 12;
      if (startRow >= maxSecondSheetRows && startRow <= thirdSheetStartRows) {
        worksheet.spliceRows(startRow, 0, [], [], [], [], [], []);
        startRow += 6;
      }
      console.log(`${startRow} после flat-roof-socket-silencer`);
    } else {
      // Удаляем строки, если нет опции
      for (let i = 0; i < 12; i += 1) {
        worksheet.spliceRows(startRow, 1);
      }
    }

    if (selectedData.selectedOptions.selectSlantRoofSocketSilencer) {
      // Добавляем данные изображений из относительного пути
      // ../images/slant-roof-socket-silencer/...

      const imagePath1 = path.join(__dirname, '../images/slant-roof-socket-silencer/slant-roof-socket-silencer.png');
      const image1Buffer = await fs.readFile(imagePath1);
      const imageId1 = workbook.addImage({
        buffer: image1Buffer,
        extension: 'png',
      });
      worksheet.addImage(imageId1, {
        tl: { col: 1, row: startRow + 2 },
        br: { col: 3, row: startRow + 8 },
        editAs: 'oneCell',
      });

      const imagePath2 = path.join(__dirname, '../images/slant-roof-socket-silencer/slant-roof-socket-silencer-dimensions.gif');
      const image2Buffer = await fs.readFile(imagePath2);
      const imageId2 = workbook.addImage({
        buffer: image2Buffer,
        extension: 'gif',
      });
      worksheet.addImage(imageId2, {
        tl: { col: 4, row: startRow + 2 },
        br: { col: 8, row: startRow + 8 },
        editAs: 'oneCell',
      });
      startRow += 12;
      if (startRow >= maxSecondSheetRows && startRow <= thirdSheetStartRows) {
        worksheet.spliceRows(startRow, 0, [], [], [], [], [], []);
        startRow += 6;
      }
      console.log(`${startRow} после slant-roof-socket-silencer`);
    } else {
      // Удаляем строки, если нет опции
      for (let i = 0; i < 12; i += 1) {
        worksheet.spliceRows(startRow, 1);
      }
    }

    if (selectedData.selectedOptions.selectBackDraftDamper) {
      // Добавляем данные изображений из относительного пути
      // ../images/back-draft-damper/...

      const imagePath1 = path.join(__dirname, '../images/back-draft-damper/back-draft-damper.jpg');
      const image1Buffer = await fs.readFile(imagePath1);
      const imageId1 = workbook.addImage({
        buffer: image1Buffer,
        extension: 'jpg',
      });
      worksheet.addImage(imageId1, {
        tl: { col: 1, row: startRow + 2 },
        br: { col: 3, row: startRow + 8 },
        editAs: 'oneCell',
      });

      const imagePath2 = path.join(__dirname, '../images/back-draft-damper/back-draft-damper-dimensions.gif');
      const image2Buffer = await fs.readFile(imagePath2);
      const imageId2 = workbook.addImage({
        buffer: image2Buffer,
        extension: 'gif',
      });
      worksheet.addImage(imageId2, {
        tl: { col: 4, row: startRow + 2 },
        br: { col: 8, row: startRow + 8 },
        editAs: 'oneCell',
      });
      startRow += 12;
      if (startRow >= maxSecondSheetRows && startRow <= thirdSheetStartRows) {
        worksheet.spliceRows(startRow, 0, [], [], [], [], [], []);
        startRow += 6;
      }
      console.log(`${startRow} после back-draft-damper`);
    } else {
      // Удаляем строки, если нет опции
      for (let i = 0; i < 12; i += 1) {
        worksheet.spliceRows(startRow, 1);
      }
    }

    if (selectedData.selectedOptions.selectFlexibleConnector) {
      // Добавляем данные изображений из относительного пути
      // ../images/flexible-connector/...

      const imagePath1 = path.join(__dirname, '../images/flexible-connector/flexible-connector.jpg');
      const image1Buffer = await fs.readFile(imagePath1);
      const imageId1 = workbook.addImage({
        buffer: image1Buffer,
        extension: 'jpg',
      });
      worksheet.addImage(imageId1, {
        tl: { col: 1, row: startRow + 2 },
        br: { col: 3, row: startRow + 8 },
        editAs: 'oneCell',
      });
      const imagePath2 = path.join(__dirname, '../images/flexible-connector/flexible-connector-dimensions.gif');
      const image2Buffer = await fs.readFile(imagePath2);
      const imageId2 = workbook.addImage({
        buffer: image2Buffer,
        extension: 'gif',
      });
      worksheet.addImage(imageId2, {
        tl: { col: 4, row: startRow + 2 },
        br: { col: 8, row: startRow + 8 },
        editAs: 'oneCell',
      });
      startRow += 12;
      if (startRow >= maxSecondSheetRows && startRow <= thirdSheetStartRows) {
        worksheet.spliceRows(startRow, 0, [], [], [], [], [], []);
        startRow += 6;
      }
      console.log(`${startRow} после flexible-connector`);
    } else {
      // Удаляем строки, если нет опции
      for (let i = 0; i < 12; i += 1) {
        worksheet.spliceRows(startRow, 1);
      }
    }

    if (selectedData.selectedOptions.selectFlange) {
      // Добавляем данные изображений из относительного пути
      // ../images/flange/...

      const imagePath1 = path.join(__dirname, '../images/flange/flange.jpg');
      const image1Buffer = await fs.readFile(imagePath1);
      const imageId1 = workbook.addImage({
        buffer: image1Buffer,
        extension: 'jpg',
      });
      worksheet.addImage(imageId1, {
        tl: { col: 1, row: startRow + 2 },
        br: { col: 3, row: startRow + 8 },
        editAs: 'oneCell',
      });

      const imagePath2 = path.join(__dirname, '../images/flange/flange-dimensions.gif');
      const image2Buffer = await fs.readFile(imagePath2);
      const imageId2 = workbook.addImage({
        buffer: image2Buffer,
        extension: 'gif',
      });
      worksheet.addImage(imageId2, {
        tl: { col: 4, row: startRow + 2 },
        br: { col: 8, row: startRow + 8 },
        editAs: 'oneCell',
      });
      startRow += 12;
      if (startRow >= maxSecondSheetRows && startRow <= thirdSheetStartRows) {
        worksheet.spliceRows(startRow, 0, [], [], [], [], [], []);
        startRow += 6;
      }
      console.log(`${startRow} после flange`);
    } else {
      // Удаляем строки, если нет опции
      for (let i = 0; i < 12; i += 1) {
        worksheet.spliceRows(startRow, 1);
      }
    }

    // Сохраняем результат в новый файл Excel

    outputPath = path.join(__dirname, `../uploads/${generateUniqueFileName()}`);
    await workbook.xlsx.writeFile(outputPath);

    // Читаем содержимое файла в бинарном формате

    const fileContent = await fs.readFile(outputPath, 'binary');

    // Устанавливаем заголовки Content-Type и Content-Disposition

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', 'attachment; filename=ZFR-Datasheet.xlsx');

    // Передаем содержимое файла в ответ

    res.write(fileContent, 'binary');
    res.end();
  } catch (error) {
    console.error('Ошибка выполнения SQL-запроса:', error);
    res.status(500).send('Внутренняя ошибка сервера');
  } finally {
    // Перемещаем код удаления файла за пределы блока catch
    try {
      if (outputPath) {
        await fs.unlink(outputPath);
        console.log('Файл успешно удален');
      }
      if (imageOutputPath) {
        await fs.unlink(imageOutputPath);
        console.log('Изображение успешно удалено');
      }
    } catch (unlinkError) {
      console.error('Ошибка удаления файла:', unlinkError);
    }
  }
};
