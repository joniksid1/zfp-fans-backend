const ExcelJS = require('exceljs');
const path = require('path');
const fs = require('fs').promises;
const Jimp = require('jimp');
const { fanDataDb } = require('../utils/db');
const { NotFoundError } = require('../utils/errors/not-found-error');

const { MYSQL_FAN_DATABASE } = process.env;

// Экспортируем функцию обработки запроса

module.exports.getDataSheet = async (req, res, next) => {
  const templatePath = path.join(__dirname, '../template/data-worksheet.xlsx');
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

    const [techDataQuery] = await fanDataDb.query(`
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
      FROM ${MYSQL_FAN_DATABASE}.zfr_data
      WHERE model = ?
    `, [selectedData.fanName]);

    const [dimensionsQuery] = await fanDataDb.query(`
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
      FROM ${MYSQL_FAN_DATABASE}.zfr_dimensions
      WHERE model = ?
    `, [selectedData.fanName]);

    const [optionsQuery] = await fanDataDb.query(`
      SELECT ZRS, ZRSI, ZRN, ZRF, ZRC, ZRD
      FROM ${MYSQL_FAN_DATABASE}.zfr_options
      WHERE model = ?
    `, [selectedData.fanName]);

    // Проверка наличия данных в ответе

    if (techDataQuery.length === 0) {
      throw new NotFoundError({ message: 'Не удалось найти данные технических характеристик вентиляторов в базе' });
    }

    if (dimensionsQuery.length === 0) {
      throw new NotFoundError({ message: 'Не удалось найти данные размеров вентиляторов в базе' });
    }

    if (optionsQuery.length === 0) {
      throw new NotFoundError({ message: 'Не удалось найти названия опций в базе' });
    }

    // Извлекаем данные из ответов на SQL-запросы

    const mySqlTechData = techDataQuery[0];
    const mySqlDimensionsData = dimensionsQuery[0];
    const fanOptions = optionsQuery[0];

    // Делаем запрос на данные опций для выбранного вентилятора

    const [socketDimensionsQuery] = await fanDataDb.query(`
    SELECT
      id,
      TypeSize,
      Model,
      Hole_Spacing_D,
      outer_socket_width_E,
      Thread_Type_M,
      inner_socket_width_G,
      outer_platform_width_F,
      height_H,
      Weight_kg
    FROM ${MYSQL_FAN_DATABASE}.zrs_zrsi_zrn_dimensions
    WHERE Model IN (?, ?, ?)
    `, [fanOptions.ZRS, fanOptions.ZRSI, fanOptions.ZRN]);

    const [zrdZrcZrfDimensionsQuery] = await fanDataDb.query(`
    SELECT
      id,
      TypeSize,
      Model,
      Inner_Diameter_d,
      Middle_Diameter_e,
      Inner_Diameter_corrected_D,
      Height_h,
      Length_L,
      Diameter_D2,
      Weight_kg
    FROM ${MYSQL_FAN_DATABASE}.zrd_zrc_zrf_dimensions
    WHERE Model IN (?, ?, ?)
    `, [fanOptions.ZRD, fanOptions.ZRC, fanOptions.ZRF]);

    // Проверка наличия данных в ответе
    if (socketDimensionsQuery.length === 0) {
      throw new NotFoundError({ message: 'Данные по монтажным стаканам не найдены в базе' });
    }

    if (zrdZrcZrfDimensionsQuery.length === 0) {
      throw new NotFoundError({ message: 'Данные по опциям "фланец", "гибкая вставка", "обратный клапан" не найдены в базе' });
    }

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
    worksheet.getCell('J4').value = selectedData.projectNameValue;

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

    let totalWeight = mySqlDimensionsData.kg;
    let startRow = 57; // Начальная строка для дополнительных опций
    const maxInstallationSecondSheetRows = 82; // Максимальное кол-во строк для вставки схемы
    // на 2-м листе
    const minThirdSheetStartRows = 104; // Строка для начала контента третьего листа

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
      // Добавляем данные об опции
      const zrsDimensions = socketDimensionsQuery.find((dimension) => dimension.Model.startsWith('ZRS'));
      const dataRow = startRow + 11;
      const optionWeigth = Math.round(zrsDimensions.Weight_kg);
      worksheet.getCell(`B${dataRow}`).value = zrsDimensions.Hole_Spacing_D;
      worksheet.getCell(`C${dataRow}`).value = zrsDimensions.outer_socket_width_E;
      worksheet.getCell(`D${dataRow}`).value = zrsDimensions.Thread_Type_M;
      worksheet.getCell(`E${dataRow}`).value = zrsDimensions.inner_socket_width_G;
      worksheet.getCell(`F${dataRow}`).value = zrsDimensions.outer_platform_width_F;
      worksheet.getCell(`G${dataRow}`).value = zrsDimensions.height_H;
      worksheet.getCell(`H${dataRow}`).value = optionWeigth;
      totalWeight += optionWeigth;
      startRow += 12;
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
      // Добавляем данные об опции
      const zrsiDimensions = socketDimensionsQuery.find((dimension) => dimension.Model.startsWith('ZRSI'));
      const dataRow = startRow + 11;
      const optionWeigth = Math.round(zrsiDimensions.Weight_kg);
      worksheet.getCell(`B${dataRow}`).value = zrsiDimensions.Hole_Spacing_D;
      worksheet.getCell(`C${dataRow}`).value = zrsiDimensions.outer_socket_width_E;
      worksheet.getCell(`D${dataRow}`).value = zrsiDimensions.Thread_Type_M;
      worksheet.getCell(`E${dataRow}`).value = zrsiDimensions.inner_socket_width_G;
      worksheet.getCell(`F${dataRow}`).value = zrsiDimensions.outer_platform_width_F;
      worksheet.getCell(`G${dataRow}`).value = zrsiDimensions.height_H;
      worksheet.getCell(`H${dataRow}`).value = optionWeigth;
      totalWeight += optionWeigth;
      startRow += 12;
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

      const imagePath2 = path.join(__dirname, '../images/slant-roof-socket-silencer/slant-roof-socket-silencer-dimensions.png');
      const image2Buffer = await fs.readFile(imagePath2);
      const imageId2 = workbook.addImage({
        buffer: image2Buffer,
        extension: 'png',
      });
      worksheet.addImage(imageId2, {
        tl: { col: 4, row: startRow + 2 },
        br: { col: 8, row: startRow + 8 },
        editAs: 'oneCell',
      });
      // Добавляем данные об опции
      const zrnDimensions = socketDimensionsQuery.find((dimension) => dimension.Model.startsWith('ZRN'));
      const dataRow = startRow + 11;
      const optionWeigth = Math.round(zrnDimensions.Weight_kg);
      worksheet.getCell(`B${dataRow}`).value = zrnDimensions.Hole_Spacing_D;
      worksheet.getCell(`C${dataRow}`).value = zrnDimensions.outer_socket_width_E;
      worksheet.getCell(`D${dataRow}`).value = zrnDimensions.Thread_Type_M;
      worksheet.getCell(`E${dataRow}`).value = zrnDimensions.inner_socket_width_G;
      worksheet.getCell(`F${dataRow}`).value = zrnDimensions.outer_platform_width_F;
      worksheet.getCell(`G${dataRow}`).value = zrnDimensions.height_H;
      worksheet.getCell(`H${dataRow}`).value = optionWeigth;
      totalWeight += optionWeigth;
      startRow += 12;
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
      // Добавляем данные об опции
      const zrdDimensions = zrdZrcZrfDimensionsQuery.find((dimension) => dimension.Model.startsWith('ZRD'));
      const dataRow = startRow + 11;
      const optionWeigth = Math.round(zrdDimensions.Weight_kg);
      worksheet.getCell(`B${dataRow}`).value = zrdDimensions.Middle_Diameter_e;
      worksheet.getCell(`C${dataRow}`).value = zrdDimensions.Diameter_D2;
      worksheet.getCell(`D${dataRow}`).value = zrdDimensions.Inner_Diameter_corrected_D;
      worksheet.getCell(`E${dataRow}`).value = zrdDimensions.Length_L;
      worksheet.getCell(`H${dataRow}`).value = optionWeigth;
      totalWeight += optionWeigth;
      startRow += 12;
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
      // Добавляем данные об опции
      const zrcDimensions = zrdZrcZrfDimensionsQuery.find((dimension) => dimension.Model.startsWith('ZRC'));
      const dataRow = startRow + 11;
      const optionWeigth = Math.round(zrcDimensions.Weight_kg);
      worksheet.getCell(`B${dataRow}`).value = zrcDimensions.Inner_Diameter_d;
      worksheet.getCell(`C${dataRow}`).value = zrcDimensions.Middle_Diameter_e;
      worksheet.getCell(`D${dataRow}`).value = zrcDimensions.Inner_Diameter_corrected_D;
      worksheet.getCell(`H${dataRow}`).value = optionWeigth;
      totalWeight += optionWeigth;
      startRow += 12;
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
      // Добавляем данные об опции
      const zrfDimensions = zrdZrcZrfDimensionsQuery.find((dimension) => dimension.Model.startsWith('ZRF'));
      const dataRow = startRow + 11;
      const optionWeigth = Math.round(zrfDimensions.Weight_kg);
      worksheet.getCell(`B${dataRow}`).value = zrfDimensions.Inner_Diameter_d;
      worksheet.getCell(`C${dataRow}`).value = zrfDimensions.Middle_Diameter_e;
      worksheet.getCell(`D${dataRow}`).value = zrfDimensions.Inner_Diameter_corrected_D;
      worksheet.getCell(`E${dataRow}`).value = zrfDimensions.Height_h;
      worksheet.getCell(`H${dataRow}`).value = optionWeigth;
      totalWeight += optionWeigth;
      startRow += 12;
      console.log(`${startRow} после flange`);
    } else {
      // Удаляем строки, если нет опции
      for (let i = 0; i < 12; i += 1) {
        worksheet.spliceRows(startRow, 1);
      }
    }
    // Перенос данных на третий лист, когда опций 3 (иначе не влезает на один лист с опциями схема)
    if (startRow >= maxInstallationSecondSheetRows && startRow <= minThirdSheetStartRows) {
      worksheet.spliceRows(startRow, 0, [], [], [], [], [], [], [], [], [], [], [], []);
      startRow += 12;
    }

    // Добавляем общую массу с учётом опций

    worksheet.getCell('J12').value = totalWeight;

    // Объединяем надпись "схемы" в одну ячейку - Баг библиотеки с разбивкой ячейки
    worksheet.mergeCells(`A${startRow + 1}:K${startRow + 1}`);

    const installationImagePath = path.join(__dirname, '../images/installation.png');
    const installationImageBuffer = await fs.readFile(installationImagePath);
    const installationImageId = workbook.addImage({
      buffer: installationImageBuffer,
      extension: 'png',
    });
    worksheet.addImage(installationImageId, {
      tl: { col: 1, row: startRow + 2 },
      br: { col: 5, row: startRow + 25 },
      editAs: 'oneCell',
    });

    // Удаляем заголовок "опции", если опций нет
    if (
      !selectedData.selectedOptions.selectFlatRoofSocket
      && !selectedData.selectedOptions.selectFlatRoofSocketSilencer
      && !selectedData.selectedOptions.selectSlantRoofSocketSilencer
      && !selectedData.selectedOptions.selectBackDraftDamper
      && !selectedData.selectedOptions.selectFlexibleConnector
      && !selectedData.selectedOptions.selectFlange
    ) {
      // Объединяем надпись "схемы" в одну ячейку - Баг библиотеки с разбивкой ячейки
      worksheet.spliceRows(startRow - 1, 1);
      worksheet.mergeCells(`A${startRow}:K${startRow}`);
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
  } catch (e) {
    next(e);
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
