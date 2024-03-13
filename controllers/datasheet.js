const ExcelJS = require('exceljs');
const path = require('path');
const fs = require('fs').promises;
const Jimp = require('jimp');
const libre = require('libreoffice-convert');
const {
  getFanTechnicalData,
  getFanDimensionsData,
  getFanOptionsName,
  getSocketDimensionsData,
  getOtherOptionsDimensionsData,
} = require('../utils/database-query-service');
const { generateUniqueFileName } = require('../utils/file-name');

// Экспортируем функцию обработки запроса

module.exports.getDataSheet = async (req, res, next) => {
  const templatePath = path.join(__dirname, '../template/data-worksheet.xlsx');
  const selectedData = req.body.historyItem;

  // Создаем изображение из base64

  const buffer = Buffer.from(selectedData.plotImage.split(',')[1], 'base64');
  const jimpImage = await Jimp.read(buffer);

  // Получаем буфер изображения

  const imageBuffer = await jimpImage.getBufferAsync(Jimp.MIME_PNG);

  // Формируем путь для сохранения изображения

  const imageOutputPath = path.join(__dirname, `../uploads/${generateUniqueFileName()}`);
  await jimpImage.writeAsync(imageOutputPath);

  let outputPath;

  try {
    // Получаем данные из базы mySQL
    const fanTechData = await getFanTechnicalData(selectedData.fanName);
    const fanDimensionsData = await getFanDimensionsData(selectedData.fanName);
    const fanOptionsName = await getFanOptionsName(selectedData.fanName);
    const socketDimensions = await getSocketDimensionsData(fanOptionsName);
    const zrdZrcZrfDimensions = await getOtherOptionsDimensionsData(fanOptionsName);

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

    worksheet.getCell('E12').value = fanTechData.model;
    worksheet.getCell('G19').value = fanTechData.voltage_V;
    worksheet.getCell('G20').value = fanTechData.power_consumption_kW;
    worksheet.getCell('G21').value = fanTechData.max_operating_current_A;
    worksheet.getCell('G22').value = fanTechData.rotation_frequency_rpm;
    worksheet.getCell('G23').value = fanTechData.sound_power_level_dBA;

    // Заполняем данные из таблицы zfr_dimensions

    worksheet.getCell('G24').value = fanDimensionsData.kg;
    worksheet.getCell('B55').value = fanDimensionsData.l;
    worksheet.getCell('C55').value = fanDimensionsData.l1;
    worksheet.getCell('D55').value = fanDimensionsData.l2;
    worksheet.getCell('E55').value = fanDimensionsData.h;
    worksheet.getCell('F55').value = fanDimensionsData.d;
    worksheet.getCell('J10').value = fanDimensionsData.h;
    worksheet.getCell('J9').value = fanDimensionsData.l1;
    worksheet.getCell('J11').value = fanDimensionsData.l1;
    worksheet.getCell('G55').value = fanDimensionsData.l3;

    // Заполняем остальные поля

    const currentDate = new Date();
    const formattedDate = currentDate.toLocaleDateString('ru-RU');
    worksheet.getCell('J5').value = formattedDate;
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

    // Это добавляет разрыв страницы после указанной строки
    worksheet.getRow(55).addPageBreak();

    let totalWeight = parseFloat(fanDimensionsData.kg, 10);
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
      const zrsDimensions = socketDimensions.find((dimension) => dimension.Model.startsWith('ZRS'));
      const dataRow = startRow + 11;
      const optionWeigth = zrsDimensions.Weight_kg;
      worksheet.getCell(`B${dataRow}`).value = zrsDimensions.Hole_Spacing_D;
      worksheet.getCell(`C${dataRow}`).value = zrsDimensions.outer_socket_width_E;
      worksheet.getCell(`D${dataRow}`).value = zrsDimensions.Thread_Type_M;
      worksheet.getCell(`E${dataRow}`).value = zrsDimensions.inner_socket_width_G;
      worksheet.getCell(`F${dataRow}`).value = zrsDimensions.outer_platform_width_F;
      worksheet.getCell(`G${dataRow}`).value = zrsDimensions.height_H;
      worksheet.getCell(`H${dataRow}`).value = optionWeigth;
      totalWeight += parseFloat(optionWeigth, 10);
      startRow += 12;
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
      const zrsiDimensions = socketDimensions.find((dimension) => dimension.Model.startsWith('ZRSI'));
      const dataRow = startRow + 11;
      const optionWeigth = zrsiDimensions.Weight_kg;
      worksheet.getCell(`B${dataRow}`).value = zrsiDimensions.Hole_Spacing_D;
      worksheet.getCell(`C${dataRow}`).value = zrsiDimensions.outer_socket_width_E;
      worksheet.getCell(`D${dataRow}`).value = zrsiDimensions.Thread_Type_M;
      worksheet.getCell(`E${dataRow}`).value = zrsiDimensions.inner_socket_width_G;
      worksheet.getCell(`F${dataRow}`).value = zrsiDimensions.outer_platform_width_F;
      worksheet.getCell(`G${dataRow}`).value = zrsiDimensions.height_H;
      worksheet.getCell(`H${dataRow}`).value = optionWeigth;
      totalWeight += parseFloat(optionWeigth, 10);
      startRow += 12;
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
      const zrnDimensions = socketDimensions.find((dimension) => dimension.Model.startsWith('ZRN'));
      const dataRow = startRow + 11;
      const optionWeigth = zrnDimensions.Weight_kg;
      worksheet.getCell(`B${dataRow}`).value = zrnDimensions.Hole_Spacing_D;
      worksheet.getCell(`C${dataRow}`).value = zrnDimensions.outer_socket_width_E;
      worksheet.getCell(`D${dataRow}`).value = zrnDimensions.Thread_Type_M;
      worksheet.getCell(`E${dataRow}`).value = zrnDimensions.inner_socket_width_G;
      worksheet.getCell(`F${dataRow}`).value = zrnDimensions.outer_platform_width_F;
      worksheet.getCell(`G${dataRow}`).value = zrnDimensions.height_H;
      worksheet.getCell(`H${dataRow}`).value = optionWeigth;
      totalWeight += parseFloat(optionWeigth, 10);
      startRow += 12;
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
      const zrdDimensions = zrdZrcZrfDimensions.find((dimension) => dimension.Model.startsWith('ZRD'));
      const dataRow = startRow + 11;
      const optionWeigth = zrdDimensions.Weight_kg;
      worksheet.getCell(`B${dataRow}`).value = zrdDimensions.Middle_Diameter_e;
      worksheet.getCell(`C${dataRow}`).value = zrdDimensions.Diameter_D2;
      worksheet.getCell(`D${dataRow}`).value = zrdDimensions.Inner_Diameter_corrected_D;
      worksheet.getCell(`E${dataRow}`).value = zrdDimensions.Length_L;
      worksheet.getCell(`H${dataRow}`).value = optionWeigth;
      totalWeight += parseFloat(optionWeigth, 10);
      startRow += 12;
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
      const zrcDimensions = zrdZrcZrfDimensions.find((dimension) => dimension.Model.startsWith('ZRC'));
      const dataRow = startRow + 11;
      const optionWeigth = zrcDimensions.Weight_kg;
      worksheet.getCell(`B${dataRow}`).value = zrcDimensions.Inner_Diameter_d;
      worksheet.getCell(`C${dataRow}`).value = zrcDimensions.Middle_Diameter_e;
      worksheet.getCell(`D${dataRow}`).value = zrcDimensions.Inner_Diameter_corrected_D;
      worksheet.getCell(`H${dataRow}`).value = optionWeigth;
      totalWeight += parseFloat(optionWeigth, 10);
      startRow += 12;
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
      const zrfDimensions = zrdZrcZrfDimensions.find((dimension) => dimension.Model.startsWith('ZRF'));
      const dataRow = startRow + 11;
      const optionWeigth = zrfDimensions.Weight_kg;
      worksheet.getCell(`B${dataRow}`).value = zrfDimensions.Inner_Diameter_d;
      worksheet.getCell(`C${dataRow}`).value = zrfDimensions.Middle_Diameter_e;
      worksheet.getCell(`D${dataRow}`).value = zrfDimensions.Inner_Diameter_corrected_D;
      worksheet.getCell(`E${dataRow}`).value = zrfDimensions.Height_h;
      worksheet.getCell(`H${dataRow}`).value = optionWeigth;
      totalWeight += parseFloat(optionWeigth, 10);
      startRow += 12;
    } else {
      // Удаляем строки, если нет опции
      for (let i = 0; i < 12; i += 1) {
        worksheet.spliceRows(startRow, 1);
      }
    }
    // Перенос данных на третий лист, когда опций 3 (иначе не влезает на один лист с опциями схема)
    if (startRow >= maxInstallationSecondSheetRows && startRow <= minThirdSheetStartRows) {
      // Это добавляет разрыв страницы после указанной строки
      worksheet.getRow(startRow).addPageBreak();
    }

    // Перенос данных на третий лист, когда опций больше 3 (иначе делает его на одну строку раньше)
    if (startRow >= maxInstallationSecondSheetRows && startRow >= minThirdSheetStartRows) {
      // Это добавляет разрыв страницы после указанной строки
      worksheet.getRow(104).addPageBreak();
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

    const xlsxBuf = await fs.readFile(outputPath);

    // Конвертируем в PDF - моим методом
    libre.convert(xlsxBuf, '.pdf', undefined, async (convertErr, pdfBuf) => {
      if (convertErr) {
        next(convertErr);
        return;
      }

      try {
        // Сохраняем PDF на диск
        const pdfOutputPath = path.join(__dirname, `../uploads/${generateUniqueFileName()}.pdf`);
        await fs.writeFile(pdfOutputPath, pdfBuf);

        // Устанавливаем заголовки Content-Type и Content-Disposition для PDF
        res.setHeader('Content-Type', 'application/pdf');
        res.setHeader('Content-Disposition', 'attachment; filename=ZFR-Datasheet.pdf');

        // Передаем содержимое файла PDF в ответ
        res.write(pdfBuf, 'binary');
        res.end();

        try {
          await fs.unlink(pdfOutputPath);
        } catch (unlinkPdfError) {
          next(unlinkPdfError);
        }
      } catch (writeFileErr) {
        next(writeFileErr);
      }
    });
  } catch (e) {
    next(e);
  } finally {
    // Перемещаем код удаления файла за пределы блока catch
    try {
      if (outputPath) {
        await fs.unlink(outputPath);
      }
      if (imageOutputPath) {
        await fs.unlink(imageOutputPath);
      }
    } catch (unlinkError) {
      next(unlinkError);
    }
  }
};
