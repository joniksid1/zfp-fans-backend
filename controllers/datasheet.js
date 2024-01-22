const ExcelJS = require('exceljs'); // Подключаем библиотеку для работы с Excel
const path = require('path'); // Подключаем модуль для работы с путями
const fs = require('fs').promises; // Используем промисифицированный fs для работы с файловой системой
const Jimp = require('jimp'); // Подключаем библиотеку для работы с изображениями
const NotFoundError = require('../utils/errors/not-found-error'); // Подключаем свою ошибку

// Экспортируем функцию обработки запроса
module.exports.getDataSheet = async (req, res) => {
  const templatePath = path.join(__dirname, '../template/data-worksheet.xlsx'); // Формируем путь к шаблону Excel
  const { db } = req; // Извлекаем объект базы данных из запроса
  const selectedData = req.body.historyItem; // Извлекаем необходимые данные из запроса

  // Создаем изображение из base64
  const buffer = Buffer.from(selectedData.plotImage.split(',')[1], 'base64');
  const jimpImage = await Jimp.read(buffer);

  // Получаем буфер изображения
  const imageBuffer = await jimpImage.getBufferAsync(Jimp.MIME_PNG);

  // Сохраняем изображение
  const imagePath = './uploads/image.png';
  await jimpImage.writeAsync(imagePath);

  let outputPath; // Объявляем переменную для пути к итоговому файлу

  try {
    // Используем ? для безопасного вставления значений в SQL-запрос
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

    // Проверяем, есть ли данные в ответе
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
    worksheet.getCell('F19').value = mySqlTechData.voltage_V;
    worksheet.getCell('F20').value = mySqlTechData.power_consumption_kW;
    worksheet.getCell('F21').value = mySqlTechData.max_operating_current_A;
    worksheet.getCell('F22').value = mySqlTechData.rotation_frequency_rpm;
    worksheet.getCell('F23').value = mySqlTechData.sound_power_level_dBA;

    // Заполняем данные из таблицы zfr_dimensions
    worksheet.getCell('F24').value = mySqlDimensionsData.kg;
    worksheet.getCell('B57').value = mySqlDimensionsData.l;
    worksheet.getCell('C57').value = mySqlDimensionsData.l1;
    worksheet.getCell('D57').value = mySqlDimensionsData.l2;
    worksheet.getCell('E57').value = mySqlDimensionsData.h;
    worksheet.getCell('F57').value = mySqlDimensionsData.d;
    worksheet.getCell('J10').value = mySqlDimensionsData.h;
    worksheet.getCell('J9').value = mySqlDimensionsData.l1;
    worksheet.getCell('J11').value = mySqlDimensionsData.l1;
    worksheet.getCell('J9').value = mySqlDimensionsData.l1;
    worksheet.getCell('G57').value = mySqlDimensionsData.l3;

    // Заполняем остальные поля
    const currentDate = new Date();
    worksheet.getCell('J5').value = currentDate;

    // Вставляем изображение в ячейку B28
    const imageId = workbook.addImage({
      buffer: imageBuffer,
      extension: 'png',
    });

    worksheet.addImage(imageId, {
      tl: { col: 1, row: 28 },
      br: { col: 8, row: 43 },
      editAs: 'oneCell',
    });

    // Сохраняем результат в новый файл Excel
    outputPath = path.join(__dirname, '../uploads/newDataSheet.xlsx');
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
    } catch (unlinkError) {
      console.error('Ошибка удаления файла:', unlinkError);
    }
  }
};
