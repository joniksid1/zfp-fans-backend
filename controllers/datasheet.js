const ExcelJS = require('exceljs');
const path = require('path');
const fs = require('fs');

module.exports.getDataSheet = async (req, res) => {
  const templatePath = path.join(__dirname, '../template/data-worksheet.xlsx');

  // Загружаем шаблон Excel-файла
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(templatePath);
  const worksheet = workbook.getWorksheet('Техника'); // Имя листа, надо проверить

  // Заполняем данные в ячейки
  worksheet.getCell('E9').value = 'blabla';

  // Сохраняем результат в новый файл Excel
  const outputPath = path.join(__dirname, '../downloads/newDataSheet.xlsx');

  workbook.xlsx.writeFile(outputPath)
    .then(() => {
      console.log('File is written');

      // Отправляем файл на фронтенд
      const file = fs.createReadStream(outputPath);

      // Устанавливаем заголовки Content-Type и Content-Disposition
      res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
      res.setHeader('Content-Disposition', 'attachment; filename=ZFR-Datasheet.xlsx');

      // Передача файла в ответ
      file.pipe(res);

      // Удаление файла после отправки
      file.on('close', () => {
        fs.unlink(outputPath, (unlinkErr) => {
          if (unlinkErr) {
            console.error('Error deleting file:', unlinkErr);
          } else {
            console.log('File deleted successfully');
          }
        });
      });
    })
    .catch((error) => {
      console.error(error.message);
      res.status(500).send('Internal Server Error');
    });
};
