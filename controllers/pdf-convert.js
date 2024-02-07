const libre = require('libreoffice-convert');
const fs = require('fs').promises;
const path = require('path');

module.exports.convertToPdf = async (req, res, next) => {
  try {
    // Конвертирование XLSX в PDF
    const ext = '.pdf';
    const inputPath = path.join(__dirname, '../pdf-test/test.xlsx');
    const outputPath = path.join(__dirname, `../pdf-test/example${ext}`);

    let xlsxBuf;
    // Read file
    try {
      console.log(inputPath);
      xlsxBuf = await fs.readFile(inputPath);
      console.log('Файл успешно прочитан');
    } catch (readFileErr) {
      console.error('Ошибка при чтении файла:', readFileErr);
      return res.status(500).send('Ошибка при чтении файла');
    }

    // Convert it to pdf format with undefined filter (see Libreoffice docs about filter)
    libre.convert(xlsxBuf, ext, undefined, (convertErr, pdfBuf) => {
      if (convertErr) {
        console.error('Ошибка при конвертации файла:', convertErr);
        return res.status(500).send('Ошибка при конвертации файла');
      }
      // Здесь у вас есть буфер PDF, который вы можете записать в файл
      fs.writeFile(outputPath, pdfBuf)
        .then(() => {
          console.log('Файл успешно сконвертирован в PDF');
          res.status(200).send('Файл успешно сконвертирован в PDF');
        })
        .catch((writeFileErr) => {
          console.error('Ошибка при записи файла PDF:', writeFileErr);
          res.status(500).send('Ошибка при записи файла PDF');
        });
    });
  } catch (e) {
    console.error(`Ошибка: ${e}`);
    next(e);
  }
};
