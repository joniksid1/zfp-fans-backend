const http = require('http');
const fs = require('fs');
const axios = require('axios');
const FormData = require('form-data');

module.exports.getPdfFromXlsx = async (req, res) => {
  try {
    // Чтение Excel-файла с диска
    const excelBuffer = fs.readFileSync('./pdf-test/test.xlsx');

    // Создаем объект FormData и добавляем Excel-файл
    const formData = new FormData();
    formData.append('formFile', excelBuffer, {
      filename: 'generated_excel.xlsx',
      contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    });

    // Отправляем POST-запрос на микросервис с использованием Axios
    const response = await axios.post('http://192.168.97.98:443/Pdf', formData, {
      headers: {
        'accept': 'text/plain',
        ...formData.getHeaders(),
      },
      responseType: 'arraybuffer',
    });

    // Получаем PDF-файл в виде буфера
    const pdfBuffer = Buffer.from(response.data);

    // Сохраняем PDF-файл локально (для проверки)
    fs.writeFileSync('converted_file.pdf', pdfBuffer);

    // Отправляем PDF-файл клиенту для скачивания
    res.download('converted_file.pdf', 'converted_file.pdf', (err) => {
      if (err) {
        console.error(err);
        res.status(500).send('Внутренняя ошибка сервера');
      }

      // Удаляем временные файлы
      fs.unlinkSync('converted_file.pdf');
    });
  } catch (error) {
    console.error(error);
    res.status(500).send('Внутренняя ошибка сервера');
  }
};
