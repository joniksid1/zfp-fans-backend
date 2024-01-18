const express = require('express');
require('dotenv').config();
const { createConnection } = require('mysql2');
const { errors } = require('celebrate');
const cors = require('cors');
const { router } = require('./routes/root');
const { NotFoundError } = require('./utils/errors/not-found-error');
const error = require('./middlewares/error');
const { requestLogger, errorLogger } = require('./middlewares/logger');
const { getDataSheet } = require('./controllers/datasheet');

const { PORT = '3000', MYSQL_HOST, MYSQL_USER, MYSQL_PASSWORD, MYSQL_DATABASE } = process.env;

// Создайте соединение с MySQL
const db = createConnection({
  host: MYSQL_HOST,
  user: MYSQL_USER,
  password: MYSQL_PASSWORD,
  database: MYSQL_DATABASE,
});

// Подключитесь к базе данных
db.connect((err) => {
  if (err) {
    console.error('Ошибка подключения к базе данных MySQL:', err);
  } else {
    console.log('Подключено к базе данных MySQL');
  }
});

// db.query('SELECT * FROM zfr_dimensions',
//   (err, results, fields) => {
//     getDataSheet();
//   }
// );

const app = express();

app.use(cors({
  origin: ['http://localhost:3001', 'http://localhost:5173'],
  credentials: true,
  maxAge: 60,
}));

app.use(express.raw());

app.use(express.json());

app.use(requestLogger);

// Передайте соединение с MySQL вашим маршрутам или контроллерам
app.use('/', (req, res, next) => {
  req.db = db;
  next();
}, router);

app.use('*', () => {
  throw new NotFoundError({ message: 'Страница не найдена' });
});

app.use(errorLogger);

// Обработчик ошибок celebrate
app.use(errors());

// Централизованный middleware-обработчик
app.use(error);

app.listen(PORT, () => {
  console.log(`app is listening on port ${PORT}`);
});
