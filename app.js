const express = require('express');
require('dotenv').config();
const { createConnection } = require('mysql2');
const { errors } = require('celebrate');
const cors = require('cors');
const { router } = require('./routes/root');
const { NotFoundError } = require('./utils/errors/not-found-error');
const error = require('./middlewares/error');
const { requestLogger, errorLogger } = require('./middlewares/logger');

const {
  PORT = '3000',
  MYSQL_HOST,
  MYSQL_USER,
  MYSQL_PASSWORD,
  MYSQL_FAN_DATABASE,
  MYSQL_PRICE_DATABASE,
} = process.env;

// Подключение MySQL к БД вентиляторов
const fanDataDb = createConnection({
  host: MYSQL_HOST,
  user: MYSQL_USER,
  password: MYSQL_PASSWORD,
  database: MYSQL_FAN_DATABASE,
});

fanDataDb.connect((err) => {
  if (err) {
    console.error('Ошибка подключения к базе данных вентиляторов MySQL:', err);
  } else {
    console.log('Подключено к базе данных вентиляторов MySQL');
  }
});

// Подключение MySQL к БД цен
const priceDb = createConnection({
  host: MYSQL_HOST,
  user: MYSQL_USER,
  password: MYSQL_PASSWORD,
  database: MYSQL_PRICE_DATABASE,
});

priceDb.connect((err) => {
  if (err) {
    console.error('Ошибка подключения к базе данных цен MySQL:', err);
  } else {
    console.log('Подключено к базе данных цен MySQL');
  }
});

const app = express();

// Нужно разрешить кросс-доменные запросы, сейчас это localhost:5173 vite или 3001 create-react-app
app.use(cors({
  origin: ['http://localhost:3001', 'http://localhost:5173'],
  credentials: true,
  maxAge: 60,
}));

app.use(express.raw());

app.use(express.json({ limit: '10mb' }));

app.use(requestLogger);

// Передаём соединения с MySQL корневому маршруту
app.use('/', (req, res, next) => {
  req.fanDataDb = fanDataDb;
  req.priceDb = priceDb;
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
  console.log(`Приложение запущенно на порте ${PORT}`);
});
