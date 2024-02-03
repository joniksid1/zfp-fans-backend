const express = require('express');
require('dotenv').config();
const { errors } = require('celebrate');
const cors = require('cors');
const { router } = require('./routes/root');
const { NotFoundError } = require('./utils/errors/not-found-error');
const error = require('./middlewares/error');
const { requestLogger, errorLogger } = require('./middlewares/logger');

const {
  PORT = '3000',
} = process.env;

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
app.use('/', router);

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
