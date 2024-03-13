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
  MODE = 'dev',
} = process.env;

const app = express();

// Нужно разрешить кросс-доменные запросы, сейчас это localhost:5173 vite или 3001 create-react-app
app.use(cors({
  origin: [
    'http://localhost:3001',
    'http://localhost:5173',
    'https://web-rooffan',
    'http://web-rooffan',
    'http://192.168.97.110',
    'https://192.168.97.110',
  ],
  credentials: true,
  maxAge: 60,
}));

app.use(express.raw());

app.use(express.json({ limit: '20mb' }));

app.use(requestLogger);

// Маршрутизация в зависимости от режима (разработка и прод)
if (MODE === 'production') {
  app.use('/api', router);
} else {
  app.use('/', router);
}

app.use('*', (req, res, next) => {
  next(new NotFoundError({ message: 'Страница не найдена' }));
});

app.use(errorLogger);

// Обработчик ошибок celebrate
app.use(errors());

// Централизованный middleware-обработчик
app.use(error);

app.listen(PORT, () => {
  console.log(`Приложение запущенно на порте ${PORT}`);
});
