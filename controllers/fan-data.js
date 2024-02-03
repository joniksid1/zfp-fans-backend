const { NotFoundError } = require('../utils/errors/not-found-error');
const { fanDataDb } = require('../utils/db');

const { MYSQL_FAN_DATABASE } = process.env;

// Экспортируем функцию обработки запроса

module.exports.getFanModels = async (req, res, next) => {
  let connection;
  try {
    // Получаем соединение из пула
    connection = await fanDataDb.getConnection();

    // Выполняем запрос
    const [allModelsQuery] = await connection.execute(`
      SELECT DISTINCT model
      FROM ${MYSQL_FAN_DATABASE}.zfr_data;
    `);
    if (allModelsQuery.length === 0) {
      throw new NotFoundError({ message: 'Не удалось найти данные названий вентиляторов в базе' });
    }

    // Извлекаем значения model и создаем массив строк
    const modelsArray = allModelsQuery.map((result) => result.model);

    // Отправляем данные на фронтенд
    res.status(200).json({
      modelsArray,
    });
  } catch (e) {
    next(e);
  } finally {
    // Закрываем соединение после использования
    if (connection) {
      await connection.release();
    }
  }
};

module.exports.getFanDataPoints = async (req, res, next) => {
  const fanModels = [
    'zfr_1_9_2e',
    'zfr_2_25_2e',
    'zfr_2_5_2e',
    'zfr_2_8_2e',
    'zfr_3_1_4e',
    'zfr_3_1_4d',
    'zfr_3_5_4d',
    'zfr_3_5_4e',
    'zfr_4_4d',
    'zfr_4_4e',
    'zfr_4_5_4d',
    'zfr_4_5_4e',
    'zfr_5_4d',
    'zfr_5_6_4d',
    'zfr_6_3_4d',
  ];
  let connection;

  try {
    const fanDataResults = [];

    // Используйте Promise.allSettled для обработки асинхронного освобождения
    const results = await Promise.allSettled(fanModels.map(async (fanModel) => {
      try {
        // Получаем соединение из пула
        connection = await fanDataDb.getConnection();

        // Выполняем запрос
        const [fanDataQuery] = await connection.execute(`
      SELECT x, y
      FROM ${MYSQL_FAN_DATABASE}.${fanModel}_dataset;
    `);

        if (fanDataQuery.length === 0) {
          throw new NotFoundError({ message: `Не удалось найти данные вентилятора ${fanModel} в базе` });
        }

        // Добавляем результаты запросов в массив
        fanDataResults.push({
          model: fanModel,
          data: fanDataQuery.map((result) => ({ x: result.x, y: result.y })),
        });
      } catch (e) {
        return e; // Возвращаем ошибку для последующей проверки
      } finally {
        // Закрываем соединение после использования
        if (connection) {
          await connection.release();
        }
      }
    }));

    // Проверяем наличие ошибок в результатах
    const errors = results.filter((result) => result.status === 'rejected').map((result) => result.reason);

    if (errors.length > 0) {
      throw errors[0]; // Бросаем первую обнаруженную ошибку
    }

    // Отправляем данные на фронтенд
    res.status(200).json({
      fanData: fanDataResults,
    });
  } catch (e) {
    next(e);
  }
};
