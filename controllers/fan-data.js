const { NotFoundError } = require('../utils/errors/not-found-error');

// Экспортируем функцию обработки запроса

module.exports.getFanModels = async (req, res, next) => {
  const { db } = req;

  try {
    // Получаем данные из базы mySQL
    const [allModelsQuery] = await db.promise().query(`
      SELECT DISTINCT model
      FROM zfrfans.zfr_data;
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
  }
};

module.exports.getFanDataPoints = async (req, res, next) => {
  const { db } = req;
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

  try {
    // Используем Promise.all для выполнения асинхронных запросов к базе данных
    const fanDataPromises = fanModels.map(async (fanModel) => {
      const [fanDataQuery] = await db.promise().query(`
        SELECT x, y
        FROM zfrfans.${fanModel}_dataset;
      `);

      if (fanDataQuery.length === 0) {
        throw new NotFoundError({ message: `Не удалось найти данные вентилятора ${fanModel} в базе` });
      }

      // Возвращаем объект с результатами для данной модели
      return {
        model: fanModel,
        data: fanDataQuery.map((result) => ({ x: result.x, y: result.y })),
      };
    });

    // Дожидаемся выполнения всех запросов
    const fanDataResults = await Promise.all(fanDataPromises);

    // Отправляем данные на фронтенд
    res.status(200).json({
      fanData: fanDataResults,
    });
  } catch (e) {
    // Передаем ошибку централизованному обработчику
    next(e);
  }
};
