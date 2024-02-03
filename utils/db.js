const mysql = require('mysql2/promise');

const {
  MYSQL_HOST,
  MYSQL_USER,
  MYSQL_PASSWORD,
  MYSQL_FAN_DATABASE,
  MYSQL_PRICE_DATABASE,
} = process.env;

let recreatePool;

const createPool = (database) => {
  const pool = mysql.createPool({
    host: MYSQL_HOST,
    user: MYSQL_USER,
    password: MYSQL_PASSWORD,
    database,
    waitForConnections: true,
    connectionLimit: 30,
    queueLimit: 0,
    connectTimeout: 20000,
  });

  // Отправляем запрос для поддержания активного соединения
  const sendKeepAliveQuery = async () => {
    let connection;
    try {
      connection = await pool.getConnection();
      await connection.query('SELECT 1');
    } catch (error) {
      if (error.code === 'PROTOCOL_CONNECTION_LOST' || error.code === 'ECONNRESET') {
        // В случае разрыва соединения, пересоздаем пул
        await recreatePool(pool, database);
      }
    } finally {
      if (connection) {
        connection.release();
      }

      // После выполнения запроса ждем некоторое время перед отправкой следующего
      setTimeout(sendKeepAliveQuery, 60000);
    }
  };

  // Начинаем отправку запросов для поддержания активного соединения
  sendKeepAliveQuery();

  return pool;
};

// Создаем пулы соединений для каждой базы данных
let fanDataDb = createPool(MYSQL_FAN_DATABASE);
let priceDb = createPool(MYSQL_PRICE_DATABASE);

// Пересоздание пула соединений - обрабатываем ECONNRESET
recreatePool = async (pool, dbName) => {
  try {
    // Закрываем старое соединение
    await pool.destroy();

    // Создаем новый пул
    const newPool = await createPool(dbName);

    // Обновляем переменную, содержащую пул, fanDataDb или priceDb
    if (dbName === MYSQL_FAN_DATABASE) {
      fanDataDb = newPool;
    } else if (dbName === MYSQL_PRICE_DATABASE) {
      priceDb = newPool;
    }

    console.log(`Успешно пересоздано соединение для базы данных ${dbName}`);
  } catch (error) {
    console.error(`Ошибка при пересоздании соединения для базы данных ${dbName}:`, error);
  }
};

// Обработка событий завершения работы сервера для корректного закрытия соединений
process.on('SIGINT', async () => {
  await fanDataDb.end();
  await priceDb.end();
  process.exit();
});

module.exports = { fanDataDb, priceDb };
