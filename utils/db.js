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
    connectionLimit: 5,
    queueLimit: 0,
    connectTimeout: 60000,
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

recreatePool = async (dbName) => {
  try {
    let newPool;
    if (dbName === MYSQL_FAN_DATABASE) {
      await fanDataDb.end();
      // Создаем новый пул
      fanDataDb = createPool(dbName);
      newPool = fanDataDb;
    } else if (dbName === MYSQL_PRICE_DATABASE) {
      await priceDb.end();
      priceDb = createPool(dbName);
      newPool = priceDb;
    }
    console.log(`Успешно пересоздано соединение для базы данных ${dbName}`);
    return newPool;
  } catch (error) {
    console.error(`Ошибка при пересоздании соединения для базы данных ${dbName}:`, error);
    throw error;
  }
};

// Проверка состояния соединения и переподключение при его потере
const checkAndReconnect = async (pool, dbName) => {
  try {
    // Проверяем и возвращаем старое соединение
    const connection = await pool.getConnection();
    connection.release();
  } catch (error) {
    if (error.code === 'PROTOCOL_CONNECTION_LOST' || error.code === 'ECONNRESET') {
      await recreatePool(dbName);
    } else {
      throw error;
    }
  }
};

// Обработка событий завершения работы сервера для корректного закрытия соединений
process.on('SIGINT', async () => {
  await Promise.all([
    fanDataDb.end(),
    priceDb.end(),
  ]);
  // Завершаем процесс
  process.exit();
});

module.exports = { fanDataDb, priceDb, checkAndReconnect };
