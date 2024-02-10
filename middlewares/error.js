module.exports = (err, req, res, next) => {
  const { statusCode = 500, message } = err;
  console.error('Error:', err);

  if (err.code === 'ETIMEDOUT') {
    // Обработка ошибки ETIMEDOUT
    return res.status(500).json({
      error: 'Ошибка соединения с базой данных: превышено время ожидания',
    });
  }

  if (err.code === 'PROTOCOL_CONNECTION_LOST') {
    // Дополнительная обработка ошибки PROTOCOL_CONNECTION_LOST
    return res.status(500).json({
      error: 'Соединение с базой данных было потеряно',
    });
  }

  if (err.code === 'ECONNRESET') {
    // Дополнительная обработка ошибки ECONNRESET
    return res.status(500).json({
      error: 'Соединение с базой данных было сброшено',
    });
  }

  if (err.code === 'ER_TOO_MANY_USER_CONNECTIONS') {
    // Дополнительная обработка ошибки ER_TOO_MANY_USER_CONNECTIONS
    return res.status(500).json({
      error: 'Превышено максимальное количество соединений с базой данных на пользователя',
    });
  }

  // Обработка других ошибок
  return res.status(statusCode).json({
    message,
  });
};
