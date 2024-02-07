module.exports = (err, req, res) => {
  const { statusCode = 500, message } = err;
  console.error('Error:', err);
  res.status(500).json({
    Ошибка: err,
  });
  if (err.code === 'ETIMEDOUT') {
    // Обработка ошибки ETIMEDOUT
    res.status(500).json({
      error: 'Ошибка соединения с базой данных: превышено время ожидания',
    });
  } else if (err.code === 'PROTOCOL_CONNECTION_LOST') {
    // Дополнительная обработка ошибки PROTOCOL_CONNECTION_LOST
    res.status(500).json({
      error: 'Соединение с базой данных было потеряно',
    });
  } else if (err.code === 'ECONNRESET') {
    // Дополнительная обработка ошибки ECONNRESET
    res.status(500).json({
      error: 'Соединение с базой данных было сброшено',
    });
  } else if (err.code === 'ER_TOO_MANY_USER_CONNECTIONS') {
    // Дополнительная обработка ошибки ER_TOO_MANY_USER_CONNECTIONS
    res.status(500).json({
      error: 'Превышено максимальное количество соединений с базой данных на пользователя',
    });
  } else {
    // Обработка других ошибок
    res.status(statusCode).json({
      message: statusCode === 500
        ? 'На сервере произошла ошибка'
        : message,
    });
  }
};
