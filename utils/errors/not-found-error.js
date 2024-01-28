const { constants } = require('http2');

const { HTTP_STATUS_NOT_FOUND } = constants;

class NotFoundError extends Error {
  constructor({ message }) {
    super(message);
    this.statusCode = HTTP_STATUS_NOT_FOUND;
  }
}

module.exports = { NotFoundError };
