const {
  HTTP_STATUS_BAD_REQUEST,
} = require('http2').constants;

class CastError extends Error {
  constructor({ message }) {
    super(message);
    this.statusCode = HTTP_STATUS_BAD_REQUEST;
  }
}

module.exports = { CastError };
