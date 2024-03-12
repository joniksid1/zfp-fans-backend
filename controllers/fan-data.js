const { getFanModels, getFanDataPoints } = require('../utils/database-query-service');

module.exports.getFanModels = async (req, res, next) => {
  try {
    const modelsArray = await getFanModels();
    res.status(200).json({
      modelsArray,
    });
  } catch (e) {
    next(e);
  }
};

module.exports.getFanDataPoints = async (req, res, next) => {
  try {
    const fanData = await getFanDataPoints();
    res.status(200).json({
      fanData,
    });
  } catch (e) {
    next(e);
  }
};
