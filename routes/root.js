const router = require('express').Router();
const { getDataSheet } = require('../controllers/datasheet');
const { getFanModels, getFanDataPoints } = require('../controllers/fan-data');

router.use('/fans', getFanModels);
router.use('/fanDataPoints', getFanDataPoints);
router.use('/pdf', getDataSheet);

module.exports = { router };
