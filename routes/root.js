const router = require('express').Router();
const { getDataSheet } = require('../controllers/datasheet');
const { getFanModels, getFanDataPoints } = require('../controllers/fan-data');
const { getPdfFromXlsx } = require('../controllers/pdf-test');

router.use('/fans', getFanModels);
router.use('/fanDataPoints', getFanDataPoints);
router.use('/excel', getDataSheet);
router.use('/excelToPdf', getPdfFromXlsx);

module.exports = { router };
