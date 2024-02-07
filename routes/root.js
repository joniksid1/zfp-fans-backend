const router = require('express').Router();
const { getDataSheet } = require('../controllers/datasheet');
const { getFanModels, getFanDataPoints } = require('../controllers/fan-data');
const { getPdfFromXlsx } = require('../controllers/pdf-test');
const { getCommercialOffer } = require('../controllers/commercial-offer');
const { convertToPdf } = require('../controllers/pdf-convert');

router.get('/fans', getFanModels);
router.get('/fanDataPoints', getFanDataPoints);
router.post('/excel', getDataSheet);
router.post('/excelToPdf', getPdfFromXlsx);
router.post('/price', getCommercialOffer);
router.post('/convert', convertToPdf);

module.exports = { router };
