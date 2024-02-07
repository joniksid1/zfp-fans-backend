const router = require('express').Router();
const { getDataSheet } = require('../controllers/datasheet');
const { getFanModels, getFanDataPoints } = require('../controllers/fan-data');
const { getCommercialOffer } = require('../controllers/commercial-offer');

router.get('/models', getFanModels);
router.get('/data-points', getFanDataPoints);
router.post('/data-sheet', getDataSheet);
router.post('/offer', getCommercialOffer);
module.exports = { router };
