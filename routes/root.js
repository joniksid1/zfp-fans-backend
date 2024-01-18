const router = require('express').Router();
const { getDataSheet } = require('../controllers/datasheet');

router.use('/pdf', getDataSheet);

module.exports = { router };
