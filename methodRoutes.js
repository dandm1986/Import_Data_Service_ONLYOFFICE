const express = require('express');
const router = express.Router();
const methodController = require('./methodController');

router.route('/import-range').post(methodController.getCellsValues);

module.exports = router;
