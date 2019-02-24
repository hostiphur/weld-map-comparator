'use strict';

const fs = require('fs');
const Hjson = require('hjson');

const configText = fs.readFileSync('./config/config.hjson', 'utf8');
const configObject = Hjson.parse(configText);

module.exports = configObject;