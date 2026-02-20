// Vercel's static bundler (nft) misses dynamically-required dialect modules.
// Explicitly requiring pg here forces the bundler to include it in the Lambda.
require('pg');
require('pg-hstore');

module.exports = require('./index.js');
