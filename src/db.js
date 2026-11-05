'use strict';
const { Pool } = require('pg');

// La configuración llega solo por variables de entorno (ver .env.example).
function crearPool() {
  return new Pool({
    host: process.env.DB_HOST || 'localhost',
    port: Number(process.env.DB_PORT || 5432),
    user: process.env.DB_USER,
    password: process.env.DB_PASSWORD,
    database: process.env.DB_NAME,
    max: 5,
    connectionTimeoutMillis: 3000,
  });
}

module.exports = { crearPool };
