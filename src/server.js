'use strict';
const { crearApp } = require('./app');
const { crearPool } = require('./db');

const pool = crearPool();
const puerto = Number(process.env.PORT || 3000);
const servidor = crearApp(pool).listen(puerto, '0.0.0.0', () => console.log(`API de tareas escuchando en el puerto ${puerto}`));

function cerrar() { servidor.close(() => pool.end().then(() => process.exit(0))); }
process.on('SIGTERM', cerrar);
process.on('SIGINT', cerrar);
