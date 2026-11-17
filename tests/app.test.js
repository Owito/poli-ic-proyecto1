'use strict';
// Pruebas unitarias con una BD simulada: no requieren Docker.
const test = require('node:test');
const assert = require('node:assert/strict');
const { crearApp } = require('../src/app');

function servir(pool) {
  return new Promise((ok) => {
    const s = crearApp(pool).listen(0, () => ok({ s, url: `http://127.0.0.1:${s.address().port}` }));
  });
}

test('POST /tareas sin titulo responde 400', async (t) => {
  const { s, url } = await servir({ query: async () => { throw new Error('no debe llegar a la BD'); } });
  t.after(() => s.close()); // cierra el servidor aunque la prueba falle
  const r = await fetch(`${url}/tareas`, { method: 'POST', headers: { 'content-type': 'application/json' }, body: '{}' });
  assert.equal(r.status, 400);
});

test('GET /tareas con estado inválido responde 400', async (t) => {
  const { s, url } = await servir({ query: async () => ({ rows: [] }) });
  t.after(() => s.close()); // cierra el servidor aunque la prueba falle
  const r = await fetch(`${url}/tareas?estado=xyz`);
  assert.equal(r.status, 400);
});

test('GET /tareas/:id inexistente responde 404', async (t) => {
  const { s, url } = await servir({ query: async () => ({ rows: [] }) });
  t.after(() => s.close()); // cierra el servidor aunque la prueba falle
  const r = await fetch(`${url}/tareas/999`);
  assert.equal(r.status, 404);
});

test('GET /salud responde 503 si la BD falla', async (t) => {
  const { s, url } = await servir({ query: async () => { throw new Error('caída'); } });
  t.after(() => s.close()); // cierra el servidor aunque la prueba falle
  const r = await fetch(`${url}/salud`);
  assert.equal(r.status, 503);
  assert.equal((await r.json()).estado, 'error');
});
