import { readFileSync } from 'node:fs';
import vm from 'node:vm';
import test from 'node:test';
import assert from 'node:assert/strict';
import path from 'node:path';
import { randomUUID } from 'node:crypto';

const sw = readFileSync(new URL('./public/sw.js', import.meta.url), 'utf8');
const html = readFileSync(new URL('./index.html', import.meta.url), 'utf8');
const registration = html.slice(html.indexOf("        if ('serviceWorker'"), html.indexOf('        // ── 공통:'));
function worker(windows = []) {
  const events = {}, deleted = [], puts = [];
  let skips = 0;
  const cache = { addAll: async () => {}, match: async () => undefined, put: async (...args) => puts.push(args) };
  const ctx = {
    self: { addEventListener: (k, f) => events[k] = f, skipWaiting: () => { skips++; },
      clients: { matchAll: async () => windows }, location: { hostname: 'example.com' }, registration: { scope: 'https://example.com/' } },
    caches: { open: async () => cache, keys: async () => ['ssapp-v13', 'ssapp-build-__PWA_BUILD_ID__', 'other-app'], delete: async k => deleted.push(k) },
    URL, Response, Request, MessageChannel, setTimeout, clearTimeout, fetch: async () => new Response('html', { headers: { 'content-type': 'text/html' } })
  };
  vm.createContext(ctx); vm.runInContext(sw, ctx);
  return { events, deleted, puts, ctx, skips: () => skips };
}
test('install waits; only explicit message skips waiting', async () => {
  const w = worker(); let pending;
  w.events.install({ waitUntil: p => pending = p }); await pending;
  assert.equal(w.skips(), 0);
  w.events.message({ data: { type: 'SKIP_WAITING' }, source: { id: 'current' }, waitUntil: p => pending = p });
  await pending; assert.equal(w.skips(), 1);
});
test('legacy other window blocks activation; safe other window allows it', async () => {
  for (const safe of [false, true]) {
    let blocked = false, pending;
    const w = worker([{ id: 'other', postMessage: (_, ports) => {
      ports[0].postMessage(safe ? 'PWA_UPDATE_SAFE' : 'legacy'); ports[0].close();
    } }]);
    w.events.message({ data: { type: 'SKIP_WAITING' },
      source: { id: 'current', postMessage: () => blocked = true }, waitUntil: p => pending = p });
    await pending; assert.equal(w.skips(), safe ? 1 : 0); assert.equal(blocked, !safe);
  }
});
test('cleanup preserves current and unrelated cache; open windows preserve old assets', async () => {
  for (const windows of [[], [{}]]) {
    const w = worker(windows); let pending;
    w.events.activate({ waitUntil: p => pending = p }); await pending;
    assert.deepEqual(w.deleted, windows.length ? [] : ['ssapp-v13']);
  }
});
test('HTML revalidates and refreshes offline cache', async () => {
  const w = worker(); let options;
  w.ctx.fetch = async (_, o) => { options = o; return new Response('new', { headers: { 'content-type': 'text/html' } }); };
  await vm.runInContext('networkHtml({})', w.ctx);
  assert.equal(options.cache, 'no-cache'); assert.equal(w.puts.length, 1);
});
function page() {
  const events = {}, windowEvents = {}, documentEvents = {}; let reloads = 0, checks = 0, messages = 0, visible = false, click;
  const reg = { waiting: { postMessage: () => messages++ }, update: async () => { checks++; }, addEventListener: (k,f) => events[k] = f };
  const ctx = {
    navigator: { serviceWorker: { controller: {}, addEventListener: (k,f) => events[k] = f, register: async () => reg, getRegistration: async () => reg } },
    window: { location: { reload: () => reloads++ }, addEventListener: (k,f) => windowEvents[k] = f },
    document: { visibilityState: 'visible', addEventListener: (k,f) => documentEvents[k] = f,
      getElementById: id => id === 'sw-update-banner' ? { classList: { add: () => visible = true } } : { addEventListener: (_,f) => click = f } },
    setInterval: () => {}, console
  };
  vm.runInNewContext(registration, ctx);
  return { events, windowEvents, documentEvents, click: () => click(), stats: () => ({ reloads, checks, messages, visible }) };
}
test('existing banner reused; only consenting page reloads once', async () => {
  const a = page(), b = page(); a.windowEvents.load(); b.windowEvents.load(); await new Promise(setImmediate);
  assert.equal(a.stats().visible, true);
  a.click(); await new Promise(setImmediate);
  a.events.controllerchange(); b.events.controllerchange(); a.events.controllerchange();
  assert.equal(a.stats().reloads, 1); assert.equal(b.stats().reloads, 0); assert.equal(a.stats().messages, 1);
});
test('startup, resume and online check updates', async () => {
  const p = page(); p.windowEvents.load(); await new Promise(setImmediate);
  p.windowEvents.online(); p.windowEvents.pageshow(); p.documentEvents.visibilitychange();
  assert.equal(p.stats().checks, 4);
});
test('PWA update code never clears user storage', () => {
  assert.doesNotMatch(sw + registration, /localStorage|indexedDB|Clear-Site-Data/);
  const config = readFileSync(new URL('./vite.config.ts', import.meta.url), 'utf8');
  assert.match(config, /randomUUID\(\)/); assert.match(config, /dist\/public\/sw.js/);
});
test('build hook creates distinct worker versions in output only', () => {
  const config = readFileSync(new URL('./vite.config.ts', import.meta.url), 'utf8');
  const body = config.match(/closeBundle\(\) \{([\s\S]*?)\n      \},/)[1];
  const outputs = [];
  const ctx = { path, randomUUID, readFileSync: () => sw,
    writeFileSync: (file, content) => outputs.push({ file, content }) };
  for (let i = 0; i < 2; i++) vm.runInNewContext('{' + body.replaceAll('import.meta.dirname', JSON.stringify('/app')) + '}', ctx);
  assert.notEqual(outputs[0].content, outputs[1].content);
  for (const output of outputs) {
    assert.ok(output.file.endsWith(path.join('dist', 'public', 'sw.js')));
    assert.ok(!output.content.includes('__PWA_BUILD_ID__'));
    new vm.Script(output.content);
  }
});
