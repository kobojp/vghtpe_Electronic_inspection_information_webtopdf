import { spawn } from 'node:child_process';
import { randomUUID } from 'node:crypto';
import { resolve } from 'node:path';

export default async function setup() {
  const root = resolve('..');
  const python = resolve(root, process.platform === 'win32' ? '.venv/Scripts/python.exe' : '.venv/bin/python');
  const token = randomUUID();
  process.env.VGHTPE_UI_TEST_TOKEN = token;
  const base = 'http://127.0.0.1:18865';
  const child = spawn(python, ['tests/ui_server.py', '--port', '18865'], {
    cwd: root, windowsHide: true, stdio: ['ignore', 'pipe', 'pipe'],
    env: { ...process.env, VGHTPE_UI_TEST_TOKEN: token },
  });
  let output = '';
  child.stdout.on('data', chunk => { output += chunk.toString(); });
  child.stderr.on('data', chunk => { output += chunk.toString(); });
  let spawnError: Error | undefined;
  child.on('error', error => { spawnError = error; });
  const headers = { 'X-Test-Token': token };
  const exit = new Promise<void>(done => child.on('exit', () => done()));
  async function stop() {
    try { await fetch(base + '/__test__/shutdown', { method: 'POST', headers, signal: AbortSignal.timeout(3000) }); } catch { /* The server may have exited during startup. */ }
    const stopped = await Promise.race([exit.then(() => true), new Promise<boolean>(done => setTimeout(() => done(false), 5000))]);
    if (!stopped) child.kill();
  }
  try {
    const deadline = Date.now() + 30_000;
    while (Date.now() < deadline) {
      if (spawnError) throw spawnError;
      if (child.exitCode !== null) throw new Error('UI test server exited: ' + output);
      try {
        const response = await fetch(base + '/__test__/health', { headers, signal: AbortSignal.timeout(1000) });
        if (response.ok) return stop;
      } catch { /* Wait until the owned server is ready. */ }
      await new Promise(done => setTimeout(done, 100));
    }
    throw new Error('UI test server did not become ready: ' + output);
  } catch (error) { await stop(); throw error; }
}
