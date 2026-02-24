const http = require('http');
const { execFile, execSync, spawnSync } = require('child_process');
const fs = require('fs');
const os = require('os');
const path = require('path');
const crypto = require('crypto');

const PORT = 7821;

// ── Find pandoc — exhaustive search ────────────────────────────────────────
function findPandoc() {
  const isWin = process.platform === 'win32';
  const home  = os.homedir();

  const candidates = [
    // PATH lookup
    isWin ? 'pandoc.exe' : 'pandoc',

    // ← This exact path found by /diagnose on your machine
    path.join(home, 'AppData', 'Local', 'Pandoc', 'pandoc.exe'),

    // Windows — all known install locations
    'C:\\Program Files\\Pandoc\\pandoc.exe',
    'C:\\Program Files (x86)\\Pandoc\\pandoc.exe',
    path.join(home, 'AppData', 'Local', 'Pandoc', 'pandoc.exe'),
    path.join(home, 'AppData', 'Local', 'Programs', 'Pandoc', 'pandoc.exe'),
    path.join(home, 'AppData', 'Roaming', 'Pandoc', 'pandoc.exe'),
    'C:\\ProgramData\\chocolatey\\bin\\pandoc.exe',    // Chocolatey
    'C:\\tools\\pandoc\\pandoc.exe',                   // Scoop / manual

    // Mac
    '/usr/local/bin/pandoc',
    '/opt/homebrew/bin/pandoc',
    '/usr/bin/pandoc',
    path.join(home, '.local', 'bin', 'pandoc'),

    // Linux
    '/usr/bin/pandoc',
    '/usr/local/bin/pandoc',
    '/snap/bin/pandoc',
    '/home/linuxbrew/.linuxbrew/bin/pandoc',
    path.join(home, '.local', 'bin', 'pandoc'),
  ];

  // 1. Try every hardcoded path
  for (const p of candidates) {
    try {
      execSync(`"${p}" --version`, { stdio: 'ignore', timeout: 3000 });
      return p;
    } catch (_) {}
  }

  // 2. Windows: try 'where pandoc' which searches PATH + App Paths registry
  if (isWin) {
    try {
      const found = execSync('where pandoc', { encoding: 'utf8', timeout: 5000 }).trim().split('\n')[0].trim();
      if (found) {
        execSync(`"${found}" --version`, { stdio: 'ignore', timeout: 3000 });
        return found;
      }
    } catch (_) {}

    // 3. Windows: check registry directly
    try {
      const reg = execSync(
        'reg query "HKLM\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\App Paths\\pandoc.exe" /ve',
        { encoding: 'utf8', timeout: 5000 }
      );
      const match = reg.match(/REG_SZ\s+(.+)/);
      if (match) {
        const regPath = match[1].trim();
        execSync(`"${regPath}" --version`, { stdio: 'ignore', timeout: 3000 });
        return regPath;
      }
    } catch (_) {}
  }

  return null;
}

const PANDOC = findPandoc();
console.log('\n' + (PANDOC ? `✅ Pandoc found: ${PANDOC}` : '❌ Pandoc not found') + '\n');

// ── HTTP server ─────────────────────────────────────────────────────────────
const server = http.createServer((req, res) => {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  if (req.method === 'OPTIONS') { res.writeHead(204); res.end(); return; }

  // ── Serve UI ──
  if (req.method === 'GET' && req.url === '/') {
    const html = fs.readFileSync(path.join(__dirname, 'index.html'));
    res.writeHead(200, { 'Content-Type': 'text/html; charset=utf-8' });
    res.end(html);
    return;
  }

  // ── Status ──
  if (req.method === 'GET' && req.url === '/status') {
    res.writeHead(PANDOC ? 200 : 503, { 'Content-Type': 'application/json' });
    res.end(JSON.stringify({ ok: !!PANDOC, pandoc: PANDOC || null }));
    return;
  }

  // ── Diagnose — call this if it's not working ──
  if (req.method === 'GET' && req.url === '/diagnose') {
    const report = {
      platform: process.platform,
      node: process.version,
      PATH: process.env.PATH,
      pandocFound: PANDOC,
      checks: {},
    };

    const toCheck = [
      'pandoc', 'pandoc.exe',
      'C:\\Program Files\\Pandoc\\pandoc.exe',
      path.join(os.homedir(), 'AppData', 'Local', 'Pandoc', 'pandoc.exe'),
      '/usr/bin/pandoc', '/usr/local/bin/pandoc', '/opt/homebrew/bin/pandoc',
    ];

    for (const p of toCheck) {
      const exists = fs.existsSync(p);
      let runs = false;
      try { execSync(`"${p}" --version`, { stdio: 'ignore', timeout: 2000 }); runs = true; } catch(_) {}
      report.checks[p] = { exists, runs };
    }

    // where/which
    try {
      report.whereResult = execSync(
        process.platform === 'win32' ? 'where pandoc' : 'which pandoc',
        { encoding: 'utf8', timeout: 3000 }
      ).trim();
    } catch(e) { report.whereResult = 'not found: ' + e.message; }

    res.writeHead(200, { 'Content-Type': 'application/json' });
    res.end(JSON.stringify(report, null, 2));
    return;
  }

  // ── Convert ──
  if (req.method === 'POST' && req.url === '/convert') {
    if (!PANDOC) {
      res.writeHead(503, { 'Content-Type': 'application/json' });
      res.end(JSON.stringify({
        error: 'Pandoc not found. Open http://127.0.0.1:' + PORT + '/diagnose in your browser to see the full diagnosis.'
      }));
      return;
    }

    let body = '';
    req.on('data', d => { body += d; if (body.length > 5e6) req.destroy(); });
    req.on('end', () => {
      let markdown;
      try { markdown = JSON.parse(body).markdown; } catch(e) {
        res.writeHead(400); res.end('Bad JSON'); return;
      }
      if (!markdown || !markdown.trim()) {
        res.writeHead(400); res.end('Empty input'); return;
      }

      const id      = crypto.randomBytes(8).toString('hex');
      const inFile  = path.join(os.tmpdir(), `cv_in_${id}.md`);
      const outFile = path.join(os.tmpdir(), `cv_out_${id}.docx`);
      const refDoc  = path.join(__dirname, 'reference.docx');

      fs.writeFileSync(inFile, normalizeMath(markdown), 'utf8');

      const args = [
        inFile, '-o', outFile,
        '--from', 'markdown+tex_math_dollars+tex_math_single_backslash',
        '--to', 'docx',
        '--metadata', 'lang=en-US',   // Force LTR — fixes RTL on Arabic Word
      ];
      if (fs.existsSync(refDoc)) args.push('--reference-doc', refDoc);

      execFile(PANDOC, args, { timeout: 30000 }, (err, _out, stderr) => {
        fs.unlink(inFile, () => {});
        if (err) {
          fs.unlink(outFile, () => {});
          res.writeHead(500, { 'Content-Type': 'application/json' });
          res.end(JSON.stringify({ error: 'Pandoc error: ' + (stderr || err.message) }));
          return;
        }
        // Post-process: inject bidi=0 into every paragraph to force LTR
        // This fixes RTL alignment on Arabic/Hebrew Word installations
        injectLTR(outFile);

        const docx = fs.readFileSync(outFile);
        fs.unlink(outFile, () => {});
        res.writeHead(200, {
          'Content-Type': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
          'Content-Disposition': 'attachment; filename="ai-output.docx"',
          'Content-Length': docx.length,
        });
        res.end(docx);
      });
    });
    return;
  }

  res.writeHead(404); res.end('Not found');
});

// ── Post-process docx: inject bidi=0 to force LTR ─────────────────────────
// Fixes RTL alignment on Arabic/Hebrew Word — calls inject_ltr.py
function injectLTR(docxPath) {
  const script = path.join(__dirname, 'inject_ltr.py');
  if (!fs.existsSync(script)) return;
  const pyExe = process.platform === 'win32' ? 'python' : 'python3';
  try {
    execSync(`"${pyExe}" "${script}" "${docxPath}"`, { timeout: 10000 });
  } catch(e) {
    // Try the other python name on Windows
    try { execSync(`python3 "${script}" "${docxPath}"`, { timeout: 10000 }); }
    catch(_) { console.warn('injectLTR skipped:', e.message); }
  }
}



function normalizeMath(text) {
  text = text.replace(/^[ \t]*\[\s*\n([\s\S]*?)\n[ \t]*\][ \t]*$/gm,
    (_, e) => `$$\n${e.trim()}\n$$`);
  text = text.replace(/^[ \t]*\[[ \t]*([^\]\n]{1,400})[ \t]*\][ \t]*$/gm, (_, e) => {
    const inner = e.trim();
    if (inner && !inner.startsWith('http') && isMath(inner)) return `$$${inner}$$`;
    return `[${e}]`;
  });
  return text;
}

function isMath(s) {
  return /[\\^_{}]|\\[a-zA-Z]|frac|sum|int|lim|sqrt|begin|vec|cdot|times|alpha|beta|gamma|sigma|nabla|partial|infty|approx|leq|geq|neq|matrix|pmatrix|bmatrix/.test(s)
    || /[+\-=×÷*\/]/.test(s) || /\d/.test(s)
    || /[a-zA-Z]\s*[+\-=*\/]/.test(s);
}

server.listen(PORT, '127.0.0.1', () => {
  console.log(`🚀 Server: http://127.0.0.1:${PORT}`);
  console.log(`🔍 Diagnose: http://127.0.0.1:${PORT}/diagnose\n`);
});
