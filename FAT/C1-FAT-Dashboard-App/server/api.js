/**
 * ==============================================================================
 * KALASIN ENGINEERING CO., LTD. — BACKEND REST API SERVICE
 * PROJECT: SPRINT 18K TPA SPRAY DRYER PLANT (KALASIN)
 * TARGET: Chassis C1 DI/DO FAT Test Verification System
 * ==============================================================================
 */

const http = require('http');
const fs = require('fs');
const path = require('path');
const url = require('url');

const PORT = process.env.PORT || 4000;
const DATA_FILE = path.join(__dirname, 'database', 'c1_dido_data.json');

// Helper to read JSON database cache
function loadDatabaseData() {
  try {
    if (fs.existsSync(DATA_FILE)) {
      const raw = fs.readFileSync(DATA_FILE, 'utf8');
      return JSON.parse(raw);
    }
  } catch (err) {
    console.error('Error reading JSON data:', err);
  }
  return { session: {}, slots: [], loops: [] };
}

const server = http.createServer((req, res) => {
  // CORS Headers
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');

  if (req.method === 'OPTIONS') {
    res.writeHead(204);
    res.end();
    return;
  }

  const parsedUrl = url.parse(req.url, true);
  const pathname = parsedUrl.pathname;
  const dbData = loadDatabaseData();

  // API 1: /api/c1-fat/summary
  if (pathname === '/api/c1-fat/summary' && req.method === 'GET') {
    res.writeHead(200, { 'Content-Type': 'application/json' });
    res.end(JSON.stringify({
      success: true,
      data: dbData.session
    }));
    return;
  }

  // API 2: /api/c1-fat/slots
  if (pathname === '/api/c1-fat/slots' && req.method === 'GET') {
    res.writeHead(200, { 'Content-Type': 'application/json' });
    res.end(JSON.stringify({
      success: true,
      count: dbData.slots.length,
      data: dbData.slots
    }));
    return;
  }

  // API 3: /api/c1-fat/loops (supports ?category= and ?slot=)
  if (pathname === '/api/c1-fat/loops' && req.method === 'GET') {
    let loops = dbData.loops;
    const category = parsedUrl.query.category;
    const slot = parsedUrl.query.slot;

    if (category) {
      loops = loops.filter(l => l.category === category);
    }
    if (slot) {
      loops = loops.filter(l => l.do_slot === slot || l.di_slot === slot);
    }

    res.writeHead(200, { 'Content-Type': 'application/json' });
    res.end(JSON.stringify({
      success: true,
      count: loops.length,
      data: loops
    }));
    return;
  }

  // API 4: /api/c1-fat/simulate (trigger pulse simulation)
  if (pathname === '/api/c1-fat/simulate' && req.method === 'POST') {
    let body = '';
    req.on('data', chunk => { body += chunk; });
    req.on('end', () => {
      let params = {};
      try { params = JSON.parse(body); } catch (e) {}

      const loopId = params.loop_id || 'C1-LPB-001';
      const latency = Math.floor(Math.random() * (42 - 22 + 1)) + 22; // 22-42 ms

      res.writeHead(200, { 'Content-Type': 'application/json' });
      res.end(JSON.stringify({
        success: true,
        test_id: loopId,
        do_command: 1,
        relay_contact_closed: true,
        di_feedback_detected: 1,
        latency_ms: latency,
        verdict: 'PASS',
        timestamp: new Date().toISOString()
      }));
    });
    return;
  }

  // Default 404
  res.writeHead(404, { 'Content-Type': 'application/json' });
  res.end(JSON.stringify({ error: 'Endpoint not found' }));
});

server.listen(PORT, () => {
  console.log(`[✓] C1 FAT Dashboard Backend API running on http://localhost:${PORT}`);
  console.log(`    Endpoints:`);
  console.log(`      GET  /api/c1-fat/summary`);
  console.log(`      GET  /api/c1-fat/slots`);
  console.log(`      GET  /api/c1-fat/loops`);
  console.log(`      POST /api/c1-fat/simulate`);
});
