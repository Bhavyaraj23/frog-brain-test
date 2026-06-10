// server.js  —  Frog Brain Test  |  Backend Server
// ─────────────────────────────────────────────────────────────
// Serves the game HTML, saves results to SQLite, exports Excel.
// ─────────────────────────────────────────────────────────────

const express = require('express');
const cors    = require('cors');
const path    = require('path');
const ExcelJS = require('exceljs');

const fetch = (...args) =>
  import('node-fetch').then(({ default: fetch }) => fetch(...args));

const { getDB } = require('./database');

const app  = express();
const PORT = process.env.PORT || 3000;

// ── Middleware ────────────────────────────────────────────────
app.use(cors());                           // Allow cross-origin requests
app.use(express.json());                   // Parse JSON request bodies
app.use(express.static(path.join(__dirname, 'public'))); // Serve the game files

// ── 0.  ROOT — serve the game  ───────────────────────────────
app.get('/', (req, res) => {
  res.sendFile(path.join(__dirname, 'game_index.html'));
});

// ── 1.  SAVE GAME RESULT  ─────────────────────────────────────
//  POST /api/save-result
//  Body: { subjectName, hits: ["1.23","MISSED","2.11",...], finalScore }
app.post('/api/save-result', async (req, res) => {
  try {
    const { subjectName, hits, finalScore, notes } = req.body;

    // Basic validation
    if (!subjectName || !hits || !Array.isArray(hits)) {
      return res.status(400).json({ error: 'Missing required fields.' });
    }
    if (hits.length !== 6) {
      return res.status(400).json({ error: 'hits must be an array of 6 values.' });
    }

    // Date and time of the test
    const now  = new Date();
    const date = now.toLocaleDateString('en-IN');   // e.g. 08/04/2026
    const time = now.toLocaleTimeString('en-IN');   // e.g. 10:34:22 AM

    const db   = getDB();
    const stmt = db.prepare(`
      INSERT INTO game_results
        (subject_name, test_date, test_time,
         hit_1, hit_2, hit_3, hit_4, hit_5, hit_6,
         final_score, total_rounds, notes)
      VALUES
        (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 6, ?)
    `);

    const result = stmt.run(
      subjectName.trim(),
      date, time,
      hits[0], hits[1], hits[2], hits[3], hits[4], hits[5],
      finalScore,
      notes || ''
    );

    console.log(`✅ Result saved — ID: ${result.lastInsertRowid} | Subject: ${subjectName} | Score: ${finalScore}/6`);

// Save to Google Sheet
try {
  await fetch('https://script.google.com/macros/s/AKfycbyZl99KBx4l81PRNCLbLlvfByGJpXWUaXSOnOmqJRtoqAafFiBGIz-K7Mv-5_p4p5nHag/exec', {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json'
    },
    body: JSON.stringify({
      subjectName,
      date,
      time,
      finalScore,
      hit1: hits[0],
      hit2: hits[1],
      hit3: hits[2],
      hit4: hits[3],
      hit5: hits[4],
      hit6: hits[5],
      notes: notes || ''
    })
  });

  console.log('✅ Saved to Google Sheets');
} catch (sheetError) {
  console.error('❌ Google Sheet Error:', sheetError.message);
}

res.json({
  success: true,
  id: result.lastInsertRowid,
  message: `Result saved for "${subjectName}"`
});

  } catch (err) {
    console.error('❌ Save error:', err.message);
    res.status(500).json({ error: 'Server error while saving result.' });
  }
});

// ── 2.  GET ALL RESULTS (JSON)  ───────────────────────────────
//  GET /api/results
//  Returns all rows, newest first
app.get('/api/results', (req, res) => {
  try {
    const db   = getDB();
    const rows = db.prepare('SELECT * FROM game_results ORDER BY id DESC').all();
    res.json({ success: true, total: rows.length, data: rows });
  } catch (err) {
    console.error('❌ Fetch error:', err.message);
    res.status(500).json({ error: 'Could not fetch results.' });
  }
});

// ── 3.  DOWNLOAD EXCEL  ───────────────────────────────────────
//  GET /api/download
//  Returns a .xlsx file with all results
app.get('/api/download', async (req, res) => {
  try {
    const db   = getDB();
    const rows = db.prepare('SELECT * FROM game_results ORDER BY id ASC').all();

    const totalTests = rows.length;
    const avgScore   = totalTests > 0
      ? (rows.reduce((s, r) => s + r.final_score, 0) / totalTests).toFixed(2)
      : 0;

    // ── Create workbook ───────────────────────────────────────
    const workbook  = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Results');

    // ── Column widths ─────────────────────────────────────────
    worksheet.columns = [
      { width: 14 }, // Test Subject #
      { width: 22 }, // Subject Name
      { width: 14 }, // Date
      { width: 14 }, // Time
      { width: 12 }, // Hit 1
      { width: 12 }, // Hit 2
      { width: 12 }, // Hit 3
      { width: 12 }, // Hit 4
      { width: 12 }, // Hit 5
      { width: 12 }, // Hit 6
      { width: 13 }, // Score
      { width: 10 }, // %
      { width: 24 }, // Notes
    ];

    // ── Title row (merged) ────────────────────────────────────
    worksheet.addRow(['FROG BRAIN SPEED TEST — Results Data']);
    worksheet.mergeCells('A1:M1');

    // ── Blank + headers ───────────────────────────────────────
    worksheet.addRow([]);
    worksheet.addRow([
      'Test Subject #', 'Subject Name', 'Test Date', 'Test Time',
      'Hit 1 (sec)', 'Hit 2 (sec)', 'Hit 3 (sec)', 'Hit 4 (sec)',
      'Hit 5 (sec)', 'Hit 6 (sec)', 'Final Score', 'Score %', 'Notes',
    ]);

    // ── Data rows ─────────────────────────────────────────────
    rows.forEach(row => worksheet.addRow([
      row.id,
      row.subject_name,
      row.test_date,
      row.test_time,
      row.hit_1 || 'MISSED',
      row.hit_2 || 'MISSED',
      row.hit_3 || 'MISSED',
      row.hit_4 || 'MISSED',
      row.hit_5 || 'MISSED',
      row.hit_6 || 'MISSED',
      `${row.final_score} / ${row.total_rounds}`,
      `${Math.round((row.final_score / row.total_rounds) * 100)}%`,
      row.notes || '',
    ]));

    // ── Summary row ───────────────────────────────────────────
    worksheet.addRow([]);
    worksheet.addRow(['', `Total Tests: ${totalTests}`, '', '', '', '', '', '', '', '', `Avg Score: ${avgScore}/6`]);

    // ── Send file ─────────────────────────────────────────────
    const filename = `FrogBrainTest_Results_${new Date().toISOString().slice(0,10)}.xlsx`;
    const buffer   = await workbook.xlsx.writeBuffer();

    res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.send(buffer);

    console.log(`📊 Excel downloaded — ${totalTests} records — "${filename}"`);

  } catch (err) {
    console.error('❌ Excel error:', err.message);
    res.status(500).json({ error: 'Could not generate Excel file.' });
  }
});

// ── 4.  DELETE A SINGLE RESULT (optional admin use) ───────────
//  DELETE /api/results/:id
app.delete('/api/results/:id', (req, res) => {
  try {
    const db   = getDB();
    const stmt = db.prepare('DELETE FROM game_results WHERE id = ?');
    const info = stmt.run(req.params.id);
    if (info.changes === 0) return res.status(404).json({ error: 'Record not found.' });
    res.json({ success: true, message: `Record #${req.params.id} deleted.` });
  } catch (err) {
    res.status(500).json({ error: 'Could not delete record.' });
  }
});

// ── 5.  ADMIN DASHBOARD page  ─────────────────────────────────
//  GET /admin  — simple HTML table of all results
app.get('/admin', (req, res) => {
  try {
    const db   = getDB();
    const rows = db.prepare('SELECT * FROM game_results ORDER BY id DESC').all();

    const tableRows = rows.map(r => `
      <tr>
        <td>${r.id}</td>
        <td><strong>${r.subject_name}</strong></td>
        <td>${r.test_date}</td>
        <td>${r.test_time}</td>
        <td class="${r.hit_1==='MISSED'?'miss':''}">${r.hit_1||'MISSED'}</td>
        <td class="${r.hit_2==='MISSED'?'miss':''}">${r.hit_2||'MISSED'}</td>
        <td class="${r.hit_3==='MISSED'?'miss':''}">${r.hit_3||'MISSED'}</td>
        <td class="${r.hit_4==='MISSED'?'miss':''}">${r.hit_4||'MISSED'}</td>
        <td class="${r.hit_5==='MISSED'?'miss':''}">${r.hit_5||'MISSED'}</td>
        <td class="${r.hit_6==='MISSED'?'miss':''}">${r.hit_6||'MISSED'}</td>
        <td><strong>${r.final_score}/6</strong></td>
        <td>${r.notes||''}</td>
      </tr>`).join('');

    res.send(`<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8"/>
<meta name="viewport" content="width=device-width,initial-scale=1"/>
<title>Admin — Frog Brain Test Results</title>
<style>
  *{box-sizing:border-box;margin:0;padding:0}
  body{font-family:system-ui,sans-serif;background:#0d1117;color:#e6edf3;padding:1.5rem}
  h1{color:#58a6ff;margin-bottom:.3rem;font-size:1.6rem}
  .meta{color:#8b949e;font-size:.85rem;margin-bottom:1.5rem}
  .btn{display:inline-block;padding:.6rem 1.4rem;border-radius:8px;text-decoration:none;
    font-weight:700;font-size:.95rem;margin-right:.5rem;margin-bottom:1.2rem;cursor:pointer;border:none}
  .btn-green{background:#238636;color:#fff}.btn-green:hover{background:#2ea043}
  .btn-blue{background:#1f6feb;color:#fff}.btn-blue:hover{background:#388bfd}
  table{width:100%;border-collapse:collapse;font-size:.82rem;background:#161b22;border-radius:10px;overflow:hidden}
  thead{background:#21262d}
  th{padding:.65rem .8rem;text-align:left;color:#8b949e;font-weight:700;white-space:nowrap}
  td{padding:.6rem .8rem;border-top:1px solid #21262d;white-space:nowrap}
  tr:hover td{background:#1c2128}
  .miss{color:#f85149;font-weight:700}
  .stat-bar{display:flex;gap:1rem;margin-bottom:1.5rem;flex-wrap:wrap}
  .stat{background:#21262d;border-radius:10px;padding:.8rem 1.2rem;min-width:140px}
  .stat-num{font-size:1.8rem;font-weight:900;color:#58a6ff}
  .stat-lbl{font-size:.75rem;color:#8b949e;text-transform:uppercase;letter-spacing:.06em}
  @media(max-width:700px){body{padding:.8rem}table{font-size:.72rem}th,td{padding:.45rem .5rem}}
</style>
</head>
<body>
<h1>🐸 Frog Brain Test — Admin Dashboard</h1>
<p class="meta">All game results • Last updated: ${new Date().toLocaleString('en-IN')}</p>

<div class="stat-bar">
  <div class="stat"><div class="stat-num">${rows.length}</div><div class="stat-lbl">Total Tests</div></div>
  <div class="stat"><div class="stat-num">${rows.length?((rows.reduce((s,r)=>s+r.final_score,0)/rows.length).toFixed(1)):0}/6</div><div class="stat-lbl">Average Score</div></div>
  <div class="stat"><div class="stat-num">${rows.filter(r=>r.final_score>=5).length}</div><div class="stat-lbl">Score ≥ 5</div></div>
  <div class="stat"><div class="stat-num">${rows.filter(r=>r.final_score<=2).length}</div><div class="stat-lbl">Score ≤ 2 (Alert)</div></div>
</div>

<a href="/api/download" class="btn btn-green">📥 Download Excel (.xlsx)</a>
<a href="/api/results" class="btn btn-blue">📋 Raw JSON</a>

<table>
  <thead>
    <tr>
      <th>#</th><th>Name</th><th>Date</th><th>Time</th>
      <th>Hit 1</th><th>Hit 2</th><th>Hit 3</th>
      <th>Hit 4</th><th>Hit 5</th><th>Hit 6</th>
      <th>Score</th><th>Notes</th>
    </tr>
  </thead>
  <tbody>
    ${tableRows || '<tr><td colspan="12" style="text-align:center;color:#8b949e;padding:2rem">No results yet — play the game first!</td></tr>'}
  </tbody>
</table>
</body>
</html>`);
  } catch (err) {
    res.status(500).send('Server error');
  }
});

// ── Start server ──────────────────────────────────────────────
app.listen(PORT, () => {
  console.log('');
  console.log('🐸 ══════════════════════════════════════════');
  console.log(`   Frog Brain Test Server running!`);
  console.log(`   🎮 Game:    http://localhost:${PORT}`);
  console.log(`   📊 Admin:   http://localhost:${PORT}/admin`);
  console.log(`   📥 Excel:   http://localhost:${PORT}/api/download`);
  console.log('🐸 ══════════════════════════════════════════');
  console.log('');
});
