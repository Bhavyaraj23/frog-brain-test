// database.js
// Sets up SQLite database and creates the results table

const Database = require('better-sqlite3');
const path = require('path');

// Database file will be created in the project folder automatically
const DB_PATH = path.join(__dirname, 'frog_results.db');

let db;

function getDB() {
  if (!db) {
    db = new Database(DB_PATH);

    // Enable WAL mode for better performance
    db.pragma('journal_mode = WAL');

    // Create the main results table if it doesn't exist
    db.exec(`
      CREATE TABLE IF NOT EXISTS game_results (
        id            INTEGER PRIMARY KEY AUTOINCREMENT,
        subject_name  TEXT    NOT NULL,
        test_date     TEXT    NOT NULL,
        test_time     TEXT    NOT NULL,

        -- Each hit column stores the time in seconds from game start,
        -- OR the string "MISSED" if the frog was not tapped
        hit_1         TEXT,
        hit_2         TEXT,
        hit_3         TEXT,
        hit_4         TEXT,
        hit_5         TEXT,
        hit_6         TEXT,

        final_score   INTEGER NOT NULL,
        total_rounds  INTEGER NOT NULL DEFAULT 6,
        notes         TEXT
      );
    `);

    console.log('✅ Database ready at:', DB_PATH);
  }
  return db;
}

module.exports = { getDB };
