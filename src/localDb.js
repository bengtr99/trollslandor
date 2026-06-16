// In-browser local database: SQLite via sql.js (WASM), persisted to IndexedDB.
// Exposes an `invoke(command, payload)` shim with the same command names the
// app previously called over Electron IPC, so App.jsx stays unchanged.
import initSqlJs from "sql.js";
import sqlWasmUrl from "sql.js/dist/sql-wasm.wasm?url";

const IDB_NAME = "trollslandeapp";
const IDB_STORE = "db";
const IDB_KEY = "observations.sqlite";

let SQL = null;
let activeDb = null;      // loaded read-only DB for queries
let buildDb = null;       // DB under construction
let progressState = { exists: false, createdAt: null, rowCount: 0, fileSize: 0, inProgress: false, progressPct: 0, progressText: "" };

async function ensureSqlJs() {
  if (SQL) return SQL;
  SQL = await initSqlJs({ locateFile: () => sqlWasmUrl });
  return SQL;
}

// ---- IndexedDB helpers (store the exported .sqlite bytes under one key) ----
function openIdb() {
  return new Promise((resolve, reject) => {
    const req = indexedDB.open(IDB_NAME, 1);
    req.onupgradeneeded = () => {
      const db = req.result;
      if (!db.objectStoreNames.contains(IDB_STORE)) db.createObjectStore(IDB_STORE);
    };
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error);
  });
}

async function idbGet(key) {
  const db = await openIdb();
  return new Promise((resolve, reject) => {
    const tx = db.transaction(IDB_STORE, "readonly");
    const req = tx.objectStore(IDB_STORE).get(key);
    req.onsuccess = () => resolve(req.result ?? null);
    req.onerror = () => reject(req.error);
    tx.oncomplete = () => db.close();
  });
}

async function idbPut(key, value) {
  const db = await openIdb();
  return new Promise((resolve, reject) => {
    const tx = db.transaction(IDB_STORE, "readwrite");
    tx.objectStore(IDB_STORE).put(value, key);
    tx.oncomplete = () => { db.close(); resolve(true); };
    tx.onerror = () => reject(tx.error);
  });
}

// ---- schema / meta helpers (mirrors the previous Electron main process) ----
function schemaSql() {
  return `
    CREATE TABLE IF NOT EXISTS observations (
      occurrenceId TEXT, species TEXT, date TEXT, province TEXT, municipality TEXT,
      locality TEXT, quantity TEXT, lifeStage TEXT, activity TEXT, recordedBy TEXT,
      latitude REAL, longitude REAL,
      isNeverFoundObservation INTEGER, isNotRediscoveredObservation INTEGER
    );
    CREATE TABLE IF NOT EXISTS meta ( key TEXT PRIMARY KEY, value TEXT );
  `;
}

function setMeta(db, key, value) {
  const stmt = db.prepare("INSERT OR REPLACE INTO meta (key, value) VALUES (?, ?)");
  stmt.run([key, String(value)]);
  stmt.free();
}

function getMeta(db, key) {
  const stmt = db.prepare("SELECT value FROM meta WHERE key = ?");
  stmt.bind([key]);
  let value = null;
  if (stmt.step()) value = stmt.getAsObject().value ?? null;
  stmt.free();
  return value;
}

function dbRowCount(db) {
  const stmt = db.prepare("SELECT COUNT(*) AS c FROM observations");
  stmt.step();
  const row = stmt.getAsObject();
  stmt.free();
  return Number(row.c || 0);
}

function closeDb(db) {
  if (db) { try { db.close(); } catch (_) {} }
}

function insertRowsInto(db, rows) {
  const stmt = db.prepare(`
    INSERT INTO observations (
      occurrenceId, species, date, province, municipality, locality, quantity,
      lifeStage, activity, recordedBy, latitude, longitude,
      isNeverFoundObservation, isNotRediscoveredObservation
    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
  `);
  for (const row of (rows || [])) {
    stmt.run([
      row.occurrenceId || "",
      row.species || "",
      row.date || "",
      row.province || "",
      row.municipality || "",
      row.locality || "",
      row.quantity == null ? "" : String(row.quantity),
      row.lifeStage || "",
      row.activity || "",
      row.recordedBy || "",
      row.latitude == null ? null : Number(row.latitude),
      row.longitude == null ? null : Number(row.longitude),
      row.isNeverFoundObservation ? 1 : 0,
      row.isNotRediscoveredObservation ? 1 : 0
    ]);
  }
  stmt.free();
}

async function loadActiveDb() {
  await ensureSqlJs();
  if (activeDb) return activeDb;
  const bytes = await idbGet(IDB_KEY);
  if (!bytes) return null;
  activeDb = new SQL.Database(new Uint8Array(bytes));
  return activeDb;
}

function normalizeString(value) {
  return String(value || "").normalize("NFC").replace(/\s+/g, " ").trim().toLowerCase();
}

function rowMatchesFilters(row, filters = {}) {
  const speciesList = Array.isArray(filters.speciesList) ? filters.speciesList.map(normalizeString) : [];
  if (speciesList.length > 0 && !speciesList.includes(normalizeString(row.species))) return false;

  const selectedLandscapes = Array.isArray(filters.selectedLandscapes) ? filters.selectedLandscapes.map(normalizeString) : [];
  if (selectedLandscapes.length > 0 && !selectedLandscapes.includes(normalizeString(row.province))) return false;

  const selectedMunicipalities = Array.isArray(filters.selectedMunicipalities) ? filters.selectedMunicipalities.map(normalizeString) : [];
  if (selectedMunicipalities.length > 0 && !selectedMunicipalities.includes(normalizeString(row.municipality))) return false;

  const date = String(row.date || "").trim();
  const yearMatch = date.match(/^(\d{4})/);
  const year = yearMatch ? yearMatch[1] : "";
  if (filters.fromYear && year && year < String(filters.fromYear)) return false;
  if (filters.toYear && year && year > String(filters.toYear)) return false;
  if (filters.month) {
    const monthMatch = date.match(/^\d{4}-(\d{2})/);
    const rowMonth = monthMatch ? monthMatch[1] : "";
    if (rowMonth !== String(filters.month)) return false;
  }

  const observerFilter = normalizeString(filters.observerFilter || "");
  if (observerFilter && !normalizeString(row.recordedBy).includes(observerFilter)) return false;

  return true;
}

async function getLocalDbStatus() {
  const bytes = await idbGet(IDB_KEY);
  if (!bytes) {
    return { exists: false, createdAt: null, rowCount: 0, fileSize: 0, inProgress: progressState.inProgress, progressPct: progressState.progressPct, progressText: progressState.progressText };
  }
  let createdAt = progressState.createdAt || null;
  let rowCount = progressState.rowCount || 0;
  try {
    const db = await loadActiveDb();
    if (db) {
      createdAt = getMeta(db, "createdAt") || createdAt;
      rowCount = Number(getMeta(db, "rowCount") || dbRowCount(db));
    }
  } catch (_) {}
  return {
    exists: true,
    createdAt,
    rowCount,
    fileSize: bytes.byteLength || bytes.length || 0,
    inProgress: progressState.inProgress,
    progressPct: progressState.progressPct,
    progressText: progressState.progressText || "Lokal databas klar"
  };
}

// ---- command handlers ----
const handlers = {
  async get_local_db_status() {
    return getLocalDbStatus();
  },

  async update_local_db_progress({ pct, text }) {
    progressState = { ...progressState, inProgress: true, progressPct: Number(pct || 0), progressText: String(text || "") };
    return progressState;
  },

  async prepare_local_db() {
    await ensureSqlJs();
    closeDb(buildDb);
    buildDb = new SQL.Database();
    buildDb.run(schemaSql());
    buildDb.run("BEGIN TRANSACTION");
    const existing = await idbGet(IDB_KEY);
    progressState = { exists: !!existing, createdAt: null, rowCount: 0, fileSize: 0, inProgress: true, progressPct: 0, progressText: "Förbereder lokal databas..." };
    return true;
  },

  async insert_local_rows_batch({ rows }) {
    if (!buildDb) throw new Error("Lokal databas är inte förberedd.");
    insertRowsInto(buildDb, rows);
    return true;
  },

  // ---- incremental update (sync only new/changed rows by `modified`) ----
  async get_sync_info() {
    const db = await loadActiveDb();
    if (!db) return { exists: false, lastModifiedSync: null };
    return { exists: true, lastModifiedSync: getMeta(db, "lastModifiedSync") };
  },

  async begin_local_update() {
    const db = await loadActiveDb();
    if (!db) throw new Error("Ingen lokal databas att uppdatera. Skapa den först.");
    try { db.run("ROLLBACK"); } catch (_) { /* no dangling transaction */ }
    db.run("BEGIN TRANSACTION");
    return true;
  },

  async rollback_local_update() {
    if (!activeDb) return false;
    try { activeDb.run("ROLLBACK"); } catch (_) {}
    // Reload from disk so the in-memory copy matches the persisted state.
    closeDb(activeDb);
    activeDb = null;
    await loadActiveDb();
    return true;
  },

  // Replace existing rows with the same occurrenceId (captures edits), then
  // insert the batch (captures new). Rows without an occurrenceId are appended.
  async upsert_local_rows_batch({ rows }) {
    const db = await loadActiveDb();
    if (!db) throw new Error("Ingen lokal databas att uppdatera.");
    const ids = [...new Set((rows || []).map((r) => r.occurrenceId).filter(Boolean))];
    if (ids.length) {
      const del = db.prepare("DELETE FROM observations WHERE occurrenceId = ?");
      for (const id of ids) del.run([id]);
      del.free();
    }
    insertRowsInto(db, rows);
    return true;
  },

  async commit_local_update({ lastModifiedSync } = {}) {
    const db = await loadActiveDb();
    if (!db) throw new Error("Ingen lokal databas att uppdatera.");
    db.run("COMMIT");
    const updatedAt = new Date().toLocaleString("sv-SE");
    const rowCount = dbRowCount(db);
    setMeta(db, "createdAt", updatedAt);
    setMeta(db, "rowCount", rowCount);
    if (lastModifiedSync) setMeta(db, "lastModifiedSync", lastModifiedSync);

    const bytes = db.export();
    await idbPut(IDB_KEY, bytes);

    progressState = { exists: true, createdAt: updatedAt, rowCount, fileSize: bytes.byteLength || bytes.length || 0, inProgress: false, progressPct: 100, progressText: "Lokal databas uppdaterad" };
    return progressState;
  },

  async finalize_local_db({ lastModifiedSync } = {}) {
    if (!buildDb) throw new Error("Ingen lokal databas att slutföra.");
    buildDb.run("COMMIT");
    buildDb.run("CREATE INDEX IF NOT EXISTS idx_species ON observations(species)");
    buildDb.run("CREATE INDEX IF NOT EXISTS idx_date ON observations(date)");
    buildDb.run("CREATE INDEX IF NOT EXISTS idx_province ON observations(province)");
    buildDb.run("CREATE INDEX IF NOT EXISTS idx_municipality ON observations(municipality)");
    buildDb.run("CREATE INDEX IF NOT EXISTS idx_occurrence ON observations(occurrenceId)");
    const createdAt = new Date().toLocaleString("sv-SE");
    const rowCount = dbRowCount(buildDb);
    setMeta(buildDb, "createdAt", createdAt);
    setMeta(buildDb, "rowCount", rowCount);
    if (lastModifiedSync) setMeta(buildDb, "lastModifiedSync", lastModifiedSync);

    const bytes = buildDb.export();
    await idbPut(IDB_KEY, bytes);
    closeDb(buildDb);
    buildDb = null;

    closeDb(activeDb);
    activeDb = null;
    await loadActiveDb();

    progressState = { exists: true, createdAt, rowCount, fileSize: bytes.byteLength || bytes.length || 0, inProgress: false, progressPct: 100, progressText: "Lokal databas klar" };
    return progressState;
  },

  async query_local_rows({ filters }) {
    const db = await loadActiveDb();
    if (!db) return [];

    // Pre-filter by species in SQL when available (dominant filter, exact stored
    // values) to avoid scanning the whole table; fine-grained filtering below.
    const speciesList = Array.isArray(filters?.speciesList) ? filters.speciesList.filter(Boolean) : [];
    let sql = `
      SELECT occurrenceId, species, date, province, municipality, locality, quantity,
             lifeStage, activity, recordedBy, latitude, longitude,
             isNeverFoundObservation, isNotRediscoveredObservation
      FROM observations
    `;
    let bindValues = [];
    if (speciesList.length > 0) {
      sql += ` WHERE species IN (${speciesList.map(() => "?").join(", ")})`;
      bindValues = speciesList;
    }

    const stmt = db.prepare(sql);
    if (bindValues.length) stmt.bind(bindValues);
    const rows = [];
    while (stmt.step()) {
      const row = stmt.getAsObject();
      const normalized = {
        occurrenceId: row.occurrenceId || "",
        species: row.species || "",
        date: row.date || "",
        province: row.province || "",
        municipality: row.municipality || "",
        locality: row.locality || "",
        quantity: row.quantity || "",
        lifeStage: row.lifeStage || "",
        activity: row.activity || "",
        recordedBy: row.recordedBy || "",
        latitude: row.latitude == null ? null : Number(row.latitude),
        longitude: row.longitude == null ? null : Number(row.longitude),
        isNeverFoundObservation: Boolean(row.isNeverFoundObservation),
        isNotRediscoveredObservation: Boolean(row.isNotRediscoveredObservation)
      };
      if (rowMatchesFilters(normalized, filters || {})) rows.push(normalized);
    }
    stmt.free();
    return rows;
  },

  async local_db_row_preview() {
    const db = await loadActiveDb();
    if (!db) return [];
    const stmt = db.prepare("SELECT species, date, province, municipality, recordedBy FROM observations LIMIT 5");
    const rows = [];
    while (stmt.step()) rows.push(stmt.getAsObject());
    stmt.free();
    return rows;
  }
};

export async function invoke(command, payload = {}) {
  const handler = handlers[command];
  if (!handler) throw new Error(`Okänt databaskommando: ${command}`);
  return handler(payload);
}
