/**
 * Survey receiver for "The Last Salvage Run" study (v3).
 *
 * POST               → appends one row per scenario to the "responses_v3" tab.
 * GET ?action=assign → returns the condition with the fewest completed
 *                      participants (ties broken at random) and reserves the
 *                      next sequential participant ID, e.g. P007_qual.
 * GET (no params)    → health check.
 *
 * Columns are taken from what the page sends: any new field the page adds
 * later gets its own column automatically instead of being rejected.
 *
 * After pasting this in, publish it with Deploy → Manage deployments →
 * (pencil) Edit → Version: New version → Deploy, so the /exec URL stays the same.
 */

const SHEET_NAME = "responses_v3";

const CONDITIONS = ["numerical", "qualitative", "no_information"];
const CONDITION_SUFFIX = { numerical: "num", qualitative: "qual", no_information: "none" };

// Sequential IDs (P001_num) come from this script; P + 10 random digits is the page's fallback.
const ID_RE =
    /^participant_(numerical|qualitative|no_information)_[1-9]\d*$/;

const MAX_BODY_CHARS = 200000;
const MAX_ROWS = 20;
const MAX_COLUMNS = 150;
const MAX_TEXT = 2000;
const COUNTER_PREFIX = "next_participant_number_";

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    const body = e && e.postData && e.postData.contents;
    if (!body || body.length > MAX_BODY_CHARS) throw new Error("Empty or oversized request");

    const data = JSON.parse(body);
    validate_(data);

    // Serialize writes so simultaneous participants can't interleave rows.
    lock.waitLock(20000);
    const sheet = getSheet_();
    const headers = ensureHeaders_(sheet, Object.keys(data.rows[0]));

    // A retry from the same browser reuses the same participant_id; don't store it twice.
    if (alreadySubmitted_(sheet, headers, data.participant_id)) {
      return json_({ ok: true, duplicate: true });
    }

    const receivedAt = new Date();
    const values = data.rows.map(r => headers.map(h => {
      if (h === "received_at") return receivedAt;
      if (h === "participant_id") return data.participant_id;
      return clean_(r[h]);
    }));
    sheet.getRange(sheet.getLastRow() + 1, 1, values.length, headers.length).setValues(values);
    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: String((err && err.message) || err) });
  } finally {
    lock.releaseLock();
  }
}

function doGet(e) {
  const action = e && e.parameter && e.parameter.action;
  if (action === "assign") {
    const lock = LockService.getScriptLock();
    try {
      lock.waitLock(10000);
      const condition = leastFilledCondition_();
      const number = reserveParticipantNumber_(condition);
      return json_({
        ok: true,
        condition: condition,
        participant_number: number,
        participant_id: `participant_${condition}_${number}`
      });
    } catch (err) {
      return json_({ ok: false, error: String((err && err.message) || err) });
    } finally {
      lock.releaseLock();
    }
  }
  return json_({ ok: true, message: "Survey endpoint is running." });
}

// Each assignment request takes the next number, so people who quit early leave gaps (P004, P006…).

function reserveParticipantNumber_(condition) {
    if (!CONDITIONS.includes(condition)) {
        throw new Error("Invalid condition");
    }

    const props = PropertiesService.getScriptProperties();
    const key = COUNTER_PREFIX + condition;

    const next = Number(props.getProperty(key) || "1");

    props.setProperty(key, String(next + 1));

    return next;
}

// Counts completed participants per condition and returns one of the smallest groups.
function leastFilledCondition_() {
  const counts = {};
  CONDITIONS.forEach(c => { counts[c] = 0; });

  const sheet = getSheet_();
  if (sheet.getLastRow() >= 2 && sheet.getLastColumn() > 0) {
    const all = sheet.getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn()).getValues();
    const idCol = all[0].indexOf("participant_id");
    const condCol = all[0].indexOf("experimental_condition");
    if (idCol >= 0 && condCol >= 0) {
      const seen = {};
      all.slice(1).forEach(row => {
        const id = row[idCol];
        const cond = row[condCol];
        if (!id || seen[id] || !(cond in counts)) return;
        seen[id] = true;
        counts[cond] += 1;
      });
    }
  }

  const min = Math.min.apply(null, CONDITIONS.map(c => counts[c]));
  const tied = CONDITIONS.filter(c => counts[c] === min);
  return tied[Math.floor(Math.random() * tied.length)];
}

function validate_(data) {
  if (!data || typeof data !== "object") throw new Error("Invalid JSON");
  const id = String(data.participant_id || "");
  if (!ID_RE.test(id)) throw new Error("Invalid participant_id");
  if (!Array.isArray(data.rows) || data.rows.length === 0 || data.rows.length > MAX_ROWS) {
    throw new Error("Invalid rows");
  }
  data.rows.forEach(r => {
    if (!r || typeof r !== "object") throw new Error("Invalid row");
    if (Object.keys(r).length > MAX_COLUMNS) throw new Error("Too many fields");
    if (!["A", "B", "C", "D"].includes(r.choice)) throw new Error("Invalid choice");
    if (!CONDITIONS.includes(r.experimental_condition)) throw new Error("Invalid condition");
    if (!id.startsWith(
    `participant_${r.experimental_condition}_`
    )) {
        throw new Error("participant_id does not match condition");
    }
  });
}

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME);
  return sheet;
}

// Returns the header row, adding received_at first and any field names not yet present at the end.
function ensureHeaders_(sheet, fieldNames) {
  const width = sheet.getLastColumn();
  let headers = width > 0
    ? sheet.getRange(1, 1, 1, width).getValues()[0].map(String).filter(h => h !== "")
    : [];
  if (headers.length === 0) headers = ["received_at"];

  const missing = fieldNames
    .map(name => String(name).replace(/[^A-Za-z0-9_]/g, "").slice(0, 60))
    .filter(name => name && headers.indexOf(name) === -1);

  if (missing.length > 0 || width === 0) {
    headers = headers.concat(missing);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
  }
  return headers;
}

function alreadySubmitted_(sheet, headers, participantId) {
  if (sheet.getLastRow() < 2) return false;
  const col = headers.indexOf("participant_id") + 1;
  if (col === 0) return false;
  const found = sheet
    .getRange(2, col, sheet.getLastRow() - 1, 1)
    .createTextFinder(participantId)
    .matchEntireCell(true)
    .findNext();
  return found !== null;
}

// Numbers and booleans are stored as-is; numeric strings (e.g. "12.345") become numbers;
// other text is length-limited and kept from being run as a formula.
function clean_(value) {
  if (value === null || value === undefined) return "";
  if (typeof value === "boolean") return value;
  if (typeof value === "number") return Number.isFinite(value) ? value : "";
  let s = String(value).slice(0, MAX_TEXT);
  if (/^-?\d+(\.\d+)?$/.test(s)) return Number(s);
  if (/^[=+\-@]/.test(s)) s = "'" + s;
  return s;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}


function initializeConditionCounters() {
  const sheet = getSheet_();
  const data = sheet.getDataRange().getValues();

  const headers = data[0];
  const idIndex = headers.indexOf("participant_id");
  const conditionIndex =
    headers.indexOf("experimental_condition");

  if (idIndex === -1 || conditionIndex === -1) {
    throw new Error("Participant ID or condition column missing");
  }

  const props = PropertiesService.getScriptProperties();

  CONDITIONS.forEach(condition => {
    const key = COUNTER_PREFIX + condition;

    // Do not overwrite an existing counter.
    if (props.getProperty(key) !== null) return;

    const ids = new Set(
      data.slice(1)
        .filter(row => row[conditionIndex] === condition)
        .map(row => String(row[idIndex]))
        .filter(Boolean)
    );

    props.setProperty(key, String(ids.size + 1));

    Logger.log(
      condition + " next participant number: " +
      (ids.size + 1)
    );
  });
}
