/**
 * さがせ！時間泥棒 → Googleスプレッドシート 送受信スクリプト
 *
 * 対象スプレッドシート:
 * https://docs.google.com/spreadsheets/d/1MRtZVSAwAokFMAZOSRIlwFz8m7qr5HOvDi8gKPNLe0E/edit
 *
 * 同一ニックネームは1行に上書き（重複行は削除）
 */

var SPREADSHEET_ID = '1MRtZVSAwAokFMAZOSRIlwFz8m7qr5HOvDi8gKPNLe0E';

var CHOICE_LABELS = {
  love: '大好き',
  like: '好き',
  neutral: '普通',
  dislike: '嫌い',
  hate: '大嫌い',
  unassigned: '担当外'
};

var LABEL_TO_CODE = {
  '大好き': 'love',
  '好き': 'like',
  '普通': 'neutral',
  '嫌い': 'dislike',
  '大嫌い': 'hate',
  '担当外': 'unassigned'
};

var EXPERIENCE_LABELS = {
  '1-5': '1〜5年',
  '6-10': '6〜10年',
  '11-20': '11〜20年',
  '21+': '21年以上'
};

function experienceLabel_(codeOrLabel) {
  var s = String(codeOrLabel || '').trim();
  if (!s) return '';
  if (EXPERIENCE_LABELS[s]) return EXPERIENCE_LABELS[s];
  // すでに日本語ラベルならそのまま
  var keys = Object.keys(EXPERIENCE_LABELS);
  for (var i = 0; i < keys.length; i++) {
    if (EXPERIENCE_LABELS[keys[i]] === s) return s;
  }
  return s;
}

function experienceCode_(codeOrLabel) {
  var s = String(codeOrLabel || '').trim();
  if (!s) return '';
  if (EXPERIENCE_LABELS[s]) return s;
  var keys = Object.keys(EXPERIENCE_LABELS);
  for (var i = 0; i < keys.length; i++) {
    if (EXPERIENCE_LABELS[keys[i]] === s) return keys[i];
  }
  return s;
}

function doGet(e) {
  try {
    var action = (e && e.parameter && e.parameter.action) || 'ping';
    if (action === 'export' || action === 'all') {
      return json_(exportAll_());
    }
    if (action === 'dedupe') {
      var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
      return json_(dedupeAllSheets_(ss));
    }
    return json_({
      ok: true,
      app: 'さがせ！時間泥棒',
      message: 'OK。action=export で全回答を取得できます。'
    });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}

function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      throw new Error('リクエスト本文が空です');
    }

    var data = JSON.parse(e.postData.contents);
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var action = data.action || 'sync_all';

    if (action === 'sync_user' || action === 'upsert_user') {
      upsertUser_(ss, data);
    } else if (action === 'dedupe') {
      return json_(dedupeAllSheets_(ss));
    } else if (action === 'export') {
      return json_(exportAll_());
    } else {
      syncAll_(ss, data);
    }

    return json_({ ok: true, action: action, at: new Date().toISOString() });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}

function exportAll_() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  // 読み取り前に重複があれば整理（分析側が二重カウントしないように）
  dedupeNicknameSheet_(ss, '回答データ', 2);
  dedupeNicknameSheet_(ss, 'マトリクス', 2);

  var fromRaw = readRawResponses_(ss);
  var responses = fromRaw.responses;
  var source = '回答データ';

  if (!responses || Object.keys(responses).length === 0) {
    responses = readMatrixResponses_(ss);
    source = 'マトリクス';
  }

  return {
    ok: true,
    source: source,
    exportedAt: new Date().toISOString(),
    respondentCount: Object.keys(responses).length,
    responses: responses,
    diagnoses: fromRaw.diagnoses || {},
    profiles: fromRaw.profiles || {}
  };
}

function syncAll_(ss, data) {
  var tasks = normalizeTasks_(data.tasks);
  var responses = data.responses || {};
  var profiles = data.profiles || {};
  // キー正規化して同一ニックネームをマージ
  var normalized = {};
  var diagnoses = data.diagnoses || {};
  var diagNorm = {};
  var profileNorm = {};
  Object.keys(responses).forEach(function (user) {
    var key = normalizeNickname_(user);
    if (!key) return;
    normalized[key] = normalizeAnswers_(responses[user] || {});
    if (diagnoses[user]) diagNorm[key] = diagnoses[user];
    var p = profiles[user] || profiles[key] || {};
    profileNorm[key] = {
      experienceYears: experienceCode_(p.experienceYears || p.experienceLabel || ''),
      experienceLabel: experienceLabel_(p.experienceLabel || p.experienceYears || '')
    };
  });
  writeRawSheet_(ss, normalized, diagNorm, profileNorm);
  writeMatrixSheet_(ss, tasks, normalized, diagNorm, profileNorm);
  writeAnswersSheet_(ss, tasks, normalized);
  writeMetaSheet_(ss, data);
}

function upsertUser_(ss, data) {
  var tasks = normalizeTasks_(data.tasks);
  var nickname = normalizeNickname_(data.nickname);
  if (!nickname) throw new Error('nickname が空です');

  var answers = normalizeAnswers_(data.answers || {});
  var diagnosis = data.diagnosis || {};
  var experienceYears = experienceCode_(data.experienceYears || data.experienceLabel || '');
  var experienceLabel = experienceLabel_(data.experienceLabel || data.experienceYears || experienceYears);
  var profile = { experienceYears: experienceYears, experienceLabel: experienceLabel };

  upsertRawRow_(ss, nickname, answers, diagnosis, profile);
  upsertMatrixRow_(ss, tasks, nickname, answers, diagnosis, profile);
  replaceAnswersForUser_(ss, tasks, nickname, answers);
  writeMetaSheet_(ss, {
    exportedAt: data.exportedAt || new Date().toISOString(),
    note: '完了送信（上書き）: ' + nickname
  });
}

/** ニックネーム正規化（前後空白除去・全角半角寄せ・連続空白つぶし） */
function normalizeNickname_(name) {
  var s = String(name == null ? '' : name).trim();
  if (!s) return '';
  try {
    if (typeof s.normalize === 'function') s = s.normalize('NFKC');
  } catch (e) { /* ignore */ }
  s = s.replace(/\s+/g, ' ').trim();
  return s;
}

/** 機械可読な正本シート（分析用にアプリが読み戻す） */
function writeRawSheet_(ss, responses, diagnoses, profiles) {
  var sheet = getOrCreateSheet_(ss, '回答データ');
  sheet.clear();
  sheet.getRange(1, 1, 1, 6).setValues([[
    '更新日時', '回答者', '経験年数', '回答JSON', '診断タイプ', 'キャッチコピー'
  ]]);
  sheet.setFrozenRows(1);

  var now = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss');
  var rows = [];
  Object.keys(responses || {}).forEach(function (user) {
    var key = normalizeNickname_(user);
    if (!key) return;
    var diag = diagnoses[key] || diagnoses[user] || {};
    var profile = (profiles && (profiles[key] || profiles[user])) || {};
    rows.push([
      now,
      key,
      experienceLabel_(profile.experienceLabel || profile.experienceYears || ''),
      JSON.stringify(normalizeAnswers_(responses[user] || {})),
      diag.name || '',
      diag.tagline || ''
    ]);
  });
  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, 6).setValues(rows);
  }
}

function ensureRawHeader_(sheet) {
  var target = ['更新日時', '回答者', '経験年数', '回答JSON', '診断タイプ', 'キャッチコピー'];
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, target.length).setValues([target]);
    sheet.setFrozenRows(1);
    return;
  }
  var width = Math.max(sheet.getLastColumn(), target.length);
  var header = sheet.getRange(1, 1, 1, width).getValues()[0].map(function (v) { return String(v || ''); });
  // 旧形式（経験年数なし）→ 列を挿入して移行
  if (header[2] !== '経験年数') {
    if (header[2] === '回答JSON') {
      sheet.insertColumnAfter(2);
      sheet.getRange(1, 3).setValue('経験年数');
    } else {
      sheet.getRange(1, 1, 1, target.length).setValues([target]);
    }
    sheet.setFrozenRows(1);
  }
}

function upsertRawRow_(ss, nickname, answers, diagnosis, profile) {
  var sheet = getOrCreateSheet_(ss, '回答データ');
  ensureRawHeader_(sheet);

  var now = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss');
  var row = [
    now,
    nickname,
    experienceLabel_((profile && (profile.experienceLabel || profile.experienceYears)) || ''),
    JSON.stringify(normalizeAnswers_(answers)),
    (diagnosis && diagnosis.name) || '',
    (diagnosis && diagnosis.tagline) || ''
  ];

  var matchRows = findNicknameRows_(sheet, 2, nickname); // 1-based sheet rows
  if (matchRows.length > 0) {
    sheet.getRange(matchRows[0], 1, 1, 6).setValues([row]);
    for (var d = matchRows.length - 1; d >= 1; d--) {
      sheet.deleteRow(matchRows[d]);
    }
    return;
  }
  sheet.appendRow(row);
}

function readRawResponses_(ss) {
  var sheet = ss.getSheetByName('回答データ');
  var result = { responses: {}, diagnoses: {}, profiles: {} };
  if (!sheet || sheet.getLastRow() < 2) return result;

  var values = sheet.getDataRange().getValues();
  var header = values[0].map(function (v) { return String(v || ''); });
  var hasExperience = header[2] === '経験年数';
  var jsonCol = hasExperience ? 3 : 2;
  var diagNameCol = hasExperience ? 4 : 3;
  var diagTagCol = hasExperience ? 5 : 4;
  var expCol = hasExperience ? 2 : -1;

  for (var r = 1; r < values.length; r++) {
    var nickname = normalizeNickname_(values[r][1]);
    if (!nickname) continue;
    var jsonText = values[r][jsonCol];
    try {
      var parsed = typeof jsonText === 'string' ? JSON.parse(jsonText) : (jsonText || {});
      // 後勝ち（最新行を優先）
      result.responses[nickname] = normalizeAnswers_(parsed);
    } catch (e) {
      result.responses[nickname] = {};
    }
    result.diagnoses[nickname] = {
      name: String(values[r][diagNameCol] || ''),
      tagline: String(values[r][diagTagCol] || '')
    };
    if (expCol >= 0) {
      var expLabel = String(values[r][expCol] || '').trim();
      result.profiles[nickname] = {
        experienceYears: experienceCode_(expLabel),
        experienceLabel: experienceLabel_(expLabel)
      };
    }
  }
  return result;
}

function readMatrixResponses_(ss) {
  var sheet = ss.getSheetByName('マトリクス');
  var responses = {};
  if (!sheet || sheet.getLastRow() < 2) return responses;

  var values = sheet.getDataRange().getValues();
  var header = values[0];
  var taskCols = [];
  for (var c = 0; c < header.length; c++) {
    var h = String(header[c] || '');
    var m = h.match(/^(\d+)_/);
    if (m) taskCols.push({ col: c, id: Number(m[1]) });
  }

  for (var r = 1; r < values.length; r++) {
    var nickname = normalizeNickname_(values[r][1]);
    if (!nickname) continue;
    var answers = {};
    taskCols.forEach(function (tc) {
      var label = String(values[r][tc.col] || '').trim();
      var code = LABEL_TO_CODE[label];
      if (code) answers[String(tc.id)] = code;
    });
    responses[nickname] = answers;
  }
  return responses;
}

function writeMatrixSheet_(ss, tasks, responses, diagnoses, profiles) {
  var sheet = getOrCreateSheet_(ss, 'マトリクス');
  sheet.clear();

  var header = ['更新日時', '回答者', '経験年数', '診断タイプ', 'キャッチコピー'];
  tasks.forEach(function (t) {
    header.push(t.id + '_' + t.name);
  });
  sheet.getRange(1, 1, 1, header.length).setValues([header]);
  sheet.setFrozenRows(1);

  var userKeys = Object.keys(responses || {});
  if (userKeys.length === 0) return;

  var now = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss');
  var rows = userKeys.map(function (user) {
    var key = normalizeNickname_(user);
    var answers = responses[user] || {};
    var diag = diagnoses[key] || diagnoses[user] || {};
    var profile = (profiles && (profiles[key] || profiles[user])) || {};
    var row = [
      now,
      key,
      experienceLabel_(profile.experienceLabel || profile.experienceYears || ''),
      diag.name || '',
      diag.tagline || ''
    ];
    tasks.forEach(function (t) {
      var val = answers[String(t.id)] != null ? answers[String(t.id)] : answers[t.id];
      row.push(labelChoice_(val));
    });
    return row;
  });

  sheet.getRange(2, 1, rows.length, header.length).setValues(rows);
}

function upsertMatrixRow_(ss, tasks, nickname, answers, diagnosis, profile) {
  var sheet = getOrCreateSheet_(ss, 'マトリクス');
  ensureMatrixHeader_(sheet, tasks);

  var now = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss');
  var row = [
    now,
    nickname,
    experienceLabel_((profile && (profile.experienceLabel || profile.experienceYears)) || ''),
    (diagnosis && diagnosis.name) || '',
    (diagnosis && diagnosis.tagline) || ''
  ];
  tasks.forEach(function (t) {
    var val = answers[String(t.id)] != null ? answers[String(t.id)] : answers[t.id];
    row.push(labelChoice_(val));
  });

  var matchRows = findNicknameRows_(sheet, 2, nickname);
  if (matchRows.length > 0) {
    sheet.getRange(matchRows[0], 1, 1, row.length).setValues([row]);
    for (var d = matchRows.length - 1; d >= 1; d--) {
      sheet.deleteRow(matchRows[d]);
    }
    return;
  }
  sheet.appendRow(row);
}

function ensureMatrixHeader_(sheet, tasks) {
  var header = ['更新日時', '回答者', '経験年数', '診断タイプ', 'キャッチコピー'];
  tasks.forEach(function (t) {
    header.push(t.id + '_' + t.name);
  });

  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, header.length).setValues([header]);
    sheet.setFrozenRows(1);
    return;
  }

  var width = Math.max(header.length, sheet.getLastColumn());
  var existing = sheet.getRange(1, 1, 1, width).getValues()[0];
  var needRewrite = false;
  for (var i = 0; i < header.length; i++) {
    if (String(existing[i] || '') !== header[i]) {
      needRewrite = true;
      break;
    }
  }
  if (needRewrite) {
    // 旧ヘッダーなら経験年数列を挿入
    if (String(existing[2] || '') !== '経験年数' && String(existing[2] || '').indexOf('_') === -1) {
      // existing[2] が診断タイプなどの場合
      if (String(existing[2] || '') === '診断タイプ') {
        sheet.insertColumnAfter(2);
      }
    }
    sheet.getRange(1, 1, 1, header.length).setValues([header]);
    sheet.setFrozenRows(1);
  }
}

function writeAnswersSheet_(ss, tasks, responses) {
  var sheet = getOrCreateSheet_(ss, '回答ログ');
  sheet.clear();
  sheet.getRange(1, 1, 1, 6).setValues([[
    '記録日時', '回答者', '業務No', '業務名', '回答コード', '回答ラベル'
  ]]);
  sheet.setFrozenRows(1);

  var now = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss');
  var rows = [];
  Object.keys(responses || {}).forEach(function (user) {
    var key = normalizeNickname_(user);
    if (!key) return;
    var answers = responses[user] || {};
    tasks.forEach(function (t) {
      var val = answers[String(t.id)] != null ? answers[String(t.id)] : answers[t.id];
      if (!val) return;
      rows.push([now, key, t.id, t.name, val, labelChoice_(val)]);
    });
  });

  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, 6).setValues(rows);
  }
}

/** 同一ニックネームの旧ログを消してから最新回答だけを書く */
function replaceAnswersForUser_(ss, tasks, nickname, answers) {
  var sheet = getOrCreateSheet_(ss, '回答ログ');
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, 6).setValues([[
      '記録日時', '回答者', '業務No', '業務名', '回答コード', '回答ラベル'
    ]]);
    sheet.setFrozenRows(1);
  }

  // 既存の同名行を下から削除
  var matchRows = findNicknameRows_(sheet, 2, nickname);
  for (var d = matchRows.length - 1; d >= 0; d--) {
    sheet.deleteRow(matchRows[d]);
  }

  var now = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss');
  var rows = [];
  tasks.forEach(function (t) {
    var val = answers[String(t.id)] != null ? answers[String(t.id)] : answers[t.id];
    if (!val) return;
    rows.push([now, nickname, t.id, t.name, val, labelChoice_(val)]);
  });
  if (rows.length > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, 6).setValues(rows);
  }
}

/**
 * 指定シートで回答者列が nickname と一致する行番号（1-based）を返す
 * @param {Sheet} sheet
 * @param {number} nicknameCol 1-based column index
 * @param {string} nickname 正規化済み
 */
function findNicknameRows_(sheet, nicknameCol, nickname) {
  var lastRow = sheet.getLastRow();
  var matches = [];
  if (lastRow < 2) return matches;

  var numRows = lastRow - 1;
  var names = sheet.getRange(2, nicknameCol, numRows, 1).getValues();
  for (var i = 0; i < names.length; i++) {
    if (normalizeNickname_(names[i][0]) === nickname) {
      matches.push(i + 2);
    }
  }
  return matches;
}

/** 同一ニックネームが複数行ある場合、最終行だけ残して削除 */
function dedupeNicknameSheet_(ss, sheetName, nicknameCol) {
  var sheet = ss.getSheetByName(sheetName);
  var removed = 0;
  if (!sheet || sheet.getLastRow() < 2) {
    return { sheet: sheetName, removed: 0 };
  }

  var lastRow = sheet.getLastRow();
  var numRows = lastRow - 1;
  var names = sheet.getRange(2, nicknameCol, numRows, 1).getValues();
  var seen = {}; // nickname -> last row index (1-based)
  var toDelete = [];

  for (var i = 0; i < names.length; i++) {
    var key = normalizeNickname_(names[i][0]);
    if (!key) continue;
    var rowNum = i + 2;
    if (seen[key] != null) {
      // 以前の行を削除候補に（後勝ち）
      toDelete.push(seen[key]);
    }
    seen[key] = rowNum;
  }

  toDelete.sort(function (a, b) { return b - a; });
  for (var d = 0; d < toDelete.length; d++) {
    sheet.deleteRow(toDelete[d]);
    removed++;
  }
  return { sheet: sheetName, removed: removed };
}

/** 回答ログはニックネーム×業務No で後勝ち */
function dedupeAnswersLog_(ss) {
  var sheet = ss.getSheetByName('回答ログ');
  var removed = 0;
  if (!sheet || sheet.getLastRow() < 2) {
    return { sheet: '回答ログ', removed: 0 };
  }

  var values = sheet.getDataRange().getValues();
  var header = values[0];
  var keep = {}; // key -> row array
  var order = [];

  for (var r = 1; r < values.length; r++) {
    var nick = normalizeNickname_(values[r][1]);
    var taskNo = String(values[r][2] == null ? '' : values[r][2]);
    if (!nick || !taskNo) continue;
    var key = nick + '\t' + taskNo;
    if (!keep[key]) order.push(key);
    keep[key] = [
      values[r][0],
      nick,
      values[r][2],
      values[r][3],
      values[r][4],
      values[r][5]
    ];
  }

  var originalDataRows = values.length - 1;
  var rows = order.map(function (k) { return keep[k]; });
  removed = Math.max(0, originalDataRows - rows.length);

  sheet.clear();
  sheet.getRange(1, 1, 1, header.length).setValues([header]);
  sheet.setFrozenRows(1);
  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, 6).setValues(rows);
  }
  return { sheet: '回答ログ', removed: removed };
}

function dedupeAllSheets_(ss) {
  var a = dedupeNicknameSheet_(ss, '回答データ', 2);
  var b = dedupeNicknameSheet_(ss, 'マトリクス', 2);
  var c = dedupeAnswersLog_(ss);
  writeMetaSheet_(ss, {
    exportedAt: new Date().toISOString(),
    note: '重複整理: 回答データ-' + a.removed + ' / マトリクス-' + b.removed + ' / 回答ログ-' + c.removed
  });
  return {
    ok: true,
    action: 'dedupe',
    results: [a, b, c],
    at: new Date().toISOString()
  };
}

function writeMetaSheet_(ss, data) {
  var sheet = getOrCreateSheet_(ss, 'メタ');
  sheet.clear();
  sheet.getRange(1, 1, 4, 2).setValues([
    ['最終受信日時', Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss')],
    ['クライアント書き出し日時', data.exportedAt || ''],
    ['メモ', data.note || ''],
    ['アプリ', 'さがせ！時間泥棒']
  ]);
}

function getOrCreateSheet_(ss, name) {
  var sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  return sheet;
}

function normalizeTasks_(tasks) {
  if (!tasks || !tasks.length) {
    var dummy = [];
    for (var i = 1; i <= 50; i++) {
      dummy.push({ id: i, name: '業務' + i, desc: '' });
    }
    return dummy;
  }
  return tasks.map(function (t, idx) {
    return {
      id: Number(t.id) || (idx + 1),
      name: String(t.name || ('業務' + (idx + 1))),
      desc: String(t.desc || '')
    };
  });
}

function normalizeAnswers_(answers) {
  var out = {};
  Object.keys(answers || {}).forEach(function (k) {
    var v = answers[k];
    if (!v) return;
    out[String(k)] = String(v);
  });
  return out;
}

function labelChoice_(val) {
  if (!val) return '';
  return CHOICE_LABELS[val] || String(val);
}

function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
