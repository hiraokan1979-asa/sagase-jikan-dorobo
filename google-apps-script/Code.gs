/**
 * さがせ！時間泥棒 → Googleスプレッドシート 送受信スクリプト
 *
 * 対象スプレッドシート:
 * https://docs.google.com/spreadsheets/d/1MRtZVSAwAokFMAZOSRIlwFz8m7qr5HOvDi8gKPNLe0E/edit
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

function doGet(e) {
  try {
    var action = (e && e.parameter && e.parameter.action) || 'ping';
    if (action === 'export' || action === 'all') {
      return json_(exportAll_());
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
    diagnoses: fromRaw.diagnoses || {}
  };
}

function syncAll_(ss, data) {
  var tasks = normalizeTasks_(data.tasks);
  var responses = data.responses || {};
  writeRawSheet_(ss, responses, data.diagnoses || {});
  writeMatrixSheet_(ss, tasks, responses, data.diagnoses || {});
  writeAnswersSheet_(ss, tasks, responses);
  writeMetaSheet_(ss, data);
}

function upsertUser_(ss, data) {
  var tasks = normalizeTasks_(data.tasks);
  var nickname = String(data.nickname || '').trim();
  if (!nickname) throw new Error('nickname が空です');

  var answers = normalizeAnswers_(data.answers || {});
  var diagnosis = data.diagnosis || {};

  upsertRawRow_(ss, nickname, answers, diagnosis);
  upsertMatrixRow_(ss, tasks, nickname, answers, diagnosis);
  appendAnswersForUser_(ss, tasks, nickname, answers);
  writeMetaSheet_(ss, {
    exportedAt: data.exportedAt || new Date().toISOString(),
    note: '完了送信: ' + nickname
  });
}

/** 機械可読な正本シート（分析用にアプリが読み戻す） */
function writeRawSheet_(ss, responses, diagnoses) {
  var sheet = getOrCreateSheet_(ss, '回答データ');
  sheet.clear();
  sheet.getRange(1, 1, 1, 5).setValues([[
    '更新日時', '回答者', '回答JSON', '診断タイプ', 'キャッチコピー'
  ]]);
  sheet.setFrozenRows(1);

  var now = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss');
  var rows = [];
  Object.keys(responses || {}).forEach(function (user) {
    var diag = diagnoses[user] || {};
    rows.push([
      now,
      user,
      JSON.stringify(normalizeAnswers_(responses[user] || {})),
      diag.name || '',
      diag.tagline || ''
    ]);
  });
  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, 5).setValues(rows);
  }
}

function upsertRawRow_(ss, nickname, answers, diagnosis) {
  var sheet = getOrCreateSheet_(ss, '回答データ');
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, 5).setValues([[
      '更新日時', '回答者', '回答JSON', '診断タイプ', 'キャッチコピー'
    ]]);
    sheet.setFrozenRows(1);
  }

  var now = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss');
  var row = [
    now,
    nickname,
    JSON.stringify(normalizeAnswers_(answers)),
    (diagnosis && diagnosis.name) || '',
    (diagnosis && diagnosis.tagline) || ''
  ];

  var lastRow = sheet.getLastRow();
  if (lastRow >= 2) {
    var names = sheet.getRange(2, 2, lastRow - 1, 1).getValues();
    for (var i = 0; i < names.length; i++) {
      if (String(names[i][0]) === nickname) {
        sheet.getRange(i + 2, 1, 1, 5).setValues([row]);
        return;
      }
    }
  }
  sheet.appendRow(row);
}

function readRawResponses_(ss) {
  var sheet = ss.getSheetByName('回答データ');
  var result = { responses: {}, diagnoses: {} };
  if (!sheet || sheet.getLastRow() < 2) return result;

  var values = sheet.getDataRange().getValues();
  for (var r = 1; r < values.length; r++) {
    var nickname = String(values[r][1] || '').trim();
    if (!nickname) continue;
    var jsonText = values[r][2];
    try {
      var parsed = typeof jsonText === 'string' ? JSON.parse(jsonText) : (jsonText || {});
      result.responses[nickname] = normalizeAnswers_(parsed);
    } catch (e) {
      result.responses[nickname] = {};
    }
    result.diagnoses[nickname] = {
      name: String(values[r][3] || ''),
      tagline: String(values[r][4] || '')
    };
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
  for (var c = 4; c < header.length; c++) {
    var h = String(header[c] || '');
    var m = h.match(/^(\d+)_/);
    if (m) taskCols.push({ col: c, id: Number(m[1]) });
  }

  for (var r = 1; r < values.length; r++) {
    var nickname = String(values[r][1] || '').trim();
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

function writeMatrixSheet_(ss, tasks, responses, diagnoses) {
  var sheet = getOrCreateSheet_(ss, 'マトリクス');
  sheet.clear();

  var header = ['更新日時', '回答者', '診断タイプ', 'キャッチコピー'];
  tasks.forEach(function (t) {
    header.push(t.id + '_' + t.name);
  });
  sheet.getRange(1, 1, 1, header.length).setValues([header]);
  sheet.setFrozenRows(1);

  var userKeys = Object.keys(responses || {});
  if (userKeys.length === 0) return;

  var now = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss');
  var rows = userKeys.map(function (user) {
    var answers = responses[user] || {};
    var diag = diagnoses[user] || {};
    var row = [now, user, diag.name || '', diag.tagline || ''];
    tasks.forEach(function (t) {
      var val = answers[String(t.id)] != null ? answers[String(t.id)] : answers[t.id];
      row.push(labelChoice_(val));
    });
    return row;
  });

  sheet.getRange(2, 1, rows.length, header.length).setValues(rows);
}

function upsertMatrixRow_(ss, tasks, nickname, answers, diagnosis) {
  var sheet = getOrCreateSheet_(ss, 'マトリクス');
  ensureMatrixHeader_(sheet, tasks);

  var now = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss');
  var row = [
    now,
    nickname,
    (diagnosis && diagnosis.name) || '',
    (diagnosis && diagnosis.tagline) || ''
  ];
  tasks.forEach(function (t) {
    var val = answers[String(t.id)] != null ? answers[String(t.id)] : answers[t.id];
    row.push(labelChoice_(val));
  });

  var lastRow = sheet.getLastRow();
  if (lastRow >= 2) {
    var names = sheet.getRange(2, 2, lastRow - 1, 1).getValues();
    for (var i = 0; i < names.length; i++) {
      if (String(names[i][0]) === nickname) {
        sheet.getRange(i + 2, 1, 1, row.length).setValues([row]);
        return;
      }
    }
  }
  sheet.appendRow(row);
}

function ensureMatrixHeader_(sheet, tasks) {
  var header = ['更新日時', '回答者', '診断タイプ', 'キャッチコピー'];
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
    var answers = responses[user] || {};
    tasks.forEach(function (t) {
      var val = answers[String(t.id)] != null ? answers[String(t.id)] : answers[t.id];
      if (!val) return;
      rows.push([now, user, t.id, t.name, val, labelChoice_(val)]);
    });
  });

  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, 6).setValues(rows);
  }
}

function appendAnswersForUser_(ss, tasks, nickname, answers) {
  var sheet = getOrCreateSheet_(ss, '回答ログ');
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, 6).setValues([[
      '記録日時', '回答者', '業務No', '業務名', '回答コード', '回答ラベル'
    ]]);
    sheet.setFrozenRows(1);
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
