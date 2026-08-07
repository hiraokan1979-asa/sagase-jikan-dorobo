/**
 * さがせ！時間泥棒 → Googleスプレッドシート受信スクリプト
 *
 * 対象スプレッドシート:
 * https://docs.google.com/spreadsheets/d/1MRtZVSAwAokFMAZOSRIlwFz8m7qr5HOvDi8gKPNLe0E/edit
 *
 * セットアップは同フォルダの「セットアップ手順.txt」を参照。
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

function doGet() {
  return json_({
    ok: true,
    app: 'さがせ！時間泥棒',
    message: 'このURLは有効です。アプリからPOSTでデータを送信してください。'
  });
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
    } else {
      syncAll_(ss, data);
    }

    return json_({ ok: true, action: action, at: new Date().toISOString() });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}

function syncAll_(ss, data) {
  var tasks = normalizeTasks_(data.tasks);
  var responses = data.responses || {};
  writeMatrixSheet_(ss, tasks, responses, data.diagnoses || {});
  writeAnswersSheet_(ss, tasks, responses);
  writeMetaSheet_(ss, data);
}

function upsertUser_(ss, data) {
  var tasks = normalizeTasks_(data.tasks);
  var nickname = String(data.nickname || '').trim();
  if (!nickname) throw new Error('nickname が空です');

  var answers = data.answers || {};
  var responses = {};
  responses[nickname] = answers;

  var diagnoses = {};
  if (data.diagnosis) {
    diagnoses[nickname] = data.diagnosis;
  }

  // マトリクスは既存行を残しつつ当該回答者だけ更新
  upsertMatrixRow_(ss, tasks, nickname, answers, diagnoses[nickname] || {});
  appendAnswersForUser_(ss, tasks, nickname, answers);
  writeMetaSheet_(ss, {
    exportedAt: data.exportedAt || new Date().toISOString(),
    note: '単一回答者の更新: ' + nickname
  });
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
    var row = [
      now,
      user,
      diag.name || '',
      diag.tagline || ''
    ];
    tasks.forEach(function (t) {
      var key = String(t.id);
      var val = answers[key] != null ? answers[key] : answers[t.id];
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
    var key = String(t.id);
    var val = answers[key] != null ? answers[key] : answers[t.id];
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

  var existing = sheet.getRange(1, 1, 1, header.length).getValues()[0];
  var needRewrite = existing.length < header.length;
  if (!needRewrite) {
    for (var i = 0; i < header.length; i++) {
      if (String(existing[i] || '') !== header[i]) {
        needRewrite = true;
        break;
      }
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
      var key = String(t.id);
      var val = answers[key] != null ? answers[key] : answers[t.id];
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
    var key = String(t.id);
    var val = answers[key] != null ? answers[key] : answers[t.id];
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

function labelChoice_(val) {
  if (!val) return '';
  return CHOICE_LABELS[val] || String(val);
}

function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
