// AutoBid Google Sheets bridge patch.
// Paste this into your existing Apps Script project that already has:
// - verifyExtensionSecret_(payload)
// - jsonResponse(data)
// - CONFIG.HEADER_ROW
// - CONFIG.NON_APPLICATION_SHEETS
//
// Then add the doPost branches shown near the bottom of this file.

const AUTOBID_BRIDGE_CONFIG = {
  QUESTION_HEADER: 'autobid_questions',
  ANSWER_HEADER: 'autobid_answers',
  STATUS_HEADER: 'autobid_status',
  UPDATED_AT_HEADER: 'autobid_updated_at',
  TAILORED_RESUME_COLUMN: 7,
  RESUME_FILE_COLUMN: 10,
  JOB_DESCRIPTION_COLUMN: 13,
  URL_HEADERS: [
    'job_url',
    'job url',
    'apply_url',
    'apply url',
    'application_url',
    'application url',
    'url',
    'link'
  ],
  RESUME_FILE_HEADERS: [
    'generated_resume',
    'generated resume',
    'generated_resume_link',
    'generated resume link',
    'resume_link',
    'resume link',
    'resume_url',
    'resume url',
    'resume_pdf',
    'resume pdf',
    'tailored_resume',
    'tailored resume',
    'tailored_resume_link',
    'tailored resume link',
    'cv_link',
    'cv link',
    'pdf_resume',
    'pdf resume'
  ]
};

function autobidListJobsFromExtension_(payload) {
  verifyExtensionSecret_(payload);

  const sheet = getAutobidSheet_(payload.sheetName);
  ensureAutobidBridgeColumns_(sheet);

  const startRow = Math.max(Number(CONFIG.HEADER_ROW || 1) + 1, Number(payload.startRow || payload.start_row || 2));
  const endRow = Math.max(startRow, Number(payload.endRow || payload.end_row || startRow));
  const headers = getAutobidHeaders_(sheet);
  const urlColumn = getAutobidUrlColumn_(headers);
  const lastColumn = Math.max(sheet.getLastColumn(), urlColumn);
  const rowCount = endRow - startRow + 1;
  const values = sheet.getRange(startRow, 1, rowCount, lastColumn).getDisplayValues();

  const jobs = values
    .map(function(row, index) {
      const rowNumber = startRow + index;
      const url = String(row[urlColumn - 1] || '').trim() || findFirstAutobidUrlInRow_(row);
      return {
        rowNumber: rowNumber,
        row: rowNumber,
        url: url,
        values: autobidRowToObject_(headers, row),
        raw: row
      };
    })
    .filter(function(job) {
      return /^https?:\/\/\S+/i.test(job.url);
    });

  return {
    status: 'success',
    action: 'autobidListJobs',
    jobs: jobs
  };
}

function autobidSaveQuestionsFromExtension_(payload) {
  verifyExtensionSecret_(payload);

  const sheet = getAutobidSheet_(payload.sheetName);
  const row = Number(payload.rowNumber || payload.row_number || payload.row);
  if (!row || row <= Number(CONFIG.HEADER_ROW || 1)) {
    throw new Error('Valid rowNumber is required.');
  }

  const columns = ensureAutobidBridgeColumns_(sheet);
  const now = new Date();
  sheet.getRange(row, columns.questions).setValue(JSON.stringify(payload.payload || {}));
  sheet.getRange(row, columns.status).setValue('questions_pending');
  sheet.getRange(row, columns.updatedAt).setValue(now);
  SpreadsheetApp.flush();

  return {
    status: 'success',
    action: 'autobidSaveQuestions',
    rowNumber: row,
    updated_at: now.toISOString()
  };
}

function autobidListPendingQuestionsFromExtension_(payload) {
  verifyExtensionSecret_(payload);

  const sheet = getAutobidSheet_(payload.sheetName);
  const columns = ensureAutobidBridgeColumns_(sheet);
  const startRow = Math.max(Number(CONFIG.HEADER_ROW || 1) + 1, Number(payload.startRow || payload.start_row || 2));
  const endRow = Math.max(startRow, Number(payload.endRow || payload.end_row || startRow));
  const headers = getAutobidHeaders_(sheet);
  const urlColumn = getAutobidUrlColumn_(headers);
  const lastColumn = Math.max(sheet.getLastColumn(), urlColumn, columns.questions, columns.answers);
  const rowCount = endRow - startRow + 1;
  const values = sheet.getRange(startRow, 1, rowCount, lastColumn).getDisplayValues();

  const rows = values
    .map(function(row, index) {
      const rowNumber = startRow + index;
      const questionsRaw = String(row[columns.questions - 1] || '').trim();
      const answersRaw = String(row[columns.answers - 1] || '').trim();
      if (!questionsRaw || answersRaw) {
        return null;
      }

      const questions = parseAutobidQuestions_(questionsRaw);
      if (!questions.fields || !questions.fields.length) {
        return null;
      }

      const url = String(row[urlColumn - 1] || '').trim() ||
        questions.row && questions.row.url ||
        questions.page && questions.page.url ||
        findFirstAutobidUrlInRow_(row);

      return {
        rowNumber: rowNumber,
        row: rowNumber,
        url: url,
        values: autobidRowToObject_(headers, row),
        raw: row,
        questions: questions
      };
    })
    .filter(function(row) {
      return row && row.rowNumber && row.questions && row.questions.fields && row.questions.fields.length;
    });

  return {
    status: 'success',
    action: 'autobidListPendingQuestions',
    rows: rows
  };
}

function autobidReadAnswersFromExtension_(payload) {
  verifyExtensionSecret_(payload);

  const sheet = getAutobidSheet_(payload.sheetName);
  const row = Number(payload.rowNumber || payload.row_number || payload.row);
  if (!row || row <= Number(CONFIG.HEADER_ROW || 1)) {
    throw new Error('Valid rowNumber is required.');
  }

  const columns = ensureAutobidBridgeColumns_(sheet);
  const raw = String(sheet.getRange(row, columns.answers).getDisplayValue() || '').trim();

  return {
    status: 'success',
    action: 'autobidReadAnswers',
    rowNumber: row,
    row_number: row,
    raw: raw,
    answers: parseAutobidAnswers_(raw)
  };
}

function autobidSaveAnswersFromExtension_(payload) {
  verifyExtensionSecret_(payload);

  const sheet = getAutobidSheet_(payload.sheetName);
  const row = Number(payload.rowNumber || payload.row_number || payload.row);
  if (!row || row <= Number(CONFIG.HEADER_ROW || 1)) {
    throw new Error('Valid rowNumber is required.');
  }

  const columns = ensureAutobidBridgeColumns_(sheet);
  const answersPayload = payload.payload || {
    answers: payload.answers || []
  };
  const normalized = {
    answers: parseAutobidAnswers_(JSON.stringify(answersPayload))
  };
  const now = new Date();

  sheet.getRange(row, columns.answers).setValue(JSON.stringify(normalized));
  sheet.getRange(row, columns.status).setValue('answers_ready');
  sheet.getRange(row, columns.updatedAt).setValue(now);
  SpreadsheetApp.flush();

  return {
    status: 'success',
    action: 'autobidSaveAnswers',
    rowNumber: row,
    savedAnswers: normalized.answers.length,
    updated_at: now.toISOString()
  };
}

function autobidReadResumeFileFromExtension_(payload) {
  verifyExtensionSecret_(payload);

  const sheet = getAutobidSheet_(payload.sheetName);
  const row = Number(payload.rowNumber || payload.row_number || payload.row);
  if (!row || row <= Number(CONFIG.HEADER_ROW || 1)) {
    throw new Error('Valid rowNumber is required.');
  }

  const headers = getAutobidHeaders_(sheet);
  const url = findFirstAutobidUrlInText_(payload.resumeUrl || payload.resume_url || '') ||
    getAutobidResumeFileUrl_(sheet, headers, row);
  if (!url) {
    return {
      status: 'success',
      action: 'autobidReadResumeFile',
      rowNumber: row,
      filename: '',
      mime_type: '',
      size: 0,
      source_url: '',
      base64: ''
    };
  }

  const fileId = extractAutobidDriveFileId_(url);
  if (!fileId) {
    throw new Error('Resume link is not a recognized Google Drive file URL: ' + url);
  }

  const file = DriveApp.getFileById(fileId);
  const blob = file.getBlob();
  const bytes = blob.getBytes();

  return {
    status: 'success',
    action: 'autobidReadResumeFile',
    rowNumber: row,
    row_number: row,
    filename: file.getName() || 'resume.pdf',
    mime_type: blob.getContentType() || 'application/pdf',
    size: bytes.length,
    source_url: url,
    base64: Utilities.base64Encode(bytes)
  };
}

function getAutobidSheet_(sheetName) {
  const name = String(sheetName || '').trim();
  if (!name) {
    throw new Error('sheetName is required.');
  }

  if (CONFIG.NON_APPLICATION_SHEETS && CONFIG.NON_APPLICATION_SHEETS.indexOf(name) !== -1) {
    throw new Error('This is not an application sheet: ' + name);
  }

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if (!sheet) {
    throw new Error('Sheet not found: ' + name);
  }

  return sheet;
}

function ensureAutobidBridgeColumns_(sheet) {
  return {
    questions: ensureAutobidHeaderColumn_(sheet, AUTOBID_BRIDGE_CONFIG.QUESTION_HEADER),
    answers: ensureAutobidHeaderColumn_(sheet, AUTOBID_BRIDGE_CONFIG.ANSWER_HEADER),
    status: ensureAutobidHeaderColumn_(sheet, AUTOBID_BRIDGE_CONFIG.STATUS_HEADER),
    updatedAt: ensureAutobidHeaderColumn_(sheet, AUTOBID_BRIDGE_CONFIG.UPDATED_AT_HEADER)
  };
}

function ensureAutobidHeaderColumn_(sheet, headerName) {
  const headers = getAutobidHeaders_(sheet);
  const key = normalizeAutobidHeader_(headerName);

  if (headers[key]) {
    return headers[key];
  }

  const nextColumn = Math.max(sheet.getLastColumn(), 1) + 1;
  sheet.getRange(Number(CONFIG.HEADER_ROW || 1), nextColumn).setValue(headerName);
  SpreadsheetApp.flush();
  return nextColumn;
}

function getAutobidHeaders_(sheet) {
  const headerRow = Number(CONFIG.HEADER_ROW || 1);
  const lastColumn = Math.max(sheet.getLastColumn(), 1);
  const row = sheet.getRange(headerRow, 1, 1, lastColumn).getDisplayValues()[0];
  const headers = {};

  row.forEach(function(value, index) {
    const key = normalizeAutobidHeader_(value);
    if (key && !headers[key]) {
      headers[key] = index + 1;
    }
  });

  return headers;
}

function getAutobidUrlColumn_(headers) {
  for (let i = 0; i < AUTOBID_BRIDGE_CONFIG.URL_HEADERS.length; i++) {
    const key = normalizeAutobidHeader_(AUTOBID_BRIDGE_CONFIG.URL_HEADERS[i]);
    if (headers[key]) {
      return headers[key];
    }
  }

  if (CONFIG.COLUMNS && CONFIG.COLUMNS.applyUrl) {
    return CONFIG.COLUMNS.applyUrl;
  }

  throw new Error('No job URL column found. Add a header named job_url, apply_url, url, or link.');
}

function autobidRowToObject_(headers, row) {
  const result = {};

  Object.keys(headers).forEach(function(key) {
    result[key] = row[headers[key] - 1] || '';
  });

  if (!String(result.column_g || '').trim()) {
    result.column_g = row[AUTOBID_BRIDGE_CONFIG.TAILORED_RESUME_COLUMN - 1] || '';
  }
  if (!String(result.column_j || '').trim()) {
    result.column_j = row[AUTOBID_BRIDGE_CONFIG.RESUME_FILE_COLUMN - 1] || '';
  }
  if (!String(result.column_m || '').trim()) {
    result.column_m = row[AUTOBID_BRIDGE_CONFIG.JOB_DESCRIPTION_COLUMN - 1] || '';
  }

  return result;
}

function findFirstAutobidUrlInRow_(row) {
  for (let i = 0; i < row.length; i++) {
    const value = String(row[i] || '').trim();
    if (/^https?:\/\/\S+/i.test(value)) {
      return value;
    }
  }

  return '';
}

function getAutobidResumeFileUrl_(sheet, headers, row) {
  if (AUTOBID_BRIDGE_CONFIG.RESUME_FILE_COLUMN) {
    const fixedValue = getAutobidCellUrlOrText_(sheet, row, AUTOBID_BRIDGE_CONFIG.RESUME_FILE_COLUMN);
    const fixedUrl = findFirstAutobidUrlInText_(fixedValue);
    if (fixedUrl) {
      return fixedUrl;
    }
  }

  for (let i = 0; i < AUTOBID_BRIDGE_CONFIG.RESUME_FILE_HEADERS.length; i++) {
    const key = normalizeAutobidHeader_(AUTOBID_BRIDGE_CONFIG.RESUME_FILE_HEADERS[i]);
    if (!headers[key]) {
      continue;
    }

    const value = getAutobidCellUrlOrText_(sheet, row, headers[key]);
    const url = findFirstAutobidUrlInText_(value);
    if (url) {
      return url;
    }
  }

  const lastColumn = Math.max(sheet.getLastColumn(), 1);
  for (let column = 1; column <= lastColumn; column++) {
    const value = getAutobidCellUrlOrText_(sheet, row, column);
    const url = findFirstAutobidUrlInText_(value);
    if (url && /drive\.google\.com|docs\.google\.com/i.test(url)) {
      return url;
    }
  }

  return '';
}

function getAutobidCellUrlOrText_(sheet, row, column) {
  const range = sheet.getRange(row, column);
  const values = [];

  const richText = range.getRichTextValue();
  if (richText) {
    values.push(richText.getLinkUrl() || '');
    const runs = richText.getRuns ? richText.getRuns() : [];
    runs.forEach(function(run) {
      values.push(run.getLinkUrl() || '');
    });
  }

  values.push(range.getFormula() || '');
  values.push(range.getDisplayValue() || '');
  return values.filter(Boolean).join(' ');
}

function findFirstAutobidUrlInText_(text) {
  const match = String(text || '').match(/https?:\/\/[^\s"'<>),]+/i);
  return match ? match[0] : '';
}

function extractAutobidDriveFileId_(url) {
  const text = String(url || '');
  const patterns = [
    /\/file\/d\/([a-zA-Z0-9_-]+)/,
    /\/document\/d\/([a-zA-Z0-9_-]+)/,
    /\/presentation\/d\/([a-zA-Z0-9_-]+)/,
    /[?&]id=([a-zA-Z0-9_-]+)/,
    /\/uc\?[^#]*id=([a-zA-Z0-9_-]+)/
  ];

  for (let i = 0; i < patterns.length; i++) {
    const match = text.match(patterns[i]);
    if (match && match[1]) {
      return match[1];
    }
  }

  return '';
}

function parseAutobidAnswers_(raw) {
  const text = String(raw || '').trim();
  if (!text) {
    return [];
  }

  try {
    const parsed = JSON.parse(text);
    if (Array.isArray(parsed)) {
      return normalizeAutobidAnswerArray_(parsed);
    }
    if (parsed && Array.isArray(parsed.answers)) {
      return normalizeAutobidAnswerArray_(parsed.answers);
    }
    if (parsed && typeof parsed === 'object') {
      return Object.keys(parsed)
        .map(function(fieldId) {
          return {
            field_id: fieldId,
            value: String(parsed[fieldId] || ''),
            source: 'sheet'
          };
        })
        .filter(function(answer) {
          return answer.field_id && answer.value;
        });
    }
  } catch (err) {
    return [];
  }

  return [];
}

function parseAutobidQuestions_(raw) {
  const text = String(raw || '').trim();
  if (!text) {
    return { fields: [] };
  }

  try {
    const parsed = JSON.parse(text);
    const fields = Array.isArray(parsed.fields)
      ? parsed.fields
        .map(function(field) {
          return {
            field_id: String(field.field_id || field.id || ''),
            id: String(field.id || field.field_id || ''),
            question: String(field.question || field.label || ''),
            option: String(field.option || ''),
            label: String(field.label || field.question || ''),
            raw_label: String(field.raw_label || ''),
            name: String(field.name || ''),
            placeholder: String(field.placeholder || ''),
            type: String(field.type || ''),
            required: Boolean(field.required),
            options: Array.isArray(field.options)
              ? field.options.map(function(option) { return String(option || ''); }).filter(Boolean)
              : []
          };
        })
        .filter(function(field) {
          return field.field_id || field.id;
        })
      : [];
    return Object.assign({}, parsed, { fields: fields });
  } catch (err) {
    return { fields: [] };
  }
}

function normalizeAutobidAnswerArray_(answers) {
  return answers
    .map(function(answer) {
      var value = Object.prototype.hasOwnProperty.call(answer, 'value') ? answer.value : answer.answer;
      return {
        field_id: String(answer.field_id || answer.id || ''),
        question: String(answer.question || ''),
        option: String(answer.option || ''),
        value: String(value == null ? '' : value),
        source: 'sheet'
      };
    })
    .filter(function(answer) {
      return answer.field_id;
    });
}

function normalizeAutobidHeader_(value) {
  return String(value || '')
    .trim()
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, '_')
    .replace(/^_+|_+$/g, '');
}

// Add these branches inside doPost(e), after:
// const action = String(payload.action || '').trim();
//
// if (action === 'autobidListJobs') {
//   return jsonResponse(autobidListJobsFromExtension_(payload));
// }
//
// if (action === 'autobidSaveQuestions') {
//   return jsonResponse(autobidSaveQuestionsFromExtension_(payload));
// }
//
// if (action === 'autobidListPendingQuestions') {
//   return jsonResponse(autobidListPendingQuestionsFromExtension_(payload));
// }
//
// if (action === 'autobidReadAnswers') {
//   return jsonResponse(autobidReadAnswersFromExtension_(payload));
// }
//
// if (action === 'autobidSaveAnswers') {
//   return jsonResponse(autobidSaveAnswersFromExtension_(payload));
// }
//
// if (action === 'autobidReadResumeFile') {
//   return jsonResponse(autobidReadResumeFileFromExtension_(payload));
// }
