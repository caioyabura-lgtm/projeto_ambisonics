const SHEET_NAME = 'eventos_site';
const APP_SECRET = 'TROCAR_POR_UMA_CHAVE_SIMPLES_E_GRANDE';

const HEADERS = [
  'timestamp',
  'eventId',
  'sessionId',
  'tipo',
  'pagina',
  'secao',
  'acao',
  'rotulo',
  'url',
  'referrer',
  'userAgent',
  'idioma',
  'larguraTela',
  'alturaTela',
  'tempoNaPaginaSeg',
  'extraJson'
];

function doPost(e) {
  try {
    const body = e && e.postData && e.postData.contents
      ? JSON.parse(e.postData.contents)
      : {};

    if (body.secret !== APP_SECRET) {
      return jsonResponse({
        ok: false,
        error: 'unauthorized'
      });
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = getOrCreateSheet_(ss);

    const payload = body.payload || {};
    const now = new Date();

    const row = [
      now,
      payload.eventId || '',
      payload.sessionId || '',
      payload.tipo || '',
      payload.pagina || '',
      payload.secao || '',
      payload.acao || '',
      payload.rotulo || '',
      payload.url || '',
      payload.referrer || '',
      payload.userAgent || '',
      payload.idioma || '',
      payload.larguraTela || '',
      payload.alturaTela || '',
      payload.tempoNaPaginaSeg || '',
      payload.extraJson ? JSON.stringify(payload.extraJson) : ''
    ];

    sheet.appendRow(row);

    return jsonResponse({
      ok: true,
      savedAt: now.toISOString()
    });

  } catch (err) {
    return jsonResponse({
      ok: false,
      error: String(err && err.message ? err.message : err)
    });
  }
}

function doGet() {
  return jsonResponse({
    ok: true,
    service: 'O Som no Espaco - eventos_site',
    message: 'Use POST para registrar eventos.'
  });
}

function getOrCreateSheet_(ss) {
  let sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }

  const firstRow = sheet.getRange(1, 1, 1, HEADERS.length).getValues()[0];
  const hasHeaders = firstRow.some(value => String(value || '').trim());

  if (!hasHeaders) {
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
    sheet.setFrozenRows(1);
  }

  return sheet;
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
