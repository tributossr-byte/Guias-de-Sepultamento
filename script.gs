// ============================================================
// GUIAS DE SEPULTAMENTO — Serra Redonda/PB
// Google Apps Script — com autenticação e gerenciamento de usuários
// ============================================================

const SHEET_GUIAS    = 'Guias';
const SHEET_CONFIG   = 'Config';
const SHEET_LOG      = 'Log';
const SHEET_USUARIOS = 'Usuarios';

// Usuário administrador fixo (não pode ser removido)
const ADMIN_USUARIO = 'julianna';

function doPost(e) { return handleRequest(e); }
function doGet(e)  { return handleRequest(e); }

function handleRequest(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let params;
  try { params = JSON.parse(e.postData.contents); } catch { params = e.parameter; }
  const action = params.action;

  try {
    let result;
    if      (action === 'login')         result = login(params);
    else if (action === 'getNextNum')    result = getNextNum(ss, params);
    else if (action === 'saveGuia')      result = saveGuia(ss, params);
    else if (action === 'getHistory')    result = getHistory(ss, params);
    else if (action === 'getUsuarios')   result = getUsuarios(ss, params);
    else if (action === 'saveUsuario')   result = saveUsuario(ss, params);
    else if (action === 'removeUsuario') result = removeUsuario(ss, params);
    else result = { error: 'Acao desconhecida: ' + action };

    return ContentService
      .createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  } catch(err) {
    return ContentService
      .createTextOutput(JSON.stringify({ error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ============================================================
// USUÁRIOS — lidos da aba "Usuarios" da planilha
// Colunas: usuario | senha | nome | admin
// ============================================================

function getSheetUsuarios(ss) {
  let sheet = ss.getSheetByName(SHEET_USUARIOS);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_USUARIOS);
    sheet.appendRow(['usuario', 'senha', 'nome', 'admin']);
    sheet.getRange(1, 1, 1, 4).setFontWeight('bold').setBackground('#1a2332').setFontColor('#ffffff');
    // Criar usuária administradora padrão
    sheet.appendRow([ADMIN_USUARIO, 'serra2024', 'Julianna Ferreira dos Santos Silva', 'sim']);
  }
  return sheet;
}

function listarUsuarios(ss) {
  const sheet = getSheetUsuarios(ss);
  if (sheet.getLastRow() < 2) return [];
  return sheet.getRange(2, 1, sheet.getLastRow() - 1, 4).getValues()
    .filter(r => r[0])
    .map(r => ({ usuario: r[0], senha: r[1], nome: r[2], admin: r[3] === 'sim' }));
}

function login(params) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const u = (params.usuario || '').toLowerCase().trim();
  const s = (params.senha || '').trim();
  const usuarios = listarUsuarios(ss);
  const found = usuarios.find(x => x.usuario.toLowerCase() === u && x.senha === s);
  if (!found) return { ok: false, error: 'Usuário ou senha incorretos.' };
  const token = Utilities.base64Encode(found.usuario + ':' + Date.now() + ':sep2025');
  return { ok: true, token, nome: found.nome, usuario: found.usuario, admin: found.admin };
}

function validarToken(token) {
  if (!token) return null;
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const decoded = Utilities.newBlob(Utilities.base64Decode(token)).getDataAsString();
    const parts = decoded.split(':');
    if (parts.length < 3) return null;
    const usuario = parts[0];
    const ts = parseInt(parts[1]);
    if (Date.now() - ts > 12 * 60 * 60 * 1000) return null;
    const usuarios = listarUsuarios(ss);
    return usuarios.find(x => x.usuario.toLowerCase() === usuario.toLowerCase()) || null;
  } catch { return null; }
}

function getUsuarios(ss, params) {
  const user = validarToken(params.token);
  if (!user) return { error: 'Sessão expirada.' };
  if (!user.admin) return { error: 'Acesso negado.' };
  const lista = listarUsuarios(ss).map(u => ({ usuario: u.usuario, nome: u.nome, admin: u.admin }));
  return { ok: true, usuarios: lista };
}

function saveUsuario(ss, params) {
  const user = validarToken(params.token);
  if (!user) return { error: 'Sessão expirada.' };
  if (!user.admin) return { error: 'Acesso negado.' };

  const sheet = getSheetUsuarios(ss);
  const novoUsuario = (params.usuario || '').toLowerCase().trim();
  const novaSenha   = (params.senha || '').trim();
  const novoNome    = (params.nome || '').trim();

  if (!novoUsuario || !novaSenha || !novoNome) return { error: 'Preencha todos os campos.' };

  // Verificar se já existe (para edição permitida)
  const rows = sheet.getLastRow() > 1
    ? sheet.getRange(2, 1, sheet.getLastRow() - 1, 4).getValues()
    : [];

  for (let i = 0; i < rows.length; i++) {
    if (rows[i][0].toString().toLowerCase() === novoUsuario) {
      // Atualizar linha existente
      sheet.getRange(i + 2, 1, 1, 4).setValues([[novoUsuario, novaSenha, novoNome, rows[i][3]]]);
      return { ok: true, acao: 'atualizado' };
    }
  }

  // Novo usuário
  sheet.appendRow([novoUsuario, novaSenha, novoNome, '']);
  return { ok: true, acao: 'criado' };
}

function removeUsuario(ss, params) {
  const user = validarToken(params.token);
  if (!user) return { error: 'Sessão expirada.' };
  if (!user.admin) return { error: 'Acesso negado.' };

  const alvo = (params.usuario || '').toLowerCase().trim();
  if (alvo === ADMIN_USUARIO) return { error: 'O usuário administrador principal não pode ser removido.' };

  const sheet = getSheetUsuarios(ss);
  if (sheet.getLastRow() < 2) return { error: 'Usuário não encontrado.' };

  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
  for (let i = rows.length - 1; i >= 0; i--) {
    if (rows[i][0].toString().toLowerCase() === alvo) {
      sheet.deleteRow(i + 2);
      return { ok: true };
    }
  }
  return { error: 'Usuário não encontrado.' };
}

// ============================================================
// GUIAS
// ============================================================

// Número inicial por ano — 2026 começa em 25, demais anos começam em 1
const INICIO_POR_ANO = { 2026: 25 };

function getNextNum(ss, params) {
  const user = validarToken(params.token);
  if (!user) return { error: 'Sessão expirada. Faça login novamente.' };

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    let sheet = ss.getSheetByName(SHEET_CONFIG);
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_CONFIG);
      sheet.getRange('A1:B1').setValues([['ano', 'contador']]);
      sheet.getRange(1,1,1,2).setFontWeight('bold');
    }

    const ano = new Date().getFullYear();

    // Procurar linha do ano atual
    const lastRow = sheet.getLastRow();
    let anoRow = -1;
    if (lastRow >= 2) {
      const anos = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
      for (let i = 0; i < anos.length; i++) {
        if (parseInt(anos[i][0]) === ano) { anoRow = i + 2; break; }
      }
    }

    let cur, cell;
    if (anoRow === -1) {
      // Primeiro registro do ano — iniciar com valor definido
      const inicio = (INICIO_POR_ANO[ano] || 1) - 1; // -1 pois vamos somar 1 abaixo
      sheet.appendRow([ano, inicio]);
      anoRow = sheet.getLastRow();
    }

    cell = sheet.getRange(anoRow, 2);
    cur = parseInt(cell.getValue()) || 0;
    const next = cur + 1;
    cell.setValue(next);

    return { registro: ano + '.' + String(next).padStart(4,'0'), num: next, ano };
  } finally { lock.releaseLock(); }
}

function saveGuia(ss, params) {
  const user = validarToken(params.token);
  if (!user) return { error: 'Sessão expirada. Faça login novamente.' };

  let sheet = ss.getSheetByName(SHEET_GUIAS);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_GUIAS);
    sheet.appendRow(['Registro','Nome','CPF','Nascimento','Idade','Pai','Mãe','CEP','Endereço','Cidade','Estado','Causa','Dec. Óbito','Data Óbito','Local Óbito','Data Sep.','Cemitério','End. Cemitério','Emitida em','Emissor','Cargo','Chave','JSON Completo']);
    sheet.getRange(1,1,1,23).setFontWeight('bold').setBackground('#1a2332').setFontColor('#ffffff');
  }

  const d = params.data || params;
  sheet.appendRow([
    d.registro, d.nome, d.cpf, d.nascimento, d.idade,
    d.nomePai, d.nomeMae, d.cep, d.endereco, d.cidade, d.estado,
    d.causa, d.decObito, d.dtObito, d.ocorrencia,
    d.dtSep, d.cemiterio, d.endCemiterio, d.emitidoEm,
    user.nome, '', d.chave,
    JSON.stringify(d)
  ]);

  registrarLog(ss, user, 'EMISSÃO', d.registro);
  return { ok: true, registro: d.registro };
}

function getHistory(ss, params) {
  const user = validarToken(params.token);
  if (!user) return { error: 'Sessão expirada. Faça login novamente.' };

  const sheet = ss.getSheetByName(SHEET_GUIAS);
  if (!sheet || sheet.getLastRow() < 2) return { guias: [] };

  const rows = sheet.getRange(2, 1, sheet.getLastRow()-1, 23).getValues();
  const guias = rows.filter(r => r[0]).map(r => {
    let parsed = {};
    try { parsed = JSON.parse(r[22] || '{}'); } catch {}
    return {
      registro: String(r[0]), nome: r[1], dtObito: r[13],
      cemiterio: r[16], emitidoEm: r[18],
      emissor: r[19], cargo: r[20],
      data: parsed
    };
  }).reverse();
  return { guias };
}

function registrarLog(ss, user, acao, referencia) {
  let sheet = ss.getSheetByName(SHEET_LOG);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_LOG);
    sheet.appendRow(['Data/Hora','Usuário','Nome','Ação','Referência']);
    sheet.getRange(1,1,1,5).setFontWeight('bold');
  }
  sheet.appendRow([new Date().toLocaleString('pt-BR'), user.usuario, user.nome, acao, referencia]);
}
