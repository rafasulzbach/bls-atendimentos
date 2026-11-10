/**
 * BLS Refrigeração · Atendimentos — servidor gratuito (Google Apps Script)
 *
 * Guarda os dados numa planilha Google e as fotos/vídeos/áudios numa pasta do Drive.
 *
 * COMO INSTALAR / ATUALIZAR:
 *   1) script.google.com → projeto "BLS Atendimentos" → apague tudo, cole este código → Salvar
 *   2) Cole o ID do cliente Google em GOOGLE_CLIENT_ID (logo abaixo)
 *   3) Escolha a função "configurar" e clique em Executar (autorize se pedir)
 *   4) Implantar → Gerenciar implantações → ✏ Editar → Versão: "Nova versão" → Implantar
 *      (a URL continua a mesma)
 */

const PASTA_ID = '1z5Vqz5UHFx0O3G6Fgc5k4tQ6mu5qrWyr';   // pasta do Drive da BLS
const GOOGLE_CLIENT_ID = 'COLE_AQUI_O_ID_DO_CLIENTE.apps.googleusercontent.com';
const VERSAO_API = '1.2';

const CAB = ['ID', 'Nº', 'Status', 'OS', 'Placa', 'Técnico', 'Criado em', 'Finalizado em',
             'Em espera', 'Enviado', 'Concluído', 'Atualizado em', 'Pasta', 'Excluído',
             'Criado por', 'Alterado por', 'Servidor', 'JSON'];
const CAB_USU = ['E-mail', 'Nome', 'Situação', 'Criado em', 'Último acesso'];
const STATUS = { pendente: 'Pendente', em_espera: 'Em espera', enviado: 'Enviado', concluido: 'Concluído' };
const DIAS = 24 * 60 * 60 * 1000;

function props() { return PropertiesService.getScriptProperties(); }

/* ---------- Instalação ---------- */
function configurar() {
  const p = props();
  const pasta = DriveApp.getFolderById(PASTA_ID);
  let ss;
  const pid = p.getProperty('PLANILHA_ID');
  if (pid) {
    ss = SpreadsheetApp.openById(pid);
  } else {
    ss = SpreadsheetApp.create('BLS Atendimentos - Dados');
    DriveApp.getFileById(ss.getId()).moveTo(pasta);
    ss.getSheets()[0].setName('Atendimentos');
    p.setProperty('PLANILHA_ID', ss.getId());
  }
  // aba de atendimentos (com migração de colunas, se necessário)
  let sh = ss.getSheetByName('Atendimentos') || ss.insertSheet('Atendimentos');
  const cabAtual = sh.getLastColumn() ? sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0] : [];
  if (cabAtual.join('|') !== CAB.join('|')) {
    const n = sh.getLastRow() - 1;
    const jsons = [];
    if (n > 0 && cabAtual.length) {
      const col = cabAtual.indexOf('JSON') + 1;
      if (col > 0) sh.getRange(2, col, n, 1).getValues().forEach(function (v) { try { jsons.push(JSON.parse(v[0])); } catch (e) {} });
      sh.getRange(2, 1, n, sh.getLastColumn()).clearContent();
    }
    sh.getRange(1, 1, 1, Math.max(CAB.length, cabAtual.length)).clearContent();
    sh.getRange(1, 1, 1, CAB.length).setValues([CAB]);
    jsons.forEach(function (a, i) { sh.getRange(i + 2, 1, 1, CAB.length).setValues([linhaDe(a)]); });
  }
  sh.getRange(1, 1, 1, CAB.length).setFontWeight('bold').setBackground('#142269').setFontColor('#ffffff');
  sh.setFrozenRows(1);
  // aba de usuários
  let su = ss.getSheetByName('Usuarios');
  if (!su) {
    su = ss.insertSheet('Usuarios');
    su.getRange(1, 1, 1, CAB_USU.length).setValues([CAB_USU])
      .setFontWeight('bold').setBackground('#142269').setFontColor('#ffffff');
    su.setFrozenRows(1);
  }
  pastaAtendimentos();
  if (!p.getProperty('CHAVE')) p.setProperty('CHAVE', Utilities.getUuid().replace(/-/g, '').slice(0, 16).toUpperCase());
  if (!p.getProperty('SEGREDO')) p.setProperty('SEGREDO', Utilities.getUuid() + Utilities.getUuid());
  p.setProperty('ADMIN_EMAIL', Session.getEffectiveUser().getEmail().toLowerCase());
  // pede a autorização usada para enviar vídeos grandes ao Drive
  UrlFetchApp.fetch('https://www.googleapis.com/drive/v3/about?fields=user',
    { headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }, muteHttpExceptions: true });

  Logger.log('✅ Tudo pronto!');
  Logger.log('Planilha: ' + ss.getUrl());
  Logger.log('CHAVE DE ACESSO: ' + p.getProperty('CHAVE'));
  Logger.log('Administrador (login Google): ' + p.getProperty('ADMIN_EMAIL'));
  if (GOOGLE_CLIENT_ID.indexOf('COLE_AQUI') === 0) Logger.log('⚠ Falta colar o GOOGLE_CLIENT_ID no topo do código.');
}

function pastaAtendimentos() {
  const p = props();
  const id = p.getProperty('PASTA_ATEND');
  if (id) { try { return DriveApp.getFolderById(id); } catch (e) {} }
  const f = DriveApp.getFolderById(PASTA_ID).createFolder('Atendimentos');
  p.setProperty('PASTA_ATEND', f.getId());
  return f;
}

function ss() {
  const id = props().getProperty('PLANILHA_ID');
  if (!id) throw new Error('Servidor não configurado: execute a função "configurar".');
  return SpreadsheetApp.openById(id);
}
function planilha() { return ss().getSheetByName('Atendimentos'); }
function abaUsuarios() { return ss().getSheetByName('Usuarios'); }

/* ---------- Entrada da API ---------- */
function doGet() {
  return ContentService.createTextOutput('BLS Atendimentos API ' + VERSAO_API + ' — funcionando.');
}

const PUBLICAS = { ping: 1, loginGoogle: 1 };
const SO_ADMIN = { usuarios: 1, usuarioAcao: 1 };

function doPost(e) {
  let r;
  try {
    const q = JSON.parse(e.postData.contents);
    if (!q.chave || q.chave !== props().getProperty('CHAVE')) throw new Error('Chave de acesso inválida');
    const f = ACOES[q.acao];
    if (!f) throw new Error('Ação desconhecida: ' + q.acao);
    let sess = null;
    if (!PUBLICAS[q.acao]) {
      sess = validarToken(q.token);
      if (SO_ADMIN[q.acao] && !sess.adm) throw new Error('Apenas o administrador pode fazer isso');
    }
    r = f(q, sess) || {};
    r.ok = true;
  } catch (err) {
    const msg = String((err && err.message) || err);
    r = { ok: false, erro: msg.replace(/^SESSAO:\s*/, ''), sessao: msg.indexOf('SESSAO:') === 0 };
  }
  return ContentService.createTextOutput(JSON.stringify(r)).setMimeType(ContentService.MimeType.JSON);
}

/* ---------- Login ---------- */
function b64u(bytes) { return Utilities.base64EncodeWebSafe(bytes).replace(/=+$/, ''); }
function assinar(payload) {
  const corpo = b64u(Utilities.newBlob(JSON.stringify(payload)).getBytes());
  const sig = b64u(Utilities.computeHmacSha256Signature(corpo, props().getProperty('SEGREDO')));
  return corpo + '.' + sig;
}
function validarToken(tok) {
  if (!tok) throw new Error('SESSAO: Faça login novamente');
  const partes = String(tok).split('.');
  const sig = b64u(Utilities.computeHmacSha256Signature(partes[0], props().getProperty('SEGREDO')));
  if (sig !== partes[1]) throw new Error('SESSAO: Sessão inválida, faça login novamente');
  const p = JSON.parse(Utilities.newBlob(Utilities.base64DecodeWebSafe(partes[0])).getDataAsString());
  if (p.exp < Date.now()) throw new Error('SESSAO: Sessão expirada, faça login novamente');
  if (!p.adm) {
    const u = acharUsuario(p.u);
    if (!u || u.situacao !== 'Liberado') throw new Error('SESSAO: Acesso não liberado');
  }
  return p;
}
function sessaoDe(usuario, nome, adm) {
  const exp = Date.now() + (adm ? 30 : 365) * DIAS;
  return { token: assinar({ u: usuario, n: nome, adm: !!adm, exp: exp }), usuario: usuario, nome: nome, admin: !!adm, exp: exp };
}
function chaveUsuario(u) { return String(u || '').trim().toLowerCase(); }
function acharUsuario(u) {
  const su = abaUsuarios(), n = su.getLastRow() - 1, k = chaveUsuario(u);
  if (n < 1) return null;
  const vals = su.getRange(2, 1, n, CAB_USU.length).getValues();
  for (let i = 0; i < vals.length; i++) {
    if (String(vals[i][0]).toLowerCase() === k) return { linha: i + 2, email: k, nome: vals[i][1], situacao: vals[i][2] };
  }
  return null;
}

/* ---------- Ações ---------- */
const ACOES = {
  ping() { return { versao: VERSAO_API, agora: Date.now(), googleClientId: GOOGLE_CLIENT_ID.indexOf('COLE_AQUI') === 0 ? '' : GOOGLE_CLIENT_ID }; },

  // Login com Google (todos). Técnico novo fica "Pendente" até o administrador liberar.
  loginGoogle(q) {
    if (GOOGLE_CLIENT_ID.indexOf('COLE_AQUI') === 0) throw new Error('Login Google ainda não configurado no servidor');
    const r = UrlFetchApp.fetch('https://oauth2.googleapis.com/tokeninfo?id_token=' + encodeURIComponent(q.credencial), { muteHttpExceptions: true });
    if (r.getResponseCode() !== 200) throw new Error('Login Google inválido');
    const t = JSON.parse(r.getContentText());
    const email = String(t.email || '').toLowerCase();
    if (t.aud !== GOOGLE_CLIENT_ID || String(t.email_verified) !== 'true') throw new Error('Login Google inválido');
    const nome = String(t.name || t.given_name || email.split('@')[0]).slice(0, 60);
    if (email === props().getProperty('ADMIN_EMAIL')) return { sessao: sessaoDe(email, nome, true) };

    const lock = LockService.getScriptLock(); lock.waitLock(30000);
    try {
      const su = abaUsuarios();
      const u = acharUsuario(email);
      if (!u) {
        su.appendRow([email, nome, 'Pendente', new Date(), new Date()]);
        return { pendente: true, nome: nome, email: email };
      }
      su.getRange(u.linha, 5).setValue(new Date());
      if (u.situacao === 'Pendente') return { pendente: true, nome: nome, email: email };
      if (u.situacao !== 'Liberado') throw new Error('Seu acesso está bloqueado. Fale com o administrador.');
      if (u.nome !== nome) su.getRange(u.linha, 2).setValue(nome);
      return { sessao: sessaoDe(email, nome, false) };
    } finally { lock.releaseLock(); }
  },

  usuarios() {
    const su = abaUsuarios(), n = su.getLastRow() - 1;
    if (n < 1) return { usuarios: [] };
    return { usuarios: su.getRange(2, 1, n, 5).getValues().map(function (v) {
      return { email: v[0], nome: v[1], situacao: v[2], criadoEm: v[3] ? new Date(v[3]).getTime() : null, ultimoAcesso: v[4] ? new Date(v[4]).getTime() : null };
    }) };
  },

  usuarioAcao(q) {
    const u = acharUsuario(q.email); if (!u) throw new Error('Usuário não encontrado');
    const su = abaUsuarios();
    if (q.op === 'liberar') su.getRange(u.linha, 3).setValue('Liberado');
    else if (q.op === 'bloquear') su.getRange(u.linha, 3).setValue('Bloqueado');
    else if (q.op === 'excluir') su.deleteRow(u.linha);
    else throw new Error('Operação inválida');
    return {};
  },

  // Grava/atualiza um atendimento (vale a versão alterada por último)
  salvar(q, sess) {
    const lock = LockService.getScriptLock();
    lock.waitLock(30000);
    try {
      const sh = planilha();
      const a = q.atend;
      if (!a || !a.id) throw new Error('Atendimento inválido');
      const lin = acharLinha(sh, a.id);
      let atual = null;
      if (lin) { try { atual = JSON.parse(sh.getRange(lin, CAB.length).getValue()); } catch (e) {} }
      if (atual && (atual.atualizadoEm || 0) > (a.atualizadoEm || 0)) return { atend: atual };

      a.numero = atual && atual.numero ? atual.numero : proximoNumero(sh);
      a.criadoPor = (atual && atual.criadoPor) || a.criadoPor || sess.n;
      a.alteradoPor = sess.n;
      const pasta = pastaDo(a);
      a.pastaId = pasta.getId();
      a.pasta = pasta.getUrl();
      const nome = nomePasta(a);
      if (pasta.getName() !== nome) pasta.setName(nome);
      if (a.excluido) { try { pasta.setTrashed(true); } catch (e) {} }
      delete a.sujo;
      a.servidorEm = Date.now();

      const linha = linhaDe(a);
      if (lin) sh.getRange(lin, 1, 1, linha.length).setValues([linha]);
      else sh.appendRow(linha);
      return { atend: a };
    } finally {
      lock.releaseLock();
    }
  },

  // Lista os atendimentos alterados desde a última consulta
  listar(q) {
    const lock = LockService.getScriptLock();
    lock.waitLock(30000);
    try {
      const sh = planilha();
      const agora = Date.now();
      const n = sh.getLastRow() - 1;
      if (n < 1) return { agora: agora, itens: [] };
      const desde = Number(q.desde) || 0;
      const vals = sh.getRange(2, CAB.length - 1, n, 2).getValues();
      const itens = [];
      vals.forEach(function (v) {
        if (Number(v[0]) > desde) { try { itens.push(JSON.parse(v[1])); } catch (e) {} }
      });
      return { agora: agora, itens: itens };
    } finally {
      lock.releaseLock();
    }
  },

  // Recebe foto/vídeo/áudio (em partes de até 4 MB) e grava na pasta do atendimento
  upload(q) {
    const bytes = Utilities.base64Decode(q.data);
    const pasta = pastaDo({ id: q.atendId });
    if (!q.sessao && q.offset === 0 && bytes.length >= q.tamanho) {
      const f = pasta.createFile(Utilities.newBlob(bytes, q.mime, q.nome));
      return { fim: true, driveId: f.getId() };
    }
    let sessao = q.sessao;
    if (!sessao) {
      const r0 = UrlFetchApp.fetch('https://www.googleapis.com/upload/drive/v3/files?uploadType=resumable&fields=id', {
        method: 'post',
        contentType: 'application/json; charset=UTF-8',
        headers: {
          Authorization: 'Bearer ' + ScriptApp.getOAuthToken(),
          'X-Upload-Content-Type': q.mime,
          'X-Upload-Content-Length': String(q.tamanho),
        },
        payload: JSON.stringify({ name: q.nome, parents: [pasta.getId()] }),
        muteHttpExceptions: true,
      });
      if (r0.getResponseCode() !== 200) throw new Error('Falha ao iniciar envio: ' + r0.getContentText().slice(0, 200));
      const h = r0.getAllHeaders();
      sessao = h['Location'] || h['location'];
    }
    const fim = q.offset + bytes.length - 1;
    const r = UrlFetchApp.fetch(sessao, {
      method: 'put',
      contentType: q.mime,
      headers: { 'Content-Range': 'bytes ' + q.offset + '-' + fim + '/' + q.tamanho },
      payload: bytes,
      muteHttpExceptions: true,
    });
    const code = r.getResponseCode();
    if (code === 308) return { fim: false, sessao: sessao };
    if (code === 200 || code === 201) return { fim: true, driveId: JSON.parse(r.getContentText()).id };
    throw new Error('Falha no envio (' + code + '): ' + r.getContentText().slice(0, 200));
  },

  // Devolve um arquivo (foto/áudio) para mostrar em outro aparelho
  baixar(q) {
    const b = DriveApp.getFileById(q.driveId).getBlob();
    return { mime: b.getContentType(), data: Utilities.base64Encode(b.getBytes()) };
  },

  // Miniatura (usada para vídeos)
  miniatura(q) {
    try {
      const t = DriveApp.getFileById(q.driveId).getThumbnail();
      return t ? { mime: t.getContentType(), data: Utilities.base64Encode(t.getBytes()) } : { data: null };
    } catch (e) {
      return { data: null };
    }
  },
};

/* ---------- Auxiliares ---------- */
function linhaDe(a) {
  const c = a.cobranca || {};
  return [
    a.id, a.numero, a.excluido ? 'Excluído' : (STATUS[a.status] || a.status), a.os || '', a.placa || '', a.tecnico || '',
    dt(a.criadoEm), dt(a.finalizadoEm), dt(c.em_espera), dt(c.enviado), dt(c.concluido), dt(a.atualizadoEm),
    a.pasta || '', a.excluido ? 'Sim' : '', a.criadoPor || '', a.alteradoPor || '', a.servidorEm || 0, JSON.stringify(a),
  ];
}

function acharLinha(sh, id) {
  const n = sh.getLastRow() - 1;
  if (n < 1) return 0;
  const ids = sh.getRange(2, 1, n, 1).getValues();
  for (let i = 0; i < ids.length; i++) if (ids[i][0] === id) return i + 2;
  return 0;
}

function proximoNumero(sh) {
  const p = props();
  let n = Number(p.getProperty('ULTIMO_NUM')) || 0;
  const linhas = sh.getLastRow() - 1;
  if (linhas > 0) {
    sh.getRange(2, 2, linhas, 1).getValues().forEach(function (v) { n = Math.max(n, Number(v[0]) || 0); });
  }
  n++;
  p.setProperty('ULTIMO_NUM', String(n));
  return n;
}

function pastaDo(a) {
  const p = props();
  const k = 'P_' + a.id;
  const id = p.getProperty(k);
  if (id) { try { return DriveApp.getFolderById(id); } catch (e) {} }
  const f = pastaAtendimentos().createFolder(nomePasta(a));
  p.setProperty(k, f.getId());
  return f;
}

function nomePasta(a) {
  return [a.numero ? '#' + a.numero : 'Atendimento', a.placa, a.os ? 'OS ' + a.os : '']
    .filter(function (x) { return x; }).join(' - ');
}

function dt(ms) { return ms ? new Date(ms) : ''; }
