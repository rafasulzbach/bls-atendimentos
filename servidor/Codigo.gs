/**
 * BLS Refrigeração · Atendimentos — servidor gratuito (Google Apps Script)
 *
 * Guarda os dados numa planilha Google e as fotos/vídeos/áudios numa pasta do Drive.
 *
 * COMO INSTALAR / ATUALIZAR:
 *   1) script.google.com → projeto "BLS Atendimentos" → apague tudo, cole este código → Salvar
 *   2) (o ID do cliente Google já está preenchido abaixo)
 *   3) Escolha a função "configurar" e clique em Executar (autorize se pedir)
 *   4) Implantar → Gerenciar implantações → ✏ Editar → Versão: "Nova versão" → Implantar
 *      (a URL continua a mesma)
 */

const PASTA_ID = '1z5Vqz5UHFx0O3G6Fgc5k4tQ6mu5qrWyr';   // pasta do Drive da BLS
const GOOGLE_CLIENT_ID = '1003853040855-he300vhna5upgis1774pagce9ltts9al.apps.googleusercontent.com';
const VERSAO_API = '1.5';

const CAB = ['ID', 'Nº', 'Status', 'OS', 'Placa', 'Técnico', 'Criado em', 'Finalizado em',
             'Em espera', 'Enviado', 'Concluído', 'Atualizado em', 'Pasta', 'Excluído',
             'Criado por', 'Alterado por', 'Servidor', 'JSON'];
const CAB_TESTE = ['ID', 'Nº', 'Status', 'OS', 'Placa', 'Data do teste', 'Avaliador', 'Local', 'Veículo (marca/ano)',
             'Equipamento (marca/modelo)', 'Nº série', 'Estacionário', 'Elétrico', 'Horímetro', 'Providências motor',
             'Vedação', 'Aspecto geral', 'Providências câmara', 'Temp. inicial', 'Temp. 1h', 'Temp. 2h', 'Temp. 3h',
             'Válido até', 'Criado por', 'Alterado por', 'Excluído', 'Servidor', 'JSON'];
const CAB_USU = ['E-mail', 'Nome (Google)', 'Situação', 'Criado em', 'Último acesso', 'Categoria', 'Nome no app'];
const ABAS = ['atendimentos', 'cobranca', 'teste', 'manuais'];                       // abas que podem ser liberadas
const PERM_PADRAO = { tecnico: ['atendimentos', 'teste', 'manuais'], geral: ['atendimentos', 'cobranca', 'teste', 'manuais'], v2: true };
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
  // aba de testes de frio
  let st = ss.getSheetByName('Testes de Frio');
  if (!st) {
    st = ss.insertSheet('Testes de Frio');
    st.getRange(1, 1, 1, CAB_TESTE.length).setValues([CAB_TESTE])
      .setFontWeight('bold').setBackground('#142269').setFontColor('#ffffff');
    st.setFrozenRows(1);
  }
  // aba de usuários
  let su = ss.getSheetByName('Usuarios');
  if (!su) {
    su = ss.insertSheet('Usuarios');
    su.getRange(1, 1, 1, CAB_USU.length).setValues([CAB_USU])
      .setFontWeight('bold').setBackground('#142269').setFontColor('#ffffff');
    su.setFrozenRows(1);
  }
  su.getRange(1, 1, 1, CAB_USU.length).setValues([CAB_USU])
    .setFontWeight('bold').setBackground('#142269').setFontColor('#ffffff');
  if (su.getLastRow() > 1) {   // usuários antigos sem categoria viram Técnico
    const r = su.getRange(2, 6, su.getLastRow() - 1, 1); const v = r.getValues();
    r.setValues(v.map(function (x) { return [x[0] || 'Técnico']; }));
  }
  pastaAtendimentos();
  abaManuais(); pastaManuais();
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
function abaTestes() { return ss().getSheetByName('Testes de Frio'); }

/* ---------- Entrada da API ---------- */
function doGet() {
  return ContentService.createTextOutput('BLS Atendimentos API ' + VERSAO_API + ' — funcionando.');
}

const PUBLICAS = { ping: 1, loginGoogle: 1 };
const SO_ADMIN = { usuarios: 1, usuarioAcao: 1, salvarPermissoes: 1, iaConfig: 1, iaTestar: 1, manualIniciar: 1, manualOcr: 1, manualConcluir: 1, manualExcluir: 1, pecasImportar: 1 };

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
    p.n = u.nomeApp || u.nome || p.n;
    p.cat = u.categoria === 'Geral' ? 'geral' : 'tecnico';
    p.reg = u;
  } else {
    p.n = props().getProperty('ADMIN_NOME') || p.n;
  }
  return p;
}
function sessaoDe(usuario, nome, adm) {
  const exp = Date.now() + (adm ? 30 : 365) * DIAS;
  return { token: assinar({ u: usuario, n: nome, adm: !!adm, exp: exp }), usuario: usuario, nome: nome, admin: !!adm, exp: exp };
}
function permissoes() {
  try {
    const p = JSON.parse(props().getProperty('PERMISSOES') || 'null');
    if (p && p.tecnico && p.geral) {
      if (!p.v2) { p.tecnico.push('manuais'); p.geral.push('manuais'); p.v2 = true; props().setProperty('PERMISSOES', JSON.stringify(p)); }
      return p;
    }
  } catch (e) {}
  return PERM_PADRAO;
}
function chaveUsuario(u) { return String(u || '').trim().toLowerCase(); }
function acharUsuario(u) {
  const su = abaUsuarios(), n = su.getLastRow() - 1, k = chaveUsuario(u);
  if (n < 1) return null;
  const vals = su.getRange(2, 1, n, CAB_USU.length).getValues();
  for (let i = 0; i < vals.length; i++) {
    if (String(vals[i][0]).toLowerCase() === k) return { linha: i + 2, email: k, nome: vals[i][1], situacao: vals[i][2], categoria: vals[i][5] || 'Técnico', nomeApp: vals[i][6] || '' };
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
        su.appendRow([email, nome, 'Pendente', new Date(), new Date(), 'Técnico', '']);
        return { pendente: true, nome: nome, email: email };
      }
      su.getRange(u.linha, 5).setValue(new Date());
      if (u.situacao === 'Pendente') return { pendente: true, nome: nome, email: email };
      if (u.situacao !== 'Liberado') throw new Error('Seu acesso está bloqueado. Fale com o administrador.');
      if (u.nome !== nome) su.getRange(u.linha, 2).setValue(nome);
      return { sessao: sessaoDe(email, u.nomeApp || nome, false) };
    } finally { lock.releaseLock(); }
  },

  // Perfil de quem está logado: nome no app, categoria e abas liberadas
  perfil(q, sess) {
    const perm = permissoes();
    if (sess.adm) {
      return { admin: true, email: sess.u, nomeApp: props().getProperty('ADMIN_NOME') || '', nomeGoogle: sess.n, categoria: 'Administrador', abas: ABAS.concat(['config']), permissoes: perm };
    }
    const u = sess.reg;
    return { admin: false, email: u.email, nomeApp: u.nomeApp, nomeGoogle: u.nome, categoria: u.categoria, abas: perm[sess.cat] || [] };
  },

  salvarPerfil(q, sess) {
    const nome = String(q.nomeApp || '').trim().replace(/\s+/g, ' ').slice(0, 60);
    if (nome.length < 2) throw new Error('Informe o nome');
    if (sess.adm) props().setProperty('ADMIN_NOME', nome);
    else abaUsuarios().getRange(sess.reg.linha, 7).setValue(nome);
    return { nomeApp: nome };
  },

  salvarPermissoes(q) {
    const limpa = function (l) { return (l || []).filter(function (a) { return ABAS.indexOf(a) >= 0; }); };
    const perm = { tecnico: limpa(q.permissoes && q.permissoes.tecnico), geral: limpa(q.permissoes && q.permissoes.geral), v2: true };
    props().setProperty('PERMISSOES', JSON.stringify(perm));
    return { permissoes: perm };
  },

  usuarios() {
    const su = abaUsuarios(), n = su.getLastRow() - 1;
    if (n < 1) return { usuarios: [] };
    return { usuarios: su.getRange(2, 1, n, CAB_USU.length).getValues().map(function (v) {
      return { email: v[0], nome: v[6] || v[1], nomeGoogle: v[1], situacao: v[2], categoria: v[5] || 'Técnico',
               criadoEm: v[3] ? new Date(v[3]).getTime() : null, ultimoAcesso: v[4] ? new Date(v[4]).getTime() : null };
    }), permissoes: permissoes() };
  },

  usuarioAcao(q) {
    const u = acharUsuario(q.email); if (!u) throw new Error('Usuário não encontrado');
    const su = abaUsuarios();
    if (q.op === 'liberar') su.getRange(u.linha, 3).setValue('Liberado');
    else if (q.op === 'bloquear') su.getRange(u.linha, 3).setValue('Bloqueado');
    else if (q.op === 'excluir') su.deleteRow(u.linha);
    else if (q.op === 'categoria') su.getRange(u.linha, 6).setValue(q.valor === 'Geral' ? 'Geral' : 'Técnico');
    else throw new Error('Operação inválida');
    return {};
  },

  // Grava/atualiza um atendimento (vale a versão alterada por último)
  salvar(q, sess) {
    const lock = LockService.getScriptLock();
    lock.waitLock(30000);
    try {
      const a = q.atend;
      if (!a || !a.id) throw new Error('Atendimento inválido');
      const teste = a.tipo === 'teste';
      const sh = teste ? abaTestes() : planilha();
      if (!sh) throw new Error('Execute a função "configurar" no servidor para criar a aba de testes.');
      const cab = teste ? CAB_TESTE : CAB;
      const lin = acharLinha(sh, a.id);
      let atual = null;
      if (lin) { try { atual = JSON.parse(sh.getRange(lin, cab.length).getValue()); } catch (e) {} }
      if (atual && (atual.atualizadoEm || 0) > (a.atualizadoEm || 0)) return { atend: atual };

      a.numero = atual && atual.numero ? atual.numero : proximoNumero(sh, teste ? 'ULTIMO_TESTE' : 'ULTIMO_NUM');
      a.criadoPor = (atual && atual.criadoPor) || a.criadoPor || sess.n;
      a.alteradoPor = sess.n;
      delete a.sujo;
      if (teste) {
        a.servidorEm = Date.now();
        const lt = linhaTeste(a);
        if (lin) sh.getRange(lin, 1, 1, lt.length).setValues([lt]); else sh.appendRow(lt);
        return { atend: a };
      }
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
      const agora = Date.now();
      const desde = Number(q.desde) || 0;
      const itens = [];
      [[planilha(), CAB], [abaTestes(), CAB_TESTE]].forEach(function (par) {
        const sh = par[0], cab = par[1];
        if (!sh) return;
        const n = sh.getLastRow() - 1;
        if (n < 1) return;
        sh.getRange(2, cab.length - 1, n, 2).getValues().forEach(function (v) {
          if (Number(v[0]) > desde) { try { itens.push(JSON.parse(v[1])); } catch (e) {} }
        });
      });
      return { agora: agora, itens: itens };
    } finally {
      lock.releaseLock();
    }
  },

  // Recebe foto/vídeo/áudio (em partes de até 4 MB) e grava na pasta do atendimento
  upload(q, sess) {
    const bytes = Utilities.base64Decode(q.data);
    let pasta;
    if (q.destino && String(q.destino).indexOf('manual:') === 0) {
      if (!sess.adm) throw new Error('Apenas o administrador pode enviar manuais');
      pasta = pastaManual(String(q.destino).slice(7));
    } else pasta = pastaDo({ id: q.atendId });
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

function linhaTeste(t) {
  const m = t.motor || {}, c = t.camara || {}, tp = t.temps || {};
  const num = function (v) { const x = parseFloat(String(v == null ? '' : v).replace(',', '.')); return isNaN(x) ? '' : x; };
  const dataTeste = t.data ? new Date(t.data + 'T12:00:00') : '';
  const validade = dataTeste ? new Date(dataTeste.getTime() + 180 * DIAS) : '';
  return [
    t.id, t.numero, t.excluido ? 'Excluído' : (t.status === 'concluido' ? 'Concluído' : 'Pendente'), t.os || '', t.placa || '',
    dataTeste, t.avaliador || '', t.local || '', t.veiculo || '', t.equipamento || '', t.serie || '',
    m.estacionario || '', m.eletrico || '', m.horimetro || '', m.prov || '', c.vedacao || '', c.aspecto || '', c.prov || '',
    num(tp.t0), num(tp.t1), num(tp.t2), num(tp.t3), validade,
    t.criadoPor || '', t.alteradoPor || '', t.excluido ? 'Sim' : '', t.servidorEm || 0, JSON.stringify(t),
  ];
}

function acharLinha(sh, id) {
  const n = sh.getLastRow() - 1;
  if (n < 1) return 0;
  const ids = sh.getRange(2, 1, n, 1).getValues();
  for (let i = 0; i < ids.length; i++) if (ids[i][0] === id) return i + 2;
  return 0;
}

function proximoNumero(sh, chave) {
  const p = props();
  let n = Number(p.getProperty(chave)) || 0;
  const linhas = sh.getLastRow() - 1;
  if (linhas > 0) {
    sh.getRange(2, 2, linhas, 1).getValues().forEach(function (v) { n = Math.max(n, Number(v[0]) || 0); });
  }
  n++;
  p.setProperty(chave, String(n));
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

/* =========================================================
   MANUAIS + IA (Gemini, plano gratuito)
   ========================================================= */
const CAB_MAN = ['ID', 'Marca', 'Modelo', 'Tipo', 'Título', 'Páginas', 'Arquivo', 'Link', 'Partes', 'Índice', 'Enviado por', 'Enviado em', 'Páginas sem texto'];
const MODELO_PADRAO = 'gemini-flash-latest';
const STOP = { de:1, da:1, do:1, das:1, dos:1, em:1, no:1, na:1, nos:1, nas:1, um:1, uma:1, o:1, a:1, os:1, as:1, e:1, ou:1, que:1, qual:1, quais:1, como:1, para:1, por:1, com:1, sem:1, se:1, ao:1, aos:1, eh:1, sao:1, ser:1, esta:1, isso:1, este:1, essa:1, esse:1, onde:1, quando:1, mais:1, menos:1, muito:1, tem:1, ter:1, fazer:1, faz:1, pelo:1, pela:1, meu:1, minha:1, seu:1, sua:1, the:1, of:1, and:1, to:1, in:1, is:1, for:1, on:1, what:1, how:1 };
const SISTEMA_IA = 'Você é o assistente técnico da Refrigeração BLS (autorizada Carrier Transicold). ' +
  'Responda em português do Brasil, de forma direta e prática para um técnico de campo. ' +
  'Use SOMENTE as informações dos TRECHOS fornecidos (páginas de manuais e linhas da tabela de peças). ' +
  'Os trechos podem estar em inglês: traduza a resposta para o português, mas mantenha códigos, valores e nomes de peças exatamente como estão. ' +
  'Nunca invente valores, códigos, torques, pressões ou procedimentos. Se os trechos não trouxerem a resposta, diga claramente que não encontrou essa informação nos manuais cadastrados. ' +
  'Responda em JSON no formato {"resposta": "texto (pode usar **negrito** e listas com - )", "fontes": [números n dos trechos [n] usados], "pecas": [números das linhas P usadas]}.';

function abaManuais() {
  const s = ss();
  let sh = s.getSheetByName('Manuais');
  if (!sh) {
    sh = s.insertSheet('Manuais');
    sh.getRange(1, 1, 1, CAB_MAN.length).setValues([CAB_MAN]).setFontWeight('bold').setBackground('#142269').setFontColor('#ffffff');
    sh.setFrozenRows(1);
  }
  return sh;
}
function abaPecas() { return ss().getSheetByName('Peças'); }
function pastaManuais() {
  const p = props();
  const id = p.getProperty('PASTA_MANUAIS');
  if (id) { try { return DriveApp.getFolderById(id); } catch (e) {} }
  const f = DriveApp.getFolderById(PASTA_ID).createFolder('Manuais');
  p.setProperty('PASTA_MANUAIS', f.getId());
  return f;
}
function pastaManual(id) {
  const fid = props().getProperty('M_' + id);
  if (!fid) throw new Error('Manual não encontrado');
  return DriveApp.getFolderById(fid);
}
function norm(s) { return String(s || '').toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, ''); }
const SINONIMOS = {
  alarme: ['alarm'], alarmes: ['alarm'], codigo: ['code'], pressao: ['pressure'], succao: ['suction'], descarga: ['discharge'],
  temperatura: ['temperature', 'temp'], oleo: ['oil'], filtro: ['filter'], correia: ['belt'], ventilador: ['fan'], valvula: ['valve'],
  fusivel: ['fuse'], bateria: ['battery'], motor: ['engine', 'motor'], degelo: ['defrost'], evaporador: ['evaporator'],
  condensador: ['condenser'], troca: ['change', 'replace'], trocar: ['change', 'replace'], intervalo: ['interval'], nivel: ['level'],
  partida: ['start', 'starter'], desligamento: ['shutdown'], aquecimento: ['heat', 'heating'], resfriamento: ['cool', 'cooling'],
  refrigerante: ['refrigerant'], gas: ['refrigerant'], carga: ['charge'], vazamento: ['leak'], sensor: ['sensor'],
  manutencao: ['maintenance', 'service'], preventiva: ['preventive'], torque: ['torque'], aperto: ['torque'], rotacao: ['rpm', 'speed'],
  velocidade: ['speed'], combustivel: ['fuel'], bomba: ['pump'], injetor: ['injector'], alternador: ['alternator'], embreagem: ['clutch'],
  compressor: ['compressor'], controlador: ['controller'], painel: ['display', 'panel'], tela: ['display'], porta: ['door'],
  falha: ['fault', 'failure'], erro: ['error', 'fault'], defeito: ['fault'], peca: ['part'], pecas: ['parts'], numero: ['number'],
  serie: ['serial'], diagrama: ['diagram'], eletrico: ['electrical', 'wiring'], fiacao: ['wiring'], rele: ['relay'], chave: ['switch'],
  agua: ['water', 'coolant'], arrefecimento: ['coolant', 'cooling'], radiador: ['radiator'], umidade: ['humidity'], ar: ['air'],
};
function termosDe(texto) {
  const vistos = {};
  return norm(texto).split(/[^a-z0-9\-\.\/]+/).map(function (t) { return t.replace(/^[\-\.\/]+|[\-\.\/]+$/g, ''); })
    .filter(function (t) { if (!t || STOP[t] || vistos[t] || (t.length < 3 && !/\d/.test(t))) return false; vistos[t] = 1; return true; })
    .map(function (t) {
      const g = [t];
      if (!/\d/.test(t) && t.length >= 6) g.push(t.slice(0, t.length - 2));          // radical: alarme → alar…, temperaturas → temperatur…
      (SINONIMOS[t] || SINONIMOS[t.replace(/s$/, '')] || []).forEach(function (x) { if (g.indexOf(x) < 0) g.push(x); });
      return g;
    });
}
function pesoTermo(g) { return /\d/.test(g[0]) ? 4 : (g[0].length >= 7 ? 2 : 1); }
function contaGrupo(txt, g) { let m = 0; g.forEach(function (v) { m = Math.max(m, txt.split(v).length - 1); }); return Math.min(5, m); }
function pontuar(txtNorm, termos) {
  let sc = 0, achou = 0;
  termos.forEach(function (g) {
    const n = contaGrupo(txtNorm, g);
    if (n > 0) { achou++; sc += pesoTermo(g) * (1 + Math.log(n)); }
  });
  for (let i = 0; i < termos.length - 1; i++) if (txtNorm.indexOf(termos[i][0] + ' ' + termos[i + 1][0]) >= 0) sc += 3;
  return achou ? sc * (achou / termos.length + 0.5) : 0;
}
function listaManuais() {
  const sh = abaManuais(), n = sh.getLastRow() - 1;
  if (n < 1) return [];
  return sh.getRange(2, 1, n, CAB_MAN.length).getValues().map(function (v, i) {
    let partes = []; try { partes = JSON.parse(v[8] || '[]'); } catch (e) {}
    return { linha: i + 2, id: v[0], marca: v[1], modelo: v[2], tipo: v[3], titulo: v[4], paginas: v[5], arquivo: v[6], link: v[7], partes: partes,
             indice: v[9], por: v[10], em: v[11] ? new Date(v[11]).getTime() : null, semTexto: v[12] || 0 };
  }).filter(function (m) { return m.id; });
}
function lerIndice(m) {
  try { return JSON.parse(Utilities.ungzip(DriveApp.getFileById(m.indice).getBlob().setContentType('application/x-gzip')).getDataAsString()); }
  catch (e) { return null; }
}
function buscarPecas(termos, max) {
  const sh = abaPecas();
  if (!sh || sh.getLastRow() < 2 || !termos.length) return { cabecalho: [], linhas: [] };
  const vals = sh.getDataRange().getDisplayValues();
  const cab = vals.shift();
  const res = [];
  vals.forEach(function (row) {
    const j = norm(row.join(' | '));
    let sc = 0;
    termos.forEach(function (g) { if (g.some(function (t) { return j.indexOf(t) >= 0; })) sc += /\d/.test(g[0]) ? 10 : (g[0].length >= 4 ? 2 : 0.5); });
    if (sc >= 2) res.push({ sc: sc, row: row });
  });
  res.sort(function (a, b) { return b.sc - a.sc; });
  return { cabecalho: cab, linhas: res.slice(0, max).map(function (r) { return r.row; }) };
}
function gemini(parts, sistema) {
  const p = props();
  const key = p.getProperty('GEMINI_KEY');
  if (!key) throw new Error('A chave da IA ainda não foi configurada (Ajustes → IA dos manuais).');
  const modelo = p.getProperty('GEMINI_MODELO') || MODELO_PADRAO;
  const body = { contents: [{ role: 'user', parts: parts }], generationConfig: { temperature: 0.2, responseMimeType: 'application/json' } };
  if (sistema) body.systemInstruction = { parts: [{ text: sistema }] };
  const r = UrlFetchApp.fetch('https://generativelanguage.googleapis.com/v1beta/models/' + encodeURIComponent(modelo) + ':generateContent', {
    method: 'post', contentType: 'application/json', headers: { 'x-goog-api-key': key }, payload: JSON.stringify(body), muteHttpExceptions: true,
  });
  const code = r.getResponseCode(), txt = r.getContentText();
  if (code === 429) throw new Error('Limite gratuito da IA atingido no momento. Tente de novo em alguns minutos.');
  if (code !== 200) { let m = ''; try { m = JSON.parse(txt).error.message; } catch (e) { m = txt.slice(0, 200); } throw new Error('Erro da IA (' + code + '): ' + m); }
  const j = JSON.parse(txt);
  const c = j.candidates && j.candidates[0];
  const out = c && c.content && c.content.parts ? c.content.parts.map(function (x) { return x.text || ''; }).join('') : '';
  if (!out) throw new Error('A IA não retornou resposta. Tente reformular a pergunta.');
  try { return JSON.parse(out); } catch (e) { const m = out.match(/\{[\s\S]*\}/); if (m) return JSON.parse(m[0]); throw new Error('Resposta inválida da IA'); }
}

const ACOES_MANUAIS = {
  // Lista de manuais + informações da tabela de peças
  manuais() {
    const pi = props().getProperty('PECAS_INFO');
    return { manuais: listaManuais().map(function (m) { delete m.indice; return m; }), pecas: pi ? JSON.parse(pi) : null };
  },

  perguntar(q) {
    const pergunta = String(q.pergunta || '').trim();
    if (pergunta.length < 3) throw new Error('Escreva a pergunta.');
    const termos = termosDe(pergunta);
    if (!termos.length) throw new Error('Escreva a pergunta com mais detalhes (modelo, código, sintoma).');
    let mans = listaManuais();
    if (q.marca) mans = mans.filter(function (m) { return m.marca === q.marca; });
    const cands = [];
    mans.forEach(function (m) {
      const idx = lerIndice(m); if (!idx) return;
      const meta = norm([m.marca, m.modelo, m.tipo, m.titulo].join(' '));
      const boost = termos.some(function (g) { return g.some(function (t) { return meta.indexOf(t) >= 0; }); }) ? 1.6 : 1;
      idx.paginas.forEach(function (txt, i) {
        if (!txt) return;
        const sc = pontuar(norm(txt), termos) * boost;
        if (sc > 0) cands.push({ m: m, p: i + 1, sc: sc, txt: txt });
      });
    });
    cands.sort(function (a, b) { return b.sc - a.sc; });
    const top = cands.slice(0, 6);
    const pecas = buscarPecas(termos, 12);
    if (!top.length && !pecas.linhas.length) {
      return { resposta: 'Não encontrei nada sobre isso nos manuais e na tabela de peças cadastrados. Tente usar outras palavras, o código do alarme ou o modelo do equipamento.', fontes: [], pecas: { cabecalho: [], linhas: [] } };
    }
    let ctx = '';
    top.forEach(function (c, i) { ctx += '\n[' + (i + 1) + '] ' + c.m.titulo + ' — página ' + c.p + '\n' + String(c.txt).slice(0, 4500) + '\n'; });
    if (pecas.linhas.length) ctx += '\n[TABELA DE PEÇAS] colunas: ' + pecas.cabecalho.join(' | ') + '\n' + pecas.linhas.map(function (l, i) { return 'P' + (i + 1) + ': ' + l.join(' | '); }).join('\n') + '\n';
    const r = gemini([{ text: 'PERGUNTA DO TÉCNICO: ' + pergunta + '\n\nTRECHOS:' + ctx }], SISTEMA_IA);
    const vistos = {};
    const fontes = (r.fontes || []).map(function (n) { return parseInt(String(n).replace(/\D/g, ''), 10); })
      .filter(function (n) { return n >= 1 && n <= top.length && !vistos[n] && (vistos[n] = 1); })
      .map(function (n) { const c = top[n - 1]; return { manualId: c.m.id, titulo: c.m.titulo, pagina: c.p, link: c.m.link, partes: c.m.partes }; });
    const pu = (r.pecas || []).map(function (n) { return parseInt(String(n).replace(/\D/g, ''), 10); })
      .filter(function (n) { return n >= 1 && n <= pecas.linhas.length; });
    return { resposta: String(r.resposta || ''), fontes: fontes, pecas: { cabecalho: pecas.cabecalho, linhas: pu.map(function (n) { return pecas.linhas[n - 1]; }) } };
  },

  pecasBuscar(q) {
    const termos = termosDe(q.busca || '');
    if (!termos.length) return { cabecalho: [], linhas: [] };
    return buscarPecas(termos, 40);
  },

  // ----- administrador -----
  iaConfig(q) {
    const p = props();
    if (q.chave) p.setProperty('GEMINI_KEY', String(q.chave).trim());
    if (q.modelo) p.setProperty('GEMINI_MODELO', String(q.modelo).trim());
    const k = p.getProperty('GEMINI_KEY') || '';
    return { temChave: !!k, final: k ? k.slice(-4) : '', modelo: p.getProperty('GEMINI_MODELO') || MODELO_PADRAO };
  },
  iaTestar() {
    const r = gemini([{ text: 'Responda apenas o JSON {"resposta":"ok"}' }]);
    return { resposta: r.resposta || 'ok' };
  },
  manualIniciar(q) {
    const id = Utilities.getUuid().slice(0, 13);
    const titulo = String(q.titulo || 'Manual').slice(0, 120);
    const f = pastaManuais().createFolder(titulo);
    props().setProperty('M_' + id, f.getId());
    return { id: id };
  },
  manualOcr(q) {
    const parts = (q.imagens || []).map(function (b64) { return { inlineData: { mimeType: 'image/jpeg', data: b64 } }; });
    parts.push({ text: 'As imagens acima são as páginas ' + (q.paginas || []).join(', ') + ' de um manual técnico de refrigeração. ' +
      'Transcreva fielmente todo o texto de cada página, incluindo tabelas (uma linha por linha da tabela, colunas separadas por " | "), códigos e valores. ' +
      'Não resuma nem traduza. Responda em JSON: {"paginas": [{"n": número da página, "texto": "..."}]} na mesma ordem das imagens.' });
    const r = gemini(parts);
    return { paginas: (r.paginas || []).map(function (x, i) { return { n: Number(x.n) || (q.paginas || [])[i], texto: String(x.texto || '') }; }) };
  },
  manualConcluir(q, sess) {
    const pasta = pastaManual(q.id);
    const idx = Utilities.gzip(Utilities.newBlob(JSON.stringify({ paginas: q.indice || [] }), 'application/json'), 'indice.json.gz');
    const fi = pasta.createFile(idx);
    let link = '';
    try {
      const f = DriveApp.getFileById(q.arquivoId);
      f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      link = 'https://drive.google.com/file/d/' + q.arquivoId + '/view';
    } catch (e) { link = 'https://drive.google.com/file/d/' + q.arquivoId + '/view'; }
    abaManuais().appendRow([q.id, q.marca || '', q.modelo || '', q.tipo || '', q.titulo || '', q.paginas || 0, q.arquivoId, link,
      JSON.stringify(q.partes || []), fi.getId(), sess.n, new Date(), q.semTexto || 0]);
    return { link: link };
  },
  manualExcluir(q) {
    const m = listaManuais().filter(function (x) { return x.id === q.id; })[0];
    if (!m) throw new Error('Manual não encontrado');
    try { pastaManual(q.id).setTrashed(true); } catch (e) {}
    abaManuais().deleteRow(m.linha);
    return {};
  },
  pecasImportar(q) {
    const s = ss();
    let sh = s.getSheetByName('Peças');
    if (q.inicio) {
      if (!sh) sh = s.insertSheet('Peças'); else sh.clear();
      sh.getRange(1, 1, 1, q.cabecalho.length).setValues([q.cabecalho]).setFontWeight('bold').setBackground('#142269').setFontColor('#ffffff');
      sh.setFrozenRows(1);
    }
    const linhas = (q.linhas || []).map(function (l) { return q.cabecalho.map(function (_, i) { return l[i] == null ? '' : String(l[i]); }); });
    if (linhas.length) sh.getRange(sh.getLastRow() + 1, 1, linhas.length, q.cabecalho.length).setNumberFormat('@').setValues(linhas);
    if (q.fim) props().setProperty('PECAS_INFO', JSON.stringify({ linhas: sh.getLastRow() - 1, arquivo: q.arquivo || '', em: Date.now() }));
    return { total: sh.getLastRow() - 1 };
  },
};
Object.keys(ACOES_MANUAIS).forEach(function (k) { ACOES[k] = ACOES_MANUAIS[k]; });
