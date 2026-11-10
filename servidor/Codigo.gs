/**
 * BLS Refrigeração · Atendimentos — servidor gratuito (Google Apps Script)
 *
 * Guarda os dados numa planilha Google e as fotos/vídeos/áudios numa pasta do Drive.
 *
 * COMO INSTALAR (uma vez só):
 *   1) script.google.com → Novo projeto → apague o que estiver lá e cole este código → Salvar
 *   2) No topo, escolha a função "configurar" e clique em Executar → autorize com a sua conta Google
 *   3) Implantar → Nova implantação → tipo "App da Web"
 *        Executar como: Eu   ·   Quem pode acessar: Qualquer pessoa   → Implantar
 *   4) Copie a URL do App da Web e a CHAVE DE ACESSO (aparece no "Registro de execução")
 *      e cole no app, em Configurações.
 */

const PASTA_ID = '1z5Vqz5UHFx0O3G6Fgc5k4tQ6mu5qrWyr';   // pasta do Drive da BLS
const VERSAO_API = '1.0';
const CAB = ['ID', 'Nº', 'Status', 'OS', 'Placa', 'Técnico', 'Criado em', 'Finalizado em',
             'Em espera', 'Enviado', 'Concluído', 'Atualizado em', 'Pasta', 'Excluído', 'Servidor', 'JSON'];
const STATUS = { pendente: 'Pendente', em_espera: 'Em espera', enviado: 'Enviado', concluido: 'Concluído' };

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
    const sh = ss.getSheets()[0];
    sh.setName('Atendimentos');
    sh.getRange(1, 1, 1, CAB.length).setValues([CAB])
      .setFontWeight('bold').setBackground('#142269').setFontColor('#ffffff');
    sh.setFrozenRows(1);
    p.setProperty('PLANILHA_ID', ss.getId());
  }
  pastaAtendimentos();
  if (!p.getProperty('CHAVE')) {
    p.setProperty('CHAVE', Utilities.getUuid().replace(/-/g, '').slice(0, 16).toUpperCase());
  }
  // pede a autorização usada para enviar vídeos grandes ao Drive
  UrlFetchApp.fetch('https://www.googleapis.com/drive/v3/about?fields=user',
    { headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }, muteHttpExceptions: true });

  Logger.log('✅ Tudo pronto!');
  Logger.log('Planilha: ' + ss.getUrl());
  Logger.log('CHAVE DE ACESSO: ' + p.getProperty('CHAVE'));
}

function pastaAtendimentos() {
  const p = props();
  const id = p.getProperty('PASTA_ATEND');
  if (id) { try { return DriveApp.getFolderById(id); } catch (e) {} }
  const f = DriveApp.getFolderById(PASTA_ID).createFolder('Atendimentos');
  p.setProperty('PASTA_ATEND', f.getId());
  return f;
}

function planilha() {
  const id = props().getProperty('PLANILHA_ID');
  if (!id) throw new Error('Servidor não configurado: execute a função "configurar".');
  return SpreadsheetApp.openById(id).getSheetByName('Atendimentos');
}

/* ---------- Entrada da API ---------- */
function doGet() {
  return ContentService.createTextOutput('BLS Atendimentos API ' + VERSAO_API + ' — funcionando.');
}

function doPost(e) {
  let r;
  try {
    const q = JSON.parse(e.postData.contents);
    if (!q.chave || q.chave !== props().getProperty('CHAVE')) throw new Error('Chave de acesso inválida');
    const f = ACOES[q.acao];
    if (!f) throw new Error('Ação desconhecida: ' + q.acao);
    r = f(q) || {};
    r.ok = true;
  } catch (err) {
    r = { ok: false, erro: String((err && err.message) || err) };
  }
  return ContentService.createTextOutput(JSON.stringify(r)).setMimeType(ContentService.MimeType.JSON);
}

const ACOES = {
  ping() { return { versao: VERSAO_API, agora: Date.now() }; },

  // Grava/atualiza um atendimento (vale a versão alterada por último)
  salvar(q) {
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
      const pasta = pastaDo(a);
      a.pastaId = pasta.getId();
      a.pasta = pasta.getUrl();
      const nome = nomePasta(a);
      if (pasta.getName() !== nome) pasta.setName(nome);
      if (a.excluido) { try { pasta.setTrashed(true); } catch (e) {} }
      delete a.sujo;
      a.servidorEm = Date.now();

      const c = a.cobranca || {};
      const linha = [
        a.id, a.numero, a.excluido ? 'Excluído' : (STATUS[a.status] || a.status), a.os || '', a.placa || '', a.tecnico || '',
        dt(a.criadoEm), dt(a.finalizadoEm), dt(c.em_espera), dt(c.enviado), dt(c.concluido), dt(a.atualizadoEm),
        a.pasta, a.excluido ? 'Sim' : '', a.servidorEm, JSON.stringify(a),
      ];
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
