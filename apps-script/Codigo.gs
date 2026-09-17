// Código do backend em Google Apps Script (Extensões > Apps Script na
// planilha "Inventário de TI - Escola Espírito Santo (Completo)").
// Mantido aqui só como referência/histórico de versão — o Apps Script não
// lê este arquivo diretamente; qualquer mudança precisa ser colada também
// no editor do Apps Script e publicada como nova versão do deploy.

function getOrCreateSheet(name, header) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    if (header) sh.appendRow(header);
  }
  return sh;
}

// ---------- Administradores ----------

function getAdmins() {
  const sh = getOrCreateSheet('Admins', ['nome', 'senha', 'permissoes']);
  const data = sh.getDataRange().getValues();
  const rows = data.slice(1).filter(function (r) { return r[0] || r[1]; });
  if (rows.length === 0) {
    sh.appendRow(['Administrador', 'mude-esta-senha-123', 'todas']);
    return [{ nome: 'Administrador', senha: 'mude-esta-senha-123', permissoes: 'todas' }];
  }
  return rows.map(function (r) {
    return { nome: String(r[0] || ''), senha: String(r[1] || ''), permissoes: String(r[2] || 'todas') };
  });
}

// Aceita uma lista já carregada (adminsList), pra evitar ler a aba "Admins"
// duas vezes quando quem chama já buscou a lista por outro motivo.
function findAdminBySecret(secret, adminsList) {
  if (!secret) return null;
  const admins = adminsList || getAdmins();
  for (let i = 0; i < admins.length; i++) {
    if (admins[i].senha === secret) return admins[i];
  }
  return null;
}

function saveAdmins(list) {
  const sh = getOrCreateSheet('Admins', ['nome', 'senha', 'permissoes']);
  const lastRow = sh.getLastRow();
  if (lastRow >= 2) {
    sh.getRange(2, 1, lastRow - 1, 3).clearContent();
  }
  const rows = list.map(function (a) { return [a.nome, a.senha, a.permissoes]; });
  if (rows.length > 0) {
    sh.getRange(2, 1, rows.length, 3).setValues(rows);
  }
}

// ---------- Solicitantes (quem abre chamados) ----------

function getSolicitantes() {
  const sh = getOrCreateSheet('Solicitantes', ['nome', 'senha']);
  const data = sh.getDataRange().getValues();
  return data.slice(1).filter(function (r) { return r[0]; }).map(function (r) {
    return { nome: String(r[0] || ''), senha: String(r[1] || '') };
  });
}

function findSolicitante(nome) {
  const alvo = String(nome || '').trim().toLowerCase();
  if (!alvo) return null;
  const lista = getSolicitantes();
  for (let i = 0; i < lista.length; i++) {
    if (lista[i].nome.trim().toLowerCase() === alvo) return lista[i];
  }
  return null;
}

function saveSolicitantes(list) {
  const sh = getOrCreateSheet('Solicitantes', ['nome', 'senha']);
  const lastRow = sh.getLastRow();
  if (lastRow >= 2) {
    sh.getRange(2, 1, lastRow - 1, 2).clearContent();
  }
  const rows = (list || []).map(function (s) { return [s.nome, s.senha]; });
  if (rows.length > 0) {
    sh.getRange(2, 1, rows.length, 2).setValues(rows);
  }
}

function checkSolicitanteLogin(nome, senha) {
  const s = findSolicitante(nome);
  if (!s) return false;
  return s.senha === String(senha || '');
}

function getConfigValor(chave) {
  const sh = getOrCreateSheet('Config', ['chave', 'valor']);
  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim().toLowerCase() === chave) {
      return String(data[i][1] || '').trim();
    }
  }
  return '';
}

function getMesAtual() {
  const d = new Date();
  const ano = d.getFullYear();
  const mes = String(d.getMonth() + 1).padStart(2, '0');
  return ano + '-' + mes;
}

function contarPor(lista, chave, valor) {
  return lista.filter(function (x) { return x[chave] === valor; }).length;
}

// Converte o valor da célula "mes" (que o Google Sheets pode ter
// auto-convertido pra Data de verdade, em vez de manter como texto
// "AAAA-MM") de volta para "AAAA-MM". Sem isso, a checagem de "esse mês já
// tem snapshot?" nunca batia e cada autosave criava uma linha duplicada em
// "Historico".
function mesDaLinha(valor) {
  if (valor instanceof Date) {
    const ano = valor.getFullYear();
    const mes = String(valor.getMonth() + 1).padStart(2, '0');
    return ano + '-' + mes;
  }
  return String(valor || '').slice(0, 7);
}

// Evita reler a aba "Historico" inteira a cada autosave: só confere a
// planilha de fato na primeira vez que o mês vira. A coluna A é forçada
// pra texto simples pra impedir o Sheets de converter "AAAA-MM" em Data.
function registrarSnapshotMensal(state) {
  try {
    const props = PropertiesService.getScriptProperties();
    const mesAtual = getMesAtual();
    if (props.getProperty('ultimoSnapshotMes') === mesAtual) return;

    const sh = getOrCreateSheet('Historico', ['mes', 'totalEquipamentos', 'emUso', 'emManutencao', 'precisaManutencao', 'emEstoque', 'chamadosAbertos', 'chamadosAndamento', 'chamadosResolvidos']);
    sh.getRange('A:A').setNumberFormat('@');
    const data = sh.getDataRange().getValues();
    const jaTem = data.slice(1).some(function (r) { return mesDaLinha(r[0]) === mesAtual; });
    if (!jaTem) {
      const inv = state.inventario || [];
      const cham = state.chamados || [];
      sh.appendRow([
        mesAtual,
        inv.length,
        contarPor(inv, 'status', 'Em uso'),
        contarPor(inv, 'status', 'Em manutenção'),
        contarPor(inv, 'status', 'Precisa de manutenção'),
        contarPor(inv, 'status', 'Em estoque'),
        contarPor(cham, 'status', 'Aberto'),
        contarPor(cham, 'status', 'Em andamento'),
        contarPor(cham, 'status', 'Resolvido')
      ]);
    }
    props.setProperty('ultimoSnapshotMes', mesAtual);
  } catch (err) {
    // não impede o salvamento principal se o histórico falhar
  }
}

function getHistoricoMensal() {
  const sh = getOrCreateSheet('Historico', ['mes', 'totalEquipamentos', 'emUso', 'emManutencao', 'precisaManutencao', 'emEstoque', 'chamadosAbertos', 'chamadosAndamento', 'chamadosResolvidos']);
  const data = sh.getDataRange().getValues();
  return data.slice(1).filter(function (r) { return r[0]; }).map(function (r) {
    return {
      mes: mesDaLinha(r[0]),
      totalEquipamentos: r[1] || 0,
      emUso: r[2] || 0,
      emManutencao: r[3] || 0,
      precisaManutencao: r[4] || 0,
      emEstoque: r[5] || 0,
      chamadosAbertos: r[6] || 0,
      chamadosAndamento: r[7] || 0,
      chamadosResolvidos: r[8] || 0
    };
  });
}

// Utilitário de execução manual (uma vez só, pelo editor do Apps Script)
// pra colapsar linhas duplicadas em "Historico", mantendo uma por mês.
// Necessário depois de corrigir o bug de duplicação acima.
function limparHistoricoDuplicado() {
  const sh = getOrCreateSheet('Historico', ['mes', 'totalEquipamentos', 'emUso', 'emManutencao', 'precisaManutencao', 'emEstoque', 'chamadosAbertos', 'chamadosAndamento', 'chamadosResolvidos']);
  const data = sh.getDataRange().getValues();
  const header = data[0];
  const vistos = {};
  const unicos = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const mes = mesDaLinha(row[0]);
    if (!mes || vistos[mes]) continue;
    vistos[mes] = true;
    row[0] = mes;
    unicos.push(row);
  }
  sh.clearContents();
  sh.appendRow(header);
  sh.getRange('A:A').setNumberFormat('@');
  if (unicos.length > 0) {
    sh.getRange(2, 1, unicos.length, header.length).setValues(unicos);
  }
  Logger.log('Historico limpo: ' + unicos.length + ' mes(es) unico(s) restante(s).');
}

function notificarNovoChamado(chamado) {
  try {
    const email = getConfigValor('email_notificacao');
    if (!email) return;
    const primeiraMsg = (chamado.mensagens && chamado.mensagens[0] && chamado.mensagens[0].texto) || '';
    MailApp.sendEmail({
      to: email,
      subject: 'Novo chamado de TI: ' + chamado.assunto,
      body:
        'Solicitante: ' + (chamado.criadoPor || chamado.solicitante || '') + '\n' +
        'Sala: ' + (chamado.sala || '-') + '\n' +
        'Categoria: ' + (chamado.categoria || '-') + '\n' +
        'Prioridade: ' + (chamado.prioridade || '-') + '\n\n' +
        primeiraMsg
    });
  } catch (err) {
    // se o envio falhar (ex: permissão não concedida ainda), não impede o chamado de ser salvo
  }
}

// ---------- Estado do app (inventário, categorias, etc.) ----------

function readState() {
  const sh = getOrCreateSheet('AppState', ['data']);
  const lastRow = sh.getLastRow();
  const defaults = { categorias: [], areas: [], responsaveis: [], inventario: [], chamados: [] };
  if (lastRow < 2) return defaults;
  const values = sh.getRange(2, 1, lastRow - 1, 1).getValues();
  const combined = values.map(function (r) { return r[0]; }).join('');
  if (!combined) return defaults;
  try {
    return JSON.parse(combined);
  } catch (e) {
    return defaults;
  }
}

function writeState(obj) {
  const sh = getOrCreateSheet('AppState', ['data']);
  const json = JSON.stringify(obj);
  const CHUNK = 45000;
  const lastRow = sh.getLastRow();
  if (lastRow >= 2) {
    sh.getRange(2, 1, lastRow - 1, 1).clearContent();
  }
  const chunks = [];
  for (let i = 0; i < json.length; i += CHUNK) {
    chunks.push([json.slice(i, i + CHUNK)]);
  }
  if (chunks.length === 0) chunks.push(['']);
  sh.getRange(2, 1, chunks.length, 1).setValues(chunks);
}

function jsonOut(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// Antes: sempre chamava readState() e getAdmins(), mesmo quando o resultado
// não era usado (sem login, senha errada, ou pra montar a lista de admins
// que já tinha sido lida). Agora só lê a planilha quando o valor é
// realmente necessário pra montar a resposta.
function doGet(e) {
  const p = (e && e.parameter) || {};
  const callback = p.callback || '';
  let payload;

  const secret = p.secret || '';
  const userNome = p.userNome || '';
  const userSenha = p.userSenha || '';

  const admins = secret ? getAdmins() : null;
  const admin = findAdminBySecret(secret, admins);

  if (admin) {
    const isMaster = admin.permissoes === 'todas';
    const state = readState();
    payload = {
      ok: true,
      isAdmin: true,
      nome: admin.nome,
      permissoes: admin.permissoes,
      state: state,
      admins: isMaster ? admins.map(function (a) { return { nome: a.nome, permissoes: a.permissoes }; }) : [],
      solicitantes: isMaster ? getSolicitantes() : [],
      historico: getHistoricoMensal()
    };
  } else if (userNome) {
    if (checkSolicitanteLogin(userNome, userSenha)) {
      const state = readState();
      const alvo = userNome.trim().toLowerCase();
      const meusChamados = (state.chamados || []).filter(function (c) {
        return (c.criadoPor || '').trim().toLowerCase() === alvo;
      });
      payload = {
        ok: true,
        isAdmin: false,
        isUser: true,
        nome: userNome,
        state: { areas: state.areas || [], categorias: state.categorias || [], chamados: meusChamados }
      };
    } else {
      payload = { ok: true, isAdmin: false, isUser: false, error: 'Nome ou senha incorretos.' };
    }
  } else {
    payload = { ok: true, isAdmin: false, isUser: false, state: { areas: [], categorias: [], chamados: [] } };
  }

  if (callback) {
    return ContentService
      .createTextOutput(callback + '(' + JSON.stringify(payload) + ');')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return jsonOut(payload);
}

// Antes: sempre chamava readState() no topo, mesmo pra 'salvarTudo' e
// 'salvarAdmins' — que sobrescrevem o estado sem nunca usar o que acabou
// de ser lido. Isso tirava uma leitura completa da planilha em TODO
// autosave (a ação mais frequente do app). Agora só lê quando precisa.
function doPost(e) {
  const body = JSON.parse(e.postData.contents);
  const admin = findAdminBySecret(body.secret);
  const isAdmin = !!admin;
  const isMaster = admin && admin.permissoes === 'todas';

  if (isAdmin && body.action === 'salvarTudo') {
    writeState(body.state);
    registrarSnapshotMensal(body.state);
    return jsonOut({ ok: true });
  }

  if (isMaster && body.action === 'salvarAdmins') {
    saveAdmins(body.admins || []);
    return jsonOut({ ok: true });
  }

  if (isMaster && body.action === 'salvarSolicitantes') {
    saveSolicitantes(body.solicitantes || []);
    return jsonOut({ ok: true });
  }

  if (body.action === 'novoChamado') {
    const state = readState();
    body.chamado.criadoPor = body.userNome || body.chamado.criadoPor || '';
    state.chamados = state.chamados || [];
    state.chamados.push(body.chamado);
    writeState(state);
    notificarNovoChamado(body.chamado);
    return jsonOut({ ok: true });
  }

  if (body.action === 'novaMensagem') {
    const state = readState();
    const chamados = state.chamados || [];
    const chamado = chamados.filter(function (c) { return c.id === body.chamadoId; })[0];
    if (!chamado) {
      return jsonOut({ ok: false, error: 'Chamado não encontrado' });
    }
    chamado.mensagens.push(body.mensagem);
    writeState(state);
    return jsonOut({ ok: true });
  }

  return jsonOut({ ok: false, error: 'Ação não autorizada' });
}
