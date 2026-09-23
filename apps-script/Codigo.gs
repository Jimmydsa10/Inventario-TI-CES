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

// Coluna "editar" (4ª): só importa pra admin restrito (permissoes != 'todas').
// true = pode adicionar/editar/excluir nas seções liberadas pra ele; false
// (ou em branco, inclusive linhas antigas de antes dessa coluna existir) =
// só visualiza, sem poder mudar nada — ver filtrarEstadoPorPermissao.
// Colunas "podeAbrirChamados"/"podeResponderChamados" (5ª/6ª): só importam
// pra admin restrito com "chamados" nas seções liberadas — controlam,
// separado do "editar" geral, se ele pode abrir chamado novo e/ou responder
// chamado existente (mudar status, excluir, mandar mensagem). Em branco
// (inclusive linhas antigas de antes dessas colunas existirem) herda o valor
// de "editar", pra não mudar o comportamento de quem já estava cadastrado.
function getAdmins() {
  const sh = getOrCreateSheet('Admins', ['nome', 'senha', 'permissoes', 'editar', 'podeAbrirChamados', 'podeResponderChamados']);
  const data = sh.getDataRange().getValues();
  const rows = data.slice(1).filter(function (r) { return r[0] || r[1]; });
  if (rows.length === 0) {
    sh.appendRow(['Administrador', 'mude-esta-senha-123', 'todas', true, true, true]);
    return [{ nome: 'Administrador', senha: 'mude-esta-senha-123', permissoes: 'todas', editar: true, podeAbrirChamados: true, podeResponderChamados: true }];
  }
  return rows.map(function (r) {
    const editar = r[3] === true;
    const abrirEmBranco = r[4] === '' || r[4] === undefined || r[4] === null;
    const responderEmBranco = r[5] === '' || r[5] === undefined || r[5] === null;
    return {
      nome: String(r[0] || ''),
      senha: String(r[1] || ''),
      permissoes: String(r[2] || 'todas'),
      editar: editar,
      podeAbrirChamados: abrirEmBranco ? editar : r[4] === true,
      podeResponderChamados: responderEmBranco ? editar : r[5] === true,
    };
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
  const sh = getOrCreateSheet('Admins', ['nome', 'senha', 'permissoes', 'editar', 'podeAbrirChamados', 'podeResponderChamados']);
  const lastRow = sh.getLastRow();
  if (lastRow >= 2) {
    sh.getRange(2, 1, lastRow - 1, 6).clearContent();
  }
  // Se quem chamou (frontend atual, sempre manda os dois campos; uma
  // chamada externa/antiga pode não mandar) não informar
  // podeAbrirChamados/podeResponderChamados, herda de "editar" — mesma
  // regra de fallback usada pra linha em branco em getAdmins, só que aqui
  // pro objeto JS recebido em vez da célula da planilha.
  const rows = list.map(function (a) {
    const abrir = a.podeAbrirChamados === undefined ? !!a.editar : !!a.podeAbrirChamados;
    const responder = a.podeResponderChamados === undefined ? !!a.editar : !!a.podeResponderChamados;
    return [a.nome, a.senha, a.permissoes, !!a.editar, abrir, responder];
  });
  if (rows.length > 0) {
    sh.getRange(2, 1, rows.length, 6).setValues(rows);
  }
}

// ---------- Solicitantes (quem abre chamados) ----------

// Coluna "tipo" (3ª): 'solicitante' (padrão, inclusive linhas antigas de
// antes dessa coluna existir) = abre e responde chamados, como sempre foi.
// 'autorizado' = só visualização (Painel, Categorias, Salas e Inventário),
// sem poder abrir chamado, responder ou editar nada — ver authenticateSolicitante
// e o ramo userNome de doGet.
// Coluna "foto" (4ª): data URL (base64) da foto de perfil, redimensionada
// pequena no cliente antes de mandar. Em branco = usa as iniciais do nome
// (ver Avatar no frontend). Só o próprio usuário altera a própria foto, pela
// action 'atualizarFotoUsuario' — nunca escrita pelo salvarSolicitantes do
// master, que preserva o que já estava (ver Usuarios no frontend).
function getSolicitantes() {
  const sh = getOrCreateSheet('Solicitantes', ['nome', 'senha', 'tipo', 'foto']);
  const data = sh.getDataRange().getValues();
  return data.slice(1).filter(function (r) { return r[0]; }).map(function (r) {
    return { nome: String(r[0] || ''), senha: String(r[1] || ''), tipo: String(r[2] || 'solicitante') || 'solicitante', foto: String(r[3] || '') };
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
  const sh = getOrCreateSheet('Solicitantes', ['nome', 'senha', 'tipo', 'foto']);
  const lastRow = sh.getLastRow();
  if (lastRow >= 2) {
    sh.getRange(2, 1, lastRow - 1, 4).clearContent();
  }
  const rows = (list || []).map(function (s) { return [s.nome, s.senha, s.tipo === 'autorizado' ? 'autorizado' : 'solicitante', s.foto || '']; });
  if (rows.length > 0) {
    sh.getRange(2, 1, rows.length, 4).setValues(rows);
  }
}

// Atualiza só a foto de UM solicitante (o autenticado), preservando o resto
// da linha dele. Separado de saveSolicitantes (que é do master, pra lista
// inteira) porque aqui quem chama só pode mudar a própria foto — nunca a de
// outra pessoa.
function atualizarFotoSolicitante(nome, novaFoto) {
  const lista = getSolicitantes();
  const alvo = String(nome || '').trim().toLowerCase();
  const atualizada = lista.map(function (s) {
    if (s.nome.trim().toLowerCase() !== alvo) return s;
    return { nome: s.nome, senha: s.senha, tipo: s.tipo, foto: novaFoto || '' };
  });
  saveSolicitantes(atualizada);
}

// Autentica só nome+senha, sem olhar o tipo — usado no LOGIN (ramo userNome
// de doGet), onde tanto solicitante quanto autorizado podem entrar.
function autenticarUsuarioLogin(nome, senha) {
  const s = findSolicitante(nome);
  if (!s) return null;
  if (s.senha !== String(senha || '')) return null;
  return s;
}

// Retorna o registro do solicitante (com o nome como está cadastrado) se
// nome+senha baterem E ele não for do tipo 'autorizado', ou null. Usado pra
// autenticar quem pode abrir chamado e responder — sem isso, doPost aceitava
// novoChamado/novaMensagem de qualquer um que soubesse a URL pública do
// backend, sem checar login algum. Um "autorizado" é só visualização: mesmo
// logado, não pode criar chamado nem responder.
function authenticateSolicitante(nome, senha) {
  const s = autenticarUsuarioLogin(nome, senha);
  if (!s || s.tipo === 'autorizado') return null;
  return s;
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
        'Categoria: ' + (chamado.categoria || '-') + '\n\n' +
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

// Cada chave do estado corresponde a uma seção do app (a mesma dividida
// pelas permissões de admin restrito, tipo "chamados" ou "inventario").
const ESTADO_CAMPO_PARA_SECAO = {
  categorias: 'categorias',
  areas: 'areas',
  responsaveis: 'responsaveis',
  inventario: 'inventario',
  chamados: 'chamados',
};

// Antes: salvarTudo só checava "é um admin válido" — um admin com acesso
// restrito (ex: só "chamados") tecnicamente conseguia sobrescrever
// categorias, salas, inventário etc., já que o autosave manda o estado
// inteiro de uma vez e nada limitava quais partes ele podia de fato mudar.
// Agora, pra um admin não-master, cada seção do estado só é aceita do jeito
// que o cliente mandou se a permissão dele cobrir aquela seção E ele puder
// escrever nela; o resto mantém o valor que já estava salvo (a mudança é
// ignorada, não rejeitada por inteiro, pra não quebrar o autosave de quem só
// mexeu no que pode). Um admin restrito sem permissão de escrita não
// consegue mudar seção nenhuma por aqui — vira, na prática, um no-op.
// "chamados" é especial: em vez do "editar" geral, usa podeAbrirChamados OU
// podeResponderChamados (abrir chamado novo e responder/mudar status/excluir
// um existente chegam aqui juntos, dentro do mesmo array — ver Chamados no
// frontend, que já esconde os botões de cada ação conforme a permissão
// específica; aqui é só o gate de "pode escrever nessa seção ou não").
function filtrarEstadoPorPermissao(novoEstado, admin) {
  const isMaster = admin && admin.permissoes === 'todas';
  if (isMaster) return novoEstado;
  const secoesPermitidas = String((admin && admin.permissoes) || '').split(',').map(function (s) { return s.trim().toLowerCase(); });
  const atual = readState();
  const resultado = {};
  for (const chave in novoEstado) {
    const secao = ESTADO_CAMPO_PARA_SECAO[chave];
    const podeEscreverSecao = secao === 'chamados'
      ? !!(admin && (admin.podeAbrirChamados || admin.podeResponderChamados))
      : !!(admin && admin.editar);
    if (secao && secoesPermitidas.indexOf(secao) !== -1 && podeEscreverSecao) {
      resultado[chave] = novoEstado[chave];
    } else if (Object.prototype.hasOwnProperty.call(atual, chave)) {
      resultado[chave] = atual[chave];
    }
  }
  return resultado;
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
      editar: !!admin.editar,
      podeAbrirChamados: !!admin.podeAbrirChamados,
      podeResponderChamados: !!admin.podeResponderChamados,
      state: state,
      // Inclui a senha (como já fazemos pra solicitantes) porque o frontend
      // (Administradores.save()) usa modal.original.senha pra manter a senha
      // de quem já existe quando o campo "nova senha" fica em branco. Sem
      // isso, editar qualquer admin (ex: só mudar uma permissão) sem
      // redigitar a senha mandava senha undefined pro salvarAdmins e travava
      // o login desse admin — bug real, encontrado numa validação de
      // segurança antes de liberar pra uso real.
      admins: isMaster ? admins.map(function (a) {
        return {
          nome: a.nome,
          senha: a.senha,
          permissoes: a.permissoes,
          editar: !!a.editar,
          podeAbrirChamados: !!a.podeAbrirChamados,
          podeResponderChamados: !!a.podeResponderChamados,
        };
      }) : [],
      solicitantes: isMaster ? getSolicitantes() : [],
      historico: getHistoricoMensal()
    };
  } else if (userNome) {
    const usuario = autenticarUsuarioLogin(userNome, userSenha);
    if (usuario && usuario.tipo === 'autorizado') {
      // Autorizado: entra como um usuário comum (não como admin), mas com
      // visão só-leitura de Painel/Categorias/Salas/Inventário — sem
      // chamados dos outros, sem admins/solicitantes/histórico.
      const state = readState();
      payload = {
        ok: true,
        isAdmin: false,
        isUser: true,
        isAutorizado: true,
        nome: usuario.nome,
        foto: usuario.foto || '',
        state: {
          categorias: state.categorias || [],
          areas: state.areas || [],
          responsaveis: state.responsaveis || [],
          inventario: state.inventario || [],
          chamados: state.chamados || []
        }
      };
    } else if (usuario) {
      const state = readState();
      const alvo = userNome.trim().toLowerCase();
      const meusChamados = (state.chamados || []).filter(function (c) {
        return (c.criadoPor || '').trim().toLowerCase() === alvo;
      });
      payload = {
        ok: true,
        isAdmin: false,
        isUser: true,
        nome: usuario.nome,
        foto: usuario.foto || '',
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
  // Admin com acesso restrito e "editar" desmarcado só pode visualizar — não
  // conta como alguém que pode escrever, mesmo sendo um admin válido.
  const adminPodeEscrever = isAdmin && (isMaster || admin.editar);
  // novoChamado/novaMensagem (chamados abertos direto pelo admin, via secret,
  // sem userNome/userSenha) usam as permissões específicas de Chamados em
  // vez do "editar" geral — mesmo critério do filtrarEstadoPorPermissao.
  const adminPodeAbrirChamados = isAdmin && (isMaster || admin.podeAbrirChamados);
  const adminPodeResponderChamados = isAdmin && (isMaster || admin.podeResponderChamados);

  if (isAdmin && body.action === 'salvarTudo') {
    const estadoFiltrado = filtrarEstadoPorPermissao(body.state, admin);
    writeState(estadoFiltrado);
    registrarSnapshotMensal(estadoFiltrado);
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

  // Autoatendimento: o próprio usuário (solicitante ou autorizado) troca a
  // própria foto de perfil, sem precisar de admin. Autentica só com
  // nome+senha (não authenticateSolicitante, que bloqueia autorizado — aqui
  // ambos os tipos podem mudar a própria foto) e só mexe na foto, nunca em
  // nome/senha/tipo de ninguém.
  if (body.action === 'atualizarFotoUsuario') {
    const usuario = autenticarUsuarioLogin(body.userNome, body.userSenha);
    if (!usuario) {
      return jsonOut({ ok: false, error: 'Não autenticado' });
    }
    atualizarFotoSolicitante(usuario.nome, body.foto);
    return jsonOut({ ok: true });
  }

  // Antes: novoChamado e novaMensagem não checavam login nenhum — qualquer
  // um com a URL pública do backend podia criar chamados falsos ou postar
  // mensagens em qualquer chamado, inclusive se passando pelo TI
  // (mensagem.autor = 'ti'). Agora exige um solicitante autenticado (ou um
  // admin) e o autor/criadoPor são fixados pelo servidor, nunca aceitos do
  // jeito que o cliente mandou.
  if (body.action === 'novoChamado') {
    const solicitante = authenticateSolicitante(body.userNome, body.userSenha);
    if (!adminPodeAbrirChamados && !solicitante) {
      return jsonOut({ ok: false, error: 'Não autenticado' });
    }
    const state = readState();
    if (solicitante) {
      body.chamado.criadoPor = solicitante.nome;
      body.chamado.solicitante = solicitante.nome;
    }
    state.chamados = state.chamados || [];
    state.chamados.push(body.chamado);
    writeState(state);
    notificarNovoChamado(body.chamado);
    return jsonOut({ ok: true });
  }

  if (body.action === 'novaMensagem') {
    const solicitante = authenticateSolicitante(body.userNome, body.userSenha);
    if (!adminPodeResponderChamados && !solicitante) {
      return jsonOut({ ok: false, error: 'Não autenticado' });
    }
    const state = readState();
    const chamados = state.chamados || [];
    const chamado = chamados.filter(function (c) { return c.id === body.chamadoId; })[0];
    if (!chamado) {
      return jsonOut({ ok: false, error: 'Chamado não encontrado' });
    }
    if (solicitante && !adminPodeResponderChamados) {
      const dono = String(chamado.criadoPor || chamado.solicitante || '').trim().toLowerCase();
      if (dono !== solicitante.nome.trim().toLowerCase()) {
        return jsonOut({ ok: false, error: 'Sem permissão para responder este chamado' });
      }
    }
    const mensagem = body.mensagem || {};
    mensagem.autor = adminPodeResponderChamados ? 'ti' : 'solicitante';
    chamado.mensagens.push(mensagem);
    writeState(state);
    return jsonOut({ ok: true });
  }

  return jsonOut({ ok: false, error: 'Ação não autorizada' });
}
