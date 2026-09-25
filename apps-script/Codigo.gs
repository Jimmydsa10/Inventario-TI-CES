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
// Coluna "responderSoProprios" (7ª): só importa junto com podeResponderChamados
// true — restringe esse admin a responder/mudar status/excluir só os
// chamados que ELE MESMO abriu (ver "abertoPorAdmin" nos chamados e o uso
// dessa flag em doPostComTrava), sem mexer nos chamados de outras pessoas ou
// de outros admins. Em branco = false (comportamento de sempre: responde
// qualquer chamado). Nunca vale pra master, que sempre tem acesso completo.
function getAdmins() {
  const sh = getOrCreateSheet('Admins', ['nome', 'senha', 'permissoes', 'editar', 'podeAbrirChamados', 'podeResponderChamados', 'responderSoProprios']);
  const data = sh.getDataRange().getValues();
  const rows = data.slice(1).filter(function (r) { return r[0] || r[1]; });
  if (rows.length === 0) {
    sh.appendRow(['Administrador', 'mude-esta-senha-123', 'todas', true, true, true, false]);
    return [{ nome: 'Administrador', senha: 'mude-esta-senha-123', permissoes: 'todas', editar: true, podeAbrirChamados: true, podeResponderChamados: true, responderSoProprios: false }];
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
      responderSoProprios: r[6] === true,
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
  const sh = getOrCreateSheet('Admins', ['nome', 'senha', 'permissoes', 'editar', 'podeAbrirChamados', 'podeResponderChamados', 'responderSoProprios']);
  const lastRow = sh.getLastRow();
  if (lastRow >= 2) {
    sh.getRange(2, 1, lastRow - 1, 7).clearContent();
  }
  // Se quem chamou (frontend atual, sempre manda os dois campos; uma
  // chamada externa/antiga pode não mandar) não informar
  // podeAbrirChamados/podeResponderChamados, herda de "editar" — mesma
  // regra de fallback usada pra linha em branco em getAdmins, só que aqui
  // pro objeto JS recebido em vez da célula da planilha.
  const rows = list.map(function (a) {
    const abrir = a.podeAbrirChamados === undefined ? !!a.editar : !!a.podeAbrirChamados;
    const responder = a.podeResponderChamados === undefined ? !!a.editar : !!a.podeResponderChamados;
    return [a.nome, a.senha, a.permissoes, !!a.editar, abrir, responder, !!a.responderSoProprios];
  });
  if (rows.length > 0) {
    sh.getRange(2, 1, rows.length, 7).setValues(rows);
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
// Coluna "email" (5ª): só preenchida por quem se cadastrou sozinho (ver
// cadastrarSolicitanteComTrava_ abaixo) — conta admin-created continuam sem
// email, e tudo bem. Serve como identificador alternativo de login (ver
// findSolicitante), além do nome.
// Coluna "aprovado" (6ª): true = pode entrar normalmente. false = cadastro
// pendente, só um admin aprovando em Usuários libera o login (ver
// autenticarUsuarioLogin/authenticateSolicitante). Em branco (inclusive
// linhas antigas de antes dessa coluna existir, e qualquer conta criada
// direto por um admin) = true, pra não exigir aprovação de quem já estava
// cadastrado ou foi criado pelo próprio admin.
// Coluna "autorizadoAbreChamados" (7ª): só importa junto com tipo='autorizado'
// — por padrão um "autorizado" é só visualização (Painel, Categorias, Salas,
// Inventário), sem chamado nenhum. Marcando essa opção, ele continua só
// visualização no resto, mas ganha a aba "Chamados" pra abrir e acompanhar
// os PRÓPRIOS chamados (igual um solicitante comum) — ver authenticateSolicitante
// e o ramo userNome de doGet. Em branco = false (comportamento de sempre).
function getSolicitantes() {
  const sh = getOrCreateSheet('Solicitantes', ['nome', 'senha', 'tipo', 'foto', 'email', 'aprovado', 'autorizadoAbreChamados']);
  const data = sh.getDataRange().getValues();
  return data.slice(1).filter(function (r) { return r[0]; }).map(function (r) {
    const aprovadoEmBranco = r[5] === '' || r[5] === undefined || r[5] === null;
    return {
      nome: String(r[0] || ''),
      senha: String(r[1] || ''),
      tipo: String(r[2] || 'solicitante') || 'solicitante',
      foto: String(r[3] || ''),
      email: String(r[4] || ''),
      aprovado: aprovadoEmBranco ? true : r[5] === true,
      autorizadoAbreChamados: r[6] === true,
    };
  });
}

// Busca por nome OU email (case-insensitive) — quem se cadastra sozinho
// entra com o email; quem foi cadastrado por um admin continua entrando com
// o nome, como sempre.
function findSolicitante(identificador) {
  const alvo = String(identificador || '').trim().toLowerCase();
  if (!alvo) return null;
  const lista = getSolicitantes();
  for (let i = 0; i < lista.length; i++) {
    const s = lista[i];
    if (s.nome.trim().toLowerCase() === alvo) return s;
    if (s.email && s.email.trim().toLowerCase() === alvo) return s;
  }
  return null;
}

function saveSolicitantes(list) {
  const sh = getOrCreateSheet('Solicitantes', ['nome', 'senha', 'tipo', 'foto', 'email', 'aprovado', 'autorizadoAbreChamados']);
  const lastRow = sh.getLastRow();
  if (lastRow >= 2) {
    sh.getRange(2, 1, lastRow - 1, 7).clearContent();
  }
  const rows = (list || []).map(function (s) {
    return [
      s.nome,
      s.senha,
      s.tipo === 'autorizado' ? 'autorizado' : 'solicitante',
      s.foto || '',
      s.email || '',
      s.aprovado === undefined ? true : !!s.aprovado,
      !!s.autorizadoAbreChamados,
    ];
  });
  if (rows.length > 0) {
    sh.getRange(2, 1, rows.length, 7).setValues(rows);
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
    return { nome: s.nome, senha: s.senha, tipo: s.tipo, foto: novaFoto || '', email: s.email, aprovado: s.aprovado, autorizadoAbreChamados: s.autorizadoAbreChamados };
  });
  saveSolicitantes(atualizada);
}

// Autentica só nome+senha, sem olhar o tipo — usado no LOGIN (ramo userNome
// de doGet), onde tanto solicitante quanto autorizado podem entrar. Devolve
// o registro mesmo se aprovado=false (quem chama decide o que fazer com
// isso — ver o ramo userNome em doGet, que dá uma mensagem específica pra
// cadastro pendente em vez de "senha incorreta").
function autenticarUsuarioLogin(nome, senha) {
  const s = findSolicitante(nome);
  if (!s) return null;
  if (s.senha !== String(senha || '')) return null;
  return s;
}

// Retorna o registro do solicitante (com o nome como está cadastrado) se
// nome+senha baterem E já estiver aprovado, ou null. Usado pra autenticar
// quem pode abrir chamado e responder — sem isso, doPost aceitava
// novoChamado/novaMensagem de qualquer um que soubesse a URL pública do
// backend, sem checar login algum. Um "autorizado" comum é só visualização:
// mesmo logado, não pode criar chamado nem responder — a não ser que tenha
// autorizadoAbreChamados marcado, aí ele pode abrir/acompanhar os PRÓPRIOS
// chamados como um solicitante normal, só que continua sem poder editar o
// resto (inventário, categorias etc.). Um cadastro ainda pendente de
// aprovação também não consegue, mesmo sabendo a senha certa — mas não
// deveria nem conseguir logar antes disso (ver doGet).
function authenticateSolicitante(nome, senha) {
  const s = autenticarUsuarioLogin(nome, senha);
  if (!s || s.aprovado === false) return null;
  if (s.tipo === 'autorizado' && !s.autorizadoAbreChamados) return null;
  return s;
}

// ---------- Cadastro público de solicitante ----------
//
// Antes, só um admin podia criar acesso de solicitante (em Usuários). Agora
// qualquer pessoa pode se cadastrar sozinha direto na tela de login (nome +
// email + senha escolhida por ela) — mas a conta nasce com aprovado=false e
// não consegue entrar até um administrador aprovar em Usuários. Roda dentro
// do doGet (não do doPost) porque o cadastro precisa de uma resposta de
// verdade pro navegador saber se o email já existe ou não (doPost, por usar
// fetch em modo no-cors pra evitar problema de CORS do Apps Script, nunca
// consegue ler a resposta — só sabe se a requisição saiu, não o que o
// servidor respondeu). Isso só é seguro por causa do LockService em
// cadastrarSolicitanteComTrava_: sem a trava, duas pessoas se cadastrando
// com o mesmo email ao mesmo tempo poderiam duplicar a conta — o mesmo
// problema de concorrência que o doPost já resolve pras outras escritas.
function validarEmail_(email) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email);
}

function processarCadastroSolicitante_(nome, email, senha) {
  nome = String(nome || '').trim();
  email = String(email || '').trim().toLowerCase();
  senha = String(senha || '');
  if (!nome || !email || !senha) {
    return { ok: false, error: 'Preencha nome, email e senha.' };
  }
  if (!validarEmail_(email)) {
    return { ok: false, error: 'Informe um email válido.' };
  }
  const existentes = getSolicitantes();
  const jaExisteEmail = existentes.some(function (s) { return s.email && s.email.trim().toLowerCase() === email; });
  if (jaExisteEmail) {
    return { ok: false, error: 'Já existe um cadastro com esse email.' };
  }
  const jaExisteNome = existentes.some(function (s) { return s.nome.trim().toLowerCase() === nome.toLowerCase(); });
  if (jaExisteNome) {
    return { ok: false, error: 'Já existe um cadastro com esse nome.' };
  }
  const novo = { nome: nome, senha: senha, tipo: 'solicitante', foto: '', email: email, aprovado: false };
  saveSolicitantes(existentes.concat([novo]));
  return { ok: true };
}

function cadastrarSolicitanteComTrava_(nome, email, senha) {
  const lock = LockService.getScriptLock();
  let temLock = false;
  try {
    temLock = lock.tryLock(30000);
  } catch (err) {
    temLock = false;
  }
  if (!temLock) {
    return { ok: false, error: 'Sistema ocupado, tente novamente em alguns segundos.' };
  }
  try {
    return processarCadastroSolicitante_(nome, email, senha);
  } finally {
    lock.releaseLock();
  }
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
      const cham = listarChamados(); // só busca aqui (1x por mês), não em toda chamada
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

// ---------- Chamados (linha própria por chamado, não mais dentro do AppState) ----------

// Antes, todo chamado vivia dentro do mesmo bloco JSON gigante que guarda
// inventário/categorias/salas/responsáveis (a aba "AppState") — então
// qualquer ação de chamado, até só mudar um status, precisava ler e
// reescrever esse bloco inteiro. Com o tempo, conforme o inventário cresce,
// isso fica mais lento — e com muita gente mexendo em chamado ao mesmo
// tempo (a trava do LockService serializa tudo), quanto mais tempo cada
// ação leva, mais tempo as próximas pessoas da fila esperam. Agora cada
// chamado é uma linha própria na aba "Chamados" (coluna A = id, pra achar
// rápido; coluna B = o chamado inteiro em JSON) — mudar um chamado só
// lê/escreve a linha dele, não o resto do app. A trava continua protegendo
// exatamente do mesmo jeito, só que fica ocupada por menos tempo em cada
// ação.
function encontrarLinhaChamado_(sh, chamadoId) {
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return -1;
  const ids = sh.getRange(2, 1, lastRow - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === String(chamadoId)) return i + 2; // +2: pula o cabeçalho, base 1
  }
  return -1;
}

function listarChamados() {
  const sh = getOrCreateSheet('Chamados', ['id', 'dados']);
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return [];
  const rows = sh.getRange(2, 1, lastRow - 1, 2).getValues();
  const out = [];
  for (let i = 0; i < rows.length; i++) {
    if (!rows[i][0]) continue;
    try {
      out.push(JSON.parse(rows[i][1]));
    } catch (e) {
      // linha corrompida (não deveria acontecer) — ignora em vez de quebrar a lista inteira
    }
  }
  return out;
}

function buscarChamadoPorId_(chamadoId) {
  const sh = getOrCreateSheet('Chamados', ['id', 'dados']);
  const linha = encontrarLinhaChamado_(sh, chamadoId);
  if (linha === -1) return null;
  const valor = sh.getRange(linha, 2, 1, 1).getValues()[0][0];
  try {
    return JSON.parse(valor);
  } catch (e) {
    return null;
  }
}

function salvarChamado_(chamado) {
  const sh = getOrCreateSheet('Chamados', ['id', 'dados']);
  const linha = encontrarLinhaChamado_(sh, chamado.id);
  const json = JSON.stringify(chamado);
  if (linha === -1) {
    sh.appendRow([chamado.id, json]);
  } else {
    sh.getRange(linha, 1, 1, 2).setValues([[chamado.id, json]]);
  }
}

function excluirChamadoDaPlanilha_(chamadoId) {
  const sh = getOrCreateSheet('Chamados', ['id', 'dados']);
  const linha = encontrarLinhaChamado_(sh, chamadoId);
  if (linha !== -1) sh.deleteRow(linha);
}

// Execução manual (uma vez só, pelo editor do Apps Script) pra mover os
// chamados que hoje estão dentro do bloco da AppState pra essa aba própria.
// Lê o bloco antigo diretamente (sem passar por readState(), que já não
// devolve mais o campo chamados) pra garantir que pega os dados de antes da
// migração mesmo depois do código novo estar publicado.
function migrarChamadosParaAbaPropria() {
  const sh = getOrCreateSheet('AppState', ['data']);
  const lastRow = sh.getLastRow();
  if (lastRow < 2) {
    Logger.log('Nada pra migrar: AppState vazio.');
    return;
  }
  const values = sh.getRange(2, 1, lastRow - 1, 1).getValues();
  const combined = values.map(function (r) { return r[0]; }).join('');
  let antigo;
  try {
    antigo = JSON.parse(combined);
  } catch (e) {
    Logger.log('Não consegui ler o AppState pra migrar: ' + e);
    return;
  }
  const chamados = antigo.chamados || [];
  chamados.forEach(function (c) { salvarChamado_(c); });
  Logger.log('Migração concluída: ' + chamados.length + ' chamado(s) movido(s) pra aba própria "Chamados".');
}

// ---------- Firebase / Firestore (sincronização ao vivo dos Chamados) ----------

// A planilha (AppState/readState/writeState) continua sendo a fonte da
// verdade dos chamados — nada muda em como eles são abertos, respondidos
// ou protegidos pela trava. A única coisa nova: depois de cada gravação
// bem-sucedida na planilha, o mesmo chamado é replicado pro Firestore, que
// é o banco que a tela do admin escuta em tempo real (sem precisar
// recarregar a página pra ver um chamado novo). Se o Firestore ainda não
// foi configurado (faltam as 3 propriedades do script) ou se a chamada
// falhar por qualquer motivo, a sincronização é simplesmente pulada — ela
// nunca pode impedir a gravação principal na planilha, que é o que
// realmente importa.
function getFirestore_() {
  const props = PropertiesService.getScriptProperties();
  const email = props.getProperty('FIREBASE_CLIENT_EMAIL');
  const chavePrivada = props.getProperty('FIREBASE_PRIVATE_KEY');
  const projectId = props.getProperty('FIREBASE_PROJECT_ID');
  if (!email || !chavePrivada || !projectId) return null;
  // O valor colado no Script Properties tem "\n" literais (duas letras,
  // barra e "n") em vez de quebra de linha de verdade — a biblioteca
  // precisa da chave no formato PEM real.
  const chaveFormatada = chavePrivada.replace(/\\n/g, '\n');
  return FirestoreApp.getFirestore(email, chaveFormatada, projectId);
}

function sincronizarChamadoNoFirestore_(chamado) {
  try {
    const db = getFirestore_();
    if (!db || !chamado || !chamado.id) return;
    db.updateDocument('chamados/' + chamado.id, chamado, true); // true = cria o documento se ainda não existir
  } catch (err) {
    Logger.log('Falha ao sincronizar chamado ' + (chamado && chamado.id) + ' com o Firestore: ' + err);
  }
}

function excluirChamadoNoFirestore_(chamadoId) {
  try {
    const db = getFirestore_();
    if (!db || !chamadoId) return;
    db.deleteDocument('chamados/' + chamadoId);
  } catch (err) {
    Logger.log('Falha ao excluir chamado ' + chamadoId + ' do Firestore: ' + err);
  }
}

// Execução manual (uma vez só, pelo editor do Apps Script) pra copiar os
// chamados que já existem na planilha pro Firestore, na hora de ligar a
// sincronização — sem isso, só chamados novos apareceriam lá.
function migrarChamadosParaFirestore() {
  const chamados = listarChamados();
  let ok = 0;
  chamados.forEach(function (c) {
    sincronizarChamadoNoFirestore_(c);
    ok++;
  });
  Logger.log('Migração concluída: ' + ok + ' chamado(s) copiado(s) pro Firestore.');
}

// ---------- Estado do app (inventário, categorias, etc.) ----------

// Não inclui mais "chamados" — eles vivem na própria aba "Chamados" (ver
// listarChamados/salvarChamado_/etc acima). Se o bloco salvo aqui ainda
// tiver um campo "chamados" de antes da migração, ele é descartado: quem
// precisar dos chamados usa listarChamados().
function readState() {
  const sh = getOrCreateSheet('AppState', ['data']);
  const lastRow = sh.getLastRow();
  const defaults = { categorias: [], areas: [], responsaveis: [], inventario: [] };
  if (lastRow < 2) return defaults;
  const values = sh.getRange(2, 1, lastRow - 1, 1).getValues();
  const combined = values.map(function (r) { return r[0]; }).join('');
  if (!combined) return defaults;
  try {
    const parsed = JSON.parse(combined);
    delete parsed.chamados;
    return parsed;
  } catch (e) {
    return defaults;
  }
}

// Por segurança, remove "chamados" antes de gravar mesmo que alguém passe
// por engano — chamados nunca devem ser persistidos aqui.
function writeState(obj) {
  const sh = getOrCreateSheet('AppState', ['data']);
  const paraGravar = Object.assign({}, obj);
  delete paraGravar.chamados;
  const json = JSON.stringify(paraGravar);
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
// "chamados" nunca passa por aqui, nem pro master: chamados têm ações
// próprias (novoChamado, novaMensagem, mudarStatusChamado, excluirChamado)
// e aba própria (ver seção "Chamados" acima), que sempre leem/escrevem o
// registro fresco na hora, em vez de confiar na cópia que o navegador de
// quem está logado guardou no login. Antes disso, um admin com a aba aberta
// há um tempo, ao salvar qualquer outra coisa (mesmo sem nada a ver com
// chamados), sobrescrevia TODOS os chamados com essa cópia antiga —
// comprovado em teste, 20 chamados novos apagados de uma vez só. Isso vale
// até pro master, que antes tinha as escritas aceitas sem filtro nenhum.
function filtrarEstadoPorPermissao(novoEstado, admin) {
  const isMaster = admin && admin.permissoes === 'todas';
  const secoesPermitidas = isMaster ? null : String((admin && admin.permissoes) || '').split(',').map(function (s) { return s.trim().toLowerCase(); });
  const atual = readState();
  const resultado = {};
  for (const chave in novoEstado) {
    if (chave === 'chamados') continue; // nunca passa por aqui — ver comentário acima
    const secao = ESTADO_CAMPO_PARA_SECAO[chave];
    const podeEscreverSecao = isMaster || !!(admin && admin.editar);
    const secaoPermitida = isMaster || (secao && secoesPermitidas.indexOf(secao) !== -1);
    if (secaoPermitida && podeEscreverSecao) {
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

  if (p.action === 'cadastro') {
    payload = cadastrarSolicitanteComTrava_(p.novoNome, p.novoEmail, p.novoSenha);
    if (callback) {
      return ContentService
        .createTextOutput(callback + '(' + JSON.stringify(payload) + ');')
        .setMimeType(ContentService.MimeType.JAVASCRIPT);
    }
    return jsonOut(payload);
  }

  const secret = p.secret || '';
  const userNome = p.userNome || '';
  const userSenha = p.userSenha || '';

  const admins = secret ? getAdmins() : null;
  const admin = findAdminBySecret(secret, admins);

  if (admin) {
    const isMaster = admin.permissoes === 'todas';
    const state = readState();
    // getSolicitantes() é lido pra qualquer admin autenticado (não só
    // master) porque a foto de perfil de quem abriu um chamado precisa
    // aparecer pro admin de Chamados também — mas só a foto vai pra quem
    // não é master (fotosSolicitantes), nunca a lista com senha
    // (solicitantes), que continua exclusiva do master.
    const solicitantesList = getSolicitantes();
    const fotosSolicitantes = {};
    solicitantesList.forEach(function (s) {
      if (s.foto) fotosSolicitantes[s.nome] = s.foto;
    });
    payload = {
      ok: true,
      isAdmin: true,
      nome: admin.nome,
      permissoes: admin.permissoes,
      editar: !!admin.editar,
      podeAbrirChamados: !!admin.podeAbrirChamados,
      podeResponderChamados: !!admin.podeResponderChamados,
      responderSoProprios: !!admin.responderSoProprios,
      state: Object.assign({}, state, { chamados: listarChamados() }),
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
          responderSoProprios: !!a.responderSoProprios,
        };
      }) : [],
      solicitantes: isMaster ? solicitantesList : [],
      fotosSolicitantes: fotosSolicitantes,
      historico: getHistoricoMensal()
    };
  } else if (userNome) {
    const usuario = autenticarUsuarioLogin(userNome, userSenha);
    if (usuario && usuario.aprovado === false) {
      payload = { ok: true, isAdmin: false, isUser: false, error: 'Seu cadastro ainda está aguardando aprovação de um administrador.' };
    } else if (usuario && usuario.tipo === 'autorizado') {
      // Autorizado: entra como um usuário comum (não como admin), com
      // visão só-leitura de Painel/Categorias/Salas/Inventário — sem
      // admins/solicitantes/histórico. Por padrão também sem chamado nenhum;
      // se autorizadoAbreChamados estiver marcado, ganha os PRÓPRIOS
      // chamados (nunca os de outras pessoas), do mesmo jeito que um
      // solicitante comum — ver o "alvo"/filtro logo abaixo, no ramo normal.
      const state = readState();
      const podeAbrirChamados = !!usuario.autorizadoAbreChamados;
      const alvoAutorizado = usuario.nome.trim().toLowerCase();
      const chamadosDoAutorizado = podeAbrirChamados
        ? listarChamados().filter(function (c) { return (c.criadoPor || '').trim().toLowerCase() === alvoAutorizado; })
        : [];
      payload = {
        ok: true,
        isAdmin: false,
        isUser: true,
        isAutorizado: true,
        podeAbrirChamados: podeAbrirChamados,
        nome: usuario.nome,
        foto: usuario.foto || '',
        state: {
          categorias: state.categorias || [],
          areas: state.areas || [],
          responsaveis: state.responsaveis || [],
          inventario: state.inventario || [],
          chamados: chamadosDoAutorizado
        }
      };
    } else if (usuario) {
      const state = readState();
      // Usa usuario.nome (o nome de verdade, resolvido pelo login) em vez do
      // que a pessoa digitou em userNome — desde que o login também aceita
      // email (ver findSolicitante), comparar direto com o que foi digitado
      // deixaria "meus chamados" vazio pra quem entra pelo email, já que
      // criadoPor sempre guarda o nome, nunca o email.
      const alvo = usuario.nome.trim().toLowerCase();
      const meusChamados = listarChamados().filter(function (c) {
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

// Toda ação de doPost segue o padrão "lê o estado inteiro -> muda um pedaço
// -> escreve o estado inteiro de volta" (readState/writeState). Sem
// nenhuma trava, duas pessoas mandando uma ação quase ao mesmo tempo (a
// diferença de tempo entre a leitura de uma e a escrita da outra é só o
// tempo de ida e volta até a planilha) fazem a segunda escrita apagar
// silenciosamente o que a primeira tinha acabado de salvar — comprovado
// num teste de concorrência: dois chamados abertos "ao mesmo tempo"
// viraram um só, sem erro nenhum. Com poucas pessoas isso é raro; com
// muita gente usando ao mesmo tempo deixa de ser raro. O LockService
// serializa as execuções deste script (cada uma espera a sua vez em vez
// de rodar em paralelo) só durante a ação em si — o resto do app
// continua funcionando normalmente, só essa gravação específica espera.
function doPost(e) {
  const lock = LockService.getScriptLock();
  let temLock = false;
  try {
    temLock = lock.tryLock(30000);
  } catch (err) {
    temLock = false;
  }
  if (!temLock) {
    // Alguém mais ficou muito tempo segurando a trava (uso pesado
    // simultâneo). Melhor avisar e deixar tentar de novo do que travar a
    // execução até estourar o limite de tempo do Apps Script.
    return jsonOut({ ok: false, error: 'Sistema ocupado, tente novamente em alguns segundos.' });
  }
  try {
    return doPostComTrava(e);
  } finally {
    lock.releaseLock();
  }
}

function doPostComTrava(e) {
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
  // Restringe um admin (nunca o master) a só responder/mudar status/excluir
  // os chamados que ELE MESMO abriu (campo "abertoPorAdmin", marcado em
  // novoChamado abaixo) — pra alguém tipo recepção, que abre chamado pra
  // outras pessoas, poder acompanhar os próprios sem mexer nas conversas de
  // outros admins ou dos solicitantes que abriram direto.
  const adminResponderSoProprios = isAdmin && !isMaster && !!admin.responderSoProprios;
  function podeMexerNesseChamado_(chamado) {
    if (!adminResponderSoProprios) return true;
    return chamado && chamado.abertoPorAdmin === admin.nome;
  }

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
    if (solicitante) {
      body.chamado.criadoPor = solicitante.nome;
      body.chamado.solicitante = solicitante.nome;
    } else if (isAdmin) {
      // Marca quem abriu quando é um admin abrindo direto (ex: recepção
      // atendendo alguém pessoalmente) — é o que permite restringir esse
      // admin a só responder os próprios chamados (responderSoProprios).
      body.chamado.abertoPorAdmin = admin.nome;
    }
    salvarChamado_(body.chamado);
    notificarNovoChamado(body.chamado);
    sincronizarChamadoNoFirestore_(body.chamado);
    return jsonOut({ ok: true });
  }

  if (body.action === 'novaMensagem') {
    const solicitante = authenticateSolicitante(body.userNome, body.userSenha);
    if (!adminPodeResponderChamados && !solicitante) {
      return jsonOut({ ok: false, error: 'Não autenticado' });
    }
    const chamado = buscarChamadoPorId_(body.chamadoId);
    if (!chamado) {
      return jsonOut({ ok: false, error: 'Chamado não encontrado' });
    }
    if (solicitante && !adminPodeResponderChamados) {
      const dono = String(chamado.criadoPor || chamado.solicitante || '').trim().toLowerCase();
      if (dono !== solicitante.nome.trim().toLowerCase()) {
        return jsonOut({ ok: false, error: 'Sem permissão para responder este chamado' });
      }
    }
    if (isAdmin && !podeMexerNesseChamado_(chamado)) {
      return jsonOut({ ok: false, error: 'Você só pode responder aos chamados que você mesma abriu' });
    }
    const mensagem = body.mensagem || {};
    mensagem.autor = adminPodeResponderChamados ? 'ti' : 'solicitante';
    // Nome de quem respondeu de verdade (nunca o que o cliente mandou) —
    // antes toda resposta de admin aparecia só como "Administrador" pra
    // todo mundo, sem dar pra saber QUAL admin respondeu.
    mensagem.nome = adminPodeResponderChamados ? admin.nome : solicitante.nome;
    chamado.mensagens.push(mensagem);
    salvarChamado_(chamado);
    sincronizarChamadoNoFirestore_(chamado);
    return jsonOut({ ok: true });
  }

  // mudarStatusChamado/excluirChamado: antes, essas duas ações do admin
  // (mudar status, excluir chamado) só mexiam no "state" local do
  // navegador e dependiam do autosave geral (salvarTudo) pra persistir —
  // o mesmo autosave que manda o ESTADO INTEIRO de volta, inclusive a
  // cópia de chamados de quando o admin abriu a página. Um admin com a
  // aba aberta por um tempo, ao salvar qualquer outra coisa (ex: editar um
  // equipamento), acabava sobrescrevendo TODOS os chamados com essa cópia
  // antiga — comprovado em teste, 20 chamados novos apagados de uma vez.
  // Agora essas duas ações leem o registro fresco na hora, igual
  // novoChamado/novaMensagem já faziam, e nunca dependem do autosave geral.
  if (body.action === 'mudarStatusChamado') {
    if (!adminPodeResponderChamados) {
      return jsonOut({ ok: false, error: 'Não autenticado' });
    }
    const chamado = buscarChamadoPorId_(body.chamadoId);
    if (!chamado) {
      return jsonOut({ ok: false, error: 'Chamado não encontrado' });
    }
    if (!podeMexerNesseChamado_(chamado)) {
      return jsonOut({ ok: false, error: 'Você só pode mudar o status dos chamados que você mesma abriu' });
    }
    chamado.status = body.status;
    salvarChamado_(chamado);
    sincronizarChamadoNoFirestore_(chamado);
    return jsonOut({ ok: true });
  }

  if (body.action === 'excluirChamado') {
    if (!adminPodeResponderChamados) {
      return jsonOut({ ok: false, error: 'Não autenticado' });
    }
    if (adminResponderSoProprios) {
      const chamado = buscarChamadoPorId_(body.chamadoId);
      if (!podeMexerNesseChamado_(chamado)) {
        return jsonOut({ ok: false, error: 'Você só pode excluir os chamados que você mesma abriu' });
      }
    }
    excluirChamadoDaPlanilha_(body.chamadoId);
    excluirChamadoNoFirestore_(body.chamadoId);
    return jsonOut({ ok: true });
  }

  return jsonOut({ ok: false, error: 'Ação não autorizada' });
}
