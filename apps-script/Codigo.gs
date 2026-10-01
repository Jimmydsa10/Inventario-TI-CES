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
  // Antes: com a aba vazia, recriava sozinho um admin master com a senha
  // padrão 'mude-esta-senha-123' — que está no código público do GitHub, ou
  // seja, qualquer um conseguiria entrar como master depois de um acidente
  // na planilha. Agora aba vazia = ninguém entra como admin; o primeiro admin
  // precisa ser digitado direto na planilha por quem tem acesso a ela.
  // Linha sem senha também é ignorada, pra um secret vazio nunca bater.
  return rows.filter(function (r) { return String(r[1] || '') !== ''; }).map(function (r) {
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
    if (senhaConfere_(secret, admins[i].senha)) return admins[i];
  }
  return null;
}

function findAdminPorNome_(nome, adminsList) {
  const alvo = String(nome || '').trim().toLowerCase();
  if (!alvo) return null;
  const admins = adminsList || getAdmins();
  for (let i = 0; i < admins.length; i++) {
    if (admins[i].nome.trim().toLowerCase() === alvo) return admins[i];
  }
  return null;
}

// ---------- Senhas guardadas com hash ----------

// Antes as senhas ficavam em texto puro na planilha (e, agora, também nos
// backups diários): quem visse a planilha ou um backup via a senha de todo
// mundo — e muita gente repete a mesma senha em outros lugares. Agora a
// planilha guarda "h1$<repetições>$<sal>$<hash>": SHA-256 com um sal
// aleatório por senha, repetido HASH_SENHA_ITERACOES vezes (pra deixar
// tentativa e erro offline caro). Não dá pra voltar do hash pra senha.
//
// Senha em texto puro ainda é aceita — é assim que o primeiro admin é
// digitado direto na planilha — e é trocada pelo hash no primeiro login
// certo (ver migrarSenhaSeTextoPuro_). Pra converter todas de uma vez, rode
// migrarSenhasParaHash pelo editor.
const HASH_SENHA_ITERACOES = 1000;

function bytesParaHex_(bytes) {
  return bytes.map(function (b) { return ('0' + (b & 0xff).toString(16)).slice(-2); }).join('');
}

function hashHex_(texto) {
  return bytesParaHex_(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, texto, Utilities.Charset.UTF_8));
}

// 64 caracteres hex a partir de dois UUIDs aleatórios (~244 bits) — o
// hash só garante o formato, a aleatoriedade vem dos UUIDs.
function hexAleatorio_() {
  return hashHex_(Utilities.getUuid() + ':' + Utilities.getUuid());
}

function derivarHashSenha_(senha, sal, iteracoes) {
  let atual = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, sal + ':' + senha, Utilities.Charset.UTF_8);
  for (let i = 1; i < iteracoes; i++) {
    atual = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, atual);
  }
  return bytesParaHex_(atual);
}

function ehHashDeSenha_(valor) {
  return /^h1\$\d{1,5}\$[0-9a-f]{32}\$[0-9a-f]{64}$/.test(String(valor || ''));
}

function gerarHashSenha_(senha) {
  const sal = hexAleatorio_().slice(0, 32);
  return ['h1', HASH_SENHA_ITERACOES, sal, derivarHashSenha_(String(senha), sal, HASH_SENHA_ITERACOES)].join('$');
}

// Compara sem parar no primeiro caractere diferente (o tempo de resposta
// não entrega quantos caracteres bateram).
function textosIguais_(a, b) {
  a = String(a); b = String(b);
  if (a.length !== b.length) return false;
  let diferenca = 0;
  for (let i = 0; i < a.length; i++) diferenca |= a.charCodeAt(i) ^ b.charCodeAt(i);
  return diferenca === 0;
}

function senhaConfere_(digitada, armazenada) {
  digitada = String(digitada || '');
  armazenada = String(armazenada || '');
  if (!digitada || !armazenada) return false;
  if (!ehHashDeSenha_(armazenada)) return textosIguais_(digitada, armazenada);
  const partes = armazenada.split('$');
  return textosIguais_(derivarHashSenha_(digitada, partes[2], Number(partes[1])), partes[3]);
}

// O que vai pra planilha: hash continua como está (é a senha de sempre,
// que o frontend do master devolve sem mudar quando o campo "nova senha"
// fica em branco); qualquer outra coisa é uma senha nova digitada e vira
// hash.
function senhaParaGravar_(valor) {
  valor = String(valor === undefined || valor === null ? '' : valor);
  if (!valor || ehHashDeSenha_(valor)) return valor;
  return gerarHashSenha_(valor);
}

// Troca por hash a senha em texto puro de UMA pessoa, logo depois de um
// login certo. Devolve o valor que ficou gravado (o hash novo, ou o de
// antes se não precisou/conseguiu trocar). Roda no doGet, fora da trava do
// doPost — por isso pega a trava aqui, sem esperar muito (se não der, fica
// pro próximo login).
function migrarSenhaSeTextoPuro_(nomeAba, nome, senhaDigitada, armazenada) {
  if (ehHashDeSenha_(armazenada)) return armazenada;
  const lock = LockService.getScriptLock();
  let temLock = false;
  try {
    temLock = lock.tryLock(5000);
  } catch (err) {
    temLock = false;
  }
  if (!temLock) return armazenada;
  try {
    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nomeAba);
    if (!sh || sh.getLastRow() < 2) return armazenada;
    const linhas = sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues();
    const alvo = String(nome || '').trim().toLowerCase();
    for (let i = 0; i < linhas.length; i++) {
      if (String(linhas[i][0] || '').trim().toLowerCase() === alvo && String(linhas[i][1] || '') === armazenada) {
        const novo = gerarHashSenha_(senhaDigitada);
        sh.getRange(i + 2, 2).setValue(novo);
        return novo;
      }
    }
    return armazenada;
  } finally {
    lock.releaseLock();
  }
}

// Execução manual (uma vez só, pelo editor do Apps Script): troca por hash
// todas as senhas que ainda estão em texto puro em Admins e Solicitantes.
function migrarSenhasParaHash() {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    ['Admins', 'Solicitantes'].forEach(function (nomeAba) {
      const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nomeAba);
      if (!sh || sh.getLastRow() < 2) return;
      const faixa = sh.getRange(2, 2, sh.getLastRow() - 1, 1);
      const senhas = faixa.getValues();
      let trocadas = 0;
      senhas.forEach(function (linha) {
        const valor = String(linha[0] || '');
        if (valor && !ehHashDeSenha_(valor)) {
          linha[0] = gerarHashSenha_(valor);
          trocadas++;
        }
      });
      if (trocadas) faixa.setValues(senhas);
      Logger.log(nomeAba + ': ' + trocadas + ' senha(s) convertida(s) pra hash.');
    });
  } finally {
    lock.releaseLock();
  }
}

// ---------- Sessões ----------

// Antes, pra continuar logado depois de um F5, o navegador guardava a
// PRÓPRIA SENHA no localStorage e mandava ela em toda requisição. Qualquer
// coisa que lesse o localStorage (uma extensão, alguém no computador
// compartilhado da sala) levava a senha junto. Agora o login devolve um
// token de sessão aleatório (sem nenhuma relação com a senha), que é o
// que o navegador guarda e manda. O servidor guarda só o hash do token nas
// Propriedades do script, com: tipo (admin/usuário), nome, validade e uma
// "impressão" da senha atual — trocar a senha de alguém derruba na hora
// todas as sessões dessa pessoa, e excluir/renomear também. Sair do app
// apaga a sessão no servidor.
const SESSAO_PREFIXO_PROP = 'sessao_';
const SESSAO_DURACAO_MS = 30 * 24 * 60 * 60 * 1000;
const SESSAO_RENOVAR_MS = 5 * 24 * 60 * 60 * 1000;

function chaveSessao_(token) {
  return SESSAO_PREFIXO_PROP + hashHex_('sessao:' + token).slice(0, 40);
}

function impressaoDaSenha_(armazenada) {
  return hashHex_('senha:' + String(armazenada || '')).slice(0, 16);
}

function limparSessoesVencidas_(props) {
  const todas = props.getProperties();
  const agora = Date.now();
  Object.keys(todas).forEach(function (chave) {
    if (chave.indexOf(SESSAO_PREFIXO_PROP) !== 0) return;
    let vencida = true;
    try {
      vencida = !(JSON.parse(todas[chave]).e > agora);
    } catch (e) {}
    if (vencida) props.deleteProperty(chave);
  });
}

// tipo: 'a' (admin) ou 'u' (solicitante/autorizado).
function criarSessao_(tipo, nome, senhaArmazenada) {
  const token = hexAleatorio_();
  const props = PropertiesService.getScriptProperties();
  limparSessoesVencidas_(props);
  props.setProperty(chaveSessao_(token), JSON.stringify({
    t: tipo,
    n: nome,
    f: impressaoDaSenha_(senhaArmazenada),
    e: Date.now() + SESSAO_DURACAO_MS,
  }));
  return token;
}

// Devolve { admin, admins } ou { usuario } — ou null se o token não vale
// (não existe, venceu, a pessoa sumiu, a senha mudou, o cadastro não está
// mais aprovado). A validade é renovada com o uso: quem abre o app pelo
// menos uma vez por mês não precisa entrar de novo.
function resolverSessao_(token) {
  if (!/^[0-9a-f]{64}$/.test(String(token || ''))) return null;
  const props = PropertiesService.getScriptProperties();
  const chave = chaveSessao_(token);
  const bruto = props.getProperty(chave);
  if (!bruto) return null;
  let sessao;
  try {
    sessao = JSON.parse(bruto);
  } catch (e) {
    sessao = null;
  }
  let resultado = null;
  if (sessao && sessao.e > Date.now()) {
    if (sessao.t === 'a') {
      const admins = getAdmins();
      const admin = findAdminPorNome_(sessao.n, admins);
      if (admin && impressaoDaSenha_(admin.senha) === sessao.f) resultado = { admin: admin, admins: admins };
    } else if (sessao.t === 'u') {
      const alvo = String(sessao.n || '').trim().toLowerCase();
      const usuario = getSolicitantes().filter(function (s) { return s.nome.trim().toLowerCase() === alvo; })[0];
      if (usuario && usuario.aprovado !== false && impressaoDaSenha_(usuario.senha) === sessao.f) resultado = { usuario: usuario };
    }
  }
  if (!resultado) {
    props.deleteProperty(chave);
    return null;
  }
  if (sessao.e - Date.now() < SESSAO_DURACAO_MS - SESSAO_RENOVAR_MS) {
    sessao.e = Date.now() + SESSAO_DURACAO_MS;
    props.setProperty(chave, JSON.stringify(sessao));
  }
  return resultado;
}

function encerrarSessao_(token) {
  if (!/^[0-9a-f]{64}$/.test(String(token || ''))) return;
  PropertiesService.getScriptProperties().deleteProperty(chaveSessao_(token));
}

// ---------- Gravação segura de abas inteiras ----------

// Antes, saveAdmins/saveSolicitantes APAGAVAM a aba primeiro e só depois
// gravavam a lista nova. Se a gravação falhasse no meio — o caso real era
// uma célula passando do limite de 50.000 caracteres do Sheets, por
// exemplo uma "foto de perfil" gigante mandada por fora do app —, a
// exceção interrompia o script com a aba JÁ vazia: todas as contas sumiam
// (reproduzido em teste). Agora: (1) confere o tamanho de todas as células
// antes de encostar na planilha — se alguma passar, nada é alterado; (2)
// grava a lista nova POR CIMA das linhas atuais; (3) só então limpa as
// linhas que sobraram embaixo (quando a lista nova é menor). Em nenhum
// momento a aba fica vazia.
const LIMITE_CARACTERES_CELULA_SEGURO = 45000;

// ---------- Proteção contra fórmulas na planilha ----------

// Texto gravado numa célula comum que começa com "=" vira FÓRMULA no Sheets.
// Como o cadastro é público, alguém podia se cadastrar com nome "=B2" (e a
// célula passava a mostrar a senha de outra pessoa) ou "=IMPORTDATA(...)"
// (tentando mandar dados da planilha pra fora). Duas camadas: (1) as colunas
// de texto de Admins/Solicitantes ficam formatadas como texto puro, o que
// também preserva senhas como "0123" (antes viravam o número 123); (2) o
// cadastro público recusa nome/email/senha começando com = + - @.
// Só as colunas de TEXTO — as de verdadeiro/falso (editar, aprovado...)
// não podem virar texto, senão "FALSE" deixaria de ser o booleano false.
const COLUNAS_TEXTO_ADMINS = 3;       // nome, senha, permissoes
const COLUNAS_TEXTO_SOLICITANTES = 5; // nome, senha, tipo, foto, email

function garantirColunasTexto_(sh, numColunas) {
  sh.getRange(1, 1, sh.getMaxRows(), numColunas).setNumberFormat('@');
}

function comecaComCaracterDeFormula_(valor) {
  return /^[=+\-@]/.test(String(valor || ''));
}

// Execução manual (pelo editor do Apps Script): lista no log qualquer
// célula de Admins/Solicitantes que tenha virado fórmula — pra conferir se
// alguém já tinha explorado isso antes da correção.
function verificarFormulasNaPlanilha() {
  let achou = 0;
  ['Admins', 'Solicitantes'].forEach(function (nome) {
    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nome);
    if (!sh) return;
    const formulas = sh.getDataRange().getFormulas();
    formulas.forEach(function (linha, i) {
      linha.forEach(function (f, j) {
        if (f) {
          achou++;
          Logger.log('Fórmula em ' + nome + ' linha ' + (i + 1) + ', coluna ' + (j + 1) + ': ' + f);
        }
      });
    });
  });
  Logger.log(achou ? achou + ' fórmula(s) encontrada(s) — confira as linhas acima.' : 'Nenhuma fórmula encontrada em Admins/Solicitantes.');
}

function substituirLinhasComSeguranca_(sh, rows, numColunas) {
  rows.forEach(function (row) {
    row.forEach(function (valor) {
      if (String(valor === undefined || valor === null ? '' : valor).length > LIMITE_CARACTERES_CELULA_SEGURO) {
        throw new Error('CELULA_MUITO_GRANDE');
      }
    });
  });
  const lastRowAntes = sh.getLastRow();
  if (rows.length > 0) {
    sh.getRange(2, 1, rows.length, numColunas).setValues(rows);
  }
  const primeiraSobrando = rows.length + 2;
  if (lastRowAntes >= primeiraSobrando) {
    sh.getRange(primeiraSobrando, 1, lastRowAntes - primeiraSobrando + 1, numColunas).clearContent();
  }
}

function saveAdmins(list) {
  const sh = getOrCreateSheet('Admins', ['nome', 'senha', 'permissoes', 'editar', 'podeAbrirChamados', 'podeResponderChamados', 'responderSoProprios']);
  // Se quem chamou (frontend atual, sempre manda os dois campos; uma
  // chamada externa/antiga pode não mandar) não informar
  // podeAbrirChamados/podeResponderChamados, herda de "editar" — mesma
  // regra de fallback usada pra linha em branco em getAdmins, só que aqui
  // pro objeto JS recebido em vez da célula da planilha.
  const rows = list.map(function (a) {
    const abrir = a.podeAbrirChamados === undefined ? !!a.editar : !!a.podeAbrirChamados;
    const responder = a.podeResponderChamados === undefined ? !!a.editar : !!a.podeResponderChamados;
    return [a.nome, senhaParaGravar_(a.senha), a.permissoes, !!a.editar, abrir, responder, !!a.responderSoProprios];
  });
  garantirColunasTexto_(sh, COLUNAS_TEXTO_ADMINS);
  substituirLinhasComSeguranca_(sh, rows, 7);
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
  const rows = (list || []).map(function (s) {
    return [
      s.nome,
      senhaParaGravar_(s.senha),
      s.tipo === 'autorizado' ? 'autorizado' : 'solicitante',
      s.foto || '',
      s.email || '',
      s.aprovado === undefined ? true : !!s.aprovado,
      !!s.autorizadoAbreChamados,
    ];
  });
  garantirColunasTexto_(sh, COLUNAS_TEXTO_SOLICITANTES);
  substituirLinhasComSeguranca_(sh, rows, 7);
}

// Foto de perfil vem do navegador já reduzida (240px, JPEG — uns 10-20 mil
// caracteres em base64). O limite aqui é só pra barrar quem manda direto
// pro backend, sem passar pelo app.
const LIMITE_FOTO_PERFIL = 40000;

// Atualiza só a foto de UM solicitante (o autenticado). Antes regravava a
// aba Solicitantes inteira (via saveSolicitantes) só pra trocar uma célula
// — e sem limite de tamanho, uma foto gigante derrubava a gravação com a
// aba já apagada, sumindo com todas as contas. Agora valida a foto e
// escreve só a célula da coluna "foto" da linha dessa pessoa; o resto da
// planilha nem é tocado. Quem chama só pode mudar a própria foto — nunca a
// de outra pessoa.
function atualizarFotoSolicitante(nome, novaFoto) {
  const foto = String(novaFoto || '');
  if (foto && (foto.indexOf('data:image/') !== 0 || foto.length > LIMITE_FOTO_PERFIL)) {
    return { ok: false, error: 'Foto inválida ou grande demais.' };
  }
  const sh = getOrCreateSheet('Solicitantes', ['nome', 'senha', 'tipo', 'foto', 'email', 'aprovado', 'autorizadoAbreChamados']);
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return { ok: false, error: 'Usuário não encontrado.' };
  const nomes = sh.getRange(2, 1, lastRow - 1, 1).getValues();
  const alvo = String(nome || '').trim().toLowerCase();
  for (let i = 0; i < nomes.length; i++) {
    if (String(nomes[i][0] || '').trim().toLowerCase() === alvo) {
      sh.getRange(i + 2, 4).setValue(foto); // +2: pula o cabeçalho, base 1; coluna 4 = foto
      return { ok: true };
    }
  }
  return { ok: false, error: 'Usuário não encontrado.' };
}

// Autentica só nome+senha, sem olhar o tipo — usado no LOGIN (ramo userNome
// de doGet), onde tanto solicitante quanto autorizado podem entrar. Devolve
// o registro mesmo se aprovado=false (quem chama decide o que fazer com
// isso — ver o ramo userNome em doGet, que dá uma mensagem específica pra
// cadastro pendente em vez de "senha incorreta").
function autenticarUsuarioLogin(nome, senha) {
  const s = findSolicitante(nome);
  if (!s) return null;
  if (!senhaConfere_(senha, s.senha)) return null;
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
  return usuarioPodeUsarChamados_(s) ? s : null;
}

function usuarioPodeUsarChamados_(s) {
  if (!s || s.aprovado === false) return false;
  if (s.tipo === 'autorizado' && !s.autorizadoAbreChamados) return false;
  return true;
}

// ---------- Proteção contra força bruta no login ----------
// Antes, dava pra tentar senha atrás de senha sem limite nenhum — com
// senha de admin curta (ex: só dígitos), um script quebra isso em horas.
// Guarda tentativas erradas no CacheService (compartilhado entre todas as
// execuções do script, expira sozinho) por nome tentado — adminNome e
// userNome chegam iguais no login unificado, então um mesmo balde cobre
// tentativa de admin ou de solicitante com aquele nome. Restauração de
// sessão (só manda o secret salvo, sem nome nenhum) não passa por este
// balde — tem o seu próprio, geral (ver secretBloqueado_ logo abaixo).
const LOGIN_MAX_TENTATIVAS = 5;
const LOGIN_BLOQUEIO_SEGUNDOS = 15 * 60;

function loginCacheChave_(identificador) {
  return 'login_tentativas_' + String(identificador || '').trim().toLowerCase();
}

function loginBloqueado_(identificador) {
  if (!identificador) return false;
  const valor = Number(CacheService.getScriptCache().get(loginCacheChave_(identificador)) || 0);
  return valor >= LOGIN_MAX_TENTATIVAS;
}

function loginRegistrarFalha_(identificador) {
  if (!identificador) return;
  const cache = CacheService.getScriptCache();
  const chave = loginCacheChave_(identificador);
  const atual = Number(cache.get(chave) || 0);
  cache.put(chave, String(atual + 1), LOGIN_BLOQUEIO_SEGUNDOS);
}

function loginLimparTentativas_(identificador) {
  if (!identificador) return;
  CacheService.getScriptCache().remove(loginCacheChave_(identificador));
}

// O balde por nome acima não cobria quem manda SÓ o secret, sem nome
// nenhum (o formato da restauração de sessão, e de todo doPost de admin):
// dava pra testar senha de admin atrás de senha sem limite, e qualquer
// acerto já virava sessão de admin (findAdminBySecret aceita a senha de
// qualquer admin). Aqui um contador geral de "secret sem nome errado", que
// vale pra doGet e doPost juntos: passou do limite, nenhum secret sem nome
// é aceito até expirar. Login normal (com nome) não conta aqui — quem só
// errou a própria senha de solicitante não trava o login de admin de
// ninguém. Efeito colateral aceito: um ataque ativo derruba a restauração
// de sessão dos admins por até 15 min (basta entrar de novo com nome+senha,
// que usa o balde por nome e não este).
const SECRET_MAX_FALHAS = 20;
const SECRET_CHAVE_CACHE = 'secret_sem_nome_falhas';

function secretBloqueado_() {
  const valor = Number(CacheService.getScriptCache().get(SECRET_CHAVE_CACHE) || 0);
  return valor >= SECRET_MAX_FALHAS;
}

function secretRegistrarFalha_() {
  const cache = CacheService.getScriptCache();
  const atual = Number(cache.get(SECRET_CHAVE_CACHE) || 0);
  cache.put(SECRET_CHAVE_CACHE, String(atual + 1), LOGIN_BLOQUEIO_SEGUNDOS);
}

// ---------- Cadastro público de solicitante ----------
//
// Antes, só um admin podia criar acesso de solicitante (em Usuários). Agora
// qualquer pessoa pode se cadastrar sozinha direto na tela de login (nome +
// email + senha escolhida por ela) — mas a conta nasce com aprovado=false e
// não consegue entrar até um administrador aprovar em Usuários. Roda dentro
// do doGet (não do doPost) porque, quando foi feito, o doPost ia em modo
// no-cors e o navegador não conseguia ler a resposta (hoje consegue — ver
// backendPost no app.jsx). Isso só é seguro por causa do LockService em
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
  // Rota pública, sem login: sem limite, dava pra encher a planilha de
  // cadastros gigantes.
  if (nome.length > 100 || email.length > 200 || senha.length > 200) {
    return { ok: false, error: 'Nome, email ou senha longos demais.' };
  }
  // Ver "Proteção contra fórmulas na planilha".
  if (comecaComCaracterDeFormula_(nome) || comecaComCaracterDeFormula_(email) || comecaComCaracterDeFormula_(senha)) {
    return { ok: false, error: 'Nome, email e senha não podem começar com = + - ou @.' };
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
  // Acrescenta só a linha nova (mesma ordem de colunas de saveSolicitantes)
  // em vez de regravar a aba inteira a cada cadastro público.
  const sh = getOrCreateSheet('Solicitantes', ['nome', 'senha', 'tipo', 'foto', 'email', 'aprovado', 'autorizadoAbreChamados']);
  garantirColunasTexto_(sh, COLUNAS_TEXTO_SOLICITANTES);
  sh.appendRow([nome, gerarHashSenha_(senha), 'solicitante', '', email, false, false]);
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

// O Google Sheets recusa gravar mais de 50.000 caracteres numa única
// célula — e o chamado inteiro (mensagens + fotos em base64) vira UMA
// célula só aqui. Sem essa checagem, uma foto grande demais fazia
// setValues/appendRow lançar uma exceção sem tratamento nenhum, que subia
// até derrubar a resposta inteira do backend (o app "travava" ao abrir ou
// responder um chamado com foto). Agora barra antes de tentar gravar, com
// uma margem de segurança abaixo do limite real.
const LIMITE_CARACTERES_CELULA = 45000;

function salvarChamado_(chamado) {
  const sh = getOrCreateSheet('Chamados', ['id', 'dados']);
  const linha = encontrarLinhaChamado_(sh, chamado.id);
  const json = JSON.stringify(chamado);
  if (json.length > LIMITE_CARACTERES_CELULA) {
    throw new Error('CHAMADO_MUITO_GRANDE');
  }
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

// ---------- Token do Firebase pro admin (leitura ao vivo dos chamados) ----------

// Antes, a tela do admin entrava no Firebase com login ANÔNIMO — e como
// qualquer pessoa na internet consegue fazer esse mesmo login anônimo (a
// config do Firebase está no index.html público), as regras do Firestore
// não tinham como diferenciar um admin de um estranho: quem abrisse a
// coleção "chamados" lia tudo. Agora o backend, depois de conferir a senha
// do admin, gera um "custom token" do Firebase assinado com a chave da
// conta de serviço (a mesma de getFirestore_), com a marca chamados=true.
// O frontend entra no Firebase com esse token (signInWithCustomToken) e as
// regras do Firestore só liberam leitura pra quem tem essa marca:
//
//   match /chamados/{id} {
//     allow read: if request.auth != null && request.auth.token.chamados == true;
//     allow write: if false;
//   }
//
// Só sai token pra admin que enxerga Chamados (master ou com "chamados"
// nas permissões) — os outros continuam sem leitura ao vivo, como já era
// na prática. O token vale 1h só pra fazer o login; depois disso o próprio
// SDK do Firebase mantém a sessão enquanto a aba estiver aberta. Se faltar
// alguma propriedade do script ou a assinatura falhar, devolve '' e o app
// segue sem atualização ao vivo (lendo pelo doGet, como sempre).
function adminVeChamados_(admin) {
  if (admin.permissoes === 'todas') return true;
  return String(admin.permissoes || '').split(',').map(function (s) { return s.trim().toLowerCase(); }).indexOf('chamados') !== -1;
}

function base64UrlSemPadding_(valor) {
  return Utilities.base64EncodeWebSafe(valor).replace(/=+$/, '');
}

function gerarTokenFirebaseAdmin_(admin) {
  if (!adminVeChamados_(admin)) return '';
  return gerarTokenFirebase_('admin', admin.nome, { chamados: true });
}

// Mesma ideia pro solicitante (e pro "autorizado" que pode abrir chamados):
// a marca "solicitante" leva o nome EXATO como está cadastrado — o mesmo
// que o servidor grava em criadoPor ao abrir o chamado (ver novoChamado) —
// e as regras do Firestore só liberam os chamados em que os dois batem:
//
//   allow read: if request.auth != null && (
//     request.auth.token.chamados == true ||
//     (request.auth.token.solicitante is string &&
//      resource.data.criadoPor == request.auth.token.solicitante));
//
// Ou seja: cada solicitante enxerga ao vivo só os próprios chamados, nunca
// os de outra pessoa — igual ao filtro "meusChamados" do doGet.
function gerarTokenFirebaseSolicitante_(usuario) {
  if (usuario.tipo === 'autorizado' && !usuario.autorizadoAbreChamados) return '';
  return gerarTokenFirebase_('user', usuario.nome, { solicitante: usuario.nome });
}

function gerarTokenFirebase_(prefixoUid, nome, claims) {
  try {
    const props = PropertiesService.getScriptProperties();
    const email = props.getProperty('FIREBASE_CLIENT_EMAIL');
    const chavePrivada = props.getProperty('FIREBASE_PRIVATE_KEY');
    if (!email || !chavePrivada) return '';
    // uid estável por pessoa, sem expor o nome (o uid aparece no painel
    // Authentication do Firebase) e sempre dentro do limite de 128 chars.
    const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, prefixoUid + ':' + String(nome || '').trim().toLowerCase(), Utilities.Charset.UTF_8);
    const uid = prefixoUid + '-' + digest.map(function (b) { return ('0' + (b & 0xff).toString(16)).slice(-2); }).join('').slice(0, 32);
    const agora = Math.floor(Date.now() / 1000);
    const header = { alg: 'RS256', typ: 'JWT' };
    const payload = {
      iss: email,
      sub: email,
      aud: 'https://identitytoolkit.googleapis.com/google.identity.identitytoolkit.v1.IdentityToolkit',
      iat: agora,
      exp: agora + 3600,
      uid: uid,
      claims: claims,
    };
    const entrada = base64UrlSemPadding_(Utilities.newBlob(JSON.stringify(header)).getBytes()) + '.' +
      base64UrlSemPadding_(Utilities.newBlob(JSON.stringify(payload)).getBytes());
    // Mesmo tratamento do "\n" literal que getFirestore_ faz na chave.
    const assinatura = Utilities.computeRsaSha256Signature(entrada, chavePrivada.replace(/\\n/g, '\n'));
    return entrada + '.' + base64UrlSemPadding_(assinatura);
  } catch (err) {
    Logger.log('Falha ao gerar token do Firebase (' + prefixoUid + '): ' + err);
    return '';
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
//
// Antes, se o JSON salvo não pudesse ser lido, readState devolvia um estado
// VAZIO como se estivesse tudo certo — e o admin que entrasse em seguida
// recebia inventário zerado, que o autosave podia gravar de volta por cima
// do inventário de verdade. Agora: planilha sem nada salvo ainda (só o
// cabeçalho) continua sendo "estado vazio" legítimo; mas se há conteúdo e
// ele não é JSON válido, tenta de novo algumas vezes (pode ter sido lido
// bem no meio de uma gravação — o doGet não usa a trava) e, se continuar
// ilegível, lança ESTADO_ILEGIVEL em vez de fingir que está vazio.
function readState() {
  const sh = getOrCreateSheet('AppState', ['data']);
  const defaults = { categorias: [], areas: [], responsaveis: [], inventario: [] };
  for (let tentativa = 0; tentativa < 3; tentativa++) {
    const lastRow = sh.getLastRow();
    if (lastRow < 2) return defaults;
    const values = sh.getRange(2, 1, lastRow - 1, 1).getValues();
    const combined = values.map(function (r) { return r[0]; }).join('');
    try {
      const parsed = JSON.parse(combined);
      delete parsed.chamados;
      return parsed;
    } catch (e) {
      Utilities.sleep(700);
    }
  }
  Logger.log('AppState ilegível depois de 3 tentativas — nada foi devolvido como vazio.');
  throw new Error('ESTADO_ILEGIVEL');
}

// Por segurança, remove "chamados" antes de gravar mesmo que alguém passe
// por engano — chamados nunca devem ser persistidos aqui.
//
// O JSON é gravado em pedaços de até ~45 mil caracteres (limite de 50 mil
// por célula do Sheets), um por linha. Dois cuidados que antes não havia:
// (1) nenhum pedaço começa com = + - @ ou ' — o Sheets trataria "=..." como
// fórmula (e "'..." perderia o apóstrofo), corrompendo o JSON; a coluna
// também é formatada como texto puro. (2) grava por cima e só depois limpa
// as linhas que sobraram, em vez de apagar tudo primeiro: se a gravação
// falhar no meio, o conteúdo anterior não some.
const CARACTERES_PERIGOSOS_NO_INICIO = "=+-@'";

function writeState(obj) {
  const sh = getOrCreateSheet('AppState', ['data']);
  const paraGravar = Object.assign({}, obj);
  delete paraGravar.chamados;
  const json = JSON.stringify(paraGravar);
  const CHUNK = 45000;
  const chunks = [];
  let inicio = 0;
  while (inicio < json.length) {
    let fim = Math.min(inicio + CHUNK, json.length);
    // Empurra o corte pra frente enquanto o próximo pedaço fosse começar
    // com um caractere perigoso (no máximo alguns caracteres a mais).
    while (fim < json.length && CARACTERES_PERIGOSOS_NO_INICIO.indexOf(json.charAt(fim)) !== -1) fim++;
    chunks.push([json.slice(inicio, fim)]);
    inicio = fim;
  }
  const lastRowAntes = sh.getLastRow();
  sh.getRange(2, 1, Math.max(chunks.length, 1), 1).setNumberFormat('@');
  sh.getRange(2, 1, chunks.length, 1).setValues(chunks);
  const primeiraSobrando = chunks.length + 2;
  if (lastRowAntes >= primeiraSobrando) {
    sh.getRange(primeiraSobrando, 1, lastRowAntes - primeiraSobrando + 1, 1).clearContent();
  }
}

// ---------- Versão do estado (controle de concorrência entre admins) ----------

// O autosave manda o ESTADO INTEIRO (inventário, categorias, salas...). Sem
// controle, um admin com a aba aberta desde cedo, ao salvar, gravava a cópia
// DELE por cima — apagando o que outro admin tinha editado nesse meio tempo.
// Agora cada gravação vem com a versão em que ela se baseou (baseVersao) e a
// versão nova que ela cria (novaVersao, gerada no navegador). Se a
// baseVersao não for a atual, alguém salvou antes: a gravação é recusada
// com CONFLITO_VERSAO e o navegador avisa a pessoa pra recarregar
// (action=versaoEstado no doGet ficou só pra navegadores antigos, da época
// em que o POST era no-cors e a resposta não podia ser lida). Navegador antigo (sem baseVersao, de antes dessa
// mudança) continua sendo aceito, pra não quebrar quem ainda não recarregou.
const VERSAO_ESTADO_PROP = 'appStateVersao';

function versaoEstadoAtual_() {
  return PropertiesService.getScriptProperties().getProperty(VERSAO_ESTADO_PROP) || '';
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
//
// ESTADO_ILEGIVEL (ver readState) vira uma resposta de erro normal — a tela
// mostra "tente de novo" em vez de abrir com inventário vazio.
function doGet(e) {
  try {
    return doGetInterno_(e);
  } catch (err) {
    if (err.message !== 'ESTADO_ILEGIVEL') throw err;
    const payload = { ok: false, error: 'Não foi possível ler o inventário agora. Tente de novo em alguns segundos; se continuar, avise o suporte de TI.' };
    const callback = (e && e.parameter && e.parameter.callback) || '';
    if (callback) {
      return ContentService
        .createTextOutput(callback + '(' + JSON.stringify(payload) + ');')
        .setMimeType(ContentService.MimeType.JAVASCRIPT);
    }
    return jsonOut(payload);
  }
}

function responderGet_(payload, callback) {
  if (callback) {
    return ContentService
      .createTextOutput(callback + '(' + JSON.stringify(payload) + ');')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return jsonOut(payload);
}

function montarPayloadAdmin_(admin, admins) {
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
  return {
    ok: true,
    isAdmin: true,
    nome: admin.nome,
    permissoes: admin.permissoes,
    editar: !!admin.editar,
    podeAbrirChamados: !!admin.podeAbrirChamados,
    podeResponderChamados: !!admin.podeResponderChamados,
    responderSoProprios: !!admin.responderSoProprios,
    state: Object.assign({}, state, { chamados: listarChamados() }),
    stateVersao: versaoEstadoAtual_(),
    // Inclui a senha (como já fazemos pra solicitantes) porque o frontend
    // (Administradores.save()) usa modal.original.senha pra manter a senha
    // de quem já existe quando o campo "nova senha" fica em branco. Sem
    // isso, editar qualquer admin (ex: só mudar uma permissão) sem
    // redigitar a senha mandava senha undefined pro salvarAdmins e travava
    // o login desse admin. Hoje o que vai é o HASH (ver senhaParaGravar_),
    // não a senha em si.
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
    historico: getHistoricoMensal(),
    firebaseToken: gerarTokenFirebaseAdmin_(admin)
  };
}

function montarPayloadUsuario_(usuario) {
  const state = readState();
  if (usuario.tipo === 'autorizado') {
    // Autorizado: entra como um usuário comum (não como admin), com
    // visão só-leitura de Painel/Categorias/Salas/Inventário — sem
    // admins/solicitantes/histórico. Por padrão também sem chamado nenhum;
    // se autorizadoAbreChamados estiver marcado, ganha os PRÓPRIOS
    // chamados (nunca os de outras pessoas), do mesmo jeito que um
    // solicitante comum.
    const podeAbrirChamados = !!usuario.autorizadoAbreChamados;
    const alvoAutorizado = usuario.nome.trim().toLowerCase();
    const chamadosDoAutorizado = podeAbrirChamados
      ? listarChamados().filter(function (c) { return (c.criadoPor || '').trim().toLowerCase() === alvoAutorizado; })
      : [];
    return {
      ok: true,
      isAdmin: false,
      isUser: true,
      isAutorizado: true,
      podeAbrirChamados: podeAbrirChamados,
      nome: usuario.nome,
      foto: usuario.foto || '',
      firebaseToken: gerarTokenFirebaseSolicitante_(usuario),
      state: {
        categorias: state.categorias || [],
        areas: state.areas || [],
        responsaveis: state.responsaveis || [],
        inventario: state.inventario || [],
        chamados: chamadosDoAutorizado
      }
    };
  }
  // Usa usuario.nome (o nome de verdade, resolvido pelo login) em vez do
  // que a pessoa digitou — desde que o login também aceita email (ver
  // findSolicitante), comparar direto com o que foi digitado deixaria "meus
  // chamados" vazio pra quem entra pelo email, já que criadoPor sempre
  // guarda o nome, nunca o email.
  const alvo = usuario.nome.trim().toLowerCase();
  const meusChamados = listarChamados().filter(function (c) {
    return (c.criadoPor || '').trim().toLowerCase() === alvo;
  });
  return {
    ok: true,
    isAdmin: false,
    isUser: true,
    nome: usuario.nome,
    foto: usuario.foto || '',
    firebaseToken: gerarTokenFirebaseSolicitante_(usuario),
    state: { areas: state.areas || [], categorias: state.categorias || [], chamados: meusChamados }
  };
}

function doGetInterno_(e) {
  const p = (e && e.parameter) || {};
  const callback = p.callback || '';

  if (p.action === 'cadastro') {
    return responderGet_(cadastrarSolicitanteComTrava_(p.novoNome, p.novoEmail, p.novoSenha), callback);
  }

  // Sessão (ver criarSessao_): o jeito normal de abrir o app depois do
  // primeiro login. Token inválido/vencido devolve sessaoInvalida — o
  // navegador apaga o token e mostra a tela de login.
  if (p.sessao) {
    const sessao = resolverSessao_(p.sessao);
    if (!sessao) {
      return responderGet_({ ok: true, isAdmin: false, isUser: false, sessaoInvalida: true }, callback);
    }
    if (sessao.admin) {
      if (p.action === 'versaoEstado') {
        return responderGet_({ ok: true, isAdmin: true, versao: versaoEstadoAtual_() }, callback);
      }
      return responderGet_(montarPayloadAdmin_(sessao.admin, sessao.admins), callback);
    }
    return responderGet_(montarPayloadUsuario_(sessao.usuario), callback);
  }

  const secret = p.secret || '';
  const adminNome = p.adminNome || '';
  const userNome = p.userNome || '';
  const userSenha = p.userSenha || '';

  // Nome que a pessoa digitou de verdade (adminNome e userNome chegam
  // iguais no login unificado) — usado só pra travar tentativas erradas
  // repetidas; nunca fica vazio numa tentativa de login real (só em
  // restauração de sessão antiga, que não passa por aqui).
  const identificadorTentativa = (userNome || adminNome || '').trim();
  const muitasTentativas = { ok: true, isAdmin: false, isUser: false, error: 'Muitas tentativas erradas. Aguarde alguns minutos e tente de novo.' };
  if (identificadorTentativa && loginBloqueado_(identificadorTentativa)) {
    return responderGet_(muitasTentativas, callback);
  }

  // Secret sem nome = restauração de sessão do jeito antigo (navegador que
  // ainda tem a senha guardada — ele recebe um token de sessão e apaga a
  // senha) ou alguém testando senhas de admin às cegas — ver secretBloqueado_.
  const secretSemNome = !!secret && !adminNome;
  if (secretSemNome && secretBloqueado_()) {
    return responderGet_(muitasTentativas, callback);
  }

  let admins = secret ? getAdmins() : null;
  let admin = null;
  if (secret && adminNome) {
    // Login de verdade: nome E senha precisam bater com o MESMO admin — não
    // só a senha sozinha, que antes bastava pra entrar em qualquer conta
    // cuja senha alguém soubesse, mesmo sem saber de quem era.
    const candidato = findAdminPorNome_(adminNome, admins);
    if (candidato && senhaConfere_(secret, candidato.senha)) admin = candidato;
  } else if (secret) {
    admin = findAdminBySecret(secret, admins);
    if (!admin) secretRegistrarFalha_();
  }

  let payload;
  if (admin && p.action === 'versaoEstado') {
    // Navegador antigo (de antes das sessões) conferindo o autosave.
    payload = { ok: true, isAdmin: true, versao: versaoEstadoAtual_() };
  } else if (admin) {
    if (identificadorTentativa) loginLimparTentativas_(identificadorTentativa);
    const senhaArmazenada = migrarSenhaSeTextoPuro_('Admins', admin.nome, secret, admin.senha);
    // Se a senha acabou de virar hash, relê a lista pra o master não
    // receber (e depois devolver) o valor antigo.
    if (senhaArmazenada !== admin.senha) admins = getAdmins();
    payload = montarPayloadAdmin_(admin, admins);
    payload.sessao = criarSessao_('a', admin.nome, senhaArmazenada);
  } else if (userNome) {
    const usuario = autenticarUsuarioLogin(userNome, userSenha);
    if (usuario && identificadorTentativa) loginLimparTentativas_(identificadorTentativa);
    if (usuario && usuario.aprovado === false) {
      payload = { ok: true, isAdmin: false, isUser: false, error: 'Seu cadastro ainda está aguardando aprovação de um administrador.' };
    } else if (usuario) {
      const senhaArmazenada = migrarSenhaSeTextoPuro_('Solicitantes', usuario.nome, userSenha, usuario.senha);
      payload = montarPayloadUsuario_(usuario);
      payload.sessao = criarSessao_('u', usuario.nome, senhaArmazenada);
    } else {
      if (identificadorTentativa) loginRegistrarFalha_(identificadorTentativa);
      payload = { ok: true, isAdmin: false, isUser: false, error: 'Nome ou senha incorretos.' };
    }
  } else {
    payload = { ok: true, isAdmin: false, isUser: false, state: { areas: [], categorias: [], chamados: [] } };
  }

  return responderGet_(payload, callback);
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

  if (body.action === 'sair') {
    encerrarSessao_(body.sessao);
    return jsonOut({ ok: true });
  }

  // Sessão (ver criarSessao_) é o jeito normal. Os campos antigos (secret
  // do admin, userNome+userSenha do solicitante) continuam aceitos pra não
  // quebrar uma aba que ficou aberta desde antes dessa mudança.
  let admin = null;
  let usuarioDaSessao = null;
  if (body.sessao) {
    const sessao = resolverSessao_(body.sessao);
    if (!sessao) return jsonOut({ ok: false, error: 'SESSAO_INVALIDA' });
    admin = sessao.admin || null;
    usuarioDaSessao = sessao.usuario || null;
  } else if (body.secret && !secretBloqueado_()) {
    // doPost de admin do jeito antigo manda só o secret, sem nome — mesmo
    // contador geral do doGet (ver secretBloqueado_). Bloqueado, o secret é
    // ignorado: a ação segue como se fosse de alguém não logado como admin.
    admin = findAdminBySecret(body.secret);
    if (!admin) secretRegistrarFalha_();
  }
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
  // Solicitante (ou autorizado que abre chamados) que está fazendo o pedido.
  function solicitanteDoPedido_() {
    if (usuarioDaSessao) return usuarioPodeUsarChamados_(usuarioDaSessao) ? usuarioDaSessao : null;
    if (body.sessao) return null;
    return authenticateSolicitante(body.userNome, body.userSenha);
  }

  // Confirmação de senha antes de uma ação perigosa (ex: apagar o
  // inventário inteiro de uma unidade, em Importar/Exportar). Antes o
  // navegador comparava com a senha guardada nele; agora ele não tem mais a
  // senha, então quem confere é o servidor — com o mesmo limite de
  // tentativas erradas do login.
  if (isAdmin && body.action === 'conferirSenha') {
    if (loginBloqueado_(admin.nome)) {
      return jsonOut({ ok: false, error: 'Muitas tentativas erradas. Aguarde alguns minutos e tente de novo.' });
    }
    const valida = senhaConfere_(body.senha, admin.senha);
    if (valida) loginLimparTentativas_(admin.nome); else loginRegistrarFalha_(admin.nome);
    return jsonOut({ ok: true, valida: valida });
  }

  if (isAdmin && body.action === 'salvarTudo') {
    // Ver VERSAO_ESTADO_PROP: recusa cópia baseada numa versão que não é
    // mais a atual (outro admin salvou antes).
    const props = PropertiesService.getScriptProperties();
    if (body.baseVersao !== undefined && String(body.baseVersao) !== versaoEstadoAtual_()) {
      return jsonOut({ ok: false, error: 'CONFLITO_VERSAO' });
    }
    let estadoFiltrado;
    try {
      estadoFiltrado = filtrarEstadoPorPermissao(body.state, admin);
    } catch (err) {
      // Estado atual ilegível (ver readState): não grava nada por cima.
      if (err.message === 'ESTADO_ILEGIVEL') return jsonOut({ ok: false, error: 'ESTADO_ILEGIVEL' });
      throw err;
    }
    writeState(estadoFiltrado);
    const novaVersao = String(body.novaVersao || '');
    props.setProperty(VERSAO_ESTADO_PROP, /^[A-Za-z0-9-]{1,40}$/.test(novaVersao) ? novaVersao : Utilities.getUuid());
    registrarSnapshotMensal(estadoFiltrado);
    return jsonOut({ ok: true });
  }

  // CELULA_MUITO_GRANDE vem de substituirLinhasComSeguranca_ — nesse caso a
  // planilha não foi alterada em nada.
  if (isMaster && body.action === 'salvarAdmins') {
    try {
      saveAdmins(body.admins || []);
    } catch (err) {
      if (err.message === 'CELULA_MUITO_GRANDE') return jsonOut({ ok: false, error: 'Algum campo está grande demais. Nada foi alterado.' });
      throw err;
    }
    return jsonOut({ ok: true });
  }

  if (isMaster && body.action === 'salvarSolicitantes') {
    try {
      saveSolicitantes(body.solicitantes || []);
    } catch (err) {
      if (err.message === 'CELULA_MUITO_GRANDE') return jsonOut({ ok: false, error: 'Algum campo está grande demais. Nada foi alterado.' });
      throw err;
    }
    return jsonOut({ ok: true });
  }

  // Autoatendimento: o próprio usuário (solicitante ou autorizado) troca a
  // própria foto de perfil, sem precisar de admin. Autentica só com
  // nome+senha (não authenticateSolicitante, que bloqueia autorizado — aqui
  // ambos os tipos podem mudar a própria foto) e só mexe na foto, nunca em
  // nome/senha/tipo de ninguém.
  if (body.action === 'atualizarFotoUsuario') {
    const usuario = usuarioDaSessao || (body.sessao ? null : autenticarUsuarioLogin(body.userNome, body.userSenha));
    if (!usuario) {
      return jsonOut({ ok: false, error: 'Não autenticado' });
    }
    return jsonOut(atualizarFotoSolicitante(usuario.nome, body.foto));
  }

  // Antes: novoChamado e novaMensagem não checavam login nenhum — qualquer
  // um com a URL pública do backend podia criar chamados falsos ou postar
  // mensagens em qualquer chamado, inclusive se passando pelo TI
  // (mensagem.autor = 'ti'). Agora exige um solicitante autenticado (ou um
  // admin) e o autor/criadoPor são fixados pelo servidor, nunca aceitos do
  // jeito que o cliente mandou.
  if (body.action === 'novoChamado') {
    const solicitante = solicitanteDoPedido_();
    if (!adminPodeAbrirChamados && !solicitante) {
      return jsonOut({ ok: false, error: 'Não autenticado' });
    }
    // Antes: o chamado era gravado do jeito que o cliente mandou — id,
    // status, mensagens e tudo. Como salvarChamado_ sobrescreve a linha se o
    // id já existe, qualquer solicitante logado conseguia substituir o
    // chamado de OUTRA pessoa mandando o mesmo id, abrir chamado já
    // "Resolvido" ou colocar mensagem com autor 'ti'. Agora o servidor monta
    // o chamado e só aproveita do cliente os campos que a pessoa preenche.
    // O id continua vindo do cliente (o frontend já usa esse id na tela pra
    // mandar as próximas mensagens), mas só se tiver o formato do uid("CH")
    // e ainda não existir.
    const recebido = body.chamado || {};
    const idRecebido = String(recebido.id || '');
    if (!/^CH-[A-Z0-9]{1,20}$/.test(idRecebido)) {
      return jsonOut({ ok: false, error: 'Chamado inválido' });
    }
    if (buscarChamadoPorId_(idRecebido)) {
      return jsonOut({ ok: false, error: 'Chamado já existe' });
    }
    const textoCampo_ = function (v, max) { return String(v || '').slice(0, max); };
    const primeiraMsg = (Array.isArray(recebido.mensagens) && recebido.mensagens[0]) || {};
    const fotoRecebida = String(recebido.foto || '');
    const agora = new Date().toISOString();
    const chamado = {
      id: idRecebido,
      assunto: textoCampo_(recebido.assunto, 60),
      tipo: textoCampo_(recebido.tipo, 100),
      unidade: textoCampo_(recebido.unidade, 100),
      sala: textoCampo_(recebido.sala, 200),
      categoria: textoCampo_(recebido.categoria, 200),
      foto: fotoRecebida.indexOf('data:image/') === 0 ? fotoRecebida : '',
      status: 'Aberto',
      criadoEm: agora,
      mensagens: [{
        // A primeira mensagem é sempre a descrição do problema, inclusive
        // quando um admin abre em nome de alguém — mesmo formato de antes.
        autor: 'solicitante',
        texto: textoCampo_(primeiraMsg.texto, 5000),
        data: agora,
      }],
    };
    if (solicitante) {
      chamado.criadoPor = solicitante.nome;
      chamado.solicitante = solicitante.nome;
    } else if (isAdmin) {
      // Marca quem abriu quando é um admin abrindo direto (ex: recepção
      // atendendo alguém pessoalmente) — é o que permite restringir esse
      // admin a só responder os próprios chamados (responderSoProprios).
      chamado.abertoPorAdmin = admin.nome;
    }
    try {
      salvarChamado_(chamado);
    } catch (err) {
      if (err.message === 'CHAMADO_MUITO_GRANDE') {
        return jsonOut({ ok: false, error: 'A foto é grande demais para salvar. Tire a foto de novo com menos detalhe ou escolha outra.' });
      }
      throw err;
    }
    notificarNovoChamado(chamado);
    sincronizarChamadoNoFirestore_(chamado);
    return jsonOut({ ok: true });
  }

  if (body.action === 'novaMensagem') {
    const solicitante = solicitanteDoPedido_();
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
    // Monta a mensagem no servidor em vez de gravar o objeto do cliente com
    // só autor/nome trocados — senão dava pra enfiar qualquer campo extra
    // (ou data falsa) dentro do chamado.
    const recebida = body.mensagem || {};
    const mensagem = {
      autor: adminPodeResponderChamados ? 'ti' : 'solicitante',
      // Nome de quem respondeu de verdade (nunca o que o cliente mandou) —
      // antes toda resposta de admin aparecia só como "Administrador" pra
      // todo mundo, sem dar pra saber QUAL admin respondeu.
      nome: adminPodeResponderChamados ? admin.nome : solicitante.nome,
      texto: String(recebida.texto || '').slice(0, 5000),
      data: new Date().toISOString(),
    };
    chamado.mensagens.push(mensagem);
    try {
      salvarChamado_(chamado);
    } catch (err) {
      if (err.message === 'CHAMADO_MUITO_GRANDE') {
        return jsonOut({ ok: false, error: 'A foto é grande demais para salvar. Tire a foto de novo com menos detalhe ou escolha outra.' });
      }
      throw err;
    }
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
    // Mesma lista de CHAMADO_STATUS_OPTIONS no frontend — qualquer outro
    // valor bagunçaria o painel e o histórico mensal (contarPor).
    if (['Aberto', 'Em andamento', 'Resolvido'].indexOf(body.status) === -1) {
      return jsonOut({ ok: false, error: 'Status inválido' });
    }
    const chamado = buscarChamadoPorId_(body.chamadoId);
    if (!chamado) {
      return jsonOut({ ok: false, error: 'Chamado não encontrado' });
    }
    if (!podeMexerNesseChamado_(chamado)) {
      return jsonOut({ ok: false, error: 'Você só pode mudar o status dos chamados que você mesma abriu' });
    }
    // Mudou o status: a resposta do solicitante ao "Ficou bom?" valia pro
    // status anterior. Se o TI marcar Resolvido de novo, ele pergunta outra vez.
    if (chamado.status !== body.status) delete chamado.confirmacao;
    chamado.status = body.status;
    try {
      salvarChamado_(chamado);
    } catch (err) {
      if (err.message === 'CHAMADO_MUITO_GRANDE') {
        return jsonOut({ ok: false, error: 'Não foi possível salvar: esse chamado ficou grande demais (provavelmente por causa de uma foto antiga). Avise o suporte.' });
      }
      throw err;
    }
    sincronizarChamadoNoFirestore_(chamado);
    return jsonOut({ ok: true });
  }

  // "Ficou bom?" — quem abriu o chamado responde depois que o TI marcou
  // como Resolvido. Sim: fica registrado (o app mostra como "Fechado").
  // Não: o chamado volta pra "Em andamento" com uma mensagem avisando o TI.
  if (body.action === 'confirmarChamado') {
    const solicitante = solicitanteDoPedido_();
    if (!solicitante) {
      return jsonOut({ ok: false, error: 'Não autenticado' });
    }
    const chamado = buscarChamadoPorId_(body.chamadoId);
    if (!chamado) {
      return jsonOut({ ok: false, error: 'Chamado não encontrado' });
    }
    const dono = String(chamado.criadoPor || chamado.solicitante || '').trim().toLowerCase();
    if (dono !== solicitante.nome.trim().toLowerCase()) {
      return jsonOut({ ok: false, error: 'Sem permissão para responder este chamado' });
    }
    if (chamado.status !== 'Resolvido') {
      return jsonOut({ ok: false, error: 'Esse chamado não está marcado como resolvido.' });
    }
    const agora = new Date().toISOString();
    if (body.resolveu === true) {
      chamado.confirmacao = { resolveu: true, data: agora };
    } else {
      delete chamado.confirmacao;
      chamado.status = 'Em andamento';
      chamado.mensagens.push({
        autor: 'solicitante',
        nome: solicitante.nome,
        texto: 'Não resolveu — o problema continua.' + (body.texto ? ' ' + String(body.texto).slice(0, 2000) : ''),
        data: agora,
        reabertura: true,
      });
    }
    try {
      salvarChamado_(chamado);
    } catch (err) {
      if (err.message === 'CHAMADO_MUITO_GRANDE') {
        return jsonOut({ ok: false, error: 'Não foi possível salvar: esse chamado ficou grande demais. Avise o suporte.' });
      }
      throw err;
    }
    sincronizarChamadoNoFirestore_(chamado);
    return jsonOut({ ok: true, chamado: chamado });
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

// ---------- Backup automático diário ----------

// Até aqui a única "cópia de segurança" era o histórico de versões do
// próprio Google Sheets — que ajuda, mas é difícil de usar numa emergência
// e não protege se a planilha for apagada. fazerBackupDiario copia TODAS as
// abas (só os dados, sem o código do Apps Script) pra uma planilha nova
// dentro da pasta "Backups - Inventário de TI" no Drive de quem é dono do
// script, e manda pra lixeira os backups com mais de BACKUP_DIAS_GUARDADOS
// dias. Os backups têm as mesmas informações da planilha original,
// inclusive senhas — a pasta é privada por padrão; não compartilhe.
//
// Pra ligar: no editor do Apps Script, escolha a função instalarBackupDiario
// no menu de cima e clique em Executar (uma vez só). O Google vai pedir
// permissão pro Drive — é pra criar a pasta e os arquivos de backup.
const BACKUP_PASTA_NOME = 'Backups - Inventário de TI';
const BACKUP_PREFIXO_ARQUIVO = 'Backup Inventário TI ';
const BACKUP_DIAS_GUARDADOS = 30;

function instalarBackupDiario() {
  // Remove gatilhos antigos da mesma função, pra não duplicar se rodar de novo.
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'fazerBackupDiario') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('fazerBackupDiario').timeBased().everyDays(1).atHour(3).create();
  fazerBackupDiario(); // já faz o primeiro agora, pra conferir que funciona
  Logger.log('Backup diário instalado (todo dia por volta das 3h). Primeiro backup feito na pasta "' + BACKUP_PASTA_NOME + '".');
}

function pastaDeBackup_() {
  const pastas = DriveApp.getFoldersByName(BACKUP_PASTA_NOME);
  return pastas.hasNext() ? pastas.next() : DriveApp.createFolder(BACKUP_PASTA_NOME);
}

function fazerBackupDiario() {
  const origem = SpreadsheetApp.getActiveSpreadsheet();
  const carimbo = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
  const copia = SpreadsheetApp.create(BACKUP_PREFIXO_ARQUIVO + carimbo);
  // A planilha nova já nasce com uma aba vazia ("Página1"/"Sheet1"); nome
  // temporário único pra não colidir com uma aba de mesmo nome da original.
  const abaVaziaInicial = copia.getSheets()[0].setName('__backup_temp_' + Date.now());
  origem.getSheets().forEach(function (aba) {
    aba.copyTo(copia).setName(aba.getName());
  });
  copia.deleteSheet(abaVaziaInicial);
  const pasta = pastaDeBackup_();
  DriveApp.getFileById(copia.getId()).moveTo(pasta);

  // Limpeza: só mexe em arquivos desta pasta com o nome de backup.
  const limite = new Date(Date.now() - BACKUP_DIAS_GUARDADOS * 24 * 60 * 60 * 1000);
  const arquivos = pasta.getFiles();
  let removidos = 0;
  while (arquivos.hasNext()) {
    const arquivo = arquivos.next();
    if (arquivo.getName().indexOf(BACKUP_PREFIXO_ARQUIVO) === 0 && arquivo.getDateCreated() < limite) {
      arquivo.setTrashed(true);
      removidos++;
    }
  }
  Logger.log('Backup criado: ' + copia.getName() + (removidos ? ' — ' + removidos + ' backup(s) antigo(s) enviado(s) pra lixeira.' : ''));
}
