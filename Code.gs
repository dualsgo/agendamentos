const NOME_PLANILHA = "Agendamentos";
const NOME_ABA = "Agenda";

// Configurações centralizadas
const CONFIG = {
  LIMITE_VOLS_PADRAO: 1000,
  STATUS_VALIDOS: ['PENDENTE', 'CONFIRMADO', 'CONCLUÍDO', 'CANCELADO', 'EM TRATATIVA'],
  STATUS_NOTIFICACAO: ['PENDENTE', 'EM TRATATIVA'],
  STATUS_CONFIRMADO: ['CONFIRMADO'],
  COLUNAS: {
    DATA_REAL: 1,
    FORNECEDOR: 2,
    NOTAS_FISCAIS: 3,
    VOLUMES: 4,
    DATA_AGENDADA: 5,
    OBSERVACOES: 6,
    STATUS: 7,
    ORIGEM: 8,
    ID: 9,
    REMETENTE: 10,
    ASSUNTO: 11,
    VOLS_RECEBIDOS: 12,
    RESUMO_DIRETO: 13
  }
};

// ==================== FUNÇÕES DE AUTORIZAÇÃO E CONFIGURAÇÃO ====================

function autorizarGmail() {
  const aliases = GmailApp.getAliases();
  console.log("Gmail autorizado com sucesso! Aliases encontrados: " + aliases.length);
}

function getSettings() {
  const props = PropertiesService.getScriptProperties();
  return {
    limite_vols: Number(props.getProperty('limite_vols')) || CONFIG.LIMITE_VOLS_PADRAO,
    data_inventario: props.getProperty('data_inventario') || ""
  };
}

function salvarSettings(s) {
  const props = PropertiesService.getScriptProperties();
  props.setProperty('limite_vols', s.limite_vols);
  props.setProperty('data_inventario', s.data_inventario);
  return "OK";
}

// ==================== FUNÇÕES DE FORMATAÇÃO DE DATAS ====================

function formatarDataStudio(valorData) {
  if (!valorData) return "";

  if (valorData instanceof Date) {
    return Utilities.formatDate(valorData, Session.getScriptTimeZone(), "dd/MM/yyyy");
  }

  let str = String(valorData).trim();

  const meses = {
    "janeiro": "01", "january": "01", "jan": "01",
    "fevereiro": "02", "february": "02", "fev": "02", "feb": "02",
    "março": "03", "marco": "03", "march": "03", "mar": "03",
    "abril": "04", "april": "04", "abr": "04", "apr": "04",
    "maio": "05", "may": "05",
    "junho": "06", "june": "06", "jun": "06",
    "julho": "07", "july": "07", "jul": "07",
    "agosto": "08", "august": "08", "ago": "08", "aug": "08",
    "setembro": "09", "september": "09", "set": "09", "sep": "09",
    "outubro": "10", "october": "10", "out": "10", "oct": "10",
    "novembro": "11", "november": "11", "nov": "11",
    "dezembro": "12", "december": "12", "dez": "12", "dec": "12"
  };

  const regexStudio = /([A-Za-zçÇ]+)\s+(\d{1,2}),?\s+(\d{4})/i;
  const match = str.match(regexStudio);

  if (match) {
    const nomeMês = match[1].toLowerCase();
    const dia = match[2].padStart(2, '0');
    const ano = match[3];
    const mesNum = meses[nomeMês];

    if (mesNum) {
      return `${dia}/${mesNum}/${ano}`;
    }
  }

  const parsedDate = new Date(str);
  if (!isNaN(parsedDate.getTime())) {
    return Utilities.formatDate(parsedDate, Session.getScriptTimeZone(), "dd/MM/yyyy");
  }

  return str;
}

function parseSafeDate(str) {
  if (!str || str === "—") return null;
  let s = String(str).trim();
  
  const formats = [
    { regex: /^(\d{2})\/(\d{2})\/(\d{4})$/, map: (d,m,y) => new Date(y, m-1, d) },
    { regex: /^(\d{4})-(\d{2})-(\d{2})$/, map: (y,m,d) => new Date(y, m-1, d) },
    { regex: /^(\d{2})-(\d{2})-(\d{4})$/, map: (d,m,y) => new Date(y, m-1, d) }
  ];
  
  for (const format of formats) {
    const match = s.match(format.regex);
    if (match) {
      const [, ...parts] = match.map(Number);
      const date = format.map(...parts);
      return isNaN(date.getTime()) ? null : date;
    }
  }
  
  const dataFormatada = formatarDataStudio(s);
  if (dataFormatada && dataFormatada !== s) {
    return parseSafeDate(dataFormatada);
  }
  
  const d = new Date(s);
  return isNaN(d.getTime()) ? null : d;
}

function getTargetRow(id, data) {
  const searchId = String(id).trim();
  for (let i = 1; i < data.length; i++) { 
    if (String(data[i][8]).trim() === searchId) return i + 1; 
  }
  return -1;
}

// ==================== FUNÇÕES DE LOG ====================

function logAcao(acao, dados, resultado) {
  const props = PropertiesService.getScriptProperties();
  let log = [];
  try {
    const logStr = props.getProperty('log');
    if (logStr) log = JSON.parse(logStr);
  } catch (e) {
    log = [];
  }
  
  log.push({
    timestamp: new Date().toISOString(),
    acao,
    dados: JSON.stringify(dados).substring(0, 500),
    resultado
  });
  
  if (log.length > 100) log.shift();
  props.setProperty('log', JSON.stringify(log));
}

function registrarAlteracao(rowIndex, campo, valorAntigo, valorNovo, usuario) {
  const props = PropertiesService.getScriptProperties();
  let historico = [];
  
  try {
    const hist = props.getProperty('historico_alteracoes');
    if (hist) historico = JSON.parse(hist);
  } catch (e) {
    historico = [];
  }
  
  historico.push({
    timestamp: new Date().toISOString(),
    rowIndex,
    campo,
    valorAntigo,
    valorNovo,
    usuario: usuario || 'Sistema'
  });
  
  if (historico.length > 500) {
    historico = historico.slice(-500);
  }
  
  props.setProperty('historico_alteracoes', JSON.stringify(historico));
}

// ==================== FUNÇÕES DE PLANILHA ====================

function getSpreadsheet() {
  try {
    let ss = SpreadsheetApp.getActiveSpreadsheet();
    if (ss && (ss.getSheetByName(NOME_ABA) || ss.getSheetByName("Agendamentos"))) return ss;

    const files = DriveApp.getFilesByName(NOME_PLANILHA);
    if (files.hasNext()) {
      ss = SpreadsheetApp.open(files.next());
      if (ss) return ss;
    }

    const filesAll = DriveApp.getFilesByType(MimeType.GOOGLE_SHEETS);
    while (filesAll.hasNext()) {
      const candidate = SpreadsheetApp.open(filesAll.next());
      if (candidate && (candidate.getSheetByName("Agenda") || candidate.getSheetByName("Agendamentos"))) {
        return candidate;
      }
    }
    return null;
  } catch (e) {
    console.error("Erro ao acessar planilha:", e);
    return null;
  }
}

function inicializarPlanilha() {
  try {
    const ss = getSpreadsheet();
    if (!ss) return null;
    
    let sheet = ss.getSheetByName("Agenda");
    
    if (!sheet) {
      sheet = ss.getSheetByName("Agendamentos");
      if (!sheet) {
        sheet = ss.insertSheet("Agenda");
        sheet.appendRow(["Data Real", "Fornecedor", "Notas Fiscais", "Volumes", "Data Agendada", "Observações", "Status", "Origem", "ID", "Remetente", "Assunto", "Volumes Recebidos", "Resumo Direto"]);
      }
    }
    
    const headers = sheet.getRange(1, 1, 1, 13).getValues()[0];
    if (!headers[8] || headers[8] !== "ID") sheet.getRange(1, 9).setValue("ID");
    if (!headers[11] || headers[11] !== "Volumes Recebidos") sheet.getRange(1, 12).setValue("Volumes Recebidos");
    if (!headers[12] || headers[12] !== "Resumo Direto") sheet.getRange(1, 13).setValue("Resumo Direto");
    
    let diffSheet = ss.getSheetByName("Divergências");
    if (!diffSheet) {
      diffSheet = ss.insertSheet("Divergências");
      diffSheet.appendRow(["Data Registro", "Código SAP", "Volumes Faltantes", "Peças Faltantes", "NF", "Fornecedor"]);
    }
    
    return sheet;
  } catch (e) { 
    throw new Error("Erro ao acessar planilha: " + e.message); 
  }
}

// ==================== FUNÇÃO PRINCIPAL COM CONSOLIDAÇÃO POR ID ====================

function getTodosOsDados() {
  const lock = LockService.getScriptLock();
  try { lock.waitLock(15000); } catch (e) { 
    return { alertas: [], agendamentos: [], divergencias: [], erro: "Planilha ocupada. Tente novamente." }; 
  }
  
  try {
    const ss = getSpreadsheet();
    if (!ss) return { alertas: [], agendamentos: [], divergencias: [], erro: "Planilha não encontrada." };
    
    const sheet = inicializarPlanilha();
    if (!sheet) return { alertas: [], agendamentos: [], divergencias: [], erro: "Aba principal não encontrada." };
    
    const rawData = sheet.getDataRange().getValues();
    if (rawData.length <= 1) return { alertas: [], agendamentos: [], divergencias: [] };

    const headers = rawData[0].map(h => String(h).trim());
    const rows = rawData.slice(1);
    
    // ============ NOVA LÓGICA DE CONSOLIDAÇÃO POR ID ============
    // 1. Primeiro, agrupa todas as linhas por ID
    const gruposPorId = new Map();
    const linhasSemId = [];
    
    for (let i = 0; i < rows.length; i++) {
      const linha = rows[i];
      if (!linha || linha.length < 3 || !linha[1]) continue;
      
      let id = linha[8] ? String(linha[8]).trim() : "";
      if (!id) {
        // Se não tem ID, gera um novo
        id = Utilities.getUuid();
        linhasSemId.push({ index: i, row: linha, id: id });
        continue;
      }
      
      if (!gruposPorId.has(id)) {
        gruposPorId.set(id, []);
      }
      gruposPorId.get(id).push({ index: i, row: linha });
    }

    // 2. Processa cada grupo, mantendo apenas o registro mais recente
    const registrosConsolidados = [];
    const idsParaGravar = [];
    let precisaGravar = false;

    // Processa grupos com ID existente
    for (const [id, linhas] of gruposPorId) {
      // Ordena as linhas por data de criação (assumindo que a última é a mais recente)
      // Usamos o índice da linha como proxy para ordem de inserção
      linhas.sort((a, b) => a.index - b.index);
      
      // Pega a última linha (mais recente)
      const linhaMaisRecente = linhas[linhas.length - 1];
      const linha = linhaMaisRecente.row;
      const index = linhaMaisRecente.index;
      
      // Se houver mais de uma linha com o mesmo ID, marca para limpeza
      if (linhas.length > 1) {
        precisaGravar = true;
        // Marca as linhas antigas para exclusão (opcional)
        // Podemos deixar para uma limpeza periódica ou apenas ignorar
        logAcao('getTodosOsDados', { id, totalLinhas: linhas.length }, 'Multiplas linhas encontradas, usando a mais recente');
      }
      
      const registro = criarRegistro(linha, index + 2, id);
      if (registro) {
        registrosConsolidados.push(registro);
      }
      idsParaGravar.push([id]);
    }

    // Processa linhas sem ID (gerar novos IDs)
    for (const item of linhasSemId) {
      const registro = criarRegistro(item.row, item.index + 2, item.id);
      if (registro) {
        registrosConsolidados.push(registro);
      }
      idsParaGravar.push([item.id]);
      precisaGravar = true;
    }

    // 3. Se houver IDs novos para gravar
    if (precisaGravar && idsParaGravar.length > 0) {
      // Atualiza apenas as linhas que precisam de ID
      for (let i = 0; i < idsParaGravar.length; i++) {
        const rowIndex = i + 2; // +2 por causa do cabeçalho e 0-based
        sheet.getRange(rowIndex, 9).setValue(idsParaGravar[i][0]);
      }
    }

    // 4. Separa entre alertas e agendamentos baseado no status
    const alertas = [];
    const agendamentos = [];

    for (const registro of registrosConsolidados) {
      const statusAtual = registro.status;
      
      // Lógica de classificação baseada no status
      if (statusAtual === 'PENDENTE' || statusAtual === 'EM TRATATIVA') {
        alertas.push(registro);
      } else if (statusAtual === 'CONFIRMADO') {
        agendamentos.push(registro);
      } else {
        // Outros status (CONCLUÍDO, RECEBIDO, CANCELADO, etc)
        agendamentos.push(registro);
      }
    }

    // 5. Processa divergências (não mudou)
    const diffSheet = ss.getSheetByName("Divergências");
    const diffData = (diffSheet && diffSheet.getLastRow() > 1) ? diffSheet.getRange(2, 1, diffSheet.getLastRow() - 1, diffSheet.getLastColumn()).getValues() : [];
    
    const divergencias = diffData.map((r, index) => ({
       rowIndex: index + 2,
       data: formatarDataStudio(r[0]),
       sap: String(r[1] || ""), 
       vols: Number(r[2]) || 0, 
       pecas: Number(r[3]) || 0, 
       nf: String(r[4] || ""), 
       fornecedor: String(r[5] || "")
    }));

    SpreadsheetApp.flush();
    
    // Ordena agendamentos por data
    agendamentos.sort((a, b) => {
      const d1 = parseSafeDate(a.data_agendada || a.data_real) || new Date('2099-01-01');
      const d2 = parseSafeDate(b.data_agendada || b.data_real) || new Date('2099-01-01');
      return d1 - d2;
    });

    return { alertas: alertas, agendamentos: agendamentos, divergencias: divergencias };
  } catch (e) { 
    logAcao('getTodosOsDados', {}, 'ERRO: ' + e.message);
    return { alertas: [], agendamentos: [], divergencias: [], erro: "ERRO CRÍTICO: " + e.message }; 
  } finally { lock.releaseLock(); }
}

/**
 * Função auxiliar para criar um registro a partir de uma linha
 */
function criarRegistro(linha, rowIndex, id) {
  if (!linha || linha.length < 3) return null;
  
  const origem = linha[7] ? String(linha[7]).trim() : "";
  const statusAtual = String(linha[6] || "PENDENTE").trim().toUpperCase();
  
  return {
    rowIndex: rowIndex,
    id: id,
    data_real: formatarDataStudio(linha[0]),
    fornecedor: String(linha[1] || "").trim(),
    notas_fiscais: String(linha[2] || "").trim(),
    volumes: Number(linha[3]) || 0,
    data_agendada: formatarDataStudio(linha[4]),
    observacoes: String(linha[5] || "").trim(),
    status: statusAtual,
    origem: origem,
    remetente: String(linha[9] || "").trim(),
    assunto: String(linha[10] || "").trim(),
    vols_recebidos: Number(linha[11]) || 0,
    resumo_direto: String(linha[12] || "").trim(),
    data_sugerida: formatarDataStudio(linha[4]),
    observacao: String(linha[5] || "").trim(),
    gmailUrl: id ? `https://mail.google.com/mail/u/0/#all/${id}` : "#"
  };
}

// ==================== FUNÇÕES DE LIMPEZA DE DUPLICATAS ====================

/**
 * Função para limpar registros duplicados (pode ser executada manualmente)
 */
function limparDuplicatas() {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
    
    const ss = getSpreadsheet();
    if (!ss) return "ERRO: Planilha não encontrada.";
    
    const sheet = ss.getSheetByName("Agenda");
    if (!sheet) return "ERRO: Aba Agenda não encontrada.";
    
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return "Nenhum dado para processar.";
    
    const idsMap = new Map();
    const linhasParaRemover = [];
    
    // Identifica duplicatas (primeira passagem)
    for (let i = 1; i < data.length; i++) {
      const linha = data[i];
      const id = linha[8] ? String(linha[8]).trim() : "";
      if (!id) continue;
      
      if (!idsMap.has(id)) {
        idsMap.set(id, []);
      }
      idsMap.get(id).push(i);
    }
    
    // Marca linhas duplicadas para remoção (mantém a mais recente)
    for (const [id, indices] of idsMap) {
      if (indices.length > 1) {
        // Ordena do mais antigo para o mais novo
        indices.sort((a, b) => a - b);
        // Remove todos exceto o último (mais recente)
        for (let j = 0; j < indices.length - 1; j++) {
          linhasParaRemover.push(indices[j] + 1); // +1 porque é 1-based
        }
      }
    }
    
    // Remove as linhas (do final para o início para não bagunçar os índices)
    linhasParaRemover.sort((a, b) => b - a);
    let removidos = 0;
    for (const rowIndex of linhasParaRemover) {
      sheet.deleteRow(rowIndex);
      removidos++;
    }
    
    SpreadsheetApp.flush();
    logAcao('limparDuplicatas', {}, `Removidos ${removidos} registros duplicados`);
    return `OK - ${removidos} duplicatas removidas`;
  } catch (e) {
    logAcao('limparDuplicatas', {}, 'ERRO: ' + e.message);
    return "ERRO: " + e.message;
  } finally {
    lock.releaseLock();
  }
}

// ==================== FUNÇÕES DE ATUALIZAÇÃO DE STATUS ====================

function atualizarStatusAgendamento(rowIndex, novoStatus, novaDataSugerida) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    
    const ss = getSpreadsheet();
    if (!ss) return "ERRO: Planilha não encontrada.";
    
    const sheet = ss.getSheetByName("Agenda");
    if (!sheet) return "ERRO: Aba Agenda não encontrada.";
    
    const currentData = sheet.getRange(rowIndex, 1, 1, sheet.getLastColumn()).getValues()[0];
    const oldStatus = currentData[6] || 'PENDENTE';
    
    sheet.getRange(rowIndex, 7).setValue(novoStatus);
    
    if (novaDataSugerida) {
      sheet.getRange(rowIndex, 5).setValue(novaDataSugerida);
    }
    
    registrarAlteracao(rowIndex, 'status', oldStatus, novoStatus, 'Sistema');
    if (novaDataSugerida) {
      const oldData = currentData[4] || '';
      registrarAlteracao(rowIndex, 'data_agendada', oldData, novaDataSugerida, 'Sistema');
    }
    
    SpreadsheetApp.flush();
    logAcao('atualizarStatusAgendamento', { rowIndex, novoStatus, novaDataSugerida }, 'OK');
    return "OK";
  } catch (e) {
    logAcao('atualizarStatusAgendamento', { rowIndex, novoStatus, novaDataSugerida }, 'ERRO: ' + e.message);
    return "ERRO INTERNO: " + e.message;
  } finally {
    lock.releaseLock();
  }
}

function confirmarAgendamento(rowIndex, dataConfirmada) {
  return atualizarStatusAgendamento(rowIndex, 'CONFIRMADO', dataConfirmada);
}

function solicitarAlteracaoData(rowIndex, novaDataSugerida) {
  return atualizarStatusAgendamento(rowIndex, 'EM TRATATIVA', novaDataSugerida);
}

// ==================== FUNÇÕES DE OPERAÇÕES EXISTENTES ====================

function atualizarAgendamento(rowIndex, dados) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    
    const ss = getSpreadsheet();
    if (!ss) return "ERRO: Planilha não encontrada.";
    
    const sheet = ss.getSheetByName("Agenda");
    if (!sheet) return "ERRO: Aba Agenda não encontrada.";
    
    const currentData = sheet.getRange(rowIndex, 1, 1, sheet.getLastColumn()).getValues()[0];
    
    const mapeamento = {
      'status': 7,
      'data_agendada': 5,
      'data_sugerida': 5,
      'observacoes': 6,
      'observacao': 6,
      'volumes': 4,
      'notas_fiscais': 3,
      'fornecedor': 2,
      'vols_recebidos': 12,
      'resumo_direto': 13
    };
    
    for (const [campo, valor] of Object.entries(dados)) {
      if (campo in mapeamento && valor !== undefined && valor !== null && valor !== '') {
        const coluna = mapeamento[campo];
        let valorFinal = valor;
        let valorAntigo = currentData[coluna - 1] || '';
        
        if (campo === 'data_agendada' || campo === 'data_sugerida') {
          const dataObj = parseSafeDate(valor);
          if (dataObj) {
            valorFinal = Utilities.formatDate(dataObj, Session.getScriptTimeZone(), "dd/MM/yyyy");
          }
          valorAntigo = formatarDataStudio(valorAntigo);
        }
        
        if (String(valorAntigo) !== String(valor)) {
          registrarAlteracao(rowIndex, campo, valorAntigo, valor, 'Sistema');
        }
        
        sheet.getRange(rowIndex, coluna).setValue(valorFinal);
      }
    }
    
    SpreadsheetApp.flush();
    logAcao('atualizarAgendamento', { rowIndex, dados }, 'OK');
    return "OK";
  } catch (e) {
    logAcao('atualizarAgendamento', { rowIndex, dados }, 'ERRO: ' + e.message);
    return "ERRO INTERNO: " + e.message;
  } finally {
    lock.releaseLock();
  }
}

function atualizarStatus(rowIndex, novoStatus) {
  if (!CONFIG.STATUS_VALIDOS.includes(novoStatus)) {
    return "ERRO: Status inválido. Use: " + CONFIG.STATUS_VALIDOS.join(', ');
  }
  return atualizarAgendamento(rowIndex, { status: novoStatus });
}

function atualizarDataSugerida(rowIndex, novaData) {
  const dataObj = parseSafeDate(novaData);
  if (!dataObj) {
    return "ERRO: Formato de data inválido. Use AAAA-MM-DD ou DD/MM/AAAA";
  }
  const dataFormatada = Utilities.formatDate(dataObj, Session.getScriptTimeZone(), "dd/MM/yyyy");
  return atualizarAgendamento(rowIndex, { data_sugerida: dataFormatada });
}

function responderEmail(id, mensagem) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    if (!id || !mensagem) return "ERRO: ID ou mensagem ausente.";
    const thread = GmailApp.getThreadById(id.trim());
    if (!thread) return "ERRO: E-mail não encontrado no Gmail.";
    
    thread.reply(mensagem);
    logAcao('responderEmail', { id }, 'OK');
    return "OK";
  } catch (e) {
    logAcao('responderEmail', { id }, 'ERRO: ' + e.message);
    return "ERRO INTERNO: " + e.message;
  } finally {
    lock.releaseLock();
  }
}

function salvarAgendamento(dados, enviarEmail = false, corpoEmail = "") {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const sheet = inicializarPlanilha();
    const data = sheet.getDataRange().getValues();
    
    let targetRow = getTargetRow(dados.id, data);
    const idUnico = dados.id || Utilities.getUuid();
    
    let dtAgendada = null;
    if (dados.data_agendada) {
      const dataObj = parseSafeDate(dados.data_agendada);
      if (dataObj) {
        dtAgendada = dataObj;
      } else {
        dtAgendada = new Date();
      }
    } else {
      dtAgendada = new Date();
    }

    if (targetRow > -1) {
      const updateData = {
        fornecedor: dados.fornecedor,
        notas_fiscais: dados.notas_fiscais,
        volumes: dados.volumes,
        data_agendada: dados.data_agendada,
        observacoes: dados.observacoes,
        status: dados.status || "PENDENTE"
      };
      if (dados.remetente) updateData.remetente = dados.remetente;
      if (dados.assunto) updateData.assunto = dados.assunto;
      if (dados.resumo_direto) updateData.resumo_direto = dados.resumo_direto;
      
      const currentData = sheet.getRange(targetRow, 1, 1, sheet.getLastColumn()).getValues()[0];
      const mapeamento = {
        'status': 7,
        'data_agendada': 5,
        'data_sugerida': 5,
        'observacoes': 6,
        'observacao': 6,
        'volumes': 4,
        'notas_fiscais': 3,
        'fornecedor': 2,
        'remetente': 10,
        'assunto': 11,
        'vols_recebidos': 12,
        'resumo_direto': 13
      };
      
      for (const [campo, valor] of Object.entries(updateData)) {
        if (valor !== undefined && valor !== null && valor !== '') {
          if (campo in mapeamento) {
            const coluna = mapeamento[campo];
            let valorFinal = valor;
            let valorAntigo = currentData[coluna - 1] || '';
            
            if (campo === 'data_agendada' || campo === 'data_sugerida') {
              const dataObj = parseSafeDate(valor);
              if (dataObj) {
                valorFinal = Utilities.formatDate(dataObj, Session.getScriptTimeZone(), "dd/MM/yyyy");
              }
              valorAntigo = formatarDataStudio(valorAntigo);
            }
            
            if (String(valorAntigo) !== String(valor)) {
              registrarAlteracao(targetRow, campo, valorAntigo, valor, 'Sistema');
            }
            
            sheet.getRange(targetRow, coluna).setValue(valorFinal);
          }
        }
      }
    } else {
      const dataFormatada = Utilities.formatDate(dtAgendada, Session.getScriptTimeZone(), "dd/MM/yyyy");
      const rowData = [
        "", dados.fornecedor, dados.notas_fiscais, dados.volumes,
        dataFormatada, dados.observacoes, dados.status || "PENDENTE", "Formulário Web", idUnico, dados.remetente || "", dados.assunto || "", dados.vols_recebidos || 0, dados.resumo_direto || ""
      ];
      sheet.appendRow(rowData);
    }

    SpreadsheetApp.flush();
    logAcao('salvarAgendamento', { dados }, 'OK');
    return "OK";
  } catch (e) { 
    logAcao('salvarAgendamento', { dados }, 'ERRO: ' + e.message);
    return "ERRO INTERNO: " + e.message; 
  } finally { lock.releaseLock(); }
}

function registrarCheckinComDivergencia(id, dados) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    const ss = getSpreadsheet();
    const sheet = inicializarPlanilha();
    const diffSheet = ss.getSheetByName("Divergências");
    const data = sheet.getDataRange().getValues();
    
    let targetRow = getTargetRow(id, data);
    if (targetRow > -1) {
      const oldStatus = sheet.getRange(targetRow, 7).getValue();
      const oldVols = sheet.getRange(targetRow, 12).getValue();
      
      const dataAtual = new Date();
      
      sheet.getRange(targetRow, 1).setValue(dataAtual); 
      sheet.getRange(targetRow, 7).setValue("CONCLUÍDO");
      sheet.getRange(targetRow, 12).setValue(Number(dados.volumes_recebidos));
      
      registrarAlteracao(targetRow, 'status', oldStatus, 'CONCLUÍDO', 'Checkin');
      registrarAlteracao(targetRow, 'vols_recebidos', oldVols, dados.volumes_recebidos, 'Checkin');
      
      if (dados.itens && Array.isArray(dados.itens)) {
         dados.itens.forEach(item => {
            if (Number(item.faltam_vols) > 0 || Number(item.faltam_pecas) > 0) {
               diffSheet.appendRow([
                  dataAtual,
                  item.sap || "N/A",
                  Number(item.faltam_vols) || 0,
                  Number(item.faltam_pecas) || 0,
                  dados.nf || "",
                  dados.fornecedor || ""
               ]);
            }
         });
      }
      SpreadsheetApp.flush();
      logAcao('registrarCheckinComDivergencia', { id, dados }, 'OK');
      return "OK";
    }
    return "NÃO ENCONTRADO";
  } catch (e) { 
    logAcao('registrarCheckinComDivergencia', { id, dados }, 'ERRO: ' + e.message);
    return "ERRO: " + e.message; 
  } finally { lock.releaseLock(); }
}

function marcarRecebidosLote(idsStr, dataRealString) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const sheet = inicializarPlanilha();
    const data = sheet.getDataRange().getValues();
    
    let idsArray = [];
    try {
      idsArray = JSON.parse(idsStr);
    } catch (e) {
      idsArray = idsStr.split(",");
    }
    
    let dtReal = new Date();
    if (dataRealString) {
      const dataObj = parseSafeDate(dataRealString);
      if (dataObj) dtReal = dataObj;
    }
    
    let processados = 0;

    for (let id of idsArray) {
      let targetRow = getTargetRow(id, data);
      if (targetRow > -1) {
        const oldStatus = sheet.getRange(targetRow, 7).getValue();
        const oldVolsRec = sheet.getRange(targetRow, 12).getValue();
        
        sheet.getRange(targetRow, 1).setValue(dtReal); 
        sheet.getRange(targetRow, 7).setValue("CONCLUÍDO");
        
        registrarAlteracao(targetRow, 'status', oldStatus, 'CONCLUÍDO', 'Recebimento Lote');
        
        const vRec = sheet.getRange(targetRow, 12).getValue();
        if (!vRec || vRec === 0) {
           const vTotal = sheet.getRange(targetRow, 4).getValue();
           sheet.getRange(targetRow, 12).setValue(vTotal);
           registrarAlteracao(targetRow, 'vols_recebidos', oldVolsRec, vTotal, 'Recebimento Lote');
        }
        processados++;
      }
    }
    SpreadsheetApp.flush();
    logAcao('marcarRecebidosLote', { total: idsArray.length, processados }, 'OK');
    return "OK - " + processados + " processados";
  } catch (e) { 
    logAcao('marcarRecebidosLote', { idsStr }, 'ERRO: ' + e.message);
    return "ERRO INTERNO: " + e.message; 
  } finally { lock.releaseLock(); }
}

function aprovarLote(idsStr) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const sheet = inicializarPlanilha();
    const data = sheet.getDataRange().getValues();
    
    let idsArray = [];
    try {
      idsArray = JSON.parse(idsStr);
    } catch (e) {
      idsArray = idsStr.split(",");
    }
    
    let processados = 0;
    
    for (let id of idsArray) {
      let targetRow = getTargetRow(id, data);
      if (targetRow > -1) {
        const currentData = sheet.getRange(targetRow, 1, 1, sheet.getLastColumn()).getValues()[0];
        const oldStatus = currentData[6] || 'PENDENTE';
        let suggestedDate = currentData[4];
        
        // Se a data de agendamento não for válida, e tivermos que definir
        if (!suggestedDate) {
           suggestedDate = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy");
           sheet.getRange(targetRow, 5).setValue(suggestedDate);
        }
        
        sheet.getRange(targetRow, 7).setValue("CONFIRMADO");
        registrarAlteracao(targetRow, 'status', oldStatus, 'CONFIRMADO', 'Aprovação Lote');
        
        processados++;
      }
    }
    SpreadsheetApp.flush();
    logAcao('aprovarLote', { total: idsArray.length, processados }, 'OK');
    return "OK - " + processados + " processados";
  } catch (e) { 
    logAcao('aprovarLote', { idsStr }, 'ERRO: ' + e.message);
    return "ERRO INTERNO: " + e.message; 
  } finally { lock.releaseLock(); }
}

function excluirAlertaFlow(id) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const sheet = inicializarPlanilha();
    const data = sheet.getDataRange().getValues();
    let targetRow = getTargetRow(id, data);
    if (targetRow > -1) { 
      sheet.deleteRow(targetRow); 
      SpreadsheetApp.flush();
      logAcao('excluirAlertaFlow', { id }, 'OK');
      return "OK"; 
    }
    return "NÃO ENCONTRADO";
  } catch (e) { 
    logAcao('excluirAlertaFlow', { id }, 'ERRO: ' + e.message);
    return "ERRO INTERNO: " + e.message; 
  } finally { lock.releaseLock(); }
}

function excluirDivergencia(index) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const ss = getSpreadsheet();
    if (!ss) return "ERRO: Planilha não encontrada.";
    const sheet = ss.getSheetByName("Divergências");
    if (!sheet) return "ERRO: Aba divergências não encontrada.";
    
    const rowToDelete = Number(index) + 2;
    sheet.deleteRow(rowToDelete);
    SpreadsheetApp.flush();
    logAcao('excluirDivergencia', { index }, 'OK');
    return "OK";
  } catch (e) { 
    logAcao('excluirDivergencia', { index }, 'ERRO: ' + e.message);
    return "ERRO INTERNO: " + e.message; 
  } finally { lock.releaseLock(); }
}

// ==================== FUNÇÕES PRINCIPAIS (WEB APP & API REST) ====================

function doGet(e) {
  // Se for acesso direto sem parâmetro action, abre a página web interna
  if (!e || !e.parameter || !e.parameter.action) {
    return HtmlService.createHtmlOutputFromFile("index")
      .setTitle("Ri Happy | Estoque e Agendamentos")
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no');
  }

  // Se for chamada via Vercel (API REST JSON)
  const action = e.parameter.action;
  let payload = {};

  try {
    if (action === 'getTodosOsDados') {
      payload = getTodosOsDados();
    } else if (action === 'getSettings') {
      payload = getSettings();
    } else {
      payload = { erro: "Ação GET inválida: " + action };
    }
  } catch (err) {
    payload = { erro: err.message };
  }

  return ContentService.createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  let payload = {};
  try {
    let data = {};
    if (e && e.postData && e.postData.contents) {
      data = JSON.parse(e.postData.contents);
    }
    const action = data.action || (e && e.parameter ? e.parameter.action : '');
    const args = data.args || [];

    if (action === 'salvarAgendamento') {
      payload = salvarAgendamento(args[0]);
    } else if (action === 'atualizarStatus') {
      payload = atualizarStatus(args[0], args[1]);
    } else if (action === 'atualizarDataSugerida') {
      payload = atualizarDataSugerida(args[0], args[1]);
    } else if (action === 'registrarCheckinComDivergencia') {
      payload = registrarCheckinComDivergencia(args[0], args[1]);
    } else if (action === 'marcarRecebidosLote') {
      payload = marcarRecebidosLote(args[0], args[1]);
    } else if (action === 'excluirAlertaFlow' || action === 'excluir') {
      payload = excluirAlertaFlow(args[0]);
    } else if (action === 'excluirDivergencia') {
      payload = excluirDivergencia(args[0]);
    } else if (action === 'salvarSettings') {
      payload = salvarSettings(args[0]);
    } else if (action === 'getTodosOsDados') {
      payload = getTodosOsDados();
    } else if (action === 'getSettings') {
      payload = getSettings();
    } else {
      payload = "ERRO: Ação POST desconhecida: " + action;
    }
  } catch (err) {
    payload = "ERRO: " + err.message;
  }

  return ContentService.createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}