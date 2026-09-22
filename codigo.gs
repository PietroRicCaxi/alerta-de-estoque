/**
 * ==================================================
 * PROJETO 2 — ALERTA DE ESTOQUE
 * ETAPA 1 — AUTENTICAÇÃO COM O BLING
 * ==================================================
 */

const PROP_CLIENT_ID = 'BLING_ESTOQUE_CLIENT_ID';
const PROP_CLIENT_SECRET = 'BLING_ESTOQUE_CLIENT_SECRET';

/**
 * Configura e retorna o serviço OAuth2 apontando para o Bling.
 */
function getBlingService() {
  const props = PropertiesService.getScriptProperties();

  return OAuth2.createService('bling-estoque')
    .setAuthorizationBaseUrl('https://www.bling.com.br/Api/v3/oauth/authorize')
    .setTokenUrl('https://api.bling.com.br/Api/v3/oauth/token')
    .setClientId(props.getProperty(PROP_CLIENT_ID))
    .setClientSecret(props.getProperty(PROP_CLIENT_SECRET))
    .setCallbackFunction('authCallback')
    .setPropertyStore(props)
    .setTokenHeaders({
      'Authorization': 'Basic ' + Utilities.base64Encode(
        props.getProperty(PROP_CLIENT_ID) + ':' + props.getProperty(PROP_CLIENT_SECRET)
      )
    });
}

/**
 * Rode manualmente para gerar o link de autorização (aparece no Registro de execução).
 */
function iniciarAutorizacao() {
  const service = getBlingService();

  if (service.hasAccess()) {
    Logger.log('Já estamos autenticados com o Bling. Nada a fazer.');
  } else {
    Logger.log('Abra este link no navegador para autorizar o acesso:');
    Logger.log(service.getAuthorizationUrl());
  }
}

/**
 * Chamada automaticamente pelo Bling após a autorização. Não rode manualmente.
 */
function authCallback(request) {
  const service = getBlingService();
  const isAuthorized = service.handleCallback(request);

  return HtmlService.createHtmlOutput(
    isAuthorized
      ? 'Autorização concluída com sucesso! Pode fechar esta aba.'
      : 'Falha na autorização. Volte ao Apps Script e tente novamente.'
  );
}

/**
 * Confirma que a conexão com o Bling está funcionando.
 */
function testarConexao() {
  const service = getBlingService();

  if (!service.hasAccess()) {
    Logger.log('Ainda não autorizado. Rode "iniciarAutorizacao" primeiro.');
    return;
  }

  Logger.log('Conectado com sucesso ao Bling!');
  Logger.log('Token (início): ' + service.getAccessToken().substring(0, 10) + '...');
}

/**
 * Desconecta o script do Bling, caso precise reautorizar do zero.
 */
function resetarAutorizacao() {
  getBlingService().reset();
  Logger.log('Autorização removida. Rode "iniciarAutorizacao" para reconectar.');
}

/**
 * Página com um botão clicável para autorizar o acesso.
 */
function doGet() {
  const service = getBlingService();

  if (service.hasAccess()) {
    return HtmlService.createHtmlOutput(
      '<p>Já estamos autenticados com o Bling. Nenhuma ação necessária.</p>'
    );
  }

  const authorizationUrl = service.getAuthorizationUrl();
  return HtmlService.createHtmlOutput(
    '<p>Clique no botão abaixo para autorizar o acesso ao Bling:</p>' +
    '<a href="' + authorizationUrl + '" target="_blank" ' +
    'style="display:inline-block;padding:12px 20px;background:#0b93f6;' +
    'color:white;text-decoration:none;border-radius:6px;font-family:sans-serif;">' +
    'Autorizar acesso ao Bling</a>'
  );
}

/**
 * ==================================================
 * MENU CUSTOMIZADO NA PLANILHA
 * ==================================================
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Alerta de Estoque')
    .addItem('Atualizar Estoque', 'atualizarEstoqueDinamico')
    .addSeparator()
    .addItem('Autorizar conexão com o Bling', 'iniciarAutorizacao')
    .addToUi();
}

/**
 * ==================================================
 * ORDENAÇÃO — cálculo do número P real
 * ==================================================
 */
function calcularPNumber(codigo) {
  codigo = String(codigo).trim();
  const padroes = [
    /^CP(\d+)-EP(\d+)-?[A-Za-z]*$/i,
    /^CP(\d+)E-P(\d+)-?[A-Za-z]*$/i,
    /^CP(\d+)EP(\d+)[A-Za-z]*$/i,
    /^CP(\d+)\s*-\s*EP(\d+)\s*-?\s*[A-Za-z]*$/i,
    /^CP(\d+)-?(\d+)EP-?[A-Za-z]*$/i,
    /^CPF(\d+)-EP(\d+)-?[A-Za-z]*$/i,
    /^CPT(\d+)-EP(\d+)-?[A-Za-z]*$/i
  ];
  for (const pat of padroes) {
    const m = codigo.match(pat);
    if (m) return parseInt(m[1] + m[2], 10);
  }
  const direto = codigo.match(/^CP[FT]?(\d+)$/i);
  if (direto) return parseInt(direto[1], 10);

  const nums = codigo.match(/\d+/g);
  return nums ? parseInt(nums.join(''), 10) : 999999999;
}

/**
 * Reordena uma aba (CP, CPF ou CPT) pelo número P real, do menor pro maior.
 */
function ordenarAbaPorPNumber(aba) {
  const ultimaLinha = aba.getLastRow();
  if (ultimaLinha < 3) return;

  const intervalo = aba.getRange(2, 1, ultimaLinha - 1, 9);
  const dados = intervalo.getValues();

  dados.sort(function(a, b) {
    return calcularPNumber(a[0]) - calcularPNumber(b[0]);
  });

  intervalo.setValues(dados);
}

/**
 * Lê as abas Registro CP/CPF/CPT e monta, para cada código, as médias
 * diárias de saída em 3 janelas: Semana (7d), Mês (30d) e Trimestre (90d).
 * Cada janela só é calculada se houver pelo menos 1 dia de dado real
 * dentro dela; senão fica "null" (o chamador decide se mantém a
 * estimativa antiga ou marca como "Aguardando").
 */
function lerRegistrosParaMedias_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoje = new Date();
  hoje.setHours(0, 0, 0, 0);

  const limite7 = new Date(hoje); limite7.setDate(limite7.getDate() - 7);
  const limite30 = new Date(hoje); limite30.setDate(limite30.getDate() - 30);
  const limite90 = new Date(hoje); limite90.setDate(limite90.getDate() - 90);

  // código -> { semana: {total, dias:Set}, mes: {...}, trimestre: {...} }
  const mapa = {};

  ['Registro CP', 'Registro CPF', 'Registro CPT'].forEach(function(nomeAba) {
    const aba = ss.getSheetByName(nomeAba);
    if (!aba || aba.getLastRow() < 2) return;

    const valores = aba.getRange(2, 1, aba.getLastRow() - 1, 4).getValues(); // Data, Código, ID, Saídas

    valores.forEach(function(linha) {
      const dataLinha = new Date(linha[0]);
      const codigo = String(linha[1]).trim();
      const saidas = Number(linha[3]) || 0;
      const dataStr = String(linha[0]);

      if (!mapa[codigo]) {
        mapa[codigo] = {
          semana: { total: 0, dias: {} },
          mes: { total: 0, dias: {} },
          trimestre: { total: 0, dias: {} }
        };
      }

      if (dataLinha >= limite90) {
        mapa[codigo].trimestre.total += saidas;
        mapa[codigo].trimestre.dias[dataStr] = true;
      }
      if (dataLinha >= limite30) {
        mapa[codigo].mes.total += saidas;
        mapa[codigo].mes.dias[dataStr] = true;
      }
      if (dataLinha >= limite7) {
        mapa[codigo].semana.total += saidas;
        mapa[codigo].semana.dias[dataStr] = true;
      }
    });
  });

  const resultado = {};
  Object.keys(mapa).forEach(function(codigo) {
    const c = mapa[codigo];
    const diasSemana = Object.keys(c.semana.dias).length;
    const diasMes = Object.keys(c.mes.dias).length;
    const diasTrimestre = Object.keys(c.trimestre.dias).length;

    resultado[codigo] = {
      semana: diasSemana > 0 ? Math.round((c.semana.total / diasSemana) * 100) / 100 : null,
      mes: diasMes > 0 ? Math.round((c.mes.total / diasMes) * 100) / 100 : null,
      trimestre: diasTrimestre > 0 ? Math.round((c.trimestre.total / diasTrimestre) * 100) / 100 : null
    };
  });

  return resultado;
}

/**
 * ==================================================
 * ATUALIZAÇÃO DINÂMICA — lê direto da planilha
 * ==================================================
 * Lê a coluna Código de cada aba (CP/CPF/CPT), descobre o ID no Bling de
 * peças novas automaticamente, busca o saldo de todo mundo em lotes,
 * recalcula as 3 médias (Semana/Mês/Trimestre) usando o Registro quando
 * já tiver dado real disponível naquela janela — senão mantém a
 * estimativa antiga (dos relatórios manuais). "Dias Restantes" usa a
 * mais conservadora entre as médias disponíveis.
 *
 * Colunas esperadas: A=Código, B=Saldo Atual, C=Média/Semana, D=Média/Mês,
 * E=Média/Trimestre, F=Dias Restantes, G=Status, H=ID Bling, I=Preço Unitário
 */
function atualizarEstoqueDinamico() {
  const service = getBlingService();
  if (!service.hasAccess()) {
    Logger.log('Ainda não autorizado. Rode "iniciarAutorizacao" primeiro.');
    return;
  }

  const mediasDoRegistro = lerRegistrosParaMedias_();

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const abas = ['CP', 'CPF', 'CPT'];
  const LIMITE_BUSCAS_NOVAS = 200;
  let buscasNovasFeitas = 0;

  abas.forEach(function(nomeAba) {
    const aba = ss.getSheetByName(nomeAba);
    if (!aba) {
      Logger.log('Aba "' + nomeAba + '" não encontrada, pulando.');
      return;
    }

    const ultimaLinha = aba.getLastRow();
    if (ultimaLinha < 2) return;

    const dados = aba.getRange(2, 1, ultimaLinha - 1, 9).getValues();

    for (let i = 0; i < dados.length; i++) {
      const codigo = String(dados[i][0]).trim();
      let id = dados[i][7];

      if (codigo === '') continue;

      if (!id && buscasNovasFeitas < LIMITE_BUSCAS_NOVAS) {
        const urlBusca = 'https://api.bling.com.br/Api/v3/produtos?codigo=' + encodeURIComponent(codigo);
        const respBusca = UrlFetchApp.fetch(urlBusca, {
          headers: { 'Authorization': 'Bearer ' + service.getAccessToken() },
          muteHttpExceptions: true
        });
        Utilities.sleep(400);
        buscasNovasFeitas++;

        if (respBusca.getResponseCode() === 200) {
          const dadosBusca = JSON.parse(respBusca.getContentText());
          if (dadosBusca.data && dadosBusca.data.length > 0) {
            id = dadosBusca.data[0].id;
            dados[i][7] = id;
            if (dadosBusca.data[0].preco !== undefined) {
              dados[i][8] = dadosBusca.data[0].preco;
            }
            Logger.log('Peça nova encontrada: ' + codigo + ' -> ID ' + id);
          } else {
            Logger.log('Peça "' + codigo + '" não encontrada no Bling. Confira o código.');
          }
        }
      }

      // Substitui cada média pela versão do Registro, SE houver dado real
      // naquela janela específica. Senão, mantém o que já estava na célula
      // (estimativa antiga dos relatórios manuais, ou "Aguardando").
      const infoRegistro = mediasDoRegistro[codigo];
      if (infoRegistro) {
        if (infoRegistro.semana !== null) dados[i][2] = infoRegistro.semana;
        if (infoRegistro.mes !== null) dados[i][3] = infoRegistro.mes;
        if (infoRegistro.trimestre !== null) dados[i][4] = infoRegistro.trimestre;
      }
    }

    const colunaIds = dados.map(function(linha) { return [linha[7]]; });
    aba.getRange(2, 8, colunaIds.length, 1).setValues(colunaIds);

    const comId = dados.map(function(linha, idx) { return { idx: idx, id: linha[7] }; })
                        .filter(function(x) { return x.id; });

    const saldosPorId = {};
    const TAMANHO_LOTE = 50;
    for (let i = 0; i < comId.length; i += TAMANHO_LOTE) {
      const lote = comId.slice(i, i + TAMANHO_LOTE);
      const params = lote.map(function(x) { return 'idsProdutos[]=' + x.id; }).join('&');
      const url = 'https://api.bling.com.br/Api/v3/estoques/saldos?' + params;

      const resp = UrlFetchApp.fetch(url, {
        headers: { 'Authorization': 'Bearer ' + service.getAccessToken() },
        muteHttpExceptions: true
      });
      Utilities.sleep(400);

      if (resp.getResponseCode() === 200) {
        const data = JSON.parse(resp.getContentText());
        (data.data || []).forEach(function(item) {
          saldosPorId[String(item.produto.id)] = item.saldoVirtualTotal;
        });
      }
    }

    for (let i = 0; i < dados.length; i++) {
      const id = dados[i][7];
      const saldo = id ? saldosPorId[String(id)] : undefined;

      if (saldo !== undefined) {
        dados[i][1] = saldo;
      }

      const mediaSemana = Number(dados[i][2]) || 0;
      const mediaMes = Number(dados[i][3]) || 0;
      const mediaTrimestre = Number(dados[i][4]) || 0;
      const melhorMedia = Math.max(mediaSemana, mediaMes, mediaTrimestre);

      let diasRestantes = '';
      let status = '';
      const saldoFinal = dados[i][1];

      if (saldoFinal === '' || saldoFinal === undefined) {
        status = 'Sem dado de saldo';
      } else if (saldoFinal <= 0) {
        status = '🔴 ZERADO';
      } else if (melhorMedia === 0) {
        status = '⚪ Sem movimento nos últimos ~3 meses';
      } else {
        diasRestantes = Math.round((saldoFinal / melhorMedia) * 10) / 10;
        if (diasRestantes <= 15) status = '🔴 Crítico (<15 dias)';
        else if (diasRestantes <= 45) status = '🟡 Atenção (<45 dias)';
        else status = '🟢 Saudável';
      }

      dados[i][5] = diasRestantes;
      dados[i][6] = status;
    }

    aba.getRange(2, 1, dados.length, 9).setValues(dados);
    ordenarAbaPorPNumber(aba);

    Logger.log('Aba "' + nomeAba + '" atualizada: ' + dados.length + ' peças.');
  });

  Logger.log('Concluído! ' + buscasNovasFeitas + ' peças novas tiveram o ID buscado nesta execução.');
  if (buscasNovasFeitas >= LIMITE_BUSCAS_NOVAS) {
    Logger.log('Atenção: atingiu o limite de buscas novas nesta execução. Rode de novo para continuar processando o restante.');
  }

  atualizarConsultaDeMedia();
}

/**
 * ==================================================
 * CONSULTA DE MÉDIA — visão consolidada
 * ==================================================
 * Junta CP + CPF + CPT numa aba só, com preço, saldo e as 3 médias,
 * pronta pra colaboradora usar o filtro nativo do Google Sheets.
 * Chamada automaticamente no final de atualizarEstoqueDinamico.
 */
function atualizarConsultaDeMedia() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let aba = ss.getSheetByName('Consulta de Média');
  if (!aba) {
    aba = ss.insertSheet('Consulta de Média');
  } else {
    aba.clear();
  }

  aba.appendRow(['Código', 'Tipo', 'Preço Unitário', 'Saldo Atual',
    'Média/Semana', 'Média/Mês', 'Média/Trimestre', 'Dias Restantes', 'Status']);

  const linhasTotais = [];

  ['CP', 'CPF', 'CPT'].forEach(function(tipo) {
    const abaOrigem = ss.getSheetByName(tipo);
    if (!abaOrigem || abaOrigem.getLastRow() < 2) return;

    const dados = abaOrigem.getRange(2, 1, abaOrigem.getLastRow() - 1, 9).getValues();
    dados.forEach(function(linha) {
      const codigo = linha[0];
      if (!codigo) return;
      linhasTotais.push([
        codigo, tipo, linha[8], linha[1], linha[2], linha[3], linha[4], linha[5], linha[6]
      ]);
    });
  });

  if (linhasTotais.length > 0) {
    aba.getRange(2, 1, linhasTotais.length, 9).setValues(linhasTotais);
  }

  Logger.log('Consulta de Média atualizada: ' + linhasTotais.length + ' peças.');
}

/**
 * ==================================================
 * WEBHOOK — recebe avisos automáticos do Bling
 * ==================================================
 * Processa eventos "stock.created"/"stock.updated" (movimento real de
 * estoque físico). Ignora "virtual_stock.updated" (reflexo em produtos
 * compostos, não é uma venda/entrada real da peça simples).
 *
 * Grava cada evento numa aba "Buffer do Dia" (rápido). A consolidação de
 * verdade (1 linha por peça por dia) acontece à noite, em função separada.
 */
/**
 * Conjunto (Set) com todos os IDs de peças que rastreamos, construído
 * a partir do MAPA_PECAS. Usado pelo webhook pra filtrar rápido, sem
 * precisar de chamada à API.
 */
function idsRastreadosSet_() {
  const ids = {};
  MAPA_PECAS.forEach(function(p) { ids[String(p.id)] = p.Tipo; });
  return ids;
}

function doPost(e) {
  try {
    const corpo = e.postData ? e.postData.contents : null;
    if (!corpo) {
      return respostaOk();
    }

    const payload = JSON.parse(corpo);

    if (payload.event !== 'stock.created' && payload.event !== 'stock.updated') {
      return respostaOk();
    }

    const dados = payload.data;
    if (!dados || !dados.produto || !dados.produto.id) {
      return respostaOk();
    }

    const idProduto = dados.produto.id;

    // Filtra: só grava se for uma peça que rastreamos (CP/CPF/CPT)
    const idsRastreados = idsRastreadosSet_();
    if (!idsRastreados[String(idProduto)]) {
      return respostaOk();
    }

    const operacao = dados.operacao || '';
    const quantidade = dados.quantidade || 0;
    const dataEvento = payload.date || new Date().toISOString();

    const lock = LockService.getScriptLock();
    lock.waitLock(4000);

    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      let aba = ss.getSheetByName('Buffer do Dia');
      if (!aba) {
        aba = ss.insertSheet('Buffer do Dia');
        aba.appendRow(['Data/Hora Evento', 'ID Produto', 'Operação', 'Quantidade']);
      }
      aba.appendRow([dataEvento, idProduto, operacao, quantidade]);
    } finally {
      lock.releaseLock();
    }

    return respostaOk();
  } catch (err) {
    console.error('Erro no doPost: ' + err.message);
    return respostaOk();
  }
}

function respostaOk() {
  return ContentService.createTextOutput(JSON.stringify({ status: 'ok' }))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * ==================================================
 * CONSOLIDAÇÃO NOTURNA
 * ==================================================
 * Roda 1x por dia (gatilho por tempo). Resume o "Buffer do Dia" em
 * 1 linha por peça, grava no Registro CP/CPF/CPT correspondente,
 * limpa o buffer, e apaga registros com mais de 90 dias.
 */
function consolidarBufferDiario() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const bufferAba = ss.getSheetByName('Buffer do Dia');

  if (!bufferAba || bufferAba.getLastRow() < 2) {
    Logger.log('Buffer do Dia vazio, nada a consolidar.');
    return;
  }

  const dados = bufferAba.getRange(2, 1, bufferAba.getLastRow() - 1, 4).getValues();

  // Agrupa por ID Produto, somando saídas e entradas separadamente.
  // Balanço (B) NÃO é somado — ele representa o saldo final absoluto
  // definido manualmente, não uma quantidade movimentada. Guardamos
  // só o valor mais recente dele, como referência/observação.
  const resumoPorId = {}; // id -> { saidas: N, entradas: N, ultimoBalanco: N|null }
  dados.forEach(function(linha) {
    const id = String(linha[1]);
    const operacao = linha[2];
    const quantidade = Number(linha[3]) || 0;

    if (!resumoPorId[id]) resumoPorId[id] = { saidas: 0, entradas: 0, ultimoBalanco: null };

    if (operacao === 'S') {
      resumoPorId[id].saidas += quantidade;
    } else if (operacao === 'E') {
      resumoPorId[id].entradas += quantidade;
    } else if (operacao === 'B') {
      resumoPorId[id].ultimoBalanco = quantidade; // guarda só o mais recente do dia
    }
  });

  const idsRastreados = idsRastreadosSet_();
  const mapaPorId = {};
  MAPA_PECAS.forEach(function(p) { mapaPorId[String(p.id)] = p; });

  const hoje = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');

  // Agrupa as linhas a escrever por aba de destino (Registro CP/CPF/CPT)
  const linhasPorAba = { 'Registro CP': [], 'Registro CPF': [], 'Registro CPT': [] };

  Object.keys(resumoPorId).forEach(function(id) {
    const peca = mapaPorId[id];
    if (!peca) return; // segurança extra, não deveria acontecer já que filtramos no doPost

    const nomeAbaRegistro = 'Registro ' + peca.Tipo;
    if (!linhasPorAba[nomeAbaRegistro]) return;

    linhasPorAba[nomeAbaRegistro].push([
      hoje, peca.codigo, id, resumoPorId[id].saidas, resumoPorId[id].entradas,
      resumoPorId[id].ultimoBalanco !== null ? resumoPorId[id].ultimoBalanco : ''
    ]);
  });

  Object.keys(linhasPorAba).forEach(function(nomeAba) {
    const linhas = linhasPorAba[nomeAba];
    if (linhas.length === 0) return;

    let aba = ss.getSheetByName(nomeAba);
    if (!aba) {
      aba = ss.insertSheet(nomeAba);
      aba.appendRow(['Data', 'Código', 'ID Produto', 'Saídas do Dia', 'Entradas do Dia', 'Balanço (se houve)']);
    }

    aba.getRange(aba.getLastRow() + 1, 1, linhas.length, 6).setValues(linhas);
    Logger.log(nomeAba + ': ' + linhas.length + ' peças consolidadas.');
  });

  // Limpa o buffer do dia (mantém o cabeçalho)
  bufferAba.getRange(2, 1, bufferAba.getLastRow() - 1, 4).clearContent();

  // Limpa registros com mais de 90 dias em cada aba de Registro
  const limiteData = new Date();
  limiteData.setDate(limiteData.getDate() - 90);

  ['Registro CP', 'Registro CPF', 'Registro CPT'].forEach(function(nomeAba) {
    const aba = ss.getSheetByName(nomeAba);
    if (!aba || aba.getLastRow() < 2) return;

    const valores = aba.getRange(2, 1, aba.getLastRow() - 1, 6).getValues();
    const linhasParaManter = valores.filter(function(linha) {
      const dataLinha = new Date(linha[0]);
      return dataLinha >= limiteData;
    });

    if (linhasParaManter.length < valores.length) {
      aba.getRange(2, 1, valores.length, 6).clearContent();
      if (linhasParaManter.length > 0) {
        aba.getRange(2, 1, linhasParaManter.length, 6).setValues(linhasParaManter);
      }
      Logger.log(nomeAba + ': removidas ' + (valores.length - linhasParaManter.length) + ' linhas com mais de 90 dias.');
    }
  });

  Logger.log('Consolidação noturna concluída.');
}

/**
 * ==================================================
 * INVESTIGAÇÕES ANTERIORES (mantidas por referência)
 * ==================================================
 */
function buscarSaldosTodasPecas() {
  const service = getBlingService();
  if (!service.hasAccess()) {
    Logger.log('Ainda não autorizado. Rode "iniciarAutorizacao" primeiro.');
    return;
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let aba = ss.getSheetByName('Saldos Atuais (Todos)');
  if (!aba) {
    aba = ss.insertSheet('Saldos Atuais (Todos)');
  } else {
    aba.clear();
  }
  aba.appendRow(['Código', 'Tipo', 'ID Bling', 'Saldo Virtual']);

  const TAMANHO_LOTE = 50;
  const resultados = {};
  let erros = 0;

  for (let i = 0; i < MAPA_PECAS.length; i += TAMANHO_LOTE) {
    const lote = MAPA_PECAS.slice(i, i + TAMANHO_LOTE);
    const params = lote.map(function(p) { return 'idsProdutos[]=' + p.id; }).join('&');
    const url = 'https://api.bling.com.br/Api/v3/estoques/saldos?' + params;

    const resp = UrlFetchApp.fetch(url, {
      headers: { 'Authorization': 'Bearer ' + service.getAccessToken() },
      muteHttpExceptions: true
    });

    if (resp.getResponseCode() !== 200) {
      Logger.log('Erro no lote ' + (i / TAMANHO_LOTE + 1) + ': status ' + resp.getResponseCode());
      erros++;
      Utilities.sleep(400);
      continue;
    }

    const data = JSON.parse(resp.getContentText());
    (data.data || []).forEach(function(item) {
      resultados[String(item.produto.id)] = item.saldoVirtualTotal;
    });

    Utilities.sleep(400);
  }

  const linhas = MAPA_PECAS.map(function(p) {
    const saldo = resultados[String(p.id)];
    return [p.codigo, p.Tipo, p.id, (saldo !== undefined) ? saldo : 'NÃO ENCONTRADO'];
  });

  aba.getRange(2, 1, linhas.length, 4).setValues(linhas);
  Logger.log('Concluído! ' + linhas.length + ' peças processadas, ' + erros + ' lotes com erro.');
}

function inspecionarFiltroPedidosPorProduto() {
  const service = getBlingService();
  if (!service.hasAccess()) {
    Logger.log('Ainda não autorizado. Rode "iniciarAutorizacao" primeiro.');
    return;
  }

  const idProduto = 10000000001;
  const dataInicial = '2026-06-01';
  const dataFinal = '2026-07-25';

  const tentativas = [
    'https://api.bling.com.br/Api/v3/pedidos/vendas?idProduto=' + idProduto + '&dataInicial=' + dataInicial + '&dataFinal=' + dataFinal,
    'https://api.bling.com.br/Api/v3/pedidos/vendas?idsProdutos[]=' + idProduto + '&dataInicial=' + dataInicial + '&dataFinal=' + dataFinal,
    'https://api.bling.com.br/Api/v3/pedidos/vendas?codigoProduto=' + encodeURIComponent('CP00-EP0') + '&dataInicial=' + dataInicial + '&dataFinal=' + dataFinal
  ];

  tentativas.forEach(function(url, i) {
    const resp = UrlFetchApp.fetch(url, {
      headers: { 'Authorization': 'Bearer ' + service.getAccessToken() },
      muteHttpExceptions: true
    });
    Logger.log('--- Tentativa ' + (i + 1) + ': ' + url);
    Logger.log('Status: ' + resp.getResponseCode());
    Utilities.sleep(400);
  });
}

function testarFiltroComIdFalso() {
  const service = getBlingService();
  if (!service.hasAccess()) {
    Logger.log('Ainda não autorizado.');
    return;
  }

  const idFalso = 999999999999;
  const url = 'https://api.bling.com.br/Api/v3/pedidos/vendas?idProduto=' + idFalso +
    '&dataInicial=2026-06-01&dataFinal=2026-07-25';

  const resp = UrlFetchApp.fetch(url, {
    headers: { 'Authorization': 'Bearer ' + service.getAccessToken() },
    muteHttpExceptions: true
  });

  const data = JSON.parse(resp.getContentText());
  Logger.log('Status: ' + resp.getResponseCode());
  Logger.log('Quantidade de pedidos retornados: ' + (data.data ? data.data.length : 0));
}

/**
 * Atualiza o TEXTO do código na planilha (colunas A de CP/CPF/CPT) pra
 * refletir os códigos que foram corrigidos/padronizados direto no Bling.
 * Busca pelo ID (que já temos guardado), não pelo texto antigo.
 * Rode isso 1x só, depois pode apagar essa função se quiser.
 */
function atualizarCodigosCorrigidos() {
  const service = getBlingService();
  if (!service.hasAccess()) {
    Logger.log('Ainda não autorizado.');
    return;
  }

    const IDS_PARA_VERIFICAR = ["10000000001", "10000000002"]; // exemplo, IDs reais removidos

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const abas = ['CP', 'CPF', 'CPT'];
  let atualizados = 0;

  // Busca o código atual de cada ID na API
  const codigoAtualPorId = {};
  IDS_PARA_VERIFICAR.forEach(function(id) {
    const url = 'https://api.bling.com.br/Api/v3/produtos/' + id;
    const resp = UrlFetchApp.fetch(url, {
      headers: { 'Authorization': 'Bearer ' + service.getAccessToken() },
      muteHttpExceptions: true
    });
    Utilities.sleep(350);

    if (resp.getResponseCode() === 200) {
      const data = JSON.parse(resp.getContentText());
      if (data.data && data.data.codigo) {
        codigoAtualPorId[String(id)] = data.data.codigo;
      }
    } else {
      Logger.log('Erro ao buscar ID ' + id + ': status ' + resp.getResponseCode());
    }
  });

  Logger.log('Códigos buscados na API: ' + Object.keys(codigoAtualPorId).length);

  // Atualiza nas 3 abas, comparando pelo ID (coluna H)
  abas.forEach(function(nomeAba) {
    const aba = ss.getSheetByName(nomeAba);
    if (!aba || aba.getLastRow() < 2) return;

    const dados = aba.getRange(2, 1, aba.getLastRow() - 1, 9).getValues();
    let mudouNestaAba = false;

    for (let i = 0; i < dados.length; i++) {
      const id = String(dados[i][7]);
      const codigoNovo = codigoAtualPorId[id];
      if (codigoNovo && codigoNovo !== dados[i][0]) {
        Logger.log(nomeAba + ': "' + dados[i][0] + '" -> "' + codigoNovo + '"');
        dados[i][0] = codigoNovo;
        atualizados++;
        mudouNestaAba = true;
      }
    }

    if (mudouNestaAba) {
      aba.getRange(2, 1, dados.length, 9).setValues(dados);
      ordenarAbaPorPNumber(aba);
    }
  });

  Logger.log('Concluído! ' + atualizados + ' códigos atualizados.');
}

/**
 * ==================================================
 * DIAGNÓSTICO 1 — códigos duplicados nas abas CP/CPF/CPT
 * ==================================================
 * Roda uma vez, olha o "Registro de execução" (Ver > Registros / Logs).
 * Lista todo código que aparece em mais de uma linha, com aba, linha,
 * ID do Bling e saldo atual de cada ocorrência.
 */
function diagnosticarDuplicados() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const abas = ['CP', 'CPF', 'CPT'];
  const vistos = {}; // código -> lista de ocorrências

  abas.forEach(function(nomeAba) {
    const aba = ss.getSheetByName(nomeAba);
    if (!aba || aba.getLastRow() < 2) return;

    const dados = aba.getRange(2, 1, aba.getLastRow() - 1, 9).getValues();
    dados.forEach(function(linha, idx) {
      const codigo = String(linha[0]).trim();
      if (!codigo) return;
      if (!vistos[codigo]) vistos[codigo] = [];
      vistos[codigo].push({
        aba: nomeAba,
        linhaPlanilha: idx + 2,
        id: linha[7],
        saldo: linha[1]
      });
    });
  });

  let totalDuplicados = 0;
  Logger.log('=== CÓDIGOS DUPLICADOS (aparecem em mais de 1 linha) ===');
  Object.keys(vistos).sort().forEach(function(codigo) {
    if (vistos[codigo].length > 1) {
      totalDuplicados++;
      Logger.log(codigo + ':');
      vistos[codigo].forEach(function(info) {
        Logger.log('  Aba ' + info.aba + ', linha ' + info.linhaPlanilha +
          ', ID Bling=' + info.id + ', Saldo na planilha=' + info.saldo);
      });
    }
  });

  Logger.log('=== TOTAL DE CÓDIGOS DUPLICADOS: ' + totalDuplicados + ' ===');
}

/**
 * ==================================================
 * DIAGNÓSTICO 2 — saldo físico vs saldo virtual
 * ==================================================
 * Passe o código exato (ex: "CP00-EP0") e ele acha o ID guardado na
 * planilha, consulta o Bling e mostra a resposta crua da API, pra
 * comparar saldoFisicoTotal com saldoVirtualTotal.
 */
function diagnosticarSaldoDetalhado(codigoAlvo) {
  const service = getBlingService();
  if (!service.hasAccess()) {
    Logger.log('Ainda não autorizado. Rode "iniciarAutorizacao" primeiro.');
    return;
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const abas = ['CP', 'CPF', 'CPT'];
  let encontrou = false;

  abas.forEach(function(nomeAba) {
    const aba = ss.getSheetByName(nomeAba);
    if (!aba || aba.getLastRow() < 2) return;

    const dados = aba.getRange(2, 1, aba.getLastRow() - 1, 9).getValues();
    dados.forEach(function(linha, idx) {
      const codigo = String(linha[0]).trim();
      if (codigo !== codigoAlvo) return;

      encontrou = true;
      const id = linha[7];
      Logger.log('--- Encontrado em ' + nomeAba + ', linha ' + (idx + 2) + ' ---');
      Logger.log('Código: ' + codigo + ' | ID Bling: ' + id + ' | Saldo na planilha: ' + linha[1]);

      if (!id) {
        Logger.log('Essa linha não tem ID Bling gravado (coluna H vazia).');
        return;
      }

      const url = 'https://api.bling.com.br/Api/v3/estoques/saldos?idsProdutos[]=' + id;
      const resp = UrlFetchApp.fetch(url, {
        headers: { 'Authorization': 'Bearer ' + service.getAccessToken() },
        muteHttpExceptions: true
      });

      Logger.log('Status API: ' + resp.getResponseCode());
      Logger.log('Resposta crua: ' + resp.getContentText());
    });
  });

  if (!encontrou) {
    Logger.log('Código "' + codigoAlvo + '" não encontrado em nenhuma das abas CP/CPF/CPT.');
  }
}

/**
 * ==================================================
 * LIMPEZA — IDs mortos (produto excluído e recriado no Bling)
 * ==================================================
 * Causa raiz: quando um produto é excluído e recriado no Bling, ele ganha
 * um ID novo. A planilha só busca ID novo quando a coluna H está vazia,
 * então a linha antiga fica presa num ID que não existe mais e nunca mais
 * atualiza o saldo.
 *
 * Esta função verifica, produto por produto, se o ID gravado na coluna H
 * ainda existe no Bling:
 * - Código único (não duplicado) com ID morto -> limpa a coluna H (força
 *   a rebusca automática na próxima "Atualizar Estoque").
 * - Código duplicado (2+ linhas) com 1 ID válido e o resto morto -> apaga
 *   a(s) linha(s) morta(s) automaticamente.
 * - Código duplicado com 2+ IDs válidos e ativos -> NÃO mexe, só avisa no
 *   log (são produtos genuinamente diferentes no Bling com código igual,
 *   precisa de decisão manual: mesclar ou renomear um deles).
 *
 * Processa em lotes (por causa do limite de execução do Apps Script) e
 * guarda o progresso, então rode de novo se aparecer "limite atingido"
 * no log, até não aparecer mais.
 */
function limparDuplicadosEIdsInvalidos() {
  const service = getBlingService();
  if (!service.hasAccess()) {
    Logger.log('Ainda não autorizado. Rode "iniciarAutorizacao" primeiro.');
    return;
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const abas = ['CP', 'CPF', 'CPT'];
  const props = PropertiesService.getScriptProperties();
  const LIMITE_VERIFICACOES = 400; // ~400 chamadas por execução (~3min)
  let verificacoesFeitas = 0;

  abas.forEach(function(nomeAba) {
    if (verificacoesFeitas >= LIMITE_VERIFICACOES) return;

    const aba = ss.getSheetByName(nomeAba);
    if (!aba || aba.getLastRow() < 2) return;

    const dados = aba.getRange(2, 1, aba.getLastRow() - 1, 9).getValues();

    const grupos = {};
    dados.forEach(function(linha, idx) {
      const codigo = String(linha[0]).trim();
      if (!codigo) return;
      if (!grupos[codigo]) grupos[codigo] = [];
      grupos[codigo].push(idx);
    });

    const chaveProgresso = 'PROGRESSO_LIMPEZA_' + nomeAba;
    const inicioIdx = parseInt(props.getProperty(chaveProgresso) || '0', 10);
    const codigosOrdenados = Object.keys(grupos);

    const linhasParaExcluir = [];
    const linhasParaLimparId = [];

    for (let ci = inicioIdx; ci < codigosOrdenados.length; ci++) {
      if (verificacoesFeitas >= LIMITE_VERIFICACOES) {
        props.setProperty(chaveProgresso, String(ci));
        Logger.log(nomeAba + ': limite atingido nesta execução (parou no código ' + ci +
          ' de ' + codigosOrdenados.length + '). Rode a função de novo pra continuar.');
        break;
      }

      const codigo = codigosOrdenados[ci];
      const indices = grupos[codigo];

      const status = indices.map(function(idx) {
        const id = dados[idx][7];
        if (!id) return { idx: idx, id: id, valido: null };

        const url = 'https://api.bling.com.br/Api/v3/produtos/' + id;
        const resp = UrlFetchApp.fetch(url, {
          headers: { 'Authorization': 'Bearer ' + service.getAccessToken() },
          muteHttpExceptions: true
        });
        Utilities.sleep(300);
        verificacoesFeitas++;
        return { idx: idx, id: id, valido: resp.getResponseCode() === 200 };
      });

      const validos = status.filter(function(s) { return s.valido === true; });
      const invalidos = status.filter(function(s) { return s.valido === false; });

      if (indices.length === 1) {
        if (invalidos.length === 1) {
          linhasParaLimparId.push(invalidos[0].idx);
          Logger.log(nomeAba + ': "' + codigo + '" (linha ' + (invalidos[0].idx + 2) +
            ') tinha ID morto (' + invalidos[0].id + '). ID limpo, será rebuscado na próxima Atualizar Estoque.');
        }
        continue;
      }

      // Código duplicado (2+ linhas)
      if (validos.length === 1 && invalidos.length >= 1) {
        invalidos.forEach(function(s) { linhasParaExcluir.push(s.idx); });
        Logger.log(nomeAba + ': "' + codigo + '" duplicado -> mantendo linha ' + (validos[0].idx + 2) +
          ' (ID ' + validos[0].id + '), removendo ' + invalidos.length + ' linha(s) com ID morto.');
      } else if (validos.length > 1) {
        Logger.log(nomeAba + ': ATENÇÃO — "' + codigo + '" tem ' + validos.length +
          ' produtos DIFERENTES e ATIVOS no Bling com o mesmo código (IDs: ' +
          validos.map(function(s) { return s.id; }).join(', ') +
          '). Precisa de decisão manual (mesclar ou renomear um deles no Bling).');
      } else {
        Logger.log(nomeAba + ': "' + codigo + '" duplicado mas nenhuma linha tem ID válido ainda — deixando como está.');
      }

      if (ci === codigosOrdenados.length - 1) {
        props.deleteProperty(chaveProgresso); // terminou de processar essa aba inteira
      }
    }

    linhasParaLimparId.forEach(function(idx) { dados[idx][7] = ''; });
    if (linhasParaLimparId.length > 0) {
      aba.getRange(2, 1, dados.length, 9).setValues(dados);
    }

    // Exclui de baixo pra cima pra não bagunçar os índices das linhas seguintes
    linhasParaExcluir.sort(function(a, b) { return b - a; });
    linhasParaExcluir.forEach(function(idx) { aba.deleteRow(idx + 2); });

    if (linhasParaExcluir.length > 0) {
      Logger.log(nomeAba + ': ' + linhasParaExcluir.length + ' linha(s) duplicada(s) removida(s).');
    }
  });

  Logger.log('=== Verificações feitas nesta execução: ' + verificacoesFeitas + ' ===');
  Logger.log('Se algum aviso de "limite atingido" apareceu acima, rode a função de novo pra continuar de onde parou.');
}

