// ═══════════════════════════════════════════════════════════════════════
// GOOGLE APPS SCRIPT — Controle de Revisão Geral de Apartamentos
// Pousada Marie Claire
// ═══════════════════════════════════════════════════════════════════════

const EMAIL_DESTINO       = 'diretoria@pousadamarieclaire.com';
const NOME_ABA_HISTORICO  = 'Histórico';
const NOME_ABA_PAINEL     = 'Painel';
const NOME_ABA_PENDENCIAS = 'Pendências';

// ID da pasta raiz no Google Drive onde ficarão as fotos
// Crie uma pasta chamada "Revisões - Pousada Marie Claire" no Drive
// e cole o ID dela aqui (está na URL: drive.google.com/drive/folders/ESTE_ID_AQUI)
const PASTA_RAIZ_DRIVE = '1RO4_T-FlwKr2XuGh04j7EICUd_2OJ-cl';

// ── POST ─────────────────────────────────────────────────────────────────
function doPost(e) {
  try {
    const dados = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // Log de debug
    if (!dados.action) {
      const tamanho = e.postData.contents.length;
      Logger.log('=== NOVA REVISÃO ===');
      Logger.log('Payload size: ' + tamanho + ' bytes (' + Math.round(tamanho/1024) + ' KB)');
      Logger.log('Apartamento: ' + dados.apartamento);
      const secFotos = ['hid','ele','mob','ele2','por','jan','pis','for','par'];
      secFotos.forEach(function(s) {
        const arr = dados['fotos_' + s];
        if (arr && arr.length > 0) {
          const kb = arr.reduce(function(t,f){ return t + (f.base64 ? f.base64.length : 0); }, 0) / 1024;
          Logger.log('fotos_' + s + ': ' + arr.length + ' foto(s), ~' + Math.round(kb) + ' KB base64');
        }
      });
    }

    // Upload de foto — salva no Drive e opcionalmente atualiza célula da planilha
    if (dados.action === 'upload_foto' || dados.action === 'upload_foto_pendencia') {
      return uploadFoto(dados, ss);
    }

    // Resolver pendência
    if (dados.action === 'resolver_pendencia') {
      return resolverPendencia(ss, dados);
    }

    // Registrar revisão completa — salva TUDO e retorna linhas das pendências
    salvarHistorico(ss, dados);
    atualizarPainel(ss, dados);
    const linhasPendencias = criarPendencias(ss, dados);
    enviarEmail(dados);

    return ContentService
      .createTextOutput(JSON.stringify({ status: 'ok', linhas_pendencias: linhasPendencias }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'erro', msg: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ── GET ───────────────────────────────────────────────────────────────────
function doGet(e) {
  const params = e && e.parameter ? e.parameter : {};

  if (params.action === 'historico') return retornarHistorico();
  if (params.action === 'pendencias') return retornarPendencias(params.apartamento || '');
  if (params.action === 'contar_pendencias') {
    try {
      const ss  = SpreadsheetApp.getActiveSpreadsheet();
      const aba = ss.getSheetByName(NOME_ABA_PENDENCIAS);
      const total = aba ? Math.max(0, aba.getLastRow() - 1) : 0; // -1 cabeçalho
      return ContentService
        .createTextOutput(JSON.stringify({ status: 'ok', total: total }))
        .setMimeType(ContentService.MimeType.JSON);
    } catch(e) {
      return ContentService
        .createTextOutput(JSON.stringify({ status: 'erro', total: null }))
        .setMimeType(ContentService.MimeType.JSON);
    }
  }

  return ContentService
    .createTextOutput(JSON.stringify({ status: 'ok', msg: 'Script ativo.' }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── Upload de foto para o Google Drive ───────────────────────────────────
// Se dados.linha_planilha existir, atualiza a coluna L da aba Pendências com o link
function uploadFoto(dados, ss) {
  try {
    // dados.foto_base64 : string base64 da imagem
    // dados.nome_arquivo: ex. "foto_1.jpg"
    // dados.apartamento : ex. "Apto 101"
    // dados.secao       : ex. "💧 Hidráulico"
    // dados.tipo        : "pendencia" ou "resolucao"

    const pastaRaiz = DriveApp.getFolderById(PASTA_RAIZ_DRIVE);

    // Pasta do apartamento
    let pastaApto;
    const foldersApto = pastaRaiz.getFoldersByName(dados.apartamento);
    if (foldersApto.hasNext()) {
      pastaApto = foldersApto.next();
    } else {
      pastaApto = pastaRaiz.createFolder(dados.apartamento);
    }

    // Pasta da data (DD-MM-AAAA)
    const hoje = new Date();
    const dia  = String(hoje.getDate()).padStart(2, '0');
    const mes  = String(hoje.getMonth() + 1).padStart(2, '0');
    const ano  = hoje.getFullYear();
    const nomeData = `${dia}-${mes}-${ano}`;

    let pastaData;
    const foldersData = pastaApto.getFoldersByName(nomeData);
    if (foldersData.hasNext()) {
      pastaData = foldersData.next();
    } else {
      pastaData = pastaApto.createFolder(nomeData);
    }

    // Subpasta: "Pendências" ou "Resoluções"
    const subNome = dados.tipo === 'resolucao' ? 'Resoluções' : 'Pendências';
    let subPasta;
    const foldersSub = pastaData.getFoldersByName(subNome);
    if (foldersSub.hasNext()) {
      subPasta = foldersSub.next();
    } else {
      subPasta = pastaData.createFolder(subNome);
    }

    // Decodificar base64 e criar arquivo
    const base64 = dados.foto_base64.replace(/^data:image\/\w+;base64,/, '');
    const blob = Utilities.newBlob(
      Utilities.base64Decode(base64),
      dados.mime_type || 'image/jpeg',
      dados.nome_arquivo || ('foto_' + new Date().getTime() + '.jpg')
    );

    const arquivo = subPasta.createFile(blob);
    try {
      arquivo.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    } catch(shareErr) {
      Logger.log('Aviso compartilhamento: ' + shareErr.message);
    }

    const linkVisualizacao = 'https://drive.google.com/file/d/' + arquivo.getId() + '/view';

    // Se linha_planilha foi fornecida, atualiza coluna L da aba Pendências
    if (ss && dados.linha_planilha) {
      try {
        const linhaNum = parseInt(dados.linha_planilha);
        const abaPend = ss.getSheetByName(NOME_ABA_PENDENCIAS);
        if (abaPend && linhaNum >= 2) {
          const celula = abaPend.getRange(linhaNum, 12);
          const valorAtual = celula.getValue();
          const novosLinks = (valorAtual && valorAtual !== '—')
            ? valorAtual + '\n' + linkVisualizacao
            : linkVisualizacao;
          celula.setValue(novosLinks);
        }
      } catch(e) {
        Logger.log('Erro ao atualizar célula: ' + e.message);
      }
    }

    return ContentService
      .createTextOutput(JSON.stringify({ status: 'ok', link: linkVisualizacao, id: arquivo.getId() }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'erro', msg: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ── Resolver pendência ────────────────────────────────────────────────────
function resolverPendencia(ss, dados) {
  try {
    const aba = ss.getSheetByName(NOME_ABA_PENDENCIAS);
    if (!aba) throw new Error('Aba Pendências não encontrada');

    const linhaNum = parseInt(dados.linha_planilha);
    if (!linhaNum || linhaNum < 2) throw new Error('Linha inválida');

    const hoje = new Date();
    const dataFormatada = formatarData(hoje);

    aba.getRange(linhaNum, 8).setValue('Resolvida');
    aba.getRange(linhaNum, 9).setValue(dataFormatada);
    aba.getRange(linhaNum, 10).setValue(dados.tecnico_resolve || '—');

    // Fazer upload das fotos de resolução e salvar links (coluna 13)
    if (dados.fotos_resolucao && dados.fotos_resolucao.length > 0) {
      const aptNome = aba.getRange(linhaNum, 2).getValue() || 'Geral';
      const linksResolucao = uploadFotosBase64(dados.fotos_resolucao, aptNome, 'resolucao');
      if (linksResolucao && linksResolucao !== '—') aba.getRange(linhaNum, 13).setValue(linksResolucao);
    }

    return ContentService
      .createTextOutput(JSON.stringify({ status: 'ok' }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'erro', msg: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ── Retornar histórico ────────────────────────────────────────────────────
function retornarHistorico() {
  try {
    const ss  = SpreadsheetApp.getActiveSpreadsheet();
    const aba = ss.getSheetByName(NOME_ABA_PAINEL);

    if (!aba) {
      return ContentService
        .createTextOutput(JSON.stringify([]))
        .setMimeType(ContentService.MimeType.JSON);
    }

    const dados     = aba.getDataRange().getValues();
    const cabecalho = dados[0];

    const iApartamento   = cabecalho.indexOf('Apartamento');
    const iAndar         = cabecalho.indexOf('Andar');
    const iUltimaRevisao = cabecalho.indexOf('Última Revisão');
    const iProxRevisao   = cabecalho.indexOf('Próxima Revisão');
    const iUltimoTecnico = cabecalho.indexOf('Último Técnico');

    const resultado = [];

    for (let i = 1; i < dados.length; i++) {
      const linha = dados[i];
      const ultimaRevisao = linha[iUltimaRevisao];
      const proxRevisao   = linha[iProxRevisao];

      if (!ultimaRevisao || ultimaRevisao === '—' || ultimaRevisao === '') continue;

      resultado.push({
        apartamento:     linha[iApartamento]   || '',
        andar:           linha[iAndar]         || '',
        data_revisao:    formatarData(ultimaRevisao),
        proxima_revisao: formatarData(proxRevisao),
        tecnico:         linha[iUltimoTecnico] || '',
      });
    }

    return ContentService
      .createTextOutput(JSON.stringify(resultado))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ erro: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ── Retornar pendências abertas ───────────────────────────────────────────
function retornarPendencias(apartamento) {
  try {
    const ss  = SpreadsheetApp.getActiveSpreadsheet();
    const aba = ss.getSheetByName(NOME_ABA_PENDENCIAS);

    if (!aba) {
      return ContentService
        .createTextOutput(JSON.stringify([]))
        .setMimeType(ContentService.MimeType.JSON);
    }

    const dados     = aba.getDataRange().getValues();
    const cabecalho = dados[0];

    const iId           = cabecalho.indexOf('ID');
    const iApartamento  = cabecalho.indexOf('Apartamento');
    const iSecao        = cabecalho.indexOf('Seção');
    const iDescricao    = cabecalho.indexOf('Descrição');
    const iDataAbertura = cabecalho.indexOf('Data Abertura');
    const iTecnicoAbre  = cabecalho.indexOf('Técnico Abertura');
    const iStatus       = cabecalho.indexOf('Status');
    const iFotosPend    = cabecalho.indexOf('Fotos Pendência') !== -1
      ? cabecalho.indexOf('Fotos Pendência')
      : 11; // coluna L como fallback

    const resultado = [];

    for (let i = 1; i < dados.length; i++) {
      const linha = dados[i];
      if (linha[iStatus] !== 'Aberta') continue;
      if (apartamento && linha[iApartamento] !== apartamento) continue;

      resultado.push({
        id:             linha[iId]           || '',
        apartamento:    linha[iApartamento]  || '',
        secao:          linha[iSecao]        || '',
        descricao:      linha[iDescricao]    || '',
        data_abertura:  formatarData(linha[iDataAbertura]),
        tecnico_abre:   linha[iTecnicoAbre]  || '',
        fotos:          linha[iFotosPend]    || '',
        linha_planilha: i + 1,
      });
    }

    return ContentService
      .createTextOutput(JSON.stringify(resultado))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ erro: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ── Upload de array de fotos base64 ──────────────────────────────────────
function uploadFotosBase64(fotosArr, apartamento, tipo) {
  if (!Array.isArray(fotosArr) || fotosArr.length === 0) return '—';

  // Verifica se os itens têm base64 (objetos) ou são strings (links prontos)
  if (typeof fotosArr[0] === 'string') {
    // Já são links — só junta
    const links = fotosArr.filter(function(l){ return l && l.trim() !== ''; });
    return links.length > 0 ? links.join('\n') : '—';
  }

  try {
    const pastaRaiz = DriveApp.getFolderById(PASTA_RAIZ_DRIVE);

    let pastaApto;
    const foldersApto = pastaRaiz.getFoldersByName(apartamento);
    if (foldersApto.hasNext()) { pastaApto = foldersApto.next(); }
    else { pastaApto = pastaRaiz.createFolder(apartamento); }

    const hoje = new Date();
    const nomeData = String(hoje.getDate()).padStart(2,'0') + '-' + String(hoje.getMonth()+1).padStart(2,'0') + '-' + hoje.getFullYear();

    let pastaData;
    const foldersData = pastaApto.getFoldersByName(nomeData);
    if (foldersData.hasNext()) { pastaData = foldersData.next(); }
    else { pastaData = pastaApto.createFolder(nomeData); }

    const subNome = tipo === 'resolucao' ? 'Resoluções' : 'Pendências';
    let subPasta;
    const foldersSub = pastaData.getFoldersByName(subNome);
    if (foldersSub.hasNext()) { subPasta = foldersSub.next(); }
    else { subPasta = pastaData.createFolder(subNome); }

    const links = [];
    fotosArr.forEach(function(f, idx) {
      try {
        if (!f || !f.base64) {
          Logger.log('Foto ' + idx + ' sem base64 — ignorada');
          return;
        }
        const base64str = f.base64.replace(/^data:image\/\w+;base64,/, '');
        Logger.log('Foto ' + idx + ': ' + Math.round(base64str.length / 1024) + ' KB base64');
        const blob = Utilities.newBlob(
          Utilities.base64Decode(base64str),
          f.mime || 'image/jpeg',
          f.nome || ('foto_' + idx + '_' + Date.now() + '.jpg')
        );
        const arquivo = subPasta.createFile(blob);
        try {
          arquivo.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        } catch(shareErr) {
          Logger.log('Aviso compartilhamento foto ' + idx + ': ' + shareErr.message);
        }
        const link = 'https://drive.google.com/file/d/' + arquivo.getId() + '/view';
        links.push(link);
        Logger.log('Foto ' + idx + ' salva: ' + link);
      } catch(errFoto) {
        Logger.log('Erro foto ' + idx + ': ' + errFoto.message);
      }
    });

    if (links.length === 0) return 'ERRO: nenhuma foto salva (ver logs Apps Script)';
    return links.join('\n');
  } catch(err) {
    Logger.log('Erro uploadFotosBase64: ' + err.message);
    return 'ERRO: ' + err.message;
  }
}

// ── Upload de foto para pendência já salva ────────────────────────────────
// Chamado pelo formulário após salvar a revisão, foto por foto
function uploadFotoPendencia(ss, dados) {
  try {
    const aba = ss.getSheetByName(NOME_ABA_PENDENCIAS);
    if (!aba) throw new Error('Aba Pendências não encontrada');

    const linhaNum = parseInt(dados.linha_planilha);
    if (!linhaNum || linhaNum < 2) throw new Error('Linha inválida: ' + dados.linha_planilha);

    // Faz upload da foto
    const pastaRaiz = DriveApp.getFolderById(PASTA_RAIZ_DRIVE);
    const aptNome = dados.apartamento || aba.getRange(linhaNum, 2).getValue() || 'Geral';

    let pastaApto;
    const fA = pastaRaiz.getFoldersByName(aptNome);
    pastaApto = fA.hasNext() ? fA.next() : pastaRaiz.createFolder(aptNome);

    const hoje = new Date();
    const nomeData = String(hoje.getDate()).padStart(2,'0') + '-' + String(hoje.getMonth()+1).padStart(2,'0') + '-' + hoje.getFullYear();
    let pastaData;
    const fD = pastaApto.getFoldersByName(nomeData);
    pastaData = fD.hasNext() ? fD.next() : pastaApto.createFolder(nomeData);

    let subPasta;
    const fS = pastaData.getFoldersByName('Pendências');
    subPasta = fS.hasNext() ? fS.next() : pastaData.createFolder('Pendências');

    const base64str = dados.foto_base64.replace(/^data:image\/\w+;base64,/, '');
    const blob = Utilities.newBlob(
      Utilities.base64Decode(base64str),
      'image/jpeg',
      dados.nome_arquivo || ('foto_' + Date.now() + '.jpg')
    );
    const arquivo = subPasta.createFile(blob);
    try {
      arquivo.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    } catch(shareErr) {
      Logger.log('Aviso compartilhamento: ' + shareErr.message);
    }
    const novoLink = 'https://drive.google.com/file/d/' + arquivo.getId() + '/view';

    // Adiciona o link na célula da coluna L (col 12), preservando links existentes
    const celula = aba.getRange(linhaNum, 12);
    const valorAtual = celula.getValue();
    const novosLinks = (valorAtual && valorAtual !== '—')
      ? valorAtual + '\n' + novoLink
      : novoLink;
    celula.setValue(novosLinks);

    Logger.log('Foto salva na linha ' + linhaNum + ': ' + novoLink);

    return ContentService
      .createTextOutput(JSON.stringify({ status: 'ok', link: novoLink }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch(err) {
    Logger.log('Erro uploadFotoPendencia: ' + err.message);
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'erro', msg: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ── Criar pendências ──────────────────────────────────────────────────────
// Retorna objeto { hid: linhaNum, ele: linhaNum, ... } para cada seção criada
function criarPendencias(ss, d) {
  let aba = ss.getSheetByName(NOME_ABA_PENDENCIAS);

  if (!aba) {
    aba = ss.insertSheet(NOME_ABA_PENDENCIAS);
    const cabecalhos = [
      'ID', 'Apartamento', 'Andar', 'Seção', 'Descrição',
      'Data Abertura', 'Técnico Abertura',
      'Status', 'Data Resolução', 'Técnico Resolução',
      'ID Revisão Origem', 'Fotos Pendência', 'Fotos Resolução'
    ];
    aba.getRange(1, 1, 1, cabecalhos.length).setValues([cabecalhos]);
    aba.getRange(1, 1, 1, cabecalhos.length).setFontWeight('bold');
    aba.setFrozenRows(1);
  }

  const secoes = [
    { chave: 'hid',  secao: '💧 Hidráulico',       obs: d.hid_obs  },
    { chave: 'ele',  secao: '⚡ Elétrica',          obs: d.ele_obs  },
    { chave: 'mob',  secao: '🛋️ Móveis',            obs: d.mob_obs  },
    { chave: 'ele2', secao: '🍳 Eletrodomésticos',  obs: d.ele2_obs },
    { chave: 'por',  secao: '🚪 Portas',            obs: d.por_obs  },
    { chave: 'jan',  secao: '🪟 Janelas',           obs: d.jan_obs  },
    { chave: 'pis',  secao: '🟫 Piso',              obs: d.pis_obs  },
    { chave: 'for',  secao: '🔲 Forro/Teto',        obs: d.for_obs  },
    { chave: 'par',  secao: '🖼️ Paredes',           obs: d.par_obs  },
  ];

  let proximaLinha = aba.getLastRow() + 1;
  const linhasCriadas = {};

  secoes.forEach(function(s) {
    if (!s.obs || s.obs === '—' || s.obs.trim() === '') return;

    const linhaNum = proximaLinha;
    const idPendencia = 'PEN-' + String(linhaNum - 1).padStart(4, '0');

    // Processar fotos desta seção já no momento de criar a pendência
    let linksFotos = '—';
    const fotosSecao = d['fotos_' + s.chave];
    if (fotosSecao && Array.isArray(fotosSecao) && fotosSecao.length > 0) {
      try {
        linksFotos = uploadFotosBase64(fotosSecao, d.apartamento || 'Geral', 'pendencia');
      } catch(errF) {
        Logger.log('Erro ao fazer upload de fotos da seção ' + s.chave + ': ' + errF.message);
      }
    }

    aba.appendRow([
      idPendencia,
      d.apartamento  || '—',
      d.andar        || '—',
      s.secao,
      s.obs.trim(),
      d.data_revisao || '—',
      d.tecnico      || '—',
      'Aberta',
      '—',
      '—',
      'REV-' + String(linhaNum - 1).padStart(3, '0'),
      linksFotos,
      '—',
    ]);

    linhasCriadas[s.chave] = linhaNum;
    Logger.log('Pendência criada: ' + idPendencia + ' na linha ' + linhaNum + ' | fotos: ' + linksFotos);
    proximaLinha++;
  });

  return linhasCriadas;
}

// ── Salvar no Histórico ───────────────────────────────────────────────────
function salvarHistorico(ss, d) {
  let aba = ss.getSheetByName(NOME_ABA_HISTORICO);

  if (!aba) {
    aba = ss.insertSheet(NOME_ABA_HISTORICO);
    const cabecalhos = [
      'ID Revisão', 'Andar', 'Apartamento', 'Data Revisão', 'Próx. Revisão',
      'Técnico',
      'Hidráulico', 'Elétrica', 'Móveis', 'Eletrodomést.',
      'Portas', 'Janelas', 'Piso', 'Forro', 'Paredes',
      'Problemas', 'Observações Gerais',
      'Obs Hidráulico', 'Obs Elétrica', 'Obs Móveis', 'Obs Eletrodomést.',
      'Obs Portas', 'Obs Janelas', 'Obs Piso', 'Obs Forro', 'Obs Paredes'
    ];
    aba.getRange(1, 1, 1, cabecalhos.length).setValues([cabecalhos]);
    aba.getRange(1, 1, 1, cabecalhos.length).setFontWeight('bold');
    aba.setFrozenRows(1);
  }

  const ultimaLinha = aba.getLastRow() + 1;
  const idRevisao   = 'REV-' + String(ultimaLinha - 1).padStart(3, '0');

  aba.getRange(ultimaLinha, 1, 1, 26).setValues([[
    idRevisao,
    d.andar               || '—',
    d.apartamento         || '—',
    d.data_revisao        || '—',
    d.proxima_revisao     || '—',
    d.tecnico             || '—',
    d.status_hidraulico   || '—',
    d.status_eletrica     || '—',
    d.status_mobiliario   || '—',
    d.status_eletrodomest || '—',
    d.status_portas       || '—',
    d.status_janelas      || '—',
    d.status_piso         || '—',
    d.status_forro        || '—',
    d.status_paredes      || '—',
    d.problemas           || '—',
    d.observacoes         || '—',
    d.hid_obs             || '—',
    d.ele_obs             || '—',
    d.mob_obs             || '—',
    d.ele2_obs            || '—',
    d.por_obs             || '—',
    d.jan_obs             || '—',
    d.pis_obs             || '—',
    d.for_obs             || '—',
    d.par_obs             || '—',
  ]]);
}

// ── Atualizar Painel ──────────────────────────────────────────────────────
function atualizarPainel(ss, d) {
  let aba = ss.getSheetByName(NOME_ABA_PAINEL);

  if (!aba) {
    aba = ss.insertSheet(NOME_ABA_PAINEL);
    const cabecalhos = [
      'Nº', 'Andar', 'Apartamento', 'Última Revisão', 'Próxima Revisão',
      'Status', 'Hidráulico', 'Elétrica', 'Móveis', 'Eletrodomést.',
      'Portas', 'Janelas', 'Piso', 'Forro', 'Paredes', 'Último Técnico'
    ];
    aba.getRange(1, 1, 1, cabecalhos.length).setValues([cabecalhos]);
    aba.getRange(1, 1, 1, cabecalhos.length).setFontWeight('bold');
    aba.setFrozenRows(1);

    const aptos = [
      [1,'Térreo','Apto 01'],[2,'Térreo','Apto 02'],[3,'Térreo','Apto 03'],
      [4,'Térreo','Apto 04'],[5,'Térreo','Apto 05'],[6,'Térreo','Apto 06'],
      [7,'Térreo','Apto 07'],[8,'Térreo','Apto 08'],[9,'Térreo','Apto 09'],
      [10,'Térreo','Apto 10'],[11,'Térreo','Apto 11'],[12,'Térreo','Apto 12'],
      [13,'Térreo','Apto 13'],
      [14,'1º Andar','Apto 101'],[15,'1º Andar','Apto 102'],[16,'1º Andar','Apto 103'],
      [17,'1º Andar','Apto 104'],[18,'1º Andar','Apto 105'],[19,'1º Andar','Apto 106'],
      [20,'1º Andar','Apto 107'],[21,'1º Andar','Apto 108'],[22,'1º Andar','Apto 109'],
      [23,'1º Andar','Apto 110'],[24,'1º Andar','Apto 111'],[25,'1º Andar','Apto 112'],
      [26,'2º Andar','Apto 201'],[27,'2º Andar','Apto 202'],[28,'2º Andar','Apto 203'],
      [29,'2º Andar','Apto 204'],[30,'2º Andar','Apto 205'],[31,'2º Andar','Apto 206'],
      [32,'2º Andar','Apto 207'],[33,'2º Andar','Apto 208'],[34,'2º Andar','Apto 209'],
      [35,'2º Andar','Apto 210'],[36,'2º Andar','Apto 211'],[37,'2º Andar','Apto 212'],
      [38,'2º Andar','Apto 213'],[39,'2º Andar','Apto 214'],
      [40,'3º Andar','Apto 301'],[41,'3º Andar','Apto 302'],[42,'3º Andar','Apto 303'],
      [43,'3º Andar','Apto 304'],[44,'3º Andar','Apto 305'],[45,'3º Andar','Apto 306'],
    ];
    aptos.forEach(([num, andar, nome]) => {
      aba.getRange(num + 1, 1, 1, 16).setValues([[
        num, andar, nome,
        '—', '—', '⏳ Sem revisão',
        '—', '—', '—', '—', '—', '—', '—', '—', '—', '—'
      ]]);
    });
  }

  const numApt = parseInt(d.apartamento_num);
  if (numApt >= 1 && numApt <= 45) {
    const linhaApt = numApt + 1;

    const hoje = new Date();
    const partes = d.proxima_revisao.split('/');
    const dataProx = new Date(partes[2], partes[1] - 1, partes[0]);
    const diasRestantes = Math.round((dataProx - hoje) / (1000 * 60 * 60 * 24));
    let status = '';
    if (diasRestantes < 0)        status = '🔴 Vencido';
    else if (diasRestantes <= 14) status = '⚠️ Próximo';
    else                          status = '✅ Em Dia';

    aba.getRange(linhaApt, 4, 1, 13).setValues([[
      d.data_revisao,
      d.proxima_revisao,
      status,
      d.status_hidraulico   || '—',
      d.status_eletrica     || '—',
      d.status_mobiliario   || '—',
      d.status_eletrodomest || '—',
      d.status_portas       || '—',
      d.status_janelas      || '—',
      d.status_piso         || '—',
      d.status_forro        || '—',
      d.status_paredes      || '—',
      d.tecnico             || '—',
    ]]);
  }
}

// ── Enviar e-mail ─────────────────────────────────────────────────────────
function enviarEmail(d) {
  const secoes = [
    { nome: '💧 Hidráulico',       obs: d.hid_obs  },
    { nome: '⚡ Elétrica',         obs: d.ele_obs  },
    { nome: '🛋️ Móveis',           obs: d.mob_obs  },
    { nome: '🍳 Eletrodomésticos', obs: d.ele2_obs },
    { nome: '🚪 Portas',           obs: d.por_obs  },
    { nome: '🪟 Janelas',          obs: d.jan_obs  },
    { nome: '🟫 Piso',             obs: d.pis_obs  },
    { nome: '🔲 Forro/Teto',       obs: d.for_obs  },
    { nome: '🖼️ Paredes',          obs: d.par_obs  },
  ];
  const pendencias = secoes
    .filter(s => s.obs && s.obs !== '—' && s.obs.trim() !== '')
    .map(s => `• ${s.nome}: ${s.obs}`)
    .join('\n');

  const assunto = `✅ Revisão registrada — ${d.apartamento} | Pousada Marie Claire`;
  const corpo = `
Olá, Leonardo!

Uma nova revisão geral foi registrada:

━━━━━━━━━━━━━━━━━━━━━━━━━━
🏨 ${d.apartamento}  (${d.andar})
📅 Data: ${d.data_revisao}
👷 Técnico: ${d.tecnico}
━━━━━━━━━━━━━━━━━━━━━━━━━━

${pendencias ? `🔧 PENDÊNCIAS ABERTAS NESTA REVISÃO:\n${pendencias}\n\n` : '✅ Nenhuma pendência registrada.\n\n'}
📅 Próxima revisão programada: ${d.proxima_revisao}

━━━━━━━━━━━━━━━━━━━━━━━━━━
📝 OBSERVAÇÕES GERAIS:
${d.observacoes || '—'}

Os dados foram salvos automaticamente no Google Sheets.

Atenciosamente,
Sistema de Controle de Revisões
Pousada Marie Claire
  `;

  GmailApp.sendEmail(EMAIL_DESTINO, assunto, corpo);
}

// ── Teste diagnóstico — execute manualmente no editor do Apps Script ─────
// Selecione esta função e clique em ▶ Executar para ver se a pasta do Drive está OK
function testarPastaDrive() {
  try {
    const pasta = DriveApp.getFolderById(PASTA_RAIZ_DRIVE);
    Logger.log('✅ Pasta encontrada: ' + pasta.getName());
    Logger.log('URL: ' + pasta.getUrl());

    // Testa criar subpasta
    const teste = pasta.createFolder('_TESTE_DELETAR');
    Logger.log('✅ Subpasta criada: ' + teste.getName());
    teste.setTrashed(true);
    Logger.log('✅ Subpasta deletada — tudo OK!');

    // Testa criar arquivo de texto
    const blob = Utilities.newBlob('teste', 'text/plain', 'teste.txt');
    const arq = pasta.createFile(blob);
    arq.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    Logger.log('✅ Arquivo criado: ' + arq.getUrl());
    arq.setTrashed(true);
    Logger.log('✅ Arquivo deletado — DRIVE FUNCIONANDO CORRETAMENTE');
  } catch(err) {
    Logger.log('❌ ERRO: ' + err.message);
  }
}

// ── Formatar data ─────────────────────────────────────────────────────────
function formatarData(valor) {
  if (!valor || valor === '—') return '—';
  const d = (valor instanceof Date) ? valor : new Date(valor);
  if (isNaN(d.getTime())) return String(valor);
  const dia = String(d.getDate()).padStart(2, '0');
  const mes = String(d.getMonth() + 1).padStart(2, '0');
  const ano = d.getFullYear();
  return `${dia}/${mes}/${ano}`;
}
