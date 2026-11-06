const APP_CONFIG = {
  propertyKey: 'SPREADSHEET_ID',
  tabs: {
    matrix: 'MATRIZ POR SETOR',
    almoxCatalog: 'CÓD. ALMOX',
    almoxQueue: 'CONTROLE ALMOX',
    codeMap: 'MAPEAMENTO ALMOX',
    movements: 'MOVIMENTAÇÕES EPI',
    destinationMaterials: 'DESTINO DE MATERIAIS'
  },
  collectivePpe: [
    'AVENTAL PLUMBÍFERO',
    'PROTETOR DE TIREÓIDE',
    'CINTA ERGONÔMICA PARA PROTEÇÃO LOMBAR',
    'PROTETOR AUDITIVO EM ESPUMA'
  ],
  occasionalCarePpe: ['ÓCULOS DE SEGURANÇA', 'PROTETOR AUDITIVO EM ESPUMA'],
  formSectors: ['CME', 'MANUTENÇÃO', 'TRANSPORTE INTERNO', 'IMAGEM-ENDOSCOPIA', 'IMAGEM-RESSONANCIA', 'CAF-SANEANTES', 'ROUPARIA', 'SESMT'],
  contractorSectors: ['MANUTENÇÃO', 'HOTELARIA', 'NAC', 'NTI', 'TRANSPORTE', 'IMAGEM', 'CME', 'GERENCIA DE ENFERMAGEM'],
  sizes: ['34', '35', '36', '37', '38', '39', '40', '41', '42', '43', '44', '45', '46'],
  otherMaterials: ['CONE', 'FITA ANTIDERRAPANTE', 'LUMINÁRIA', 'PLACA DE SINALIZAÇÃO', 'APOIO PARA OS PÉS', 'EXTINTOR', 'MANGUEIRA'],
  movementHeaders: [
    'ID MOVIMENTAÇÃO', 'DATA/HORA', 'OPERAÇÃO', 'DESTINATÁRIO', 'VÍNCULO', 'EMPRESA',
    'MATRÍCULA', 'NOME', 'SETOR', 'FUNÇÃO', 'TIPO DE ENTREGA', 'TIPO DE ITEM',
    'ITEM', 'TAMANHO', 'CA', 'QUANTIDADE', 'DEVOLUÇÃO PREVISTA', 'ID DO EMPRÉSTIMO',
    'RESPONSÁVEL', 'OBSERVAÇÃO'
  ]
};

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('CONTROLE DE EPIs E EPCs - SESMT HRC')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function getBootstrapData() {
  return {
    sectors: APP_CONFIG.formSectors,
    contractorSectors: APP_CONFIG.contractorSectors,
    catalog: getAlmoxCatalog_().map(x => ({code: x.code, description: x.description})),
    sizes: APP_CONFIG.sizes,
    otherMaterials: APP_CONFIG.otherMaterials.map(item => ({value:item,label:item === 'LUMINÁRIA' ? 'LUMINÁRIA DE EMERGÊNCIA' : item})).concat([{value:'OUTRO',label:'OUTRO'}]),
    collectivePpe: APP_CONFIG.collectivePpe
  };
}

function setupEpiModule() {
  const ss = getSpreadsheet_();
  getMovementsSheet_(true);
  ensureAlmoxQueueSheet_(ss);
  ensureAlmoxEditTriggers_(ss);
  ensureDestinationMaterialsSheet_(ss);
  let map = ss.getSheetByName(APP_CONFIG.tabs.codeMap);
  if (!map) {
    map = ss.insertSheet(APP_CONFIG.tabs.codeMap);
    const headers = ['TIPO', 'ITEM PADRONIZADO', 'TAMANHO', 'DESCRIÇÃO NO ALMOX', 'CÓDIGO ALMOX', 'ATIVO?'];
    map.getRange(1, 1, 1, headers.length).setValues([headers]);
    map.setFrozenRows(1);
    map.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#e8efff');
    const products = getAlmoxCatalog_();
    if (products.length) map.getRange(2, 4, products.length, 2).setValues(products.map(x => [x.description, x.code]));
  }
  return {ok:true, message:'Abas de movimentações e mapeamento do almox prontas. Complete as colunas TIPO, ITEM PADRONIZADO e TAMANHO para cruzar os códigos.'};
}

function getAllowedPpe(sector, role, recipientType) {
  const sheet = getSheet_(APP_CONFIG.tabs.destinationMaterials);
  if (!sheet) throw new Error('A aba DESTINO DE MATERIAIS não foi encontrada.');
  const last = sheet.getLastRow();
  if (last < 2) return [];
  const rows = sheet.getRange(2, 1, last - 1, 3).getDisplayValues();
  const wanted = recipientType === 'SETOR' ? 'COLETIVO' : 'INDIVIDUAL';
  const destination = String(sector || '').trim();
  return unique_(rows.filter(r => {
    if (!r[0] || !r[1] || !r[2] || isOtherMaterial_(r[0])) return false;
    const exits = normalizeExitType_(r[1]);
    if (!exits.includes(wanted)) return false;
    return String(r[2]).split(',').some(clause => destinationClauseApplies_(clause, destination, wanted));
  }).map(r => r[0]));
}

function normalizeExitType_(value) {
  // A coluna SAÍDA contém variações de digitação de INDIVIDUAL (ex.: INDIDIVUAL).
  return normalize_(value).replace(/INDIDIVUAL|INDIDIVIDUAL/g, 'INDIVIDUAL');
}

function destinationClauseApplies_(clause, sector, wanted) {
  const normalizedClause = normalize_(clause);
  const hasUsageQualifier = normalizedClause.includes('INDIVIDUAL') || normalizedClause.includes('COLETIVO');
  if (hasUsageQualifier && !normalizedClause.includes(wanted)) return false;
  let target = normalizedClause
    .replace(/\bINDIVIDUAL\b/g, ' ')
    .replace(/\bCOLETIVO\b/g, ' ')
    .replace(/\bSETOR\b/g, ' ')
    .replace(/\bTODOS OS SETORES\b/g, 'TODOS');
  if (target.includes('TODOS')) return true;
  const selected = normalize_(sector);
  if (!selected) return false;
  const aliases = {
    'HOTELARIA': ['ROUPARIA'],
    'ROUPARIA': ['HOTELARIA'],
    'TRANSPORTE INTERNO': ['TRANSPORTE'],
    'IMAGEM ENDOSCOPIA': ['IMAGEM'],
    'IMAGEM RESSONANCIA': ['IMAGEM']
  };
  const candidates = [selected].concat(aliases[selected] || []);
  return candidates.some(value => target.includes(value) || value.includes(target.trim()));
}

function getOpenLoans() {
  const sheet = getMovementsSheet_(false);
  if (!sheet || sheet.getLastRow() < 2) return [];
  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, APP_CONFIG.movementHeaders.length).getDisplayValues();
  const open = new Map();
  rows.forEach(r => {
    const op = r[2];
    const loanId = r[17];
    if (op === 'EMPRÉSTIMO' && loanId) {
      open.set(loanId, {id: loanId, item: r[12], kind: r[11], size: r[13], ca: r[14], qty: r[15], recipient: r[3], sector: r[8], due: r[16]});
    } else if (op === 'DEVOLUÇÃO' && loanId) {
      open.delete(loanId);
    }
  });
  return Array.from(open.values());
}

function submitMovement(payload) {
  validatePayload_(payload);
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    if (payload.operation === 'DEVOLUÇÃO') return registerReturn_(payload);

    const id = Utilities.getUuid().split('-')[0].toUpperCase();
    const timestamp = new Date();
    const person = payload.recipientType === 'COLABORADOR';
    const sector = payload.sector;
    const role = payload.role || '';
    const name = payload.name || '';
    const company = payload.link === 'TERCEIRIZADO' ? (payload.company || '') : '';
    const recipient = person ? 'COLABORADOR' : 'SETOR';
    const movementSheet = getMovementsSheet_(true);
    const queue = ensureAlmoxQueueSheet_(getSpreadsheet_());
    const catalog = getAlmoxCatalog_();
    const records = [];

    payload.items.forEach((line, index) => {
      const loanId = payload.operation === 'EMPRÉSTIMO' ? id + '-' + (index + 1) : '';
      records.push([
        id, timestamp, payload.operation, recipient, payload.link || '', company,
        person ? (payload.link === 'PRÓPRIO' ? payload.matricula : '') : '',
        person ? name : '', sector, role, payload.deliveryType || '', line.kind, line.item,
        line.size || 'NA', line.ca || '', Number(line.quantity),
        payload.dueDate || '', loanId, payload.responsible, payload.notes || ''
      ]);
    });

    movementSheet.getRange(movementSheet.getLastRow() + 1, 1, records.length, APP_CONFIG.movementHeaders.length).setValues(records);
    // A saída do sistema Almox recebe uma linha por item. O código é cruzado pelo catálogo;
    // quando a descrição não permite uma associação inequívoca, a linha fica marcada para conferência.
    if (queue) {
      const almoxRows = payload.items.map(line => {
        const code = resolveAlmoxCode_(line, catalog);
        return [timestamp, sector, line.kind, line.item, line.size || 'NA',
          Number(line.quantity), code, false];
      });
      const firstQueueRow = queue.getLastRow() + 1;
      queue.getRange(firstQueueRow, 1, almoxRows.length, 8).setValues(almoxRows);
      queue.getRange(firstQueueRow, 8, almoxRows.length, 1).insertCheckboxes();
    }
    return {ok: true, id, message: payload.operation === 'EMPRÉSTIMO'
      ? 'Empréstimo registrado. Guarde o identificador para registrar a devolução.'
      : 'Movimentação registrada e encaminhada para a lista de saída do almox.'};
  } finally {
    lock.releaseLock();
  }
}

function registerReturn_(payload) {
  const loans = getOpenLoans();
  const loan = loans.find(x => x.id === payload.loanId);
  if (!loan) throw new Error('Empréstimo não encontrado ou já devolvido. Atualize a lista e tente novamente.');
  const sheet = getMovementsSheet_(true);
  const id = Utilities.getUuid().split('-')[0].toUpperCase();
  const row = [
    id, new Date(), 'DEVOLUÇÃO', loan.recipient === 'SETOR' ? 'SETOR' : 'COLABORADOR', '', '', '', '',
    loan.sector, '', '', loan.kind || 'EPI', loan.item, loan.size || 'NA', loan.ca || '', Number(loan.qty),
    '', loan.id, payload.responsible, payload.notes || ''
  ];
  sheet.appendRow(row);
  return {ok: true, id, message: 'Devolução registrada.'};
}

function validatePayload_(p) {
  if (!p || !['ENTREGA', 'EMPRÉSTIMO', 'DEVOLUÇÃO'].includes(p.operation)) throw new Error('Selecione a operação.');
  if (!p.responsible || !String(p.responsible).trim()) throw new Error('Selecione quem está registrando a movimentação.');
  if (p.operation === 'DEVOLUÇÃO') {
    if (!p.loanId) throw new Error('Selecione o empréstimo que está sendo devolvido.');
    return;
  }
  if (!['COLABORADOR', 'SETOR'].includes(p.recipientType)) throw new Error('Selecione o destinatário.');
  if (!p.sector) throw new Error('Selecione ou informe o setor.');
  if (p.recipientType === 'COLABORADOR') {
    if (!['PRÓPRIO', 'TERCEIRIZADO'].includes(p.link)) throw new Error('Selecione o vínculo do colaborador.');
    if (p.link === 'PRÓPRIO' && !p.matricula) throw new Error('Informe a matrícula do colaborador próprio.');
    if (p.link === 'TERCEIRIZADO' && (!p.company || !p.name)) throw new Error('Informe a empresa e o nome do colaborador terceirizado.');
  }
  if (p.operation === 'ENTREGA' && !['INICIAL', 'SUBSTITUIÇÃO'].includes(p.deliveryType)) throw new Error('Selecione se a entrega é inicial ou substituição.');
  if (p.operation === 'EMPRÉSTIMO' && !p.dueDate) throw new Error('Informe a data prevista para devolução.');
  if (!Array.isArray(p.items) || !p.items.length) throw new Error('Adicione pelo menos um item.');
  const allowed = getAllowedPpe(p.sector, '', p.recipientType);
  p.items.forEach(line => {
    if (!line || !line.item || !['EPI', 'MATERIAL'].includes(line.kind)) throw new Error('Preencha o tipo e a descrição de cada item.');
    if (!Number.isInteger(Number(line.quantity)) || Number(line.quantity) < 1) throw new Error('A quantidade precisa ser um número inteiro maior que zero.');
    if (line.kind === 'EPI') {
      if (!allowed.includes(line.item)) throw new Error('O EPI selecionado não consta como permitido para esse setor e tipo de saída.');
      if (requiresNumericSize_(line.item)) {
        if (!APP_CONFIG.sizes.includes(String(line.size || ''))) throw new Error('Selecione o número do calçado/bota.');
      } else {
        line.size = 'NA';
      }
    } else {
      const isCustom = line.customMaterial === true;
      const canonical = String(line.item).trim();
      if (isCustom && !canonical) throw new Error('Informe a descrição do outro material.');
      if (!isCustom && !APP_CONFIG.otherMaterials.includes(canonical)) throw new Error('Selecione um material válido.');
      line.size = 'NA';
      line.ca = '';
    }
  });
}

function isOtherMaterial_(item) {
  const value = normalize_(item);
  return APP_CONFIG.otherMaterials.some(x => normalize_(x) === value) ||
    value === normalize_('LUMINÁRIA DE EMERGÊNCIA');
}

function requiresNumericSize_(item) {
  const value = normalize_(item);
  return /\bBOTA\b/.test(value) || /\bCALCADO\b/.test(value) || /\bBOTINA\b/.test(value);
}

function resolveAlmoxCode_(line, catalog) {
  const size = normalize_(line.size || 'NA');
  const map = getSheet_(APP_CONFIG.tabs.codeMap);
  if (map && map.getLastRow() > 1) {
    const rows = map.getRange(2, 1, map.getLastRow() - 1, 6).getDisplayValues();
    const mapped = rows.filter(r => normalize_(r[0]) === normalize_(line.kind) &&
      normalize_(r[1]) === normalize_(line.item) && normalize_(r[2] || 'NA') === size &&
      normalize_(r[5]) === 'SIM' && r[4]);
    if (mapped.length === 1) return String(mapped[0][4]).trim();
  }
  // Materials use the exact product description from CÓD. ALMOX. PPEs only auto-match
  // an exact catalog description; ambiguous or incomplete links are sent for review.
  const exact = catalog.filter(x => normalize_(x.description) === normalize_(line.item));
  return exact.length === 1 ? exact[0].code : 'CONFERIR';
}

function getMatrix_() {
  const sheet = getSheet_(APP_CONFIG.tabs.matrix);
  if (!sheet) throw new Error('A aba de matriz por setor não foi encontrada.');
  const last = sheet.getLastRow();
  const rows = last < 2 ? [] : sheet.getRange(2, 1, last - 1, 6).getDisplayValues()
    .map(r => r.map(v => String(v || '').trim())).filter(r => r[0]);
  return {rows};
}

function getAlmoxCatalog_() {
  const sheet = getSheet_(APP_CONFIG.tabs.almoxCatalog);
  if (!sheet) throw new Error('A aba de códigos do almox não foi encontrada.');
  const last = sheet.getLastRow();
  if (last < 3) return [];
  return sheet.getRange(3, 1, last - 2, 2).getDisplayValues()
    .filter(r => r[0] && r[1])
    .map(r => ({code: String(r[0]).trim(), description: String(r[1]).trim()}));
}

function ensureAlmoxQueueSheet_(ss) {
  let sheet = ss.getSheetByName(APP_CONFIG.tabs.almoxQueue);
  if (!sheet) sheet = ss.insertSheet(APP_CONFIG.tabs.almoxQueue);
  const headers = ['DATA/HORA', 'SETOR', 'TIPO', 'ITEM', 'TAMANHO', 'QUANTIDADE', 'CÓDIGO ALMOX', 'LANÇADO NO ALMOX?'];
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#d9ead3');
    sheet.setColumnWidths(1, headers.length, 150);
    sheet.setColumnWidth(4, 320);
  }
  return sheet;
}

function ensureAlmoxEditTriggers_(ss) {
  const triggers = ScriptApp.getProjectTriggers();
  if (!triggers.some(t => t.getHandlerFunction() === 'handleAlmoxEdit_')) {
    ScriptApp.newTrigger('handleAlmoxEdit_').forSpreadsheet(ss).onEdit().create();
  }
  if (!triggers.some(t => t.getHandlerFunction() === 'addAlmoxMenu_')) {
    ScriptApp.newTrigger('addAlmoxMenu_').forSpreadsheet(ss).onOpen().create();
  }
}

function addAlmoxMenu_() {
  SpreadsheetApp.getUi().createMenu('Controle Almox')
    .addItem('Exibir todos os registros (inclui lançados)', 'showAllAlmoxRows_')
    .addToUi();
}

function showAllAlmoxRows_() {
  const sheet = getSheet_(APP_CONFIG.tabs.almoxQueue);
  if (sheet && sheet.getMaxRows() > 1) sheet.showRows(2, sheet.getMaxRows() - 1);
}

function handleAlmoxEdit_(e) {
  if (!e || !e.range) return;
  const sheet = e.range.getSheet();
  if (sheet.getName() !== APP_CONFIG.tabs.almoxQueue || e.range.getColumn() !== 8 || e.range.getRow() < 2) return;
  if (String(e.value).toUpperCase() === 'TRUE') sheet.hideRows(e.range.getRow());
  else sheet.showRows(e.range.getRow());
}

function ensureDestinationMaterialsSheet_(ss) {
  let sheet = ss.getSheetByName(APP_CONFIG.tabs.destinationMaterials);
  if (!sheet) {
    sheet = ss.insertSheet(APP_CONFIG.tabs.destinationMaterials);
    const headers = ['DESCRIÇÃO', 'SAÍDA', 'DESTINO'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#e8efff');
  }
  return sheet;
}

function getMovementsSheet_(create) {
  const ss = getSpreadsheet_();
  let sheet = ss.getSheetByName(APP_CONFIG.tabs.movements);
  if (!sheet && create) {
    sheet = ss.insertSheet(APP_CONFIG.tabs.movements);
    sheet.getRange(1, 1, 1, APP_CONFIG.movementHeaders.length).setValues([APP_CONFIG.movementHeaders]);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, APP_CONFIG.movementHeaders.length).setFontWeight('bold').setBackground('#e8efff');
  }
  return sheet;
}

function getSpreadsheet_() {
  const configured = PropertiesService.getScriptProperties().getProperty(APP_CONFIG.propertyKey);
  if (!configured) throw new Error('Configure a propriedade SPREADSHEET_ID nas propriedades do projeto Apps Script.');
  // Aceita tanto o ID puro quanto a URL completa da planilha.
  const match = String(configured).trim().match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
  const id = match ? match[1] : String(configured).trim();
  return SpreadsheetApp.openById(id);
}

function getSheet_(name) { return getSpreadsheet_().getSheetByName(name); }
function unique_(values) { return Array.from(new Set(values.map(v => String(v).trim()).filter(Boolean))).sort((a, b) => a.localeCompare(b, 'pt-BR')); }
function normalize_(value) { return String(value || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').toUpperCase().replace(/[^A-Z0-9]+/g, ' ').trim().replace(/\s+/g, ' '); }
