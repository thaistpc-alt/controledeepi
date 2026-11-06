const APP_CONFIG = {
  propertyKey: 'SPREADSHEET_ID',
  tabs: {
    matrix: 'MATRIZ POR SETOR',
    almoxCatalog: 'CÓD. ALMOX',
    almoxQueue: 'LANÇAR ALMOX',
    codeMap: 'MAPEAMENTO ALMOX',
    movements: 'MOVIMENTAÇÕES EPI'
  },
  collectivePpe: [
    'AVENTAL PLUMBÍFERO',
    'PROTETOR DE TIREÓIDE',
    'CINTA ERGONÔMICA PARA PROTEÇÃO LOMBAR',
    'PROTETOR AUDITIVO EM ESPUMA'
  ],
  occasionalCarePpe: ['ÓCULOS DE SEGURANÇA', 'PROTETOR AUDITIVO EM ESPUMA'],
  sizes: ['NA', 'ÚNICO', 'PP', 'P', 'M', 'G', 'GG', 'XG', '34', '35', '36', '37', '38', '39', '40', '41', '42', '43', '44', '45', '46'],
  movementHeaders: [
    'ID MOVIMENTAÇÃO', 'DATA/HORA', 'OPERAÇÃO', 'DESTINATÁRIO', 'VÍNCULO', 'EMPRESA',
    'MATRÍCULA', 'NOME', 'SETOR', 'FUNÇÃO', 'TIPO DE ENTREGA', 'TIPO DE ITEM',
    'ITEM', 'TAMANHO', 'CA', 'QUANTIDADE', 'CÓD. ALMOX', 'STATUS',
    'DEVOLUÇÃO PREVISTA', 'ID DO EMPRÉSTIMO', 'RESPONSÁVEL', 'OBSERVAÇÃO'
  ]
};

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Saída e empréstimo de itens')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function getBootstrapData() {
  const matrix = getMatrix_();
  const sectors = unique_(matrix.rows.map(r => r[0]).filter(Boolean));
  const functionsBySector = {};
  matrix.rows.forEach(r => {
    if (!r[0] || !r[1]) return;
    functionsBySector[r[0]] = functionsBySector[r[0]] || [];
    functionsBySector[r[0]].push(r[1]);
  });
  Object.keys(functionsBySector).forEach(k => functionsBySector[k] = unique_(functionsBySector[k]));
  return {
    sectors,
    functionsBySector,
    catalog: getAlmoxCatalog_().map(x => ({code: x.code, description: x.description})),
    sizes: APP_CONFIG.sizes,
    collectivePpe: APP_CONFIG.collectivePpe
  };
}

function setupEpiModule() {
  const ss = getSpreadsheet_();
  getMovementsSheet_(true);
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

function lookupEmployee(matricula) {
  const key = normalize_(matricula).replace(/\D/g, '');
  if (!key) return null;
  const rows = getMatrix_().rows;
  const match = rows.find(r => normalize_(r[2]).replace(/\D/g, '') === key && normalize_(r[4]) === 'ATIVO');
  if (!match) return null;
  return {matricula: String(match[2]), name: match[3], sector: match[0], role: match[1]};
}

function getAllowedPpe(sector, role, recipientType) {
  const s = String(sector || '').trim();
  const f = String(role || '').trim();
  if (recipientType === 'SETOR') return APP_CONFIG.collectivePpe.slice();
  const rows = getMatrix_().rows.filter(r => r[0] === s && (!f || r[1] === f));
  let items = unique_(rows.map(r => r[5]).filter(Boolean));
  if (!items.length) items = APP_CONFIG.occasionalCarePpe.slice();
  return items;
}

function getOpenLoans() {
  const sheet = getMovementsSheet_(false);
  if (!sheet || sheet.getLastRow() < 2) return [];
  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, APP_CONFIG.movementHeaders.length).getDisplayValues();
  const open = new Map();
  rows.forEach(r => {
    const op = r[2];
    const status = r[17];
    if (op === 'EMPRÉSTIMO' && status === 'EM ABERTO') {
      open.set(r[19], {id: r[19], item: r[12], kind: r[11], size: r[13], ca: r[14], qty: r[15], recipient: r[3], sector: r[8], due: r[18]});
    } else if (op === 'DEVOLUÇÃO' && r[19]) {
      open.delete(r[19]);
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
    const employee = person && payload.link === 'PRÓPRIO' ? lookupEmployee(payload.matricula) : null;
    const sector = (employee && employee.sector) || payload.sector;
    const role = (employee && employee.role) || payload.role || '';
    const name = (employee && employee.name) || payload.name || '';
    const company = payload.link === 'TERCEIRIZADO' ? (payload.company || '') : '';
    const recipient = person ? 'COLABORADOR' : 'SETOR';
    const movementSheet = getMovementsSheet_(true);
    const queue = getSheet_(APP_CONFIG.tabs.almoxQueue);
    const catalog = getAlmoxCatalog_();
    const records = [];

    payload.items.forEach((line, index) => {
      const code = resolveAlmoxCode_(line, catalog);
      const loanId = payload.operation === 'EMPRÉSTIMO' ? id + '-' + (index + 1) : '';
      const status = payload.operation === 'EMPRÉSTIMO' ? 'EM ABERTO' : (code === 'CONFERIR' ? 'CONFERIR CÓDIGO' : 'PENDENTE NO ALMOX');
      records.push([
        id, timestamp, payload.operation, recipient, payload.link || '', company,
        person ? (payload.link === 'PRÓPRIO' ? payload.matricula : '') : '',
        person ? name : '', sector, role, payload.deliveryType || '', line.kind, line.item,
        line.size || 'NA', line.ca || '', Number(line.quantity), code, status,
        payload.dueDate || '', loanId, payload.responsible, payload.notes || ''
      ]);
    });

    movementSheet.getRange(movementSheet.getLastRow() + 1, 1, records.length, APP_CONFIG.movementHeaders.length).setValues(records);
    // A saída do sistema Almox recebe uma linha por item. O código é cruzado pelo catálogo;
    // quando a descrição não permite uma associação inequívoca, a linha fica marcada para conferência.
    if (queue) {
      const almoxRows = payload.items.map(line => {
        const code = resolveAlmoxCode_(line, catalog);
        return [timestamp.getFullYear(), timestamp.getMonth() + 1, timestamp, sector, line.item,
          line.size || 'NA', String(line.quantity), code, ''];
      });
      queue.getRange(queue.getLastRow() + 1, 1, almoxRows.length, 9).setValues(almoxRows);
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
    loan.sector, '', '', loan.kind || 'EPI', loan.item, loan.size || 'NA', loan.ca || '', Number(loan.qty), '', 'DEVOLVIDO', '',
    loan.id, payload.responsible, payload.notes || ''
  ];
  sheet.appendRow(row);
  return {ok: true, id, message: 'Devolução registrada.'};
}

function validatePayload_(p) {
  if (!p || !['ENTREGA', 'EMPRÉSTIMO', 'DEVOLUÇÃO'].includes(p.operation)) throw new Error('Selecione a operação.');
  if (!p.responsible || !String(p.responsible).trim()) throw new Error('Informe quem está registrando a movimentação.');
  if (p.operation === 'DEVOLUÇÃO') {
    if (!p.loanId) throw new Error('Selecione o empréstimo que está sendo devolvido.');
    return;
  }
  if (!['COLABORADOR', 'SETOR'].includes(p.recipientType)) throw new Error('Selecione o destinatário.');
  if (!p.sector) throw new Error('Selecione ou informe o setor.');
  if (p.recipientType === 'COLABORADOR') {
    if (!['PRÓPRIO', 'TERCEIRIZADO'].includes(p.link)) throw new Error('Selecione o vínculo do colaborador.');
    if (p.link === 'PRÓPRIO' && (!p.matricula || !p.name)) throw new Error('Informe a matrícula e o nome do colaborador próprio.');
    if (p.link === 'TERCEIRIZADO' && (!p.company || !p.name)) throw new Error('Informe a empresa e o nome do colaborador terceirizado.');
  }
  if (p.operation === 'ENTREGA' && !p.deliveryType) throw new Error('Informe se a entrega é inicial, substituição ou para uso coletivo.');
  if (p.operation === 'EMPRÉSTIMO' && !p.dueDate) throw new Error('Informe a data prevista para devolução.');
  if (!Array.isArray(p.items) || !p.items.length) throw new Error('Adicione pelo menos um item.');
  const allowed = p.recipientType === 'SETOR' ? APP_CONFIG.collectivePpe : getAllowedPpe(p.sector, '');
  p.items.forEach(line => {
    if (!line.item || !['EPI', 'MATERIAL'].includes(line.kind)) throw new Error('Confira o tipo e a descrição de cada item.');
    if (!Number.isInteger(Number(line.quantity)) || Number(line.quantity) < 1) throw new Error('A quantidade precisa ser um número inteiro maior que zero.');
    if (!line.size) line.size = 'NA';
    if (line.kind === 'EPI' && !allowed.includes(line.item)) throw new Error('O EPI selecionado não consta como permitido para esse setor/função.');
  });
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
