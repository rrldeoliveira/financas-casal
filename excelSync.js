/**
 * excelSync.js — Integração com planilha Excel no OneDrive via Microsoft Graph
 *
 * Fluxo:
 *   1. Converte o share link em shareId (u! + base64url)
 *   2. Resolve shareId → driveId + itemId (com cache em localStorage)
 *   3. Auto-detecta tabela no workbook (ou cria "Lançamentos" se não existir)
 *   4. Lê colunas da tabela e mapeia o objeto de transação
 *   5. Faz POST para adicionar linha
 *   6. Se offline ou erro → enfileira localmente e sincroniza quando online
 *
 * Dependência: graphToken deve estar disponível no escopo global (setado pelo MSAL)
 */

// ─────────────────────────────────────────────────────────────
//  CONFIG — cole aqui o link compartilhado do OneDrive
// ─────────────────────────────────────────────────────────────
const EXCEL_SHARE_URL = 'https://1drv.ms/x/c/fd040843751e5a14/IQBS5g7KWOzMSqGAx1J0SAjkAcXj9Ne-HnbcEIQJK7gSViw?e=VBwEy5';

// Nome da tabela Excel — deixe '' para auto-detectar
const TABLE_NAME_CONFIG = '';

// Nome da aba e da tabela que serão CRIADOS caso não exista nenhuma tabela
const AUTO_SHEET_NAME = 'Lançamentos';
const AUTO_TABLE_NAME = 'TbLancamentos';
const AUTO_TABLE_COLS = ['Data', 'Pessoa', 'Tipo', 'Categoria', 'Descrição', 'Valor'];

const GRAPH = 'https://graph.microsoft.com/v1.0';
const LS_ITEM_ID    = 'excelSync_itemId';
const LS_DRIVE_ID   = 'excelSync_driveId';
const LS_TABLE_NAME = 'excelSync_tableName';
const LS_TABLE_COLS = 'excelSync_tableCols';
const LS_QUEUE      = 'excelSync_queue';

// ─────────────────────────────────────────────────────────────
//  1. SHARE LINK → shareId
// ─────────────────────────────────────────────────────────────
function shareUrlToId(url) {
  // Especificação Microsoft: "u!" + base64url(shareUrl)
  const b64 = btoa(encodeURIComponent(url).replace(/%([0-9A-F]{2})/g, (_, p) => String.fromCharCode(parseInt(p, 16))));
  return 'u!' + b64.replace(/\+/g, '-').replace(/\//g, '_').replace(/=/g, '');
}

// ─────────────────────────────────────────────────────────────
//  2. RESOLVE itemId + driveId (com cache)
// ─────────────────────────────────────────────────────────────
async function resolveShareLink(token) {
  const cachedItem  = localStorage.getItem(LS_ITEM_ID);
  const cachedDrive = localStorage.getItem(LS_DRIVE_ID);
  if (cachedItem && cachedDrive) {
    console.info('[excelSync] itemId (cache):', cachedItem, '| driveId:', cachedDrive);
    return { itemId: cachedItem, driveId: cachedDrive };
  }

  const shareId = shareUrlToId(EXCEL_SHARE_URL);
  console.info('[excelSync] Resolvendo share link → shareId:', shareId);
  const res = await fetch(`${GRAPH}/shares/${shareId}/driveItem`, {
    headers: { Authorization: `Bearer ${token}` }
  });

  if (!res.ok) {
    const err = await res.text();
    throw new Error(`Não foi possível acessar o arquivo compartilhado (${res.status}): ${err}`);
  }

  const data     = await res.json();
  const itemId   = data.id;
  const driveId  = data.parentReference?.driveId;

  if (!itemId || !driveId) throw new Error('Resposta do Graph sem itemId/driveId.');

  console.info('[excelSync] itemId resolvido:', itemId, '| driveId:', driveId);
  localStorage.setItem(LS_ITEM_ID,  itemId);
  localStorage.setItem(LS_DRIVE_ID, driveId);
  return { itemId, driveId };
}

// ─────────────────────────────────────────────────────────────
//  3. AUTO-DETECTA OU CRIA TABELA
// ─────────────────────────────────────────────────────────────
async function getOrCreateTable(token, driveId, itemId) {
  const cachedName = localStorage.getItem(LS_TABLE_NAME);
  const cachedCols = localStorage.getItem(LS_TABLE_COLS);
  if (cachedName && cachedCols) {
    const cols = JSON.parse(cachedCols);
    console.info('[excelSync] Tabela (cache):', cachedName, '| colunas:', cols);
    return { tableName: cachedName, columns: cols };
  }

  const base = `${GRAPH}/drives/${driveId}/items/${itemId}/workbook`;

  // Lista tabelas existentes
  const listRes = await fetch(`${base}/tables?$select=name`, {
    headers: { Authorization: `Bearer ${token}` }
  });
  if (!listRes.ok) throw new Error(`Erro ao listar tabelas (${listRes.status})`);

  const listData = await listRes.json();
  let tableName = TABLE_NAME_CONFIG || listData.value?.[0]?.name || '';

  if (!tableName) {
    // Nenhuma tabela encontrada → cria aba + tabela automaticamente
    console.info('[excelSync] Nenhuma tabela encontrada. Criando aba e tabela...');
    tableName = await createLancamentosTable(token, driveId, itemId, base);
  } else {
    console.info('[excelSync] Tabela escolhida:', tableName);
  }

  // Lê colunas da tabela escolhida
  const colRes = await fetch(`${base}/tables/${encodeURIComponent(tableName)}/columns?$select=name`, {
    headers: { Authorization: `Bearer ${token}` }
  });
  if (!colRes.ok) throw new Error(`Erro ao ler colunas da tabela "${tableName}" (${colRes.status})`);

  const colData = await colRes.json();
  const columns = colData.value.map(c => c.name);

  console.info('[excelSync] Colunas detectadas:', columns);
  localStorage.setItem(LS_TABLE_NAME, tableName);
  localStorage.setItem(LS_TABLE_COLS, JSON.stringify(columns));
  return { tableName, columns };
}

// Cria aba "Lançamentos" com tabela TbLancamentos caso não exista nada
async function createLancamentosTable(token, _driveId, _itemId, base) {
  // Cria worksheet
  const wsRes = await fetch(`${base}/worksheets/add`, {
    method: 'POST',
    headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
    body: JSON.stringify({ name: AUTO_SHEET_NAME })
  });
  // 409 = aba já existe, tudo bem
  if (!wsRes.ok && wsRes.status !== 409) {
    throw new Error(`Erro ao criar aba (${wsRes.status})`);
  }

  // Escreve cabeçalho na linha 1
  const headerRange = `A1:${String.fromCharCode(64 + AUTO_TABLE_COLS.length)}1`;
  await fetch(`${base}/worksheets('${encodeURIComponent(AUTO_SHEET_NAME)}')/range(address='${headerRange}')`, {
    method: 'PATCH',
    headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
    body: JSON.stringify({ values: [AUTO_TABLE_COLS] })
  });

  // Cria tabela a partir do range do cabeçalho
  const tblRes = await fetch(`${base}/worksheets('${encodeURIComponent(AUTO_SHEET_NAME)}')/tables/add`, {
    method: 'POST',
    headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
    body: JSON.stringify({ address: headerRange, hasHeaders: true })
  });
  if (!tblRes.ok) throw new Error(`Erro ao criar tabela (${tblRes.status})`);

  // Renomeia a tabela
  const tblData  = await tblRes.json();
  const tableId  = tblData.id || tblData.name;
  await fetch(`${base}/tables/${tableId}`, {
    method: 'PATCH',
    headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
    body: JSON.stringify({ name: AUTO_TABLE_NAME })
  });

  console.info(`[excelSync] Tabela "${AUTO_TABLE_NAME}" criada na aba "${AUTO_SHEET_NAME}".`);
  return AUTO_TABLE_NAME;
}

// ─────────────────────────────────────────────────────────────
//  4. MAPEAMENTO transação → linha da tabela
// ─────────────────────────────────────────────────────────────
function mapTxToRow(columns, tx) {
  return columns.map(col => {
    const c = col.toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '');
    if (c === 'data' || c === 'date')         return tx.date   || tx.data  || '';
    if (c === 'pessoa' || c === 'person'
     || c === 'quem'  || c === 'who')         return tx.person || tx.pessoa || '';
    if (c === 'tipo'  || c === 'type')         return tx.type   || tx.tipo  || '';
    if (c === 'categoria' || c === 'category') return tx.category || tx.categoria || '';
    if (c.startsWith('descri'))                return tx.description || tx.descricao || tx.desc || '';
    if (c === 'valor' || c === 'value'
     || c === 'amount'|| c === 'montante')     return parseFloat((tx.amount ?? tx.valor ?? 0).toString().replace(',', '.')) || 0;
    if (c === 'mes'   || c === 'month')        return tx.month  || tx.mes   || '';
    if (c === 'ano'   || c === 'year')         return tx.ano    || (tx.date ? tx.date.slice(0,4) : '');
    if (c === 'id')                            return tx.id     || '';
    return '';   // coluna desconhecida
  });
}

// ─────────────────────────────────────────────────────────────
//  5. ADICIONA LINHA NA TABELA
// ─────────────────────────────────────────────────────────────
async function appendRow(token, driveId, itemId, tableName, row) {
  const url = `${GRAPH}/drives/${driveId}/items/${itemId}/workbook/tables/${encodeURIComponent(tableName)}/rows/add`;
  const payload = { values: [row] };
  console.info('[excelSync] rows/add payload:', JSON.stringify(payload));
  const res = await fetch(url, {
    method: 'POST',
    headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
    body: JSON.stringify(payload)
  });
  if (!res.ok) {
    const body = await res.text();
    console.error('[excelSync] rows/add erro', res.status, body);
    throw new Error(`Erro ao adicionar linha (${res.status}): ${body}`);
  }
  const resData = await res.json();
  console.info('[excelSync] rows/add resposta (index):', resData.index ?? resData);
}

// ─────────────────────────────────────────────────────────────
//  6. FILA OFFLINE
// ─────────────────────────────────────────────────────────────
function getQueue()         { return JSON.parse(localStorage.getItem(LS_QUEUE) || '[]'); }
function saveQueue(q)       { localStorage.setItem(LS_QUEUE, JSON.stringify(q)); }
function enqueue(tx)        { const q = getQueue(); q.push(tx); saveQueue(q); }
function dequeue(tx)        { const q = getQueue().filter(t => t.id !== tx.id); saveQueue(q); }

// ─────────────────────────────────────────────────────────────
//  API PÚBLICA — chamada pelo app após salvar lançamento
// ─────────────────────────────────────────────────────────────

/**
 * syncExcelRow(tx)
 * Envia UMA transação para a tabela Excel.
 * Se offline ou falhar, enfileira e tenta depois.
 * Retorna { ok: true } ou { ok: false, error: string }
 */
async function syncExcelRow(tx) {
  if (!navigator.onLine) {
    enqueue(tx);
    console.info('[excelSync] Offline — lançamento enfileirado.');
    return { ok: false, error: 'offline' };
  }

  // graphToken vem do escopo global do app
  const token = typeof graphToken !== 'undefined' ? graphToken : null;
  if (!token) {
    enqueue(tx);
    console.info('[excelSync] Sem token — lançamento enfileirado.');
    return { ok: false, error: 'sem_token' };
  }

  try {
    const { itemId, driveId }    = await resolveShareLink(token);
    const { tableName, columns } = await getOrCreateTable(token, driveId, itemId);
    const row                    = mapTxToRow(columns, tx);
    console.info('[excelSync] Linha mapeada (sem token):', row);
    await appendRow(token, driveId, itemId, tableName, row);
    dequeue(tx);   // garante que não fique na fila se já estava
    console.info('[excelSync] ✅ Sincronização concluída.');
    return { ok: true };
  } catch (e) {
    enqueue(tx);
    console.error('[excelSync] Falha ao sincronizar:', e.message);
    return { ok: false, error: e.message };
  }
}

/**
 * flushExcelQueue()
 * Tenta sincronizar todos os lançamentos pendentes da fila offline.
 * Chamado automaticamente quando o browser volta a ficar online.
 */
async function flushExcelQueue() {
  const q = getQueue();
  if (!q.length) return;

  const token = typeof graphToken !== 'undefined' ? graphToken : null;
  if (!token) return;

  console.info(`[excelSync] Flush: ${q.length} lançamento(s) pendente(s).`);
  for (const tx of q) {
    const result = await syncExcelRow(tx);
    if (!result.ok && result.error !== 'offline') {
      // Falha não-offline: para e mantém o restante na fila
      break;
    }
  }
}

/**
 * resetExcelSyncCache()
 * Limpa o cache de itemId/tableName para forçar re-resolução.
 * Útil se o arquivo foi movido ou a tabela foi renomeada.
 */
function resetExcelSyncCache() {
  [LS_ITEM_ID, LS_DRIVE_ID, LS_TABLE_NAME, LS_TABLE_COLS].forEach(k => localStorage.removeItem(k));
  console.info('[excelSync] Cache limpo. Próxima sincronização vai re-detectar tudo.');
}

// Faz flush automático quando voltar online
window.addEventListener('online', flushExcelQueue);
