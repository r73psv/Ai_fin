import type {
  YclientsFinanceDiagnostics,
  YclientsFinanceRow,
  YclientsServiceRow,
  YclientsServicesDiagnostics,
} from '../src/yclientsTypes';

const API_BASE = 'https://api.yclients.com/api/v1';

// ---------------------------------------------------------------------------
// Конфигурация (из переменных окружения; токены никогда не уходят в браузер)
// ---------------------------------------------------------------------------

export interface YclientsConfig {
  partnerToken: string;
  userToken: string;
  companyId: string;
}

export interface YclientsConfigState {
  configured: boolean;
  companyId: string | null;
  missing: string[];
}

export function getConfigState(env: NodeJS.ProcessEnv = process.env): YclientsConfigState {
  const missing: string[] = [];
  if (!env.YCLIENTS_PARTNER_TOKEN) missing.push('YCLIENTS_PARTNER_TOKEN');
  if (!env.YCLIENTS_COMPANY_ID) missing.push('YCLIENTS_COMPANY_ID');
  const hasUserToken = !!env.YCLIENTS_USER_TOKEN;
  const hasLogin = !!env.YCLIENTS_LOGIN && !!env.YCLIENTS_PASSWORD;
  if (!hasUserToken && !hasLogin) {
    missing.push('YCLIENTS_USER_TOKEN (или пара YCLIENTS_LOGIN + YCLIENTS_PASSWORD)');
  }
  return {
    configured: missing.length === 0,
    companyId: env.YCLIENTS_COMPANY_ID || null,
    missing,
  };
}

export class YclientsError extends Error {
  status: number;
  constructor(message: string, status = 502) {
    super(message);
    this.name = 'YclientsError';
    this.status = status;
  }
}

let cachedLoginToken: string | null = null;

/** Собирает конфиг; если задан только логин/пароль — получает пользовательский токен через POST /auth. */
export async function resolveConfig(env: NodeJS.ProcessEnv = process.env): Promise<YclientsConfig> {
  const state = getConfigState(env);
  if (!state.configured) {
    throw new YclientsError(
      `Интеграция с Yclients не настроена. Не заданы переменные окружения: ${state.missing.join(', ')}.`,
      503
    );
  }

  const partnerToken = env.YCLIENTS_PARTNER_TOKEN as string;
  const companyId = String(env.YCLIENTS_COMPANY_ID).trim();
  if (!/^\d+$/.test(companyId)) {
    throw new YclientsError('YCLIENTS_COMPANY_ID должен быть числом (ID филиала в Yclients).', 503);
  }

  let userToken = env.YCLIENTS_USER_TOKEN || cachedLoginToken || '';
  if (!userToken) {
    const res = await fetch(`${API_BASE}/auth`, {
      method: 'POST',
      headers: {
        Authorization: `Bearer ${partnerToken}`,
        Accept: 'application/vnd.yclients.v2+json',
        'Content-Type': 'application/json',
      },
      body: JSON.stringify({ login: env.YCLIENTS_LOGIN, password: env.YCLIENTS_PASSWORD }),
    });
    const json: any = await res.json().catch(() => null);
    const token = json?.data?.user_token;
    if (!res.ok || !token) {
      throw new YclientsError(
        `Не удалось авторизоваться в Yclients по логину и паролю: ${json?.meta?.message || res.status}`,
        401
      );
    }
    cachedLoginToken = token;
    userToken = token;
  }

  return { partnerToken, userToken, companyId };
}

// ---------------------------------------------------------------------------
// HTTP-слой
// ---------------------------------------------------------------------------

const sleep = (ms: number) => new Promise((r) => setTimeout(r, ms));

async function yclientsGet(
  cfg: YclientsConfig,
  path: string,
  query: Record<string, string | number> = {}
): Promise<any> {
  const url = new URL(API_BASE + path);
  Object.entries(query).forEach(([k, v]) => url.searchParams.set(k, String(v)));

  for (let attempt = 0; attempt < 4; attempt++) {
    const res = await fetch(url, {
      headers: {
        Authorization: `Bearer ${cfg.partnerToken}, User ${cfg.userToken}`,
        Accept: 'application/vnd.yclients.v2+json',
        'Content-Type': 'application/json',
      },
    });

    // Ограничение частоты запросов — ждём и повторяем
    if (res.status === 429) {
      await sleep(1000 * (attempt + 1));
      continue;
    }

    const text = await res.text();
    let json: any = null;
    try {
      json = JSON.parse(text);
    } catch {
      /* не JSON */
    }

    if (!res.ok || json?.success === false) {
      const apiMessage = json?.meta?.message || json?.message;
      if (res.status === 401) {
        cachedLoginToken = null;
        throw new YclientsError(`Yclients отклонил токены (401). ${apiMessage || 'Проверьте YCLIENTS_PARTNER_TOKEN и YCLIENTS_USER_TOKEN.'}`, 401);
      }
      if (res.status === 403) {
        throw new YclientsError(`Нет доступа к данным филиала (403). ${apiMessage || 'Проверьте права пользователя на финансы и записи.'}`, 403);
      }
      throw new YclientsError(`Ошибка Yclients API (${res.status}): ${apiMessage || 'неизвестная ошибка'}`, 502);
    }
    return json;
  }
  throw new YclientsError('Yclients API: превышен лимит запросов (429), попробуйте позже.', 429);
}

/** Постраничная загрузка: идём, пока страницы не закончатся или не перестанут приходить новые id. */
async function fetchAllPages<T extends { id?: number }>(
  cfg: YclientsConfig,
  path: string,
  query: Record<string, string | number>,
  pageSize = 200,
  maxPages = 500
): Promise<T[]> {
  const all: T[] = [];
  const seen = new Set<number>();
  for (let page = 1; page <= maxPages; page++) {
    const json = await yclientsGet(cfg, path, { ...query, page, count: pageSize });
    const data: T[] = Array.isArray(json?.data) ? json.data : [];
    if (data.length === 0) break;

    let added = 0;
    for (const item of data) {
      const id = item?.id;
      if (typeof id === 'number') {
        if (seen.has(id)) continue;
        seen.add(id);
      }
      all.push(item);
      added++;
    }
    // API не поддержал пагинацию и вернул те же строки — выходим, чтобы не зациклиться
    if (added === 0) break;
  }
  return all;
}

// ---------------------------------------------------------------------------
// Сырые структуры Yclients (описаны с запасом: все поля необязательные)
// ---------------------------------------------------------------------------

export interface YcTransaction {
  id?: number;
  date?: string; // "2026-05-10 14:00:00"
  amount?: number;
  comment?: string;
  record_id?: number | null;
  expense?: { id?: number; title?: string } | null;
  account?: { id?: number; title?: string } | null;
  client?: { name?: string } | null;
  master?: { name?: string } | null;
  supplier?: { title?: string } | null;
}

export interface YcRecordService {
  title?: string;
  cost?: number;
  cost_to_pay?: number;
  first_cost?: number;
  amount?: number;
}

export interface YcRecord {
  id?: number;
  deleted?: boolean;
  attendance?: number; // 1 — клиент пришёл
  paid_full?: number;
  services?: YcRecordService[];
  finance_transactions?: { amount?: number; account?: { title?: string } | null }[];
}

export async function fetchTransactions(cfg: YclientsConfig, startDate: string, endDate: string): Promise<YcTransaction[]> {
  return fetchAllPages<YcTransaction>(cfg, `/transactions/${cfg.companyId}`, {
    start_date: startDate,
    end_date: endDate,
  });
}

export async function fetchRecords(
  cfg: YclientsConfig,
  startDate: string,
  endDate: string,
  withFinanceTransactions = false
): Promise<YcRecord[]> {
  const query: Record<string, string | number> = { start_date: startDate, end_date: endDate };
  if (withFinanceTransactions) query.include_finance_transactions = 1;
  return fetchAllPages<YcRecord>(cfg, `/records/${cfg.companyId}`, query);
}

// ---------------------------------------------------------------------------
// Вспомогательные функции
// ---------------------------------------------------------------------------

const round2 = (n: number) => Math.round(n * 100) / 100;

/** "2026-05-10 14:00:00" | "2026-05-10T14:00:00+03:00" -> "10.05.2026" */
export function toDdMmYyyy(raw: string | undefined | null): string {
  const m = /^(\d{4})-(\d{2})-(\d{2})/.exec(String(raw || ''));
  return m ? `${m[3]}.${m[2]}.${m[1]}` : '';
}

function lineNet(s: YcRecordService): number {
  const v = typeof s.cost_to_pay === 'number' ? s.cost_to_pay : s.cost;
  return typeof v === 'number' && isFinite(v) ? Math.max(0, v) : 0;
}

// ---------------------------------------------------------------------------
// Финансовые операции -> формат «Вкладки 1»
// ---------------------------------------------------------------------------

export interface RecordServiceLine {
  title: string;
  net: number;
}

/** Справочник: id записи -> её услуги (для расшифровки доходов по услугам). */
export function buildRecordServicesMap(records: YcRecord[]): Map<number, RecordServiceLine[]> {
  const map = new Map<number, RecordServiceLine[]>();
  for (const r of records) {
    if (typeof r.id !== 'number' || r.deleted) continue;
    const lines = (r.services || [])
      .map((s) => ({ title: String(s.title || '').trim(), net: lineNet(s) }))
      .filter((l) => l.title);
    if (lines.length) map.set(r.id, lines);
  }
  return map;
}

/**
 * Приводит название статьи Yclients к названиям статей шаблона отчёта.
 * Ручные соответствия можно задать переменной YCLIENTS_CATEGORY_MAP (JSON: {"Статья в Yclients": "Статья в отчёте"}).
 */
function mapCategory(tx: YcTransaction, isIncome: boolean, overrides: Record<string, string>): string {
  const title = String(tx.expense?.title || '').trim();

  if (title && overrides[title]) return overrides[title];
  if (/перевод/i.test(title)) return 'Перевод средств';

  if (isIncome) {
    if (/абонемент/i.test(title)) return 'Доход от продажи абонементов';
    if (/сертификат/i.test(title)) return 'Выручка с сертификатов';
    if (/пополнени|депозит/i.test(title)) return 'Пополнение счёта';
    if (/товар/i.test(title)) return 'Выручка с товаров';
    if (tx.record_id || /услуг/i.test(title)) return 'Выручка с услуг';
    return title || 'Прочие доходы';
  }
  return title || 'Без статьи';
}

export function parseCategoryOverrides(raw: string | undefined): Record<string, string> {
  if (!raw) return {};
  try {
    const parsed = JSON.parse(raw);
    return parsed && typeof parsed === 'object' ? parsed : {};
  } catch {
    return {};
  }
}

export function mapTransactionsToFinanceRows(
  txs: YcTransaction[],
  recordServices: Map<number, RecordServiceLine[]>,
  overrides: Record<string, string> = {}
): { records: YclientsFinanceRow[]; diagnostics: YclientsFinanceDiagnostics } {
  const rows: YclientsFinanceRow[] = [];
  const catStats = new Map<string, { count: number; sum: number }>();
  let incomeCount = 0;
  let expenseCount = 0;
  let transfersSkipped = 0;
  let incomeLinked = 0;
  let incomeNoService = 0;

  const push = (row: YclientsFinanceRow) => {
    rows.push(row);
    const s = catStats.get(row.category) || { count: 0, sum: 0 };
    s.count += 1;
    s.sum += row.sum;
    catStats.set(row.category, s);
  };

  for (const tx of txs) {
    const amount = Number(tx.amount);
    if (!isFinite(amount) || amount === 0) continue;

    // Знак суммы в Yclients: приход > 0, расход < 0
    const isIncome = amount > 0;
    const category = mapCategory(tx, isIncome, overrides);
    const dateStr = toDdMmYyyy(tx.date);
    const counterparty = String(tx.master?.name || tx.supplier?.title || tx.client?.name || '').trim();
    const comment = String(tx.comment || '').trim();
    const abs = Math.abs(amount);

    if (category === 'Перевод средств') {
      transfersSkipped++;
      // Переводы всё равно передаём — движок сам исключает их из расчёта P&L
      push({ sum: abs, category, counterparty, comment, dateStr });
      continue;
    }

    if (isIncome) {
      incomeCount++;
      const lines = tx.record_id ? recordServices.get(tx.record_id) : undefined;

      if (category === 'Выручка с услуг' && lines && lines.length > 0) {
        incomeLinked++;
        // Платёж по записи с несколькими услугами делим пропорционально стоимости услуг
        const totalNet = lines.reduce((acc, l) => acc + l.net, 0);
        let allocated = 0;
        lines.forEach((l, i) => {
          const share = totalNet > 0 ? l.net / totalNet : 1 / lines.length;
          const part = i === lines.length - 1 ? round2(abs - allocated) : round2(abs * share);
          allocated += part;
          if (part > 0) push({ sum: part, category, counterparty, comment, dateStr, serviceProduct: l.title });
        });
      } else {
        if (category === 'Выручка с услуг') incomeNoService++;
        push({ sum: abs, category, counterparty, comment, dateStr });
      }
    } else {
      expenseCount++;
      push({ sum: abs, category, counterparty, comment, dateStr });
    }
  }

  const topCategories = Array.from(catStats.entries())
    .map(([category, v]) => ({ category, count: v.count, sum: round2(v.sum) }))
    .sort((a, b) => b.sum - a.sum)
    .slice(0, 15);

  return {
    records: rows,
    diagnostics: {
      transactionsTotal: txs.length,
      incomeCount,
      expenseCount,
      transfersSkipped,
      incomeLinkedToRecords: incomeLinked,
      incomeWithoutService: incomeNoService,
      topCategories,
      warnings: [],
    },
  };
}

// ---------------------------------------------------------------------------
// Записи клиентов -> отчёт по услугам («Вкладка 2»)
// ---------------------------------------------------------------------------

type PayBucket = 'points' | 'subscriptions' | 'certificates' | 'clientAccounts' | 'money';

function bucketForAccount(title: string): PayBucket {
  const t = title.toLowerCase();
  if (/балл|бонус/.test(t)) return 'points';
  if (/абонемент/.test(t)) return 'subscriptions';
  if (/сертификат/.test(t)) return 'certificates';
  if (/депозит|счет клиента|счёт клиента|со счет|со счёт|баланс/.test(t)) return 'clientAccounts';
  return 'money';
}

export function mapRecordsToServiceRows(
  records: YcRecord[],
  options: { onlyAttended?: boolean } = {}
): { records: YclientsServiceRow[]; diagnostics: YclientsServicesDiagnostics } {
  const onlyAttended = options.onlyAttended !== false;
  const rows: YclientsServiceRow[] = [];
  let used = 0;
  let skipped = 0;
  let serviceLines = 0;
  let withBreakdown = 0;

  for (const rec of records) {
    const attended = rec.attendance === 1 || rec.paid_full === 1;
    if (rec.deleted || (onlyAttended && !attended)) {
      skipped++;
      continue;
    }
    const services = (rec.services || []).filter((s) => String(s.title || '').trim());
    if (services.length === 0) {
      skipped++;
      continue;
    }
    used++;

    // Разбивка способов оплаты по записи (если Yclients вернул финансовые операции записи)
    const buckets: Record<PayBucket, number> = { points: 0, subscriptions: 0, certificates: 0, clientAccounts: 0, money: 0 };
    let paidTotal = 0;
    for (const ft of rec.finance_transactions || []) {
      const a = Math.abs(Number(ft.amount) || 0);
      if (a <= 0) continue;
      buckets[bucketForAccount(String(ft.account?.title || ''))] += a;
      paidTotal += a;
    }
    const hasBreakdown = paidTotal > 0;
    if (hasBreakdown) withBreakdown++;
    const share = (b: PayBucket) => (hasBreakdown ? buckets[b] / paidTotal : b === 'money' ? 1 : 0);

    for (const s of services) {
      serviceLines++;
      const net = lineNet(s);
      const qty = Math.max(1, Math.round(Number(s.amount) || 1));
      const gross = typeof s.first_cost === 'number' && isFinite(s.first_cost) ? s.first_cost * qty : net;
      const discount = Math.max(0, gross - net);

      // Каждое оказание услуги — отдельная строка (приложение считает количество по строкам)
      for (let i = 0; i < qty; i++) {
        rows.push({
          service: String(s.title).trim(),
          discount: round2(discount / qty),
          points: round2((net * share('points')) / qty),
          subscriptions: round2((net * share('subscriptions')) / qty),
          certificates: round2((net * share('certificates')) / qty),
          clientAccounts: round2((net * share('clientAccounts')) / qty),
          money: round2((net * share('money')) / qty),
        });
      }
    }
  }

  return {
    records: rows,
    diagnostics: {
      recordsTotal: records.length,
      recordsUsed: used,
      recordsSkipped: skipped,
      serviceLines,
      paymentBreakdownFromTransactions: withBreakdown,
      warnings: [],
    },
  };
}
