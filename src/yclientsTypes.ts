// Типы обмена данными между фронтендом и серверным прокси Yclients.
// Файл используется и сервером (server/yclients.ts), и браузером (src/yclientsApi.ts).

/** Строка финансовой операции в формате, совместимом с RawRecord (без id и Date). */
export interface YclientsFinanceRow {
  sum: number; // всегда положительное число (знак определяется статьёй)
  category: string; // «Назначение»
  counterparty: string; // «Контрагент»
  comment: string; // «Комментарий»
  dateStr: string; // «Дата» в формате ДД.ММ.ГГГГ
  serviceProduct?: string; // «Услуга/Товар»
}

/** Строка отчёта по услугам в формате, совместимом с RawServiceRecord (без id, rawIndex, revenue). */
export interface YclientsServiceRow {
  service: string;
  discount: number;
  points: number;
  subscriptions: number;
  certificates: number;
  clientAccounts: number;
  money: number;
}

export interface YclientsFinanceDiagnostics {
  transactionsTotal: number; // сколько операций вернул Yclients
  incomeCount: number;
  expenseCount: number;
  transfersSkipped: number; // переводы между кассами (в отчёте не участвуют)
  incomeLinkedToRecords: number; // доходы, сопоставленные с записями клиентов (есть услуга)
  incomeWithoutService: number; // доходы без сопоставленной услуги
  topCategories: { category: string; count: number; sum: number }[];
  warnings: string[];
}

export interface YclientsFinanceResponse {
  records: YclientsFinanceRow[];
  diagnostics: YclientsFinanceDiagnostics;
}

export interface YclientsServicesDiagnostics {
  recordsTotal: number; // записей клиентов в периоде
  recordsUsed: number; // записей, попавших в отчёт
  recordsSkipped: number; // удалённые / клиент не пришёл
  serviceLines: number;
  paymentBreakdownFromTransactions: number; // записей, где разбивка оплат взята из финансовых операций
  warnings: string[];
}

export interface YclientsServicesResponse {
  records: YclientsServiceRow[];
  diagnostics: YclientsServicesDiagnostics;
}

export interface YclientsStatusResponse {
  configured: boolean;
  companyId: string | null;
  missing: string[]; // какие переменные окружения не заданы
}
