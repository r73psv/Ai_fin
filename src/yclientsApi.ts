import type {
  YclientsFinanceResponse,
  YclientsServicesResponse,
  YclientsStatusResponse,
} from './yclientsTypes';

async function parseJsonOrThrow<T>(res: Response): Promise<T> {
  let body: any = null;
  try {
    body = await res.json();
  } catch {
    /* ответ не JSON */
  }
  if (!res.ok) {
    throw new Error(body?.error || `Сервер вернул ошибку ${res.status}. Убедитесь, что приложение запущено через «npm run dev» (серверная часть нужна для связи с Yclients).`);
  }
  return body as T;
}

export async function fetchYclientsStatus(): Promise<YclientsStatusResponse> {
  const res = await fetch('/api/yclients/status');
  return parseJsonOrThrow<YclientsStatusResponse>(res);
}

export async function fetchYclientsFinance(startDate: string, endDate: string): Promise<YclientsFinanceResponse> {
  const res = await fetch('/api/yclients/transactions', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ startDate, endDate }),
  });
  return parseJsonOrThrow<YclientsFinanceResponse>(res);
}

export async function fetchYclientsServices(startDate: string, endDate: string): Promise<YclientsServicesResponse> {
  const res = await fetch('/api/yclients/services', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ startDate, endDate }),
  });
  return parseJsonOrThrow<YclientsServicesResponse>(res);
}

/** "2026-05-10" -> "10.05.2026" */
export function isoToRu(iso: string): string {
  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(iso);
  return m ? `${m[3]}.${m[2]}.${m[1]}` : iso;
}
