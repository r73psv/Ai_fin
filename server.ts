import 'dotenv/config';
import express from 'express';
import path from 'path';
import { fileURLToPath } from 'url';
import {
  YclientsError,
  buildRecordServicesMap,
  fetchRecords,
  fetchTransactions,
  getConfigState,
  mapRecordsToServiceRows,
  mapTransactionsToFinanceRows,
  parseCategoryOverrides,
  resolveConfig,
} from './server/yclients';
import type {
  YclientsFinanceResponse,
  YclientsServicesResponse,
  YclientsStatusResponse,
} from './src/yclientsTypes';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const isProduction = process.env.NODE_ENV === 'production';
const PORT = Number(process.env.PORT) || 3000;
const MAX_RANGE_DAYS = 366;

function validateRange(body: any): { startDate: string; endDate: string } {
  const startDate = String(body?.startDate || '');
  const endDate = String(body?.endDate || '');
  const re = /^\d{4}-\d{2}-\d{2}$/;
  if (!re.test(startDate) || !re.test(endDate)) {
    throw new YclientsError('Укажите период в формате ГГГГ-ММ-ДД (startDate, endDate).', 400);
  }
  const s = new Date(startDate + 'T00:00:00Z').getTime();
  const e = new Date(endDate + 'T00:00:00Z').getTime();
  if (isNaN(s) || isNaN(e) || s > e) {
    throw new YclientsError('Некорректный период: дата начала позже даты окончания.', 400);
  }
  if ((e - s) / 86400000 > MAX_RANGE_DAYS) {
    throw new YclientsError(`Период не должен превышать ${MAX_RANGE_DAYS} дней.`, 400);
  }
  return { startDate, endDate };
}

function sendError(res: express.Response, err: unknown) {
  if (err instanceof YclientsError) {
    res.status(err.status >= 400 && err.status < 600 ? err.status : 502).json({ error: err.message });
    return;
  }
  console.error('[yclients] unexpected error:', err);
  res.status(500).json({ error: 'Внутренняя ошибка сервера при обращении к Yclients.' });
}

async function startServer() {
  const app = express();
  app.use(express.json({ limit: '1mb' }));

  // --- Yclients proxy -------------------------------------------------------

  app.get('/api/yclients/status', (_req, res) => {
    const state = getConfigState();
    const body: YclientsStatusResponse = {
      configured: state.configured,
      companyId: state.companyId,
      missing: state.missing,
    };
    res.json(body);
  });

  // Финансовые операции -> строки для «Вкладки 1»
  app.post('/api/yclients/transactions', async (req, res) => {
    try {
      const { startDate, endDate } = validateRange(req.body);
      const cfg = await resolveConfig();

      const txs = await fetchTransactions(cfg, startDate, endDate);

      // Записи клиентов нужны, чтобы расшифровать доходы по услугам; без них отчёт всё равно строится
      let recordServices = new Map<number, { title: string; net: number }[]>();
      const warnings: string[] = [];
      try {
        const records = await fetchRecords(cfg, startDate, endDate);
        recordServices = buildRecordServicesMap(records);
      } catch (e: any) {
        warnings.push(
          `Не удалось получить записи клиентов для расшифровки доходов по услугам: ${e?.message || 'ошибка'}. Доходы загружены без разбивки по услугам.`
        );
      }

      const overrides = parseCategoryOverrides(process.env.YCLIENTS_CATEGORY_MAP);
      const { records, diagnostics } = mapTransactionsToFinanceRows(txs, recordServices, overrides);
      diagnostics.warnings.push(...warnings);

      const body: YclientsFinanceResponse = { records, diagnostics };
      res.json(body);
    } catch (err) {
      sendError(res, err);
    }
  });

  // Записи клиентов -> отчёт по услугам для «Вкладки 2»
  app.post('/api/yclients/services', async (req, res) => {
    try {
      const { startDate, endDate } = validateRange(req.body);
      const cfg = await resolveConfig();

      const warnings: string[] = [];
      let yRecords;
      try {
        yRecords = await fetchRecords(cfg, startDate, endDate, true);
      } catch (e) {
        // Если расширенная выборка не поддерживается — пробуем обычную (без разбивки способов оплаты)
        if (e instanceof YclientsError && (e.status === 401 || e.status === 403 || e.status === 429)) throw e;
        warnings.push('Разбивка способов оплаты недоступна — вся выручка отнесена к «Оплачено деньгами».');
        yRecords = await fetchRecords(cfg, startDate, endDate, false);
      }

      const { records, diagnostics } = mapRecordsToServiceRows(yRecords, {
        onlyAttended: req.body?.onlyAttended !== false,
      });
      diagnostics.warnings.push(...warnings);

      const body: YclientsServicesResponse = { records, diagnostics };
      res.json(body);
    } catch (err) {
      sendError(res, err);
    }
  });

  // --- Frontend -------------------------------------------------------------

  if (isProduction) {
    const distPath = path.join(__dirname, 'dist');
    app.use(express.static(distPath));
    app.get('*', (_req, res) => {
      res.sendFile(path.join(distPath, 'index.html'));
    });
  } else {
    const { createServer: createViteServer } = await import('vite');
    const vite = await createViteServer({
      server: { middlewareMode: true },
      appType: 'spa',
    });
    app.use(vite.middlewares);
  }

  app.listen(PORT, '0.0.0.0', () => {
    console.log(`Сервер запущен: http://localhost:${PORT} (${isProduction ? 'production' : 'development'})`);
    const state = getConfigState();
    if (!state.configured) {
      console.log(`Yclients: интеграция не настроена (не заданы: ${state.missing.join(', ')})`);
    }
  });
}

startServer();
