import React, { useEffect, useState } from 'react';
import { CloudDownload, Loader2, AlertCircle, CheckCircle } from 'lucide-react';
import { fetchYclientsStatus } from '../yclientsApi';
import type { YclientsStatusResponse } from '../yclientsTypes';

interface YclientsPanelProps {
  title: string;
  description: string;
  /** Загружает данные за период (даты в формате ГГГГ-ММ-ДД). Возвращает краткий итог для показа под кнопкой. */
  onLoad: (startIso: string, endIso: string) => Promise<string>;
  variant?: 'light' | 'dark';
}

const pad = (n: number) => String(n).padStart(2, '0');
const toIso = (d: Date) => `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;

export default function YclientsPanel({ title, description, onLoad, variant = 'light' }: YclientsPanelProps) {
  const now = new Date();
  const [startIso, setStartIso] = useState<string>(toIso(new Date(now.getFullYear(), now.getMonth(), 1)));
  const [endIso, setEndIso] = useState<string>(toIso(now));
  const [status, setStatus] = useState<YclientsStatusResponse | null>(null);
  const [statusError, setStatusError] = useState<string | null>(null);
  const [loading, setLoading] = useState(false);
  const [result, setResult] = useState<{ ok: boolean; text: string } | null>(null);

  useEffect(() => {
    let cancelled = false;
    fetchYclientsStatus()
      .then((s) => {
        if (!cancelled) setStatus(s);
      })
      .catch((e: any) => {
        if (!cancelled) setStatusError(e?.message || 'Не удалось проверить настройки Yclients');
      });
    return () => {
      cancelled = true;
    };
  }, []);

  const handleClick = async () => {
    setLoading(true);
    setResult(null);
    try {
      const summary = await onLoad(startIso, endIso);
      setResult({ ok: true, text: summary });
    } catch (e: any) {
      setResult({ ok: false, text: e?.message || 'Не удалось загрузить данные из Yclients' });
    } finally {
      setLoading(false);
    }
  };

  const dark = variant === 'dark';
  const box = dark
    ? 'bg-white/10 border border-white/20 text-white'
    : 'bg-white border border-gray-150 text-gray-900 shadow-sm';
  const muted = dark ? 'text-indigo-200' : 'text-gray-500';
  const input = dark
    ? 'bg-white/10 border border-white/20 text-white'
    : 'bg-gray-50 border border-gray-200 text-gray-800';

  const datesValid = !!startIso && !!endIso && startIso <= endIso;
  const canLoad = !!status?.configured && datesValid && !loading;

  return (
    <div className={`rounded-2xl p-4 space-y-3 ${box}`}>
      <div className="flex items-center gap-2">
        <CloudDownload className={`w-4 h-4 ${dark ? 'text-indigo-200' : 'text-indigo-600'}`} />
        <h3 className="font-display font-semibold text-sm">{title}</h3>
      </div>
      <p className={`text-xs leading-relaxed ${muted}`}>{description}</p>

      {statusError && (
        <div className="flex items-start gap-2 text-xs text-amber-600 bg-amber-50 border border-amber-200 rounded-xl p-2.5">
          <AlertCircle className="w-4 h-4 flex-shrink-0 mt-px" />
          <span>{statusError}</span>
        </div>
      )}

      {status && !status.configured && (
        <div className="flex items-start gap-2 text-xs text-amber-700 bg-amber-50 border border-amber-200 rounded-xl p-2.5">
          <AlertCircle className="w-4 h-4 flex-shrink-0 mt-px" />
          <span>
            Интеграция не настроена. Задайте в файле <code className="font-mono">.env</code>: {status.missing.join(', ')}. См. YCLIENTS.md.
          </span>
        </div>
      )}

      <div className="flex flex-wrap items-end gap-3">
        <label className="flex flex-col gap-1 text-[10px] uppercase tracking-wider font-medium">
          <span className={muted}>С</span>
          <input
            type="date"
            value={startIso}
            onChange={(e) => setStartIso(e.target.value)}
            className={`rounded-lg px-2.5 py-1.5 text-xs font-mono ${input}`}
          />
        </label>
        <label className="flex flex-col gap-1 text-[10px] uppercase tracking-wider font-medium">
          <span className={muted}>По</span>
          <input
            type="date"
            value={endIso}
            onChange={(e) => setEndIso(e.target.value)}
            className={`rounded-lg px-2.5 py-1.5 text-xs font-mono ${input}`}
          />
        </label>
        <button
          type="button"
          onClick={handleClick}
          disabled={!canLoad}
          className="px-4 py-2 bg-indigo-600 text-white rounded-xl text-xs font-semibold shadow-md shadow-indigo-200/40 hover:bg-indigo-700 transition-colors cursor-pointer disabled:opacity-50 disabled:cursor-not-allowed flex items-center gap-2"
        >
          {loading ? <Loader2 className="w-3.5 h-3.5 animate-spin" /> : <CloudDownload className="w-3.5 h-3.5" />}
          <span>{loading ? 'Загрузка…' : 'Загрузить из Yclients'}</span>
        </button>
      </div>

      {!datesValid && (
        <p className="text-[11px] text-amber-600">Дата начала должна быть не позже даты окончания.</p>
      )}

      {result && (
        <div
          className={`flex items-start gap-2 text-xs rounded-xl p-2.5 border ${
            result.ok
              ? 'text-emerald-700 bg-emerald-50 border-emerald-200'
              : 'text-red-700 bg-red-50 border-red-200'
          }`}
        >
          {result.ok ? (
            <CheckCircle className="w-4 h-4 flex-shrink-0 mt-px" />
          ) : (
            <AlertCircle className="w-4 h-4 flex-shrink-0 mt-px" />
          )}
          <span className="whitespace-pre-line">{result.text}</span>
        </div>
      )}
    </div>
  );
}
