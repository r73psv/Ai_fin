import React, { useState, useRef, useMemo } from 'react';
import * as XLSX from 'xlsx';
import { 
  Upload, 
  Download, 
  Search, 
  Filter, 
  Calendar, 
  TrendingUp, 
  DollarSign, 
  AlertCircle, 
  CheckCircle, 
  RefreshCw, 
  BookOpen, 
  Users, 
  Table, 
  FileText,
  Lock,
  Edit2,
  Trash2,
  HelpCircle,
  Undo,
  Play,
  Loader2,
  Check,
  ChevronDown,
  ChevronUp,
  ChevronsUpDown,
  ChevronsDownUp,
  Layers,
  Package,
  Eye,
  EyeOff,
  Scissors,
  Calculator,
  PieChart,
  Percent,
  BarChart3,
  SlidersHorizontal,
  Sparkles,
  Building2,
  ChevronRight,
  Settings2,
  X
} from 'lucide-react';
import { motion, AnimatePresence } from 'motion/react';

import { 
  RawRecord, 
  MetricRow, 
  SalaryEmployee, 
  SalaryProcessingMetrics, 
  RevenueItemDetail,
  RawServiceRecord,
  ServiceReportItem,
  ServicesReportSummary,
  ServicesPreprocessingStats,
  HallReportGroup,
  ExpenseAllocationMethod
} from './types';
import { 
  parseUploadedSpreadsheet, 
  generateFinancialReport, 
  processSalaryPayments,
  calculateReportFormulas,
  downloadFinancialSpreadsheet,
  downloadSalarySpreadsheet,
  downloadRevenueSpreadsheet,
  groupRevenueByServiceProduct,
  parseDate,
  parseServicesSpreadsheet,
  generateServicesReport,
  downloadServicesSpreadsheet
} from './processingEngine';
import { HALL_NAMES, HallName, HALL_METADATA } from './hallsClassifier';
import { FINANCIAL_TEMPLATE } from './financialTemplate';
import { SUPPLIER_HALL_RULES, SuppliersAllocationBreakdown } from './suppliersClassifier';
import TestPlayground from './components/TestPlayground';
import UserGuideModal from './components/UserGuideModal';
import YclientsPanel from './components/YclientsPanel';
import { fetchYclientsFinance, fetchYclientsServices, isoToRu } from './yclientsApi';
import { downloadDocxReport } from './docxGenerator';

export interface AggregatedArticleConfig {
  id: string;
  title: string;
  shortTitle: string;
  icon: string;
  headerLabel: string;
  isChildRow: (row: MetricRow, all: MetricRow[]) => boolean;
}

export const AGGREGATED_SECTIONS: AggregatedArticleConfig[] = [
  {
    id: 'revenue',
    title: 'Выручка общая (Доходы)',
    shortTitle: 'Доходы (Выручка)',
    icon: '📈',
    headerLabel: 'Выручка общая',
    isChildRow: (row, all) => {
      const pIdx = all.findIndex(r => r.label === 'Выручка общая');
      const nIdx = all.findIndex(r => r.label === 'Расходы');
      const rIdx = all.indexOf(row);
      return pIdx !== -1 && nIdx !== -1 && rIdx > pIdx && rIdx < nIdx;
    }
  },
  {
    id: 'variables',
    title: 'Расходы переменные',
    shortTitle: 'Переменные',
    icon: '⚡',
    headerLabel: 'Расходы переменные',
    isChildRow: (row, all) => {
      const pIdx = all.findIndex(r => r.label === 'Расходы переменные');
      const nIdx = all.findIndex(r => r.label === 'Расходы постоянные');
      const rIdx = all.indexOf(row);
      return pIdx !== -1 && nIdx !== -1 && rIdx > pIdx && rIdx < nIdx;
    }
  },
  {
    id: 'fixed',
    title: 'Расходы постоянные',
    shortTitle: 'Постоянные',
    icon: '🏢',
    headerLabel: 'Расходы постоянные',
    isChildRow: (row, all) => {
      const pIdx = all.findIndex(r => r.label === 'Расходы постоянные');
      const nIdx = all.findIndex(r => r.label === 'Поставщики');
      const rIdx = all.indexOf(row);
      return pIdx !== -1 && nIdx !== -1 && rIdx > pIdx && rIdx < nIdx;
    }
  },
  {
    id: 'suppliers',
    title: 'Поставщики (материалы)',
    shortTitle: 'Поставщики',
    icon: '📦',
    headerLabel: 'Поставщики',
    isChildRow: (row, all) => {
      const pIdx = all.findIndex(r => r.label === 'Поставщики');
      const nIdx = all.findIndex(r => r.label === 'Расходы на заработную плату');
      const rIdx = all.indexOf(row);
      return pIdx !== -1 && nIdx !== -1 && rIdx > pIdx && rIdx < nIdx;
    }
  },
  {
    id: 'salary',
    title: 'Расходы на заработную плату',
    shortTitle: 'Зарплата (ФОТ)',
    icon: '👥',
    headerLabel: 'Расходы на заработную плату',
    isChildRow: (row, all) => {
      const pIdx = all.findIndex(r => r.label === 'Расходы на заработную плату');
      const nIdx = all.findIndex(r => r.label === 'Прочие' || r.label === 'НАЛОГИ');
      const rIdx = all.indexOf(row);
      return pIdx !== -1 && nIdx !== -1 && rIdx > pIdx && rIdx < nIdx;
    }
  },
  {
    id: 'additional',
    title: 'Дополнительные расходы',
    shortTitle: 'Доп. расходы',
    icon: '➕',
    headerLabel: 'Дополнительные расходы',
    isChildRow: (row, all) => {
      const pIdx = all.findIndex(r => r.label === 'Дополнительные расходы');
      const nIdx = all.findIndex(r => r.label === 'Отчисления в фонды');
      const rIdx = all.indexOf(row);
      return pIdx !== -1 && nIdx !== -1 && rIdx > pIdx && rIdx < nIdx;
    }
  },
  {
    id: 'funds_expenses',
    title: 'Расходы за счёт фондов',
    shortTitle: 'Расходы фондов',
    icon: '🛡️',
    headerLabel: 'Расходы за счёт фондов',
    isChildRow: (row, all) => {
      const pIdx = all.findIndex(r => r.label === 'Расходы за счёт фондов');
      const nIdx = all.findIndex(r => r.label === 'ЧИСТАЯ ПРИБЫЛЬ');
      const rIdx = all.indexOf(row);
      return pIdx !== -1 && nIdx !== -1 && rIdx > pIdx && rIdx < nIdx;
    }
  },
  {
    id: 'funds',
    title: 'ФОНДЫ',
    shortTitle: 'Фонды',
    icon: '🏦',
    headerLabel: 'ФОНДЫ',
    isChildRow: (row, all) => {
      const pIdx = all.findIndex(r => r.label === 'ФОНДЫ');
      const nIdx = all.findIndex(r => r.label === 'Распределение прибыли');
      const endIdx = nIdx !== -1 ? nIdx : all.length;
      const rIdx = all.indexOf(row);
      return pIdx !== -1 && rIdx > pIdx && rIdx < endIdx;
    }
  },
  {
    id: 'profit_distribution',
    title: 'Распределение прибыли',
    shortTitle: 'Распределение прибыли',
    icon: '🤝',
    headerLabel: 'Распределение прибыли',
    isChildRow: (row, all) => {
      const pIdx = all.findIndex(r => r.label === 'Распределение прибыли');
      const rIdx = all.indexOf(row);
      return pIdx !== -1 && rIdx > pIdx;
    }
  }
];

// Realistic mock data so users can test the dashboard instantly without preparing files.
const DEMO_RECORDS: Omit<RawRecord, 'id' | 'date'>[] = [
  // Revenue with "Услуга/Товар" detailing (matching rows grouped & summed)
  { sum: 140000, category: 'Выручка с услуг', counterparty: 'Клиенты салона', serviceProduct: 'Стрижка женская (модельная)', comment: 'Стрижки май 1 смена', dateStr: '10.05.2026', rawIndex: 2 },
  { sum: 110000, category: 'Выручка с услуг', counterparty: 'Клиенты салона', serviceProduct: 'Стрижка женская (модельная)', comment: 'Стрижки май 2 смена', dateStr: '14.05.2026', rawIndex: 3 },
  { sum: 90000, category: 'Выручка с услуг', counterparty: 'Клиенты салона', serviceProduct: 'Сложное окрашивание (Airtouch)', comment: 'Окрашивание волос сложные техники', dateStr: '15.05.2026', rawIndex: 4 },
  { sum: 65000, category: 'Выручка с товаров', counterparty: 'Розничные покупатели', serviceProduct: 'Шампунь Kerastase Nutritive 250ml', comment: 'Продажа косметики витрина', dateStr: '18.05.2026', rawIndex: 5 },
  { sum: 55000, category: 'Выручка с товаров', counterparty: 'Розничные покупатели', serviceProduct: 'Шампунь Kerastase Nutritive 250ml', comment: 'Продажа косметики витрина', dateStr: '19.05.2026', rawIndex: 6 },
  { sum: 30000, category: 'Пополнение счёта', counterparty: 'Абоненты', serviceProduct: 'Пополнение депозитного баланса', comment: 'Депозиты на баланс', dateStr: '20.05.2026', rawIndex: 7 },
  { sum: 45000, category: 'Выручка с сертификатов', counterparty: 'Подарочные продажи', serviceProduct: 'Подарочный сертификат 5000 ₽', comment: 'Сертификаты подарочные', dateStr: '22.05.2026', rawIndex: 8 },
  // Operating Variable Expenses
  { sum: 15400, category: '2026 Жизнеобеспечение салона', counterparty: 'Пекарня Сласть', comment: 'Вкусняшки для утренней смены', dateStr: '02.05.2026', rawIndex: 6 },
  { sum: 3400, category: 'Доставка', counterparty: 'Яндекс Доставка', comment: 'Доставка одноразовых простыней', dateStr: '05.05.2026', rawIndex: 7 },
  { sum: 12000, category: 'Кофе', counterparty: 'Спешелти Ростерс', comment: 'Кофе свежей обжарки', dateStr: '15.05.2026', rawIndex: 8 },
  // Suppliers Grouping Rules (распределяются между залами согласно таблице)
  { sum: 45000, category: 'Закупка материалов', counterparty: 'ИП Щеголькова (ProfiLine)', comment: 'Краски для волос, оксиды', dateStr: '10.05.2026', rawIndex: 9 }, // -> Зал макияжей и укладок
  { sum: 28000, category: 'Закупка товаров', counterparty: 'ООО Солком Регион (Лореаль витрина)', comment: 'Шампуни Лореаль на витрину', dateStr: '12.05.2026', rawIndex: 10 }, // -> Зал колористики/парикмахеры
  { sum: 19500, category: 'Закупка материалов', counterparty: 'Калашникова (Matrix расходники)', comment: 'Расходники Матрикс красители', dateStr: '13.05.2026', rawIndex: 11 }, // -> Зал колористики/парикмахеры
  { sum: 34000, category: 'Закупка материалов', counterparty: 'Материалы для ногтевиков', comment: 'Гель-лаки, базы, фрезы, жидкости', dateStr: '14.05.2026', rawIndex: 23 }, // -> Зал ногтевого сервиса
  { sum: 12500, category: 'Закупка материалов', counterparty: 'Материалы для бровистов', comment: 'Краски Thuya, воски, пинцеты', dateStr: '15.05.2026', rawIndex: 24 }, // -> Зал бровей
  { sum: 16000, category: 'Закупка материалов', counterparty: 'ИП Щеголькова (American CREW)', comment: 'Глины, помады, шампуни для барберов', dateStr: '16.05.2026', rawIndex: 25 }, // -> Мужской зал
  { sum: 18000, category: 'Закупка материалов', counterparty: 'OOO «Торговый Дом „БлескСервис“» (TORK) (Расходники)', comment: 'Бумажная продукция для всех залов', dateStr: '17.05.2026', rawIndex: 26 }, // -> распределяются между всеми
  // Overheads / Constants
  { sum: 120000, category: 'Аренда', counterparty: 'Бизнес-Центр Меридиан', comment: 'Арендная плата май', dateStr: '01.05.2026', rawIndex: 12 },
  { sum: 25000, category: 'Бухгалтер', counterparty: 'ИП Петрова Капитал', comment: 'Сдача отчетности и учет', dateStr: '05.05.2026', rawIndex: 13 },
  { sum: 18000, category: 'Коммунальные платежи', counterparty: 'Энергосбыт', comment: 'Оплата электричества', dateStr: '10.05.2026', rawIndex: 14 },
  { sum: 35000, category: 'Реклама, маркетинг, таргет', counterparty: 'ТаргетПро Агентство', comment: 'Кампания ВК и Яндекс Карты', dateStr: '11.05.2026', rawIndex: 15 },
  // Salary - Employee Payments (Various rules to showcase the split logic)
  { sum: 50000, category: 'Зарплата', counterparty: 'Ковалева Лариса (Мастер)', comment: 'Зарплата за апрель', dateStr: '10.05.2026', rawIndex: 16 },
  { sum: 35000, category: 'Зарплата masters', counterparty: 'Ковалева Лариса (Мастер)', comment: 'Аванс за май', dateStr: '25.05.2026', rawIndex: 17 }, // No comment keyword -> reporting
  { sum: 45000, category: 'Заработная плата', counterparty: 'Цветкова Анна (Администратор)', comment: 'з.п. остаток за апрель', dateStr: '15.05.2026', rawIndex: 18 },
  { sum: 20000, category: 'Зарплата', counterparty: 'Цветкова Анна (Администратор)', comment: 'Бонус за выполнение плана продаж', dateStr: '28.05.2026', rawIndex: 19 }, // No keyword -> reporting
  { sum: 55000, category: 'Зарплата', counterparty: 'Власов Сергей (Топ-Барбер)', comment: 'Зарплата за апрель по ведомости', dateStr: '14.05.2026', rawIndex: 20 },
  { sum: 40000, category: 'Зарплата', counterparty: 'Власов Сергей (Топ-Барбер)', comment: 'ЗП авансовая выплата май', dateStr: '26.05.2026', rawIndex: 21 },
  // Founders payouts («2026/Учредители» -> «Выплаты учредителям»)
  { sum: 150000, category: '2026/Учредители', counterparty: 'Учредитель (Ренат / Надежда)', comment: 'Выплата учредителям за май', dateStr: '28.05.2026', rawIndex: 27 },
  // Unmatched additional rows (tests appending extra lines automatically)
  { sum: 14200, category: 'Непредвиденный форс-мажор ураган', counterparty: 'ООО РемСтройВолга', comment: 'Срочный ремонт козырька входа', dateStr: '16.05.2026', rawIndex: 22 }
];

// Realistic mock services data for instant testing of Tab 2 (covering all salon halls)
const DEMO_SERVICES_RECORDS: Omit<RawServiceRecord, 'id' | 'revenue'>[] = [
  // Зал колористики/парикмахеры
  { rawIndex: 2, service: 'Стрижка женская (модельная)', discount: 500, points: 200, subscriptions: 0, certificates: 0, clientAccounts: 0, money: 2800 },
  { rawIndex: 3, service: 'Стрижка женская (модельная)', discount: 0, points: 0, subscriptions: 0, certificates: 1500, clientAccounts: 0, money: 2000 },
  { rawIndex: 4, service: 'Стрижка женская (модельная)', discount: 300, points: 100, subscriptions: 0, certificates: 0, clientAccounts: 1200, money: 1900 },
  { rawIndex: 5, service: 'Сложное окрашивание (Airtouch)', discount: 1500, points: 500, subscriptions: 0, certificates: 3000, clientAccounts: 2000, money: 8500 },
  { rawIndex: 6, service: 'Сложное окрашивание (Airtouch)', discount: 0, points: 0, subscriptions: 0, certificates: 0, clientAccounts: 0, money: 15000 },
  { rawIndex: 7, service: 'Уход для волос глубокое восстановление Kerastase', discount: 500, points: 300, subscriptions: 0, certificates: 2000, clientAccounts: 0, money: 2700 },
  { rawIndex: 8, service: 'Окрашивание корней Matrix', discount: 0, points: 250, subscriptions: 0, certificates: 1500, clientAccounts: 0, money: 3250 },
  
  // Зал ногтевого сервиса
  { rawIndex: 9, service: 'Маникюр аппаратный с покрытием гель-лак', discount: 200, points: 150, subscriptions: 2000, certificates: 0, clientAccounts: 0, money: 350 },
  { rawIndex: 10, service: 'Маникюр аппаратный с покрытием гель-лак', discount: 0, points: 0, subscriptions: 0, certificates: 1000, clientAccounts: 500, money: 1000 },
  { rawIndex: 11, service: 'Маникюр аппаратный с покрытием гель-лак', discount: 400, points: 200, subscriptions: 0, certificates: 0, clientAccounts: 0, money: 1900 },
  { rawIndex: 12, service: 'Педикюр Smart классический', discount: 0, points: 100, subscriptions: 0, certificates: 0, clientAccounts: 1000, money: 2100 },
  { rawIndex: 13, service: 'Педикюр Smart классический', discount: 300, points: 0, subscriptions: 2500, certificates: 0, clientAccounts: 0, money: 400 },

  // Зал макияжей и укладок
  { rawIndex: 14, service: 'Макияж вечерний / никах', discount: 300, points: 200, subscriptions: 0, certificates: 1000, clientAccounts: 0, money: 3500 },
  { rawIndex: 15, service: 'Укладка на dyson + спа g-prof', discount: 0, points: 150, subscriptions: 0, certificates: 0, clientAccounts: 500, money: 2850 },
  { rawIndex: 16, service: 'Прическа невесты', discount: 500, points: 0, subscriptions: 0, certificates: 2000, clientAccounts: 0, money: 4500 },

  // Зал бровей
  { rawIndex: 17, service: 'Архитектура и коррекция бровей', discount: 100, points: 50, subscriptions: 0, certificates: 0, clientAccounts: 0, money: 1050 },
  { rawIndex: 18, service: 'Ламинирование ресниц', discount: 200, points: 100, subscriptions: 0, certificates: 500, clientAccounts: 0, money: 1900 },

  // Мужской зал
  { rawIndex: 19, service: 'Стрижка мужская Fade / Кроп', discount: 0, points: 100, subscriptions: 0, certificates: 0, clientAccounts: 0, money: 1900 },
  { rawIndex: 20, service: 'Стрижка мужская Fade / Кроп', discount: 200, points: 0, subscriptions: 0, certificates: 0, clientAccounts: 500, money: 1300 },
  { rawIndex: 21, service: 'Оформление бороды', discount: 0, points: 50, subscriptions: 0, certificates: 0, clientAccounts: 0, money: 1150 },

  // Прочие услуги (для демонстрации отнесения услуг к залам салона с запоминанием выбора)
  { rawIndex: 22, service: 'Аренда рабочего места (коворкинг мастера)', discount: 0, points: 0, subscriptions: 0, certificates: 0, clientAccounts: 0, money: 3500 },
  { rawIndex: 23, service: 'Консультация трихолога с микровидеодиагностикой', discount: 200, points: 100, subscriptions: 0, certificates: 0, clientAccounts: 500, money: 2500 },
  { rawIndex: 24, service: 'Сервисный сбор и сопровождение VIP-гостя', discount: 0, points: 0, subscriptions: 0, certificates: 1000, clientAccounts: 0, money: 1500 }
];

const STORAGE_KEY_HALL_OVERRIDES = 'salon_service_hall_overrides';

function loadSavedHallOverrides(): Record<string, HallName> {
  try {
    const raw = localStorage.getItem(STORAGE_KEY_HALL_OVERRIDES);
    if (raw) {
      const parsed = JSON.parse(raw);
      if (parsed && typeof parsed === 'object') {
        return parsed as Record<string, HallName>;
      }
    }
  } catch (e) {
    console.error('Failed to load saved hall overrides', e);
  }
  return {};
}

function saveHallOverrides(overrides: Record<string, HallName>) {
  try {
    localStorage.setItem(STORAGE_KEY_HALL_OVERRIDES, JSON.stringify(overrides));
  } catch (e) {
    console.error('Failed to save hall overrides', e);
  }
}

export default function App() {
  // Primary top-level navigation: Tab 1 (Финансовый отчет) vs Tab 2 (Отчет по услугам)
  const [primaryTab, setPrimaryTab] = useState<'finance' | 'services'>('finance');

  // Application State - Tab 1 (Finance)
  const [cutoffDateStr, setCutoffDateStr] = useState('20.05.2026'); // June 2026 is current. June starts on 1st. Let's make cutoff May 20th, 2026
  const [fileName, setFileName] = useState<string | null>(null);
  const [rawRecords, setRawRecords] = useState<RawRecord[]>([]);
  const [userOverrides, setUserOverrides] = useState<Record<string | number, number>>({});
  const [manualMastersSalary, setManualMastersSalary] = useState<string>('');
  const [manualAdminsSalary, setManualAdminsSalary] = useState<string>('');
  const [categoryMappings, setCategoryMappings] = useState<Record<string, string>>({});
  const [transactionMappings, setTransactionMappings] = useState<Record<string, string>>({});
  const [dragActive, setDragActive] = useState(false);
  const [searchQuery, setSearchQuery] = useState('');
  const [tableFilter, setTableFilter] = useState<'all' | 'data' | 'formulas' | 'nonzero'>('all');
  const [activeTab, setActiveTab] = useState<'financial' | 'salary' | 'revenue-details' | 'raw' | 'faq'>('financial');
  const [showRevenueDetailInline, setShowRevenueDetailInline] = useState(true);
  const [revenueSearch, setRevenueSearch] = useState('');
  const [revenueCategoryFilter, setRevenueCategoryFilter] = useState<'all' | 'services' | 'goods' | 'other'>('all');
  const [expandedRevenueItems, setExpandedRevenueItems] = useState<Record<string, boolean>>({});
  const [editingRowIdx, setEditingRowIdx] = useState<number | null>(null);
  const [editingRowValue, setEditingRowValue] = useState<string>('');
  const [expandedExtraRows, setExpandedExtraRows] = useState<Record<number, boolean>>({});
  const [expandedReportRows, setExpandedReportRows] = useState<Record<number, boolean>>({});
  // Collapsed Aggregated Sections in Financial Report (Tab 1)
  const [collapsedSections, setCollapsedSections] = useState<Record<string, boolean>>({});
  const [notification, setNotification] = useState<{ status: 'success' | 'error', message: string } | null>(null);

  // Application State - Tab 2 (Services Report)
  const [servicesRecords, setServicesRecords] = useState<RawServiceRecord[]>([]);
  const [servicesFileName, setServicesFileName] = useState<string | null>(null);
  const [servicesSearchQuery, setServicesSearchQuery] = useState('');
  const [servicesDragActive, setServicesDragActive] = useState(false);
  const [expandedServiceItems, setExpandedServiceItems] = useState<Record<string, boolean>>({});
  const [customVariableExpenses, setCustomVariableExpenses] = useState<string>('');
  const [customFixedExpenses, setCustomFixedExpenses] = useState<string>('');
  const [customSupplierExpenses, setCustomSupplierExpenses] = useState<string>('');
  const [isServicesCustomExpensesActive, setIsServicesCustomExpensesActive] = useState(false);
  const [showSuppliersRulesModal, setShowSuppliersRulesModal] = useState<boolean>(false);
  const [servicesStats, setServicesStats] = useState<ServicesPreprocessingStats | null>(null);
  const [selectedHallFilter, setSelectedHallFilter] = useState<string>('all');
  const [collapsedHalls, setCollapsedHalls] = useState<Record<string, boolean>>({});
  
  // Hall Areas & Salaries State (Tab 2) - initial values cleared of demo data
  const [hallAreas, setHallAreas] = useState<Record<string, number>>({
    'Зал колористики/парикмахеры': 0,
    'Зал макияжей и укладок': 0,
    'Зал ногтевого сервиса': 0,
    'Зал бровей': 0,
    'Мужской зал': 0,
    'Зал ресниц': 0,
    'Фотостудия': 0,
    'Зал эпиляции': 0,
    'ЗАЛ SPA': 0,
    'Зал массажный': 0,
    'Прочие услуги': 0
  });
  const [hallMastersSalaries, setHallMastersSalaries] = useState<Record<string, number>>({
    'Зал колористики/парикмахеры': 0,
    'Зал макияжей и укладок': 0,
    'Зал ногтевого сервиса': 0,
    'Зал бровей': 0,
    'Мужской зал': 0,
    'Зал ресниц': 0,
    'Фотостудия': 0,
    'Зал эпиляции': 0,
    'ЗАЛ SPA': 0,
    'Зал массажный': 0,
    'Прочие услуги': 0
  });
  const [serviceAdminSalary, setServiceAdminSalary] = useState<number>(0);
  // Cost allocation method: 'revenue' (пропорционально выручке) or 'area' (пропорционально площади залов)
  const [expenseAllocationMethod, setExpenseAllocationMethod] = useState<ExpenseAllocationMethod>('revenue');
  const [collapsedRevenueDetails, setCollapsedRevenueDetails] = useState<Record<number, boolean>>({});
  const [customAdminSalary, setCustomAdminSalary] = useState<string>('');
  const [serviceHallOverrides, setServiceHallOverrides] = useState<Record<string, HallName>>(() => loadSavedHallOverrides());
  const [showHallOverridesModal, setShowHallOverridesModal] = useState<boolean>(false);
  const [showUserGuideModal, setShowUserGuideModal] = useState<boolean>(false);
  const servicesFileInputRef = useRef<HTMLInputElement>(null);

  // Handlers for assigning services from "Прочие услуги" to a hall with persistence
  const handleAssignServiceToHall = (serviceName: string, targetHall: HallName) => {
    setServiceHallOverrides(prev => {
      let next: Record<string, HallName>;
      if (targetHall === 'Прочие услуги') {
        next = { ...prev };
        delete next[serviceName];
      } else {
        next = { ...prev, [serviceName]: targetHall };
      }
      saveHallOverrides(next);
      return next;
    });
    setNotification({
      status: 'success',
      message: targetHall === 'Прочие услуги'
        ? `Услуга «${serviceName}» возвращена в категорию «Прочие услуги».`
        : `Услуга «${serviceName}» успешно отнесена к залу «${targetHall}». Выбор сохранен и применен во всех отчетах!`
    });
  };

  const handleResetServiceHall = (serviceName: string) => {
    setServiceHallOverrides(prev => {
      const next = { ...prev };
      delete next[serviceName];
      saveHallOverrides(next);
      return next;
    });
    setNotification({
      status: 'success',
      message: `Сброшено ручное назначение зала для «${serviceName}».`
    });
  };

  const handleClearAllHallOverrides = () => {
    setServiceHallOverrides({});
    saveHallOverrides({});
    setNotification({
      status: 'success',
      message: 'Все сохраненные ручные назначения услуг по залам успешно сброшены.'
    });
  };

  // Staging and processing states
  const [stagedRecords, setStagedRecords] = useState<RawRecord[]>([]);
  const [stagedFileName, setStagedFileName] = useState<string | null>(null);
  const [isProcessed, setIsProcessed] = useState(false);
  const [isProcessing, setIsProcessing] = useState(false);
  const [processingStep, setProcessingStep] = useState(0);

  const fileInputRef = useRef<HTMLInputElement>(null);

  // Parse cutoff date. Re-computes whenever cutoff string changes.
  const parsedCutoffDate = useMemo(() => {
    return parseDate(cutoffDateStr);
  }, [cutoffDateStr]);

  // Validates if input matches standard DD.MM.YYYY
  const isCutoffDateValid = useMemo(() => {
    if (!cutoffDateStr) return false;
    const regex = /^\d{2}\.\d{2}\.\d{4}$/;
    if (!regex.test(cutoffDateStr)) return false;
    const d = parsedCutoffDate;
    return d !== null && !isNaN(d.getTime());
  }, [cutoffDateStr, parsedCutoffDate]);

  // Handle uploaded data
  const handleDataParsed = (recordsList: RawRecord[], name: string) => {
    setStagedRecords(recordsList);
    setStagedFileName(name);
    setIsProcessed(false);
    setUserOverrides({});
    setCategoryMappings({});
    setTransactionMappings({});
    setNotification({
      status: 'success',
      message: `Файл "${name}" (${recordsList.length} транзакций) успешно загружен и подготовлен к обработке. Нажмите кнопку «Запустить программу» ниже.`
    });
  };

  // Run the processing engine & trigger automatic download of both reports
  const handleRunProgram = async () => {
    if (stagedRecords.length === 0) {
      setNotification({
        status: 'error',
        message: 'Пожалуйста, сначала выберите или загрузите файл данных.'
      });
      return;
    }

    setIsProcessing(true);
    setProcessingStep(1);

    // Step 1: Parsing entries
    await new Promise((resolve) => setTimeout(resolve, 500));
    setProcessingStep(2);

    // Step 2: Extracting contractor groupings & materials
    await new Promise((resolve) => setTimeout(resolve, 500));
    setProcessingStep(3);

    // Step 3: Performing math & formulas validation
    await new Promise((resolve) => setTimeout(resolve, 500));
    setProcessingStep(4);

    // Final state sync
    const finalRecords = [...stagedRecords];
    const finalName = stagedFileName || 'Обработанные_Данные.xlsx';

    // Parse manual salaries if entered or take values from Services tab
    const newOverrides = { ...userOverrides };
    const currentTotalMasters: number = Object.values(hallMastersSalaries).reduce<number>((acc, v) => acc + (typeof v === 'number' ? v : Number(v) || 0), 0);
    newOverrides["Зарплата мастеров"] = -Math.abs(currentTotalMasters);
    newOverrides["Зарплата административного персонала"] = -Math.abs(serviceAdminSalary);
    if (manualMastersSalary.trim() !== '') {
      const parsedMasters = parseFloat(manualMastersSalary.replace(/\s/g, '').replace(',', '.'));
      if (!isNaN(parsedMasters)) {
        newOverrides["Зарплата мастеров"] = -Math.abs(parsedMasters);
      }
    }
    if (manualAdminsSalary.trim() !== '') {
      const parsedAdmins = parseFloat(manualAdminsSalary.replace(/\s/g, '').replace(',', '.'));
      if (!isNaN(parsedAdmins)) {
        newOverrides["Зарплата административного персонала"] = -Math.abs(parsedAdmins);
      }
    }
    setUserOverrides(newOverrides);

    // Set records so that useMemo automatically generates report & salary rows
    setRawRecords(finalRecords);
    setFileName(finalName);
    setIsProcessed(true);
    setIsProcessing(false);
    setProcessingStep(0);

    // Automatic download disabled per user instruction. Reports are ready on-screen.
    setNotification({
      status: 'success',
      message: 'Программа успешно запущена! Данные рассчитаны, отчеты сформированы на экране. При необходимости выгрузите Excel-файлы вручную.'
    });
  };

  const handleFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const selectedFile = e.target.files?.[0];
    if (!selectedFile) return;

    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const rawData = evt.target?.result;
        if (!rawData) throw new Error('Не удалось прочитать файл');
        
        const dataArray = new Uint8Array(rawData as ArrayBuffer);
        const workbook = XLSX.read(dataArray, { type: 'array', cellDates: true });
        
        const { records, errors } = parseUploadedSpreadsheet(workbook);
        
        if (errors.length > 0) {
          setNotification({ status: 'error', message: errors.join(' ') });
          return;
        }

        handleDataParsed(records, selectedFile.name);
      } catch (err: any) {
        setNotification({
          status: 'error',
          message: `Ошибка чтения документа: ${err.message || 'неподдерживаемый формат'}`
        });
      }
    };
    reader.readAsArrayBuffer(selectedFile);
  };

  const handleDrag = (e: React.DragEvent) => {
    e.preventDefault();
    e.stopPropagation();
    if (e.type === "dragenter" || e.type === "dragover") {
      setDragActive(true);
    } else if (e.type === "dragleave") {
      setDragActive(false);
    }
  };

  const handleDrop = (e: React.DragEvent) => {
    e.preventDefault();
    e.stopPropagation();
    setDragActive(false);

    if (e.dataTransfer.files && e.dataTransfer.files[0]) {
      const selectedFile = e.dataTransfer.files[0];
      const ext = selectedFile.name.split('.').pop()?.toLowerCase();
      if (ext !== 'xlsx' && ext !== 'xls' && ext !== 'csv') {
        setNotification({
          status: 'error',
          message: 'Недопустимый формат файла. Пожалуйста, загрузите файл .xlsx, .xls или .csv'
        });
        return;
      }

      const reader = new FileReader();
      reader.onload = (evt) => {
        try {
          const rawData = evt.target?.result;
          if (!rawData) throw new Error('Не удалось прочитать файл');
          const dataArray = new Uint8Array(rawData as ArrayBuffer);
          const workbook = XLSX.read(dataArray, { type: 'array', cellDates: true });
          const { records, errors } = parseUploadedSpreadsheet(workbook);

          if (errors.length > 0) {
            setNotification({ status: 'error', message: errors.join(' ') });
            return;
          }

          handleDataParsed(records, selectedFile.name);
        } catch (err: any) {
          setNotification({
            status: 'error',
            message: `Ошибка парсинга: ${err.message}`
          });
        }
      };
      reader.readAsArrayBuffer(selectedFile);
    }
  };

  // Load realistic demo data
  const loadDemoData = () => {
    const structured = DEMO_RECORDS.map((dem, idx) => ({
      ...dem,
      id: `demo_rec_${idx}_${Date.now()}`,
      date: parseDate(dem.dateStr),
    }));
    handleDataParsed(structured, 'Финансовые_операции_(25.01.2026_-_31.01.2026).xlsx');
  };

  // Master salaries sum from Tab 2 (Services Report)
  const totalMastersSalaryFromServices = useMemo(() => {
    return Object.values(hallMastersSalaries).reduce<number>((acc, v) => acc + (typeof v === 'number' ? v : Number(v) || 0), 0);
  }, [hallMastersSalaries]);

  // Effective overrides synchronizing salaries from Services tab into Financial Report
  const effectiveUserOverrides = useMemo(() => {
    const overrides: Record<string, number> = { ...userOverrides };
    // Articles "Зарплата мастеров" and "Зарплата административного персонала" are populated from Services tab
    overrides["Зарплата мастеров"] = -Math.abs(totalMastersSalaryFromServices);
    overrides["Зарплата административного персонала"] = -Math.abs(serviceAdminSalary);
    return overrides;
  }, [userOverrides, totalMastersSalaryFromServices, serviceAdminSalary]);

  // Processing results: Financial Report
  const financialResults = useMemo(() => {
    return generateFinancialReport(rawRecords, effectiveUserOverrides, categoryMappings, transactionMappings, fileName, serviceHallOverrides);
  }, [rawRecords, effectiveUserOverrides, categoryMappings, transactionMappings, fileName, serviceHallOverrides]);

  const { report: financialReport, extraRows, revenueDetails } = financialResults;

  // Filtered revenue details for the dedicated revenue tab
  const filteredRevenueDetails = useMemo(() => {
    if (!revenueDetails) return [];
    return revenueDetails.filter((item) => {
      const matchesSearch = !revenueSearch || 
        item.name.toLowerCase().includes(revenueSearch.toLowerCase()) ||
        item.category.toLowerCase().includes(revenueSearch.toLowerCase());
      
      let matchesCat = true;
      if (revenueCategoryFilter === 'services') {
        matchesCat = item.category.toLowerCase().includes('услуг');
      } else if (revenueCategoryFilter === 'goods') {
        matchesCat = item.category.toLowerCase().includes('товар');
      } else if (revenueCategoryFilter === 'other') {
        matchesCat = !item.category.toLowerCase().includes('услуг') && !item.category.toLowerCase().includes('товар');
      }

      return matchesSearch && matchesCat;
    });
  }, [revenueDetails, revenueSearch, revenueCategoryFilter]);

  // Aggregate stats for revenue detailing
  const revenueStats = useMemo(() => {
    if (!revenueDetails || revenueDetails.length === 0) {
      return { totalSum: 0, totalCount: 0, servicesSum: 0, goodsSum: 0 };
    }
    let totalSum = 0;
    let totalCount = 0;
    let servicesSum = 0;
    let goodsSum = 0;
    revenueDetails.forEach(item => {
      totalSum += item.sum;
      totalCount += item.count;
      if (item.category.toLowerCase().includes('услуг')) {
        servicesSum += item.sum;
      } else if (item.category.toLowerCase().includes('товар')) {
        goodsSum += item.sum;
      }
    });
    return {
      totalSum,
      totalCount,
      servicesSum,
      goodsSum
    };
  }, [revenueDetails]);

  // Processing results: Salary Payments
  const salaryResults = useMemo(() => {
    return processSalaryPayments(rawRecords, parsedCutoffDate);
  }, [rawRecords, parsedCutoffDate]);

  const { employees: salaryEmployees, metrics: salaryMetrics } = salaryResults;

  // KPIs
  const totalRevenue = useMemo(() => {
    return financialReport.find(r => r.label === "Выручка общая")?.value || 0;
  }, [financialReport]);

  const rawExpensesRow = financialReport.find(r => r.label === "Расходы");
  const totalExpenses = rawExpensesRow ? rawExpensesRow.value : 0;

  const rawNetProfitRow = financialReport.find(r => r.label === "ЧИСТАЯ ПРИБЫЛЬ");
  const netProfit = rawNetProfitRow ? rawNetProfitRow.value : 0;

  const rawProfitabilityRow = financialReport.find(r => r.label === "Рентабельность");
  const profitability = rawProfitabilityRow ? rawProfitabilityRow.value : 0;

  // Edit cell value actions
  const startEditingRow = (idx: number, currentVal: number) => {
    setEditingRowIdx(idx);
    setEditingRowValue(String(currentVal));
  };

  const saveRowOverride = (idx: number) => {
    let num = parseFloat(editingRowValue);
    if (isNaN(num)) {
      setNotification({ status: 'error', message: 'Введите корректное числовое значение' });
      return;
    }
    const label = financialReport[idx]?.label;
    if (label === "Зарплата административного персонала") {
      num = -Math.abs(num);
      setServiceAdminSalary(Math.abs(num));
    } else if (label === "Зарплата мастеров") {
      num = -Math.abs(num);
    }
    const updated = { ...userOverrides, [idx]: num };
    setUserOverrides(updated);
    setEditingRowIdx(null);
    setNotification({
      status: 'success',
      message: `Показатель "${financialReport[idx].label}" изменен вручную.`
    });
  };

  const resetRowOverride = (idx: number) => {
    const updated = { ...userOverrides };
    delete updated[idx];
    setUserOverrides(updated);
    setNotification({
      status: 'success',
      message: `Сброшено ручное изменение для "${financialReport[idx].label}"`
    });
  };

  // Header quick statistics
  const fileLoaded = rawRecords.length > 0;

  // Filtered target list of indicators to display
  const filteredReportRows = useMemo(() => {
    return financialReport.filter((row) => {
      // 1. Label matches search query
      const matchesSearch = row.label.toLowerCase().includes(searchQuery.toLowerCase());
      
      // 2. Type matches filters
      let matchesFilter = true;
      if (tableFilter === 'data') {
        matchesFilter = row.type === 'data';
      } else if (tableFilter === 'formulas') {
        matchesFilter = row.type === 'formula';
      } else if (tableFilter === 'nonzero') {
        matchesFilter = row.value !== 0;
      }

      if (!matchesSearch || !matchesFilter) return false;

      // 3. Section collapse filter: If row belongs to a collapsed aggregated section, hide it
      for (const sec of AGGREGATED_SECTIONS) {
        if (collapsedSections[sec.id] && sec.isChildRow(row, financialReport)) {
          return false;
        }
      }

      return true;
    });
  }, [financialReport, searchQuery, tableFilter, collapsedSections]);

  // -------------------------------------------------------------
  // TAB 2 (SERVICES REPORT) LOGIC & COST ALLOCATION
  // -------------------------------------------------------------
  // Variable and Fixed expenses extracted from Financial Report (Tab 1)
  const variableExpensesFromReport = useMemo(() => {
    const row = financialReport.find(r => r.label === "Расходы переменные");
    return row ? Math.abs(row.value) : 0;
  }, [financialReport]);

  const fixedExpensesFromReport = useMemo(() => {
    const row = financialReport.find(r => r.label === "Расходы постоянные");
    return row ? Math.abs(row.value) : 0;
  }, [financialReport]);

  // Supplier expenses extracted from Financial Report (Tab 1, section Поставщики)
  const supplierRowsFromReport = useMemo(() => {
    const startIdx = financialReport.findIndex(r => r.label === "Поставщики");
    const endIdx = financialReport.findIndex(r => r.label === "Расходы на заработную плату");
    if (startIdx !== -1 && endIdx !== -1) {
      return financialReport.slice(startIdx + 1, endIdx).filter(r => r.type === 'data');
    }
    return financialReport.filter(r => r.index >= 54 && r.index <= 86 && r.type === 'data');
  }, [financialReport]);

  const totalSupplierExpensesFromReport = useMemo(() => {
    const row = financialReport.find(r => r.label === "Поставщики");
    return row ? Math.abs(row.value) : supplierRowsFromReport.reduce((acc, r) => acc + Math.abs(r.value), 0);
  }, [financialReport, supplierRowsFromReport]);

  // Admin salary extracted from Financial Report (row Зарплата административного персонала)
  const adminSalaryFromReport = useMemo(() => {
    const row = financialReport.find(r => r.label === "Зарплата административного персонала");
    return row ? Math.abs(row.value) : 85000;
  }, [financialReport]);

  const activeAdminSalary = useMemo(() => {
    if (customAdminSalary !== '') {
      return parseFloat(customAdminSalary) || 0;
    }
    return adminSalaryFromReport;
  }, [customAdminSalary, adminSalaryFromReport]);

  const activeVariableExpenses = useMemo(() => {
    if (isServicesCustomExpensesActive && customVariableExpenses !== '') {
      return parseFloat(customVariableExpenses) || 0;
    }
    return variableExpensesFromReport;
  }, [isServicesCustomExpensesActive, customVariableExpenses, variableExpensesFromReport]);

  const activeFixedExpenses = useMemo(() => {
    if (isServicesCustomExpensesActive && customFixedExpenses !== '') {
      return parseFloat(customFixedExpenses) || 0;
    }
    return fixedExpensesFromReport;
  }, [isServicesCustomExpensesActive, customFixedExpenses, fixedExpensesFromReport]);

  const activeCustomSupplierExpenses = useMemo(() => {
    if (isServicesCustomExpensesActive && customSupplierExpenses !== '') {
      return parseFloat(customSupplierExpenses) || 0;
    }
    return 0;
  }, [isServicesCustomExpensesActive, customSupplierExpenses]);

  // Main Services Report Generation with supplier distribution according to regulation table, hall salaries, areas, and chosen expense allocation method
  const servicesReport = useMemo(() => {
    return generateServicesReport(
      servicesRecords, 
      activeVariableExpenses, 
      activeFixedExpenses,
      supplierRowsFromReport,
      activeCustomSupplierExpenses,
      hallMastersSalaries,
      serviceAdminSalary,
      hallAreas,
      serviceHallOverrides,
      expenseAllocationMethod
    );
  }, [
    servicesRecords, 
    activeVariableExpenses, 
    activeFixedExpenses, 
    supplierRowsFromReport, 
    activeCustomSupplierExpenses,
    hallMastersSalaries,
    serviceAdminSalary,
    hallAreas,
    serviceHallOverrides,
    expenseAllocationMethod
  ]);

  // Items currently remaining in "Прочие услуги"
  const otherServicesItems = useMemo(() => {
    const grp = servicesReport.halls.find(h => h.hallName === 'Прочие услуги');
    return grp ? grp.items : [];
  }, [servicesReport]);

  // Filtered Services List for Search & Selected Hall
  const filteredServiceItems = useMemo(() => {
    let list = servicesReport.items;
    if (selectedHallFilter !== 'all') {
      list = list.filter(item => item.hall === selectedHallFilter);
    }
    if (servicesSearchQuery.trim()) {
      const q = servicesSearchQuery.toLowerCase().trim();
      list = list.filter(item => item.service.toLowerCase().includes(q));
    }
    return list;
  }, [servicesReport.items, selectedHallFilter, servicesSearchQuery]);

  // Filtered Halls List for Hall-Grouped View
  const filteredHalls = useMemo(() => {
    const q = servicesSearchQuery.toLowerCase().trim();
    return servicesReport.halls.map(hallGroup => {
      if (selectedHallFilter !== 'all' && hallGroup.hallName !== selectedHallFilter) {
        return null;
      }
      let items = hallGroup.items;
      if (q) {
        items = items.filter(it => it.service.toLowerCase().includes(q));
      }
      if (items.length === 0 && (q || selectedHallFilter !== 'all')) {
        return null;
      }
      return {
        ...hallGroup,
        items
      };
    }).filter((h): h is HallReportGroup => h !== null);
  }, [servicesReport.halls, selectedHallFilter, servicesSearchQuery]);

  // Toggle individual revenue row detailing in Financial Report (Tab 1)
  const toggleRevenueDetailRow = (rowIdx: number) => {
    setCollapsedRevenueDetails(prev => ({
      ...prev,
      [rowIdx]: !prev[rowIdx]
    }));
  };

  // Bulk expand/collapse revenue details in Financial Report (Tab 1)
  const toggleAllTab1RevenueDetails = (collapse: boolean) => {
    if (collapse) {
      const all: Record<number, boolean> = {};
      for (let i = 0; i <= 7; i++) all[i] = true;
      setCollapsedRevenueDetails(all);
      setShowRevenueDetailInline(false);
    } else {
      setCollapsedRevenueDetails({});
      setShowRevenueDetailInline(true);
    }
  };

  // Bulk expand/collapse handlers for Aggregated Sections in Report 1 (Financial Report)
  const toggleAllSections = (collapse: boolean) => {
    if (!collapse) {
      setCollapsedSections({});
      setCollapsedRevenueDetails({});
      setShowRevenueDetailInline(true);
    } else {
      const all: Record<string, boolean> = {};
      AGGREGATED_SECTIONS.forEach(sec => {
        all[sec.id] = true;
      });
      setCollapsedSections(all);
      const allRev: Record<number, boolean> = {};
      for (let i = 0; i <= 7; i++) allRev[i] = true;
      setCollapsedRevenueDetails(allRev);
      setShowRevenueDetailInline(false);
    }
  };

  // Bulk expand/collapse handlers for Report 1 (Financial Report - transactions)
  const toggleAllReportRows = (expand: boolean) => {
    if (!expand) {
      setExpandedReportRows({});
      const allRev: Record<number, boolean> = {};
      for (let i = 0; i <= 7; i++) allRev[i] = true;
      setCollapsedRevenueDetails(allRev);
    } else {
      const all: Record<number, boolean> = {};
      filteredReportRows.forEach(r => {
        if (r.type !== 'section') {
          all[r.index] = true;
        }
      });
      setExpandedReportRows(all);
      setCollapsedRevenueDetails({});
      setShowRevenueDetailInline(true);
    }
  };

  const toggleAllRevenueDetails = (expand: boolean) => {
    if (!expand) {
      setExpandedRevenueItems({});
    } else {
      const all: Record<string, boolean> = {};
      filteredRevenueDetails.forEach(item => {
        all[item.name] = true;
      });
      setExpandedRevenueItems(all);
    }
  };

  const toggleAllExtraRows = (expand: boolean) => {
    if (!expand) {
      setExpandedExtraRows({});
    } else {
      const all: Record<number, boolean> = {};
      extraRows.forEach((_, idx) => {
        all[idx] = true;
      });
      setExpandedExtraRows(all);
    }
  };

  // Bulk expand/collapse handlers for Report 2 (Services Report)
  const toggleAllHalls = (collapse: boolean) => {
    if (!collapse) {
      setCollapsedHalls({});
    } else {
      const all: Record<string, boolean> = {};
      filteredHalls.forEach(h => {
        all[h.hallName] = true;
      });
      setCollapsedHalls(all);
    }
  };

  const toggleAllServiceItems = (expand: boolean) => {
    if (!expand) {
      setExpandedServiceItems({});
    } else {
      const all: Record<string, boolean> = {};
      filteredServiceItems.forEach(item => {
        all[item.service] = true;
      });
      setExpandedServiceItems(all);
    }
  };

  // Handle uploaded services file with numeric preprocessing
  const processServicesFile = (selectedFile: File) => {
    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const rawData = evt.target?.result;
        if (!rawData) throw new Error('Не удалось прочитать файл');
        const dataArray = new Uint8Array(rawData as ArrayBuffer);
        const workbook = XLSX.read(dataArray, { type: 'array', cellDates: true });
        const { records, errors, stats } = parseServicesSpreadsheet(workbook);

        if (errors.length > 0) {
          setNotification({ status: 'error', message: errors.join(' ') });
          return;
        }

        setServicesRecords(records);
        setServicesFileName(selectedFile.name);
        setServicesStats(stats);
        setNotification({
          status: 'success',
          message: `Файл по услугам «${selectedFile.name}» (${records.length} строк) успешно загружен! Выполнена предобработка числовых показателей: очищено и нормализовано ${stats.cleanedFieldsCount} значений (валюты, пробелы, скобки, разделители). Выручка рассчитана.`
        });
      } catch (err: any) {
        setNotification({
          status: 'error',
          message: `Ошибка чтения файла по услугам: ${err.message || 'неподдерживаемый формат'}`
        });
      }
    };
    reader.readAsArrayBuffer(selectedFile);
  };

  const handleServicesFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const selectedFile = e.target.files?.[0];
    if (!selectedFile) return;
    processServicesFile(selectedFile);
  };

  const handleServicesDrop = (e: React.DragEvent) => {
    e.preventDefault();
    e.stopPropagation();
    setServicesDragActive(false);

    if (e.dataTransfer.files && e.dataTransfer.files[0]) {
      const selectedFile = e.dataTransfer.files[0];
      const ext = selectedFile.name.split('.').pop()?.toLowerCase();
      if (ext !== 'xlsx' && ext !== 'xls' && ext !== 'csv') {
        setNotification({
          status: 'error',
          message: 'Недопустимый формат файла. Пожалуйста, загрузите файл .xlsx, .xls или .csv'
        });
        return;
      }
      processServicesFile(selectedFile);
    }
  };

  const loadDemoServicesData = () => {
    const structured: RawServiceRecord[] = DEMO_SERVICES_RECORDS.map((dem, idx) => {
      const rev = dem.points + dem.subscriptions + dem.certificates + dem.clientAccounts + dem.money;
      return {
        ...dem,
        id: `demo_svc_${idx}_${Date.now()}`,
        revenue: parseFloat(rev.toFixed(2))
      };
    });
    setServicesRecords(structured);
    setServicesFileName('Услуги_салона_(25.01.2026_-_31.01.2026).xlsx');
    setServicesStats({
      totalRows: DEMO_SERVICES_RECORDS.length,
      validServiceRows: DEMO_SERVICES_RECORDS.length,
      cleanedFieldsCount: DEMO_SERVICES_RECORDS.length * 6,
      recognizedColumns: {
        service: 'Услуга',
        discount: 'Скидка',
        points: 'Баллы',
        subscriptions: 'Абонементы',
        certificates: 'Сертификаты',
        clientAccounts: 'Оплачено со счетов клиентов, ₽',
        money: 'Оплачено деньгами, ₽'
      }
    });
    if (rawRecords.length === 0) {
      loadDemoData();
    }
    setNotification({
      status: 'success',
      message: 'Демо-данные по услугам успешно загружены! Предобработка числовых показателей выполнена, затраты распределены пропорционально выручке.'
    });
  };

  // ---------------------------------------------------------------------------
  // Интеграция с Yclients (данные приходят с серверного прокси /api/yclients/*)
  // ---------------------------------------------------------------------------

  // Вкладка 1: финансовые операции
  const loadFinanceFromYclients = async (startIso: string, endIso: string): Promise<string> => {
    const { records, diagnostics } = await fetchYclientsFinance(startIso, endIso);
    if (records.length === 0) {
      throw new Error('За выбранный период Yclients не вернул финансовых операций.');
    }
    const structured: RawRecord[] = records.map((r, idx) => ({
      ...r,
      id: `yc_rec_${idx}_${Date.now()}`,
      date: parseDate(r.dateStr),
      rawIndex: idx + 1,
    }));
    const period = `${isoToRu(startIso)}_-_${isoToRu(endIso)}`;
    handleDataParsed(structured, `Финансовые_операции_Yclients_(${period}).xlsx`);

    const lines = [
      `Загружено операций: ${records.length} (доходов: ${diagnostics.incomeCount}, расходов: ${diagnostics.expenseCount}, переводов между кассами: ${diagnostics.transfersSkipped}).`,
    ];
    if (diagnostics.incomeCount > 0) {
      lines.push(`Доходов с расшифровкой по услугам: ${diagnostics.incomeLinkedToRecords}, без услуги: ${diagnostics.incomeWithoutService}.`);
    }
    diagnostics.warnings.forEach((w) => lines.push(w));
    lines.push('Нажмите «Запустить программу», чтобы сформировать отчёты.');
    return lines.join('\n');
  };

  // Вкладка 2: отчёт по услугам
  const loadServicesFromYclients = async (startIso: string, endIso: string): Promise<string> => {
    const { records, diagnostics } = await fetchYclientsServices(startIso, endIso);
    if (records.length === 0) {
      throw new Error(
        `За выбранный период не найдено оказанных услуг (записей: ${diagnostics.recordsTotal}, пропущено: ${diagnostics.recordsSkipped} — удалённые или клиент не пришёл).`
      );
    }
    const structured: RawServiceRecord[] = records.map((r, idx) => {
      const rev = r.points + r.subscriptions + r.certificates + r.clientAccounts + r.money;
      return {
        ...r,
        id: `yc_svc_${idx}_${Date.now()}`,
        rawIndex: idx + 1,
        revenue: parseFloat(rev.toFixed(2)),
      };
    });
    const period = `${isoToRu(startIso)}_-_${isoToRu(endIso)}`;
    setServicesRecords(structured);
    setServicesFileName(`Услуги_салона_Yclients_(${period}).xlsx`);
    setServicesStats({
      totalRows: structured.length,
      validServiceRows: structured.length,
      cleanedFieldsCount: 0,
      recognizedColumns: {
        service: 'Yclients: услуга записи',
        discount: 'Yclients: скидка',
        points: 'Yclients: баллы',
        subscriptions: 'Yclients: абонементы',
        certificates: 'Yclients: сертификаты',
        clientAccounts: 'Yclients: счета клиентов',
        money: 'Yclients: оплачено деньгами'
      }
    });
    setNotification({
      status: 'success',
      message: `Данные по услугам загружены из Yclients: ${structured.length} строк из ${diagnostics.recordsUsed} записей.`
    });

    const lines = [
      `Загружено строк услуг: ${structured.length} (записей в отчёте: ${diagnostics.recordsUsed} из ${diagnostics.recordsTotal}).`,
    ];
    if (diagnostics.paymentBreakdownFromTransactions > 0) {
      lines.push(`Разбивка способов оплаты взята из финансовых операций по ${diagnostics.paymentBreakdownFromTransactions} записям.`);
    }
    diagnostics.warnings.forEach((w) => lines.push(w));
    return lines.join('\n');
  };

  return (
    <div className="min-h-screen bg-gray-50/70 pb-20 font-sans">
      <input
        type="file"
        ref={fileInputRef}
        onChange={handleFileUpload}
        accept=".xlsx,.xls,.csv"
        className="hidden"
      />
      <input
        type="file"
        ref={servicesFileInputRef}
        onChange={handleServicesFileUpload}
        accept=".xlsx,.xls,.csv"
        className="hidden"
      />
      
      {/* Upper Brand Premium Bar */}
      <header className="bg-white border-b border-gray-100 shadow-sm sticky top-0 z-40 backdrop-blur-md bg-white/90">
        <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 py-4 flex flex-col sm:flex-row items-start sm:items-center justify-between gap-4">
          <div className="flex items-center gap-3">
            <div className="w-10 h-10 rounded-xl bg-indigo-600 flex items-center justify-center shadow-lg shadow-indigo-150">
              <TrendingUp className="w-5 h-5 text-white" />
            </div>
            <div>
              <h1 className="font-display font-extrabold text-xl text-gray-900 tracking-tight flex items-center gap-2">
                AI‑Помощник Финансовых Отчетов
                <span className="text-[10px] px-2 py-0.5 rounded-full font-sans font-medium bg-indigo-50 text-indigo-700 tracking-normal border border-indigo-100">v1.2</span>
              </h1>
              <p className="text-[11px] text-gray-500 mt-0.5">
                Автоматизированный разбор статей доходов, расходов и умных зарплатных периодов
              </p>
            </div>
          </div>

          <div className="flex items-center gap-3 w-full sm:w-auto">
            {/* Cutoff Date Input with Visual Validation */}
            <div className="flex-1 sm:flex-initial flex items-center gap-2.5 bg-gray-50 px-3.5 py-1.5 rounded-xl border border-gray-150 relative">
              <Calendar className="w-4 h-4 text-indigo-500" />
              <div className="flex flex-col text-left">
                <label className="text-[9px] text-gray-400 font-medium uppercase tracking-wider">Дата Отсечения ЗП</label>
                <input
                  type="text"
                  value={cutoffDateStr}
                  onChange={(e) => setCutoffDateStr(e.target.value)}
                  className="bg-transparent font-mono text-xs font-semibold text-gray-800 border-none outline-none focus:ring-0 p-0 w-24"
                  placeholder="ДД.ММ.ГГГГ"
                  title="Укажите дату отсечения для разделения периодов выплат (все выплаты до этой даты попадут в прошлый период)"
                />
              </div>
              
              {isCutoffDateValid ? (
                <span className="w-2 h-2 rounded-full bg-emerald-500" title="Дата верна" />
              ) : (
                <span className="w-2 h-2 rounded-full bg-amber-500 animate-ping" title="Невалидный формат даты. Ожидается ДД.ММ.ГГГГ" />
              )}
            </div>

            {/* Quick Demo Loader */}
            {stagedRecords.length === 0 && (
              <button
                onClick={loadDemoData}
                className="px-4 py-2 text-xs font-semibold text-indigo-700 bg-indigo-50 border border-indigo-150 hover:bg-indigo-100/70 transition-all rounded-xl cursor-pointer"
              >
                Загрузить Демо
              </button>
            )}

            {/* Detailed User Guide Button */}
            <button
              type="button"
              onClick={() => setShowUserGuideModal(true)}
              className="px-3.5 py-2 text-xs font-semibold text-gray-700 bg-white border border-gray-250 hover:bg-indigo-50 hover:text-indigo-700 hover:border-indigo-200 transition-all rounded-xl cursor-pointer flex items-center gap-1.5 shadow-2xs"
              title="Открыть подробное руководство пользователя"
            >
              <BookOpen className="w-4 h-4 text-indigo-600" />
              <span className="hidden sm:inline">Руководство</span>
            </button>

            {stagedRecords.length > 0 && !isProcessed && (
              <button
                onClick={handleRunProgram}
                disabled={isProcessing || !isCutoffDateValid}
                className="px-4 py-2 text-xs font-semibold text-white bg-indigo-600 hover:bg-indigo-700 transition-all rounded-xl cursor-pointer flex items-center gap-1.5 shadow-md shadow-indigo-100 disabled:opacity-55 disabled:cursor-not-allowed animate-pulse"
                title="Запустить расчет и скачать отчеты"
              >
                <Play className="w-3.5 h-3.5 fill-current" />
                <span>Запуск</span>
              </button>
            )}

            {isProcessed && (
              <button
                onClick={handleRunProgram}
                disabled={isProcessing || !isCutoffDateValid}
                className="px-4 py-2 text-xs font-semibold text-indigo-700 bg-indigo-50 border border-indigo-150 hover:bg-indigo-100 transition-all rounded-xl cursor-pointer flex items-center gap-1.5 disabled:opacity-50"
                title="Перезапустить обработку программы и повторно скачать отчеты"
              >
                <RefreshCw className="w-3.5 h-3.5" />
                <span>Запустить повторно</span>
              </button>
            )}
          </div>
        </div>
      </header>

      {/* Main Container */}
      <main className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 mt-6">
        
        {/* Dynamic Alerts/Notifications */}
        <AnimatePresence>
          {notification && (
            <motion.div 
              initial={{ opacity: 0, y: -10 }}
              animate={{ opacity: 1, y: 0 }}
              exit={{ opacity: 0, y: -10 }}
              className={`p-4 mb-6 rounded-xl border flex items-start gap-3 justify-between ${
                notification.status === 'success' 
                  ? 'bg-emerald-50 text-emerald-800 border-emerald-150' 
                  : 'bg-rose-50 text-rose-800 border-rose-150'
              }`}
            >
              <div className="flex gap-2.5 items-start">
                {notification.status === 'success' ? (
                  <CheckCircle className="w-5 h-5 text-emerald-600 flex-shrink-0 mt-0.5" />
                ) : (
                  <AlertCircle className="w-5 h-5 text-rose-600 flex-shrink-0 mt-0.5" />
                )}
                <span className="text-xs font-medium leading-relaxed">{notification.message}</span>
              </div>
              <button 
                onClick={() => setNotification(null)}
                className="text-[10px] uppercase font-semibold text-gray-400 hover:text-gray-600 px-1"
              >
                Закрыть
              </button>
            </motion.div>
          )}
        </AnimatePresence>

        {/* Primary Tabs Switcher: Вкладка 1 (Финансовый отчет) | Вкладка 2 (Отчет по услугам) */}
        <div className="bg-white p-2 rounded-2xl border border-gray-200/80 shadow-xs mb-6 flex flex-col sm:flex-row items-stretch sm:items-center justify-between gap-3">
          <div className="flex items-center gap-2">
            <button
              onClick={() => setPrimaryTab('finance')}
              className={`flex-1 sm:flex-initial flex items-center justify-center gap-2.5 px-5 py-2.5 rounded-xl font-bold text-xs transition-all cursor-pointer ${
                primaryTab === 'finance'
                  ? 'bg-indigo-600 text-white shadow-md shadow-indigo-150'
                  : 'text-gray-600 hover:text-gray-900 hover:bg-gray-50'
              }`}
            >
              <Table className="w-4 h-4" />
              <span>Вкладка 1: Финансовый отчет</span>
              {isProcessed && (
                <span className="w-2 h-2 rounded-full bg-emerald-400" title="Данные обработаны"></span>
              )}
            </button>

            <button
              onClick={() => setPrimaryTab('services')}
              className={`flex-1 sm:flex-initial flex items-center justify-center gap-2.5 px-5 py-2.5 rounded-xl font-bold text-xs transition-all cursor-pointer ${
                primaryTab === 'services'
                  ? 'bg-indigo-600 text-white shadow-md shadow-indigo-150'
                  : 'text-gray-600 hover:text-gray-900 hover:bg-gray-50'
              }`}
            >
              <Scissors className="w-4 h-4" />
              <span>Вкладка 2: Отчет по услугам</span>
              {servicesRecords.length > 0 ? (
                <span className={`px-2 py-0.5 rounded-full text-[10px] font-bold ${
                  primaryTab === 'services' ? 'bg-white/20 text-white' : 'bg-indigo-50 text-indigo-700 border border-indigo-150'
                }`}>
                  {servicesReport.items.length} услуг
                </span>
              ) : (
                <span className="px-2 py-0.5 rounded-full text-[10px] font-semibold bg-amber-50 text-amber-700 border border-amber-150">
                  Загрузить файл
                </span>
              )}
            </button>
          </div>

          <div className="text-[11px] text-gray-500 flex items-center gap-2 px-3 py-1.5 bg-gray-50 rounded-xl self-end sm:self-auto border border-gray-100">
            {primaryTab === 'finance' ? (
              <span>Основной финансовый отчет, зарплаты и детализация</span>
            ) : (
              <span>Группировка услуг, расчет выручки и распределение затрат</span>
            )}
          </div>
        </div>

        {/* TAB 1: ФИНАНСОВЫЙ ОТЧЕТ */}
        {primaryTab === 'finance' && (
          <div>
            {/* If no file uploaded, show Upload Prompt Jumbotron */}
            {!isProcessed ? (
          <div>
            {isProcessing ? (
              /* Premium Step-by-Step Processing Loader */
              <div className="max-w-3xl mx-auto bg-white rounded-3xl border border-gray-150 p-10 shadow-xl my-10 relative overflow-hidden">
                <div className="absolute top-0 left-0 h-1.5 bg-indigo-600 transition-all duration-500" style={{ width: `${(processingStep / 4) * 100}%` }}></div>
                
                <div className="text-center mb-8">
                  <div className="w-16 h-16 rounded-2xl bg-indigo-50 flex items-center justify-center mx-auto mb-4 animate-bounce">
                    <Loader2 className="w-8 h-8 text-indigo-600 animate-spin" />
                  </div>
                  <h2 className="font-display font-extrabold text-2xl text-gray-900 tracking-tight">Обработка финансовых данных</h2>
                  <p className="text-xs text-gray-500 mt-1 max-w-sm mx-auto">Пожалуйста, подождите. Алгоритм структурирует статьи доходов и расходов в реальном времени</p>
                </div>

                <div className="space-y-6 max-w-lg mx-auto">
                  <div className="flex items-start gap-4 p-3.5 rounded-2xl transition-all duration-300 border border-gray-100 bg-gray-50">
                    <div className="flex-shrink-0 w-8 h-8 rounded-full flex items-center justify-center font-bold text-xs bg-indigo-600 text-white">1</div>
                    <div className="flex-1">
                      <div className="flex items-center justify-between">
                        <span className="font-bold text-xs text-gray-800">Нормализация и разбор транзакций</span>
                        {processingStep > 1 ? (
                          <div className="w-5 h-5 rounded-full bg-emerald-100 flex items-center justify-center text-emerald-700 font-bold text-[10px]"><Check className="w-3 h-3" /></div>
                        ) : processingStep === 1 ? (
                          <Loader2 className="w-4 h-4 text-indigo-600 animate-spin" />
                        ) : (
                          <span className="w-1.5 h-1.5 rounded-full bg-gray-300"></span>
                        )}
                      </div>
                      <p className="text-[11px] text-gray-500 mt-0.5">Чтение исходных Excel-строк: сумм, дат, категорий и примечаний.</p>
                    </div>
                  </div>

                  <div className="flex items-start gap-4 p-3.5 rounded-2xl transition-all duration-300 border border-gray-100 bg-gray-50">
                    <div className="flex-shrink-0 w-8 h-8 rounded-full flex items-center justify-center font-bold text-xs bg-indigo-600 text-white">2</div>
                    <div className="flex-1">
                      <div className="flex items-center justify-between">
                        <span className="font-bold text-xs text-gray-800">Фильтрация поставщиков косметики и материалов</span>
                        {processingStep > 2 ? (
                          <div className="w-5 h-5 rounded-full bg-emerald-100 flex items-center justify-center text-emerald-700 font-bold text-[10px]"><Check className="w-3 h-3" /></div>
                        ) : processingStep === 2 ? (
                          <Loader2 className="w-4 h-4 text-indigo-600 animate-spin" />
                        ) : (
                          <span className="w-1.5 h-1.5 rounded-full bg-gray-300"></span>
                        )}
                      </div>
                      <p className="text-[11px] text-gray-500 mt-0.5">Поиск вхождений контрагентов (Loreal, Matrix, ProfiLine) в статьи расходов.</p>
                    </div>
                  </div>

                  <div className="flex items-start gap-4 p-3.5 rounded-2xl transition-all duration-300 border border-gray-100 bg-gray-50">
                    <div className="flex-shrink-0 w-8 h-8 rounded-full flex items-center justify-center font-bold text-xs bg-indigo-600 text-white">3</div>
                    <div className="flex-1">
                      <div className="flex items-center justify-between">
                        <span className="font-bold text-xs text-gray-800">Умный расчет зарплатных периодов</span>
                        {processingStep > 3 ? (
                          <div className="w-5 h-5 rounded-full bg-emerald-100 flex items-center justify-center text-emerald-700 font-bold text-[10px]"><Check className="w-3 h-3" /></div>
                        ) : processingStep === 3 ? (
                          <Loader2 className="w-4 h-4 text-indigo-600 animate-spin" />
                        ) : (
                          <span className="w-1.5 h-1.5 rounded-full bg-gray-300"></span>
                        )}
                      </div>
                      <p className="text-[11px] text-gray-500 mt-0.5">Разделение на прошлые и отчетные выплаты по дате отсечения {cutoffDateStr}.</p>
                    </div>
                  </div>

                  <div className="flex items-start gap-4 p-3.5 rounded-2xl transition-all duration-300 border border-gray-100 bg-gray-50">
                    <div className="flex-shrink-0 w-8 h-8 rounded-full flex items-center justify-center font-bold text-xs bg-indigo-600 text-white">4</div>
                    <div className="flex-1">
                      <div className="flex items-center justify-between">
                        <span className="font-bold text-xs text-gray-800">Автоматическая выгрузка XLSX-файлов</span>
                        {processingStep > 4 ? (
                          <div className="w-5 h-5 rounded-full bg-emerald-100 flex items-center justify-center text-emerald-700 font-bold text-[10px]"><Check className="w-3 h-3" /></div>
                        ) : processingStep === 4 ? (
                          <Loader2 className="w-4 h-4 text-indigo-600 animate-spin" />
                        ) : (
                          <span className="w-1.5 h-1.5 rounded-full bg-gray-300"></span>
                        )}
                      </div>
                      <p className="text-[11px] text-gray-500 mt-0.5">Генерация двух сводных таблиц (Финансовый отчет и Зарплатная ведомость).</p>
                    </div>
                  </div>
                </div>
              </div>
            ) : stagedRecords.length > 0 ? (
              /* Gorgeous Staging Preparation View */
              <div className="grid grid-cols-1 lg:grid-cols-12 gap-6 my-4">
                
                {/* Left Card: Staging Info & Parameters */}
                <div className="lg:col-span-7 flex flex-col gap-6">
                  <div className="bg-white rounded-3xl border border-gray-150 p-8 shadow-sm flex flex-col justify-between flex-grow relative overflow-hidden">
                    <div className="absolute top-0 right-0 p-4 opacity-[0.03] text-indigo-900">
                      <FileText className="w-48 h-48 pointer-events-none" />
                    </div>

                    <div>
                      <span className="px-3 py-1 bg-amber-50 text-amber-700 border border-amber-100 text-[10px] rounded-lg uppercase tracking-wider font-bold">Файлы готовы</span>
                      <h2 className="font-display font-extrabold text-2xl text-gray-900 tracking-tight mt-3">Файл успешно загружен в память</h2>
                      <p className="text-xs text-gray-500 mt-1">Доступ к данным полностью защищен и готов к запуску программы обработки.</p>
                      
                      {/* File metadata badges */}
                      <div className="mt-6 p-4 rounded-2xl bg-gray-50 border border-gray-100 flex flex-col sm:flex-row gap-4 justify-between items-start sm:items-center">
                        <div className="flex items-center gap-3">
                          <div className="w-10 h-10 rounded-xl bg-indigo-50 flex items-center justify-center">
                            <Table className="w-5 h-5 text-indigo-600" />
                          </div>
                          <div>
                            <div className="font-bold text-xs text-gray-800 truncate max-w-[220px]" title={stagedFileName || ''}>{stagedFileName}</div>
                            <div className="text-[10px] text-gray-500 mt-0.5">{stagedRecords.length} транзакций распознано</div>
                          </div>
                        </div>

                        <div className="flex items-center gap-1.5 bg-indigo-50 border border-indigo-100/50 text-indigo-700 px-3 py-1.5 rounded-xl font-mono text-[11px] font-bold">
                          <CheckCircle className="w-3.5 h-3.5" />
                          <span>Проверка валидности: ОК</span>
                        </div>
                      </div>

                      {/* Setup Cutoff Parameter right here */}
                      <div className="mt-8 border-t border-gray-100 pt-6">
                        <h3 className="font-display font-bold text-xs uppercase tracking-wider text-gray-400 mb-3 block">Укажите или проверьте дату отсечения зарплат</h3>
                        <div className="flex items-start gap-4">
                          <div className="flex items-center gap-2.5 bg-gray-55 px-4 py-2.5 rounded-xl border border-gray-150 relative">
                            <Calendar className="w-5 h-5 text-indigo-500" />
                            <div className="flex flex-col text-left">
                              <label className="text-[9px] text-gray-400 font-medium uppercase tracking-wider">Дата отрезка ЗП</label>
                              <input
                                type="text"
                                value={cutoffDateStr}
                                onChange={(e) => setCutoffDateStr(e.target.value)}
                                className="bg-transparent font-mono text-sm font-semibold text-gray-800 border-none outline-none focus:ring-0 p-0 w-28"
                                placeholder="ДД.ММ.ГГГГ"
                                title="Укажите дату отсечения"
                              />
                            </div>
                            
                            {isCutoffDateValid ? (
                              <span className="w-2.5 h-2.5 rounded-full bg-emerald-500" title="Дата верна" />
                            ) : (
                              <span className="w-2.5 h-2.5 rounded-full bg-amber-500 animate-ping" title="Невалидный формат даты. Ожидается ДД.ММ.ГГГГ" />
                            )}
                          </div>

                          <p className="text-[11px] text-gray-500 leading-relaxed pt-1.5">
                            Все транзакции со словом «Зарплата» в категории и ключевыми словами в комментарии с датой <strong>строго до</strong> указанной будут отнесены к прошлому периоду. Остальные — к отчетному.
                          </p>
                        </div>
                      </div>

                      {/* Manual Salaries Input right here */}
                      <div className="mt-8 border-t border-gray-100 pt-6">
                        <h3 className="font-display font-bold text-xs uppercase tracking-wider text-gray-400 mb-3 block">Ручной ввод финансовых показателей</h3>
                        <p className="text-[11px] text-gray-500 mb-4">
                          Заполните эти поля, если хотите дополнить или переопределить суммы выплат зарплаты напрямую при загрузке документа:
                        </p>
                        
                        <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                          <div className="flex items-center gap-2.5 bg-gray-55 px-4 py-2.5 rounded-xl border border-gray-150 relative">
                            <div className="flex flex-col text-left w-full">
                              <label className="text-[9px] text-gray-400 font-medium uppercase tracking-wider">Зарплата административного персонала (₽)</label>
                              <input
                                type="text"
                                id="manual_admins_salary_input"
                                value={manualAdminsSalary}
                                onChange={(e) => setManualAdminsSalary(e.target.value)}
                                className="bg-transparent font-mono text-sm font-semibold text-gray-800 border-none outline-none focus:ring-0 p-0 w-full"
                                placeholder="Например: 125000"
                              />
                            </div>
                          </div>

                          <div className="flex items-center gap-2.5 bg-gray-55 px-4 py-2.5 rounded-xl border border-gray-150 relative">
                            <div className="flex flex-col text-left w-full">
                              <label className="text-[9px] text-gray-400 font-medium uppercase tracking-wider">Зарплата мастеров (₽)</label>
                              <input
                                type="text"
                                id="manual_masters_salary_input"
                                value={manualMastersSalary}
                                onChange={(e) => setManualMastersSalary(e.target.value)}
                                className="bg-transparent font-mono text-sm font-semibold text-gray-800 border-none outline-none focus:ring-0 p-0 w-full"
                                placeholder="Например: 340000"
                              />
                            </div>
                          </div>
                        </div>
                      </div>
                    </div>

                    <div className="mt-8 pt-6 border-t border-gray-100 flex flex-col sm:flex-row gap-3">
                      <button
                        onClick={handleRunProgram}
                        disabled={!isCutoffDateValid}
                        className={`flex-1 py-3.5 px-6 rounded-xl font-display font-bold text-xs text-center flex items-center justify-center gap-2.5 shadow-md shadow-indigo-100 transition-all duration-300 ${
                          isCutoffDateValid 
                            ? 'bg-indigo-600 hover:bg-indigo-700 text-white cursor-pointer active:scale-95' 
                            : 'bg-gray-200 text-gray-400 cursor-not-allowed'
                        }`}
                      >
                        <Play className="w-4 h-4 fill-current" />
                        Запустить расчет и сформировать отчеты
                      </button>
                      <button
                        onClick={() => {
                          setStagedRecords([]);
                          setStagedFileName(null);
                        }}
                        className="py-3.5 px-5 rounded-xl border border-gray-250 text-gray-500 hover:text-gray-700 bg-white hover:bg-gray-50 text-xs font-bold transition-all cursor-pointer"
                      >
                        Сбросить файл
                      </button>
                    </div>
                  </div>
                </div>

                {/* Right Card: Guide Sidebar info summary */}
                <div className="lg:col-span-5 flex flex-col gap-6">
                  <div className="bg-white rounded-3xl border border-gray-150 p-6 shadow-sm flex-grow">
                    <h3 className="font-display font-semibold text-gray-900 text-md flex items-center gap-2 mb-4">
                      <BookOpen className="w-4 h-4 text-indigo-500" />
                      Что произойдет при запуске?
                    </h3>
                    
                    <ul className="space-y-4 text-xs text-gray-600 leading-relaxed font-sans">
                      <li className="flex items-start gap-2.5">
                        <span className="w-5 h-5 rounded-full bg-indigo-50 text-indigo-600 flex items-center justify-center font-bold text-[10px] mt-0.5">1</span>
                        <div>
                          <strong className="text-gray-800 block">Разделение заработной платы:</strong> 
                          Ведомости по каждому сотруднику будут разбиты на периоды и собраны с остатками.
                        </div>
                      </li>
                      <li className="flex items-start gap-2.5">
                        <span className="w-5 h-5 rounded-full bg-indigo-50 text-indigo-600 flex items-center justify-center font-bold text-[10px] mt-0.5">2</span>
                        <div>
                          <strong className="text-gray-800 block">Группировка Salon-Suppliers:</strong> 
                          Программа сама найдет расходы на Matrix красители, Loreal витрину и Profiline.
                        </div>
                      </li>
                      <li className="flex items-start gap-2.5">
                        <span className="w-5 h-5 rounded-full bg-indigo-50 text-indigo-600 flex items-center justify-center font-bold text-[10px] mt-0.5">3</span>
                        <div>
                          <strong className="text-gray-800 block">Автоматический экспорт:</strong> 
                          Готовые файлы <span className="font-semibold text-indigo-750">«Финансовый_отчет.xlsx»</span> и <span className="font-semibold text-emerald-800">«Зарплатная_ведомость.xlsx»</span> автоматически скачаются на компьютер.
                        </div>
                      </li>
                    </ul>

                    <div className="mt-8 pt-6 border-t border-gray-100">
                      <div className="flex items-center gap-2 text-xs text-indigo-600 bg-indigo-50/50 p-3.5 rounded-xl border border-indigo-100/30">
                        <AlertCircle className="w-4.5 h-4.5 flex-shrink-0" />
                        <span>Система полностью готова. Нажмите кнопку <strong>«Запустить обработку»</strong> слева или кнопку <strong>«Запуск»</strong> в верхнем меню.</span>
                      </div>
                    </div>
                  </div>
                </div>

                <div className="lg:col-span-12">
                  <TestPlayground />
                </div>
              </div>
            ) : (
              /* Original File Upload dropzone */
              <div className="grid grid-cols-1 lg:grid-cols-12 gap-6 my-4">
                
                {/* Left Big Upload Box */}
                <div className="lg:col-span-7">
                  <div 
                    onDragEnter={handleDrag}
                    onDragOver={handleDrag}
                    onDragLeave={handleDrag}
                    onDrop={handleDrop}
                    className={`border-2 border-dashed rounded-3xl p-10 text-center flex flex-col items-center justify-center min-h-[420px] transition-all duration-300 ${
                      dragActive 
                        ? 'border-indigo-600 bg-indigo-50/40 scale-[0.99] shadow-inner' 
                        : 'border-gray-250 bg-white hover:border-indigo-400'
                    }`}
                  >
                    <div className="w-16 h-16 rounded-2xl bg-indigo-50 flex items-center justify-center mb-5 transition-transform hover:scale-105">
                      <Upload className="w-8 h-8 text-indigo-600" />
                    </div>
                    
                    <h2 className="font-display font-bold text-xl text-gray-900 tracking-tight">
                      Перетащите файл финансового отчета
                    </h2>
                    <p className="text-sm text-gray-500 mt-2 max-w-sm mx-auto">
                      Рекомендуемые форматы: <code className="font-mono bg-gray-50 px-1 py-0.5 text-xs text-indigo-600">.xlsx</code> или <code className="font-mono bg-gray-50 px-1 py-0.5 text-xs text-indigo-600">.csv</code>. 
                      Минимальные требуемые колонки: «Сумма, Р», «Назначение», «Контрагент», «Комментарий», «Дата».
                    </p>

                    <div className="mt-8 flex gap-3">
                      <button
                        onClick={() => fileInputRef.current?.click()}
                        className="px-6 py-3 bg-indigo-600 text-white rounded-xl text-xs font-semibold shadow-md shadow-indigo-200 hover:bg-indigo-700 transition-colors cursor-pointer"
                      >
                        Выбрать файл на устройстве
                      </button>
                    </div>
                  </div>

                  <div className="mt-4">
                    <YclientsPanel
                      title="Загрузить финансовые операции из Yclients"
                      description="Операции по кассам и счетам за выбранный период загрузятся напрямую из Yclients, файл выписки не нужен."
                      onLoad={loadFinanceFromYclients}
                    />
                  </div>
                </div>

                {/* Right Guide & FAQ Sidebar */}
                <div className="lg:col-span-5 flex flex-col justify-between gap-6">
                  <div className="bg-white rounded-3xl border border-gray-150 p-6 shadow-sm flex-1">
                    <h3 className="font-display font-semibold text-gray-900 text-md flex items-center gap-2 mb-4">
                      <BookOpen className="w-4 h-4 text-indigo-500" />
                      Руководство по обработке данных
                    </h3>
                    
                    <ul className="space-y-4 text-xs text-gray-600 leading-relaxed font-sans">
                      <li className="flex items-start gap-2.5">
                        <span className="w-5 h-5 rounded-full bg-indigo-50 text-indigo-600 flex items-center justify-center font-bold text-[10px] mt-0.5">1</span>
                        <div>
                          <strong className="text-gray-800">Умное объединение:</strong> 
                          Система ищет вхождение строк отчёта в исходные данные (например, исходная <code className="px-1 bg-gray-100 rounded">2026 Аренда</code> запишется в строке <code className="px-1 bg-gray-100 rounded">Аренда</code>).
                        </div>
                      </li>
                      <li className="flex items-start gap-2.5">
                        <span className="w-5 h-5 rounded-full bg-indigo-50 text-indigo-600 flex items-center justify-center font-bold text-[10px] mt-0.5">2</span>
                        <div>
                          <strong className="text-gray-800">Фильтрация Поставщиков:</strong> 
                          Закупки материалов/товаров автоматически распределяются по соответствующим контрагентам (из строк 54–86).
                        </div>
                      </li>
                      <li className="flex items-start gap-2.5">
                        <span className="w-5 h-5 rounded-full bg-indigo-50 text-indigo-600 flex items-center justify-center font-bold text-[10px] mt-0.5">3</span>
                        <div>
                          <strong className="text-gray-800">Разделение зарплат:</strong> 
                          Формируется отдельная зарплатная ведомость, группируемая по ФИО, с разделением на «прошлый период» на основе указанной вверху даты отсечения.
                        </div>
                      </li>
                    </ul>

                    <div className="mt-8 pt-6 border-t border-gray-100 space-y-3">
                      <div className="flex items-center gap-2 text-xs text-indigo-600 bg-indigo-50/50 p-3.5 rounded-xl border border-indigo-100/30">
                        <AlertCircle className="w-4.5 h-4.5 flex-shrink-0" />
                        <span>Для быстрого ознакомления нажмите <strong>«Загрузить Демо»</strong> в верхнем меню.</span>
                      </div>
                      <button
                        type="button"
                        onClick={() => setShowUserGuideModal(true)}
                        className="w-full py-2.5 px-3 bg-white hover:bg-gray-50 text-indigo-700 border border-indigo-200 rounded-xl text-xs font-semibold transition-all flex items-center justify-center gap-2 cursor-pointer shadow-2xs"
                      >
                        <BookOpen className="w-4 h-4 text-indigo-600" />
                        <span>Открыть подробное руководство пользователя</span>
                      </button>
                    </div>
                  </div>
                </div>

                <div className="lg:col-span-12">
                  <TestPlayground />
                </div>
              </div>
            )}
          </div>
        ) : (
          /* Main Dashboard with Tabs on data loaded */
          <div className="space-y-6">
            
            {/* KPI Summary Bento Grid */}
            <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4">
              
              {/* Loaded stats card */}
              <div className="bg-white rounded-2xl border border-gray-150 p-5 shadow-sm">
                <div className="text-[10px] uppercase font-bold text-gray-400 tracking-wider">Файл Распознан</div>
                <div className="font-display font-extrabold text-lg text-gray-900 mt-2 truncate max-w-[220px]" title={fileName || ''}>
                  {fileName}
                </div>
                <div className="flex items-center gap-1.5 text-xs text-indigo-600 mt-1">
                  <FileText className="w-3.5 h-3.5" />
                  <span>{rawRecords.length} записей импортировано</span>
                </div>
              </div>

              {/* Total revenues card */}
              <div className="bg-white rounded-2xl border border-gray-150 p-5 shadow-sm">
                <div className="text-[10px] uppercase font-bold text-gray-400 tracking-wider">Выручка Общая</div>
                <div className="font-display font-extrabold text-2xl text-emerald-600 mt-2">
                  {totalRevenue.toLocaleString('ru-RU')} ₽
                </div>
                <div className="text-xs text-gray-500 mt-1">
                  На основе исходных статей
                </div>
              </div>

              {/* Expenses card */}
              <div className="bg-white rounded-2xl border border-gray-150 p-5 shadow-sm">
                <div className="text-[10px] uppercase font-bold text-gray-400 tracking-wider">Всего Расходов</div>
                <div className="font-display font-extrabold text-2xl text-rose-500 mt-2">
                  {totalExpenses.toLocaleString('ru-RU')} ₽
                </div>
                <div className="text-xs text-gray-500 mt-1">
                  Автоматический расчет
                </div>
              </div>

              {/* Theoretical Profit card */}
              <div className="bg-white rounded-2xl border border-gray-150 p-5 shadow-sm">
                <div className="text-[10px] uppercase font-bold text-gray-400 tracking-wider">Чистая Прибыль / Рентабельность</div>
                <div className="font-display font-extrabold text-2xl text-indigo-700 mt-2">
                  {netProfit.toLocaleString('ru-RU')} ₽
                </div>
                <div className="flex items-center gap-1.5 text-xs text-gray-500 mt-1">
                  <span className="font-mono bg-indigo-50 text-indigo-700 font-bold px-1.5 py-0.5 rounded text-[10px]">
                    {profitability.toFixed(1)}% р-сть
                  </span>
                  <span>теоретический итог</span>
                </div>
              </div>

            </div>

            {/* Crucial Custom Message block for salary counts metrics */}
            <div className="bg-indigo-50/50 rounded-2xl p-4.5 border border-indigo-100/40 flex flex-col md:flex-row items-start md:items-center justify-between gap-4">
              <div className="flex gap-3 items-start">
                <Users className="w-5 h-5 text-indigo-600 mt-0.5 flex-shrink-0" />
                <div>
                  <h4 className="text-xs font-bold text-indigo-950">Статистика по Зарплатным Выплатам</h4>
                  <p className="text-[11px] text-gray-600 mt-0.5 line-clamp-2 md:line-clamp-none">
                    «Из <strong className="text-indigo-900">{salaryMetrics.totalCount}</strong> записей <strong className="text-indigo-900">{salaryMetrics.matchingSalaryCount}</strong> соответствуют критериям зарплатных выплат, из них <strong className="text-indigo-900">{salaryMetrics.pastPeriodCount}</strong> отнесены к прошлому периоду».
                  </p>
                </div>
              </div>
              <div className="flex gap-2 flex-wrap">
                <span className="text-[9px] uppercase font-bold bg-amber-50 text-amber-700 border border-amber-100/70 p-1.5 rounded-lg">
                  Прошлый период: {salaryMetrics.pastPeriodCount} выплат
                </span>
                <span className="text-[9px] uppercase font-bold bg-emerald-50 text-emerald-700 border border-emerald-100/70 p-1.5 rounded-lg">
                  Отчетный период: {salaryMetrics.reportPeriodCount} выплат
                </span>
              </div>
            </div>

            {/* Download and Tabs control block */}
            <div className="bg-white rounded-2xl border border-gray-150 p-4 shadow-sm flex flex-col md:flex-row justify-between items-center gap-4">
              {/* Tab selector */}
              <div className="flex bg-gray-50 p-1.5 rounded-xl border border-gray-100 w-full md:w-auto overflow-x-auto">
                <button
                  onClick={() => setActiveTab('financial')}
                  className={`flex-1 md:flex-none flex items-center justify-center gap-2 px-4 py-2 rounded-lg text-xs font-semibold whitespace-nowrap transition-all cursor-pointer ${
                    activeTab === 'financial' 
                      ? 'bg-white text-indigo-700 shadow-sm' 
                      : 'text-gray-500 hover:text-gray-700'
                  }`}
                >
                  <Table className="w-4 h-4" />
                  Финансовый отчет
                </button>
                <button
                  onClick={() => setActiveTab('revenue-details')}
                  className={`flex-1 md:flex-none flex items-center justify-center gap-2 px-4 py-2 rounded-lg text-xs font-semibold whitespace-nowrap transition-all cursor-pointer ${
                    activeTab === 'revenue-details' 
                      ? 'bg-white text-indigo-700 shadow-sm' 
                      : 'text-gray-500 hover:text-gray-700'
                  }`}
                >
                  <TrendingUp className="w-4 h-4" />
                  <span>Детализация доходов</span>
                  {revenueDetails && revenueDetails.length > 0 && (
                    <span className="px-1.5 py-0.2 rounded-full text-[10px] bg-indigo-100 text-indigo-800 font-bold">
                      {revenueDetails.length}
                    </span>
                  )}
                </button>
                <button
                  onClick={() => setActiveTab('salary')}
                  className={`flex-1 md:flex-none flex items-center justify-center gap-2 px-4 py-2 rounded-lg text-xs font-semibold whitespace-nowrap transition-all cursor-pointer ${
                    activeTab === 'salary' 
                      ? 'bg-white text-indigo-700 shadow-sm' 
                      : 'text-gray-500 hover:text-gray-700'
                  }`}
                >
                  <Users className="w-4 h-4" />
                  Зарплатные выплаты
                </button>
                <button
                  onClick={() => setActiveTab('raw')}
                  className={`flex-1 md:flex-none flex items-center justify-center gap-2 px-4 py-2 rounded-lg text-xs font-semibold whitespace-nowrap transition-all cursor-pointer ${
                    activeTab === 'raw' 
                      ? 'bg-white text-indigo-700 shadow-sm' 
                      : 'text-gray-500 hover:text-gray-700'
                  }`}
                >
                  <FileText className="w-4 h-4" />
                  Исходная таблица
                </button>
                <button
                  onClick={() => setPrimaryTab('services')}
                  className="flex-1 md:flex-none flex items-center justify-center gap-1.5 px-3 py-2 rounded-lg text-xs font-semibold whitespace-nowrap transition-all cursor-pointer text-indigo-700 bg-indigo-50/70 hover:bg-indigo-100 border border-indigo-150 ml-1"
                  title="Перейти во 2-ю вкладку: Отчет по услугам с распределением затрат"
                >
                  <Scissors className="w-3.5 h-3.5" />
                  <span>Вкладка 2: Отчет по услугам →</span>
                </button>
              </div>

              {/* Downloads Actions */}
              <div className="flex gap-2 flex-wrap w-full md:w-auto">
                <button
                  onClick={() => fileInputRef.current?.click()}
                  className="flex-1 md:flex-initial flex items-center justify-center gap-2 px-4 py-2 bg-indigo-600 hover:bg-indigo-700 text-white rounded-xl text-xs font-bold shadow-md shadow-indigo-100 transition-all cursor-pointer"
                  title="Выбрать новый исходный файл для расчета нового отчета"
                >
                  <RefreshCw className="w-4 h-4" />
                  <span>Рассчитать новый отчет</span>
                </button>
                <button
                  onClick={() => downloadFinancialSpreadsheet(financialReport, extraRows, revenueDetails)}
                  className="flex-1 md:flex-initial flex items-center justify-center gap-2 px-4 py-2 bg-indigo-50 border border-indigo-150 hover:bg-indigo-100 text-indigo-700 rounded-xl text-xs font-semibold transition-all cursor-pointer"
                  title="Скачать полный финансовый отчет со сводным и подробным листом"
                >
                  <Download className="w-3.5 h-3.5" />
                  Метрики (XLSX)
                </button>
                <button
                  onClick={() => downloadRevenueSpreadsheet(revenueDetails, totalRevenue)}
                  className="flex-1 md:flex-initial flex items-center justify-center gap-2 px-4 py-2 bg-amber-50 border border-amber-200 hover:bg-amber-100 text-amber-800 rounded-xl text-xs font-semibold transition-all cursor-pointer"
                  title="Скачать отдельную ведомость по графе Услуга/Товар"
                >
                  <Download className="w-3.5 h-3.5" />
                  Услуги/Товары (XLSX)
                </button>
                <button
                  onClick={() => downloadSalarySpreadsheet(salaryEmployees)}
                  className="flex-1 md:flex-initial flex items-center justify-center gap-2 px-4 py-2 bg-emerald-600 hover:bg-emerald-700 text-white rounded-xl text-xs font-semibold shadow-sm transition-all cursor-pointer"
                  title="Скачать сформированную зарплатную таблицу"
                >
                  <Download className="w-3.5 h-3.5" />
                  Зарплатная ведомость (XLSX)
                </button>
                <button
                  onClick={() => downloadDocxReport(financialReport, rawRecords, categoryMappings, revenueDetails)}
                  className="flex-1 md:flex-initial flex items-center justify-center gap-2 px-4 py-2 bg-blue-600 hover:bg-blue-700 text-white rounded-xl text-xs font-semibold shadow-sm transition-all cursor-pointer"
                  title="Скачать финансовый отчет в формате .docx с детализацией доходов"
                >
                  <FileText className="w-3.5 h-3.5" />
                  Отчет-ведомость (DOCX)
                </button>
                <button
                  onClick={() => {
                    setRawRecords([]);
                    setStagedRecords([]);
                    setFileName(null);
                    setStagedFileName(null);
                    setUserOverrides({});
                    setCategoryMappings({});
                    setTransactionMappings({});
                    setIsProcessed(false);
                    setNotification({
                      status: 'success',
                      message: 'Все очищено. Вы можете загрузить новые данные.'
                    });
                  }}
                  className="px-3 py-2 bg-gray-50 hover:bg-gray-100 border border-gray-200 text-gray-500 rounded-xl text-xs font-semibold transition-all cursor-pointer"
                  title="Очистить и сбросить все данные"
                >
                  <RefreshCw className="w-3.5 h-3.5" />
                </button>
              </div>
            </div>

            {/* Financial tab content */}
            {activeTab === 'financial' && (
              <div className="space-y-6">
                
                {/* Filters Row */}
                <div className="bg-white rounded-2xl border border-gray-150 p-4.5 shadow-sm flex flex-col sm:flex-row gap-4 items-center justify-between">
                  {/* Search query input */}
                  <div className="flex items-center gap-2 bg-gray-50 px-3.5 py-1.5 rounded-xl border border-gray-150 w-full sm:max-w-sm">
                    <Search className="w-4 h-4 text-gray-400" />
                    <input
                      type="text"
                      value={searchQuery}
                      onChange={(e) => setSearchQuery(e.target.value)}
                      placeholder="Поиск по показателям..."
                      className="bg-transparent text-xs w-full outline-none text-gray-700 border-none inline-block focus:ring-0"
                    />
                    {searchQuery && (
                      <button onClick={() => setSearchQuery('')} className="text-gray-400 hover:text-gray-600 text-xs text-[11px] select-none font-medium">Очистить</button>
                    )}
                  </div>

                  {/* Toggle and Dropdown filters */}
                  <div className="flex items-center gap-2.5 w-full sm:w-auto flex-wrap">
                    {/* Bulk Aggregated Sections Collapsing Buttons */}
                    <div className="flex items-center gap-1 bg-gray-50 p-1 rounded-xl border border-gray-150">
                      <button
                        onClick={() => toggleAllSections(false)}
                        className="flex items-center gap-1 px-2.5 py-1 rounded-lg text-xs font-semibold text-gray-700 hover:text-indigo-700 hover:bg-white transition-all cursor-pointer shadow-2xs"
                        title="Развернуть детализацию всех укрупненных статей отчета"
                      >
                        <Eye className="w-3.5 h-3.5 text-indigo-600" />
                        <span>Развернуть все статьи</span>
                      </button>
                      <button
                        onClick={() => toggleAllSections(true)}
                        className="flex items-center gap-1 px-2.5 py-1 rounded-lg text-xs font-semibold text-gray-700 hover:text-indigo-700 hover:bg-white transition-all cursor-pointer shadow-2xs"
                        title="Свернуть детализации по всем укрупненным статьям (компактный P&L вид)"
                      >
                        <EyeOff className="w-3.5 h-3.5 text-amber-600" />
                        <span>Свернуть все статьи</span>
                      </button>
                    </div>

                    {/* Bulk Expand / Collapse Transactions Buttons */}
                    <div className="flex items-center gap-1 bg-gray-50 p-1 rounded-xl border border-gray-150">
                      <button
                        onClick={() => toggleAllReportRows(true)}
                        className="flex items-center gap-1 px-2.5 py-1 rounded-lg text-xs font-semibold text-gray-700 hover:text-indigo-700 hover:bg-white transition-all cursor-pointer shadow-2xs"
                        title="Развернуть детализацию транзакций всех статей отчета"
                      >
                        <ChevronsUpDown className="w-3.5 h-3.5 text-indigo-600" />
                        <span>Транзакции</span>
                      </button>
                      <button
                        onClick={() => toggleAllReportRows(false)}
                        className="flex items-center gap-1 px-2.5 py-1 rounded-lg text-xs font-semibold text-gray-700 hover:text-indigo-700 hover:bg-white transition-all cursor-pointer shadow-2xs"
                        title="Свернуть открытые транзакции"
                      >
                        <ChevronsDownUp className="w-3.5 h-3.5 text-gray-500" />
                        <span>Скрыть</span>
                      </button>
                    </div>

                    {/* Bulk Expand / Collapse Revenue Details Buttons */}
                    <div className="flex items-center gap-1 bg-gray-50 p-1 rounded-xl border border-gray-150">
                      <button
                        onClick={() => toggleAllTab1RevenueDetails(false)}
                        className="flex items-center gap-1 px-2.5 py-1 rounded-lg text-xs font-semibold text-gray-700 hover:text-indigo-700 hover:bg-white transition-all cursor-pointer shadow-2xs"
                        title="Развернуть детализацию доходов во всех статьях выручки"
                      >
                        <TrendingUp className="w-3.5 h-3.5 text-indigo-600" />
                        <span>Доходы: Развернуть</span>
                      </button>
                      <button
                        onClick={() => toggleAllTab1RevenueDetails(true)}
                        className="flex items-center gap-1 px-2.5 py-1 rounded-lg text-xs font-semibold text-gray-700 hover:text-indigo-700 hover:bg-white transition-all cursor-pointer shadow-2xs"
                        title="Свернуть все детализации доходов в статьях выручки"
                      >
                        <ChevronUp className="w-3.5 h-3.5 text-amber-600" />
                        <span>Свернуть доходы</span>
                      </button>
                    </div>

                    <button
                      onClick={() => {
                        const nextVal = !showRevenueDetailInline;
                        setShowRevenueDetailInline(nextVal);
                        if (!nextVal) {
                          const allRev: Record<number, boolean> = {};
                          for (let i = 0; i <= 7; i++) allRev[i] = true;
                          setCollapsedRevenueDetails(allRev);
                        } else {
                          setCollapsedRevenueDetails({});
                        }
                      }}
                      className={`flex items-center gap-1.5 px-3 py-1.5 rounded-xl border text-xs font-semibold transition-all cursor-pointer ${
                        showRevenueDetailInline
                          ? 'bg-indigo-50 border-indigo-200 text-indigo-700 shadow-sm'
                          : 'bg-gray-50 border-gray-150 text-gray-500 hover:text-gray-700'
                      }`}
                      title="Показывать детализацию по графе Услуга/Товар прямо в строках отчета"
                    >
                      <Layers className="w-3.5 h-3.5" />
                      <span>Детализация по «Услуга/Товар»: {showRevenueDetailInline ? 'Вкл' : 'Выкл'}</span>
                    </button>

                    <div className="flex group items-center gap-2 bg-gray-50 px-3.5 py-1.5 rounded-xl border border-gray-150">
                      <Filter className="w-4 h-4 text-gray-400" />
                      <span className="text-xs text-gray-500 font-medium whitespace-nowrap">Показывать:</span>
                      <select
                        value={tableFilter}
                        onChange={(e) => setTableFilter(e.target.value as any)}
                        className="bg-transparent border-none text-xs text-gray-700 outline-none pr-6 focus:ring-0 font-medium cursor-pointer"
                      >
                        <option value="all">Все строки</option>
                        <option value="data">Только вводимые статьи</option>
                        <option value="formulas">Только формулы</option>
                        <option value="nonzero">Только ненулевые</option>
                      </select>
                    </div>
                  </div>
                </div>

                {/* Quick Section Collapse Chips Bar */}
                <div className="bg-white rounded-2xl border border-gray-150 p-3.5 shadow-2xs flex flex-wrap items-center gap-2">
                  <span className="text-xs font-bold text-gray-600 flex items-center gap-1 mr-1">
                    <SlidersHorizontal className="w-3.5 h-3.5 text-indigo-600" />
                    <span>Сворачивание по статьям:</span>
                  </span>
                  {AGGREGATED_SECTIONS.map((sec) => {
                    const isCollapsed = !!collapsedSections[sec.id];
                    const childCount = financialReport.filter(r => sec.isChildRow(r, financialReport)).length;
                    return (
                      <button
                        key={sec.id}
                        type="button"
                        onClick={() => setCollapsedSections(prev => ({ ...prev, [sec.id]: !prev[sec.id] }))}
                        className={`px-2.5 py-1 rounded-xl text-xs font-semibold transition-all flex items-center gap-1.5 border cursor-pointer ${
                          isCollapsed
                            ? 'bg-amber-50/90 text-amber-900 border-amber-300 hover:bg-amber-100'
                            : 'bg-white hover:bg-indigo-50/60 text-gray-750 hover:text-indigo-900 border-gray-200 shadow-2xs'
                        }`}
                        title={`${isCollapsed ? 'Развернуть' : 'Свернуть'} детализацию статьи «${sec.title}» (${childCount} позиций)`}
                      >
                        <span>{sec.icon}</span>
                        <span>{sec.shortTitle}</span>
                        <span className={`text-[10px] font-mono px-1.5 py-0.2 rounded-md ${isCollapsed ? 'bg-amber-200 text-amber-950 font-bold' : 'bg-gray-100 text-gray-600'}`}>
                          {childCount}
                        </span>
                        {isCollapsed ? (
                          <span className="text-[10px] text-amber-700 font-bold ml-0.5">свернуто</span>
                        ) : (
                          <ChevronDown className="w-3 h-3 text-gray-400" />
                        )}
                      </button>
                    );
                  })}
                </div>

                {/* Main spreadsheet card */}
                <div className="bg-white rounded-3xl border border-gray-150 shadow-sm overflow-hidden">
                  <div className="overflow-x-auto">
                    <table className="w-full text-left text-xs border-collapse">
                      <thead>
                        <tr className="bg-gray-50 border-b border-gray-150 text-gray-600 font-bold select-none h-11">
                          <th className="p-3.5 pl-6 w-16 text-center font-mono text-[10px] text-gray-400">№</th>
                          <th className="p-3.5">Показатели</th>
                          <th className="p-3.5 w-40">Тип статьи</th>
                          <th className="p-3.5 text-right w-64 pr-6">Сумма (Руб)</th>
                        </tr>
                      </thead>
                      <tbody className="divide-y divide-gray-100 font-sans">
                        {filteredReportRows.map((row) => {
                          const isFormula = row.type === 'formula';
                          const isHeaderSection = row.type === 'section';
                          const isDataRow = !isFormula && !isHeaderSection;
                          const hasTransactions = row.transactions && row.transactions.length > 0;
                          const isExpanded = !!expandedReportRows[row.index];
                          const sectionForHeaderRow = AGGREGATED_SECTIONS.find(sec => sec.headerLabel === row.label);
                          const isSectionCollapsed = sectionForHeaderRow ? !!collapsedSections[sectionForHeaderRow.id] : false;
                          const childCount = sectionForHeaderRow ? financialReport.filter(r => sectionForHeaderRow.isChildRow(r, financialReport)).length : 0;
                          
                          const isRevenueRow = (row.index >= 1 && row.index <= 7) && !!row.revenueDetails && row.revenueDetails.length > 0;
                          const isRevDetailCollapsed = isRevenueRow ? (!!collapsedRevenueDetails[row.index] || !showRevenueDetailInline) : false;

                          // Style based on type
                          let bgStyle = "bg-white hover:bg-gray-50/50";
                          if (isHeaderSection) {
                            bgStyle = "bg-gray-50/90 font-bold text-gray-900";
                          } else if (isFormula) {
                            bgStyle = "bg-indigo-50/10 font-medium hover:bg-indigo-50/25";
                          } else if (isDataRow) {
                            bgStyle = "bg-white hover:bg-indigo-50/10 cursor-pointer";
                          }
                          if (sectionForHeaderRow) {
                            bgStyle = "bg-indigo-50/30 hover:bg-indigo-50/60 font-semibold cursor-pointer border-t border-b border-indigo-150/50";
                          }

                          return (
                            <React.Fragment key={row.index}>
                              <tr 
                                className={`transition-colors h-11 ${bgStyle}`}
                                onClick={() => {
                                  if (sectionForHeaderRow) {
                                    setCollapsedSections(prev => ({
                                      ...prev,
                                      [sectionForHeaderRow.id]: !prev[sectionForHeaderRow.id]
                                    }));
                                  } else if (isRevenueRow) {
                                    toggleRevenueDetailRow(row.index);
                                  } else if (isDataRow) {
                                    setExpandedReportRows(prev => ({ ...prev, [row.index]: !prev[row.index] }));
                                  }
                                }}
                              >
                                <td className="p-3 text-center font-mono text-[11px] text-gray-400 border-r border-gray-50">
                                  {row.index + 1}
                                </td>
                                <td className="p-3 pl-4 font-sans font-medium text-gray-800">
                                  <div className="flex items-center gap-2">
                                    {sectionForHeaderRow ? (
                                      <button
                                        type="button"
                                        onClick={(e) => {
                                          e.stopPropagation();
                                          setCollapsedSections(prev => ({
                                            ...prev,
                                            [sectionForHeaderRow.id]: !prev[sectionForHeaderRow.id]
                                          }));
                                        }}
                                        className="shrink-0 p-1 rounded-md bg-indigo-100 hover:bg-indigo-200 text-indigo-700 transition-colors"
                                        title={isSectionCollapsed ? "Развернуть детализацию статьи" : "Свернуть детализацию статьи"}
                                      >
                                        {isSectionCollapsed ? (
                                          <ChevronRight className="w-3.5 h-3.5 text-indigo-800" />
                                        ) : (
                                          <ChevronDown className="w-3.5 h-3.5 text-indigo-800" />
                                        )}
                                      </button>
                                    ) : isRevenueRow ? (
                                      <button
                                        type="button"
                                        onClick={(e) => {
                                          e.stopPropagation();
                                          toggleRevenueDetailRow(row.index);
                                        }}
                                        className="shrink-0 p-1 rounded-md bg-indigo-50 hover:bg-indigo-150 text-indigo-700 transition-colors cursor-pointer"
                                        title={isRevDetailCollapsed ? "Развернуть детализацию доходов" : "Свернуть детализацию доходов"}
                                      >
                                        <ChevronDown className={`w-3.5 h-3.5 text-indigo-800 transition-transform ${isRevDetailCollapsed ? '-rotate-90' : ''}`} />
                                      </button>
                                    ) : isDataRow ? (
                                      <ChevronDown className={`w-3.5 h-3.5 text-indigo-500 transition-transform shrink-0 ${isExpanded ? 'rotate-180' : ''}`} />
                                    ) : null}

                                    <span className={isHeaderSection ? "uppercase tracking-wide text-indigo-950 font-bold text-[11px]" : isFormula ? "text-indigo-900 font-semibold" : "text-gray-750 font-semibold"}>
                                      {row.label}
                                    </span>

                                    {sectionForHeaderRow && childCount > 0 && (
                                      <button
                                        type="button"
                                        onClick={(e) => {
                                          e.stopPropagation();
                                          setCollapsedSections(prev => ({
                                            ...prev,
                                            [sectionForHeaderRow.id]: !prev[sectionForHeaderRow.id]
                                          }));
                                        }}
                                        className={`px-2 py-0.5 rounded-lg text-[10px] font-bold border transition-all cursor-pointer ${
                                          isSectionCollapsed 
                                            ? 'bg-amber-100/80 text-amber-900 border-amber-300 hover:bg-amber-200' 
                                            : 'bg-indigo-100/60 text-indigo-800 border-indigo-250 hover:bg-indigo-200'
                                        }`}
                                      >
                                        {isSectionCollapsed ? `[+] Свернуто (${childCount})` : `[-] ${childCount} статей`}
                                      </button>
                                    )}

                                    {isRevenueRow && (
                                      <button
                                        type="button"
                                        onClick={(e) => {
                                          e.stopPropagation();
                                          toggleRevenueDetailRow(row.index);
                                        }}
                                        className={`px-2 py-0.5 rounded-lg text-[10px] font-bold border transition-all cursor-pointer ${
                                          isRevDetailCollapsed 
                                            ? 'bg-amber-100/80 text-amber-900 border-amber-300 hover:bg-amber-200' 
                                            : 'bg-indigo-100/60 text-indigo-800 border-indigo-250 hover:bg-indigo-200'
                                        }`}
                                        title={isRevDetailCollapsed ? "Развернуть детализацию доходов" : "Свернуть детализацию доходов"}
                                      >
                                        {isRevDetailCollapsed ? `[+] Детализация (${row.revenueDetails!.length})` : `[-] Свернуть (${row.revenueDetails!.length})`}
                                      </button>
                                    )}

                                    {row.label === "Зарплата мастеров" && (
                                      <span className="px-2 py-0.5 rounded-lg text-[10px] font-semibold bg-sky-50 text-sky-850 border border-sky-200" title="Сумма рассчитана по залам из вкладки «Отчет по услугам»">
                                        По залам: {totalMastersSalaryFromServices.toLocaleString('ru-RU')} ₽
                                      </span>
                                    )}

                                    {row.label === "Зарплата административного персонала" && (
                                      <span className="px-2 py-0.5 rounded-lg text-[10px] font-semibold bg-teal-50 text-teal-850 border border-teal-200" title="Введено на вкладке «Отчет по услугам» (распределяется пропорционально выручке)">
                                        Из отчета по услугам: {serviceAdminSalary.toLocaleString('ru-RU')} ₽
                                      </span>
                                    )}

                                    {isDataRow && (
                                      <span className="text-[10px] text-gray-400 font-normal font-mono" title={`${row.transactions ? row.transactions.length : 0} транзакций из исходного файла`}>
                                        ({row.transactions ? row.transactions.length : 0})
                                      </span>
                                    )}
                                  </div>
                                </td>
                                <td className="p-3">
                                  {isHeaderSection ? (
                                    <span className="px-2 py-0.5 rounded text-[9px] uppercase font-bold bg-indigo-50 text-indigo-700 border border-indigo-100/50">Раздел</span>
                                  ) : isFormula ? (
                                    <div className="flex flex-col">
                                      <span className="px-2 py-0.5 rounded text-[9px] uppercase font-bold bg-indigo-50/50 text-indigo-600 border border-indigo-100/30 inline-block w-max">Формула</span>
                                      {row.formula && <span className="text-[9px] text-gray-400 mt-0.5">{row.formula}</span>}
                                    </div>
                                  ) : (
                                    <span className="px-2 py-0.5 rounded text-[9px] uppercase font-bold bg-gray-100 text-gray-600 border border-gray-150/40">Статья ввода</span>
                                  )}
                                </td>
                                
                                <td className="p-3 text-right pr-6">
                                  <div className="flex items-center justify-end gap-2.5">
                                    {row.isCustomOverride && (
                                      <button 
                                        onClick={(e) => {
                                          e.stopPropagation();
                                          resetRowOverride(row.index);
                                        }}
                                        className="text-amber-600 hover:text-amber-700 flex items-center gap-0.5 text-[9px] font-bold bg-amber-50 px-1 py-0.5 rounded border border-amber-100"
                                        title="Сбросить ручную коррекцию"
                                      >
                                        <Undo className="w-3 h-3" /> Коррекция
                                      </button>
                                    )}

                                    {editingRowIdx === row.index ? (
                                      // Live edit cell input
                                      <div className="flex items-center gap-1.5" onClick={(e) => e.stopPropagation()}>
                                        <input
                                          type="number"
                                          value={editingRowValue}
                                          onChange={(e) => setEditingRowValue(e.target.value)}
                                          className="w-28 text-right px-2 py-1 text-xs font-mono border border-indigo-400 rounded-lg focus:outline-none focus:ring-1 focus:ring-indigo-400 text-gray-800"
                                          title="Введите сумму в рублях"
                                          autoFocus
                                        />
                                        <button 
                                          onClick={() => saveRowOverride(row.index)}
                                          className="px-2 py-1 bg-emerald-600 text-white text-[10px] uppercase font-bold rounded-md hover:bg-emerald-700"
                                        >
                                          Сохр
                                        </button>
                                        <button 
                                          onClick={() => setEditingRowIdx(null)}
                                          className="px-2 py-1 bg-gray-100 text-gray-500 text-[10px] uppercase font-bold rounded-md hover:bg-gray-200"
                                        >
                                          Отмена
                                        </button>
                                      </div>
                                    ) : (
                                      // Row value display
                                      <div className="flex items-center gap-2 group-hover:bg-transparent">
                                        <span className={`font-mono font-semibold text-[13px] ${
                                          isFormula 
                                            ? 'text-indigo-800' 
                                            : isHeaderSection 
                                            ? 'text-indigo-950 font-bold' 
                                            : row.value > 0 
                                            ? 'text-gray-900' 
                                            : 'text-gray-400'
                                        }`}>
                                          {row.type === 'formula' && row.label.includes('Рентабельность') 
                                            ? `${row.value.toFixed(1)} %` 
                                            : `${row.value.toLocaleString('ru-RU')} ₽`
                                          }
                                        </span>
                                        
                                        {!isHeaderSection && !isFormula && (
                                          <button
                                            onClick={(e) => {
                                              e.stopPropagation();
                                              startEditingRow(row.index, row.value);
                                            }}
                                            className="text-gray-400 hover:text-indigo-600 p-1 rounded-md opacity-0 group-hover:opacity-100 focus:opacity-100 transition-opacity"
                                            title="Ручная коррекция показателя"
                                          >
                                            <Edit2 className="w-3.5 h-3.5" />
                                          </button>
                                        )}
                                      </div>
                                    )}
                                  </div>
                                </td>
                              </tr>
                              {isRevenueRow && !isRevDetailCollapsed && (
                                <>
                                  {row.revenueDetails!.map((detail, dIdx) => (
                                    <tr 
                                      key={`rev_inline_${row.index}_${dIdx}`}
                                      className="bg-indigo-50/15 hover:bg-indigo-50/35 transition-colors border-l-4 border-l-indigo-400 h-9"
                                    >
                                      <td className="p-2 text-center font-mono text-[10px] text-indigo-400 border-r border-gray-50">
                                        ↳
                                      </td>
                                      <td className="p-2 pl-8 font-sans text-gray-800">
                                        <div className="flex items-center gap-2">
                                          <span className="text-gray-400 font-mono text-xs">↳</span>
                                          <span className="font-medium text-gray-850 text-[11.5px]">{detail.name}</span>
                                          <span className="text-[10px] bg-indigo-50 text-indigo-700 font-mono px-1.5 py-0.2 rounded border border-indigo-100/60" title="Количество сгруппированных строк">
                                            {detail.count} оп.
                                          </span>
                                        </div>
                                      </td>
                                      <td className="p-2">
                                        <span className="px-1.5 py-0.5 rounded text-[8.5px] uppercase font-bold bg-indigo-50/80 text-indigo-650 border border-indigo-100/50">
                                          Услуга/Товар
                                        </span>
                                      </td>
                                      <td className="p-2 text-right pr-6 font-mono font-bold text-gray-850 text-[12px]">
                                        {detail.sum.toLocaleString('ru-RU', { minimumFractionDigits: 2, maximumFractionDigits: 2 })} ₽
                                      </td>
                                    </tr>
                                  ))}
                                  <tr key={`rev_inline_collapse_btn_${row.index}`} className="bg-indigo-50/20 border-b border-indigo-100/40 h-8">
                                    <td colSpan={4} className="py-1 px-8 text-right">
                                      <button
                                        type="button"
                                        onClick={(e) => {
                                          e.stopPropagation();
                                          toggleRevenueDetailRow(row.index);
                                        }}
                                        className="inline-flex items-center gap-1 text-[10.5px] font-semibold text-indigo-700 hover:text-indigo-900 bg-white hover:bg-indigo-50 px-2.5 py-0.5 rounded-md border border-indigo-200 cursor-pointer shadow-2xs transition-colors"
                                        title="Свернуть детализацию доходов этой статьи"
                                      >
                                        <ChevronUp className="w-3 h-3 text-indigo-500" />
                                        <span>Свернуть детализацию доходов «{row.label}»</span>
                                      </button>
                                    </td>
                                  </tr>
                                </>
                              )}
                              {isExpanded && (isDataRow || (row.revenueDetails && row.revenueDetails.length > 0)) && (
                                <tr className="bg-indigo-50-[0.02]">
                                  <td colSpan={4} className="p-4 pl-12 bg-indigo-50/20 border-b border-gray-100">
                                    {/* Header bar with Collapse button */}
                                    <div className="flex items-center justify-between pb-2.5 mb-3 border-b border-indigo-150/60">
                                      <div className="flex items-center gap-2">
                                        <span className="text-xs font-bold text-indigo-950">Детализация статьи #{row.index + 1}: «{row.label}»</span>
                                        <span className="text-[10.5px] text-gray-500 font-mono">
                                          ({row.transactions ? row.transactions.length : 0} оп. • Сумма: {row.value.toLocaleString('ru-RU')} ₽)
                                        </span>
                                      </div>
                                      <button
                                        onClick={(e) => {
                                          e.stopPropagation();
                                          setExpandedReportRows(prev => ({ ...prev, [row.index]: false }));
                                        }}
                                        className="px-2.5 py-1 text-[11px] font-semibold text-indigo-700 hover:text-indigo-900 bg-white hover:bg-indigo-50 rounded-lg border border-indigo-200 transition-all flex items-center gap-1 cursor-pointer shadow-2xs"
                                        title="Свернуть детализацию этой статьи"
                                      >
                                        <ChevronUp className="w-3.5 h-3.5" />
                                        <span>Свернуть детализацию</span>
                                      </button>
                                    </div>

                                    {/* Grouped Revenue Detailing Card if present */}
                                    {row.revenueDetails && row.revenueDetails.length > 0 && (
                                      <div className="mb-4 bg-white p-3.5 rounded-xl border border-indigo-150 shadow-sm">
                                        <div className="flex items-center justify-between mb-2 pb-2 border-b border-gray-100">
                                          <div className="text-[11px] font-bold text-indigo-950 flex items-center gap-1.5">
                                            <Layers className="w-3.5 h-3.5 text-indigo-600" />
                                            <span>Детализация по графе «Услуга/Товар» (совпадающие строки сгруппированы и просуммированы):</span>
                                          </div>
                                          <span className="text-[10px] bg-indigo-50 text-indigo-700 font-mono font-semibold px-2 py-0.5 rounded border border-indigo-100">
                                            {row.revenueDetails.length} уникальных позиций
                                          </span>
                                        </div>
                                        <div className="overflow-x-auto">
                                          <table className="w-full text-left text-[11px] divide-y divide-gray-100">
                                            <thead>
                                              <tr className="bg-indigo-50/30 text-indigo-900 font-bold text-[9px] uppercase select-none">
                                                <th className="p-2 pl-3">№</th>
                                                <th className="p-2">Наименование услуги/товара</th>
                                                <th className="p-2 text-center">Кол-во операций</th>
                                                <th className="p-2 text-center">Доля в статье</th>
                                                <th className="p-2 text-right pr-3">Итоговая сумма</th>
                                              </tr>
                                            </thead>
                                            <tbody className="divide-y divide-gray-50">
                                              {row.revenueDetails.map((item, itIdx) => (
                                                <tr key={itIdx} className="hover:bg-gray-50/60">
                                                  <td className="p-2 pl-3 text-gray-400 font-mono">{itIdx + 1}</td>
                                                  <td className="p-2 font-medium text-gray-800">{item.name}</td>
                                                  <td className="p-2 text-center font-mono font-semibold text-indigo-700">{item.count}</td>
                                                  <td className="p-2 text-center font-mono text-gray-500">
                                                    {row.value > 0 ? `${((item.sum / row.value) * 100).toFixed(1)}%` : '—'}
                                                  </td>
                                                  <td className="p-2 text-right pr-3 font-mono font-bold text-gray-900">
                                                    {item.sum.toLocaleString('ru-RU', { minimumFractionDigits: 2, maximumFractionDigits: 2 })} ₽
                                                  </td>
                                                </tr>
                                              ))}
                                            </tbody>
                                          </table>
                                        </div>
                                      </div>
                                    )}

                                    {hasTransactions ? (
                                      <div className="overflow-x-auto rounded-xl border border-gray-200/60 shadow-sm bg-white">
                                        <table className="min-w-full divide-y divide-gray-150 text-left text-[11px]">
                                          <thead className="bg-indigo-50/50 font-semibold text-indigo-950 uppercase tracking-wider text-[9px] select-none">
                                            <tr>
                                              <th className="p-2.5 pl-4 w-28">Дата</th>
                                              <th className="p-2.5 w-40">Контрагент</th>
                                              <th className="p-2.5 w-44 font-bold text-indigo-950">Услуга/Товар</th>
                                              <th className="p-2.5">Назначение</th>
                                              <th className="p-2.5">Комментарий</th>
                                              <th className="p-2.5 text-right pr-4 w-32">Сумма</th>
                                            </tr>
                                          </thead>
                                          <tbody className="divide-y divide-gray-100 font-sans text-gray-650">
                                            {row.transactions!.map((t, tIdx) => (
                                              <tr key={tIdx} className="hover:bg-gray-50/40 transition-colors">
                                                <td className="p-2.5 pl-4 font-mono text-[10.5px]">
                                                  {t.date ? new Date(t.date).toLocaleDateString('ru-RU') : t.dateStr || '—'}
                                                </td>
                                                <td className="p-2.5 font-medium text-gray-800">
                                                   {t.counterparty || <span className="text-gray-400 italic">Не указан</span>}
                                                 </td>
                                                <td className="p-2.5 font-medium text-indigo-950 max-w-[200px] truncate" title={t.serviceProduct}>
                                                  {t.serviceProduct || <span className="text-gray-300">—</span>}
                                                </td>
                                                <td className="p-2.5 text-gray-700 max-w-[200px] truncate" title={t.category}>
                                                  {t.category || <span className="text-gray-300">—</span>}
                                                </td>
                                                <td className="p-2.5 text-gray-500 max-w-[200px] truncate" title={t.comment}>
                                                  {t.comment || <span className="text-gray-300">—</span>}
                                                </td>
                                                <td className="p-2.5 text-right pr-4 font-mono font-semibold text-gray-800">
                                                  {t.sum.toLocaleString('ru-RU', { minimumFractionDigits: 2, maximumFractionDigits: 2 })} ₽
                                                </td>
                                              </tr>
                                            ))}
                                          </tbody>
                                        </table>
                                      </div>
                                    ) : (
                                      <div className="p-4 text-center text-gray-400 italic bg-white rounded-xl border border-gray-200/60 shadow-sm text-xs">
                                        Нет связанных транзакций в импортированном периоде
                                      </div>
                                    )}

                                    {/* Footer Collapse Button */}
                                    <div className="mt-3 flex justify-end">
                                      <button
                                        onClick={(e) => {
                                          e.stopPropagation();
                                          setExpandedReportRows(prev => ({ ...prev, [row.index]: false }));
                                        }}
                                        className="px-3 py-1 text-[11px] font-semibold text-gray-600 hover:text-indigo-700 bg-white hover:bg-gray-50 rounded-lg border border-gray-200 transition-all flex items-center gap-1 cursor-pointer shadow-2xs"
                                      >
                                        <ChevronUp className="w-3.5 h-3.5 text-gray-400" />
                                        <span>Свернуть детализацию статьи #{row.index + 1}</span>
                                      </button>
                                    </div>
                                  </td>
                                </tr>
                              )}
                            </React.Fragment>
                          );
                        })}
                        
                        {/* Display unmatched додатковые (unexpected) rows if found */}
                        {(extraRows.length > 0 || Object.keys(categoryMappings).length > 0 || Object.keys(transactionMappings).length > 0) && (
                          <>
                            <tr className="bg-amber-50/50 font-bold text-amber-950 h-11 select-none">
                              <td className="p-3 text-center text-amber-400 border-r border-amber-100">+</td>
                              <td className="p-3 pl-4 uppercase tracking-wide text-[10px]" colSpan={2}>
                                Дополнительные нераспределенные показатели и ручные привязки
                              </td>
                              <td className="p-3 text-right pr-6">
                                <div className="flex items-center justify-end gap-1.5">
                                  <button
                                    onClick={() => toggleAllExtraRows(true)}
                                    className="px-2 py-0.5 rounded text-[10px] font-semibold bg-white hover:bg-amber-100 text-amber-900 border border-amber-250 transition-all cursor-pointer shadow-2xs"
                                    title="Развернуть детализацию всех нераспределенных строк"
                                  >
                                    Развернуть все
                                  </button>
                                  <button
                                    onClick={() => toggleAllExtraRows(false)}
                                    className="px-2 py-0.5 rounded text-[10px] font-semibold bg-white hover:bg-amber-100 text-amber-900 border border-amber-250 transition-all cursor-pointer shadow-2xs"
                                    title="Свернуть все нераспределенные строки"
                                  >
                                    Свернуть все
                                  </button>
                                </div>
                              </td>
                            </tr>

                            {/* Active mappings summary view */}
                            {(Object.keys(categoryMappings).length > 0 || Object.keys(transactionMappings).length > 0) && (
                              <tr>
                                <td className="p-4" colSpan={4}>
                                  <div className="bg-indigo-50/30 p-4.5 rounded-2xl border border-indigo-100/40 space-y-3 text-left">
                                    {Object.keys(categoryMappings).length > 0 && (
                                      <div>
                                        <div className="text-[10px] uppercase font-bold text-indigo-950 tracking-wider mb-2">
                                          Активные привязки по категориям (Смена по умолчанию):
                                        </div>
                                        <div className="flex flex-wrap gap-1.5 justify-start">
                                          {Object.entries(categoryMappings).map(([unmatchedCat, targetLabel]) => (
                                            <div key={unmatchedCat} className="flex items-center gap-1.5 bg-white border border-indigo-150 px-2.5 py-1 rounded-lg text-[11px] font-sans">
                                              <span className="font-semibold text-gray-750 truncate max-w-[150px]">{unmatchedCat}</span>
                                              <span className="text-gray-400">→</span>
                                              <span className="font-bold text-indigo-700">{targetLabel}</span>
                                              <button
                                                onClick={(e) => {
                                                  e.stopPropagation();
                                                  const d = { ...categoryMappings };
                                                  delete d[unmatchedCat];
                                                  setCategoryMappings(d);
                                                }}
                                                className="text-gray-400 hover:text-rose-500 ml-1 font-bold text-xs cursor-pointer px-0.5"
                                                title="Сбросить привязку"
                                              >
                                                &times;
                                              </button>
                                            </div>
                                          ))}
                                        </div>
                                      </div>
                                    )}

                                    {Object.keys(transactionMappings).length > 0 && (
                                      <div>
                                        <div className="text-[10px] uppercase font-bold text-emerald-950 tracking-wider mb-1.5 mt-2">
                                          Активные перенаправления транзакций:
                                        </div>
                                        <div className="flex flex-wrap gap-1.5 justify-start">
                                          {Object.entries(transactionMappings).map(([tid, targetLabel]) => {
                                            const transaction = rawRecords.find(r => r.id === tid);
                                            const desc = transaction 
                                              ? `${transaction.sum.toLocaleString('ru-RU')} ₽ (${transaction.comment || transaction.category})` 
                                              : tid;
                                            return (
                                              <div key={tid} className="flex items-center gap-1.5 bg-white border border-emerald-150 px-2.5 py-1 rounded-lg text-[11px] font-sans">
                                                <span className="font-semibold text-gray-750 truncate max-w-[180px]">{desc}</span>
                                                <span className="text-gray-400">→</span>
                                                <span className="font-bold text-emerald-700">{targetLabel}</span>
                                                <button
                                                  onClick={(e) => {
                                                    e.stopPropagation();
                                                    const d = { ...transactionMappings };
                                                    delete d[tid];
                                                    setTransactionMappings(d);
                                                  }}
                                                  className="text-gray-400 hover:text-rose-500 ml-1 font-bold text-xs cursor-pointer px-0.5"
                                                  title="Сбросить перенаправление"
                                                >
                                                  &times;
                                                </button>
                                              </div>
                                            );
                                          })}
                                        </div>
                                      </div>
                                    )}
                                  </div>
                                </td>
                              </tr>
                            )}

                            {extraRows.map((row, idx) => (
                              <React.Fragment key={`extra_${idx}`}>
                                <tr 
                                  className="bg-white hover:bg-amber-50/20 transition-colors h-11 cursor-pointer border-b border-gray-100 table-row text-left"
                                  onClick={() => setExpandedExtraRows(prev => ({ ...prev, [idx]: !prev[idx] }))}
                                >
                                  <td className="p-3 text-center font-mono text-[11px] text-amber-400/80 border-r border-amber-50">
                                    {idx + 1 + financialReport.length}
                                  </td>
                                  <td className="p-3 pl-4 font-sans text-gray-750 font-medium flex items-center gap-2">
                                    <ChevronDown className={`w-3.5 h-3.5 text-amber-500 transition-transform ${expandedExtraRows[idx] ? 'rotate-180' : ''}`} />
                                    <span>{row.label}</span>
                                    {row.transactions && (
                                      <span className="text-[10px] text-gray-400 font-normal font-mono">
                                        ({row.transactions.length})
                                      </span>
                                    )}
                                  </td>
                                  <td className="p-3">
                                    <div className="flex items-center gap-3" onClick={(e) => e.stopPropagation()}>
                                      <span className="px-2 py-0.5 rounded text-[9px] uppercase font-bold bg-amber-50 text-amber-700 border border-amber-100 whitespace-nowrap">
                                        НЕТ В ТЕМПЛЕЙТЕ
                                      </span>
                                      <div className="inline-flex items-center gap-1.5">
                                        <span className="text-[10px] text-gray-400 font-medium whitespace-nowrap">Связать всё:</span>
                                        <select
                                          value={categoryMappings[row.label] || ''}
                                          onChange={(e) => {
                                            const val = e.target.value;
                                            const updated = { ...categoryMappings };
                                            if (val) {
                                              updated[row.label] = val;
                                              setNotification({
                                                status: 'success',
                                                message: `Категория "${row.label}" перенаправлена на показатель "${val}". Финансовые показатели пересчитаны мгновенно!`
                                              });
                                            } else {
                                              delete updated[row.label];
                                            }
                                            setCategoryMappings(updated);
                                          }}
                                          className="bg-gray-55 hover:bg-gray-100 border border-gray-250 text-gray-700 rounded-lg text-[10.5px] px-2 py-1 font-sans focus:outline-none cursor-pointer max-w-[130px] sm:max-w-[180px] truncate"
                                        >
                                          <option value="">-- выбрать --</option>
                                          {FINANCIAL_TEMPLATE.filter(r => r.type === 'data').sort((a,b) => a.label.localeCompare(b.label)).map((tRow) => (
                                            <option key={tRow.index} value={tRow.label}>
                                              {tRow.label}
                                            </option>
                                          ))}
                                        </select>
                                      </div>
                                    </div>
                                  </td>
                                  <td className="p-3 text-right pr-6">
                                    <span className="font-mono font-semibold text-gray-900 text-[13px]">
                                      {row.value.toLocaleString('ru-RU')} ₽
                                    </span>
                                  </td>
                                </tr>
                                {expandedExtraRows[idx] && row.transactions && (
                                  <tr className="bg-amber-50/10">
                                    <td colSpan={4} className="p-4 pl-12 bg-gray-55/65 border-b border-gray-100">
                                      <div className="flex items-center justify-between pb-2 mb-2.5 border-b border-amber-200/60">
                                        <span className="text-xs font-bold text-amber-950">
                                          Операции по нераспределенной строке «{row.label}» ({row.transactions.length} шт.):
                                        </span>
                                        <button
                                          onClick={(e) => {
                                            e.stopPropagation();
                                            setExpandedExtraRows(prev => ({ ...prev, [idx]: false }));
                                          }}
                                          className="px-2.5 py-1 text-[11px] font-semibold text-amber-850 hover:text-amber-950 bg-white hover:bg-amber-50 rounded-lg border border-amber-200 transition-all flex items-center gap-1 cursor-pointer shadow-2xs"
                                          title="Свернуть детализацию"
                                        >
                                          <ChevronUp className="w-3.5 h-3.5" />
                                          <span>Свернуть детализацию</span>
                                        </button>
                                      </div>
                                      <div className="overflow-x-auto rounded-xl border border-gray-200/60 shadow-sm bg-white">
                                        <table className="min-w-full divide-y divide-gray-150 text-left text-[11px]">
                                          <thead className="bg-gray-50/85 font-semibold text-gray-600 uppercase tracking-wider text-[9px] select-none">
                                            <tr>
                                              <th className="p-2.5 pl-4 w-28">Дата</th>
                                              <th className="p-2.5 w-48">Контрагент</th>
                                              <th className="p-2.5 font-bold text-indigo-950">Назначение</th>
                                              <th className="p-2.5">Комментарий</th>
                                              <th className="p-2.5 text-right w-36">Сумма</th>
                                              <th className="p-2.5 text-center pr-4 w-48">Определить показатель</th>
                                            </tr>
                                          </thead>
                                          <tbody className="divide-y divide-gray-150 font-sans text-gray-650">
                                            {row.transactions.map((t, tIdx) => (
                                              <tr key={tIdx} className="hover:bg-gray-50/40 transition-colors">
                                                <td className="p-2.5 pl-4 font-mono text-[10.5px]">
                                                  {t.date ? new Date(t.date).toLocaleDateString('ru-RU') : t.dateStr || '—'}
                                                </td>
                                                <td className="p-2.5 font-medium text-gray-800">
                                                   {t.counterparty || <span className="text-gray-400 italic">Не указан</span>}
                                                 </td>
                                                 <td className="p-2.5 text-gray-700 max-w-[200px] truncate" title={t.category}>
                                                   {t.category || <span className="text-gray-300">—</span>}
                                                 </td>
                                                 <td className="p-2.5 text-gray-500 max-w-[200px] truncate" title={t.comment}>
                                                  {t.comment || <span className="text-gray-300">—</span>}
                                                </td>
                                                <td className="p-2.5 text-right font-mono font-semibold text-gray-800">
                                                  {t.sum.toLocaleString('ru-RU', { minimumFractionDigits: 2, maximumFractionDigits: 2 })} ₽
                                                </td>
                                                <td className="p-2 text-center pr-4">
                                                  <select
                                                    value={transactionMappings[t.id] || ''}
                                                    onChange={(e) => {
                                                      const val = e.target.value;
                                                      const updated = { ...transactionMappings };
                                                      if (val) {
                                                        updated[t.id] = val;
                                                        setNotification({
                                                          status: 'success',
                                                          message: `Транзакция на ${t.sum} ₽ перенаправлена на "${val}". Финансовые показатели пересчитаны мгновенно!`
                                                        });
                                                      } else {
                                                        delete updated[t.id];
                                                      }
                                                      setTransactionMappings(updated);
                                                    }}
                                                    className="bg-gray-50 hover:bg-gray-100 border border-gray-250 text-gray-600 rounded-lg px-2 py-1 text-[10px] font-sans focus:outline-none cursor-pointer max-w-[150px] truncate"
                                                  >
                                                    <option value="">-- авто --</option>
                                                    {FINANCIAL_TEMPLATE.filter(r => r.type === 'data').sort((a,b) => a.label.localeCompare(b.label)).map((tRow) => (
                                                      <option key={tRow.index} value={tRow.label}>
                                                        {tRow.label}
                                                      </option>
                                                    ))}
                                                  </select>
                                                </td>
                                              </tr>
                                            ))}
                                          </tbody>
                                        </table>
                                      </div>
                                      <div className="mt-2.5 flex justify-end">
                                        <button
                                          onClick={(e) => {
                                            e.stopPropagation();
                                            setExpandedExtraRows(prev => ({ ...prev, [idx]: false }));
                                          }}
                                          className="px-2.5 py-1 text-[11px] font-semibold text-amber-850 hover:text-amber-950 bg-white hover:bg-amber-50 rounded-lg border border-amber-200 transition-all flex items-center gap-1 cursor-pointer shadow-2xs"
                                        >
                                          <ChevronUp className="w-3.5 h-3.5 text-amber-500" />
                                          <span>Свернуть</span>
                                        </button>
                                      </div>
                                    </td>
                                  </tr>
                                )}
                              </React.Fragment>
                            ))}
                          </>
                        )}
                      </tbody>
                    </table>
                  </div>
                </div>
              </div>
            )}

            {/* Revenue Details by Service/Product Tab Content */}
            {activeTab === 'revenue-details' && (
              <div className="space-y-6">
                
                {/* Header Information Banner */}
                <div className="bg-gradient-to-r from-indigo-900 to-indigo-800 rounded-3xl p-6 text-white shadow-md">
                  <div className="flex flex-col md:flex-row items-start md:items-center justify-between gap-4">
                    <div className="space-y-1">
                      <div className="flex items-center gap-2">
                        <span className="p-2 bg-white/10 rounded-xl">
                          <Layers className="w-5 h-5 text-indigo-200" />
                        </span>
                        <h3 className="font-display font-bold text-lg text-white">
                          Детализация доходной части (графа «Услуга/Товар»)
                        </h3>
                      </div>
                      <p className="text-xs text-indigo-200 max-w-2xl leading-relaxed">
                        Строки отчета, по которым записи в графе «Услуга/Товар» совпадают, автоматически сгруппированы и просуммированы. 
                        Вы можете просмотреть как итоговую сумму по каждой услуге или товару, так и раскрыть все связанные транзакции.
                      </p>
                    </div>

                    <button
                      onClick={() => downloadRevenueSpreadsheet(revenueDetails || [], totalRevenue)}
                      className="px-4 py-2.5 bg-white text-indigo-900 hover:bg-indigo-50 font-bold rounded-xl text-xs flex items-center gap-2 shadow-sm transition-all cursor-pointer whitespace-nowrap self-stretch md:self-auto justify-center"
                      title="Выгрузить отдельный Excel файл с детализацией по Услугам/Товарам"
                    >
                      <Download className="w-4 h-4 text-indigo-700" />
                      <span>Экспорт в Excel (XLSX)</span>
                    </button>
                  </div>

                  {/* Summary KPI Counters inside Banner */}
                  <div className="grid grid-cols-2 sm:grid-cols-4 gap-3 mt-6 pt-5 border-t border-indigo-700/60">
                    <div className="bg-white/10 rounded-2xl p-3">
                      <div className="text-[10px] uppercase font-bold text-indigo-200 tracking-wider">Всего доходов</div>
                      <div className="text-xl font-bold font-mono mt-1 text-white">
                        {revenueStats.totalSum.toLocaleString('ru-RU')} ₽
                      </div>
                    </div>
                    <div className="bg-white/10 rounded-2xl p-3">
                      <div className="text-[10px] uppercase font-bold text-indigo-200 tracking-wider">Позиций (Услуг/Товаров)</div>
                      <div className="text-xl font-bold font-mono mt-1 text-white">
                        {filteredRevenueDetails.length} <span className="text-xs text-indigo-300 font-normal">из {revenueDetails?.length || 0}</span>
                      </div>
                    </div>
                    <div className="bg-white/10 rounded-2xl p-3">
                      <div className="text-[10px] uppercase font-bold text-indigo-200 tracking-wider">Всего операций</div>
                      <div className="text-xl font-bold font-mono mt-1 text-white">
                        {revenueStats.totalCount} <span className="text-xs text-indigo-300 font-normal">строк</span>
                      </div>
                    </div>
                    <div className="bg-white/10 rounded-2xl p-3">
                      <div className="text-[10px] uppercase font-bold text-indigo-200 tracking-wider">Услуги / Товары</div>
                      <div className="text-xs font-mono mt-1 text-white flex flex-col gap-0.5">
                        <span className="truncate">Услуги: {revenueStats.servicesSum.toLocaleString('ru-RU')} ₽</span>
                        <span className="truncate">Товары: {revenueStats.goodsSum.toLocaleString('ru-RU')} ₽</span>
                      </div>
                    </div>
                  </div>
                </div>

                {/* Search & Category Filter Toolbar */}
                <div className="bg-white rounded-2xl border border-gray-150 p-4 shadow-sm flex flex-col sm:flex-row gap-4 items-center justify-between">
                  {/* Search query input */}
                  <div className="flex items-center gap-2 bg-gray-50 px-3.5 py-2 rounded-xl border border-gray-150 w-full sm:max-w-md">
                    <Search className="w-4 h-4 text-gray-400" />
                    <input
                      type="text"
                      value={revenueSearch}
                      onChange={(e) => setRevenueSearch(e.target.value)}
                      placeholder="Поиск по названию услуги, товара или категории..."
                      className="bg-transparent text-xs w-full outline-none text-gray-700 border-none inline-block focus:ring-0"
                    />
                    {revenueSearch && (
                      <button onClick={() => setRevenueSearch('')} className="text-gray-400 hover:text-gray-600 text-xs select-none font-medium">
                        Очистить
                      </button>
                    )}
                  </div>

                  {/* Filter category pills and bulk collapse/expand */}
                  <div className="flex items-center gap-2.5 w-full sm:w-auto flex-wrap">
                    {/* Bulk Expand / Collapse Details Buttons */}
                    <div className="flex items-center gap-1 bg-gray-50 p-1 rounded-xl border border-gray-150">
                      <button
                        onClick={() => toggleAllRevenueDetails(true)}
                        className="flex items-center gap-1 px-2.5 py-1 rounded-lg text-xs font-semibold text-gray-700 hover:text-indigo-700 hover:bg-white transition-all cursor-pointer shadow-2xs"
                        title="Развернуть детализацию всех услуг и товаров"
                      >
                        <ChevronsUpDown className="w-3.5 h-3.5 text-indigo-600" />
                        <span>Развернуть все</span>
                      </button>
                      <button
                        onClick={() => toggleAllRevenueDetails(false)}
                        className="flex items-center gap-1 px-2.5 py-1 rounded-lg text-xs font-semibold text-gray-700 hover:text-indigo-700 hover:bg-white transition-all cursor-pointer shadow-2xs"
                        title="Свернуть все открытые детализации"
                      >
                        <ChevronsDownUp className="w-3.5 h-3.5 text-gray-500" />
                        <span>Свернуть все</span>
                      </button>
                    </div>

                    <div className="flex items-center gap-1.5 overflow-x-auto pb-1 sm:pb-0">
                      <button
                        onClick={() => setRevenueCategoryFilter('all')}
                        className={`px-3 py-1.5 rounded-xl text-xs font-semibold whitespace-nowrap transition-all cursor-pointer ${
                          revenueCategoryFilter === 'all'
                            ? 'bg-indigo-600 text-white shadow-sm'
                            : 'bg-gray-50 text-gray-600 hover:bg-gray-100'
                        }`}
                      >
                        Все ({revenueDetails?.length || 0})
                      </button>
                      <button
                        onClick={() => setRevenueCategoryFilter('services')}
                        className={`px-3 py-1.5 rounded-xl text-xs font-semibold whitespace-nowrap transition-all cursor-pointer ${
                          revenueCategoryFilter === 'services'
                            ? 'bg-indigo-600 text-white shadow-sm'
                            : 'bg-gray-50 text-gray-600 hover:bg-gray-100'
                        }`}
                      >
                        Только услуги
                      </button>
                      <button
                        onClick={() => setRevenueCategoryFilter('goods')}
                        className={`px-3 py-1.5 rounded-xl text-xs font-semibold whitespace-nowrap transition-all cursor-pointer ${
                          revenueCategoryFilter === 'goods'
                            ? 'bg-indigo-600 text-white shadow-sm'
                            : 'bg-gray-50 text-gray-600 hover:bg-gray-100'
                        }`}
                      >
                        Только товары
                      </button>
                      <button
                        onClick={() => setRevenueCategoryFilter('other')}
                        className={`px-3 py-1.5 rounded-xl text-xs font-semibold whitespace-nowrap transition-all cursor-pointer ${
                          revenueCategoryFilter === 'other'
                            ? 'bg-indigo-600 text-white shadow-sm'
                            : 'bg-gray-50 text-gray-600 hover:bg-gray-100'
                        }`}
                      >
                        Прочие / Сертификаты
                      </button>
                    </div>
                  </div>
                </div>

                {/* Main Detailing Table */}
                <div className="bg-white rounded-3xl border border-gray-150 shadow-sm overflow-hidden">
                  <div className="p-4 bg-gray-50/70 border-b border-gray-150 flex items-center justify-between select-none">
                    <div className="flex items-center gap-2">
                      <Table className="w-4 h-4 text-indigo-600" />
                      <h3 className="font-display font-bold text-xs uppercase tracking-wider text-gray-700">
                        Сгруппированные записи по графе «Услуга/Товар»
                      </h3>
                    </div>
                    <span className="text-[11px] text-indigo-700 bg-indigo-50 border border-indigo-150 px-2.5 py-0.5 rounded-md font-semibold">
                      Показано строк: {filteredRevenueDetails.length}
                    </span>
                  </div>

                  {filteredRevenueDetails.length === 0 ? (
                    <div className="p-12 text-center text-gray-400 text-xs">
                      {revenueSearch 
                        ? `По запросу «${revenueSearch}» ничего не найдено.` 
                        : 'Нет данных о доходах или в загруженном файле не заполнена графа «Услуга/Товар».'}
                    </div>
                  ) : (
                    <div className="overflow-x-auto">
                      <table className="w-full text-left text-xs border-collapse">
                        <thead>
                          <tr className="bg-gray-50 text-gray-600 font-bold select-none h-11 border-b border-gray-150">
                            <th className="p-3.5 pl-6 w-16 text-center font-mono text-[10px] text-gray-400">№</th>
                            <th className="p-3.5">Наименование из графы «Услуга/Товар»</th>
                            <th className="p-3.5 w-48">Категория отчета</th>
                            <th className="p-3.5 text-center w-36">Кол-во операций</th>
                            <th className="p-3.5 text-center w-32">Доля в выручке</th>
                            <th className="p-3.5 text-right pr-6 w-44">Итоговая сумма</th>
                          </tr>
                        </thead>
                        <tbody className="divide-y divide-gray-100 font-sans">
                          {filteredRevenueDetails.map((item, idx) => {
                            const isExpanded = !!expandedRevenueItems[item.name];
                            const share = totalRevenue > 0 ? (item.sum / totalRevenue) * 100 : 0;
                            const isService = item.category.toLowerCase().includes('услуг');
                            const isGoods = item.category.toLowerCase().includes('товар');

                            return (
                              <React.Fragment key={`rev_group_${idx}`}>
                                <tr
                                  className="hover:bg-indigo-50/20 transition-colors h-12 cursor-pointer"
                                  onClick={() => setExpandedRevenueItems(prev => ({ ...prev, [item.name]: !prev[item.name] }))}
                                >
                                  <td className="p-3 text-center font-mono text-[11px] text-gray-400 border-r border-gray-50 pl-6">
                                    {idx + 1}
                                  </td>
                                  <td className="p-3 font-medium text-gray-850">
                                    <div className="flex items-center gap-2 flex-wrap">
                                      <div className="flex items-center gap-2 shrink-0">
                                        <ChevronDown className={`w-3.5 h-3.5 text-indigo-500 transition-transform shrink-0 ${isExpanded ? 'rotate-180' : ''}`} />
                                        <span className="font-semibold text-gray-900">{item.name}</span>
                                      </div>

                                      {item.hall && (
                                        <span className={`px-2 py-0.5 rounded-md text-[10px] font-bold border ${
                                          item.hall === 'Прочие услуги'
                                            ? 'bg-amber-50 text-amber-900 border-amber-250'
                                            : `${HALL_METADATA[item.hall as HallName]?.badgeBg || 'bg-indigo-50'} ${HALL_METADATA[item.hall as HallName]?.badgeText || 'text-indigo-800'} ${HALL_METADATA[item.hall as HallName]?.badgeBorder || 'border-indigo-200'}`
                                        }`}>
                                          {HALL_METADATA[item.hall as HallName]?.shortLabel || item.hall}
                                        </span>
                                      )}

                                      {item.hall === 'Прочие услуги' && (
                                        <div className="inline-flex items-center gap-1" onClick={(e) => e.stopPropagation()}>
                                          <select
                                            value=""
                                            onChange={(e) => {
                                              if (e.target.value) {
                                                handleAssignServiceToHall(item.name, e.target.value as HallName);
                                              }
                                            }}
                                            className="text-[10px] font-bold bg-amber-50 hover:bg-amber-100 text-amber-900 border border-amber-300 rounded-lg px-2 py-0.5 outline-none cursor-pointer shadow-2xs"
                                            title="Отнести эту позицию к одному из залов салона с сохранением выбора"
                                          >
                                            <option value="">➡️ Отнести к залу...</option>
                                            {HALL_NAMES.filter(h => h !== 'Прочие услуги').map(h => (
                                              <option key={`tab3_opt_${item.name}_${h}`} value={h}>
                                                {h}
                                              </option>
                                            ))}
                                          </select>
                                        </div>
                                      )}

                                      {serviceHallOverrides[item.name] && (
                                        <div 
                                          className="inline-flex items-center gap-1 px-1.5 py-0.5 rounded-md text-[9.5px] font-bold bg-indigo-50 text-indigo-800 border border-indigo-200" 
                                          onClick={(e) => e.stopPropagation()}
                                          title={`Назначено пользователем в «${serviceHallOverrides[item.name]}». Сохранено в браузере.`}
                                        >
                                          <span>✏️ Вручную</span>
                                          <button
                                            type="button"
                                            onClick={() => handleResetServiceHall(item.name)}
                                            className="text-indigo-400 hover:text-rose-600 transition-colors ml-0.5 cursor-pointer font-bold"
                                            title="Сбросить назначение"
                                          >
                                            ✕
                                          </button>
                                        </div>
                                      )}
                                    </div>
                                  </td>
                                  <td className="p-3">
                                    <span className={`px-2 py-0.5 rounded text-[10px] font-semibold border ${
                                      isService 
                                        ? 'bg-blue-50 text-blue-700 border-blue-150'
                                        : isGoods
                                        ? 'bg-amber-50 text-amber-700 border-amber-150'
                                        : 'bg-purple-50 text-purple-700 border-purple-150'
                                    }`}>
                                      {item.category}
                                    </span>
                                  </td>
                                  <td className="p-3 text-center">
                                    <span className="font-mono font-bold text-indigo-700 bg-indigo-50 border border-indigo-150 px-2 py-0.5 rounded text-[11px]">
                                      {item.count} {item.count === 1 ? 'операция' : item.count < 5 ? 'операции' : 'операций'}
                                    </span>
                                  </td>
                                  <td className="p-3 text-center font-mono text-gray-600 font-semibold text-[11px]">
                                    {share.toFixed(1)}%
                                  </td>
                                  <td className="p-3 text-right pr-6 font-mono font-bold text-[13px] text-gray-900">
                                    {item.sum.toLocaleString('ru-RU', { minimumFractionDigits: 2, maximumFractionDigits: 2 })} ₽
                                  </td>
                                </tr>

                                {/* Expanded Drawer: All underlying transactions for this grouped service/product */}
                                {isExpanded && item.transactions && item.transactions.length > 0 && (
                                  <tr className="bg-indigo-50/15">
                                    <td colSpan={6} className="p-4 pl-12 bg-indigo-50/25 border-b border-gray-150">
                                      <div className="bg-white rounded-2xl border border-indigo-150 shadow-sm overflow-hidden">
                                        <div className="p-3 bg-indigo-50/60 border-b border-indigo-100 flex items-center justify-between text-xs font-semibold text-indigo-950">
                                          <div className="flex items-center gap-2">
                                            <span>Все входящие операции по позиции «{item.name}» ({item.transactions.length} шт.):</span>
                                            <span className="font-mono text-indigo-700 font-bold">
                                              Сумма: {item.sum.toLocaleString('ru-RU', { minimumFractionDigits: 2, maximumFractionDigits: 2 })} ₽
                                            </span>
                                          </div>
                                          <button
                                            onClick={(e) => {
                                              e.stopPropagation();
                                              setExpandedRevenueItems(prev => ({ ...prev, [item.name]: false }));
                                            }}
                                            className="px-2.5 py-1 text-[11px] font-semibold text-indigo-700 hover:text-indigo-900 bg-white hover:bg-indigo-50 rounded-lg border border-indigo-200 transition-all flex items-center gap-1 cursor-pointer shadow-2xs"
                                            title="Свернуть детализацию этой позиции"
                                          >
                                            <ChevronUp className="w-3.5 h-3.5" />
                                            <span>Свернуть</span>
                                          </button>
                                        </div>
                                        <div className="overflow-x-auto">
                                          <table className="w-full text-left text-[11px] divide-y divide-gray-150">
                                            <thead className="bg-gray-50/80 font-bold text-gray-600 uppercase text-[9px] select-none">
                                              <tr>
                                                <th className="p-2.5 pl-4 w-28">Дата</th>
                                                <th className="p-2.5 w-44">Контрагент / Клиент</th>
                                                <th className="p-2.5">Назначение платежа</th>
                                                <th className="p-2.5">Комментарий</th>
                                                <th className="p-2.5 text-right pr-4 w-36">Сумма операции</th>
                                              </tr>
                                            </thead>
                                            <tbody className="divide-y divide-gray-100 font-sans text-gray-700">
                                              {item.transactions.map((t, tIdx) => (
                                                <tr key={`tx_${idx}_${tIdx}`} className="hover:bg-gray-50/60 transition-colors">
                                                  <td className="p-2.5 pl-4 font-mono text-gray-500">
                                                    {t.date ? new Date(t.date).toLocaleDateString('ru-RU') : t.dateStr || '—'}
                                                  </td>
                                                  <td className="p-2.5 font-medium text-gray-850">
                                                    {t.counterparty || <span className="text-gray-400 italic">Не указан</span>}
                                                  </td>
                                                  <td className="p-2.5 text-gray-700">
                                                    {t.category || '—'}
                                                  </td>
                                                  <td className="p-2.5 text-gray-500 italic max-w-xs truncate" title={t.comment}>
                                                    {t.comment || '—'}
                                                  </td>
                                                  <td className="p-2.5 text-right pr-4 font-mono font-bold text-gray-900">
                                                    {t.sum.toLocaleString('ru-RU', { minimumFractionDigits: 2, maximumFractionDigits: 2 })} ₽
                                                  </td>
                                                </tr>
                                              ))}
                                            </tbody>
                                          </table>
                                        </div>
                                        {/* Footer Collapse Button */}
                                        <div className="p-2 bg-gray-50/60 border-t border-gray-100 flex justify-end">
                                          <button
                                            onClick={(e) => {
                                              e.stopPropagation();
                                              setExpandedRevenueItems(prev => ({ ...prev, [item.name]: false }));
                                            }}
                                            className="px-2.5 py-1 text-[11px] font-semibold text-gray-600 hover:text-indigo-700 bg-white hover:bg-gray-100 rounded-lg border border-gray-200 transition-all flex items-center gap-1 cursor-pointer shadow-2xs"
                                          >
                                            <ChevronUp className="w-3.5 h-3.5 text-gray-400" />
                                            <span>Свернуть детализацию «{item.name}»</span>
                                          </button>
                                        </div>
                                      </div>
                                    </td>
                                  </tr>
                                )}
                              </React.Fragment>
                            );
                          })}
                        </tbody>
                        {/* Table Footer with totals */}
                        <tfoot className="bg-gray-50 font-bold border-t border-gray-200">
                          <tr className="h-11">
                            <td className="p-3 pl-6 text-center text-gray-400" colSpan={3}>
                              Итого по сгруппированным строкам
                            </td>
                            <td className="p-3 text-center font-mono text-indigo-700">
                              {filteredRevenueDetails.reduce((sum, d) => sum + d.count, 0)} оп.
                            </td>
                            <td className="p-3 text-center font-mono text-gray-600">
                              {totalRevenue > 0 
                                ? `${((filteredRevenueDetails.reduce((sum, d) => sum + d.sum, 0) / totalRevenue) * 100).toFixed(1)}%`
                                : '—'}
                            </td>
                            <td className="p-3 text-right pr-6 font-mono text-gray-900 text-[13px]">
                              {filteredRevenueDetails.reduce((sum, d) => sum + d.sum, 0).toLocaleString('ru-RU', { minimumFractionDigits: 2, maximumFractionDigits: 2 })} ₽
                            </td>
                          </tr>
                        </tfoot>
                      </table>
                    </div>
                  )}
                </div>
              </div>
            )}

            {/* Salary Payments tab content */}
            {activeTab === 'salary' && (
              <div className="space-y-6">
                
                {/* Visual Indicator rules callout */}
                <div className="bg-white rounded-3xl border border-gray-150 p-6 shadow-sm">
                  <h3 className="font-display font-semibold text-md text-gray-900 mb-3 flex items-center gap-2">
                    <Calendar className="w-4.5 h-4.5 text-indigo-500" />
                    Алгоритм проверки зарплатной ведомости
                  </h3>
                  <div className="grid grid-cols-1 md:grid-cols-2 gap-4 text-xs leading-relaxed text-gray-600">
                    <div className="p-3.5 rounded-xl bg-amber-50/40 border border-amber-150/20">
                      <strong className="text-amber-900 block mb-1">Выплачено за прошлый период:</strong>
                      Попадают записи, в назначении которых есть «Зарплата»/«Заработная», 
                      в комментарии содержится ключевое слово (<code className="font-mono bg-white px-1">зп</code>, <code className="font-mono bg-white px-1">з.п.</code>, <code className="font-mono bg-white px-1">з п</code>, <code className="font-mono bg-white px-1">зарплата</code>, <code className="font-mono bg-white px-1">заработная</code>) 
                      и дата операции меньше заданной даты отсечения (<strong className="font-mono text-indigo-900">{cutoffDateStr}</strong>).
                    </div>
                    <div className="p-3.5 rounded-xl bg-emerald-50/40 border border-emerald-150/20">
                      <strong className="text-emerald-900 block mb-1">Выплачено за отчетный период:</strong>
                      Остальные выплаты, включая: 
                      1) записи без ключевых слов в комментарии; 
                      2) записи с датой позже или равной дате отсечения.
                    </div>
                  </div>
                </div>

                {/* Salary report table */}
                <div className="bg-white rounded-3xl border border-gray-150 shadow-sm overflow-hidden">
                  <div className="p-4 bg-gray-50/50 border-b border-gray-100 flex items-center justify-between">
                    <h3 className="font-display font-semibold text-xs uppercase tracking-wider text-gray-500">
                      Зарплатная таблица (Группировка по сотрудникам)
                    </h3>
                    <span className="text-[11px] text-indigo-700 bg-indigo-50 px-2.5 py-1 rounded-md font-medium border border-indigo-100/50">
                      Сотрудников: {salaryEmployees.length}
                    </span>
                  </div>

                  {salaryEmployees.length === 0 ? (
                    <div className="p-12 text-center text-gray-400">
                      Нет сотрудников, подходящих под критерии начислений зарплат (Проверьте, содержит ли "Назначение" слова "Зарплата" или "Заработная").
                    </div>
                  ) : (
                    <div className="overflow-x-auto">
                      <table className="w-full text-left text-xs border-collapse">
                        <thead>
                          <tr className="bg-gray-50 text-gray-500 font-bold select-none h-11 border-b border-gray-100">
                            <th className="p-3.5 pl-6">ФИО (Контрагент)</th>
                            <th className="p-3.5 text-right">Выплачено за прошлый период</th>
                            <th className="p-3.5 text-right">Выплачено за отчётный период</th>
                            <th className="p-3.5 text-right pr-6">Итого выплат</th>
                          </tr>
                        </thead>
                        <tbody className="divide-y divide-gray-100 font-sans">
                          {salaryEmployees.map((emp, idx) => (
                            <tr key={idx} className="hover:bg-gray-50/50 transition-colors h-11">
                              <td className="p-3 pl-6 font-medium text-gray-800">
                                {emp.fio}
                              </td>
                              <td className="p-3 text-right font-mono font-semibold text-amber-700">
                                {emp.pastPeriodSum > 0 ? `${emp.pastPeriodSum.toLocaleString('ru-RU')} ₽` : '—'}
                              </td>
                              <td className="p-3 text-right font-mono font-semibold text-emerald-700">
                                {emp.reportPeriodSum > 0 ? `${emp.reportPeriodSum.toLocaleString('ru-RU')} ₽` : '—'}
                              </td>
                              <td className="p-3 text-right font-mono font-bold text-gray-900 pr-6">
                                {emp.totalSum.toLocaleString('ru-RU')} ₽
                              </td>
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    </div>
                  )}
                </div>
              </div>
            )}

            {/* Raw records list tab */}
            {activeTab === 'raw' && (
              <div className="bg-white rounded-3xl border border-gray-150 shadow-sm overflow-hidden">
                <div className="p-4 bg-gray-50/50 border-b border-gray-100 flex items-center justify-between select-none">
                  <h3 className="font-display font-semibold text-xs uppercase tracking-wider text-gray-500">
                    Распознанные транзакции из файла
                  </h3>
                  <span className="text-[11px] text-indigo-700 bg-indigo-50 border border-indigo-100/50 px-2.5 py-1 rounded-md font-medium">
                    Всего строк: {rawRecords.length}
                  </span>
                </div>
                
                <div className="overflow-x-auto">
                  <table className="w-full text-left text-xs border-collapse">
                    <thead>
                      <tr className="bg-gray-50 text-gray-400 font-medium h-10 border-b border-gray-100 font-mono text-[10px]">
                        <th className="p-3 text-center pl-6 w-16">№</th>
                        <th className="p-3">Дата</th>
                        <th className="p-3">Назначение</th>
                        <th className="p-3">Контрагент</th>
                        <th className="p-3">Комментарий</th>
                        <th className="p-3 text-right pr-6">Сумма (Руб)</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-gray-100 font-sans">
                      {rawRecords.map((r) => (
                        <tr key={r.id} className="hover:bg-gray-50/40 transition-colors text-gray-600">
                          <td className="p-3 text-center font-mono text-[11px] text-gray-400 border-r border-gray-50 pl-6">
                            {r.rawIndex}
                          </td>
                          <td className="p-3 font-mono text-gray-500">
                            {r.dateStr || '—'}
                          </td>
                          <td className="p-3 text-gray-800 font-medium">
                            {r.category}
                          </td>
                          <td className="p-3 text-gray-700">
                            {r.counterparty || '—'}
                          </td>
                          <td className="p-3 text-gray-500 italic max-w-xs truncate">
                            {r.comment || '—'}
                          </td>
                          <td className="p-3 text-right pr-6 font-mono font-medium text-gray-900">
                            {r.sum.toLocaleString('ru-RU')} ₽
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </div>
            )}

            {/* Embed test verification sandbox always at the of loaded state as well to let them easily test and verify */}
            <div className="pt-6 border-t border-gray-100">
              <TestPlayground />
            </div>

          </div>
        )}
      </div>
    )}

    {/* ========================================================================= */}
    {/* TAB 2: ОТЧЕТ ПО УСЛУГАМ С РАСПРЕДЕЛЕНИЕМ ЗАТРАТ                          */}
    {/* ========================================================================= */}
    {primaryTab === 'services' && (
      <div className="space-y-6">
        
        {/* Header Information Banner */}
        <div className="bg-gradient-to-r from-indigo-900 via-indigo-850 to-indigo-800 rounded-3xl p-6 text-white shadow-md">
          <div className="flex flex-col md:flex-row items-start md:items-center justify-between gap-4">
            <div className="space-y-1.5">
              <div className="flex items-center gap-2">
                <span className="p-2 bg-white/10 rounded-xl">
                  <Scissors className="w-5 h-5 text-indigo-200" />
                </span>
                <h2 className="font-display font-extrabold text-xl text-white">
                  Отчет по услугам и распределению затрат (Вкладка 2)
                </h2>
              </div>
              <p className="text-xs text-indigo-200 max-w-3xl leading-relaxed">
                Отчет формируется на основе загружаемого файла по услугам и финансового отчета из 1-й вкладки:
                графа <strong>«Услуга»</strong> группирует одинаковые записи, графа <strong>«Скидка»</strong> суммирует скидки, 
                графа <strong>«Сумма выручки»</strong> складывает «Баллы» + «Абонементы» + «Сертификаты» + «Оплачено со счетов клиентов» + «Оплачено деньгами». 
                Графы <strong>«Переменные затраты»</strong> и <strong>«Постоянные затраты»</strong> заполняются путем распределения затрат из финансового отчета пропорционально выручке.
              </p>
            </div>

            <div className="flex gap-2 flex-wrap self-stretch md:self-auto justify-end">
              {servicesRecords.length === 0 && (
                <button
                  onClick={loadDemoServicesData}
                  className="px-4 py-2.5 bg-indigo-700/80 hover:bg-indigo-700 text-white font-semibold rounded-xl text-xs flex items-center gap-1.5 border border-indigo-500/50 transition-all cursor-pointer"
                  title="Загрузить тестовые услуги салона красоты"
                >
                  <Sparkles className="w-3.5 h-3.5 text-amber-300" />
                  <span>Загрузить Демо Услуг</span>
                </button>
              )}
              {servicesRecords.length > 0 && (
                <button
                  onClick={() => downloadServicesSpreadsheet(
                    servicesReport.items, 
                    servicesReport.summary, 
                    servicesFileName ? `Отчет_${servicesFileName}` : 'Отчет_по_услугам.xlsx',
                    servicesReport.halls,
                    servicesReport.suppliersBreakdown
                  )}
                  className="px-4 py-2.5 bg-white text-indigo-900 hover:bg-indigo-50 font-bold rounded-xl text-xs flex items-center gap-2 shadow-sm transition-all cursor-pointer whitespace-nowrap"
                  title="Скачать сформированный отчет по услугам в Excel"
                >
                  <Download className="w-4 h-4 text-indigo-700" />
                  <span>Экспорт в Excel (XLSX)</span>
                </button>
              )}
            </div>
          </div>
        </div>

        <YclientsPanel
          title="Загрузить отчёт по услугам из Yclients"
          description="Услуги из записей клиентов за выбранный период (оказанные визиты) загрузятся напрямую из Yclients, файл не нужен."
          onLoad={loadServicesFromYclients}
        />

        {/* Upload Box or Active File Summary */}
        {servicesRecords.length === 0 ? (
          <div 
            onDragEnter={(e) => { e.preventDefault(); setServicesDragActive(true); }}
            onDragOver={(e) => { e.preventDefault(); setServicesDragActive(true); }}
            onDragLeave={(e) => { e.preventDefault(); setServicesDragActive(false); }}
            onDrop={handleServicesDrop}
            className={`bg-white rounded-3xl border-2 border-dashed p-8 text-center transition-all shadow-sm ${
              servicesDragActive ? 'border-indigo-500 bg-indigo-50/20' : 'border-gray-200 hover:border-indigo-300'
            }`}
          >
            <div className="max-w-md mx-auto space-y-4">
              <div className="w-16 h-16 rounded-2xl bg-indigo-50 flex items-center justify-center mx-auto text-indigo-600">
                <Upload className="w-8 h-8" />
              </div>
              <div>
                <h3 className="font-display font-bold text-lg text-gray-900">
                  Загрузите файл с данными по услугам
                </h3>
                <p className="text-xs text-gray-500 mt-1 leading-relaxed">
                  Перетащите файл сюда или нажмите кнопку ниже. Поддерживаются форматы Excel <strong>.xlsx, .xls</strong> и <strong>.csv</strong>
                </p>
              </div>

              <div className="flex flex-col sm:flex-row gap-2.5 justify-center pt-2">
                <button
                  onClick={() => servicesFileInputRef.current?.click()}
                  className="px-5 py-2.5 bg-indigo-600 hover:bg-indigo-700 text-white rounded-xl text-xs font-bold shadow-md shadow-indigo-150 transition-all cursor-pointer flex items-center justify-center gap-2"
                >
                  <Upload className="w-4 h-4" />
                  <span>Выбрать файл по услугам</span>
                </button>
                <button
                  onClick={loadDemoServicesData}
                  className="px-5 py-2.5 bg-indigo-50 hover:bg-indigo-100 text-indigo-700 border border-indigo-200 rounded-xl text-xs font-semibold transition-all cursor-pointer flex items-center justify-center gap-2"
                >
                  <Sparkles className="w-4 h-4 text-amber-500" />
                  <span>Загрузить Демо-данные услуг</span>
                </button>
              </div>

              <div className="pt-4 border-t border-gray-100 text-[11px] text-gray-400 space-y-1">
                <div>Ожидаемые колонки: <span className="font-mono text-gray-600">«Услуга», «Скидка», «Баллы», «Абонементы», «Сертификаты», «Оплачено со счетов клиентов, ₽», «Оплачено деньгами, ₽»</span></div>
                <div className="text-indigo-600 font-medium flex items-center justify-center gap-1">
                  <Sparkles className="w-3 h-3 text-indigo-500" />
                  <span>Включена автоматическая предобработка числовых показателей (удаление пробелов, валют ₽/$, бухгалтерских скобок)</span>
                </div>
              </div>
            </div>
          </div>
        ) : (
          <div className="bg-white rounded-2xl border border-gray-150 p-4 shadow-sm flex flex-col sm:flex-row items-center justify-between gap-4">
            <div className="flex items-center gap-3">
              <div className="w-10 h-10 rounded-xl bg-indigo-50 border border-indigo-100 flex items-center justify-center text-indigo-600">
                <FileText className="w-5 h-5" />
              </div>
              <div>
                <div className="flex items-center gap-2 flex-wrap">
                  <span className="font-bold text-sm text-gray-900">{servicesFileName || 'Файл_по_услугам.xlsx'}</span>
                  <span className="text-[10px] px-2 py-0.2 rounded-full font-bold bg-emerald-50 text-emerald-700 border border-emerald-150">
                    Загружен
                  </span>
                  <span className="text-[10px] px-2 py-0.2 rounded-full font-bold bg-indigo-50 text-indigo-700 border border-indigo-150 flex items-center gap-1">
                    <Sparkles className="w-2.5 h-2.5 text-indigo-500" />
                    <span>Числовая предобработка активна</span>
                  </span>
                </div>
                <div className="text-xs text-gray-500 mt-0.5 flex flex-wrap items-center gap-2 sm:gap-3">
                  <span>Исходных строк: <strong className="text-gray-700">{servicesRecords.length}</strong></span>
                  <span>•</span>
                  <span>Сгруппировано услуг: <strong className="text-indigo-700">{servicesReport.items.length}</strong></span>
                  <span>•</span>
                  <span>Очищено числовых ячеек: <strong className="text-emerald-700">{servicesStats?.cleanedFieldsCount || (servicesRecords.length * 6)}</strong></span>
                </div>
              </div>
            </div>

            <div className="flex items-center gap-2 flex-wrap">
              <button
                onClick={() => servicesFileInputRef.current?.click()}
                className="px-3.5 py-1.5 bg-gray-50 hover:bg-gray-100 border border-gray-200 text-gray-700 rounded-xl text-xs font-semibold transition-all cursor-pointer flex items-center gap-1.5"
                title="Загрузить другой файл по услугам"
              >
                <RefreshCw className="w-3.5 h-3.5" />
                <span>Сменить файл</span>
              </button>
              <button
                onClick={() => {
                  setServicesRecords([]);
                  setServicesFileName(null);
                  setNotification({
                    status: 'success',
                    message: 'Данные по услугам сброшены. Вы можете загрузить новый файл.'
                  });
                }}
                className="px-3 py-1.5 bg-rose-50 hover:bg-rose-100 border border-rose-200 text-rose-700 rounded-xl text-xs font-semibold transition-all cursor-pointer flex items-center gap-1"
                title="Очистить загруженные услуги"
              >
                <Trash2 className="w-3.5 h-3.5" />
                <span>Очистить</span>
              </button>
            </div>
          </div>
        )}

        {/* Financial Report Costs Allocation Linkage Panel */}
        <div className="bg-white rounded-3xl border border-gray-150 p-6 shadow-sm space-y-4">
          <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3 border-b border-gray-100 pb-3">
            <div className="flex items-center gap-2">
              <SlidersHorizontal className="w-4.5 h-4.5 text-indigo-600" />
              <h3 className="font-display font-bold text-sm text-gray-900">
                Распределение затрат из финансового отчета (1-я вкладка)
              </h3>
            </div>
            
            {/* Status indicator badge */}
            <div className="flex items-center gap-2">
              {(totalSupplierExpensesFromReport > 0 || variableExpensesFromReport > 0 || fixedExpensesFromReport > 0) ? (
                <span className="text-[11px] font-semibold text-emerald-700 bg-emerald-50 border border-emerald-150 px-2.5 py-1 rounded-xl flex items-center gap-1.5">
                  <Check className="w-3.5 h-3.5 text-emerald-600" />
                  <span>Затраты подгружены из финансового отчета</span>
                </span>
              ) : (
                <span className="text-[11px] font-semibold text-amber-700 bg-amber-50 border border-amber-150 px-2.5 py-1 rounded-xl flex items-center gap-1.5">
                  <AlertCircle className="w-3.5 h-3.5 text-amber-600" />
                  <span>В 1-й вкладке финотчет еще не рассчитан (можно ввести вручную)</span>
                </span>
              )}

              <button
                onClick={() => {
                  if (!isServicesCustomExpensesActive) {
                    setCustomVariableExpenses(String(activeVariableExpenses));
                    setCustomFixedExpenses(String(activeFixedExpenses));
                    setCustomSupplierExpenses(String(servicesReport.summary.totalSupplierCost));
                  }
                  setIsServicesCustomExpensesActive(!isServicesCustomExpensesActive);
                }}
                className="text-[11px] font-semibold text-indigo-700 hover:text-indigo-800 bg-indigo-50 border border-indigo-150 px-2.5 py-1 rounded-xl cursor-pointer transition-all"
              >
                {isServicesCustomExpensesActive ? 'Вернуть авто-связку' : 'Ручная корректировка'}
              </button>
            </div>
          </div>

          <div className="grid grid-cols-1 md:grid-cols-2 xl:grid-cols-4 gap-4">
            {/* Suppliers Cost Card */}
            <div className="bg-blue-50/40 rounded-2xl p-4 border border-blue-150/60 flex flex-col justify-between">
              <div>
                <div className="flex items-center justify-between">
                  <span className="text-[10px] uppercase font-bold text-blue-900 tracking-wider">Поставщики (материалы)</span>
                  <span className="text-[10px] bg-blue-100 text-blue-800 font-mono px-1.5 py-0.2 rounded font-bold">Статья 53 • По залам</span>
                </div>

                {isServicesCustomExpensesActive ? (
                  <div className="mt-2">
                    <input
                      type="number"
                      value={customSupplierExpenses}
                      onChange={(e) => setCustomSupplierExpenses(e.target.value)}
                      placeholder="Сумма в рублях"
                      className="w-full text-lg font-mono font-bold bg-white border border-blue-300 rounded-xl px-3 py-1.5 text-blue-900 focus:outline-none focus:ring-1 focus:ring-blue-400"
                    />
                    <span className="text-[10px] text-blue-700 mt-1 block">Ручной ввод расходов на поставщиков</span>
                  </div>
                ) : (
                  <div className="mt-2">
                    <div className="font-display font-extrabold text-2xl text-blue-950">
                      {servicesReport.summary.totalSupplierCost.toLocaleString('ru-RU')} ₽
                    </div>
                    <div className="text-[11px] text-blue-700/90 mt-1 leading-relaxed">
                      Распределяются строго по залам согласно регламенту таблицы (прямые затраты зала + доля общих закупок пропорционально выручке).
                    </div>
                  </div>
                )}
              </div>

              <div className="mt-3 pt-2.5 border-t border-blue-100 flex items-center justify-between">
                <button
                  type="button"
                  onClick={() => setShowSuppliersRulesModal(prev => !prev)}
                  className="text-[11px] font-bold text-blue-700 hover:text-blue-900 flex items-center gap-1 cursor-pointer transition-colors"
                >
                  <span>{showSuppliersRulesModal ? 'Скрыть таблицу регламента ▲' : 'Таблица распределения по залам ▼'}</span>
                </button>
                <span className="text-[10px] font-mono text-blue-600 bg-blue-100/70 px-1.5 py-0.5 rounded">
                  {servicesReport.suppliersBreakdown?.items?.length || 0} поставщиков
                </span>
              </div>
            </div>

            {/* Variable Costs Card */}
            <div className="bg-amber-50/40 rounded-2xl p-4 border border-amber-150/50">
              <div className="flex items-center justify-between">
                <span className="text-[10px] uppercase font-bold text-amber-800 tracking-wider">Переменные затраты</span>
                <span className="text-[10px] bg-amber-100 text-amber-800 font-mono px-1.5 py-0.2 rounded font-bold">Статья 9</span>
              </div>

              {isServicesCustomExpensesActive ? (
                <div className="mt-2">
                  <input
                    type="number"
                    value={customVariableExpenses}
                    onChange={(e) => setCustomVariableExpenses(e.target.value)}
                    placeholder="Сумма в рублях"
                    className="w-full text-lg font-mono font-bold bg-white border border-amber-300 rounded-xl px-3 py-1.5 text-amber-900 focus:outline-none focus:ring-1 focus:ring-amber-400"
                  />
                  <span className="text-[10px] text-amber-700 mt-1 block">Ручной ввод переменных затрат</span>
                </div>
              ) : (
                <div className="mt-2">
                  <div className="font-display font-extrabold text-2xl text-amber-900">
                    {activeVariableExpenses.toLocaleString('ru-RU')} ₽
                  </div>
                  <div className="text-[11px] text-amber-700/90 mt-1 leading-relaxed">
                    Сумма строки «Расходы переменные» из финансового отчета. Распределяется пропорционально выручке.
                  </div>
                </div>
              )}
            </div>

            {/* Fixed Costs Card */}
            <div className="bg-purple-50/40 rounded-2xl p-4 border border-purple-150/50">
              <div className="flex items-center justify-between">
                <span className="text-[10px] uppercase font-bold text-purple-800 tracking-wider">Постоянные затраты</span>
                <span className="text-[10px] bg-purple-100 text-purple-800 font-mono px-1.5 py-0.2 rounded font-bold">Статья 36</span>
              </div>

              {isServicesCustomExpensesActive ? (
                <div className="mt-2">
                  <input
                    type="number"
                    value={customFixedExpenses}
                    onChange={(e) => setCustomFixedExpenses(e.target.value)}
                    placeholder="Сумма в рублях"
                    className="w-full text-lg font-mono font-bold bg-white border border-purple-300 rounded-xl px-3 py-1.5 text-purple-900 focus:outline-none focus:ring-1 focus:ring-purple-400"
                  />
                  <span className="text-[10px] text-purple-700 mt-1 block">Ручной ввод постоянных затрат</span>
                </div>
              ) : (
                <div className="mt-2">
                  <div className="font-display font-extrabold text-2xl text-purple-900">
                    {activeFixedExpenses.toLocaleString('ru-RU')} ₽
                  </div>
                  <div className="text-[11px] text-purple-700/90 mt-1 leading-relaxed">
                    Сумма строки «Расходы постоянные» из финансового отчета. Распределяется пропорционально выручке.
                  </div>
                </div>
              )}
            </div>

            {/* Total Cost to Allocate Card */}
            <div className="bg-indigo-50/40 rounded-2xl p-4 border border-indigo-150/50 flex flex-col justify-between">
              <div>
                <div className="flex items-center justify-between">
                  <span className="text-[10px] uppercase font-bold text-indigo-800 tracking-wider">Всего затрат к распределению</span>
                  <span className="text-[10px] bg-indigo-100 text-indigo-800 font-mono px-1.5 py-0.2 rounded font-bold">100% базы</span>
                </div>
                <div className="font-display font-extrabold text-2xl text-indigo-950 mt-2">
                  {servicesReport.summary.totalCost.toLocaleString('ru-RU')} ₽
                </div>
                <div className="text-[11px] text-indigo-700/90 mt-1 leading-relaxed">
                  Поставщики ({servicesReport.summary.totalSupplierCost.toLocaleString('ru-RU')} ₽) + Переменные ({servicesReport.summary.totalVariableCost.toLocaleString('ru-RU')} ₽) + Постоянные ({servicesReport.summary.totalFixedCost.toLocaleString('ru-RU')} ₽).
                </div>
              </div>

              {rawRecords.length === 0 && (
                <button
                  onClick={loadDemoData}
                  className="mt-3 text-[11px] font-bold text-indigo-600 hover:text-indigo-800 flex items-center gap-1 cursor-pointer"
                >
                  <span>Загрузить Демо-финансовый отчет для связки →</span>
                </button>
              )}
            </div>
          </div>

          {/* Interactive Collapsible Suppliers Allocation Regulation Panel */}
          {showSuppliersRulesModal && (
            <div className="mt-4 p-5 bg-gradient-to-br from-blue-50/70 via-indigo-50/50 to-white rounded-2xl border-2 border-blue-200/80 shadow-sm space-y-4">
              <div className="flex items-center justify-between border-b border-blue-150 pb-3">
                <div className="flex items-center gap-2">
                  <Package className="w-5 h-5 text-blue-700" />
                  <div>
                    <h4 className="font-display font-bold text-sm text-blue-950">
                      Регламент распределения расходов на поставщиков между залами
                    </h4>
                    <p className="text-xs text-blue-800/80 mt-0.5">
                      Прямые поставщики относятся на закрепленный зал, общие закупки («распределяются между всеми») распределяются пропорционально доле зала в общей выручке.
                    </p>
                  </div>
                </div>
                <button
                  onClick={() => setShowSuppliersRulesModal(false)}
                  className="px-3 py-1 bg-white hover:bg-blue-100 text-blue-800 border border-blue-200 rounded-xl text-xs font-semibold shadow-2xs transition-colors flex items-center gap-1 cursor-pointer"
                >
                  <ChevronUp className="w-3.5 h-3.5" />
                  <span>Свернуть регламент</span>
                </button>
              </div>

              {/* Hall Suppliers Allocation Summary Cards */}
              <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-5 gap-3">
                {servicesReport.suppliersBreakdown?.hallSuppliersSummary?.map((hSum) => {
                  const meta = HALL_METADATA[hSum.hallName as HallName] || HALL_METADATA['Прочие услуги'];
                  return (
                    <div 
                      key={`h_supp_card_${hSum.hallName}`}
                      className="bg-white p-3.5 rounded-xl border border-blue-150/80 shadow-2xs flex flex-col justify-between"
                    >
                      <div>
                        <div className="flex items-center justify-between mb-1.5">
                          <span className={`text-[10px] font-bold px-2 py-0.5 rounded-md ${meta.badgeBg} ${meta.badgeText} border ${meta.badgeBorder}`}>
                            {meta.shortLabel}
                          </span>
                          <span className="text-[10px] font-mono text-gray-400">
                            {hSum.items.length} поставщ.
                          </span>
                        </div>
                        <div className="font-bold text-xs text-gray-900 line-clamp-1" title={hSum.hallName}>
                          {hSum.hallName}
                        </div>
                      </div>

                      <div className="mt-2.5 pt-2 border-t border-gray-100 space-y-1 text-[11px] font-mono">
                        <div className="flex justify-between items-baseline text-gray-500">
                          <span>Прямые:</span>
                          <span className="font-bold text-gray-800">{hSum.directCost.toLocaleString('ru-RU')} ₽</span>
                        </div>
                        <div className="flex justify-between items-baseline text-gray-500">
                          <span>Доля общих:</span>
                          <span className="font-semibold text-blue-700">{hSum.allocatedSharedCost.toLocaleString('ru-RU')} ₽</span>
                        </div>
                        <div className="flex justify-between items-baseline pt-1 border-t border-dashed border-gray-200">
                          <span className="font-sans font-bold text-blue-950 text-[10px] uppercase">Итого залу:</span>
                          <span className="font-extrabold text-blue-900 text-xs">{hSum.totalSupplierCost.toLocaleString('ru-RU')} ₽</span>
                        </div>
                      </div>
                    </div>
                  );
                })}
              </div>

              {/* Classification rules reference table based on user's specification */}
              <div className="bg-white rounded-xl border border-blue-150 p-4 shadow-2xs">
                <div className="text-xs font-bold text-gray-800 mb-2 flex items-center justify-between">
                  <span>Таблица регламента соответствия статей и залов:</span>
                  <span className="text-[10px] font-mono text-gray-500 font-normal">
                    Всего правил в справочнике: {SUPPLIER_HALL_RULES.length}
                  </span>
                </div>
                <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-3 text-xs">
                  <div className="p-3 bg-blue-50/40 rounded-lg border border-blue-100">
                    <div className="font-bold text-blue-950 text-[11px] mb-1">1. Зал ногтевого сервиса</div>
                    <ul className="list-disc list-inside text-[11px] text-gray-700 space-y-0.5">
                      <li>Материалы для ногтевиков</li>
                      <li>ООО "ВЕЛЛНЕС БЬЮТИ" (La Sultane de Saba)</li>
                      <li>Материалы для эпиляции</li>
                    </ul>
                  </div>

                  <div className="p-3 bg-amber-50/40 rounded-lg border border-amber-100">
                    <div className="font-bold text-amber-950 text-[11px] mb-1">2. Зал бровей</div>
                    <ul className="list-disc list-inside text-[11px] text-gray-700 space-y-0.5">
                      <li>Материалы для бровистов</li>
                      <li>Материалы для ресниц</li>
                    </ul>
                  </div>

                  <div className="p-3 bg-pink-50/40 rounded-lg border border-pink-100">
                    <div className="font-bold text-pink-950 text-[11px] mb-1">6. Зал макияжей и укладок</div>
                    <ul className="list-disc list-inside text-[11px] text-gray-700 space-y-0.5">
                      <li>ИП Щеголькова (ProfiLine)</li>
                    </ul>
                  </div>

                  <div className="p-3 bg-cyan-50/40 rounded-lg border border-cyan-100">
                    <div className="font-bold text-cyan-950 text-[11px] mb-1">4. Мужской зал</div>
                    <ul className="list-disc list-inside text-[11px] text-gray-700 space-y-0.5">
                      <li>ИП Щеголькова (American CREW)</li>
                      <li>Расходники для барберов</li>
                    </ul>
                  </div>

                  <div className="p-3 bg-emerald-50/40 rounded-lg border border-emerald-100">
                    <div className="font-bold text-emerald-950 text-[11px] mb-1">5. Распределяются между всеми</div>
                    <ul className="list-disc list-inside text-[11px] text-gray-700 space-y-0.5">
                      <li>Торговый Дом "БлескСервис" (TORK)</li>
                      <li>Прочие закупки, Заказ с маркетплейсов</li>
                      <li>(распределяются пропорционально выручке залов)</li>
                    </ul>
                  </div>

                  <div className="p-3 bg-purple-50/40 rounded-lg border border-purple-100">
                    <div className="font-bold text-purple-950 text-[11px] mb-1">3. Зал колористики / парикмахеры</div>
                    <ul className="list-disc list-inside text-[10.5px] text-gray-700 space-y-0.5 line-clamp-4" title="Constant Delight, Худорожков/Демин, Калашникова (Matrix, L'Oreal), Коротков (Дайвинес), Керастас, ГЕРА ГРУП, Эстель Кутюр, Уходы без поставщика, Кэвиар Групп (NASHI), TOKIO, KEUNE, Солком Регион, DR.HURTINA, Organic Works">
                      <li>Constant Delight (расходники, витрина)</li>
                      <li>Худорожков/Демин, Калашникова (Matrix, L'Oreal)</li>
                      <li>Коротков (Дайвинес), Керастас, ГЕРА ГРУП</li>
                      <li>Кэвиар Групп (NASHI), TOKIO, KEUNE, Солком и др.</li>
                    </ul>
                  </div>
                </div>

                <div className="mt-3 pt-2.5 border-t border-gray-150 flex items-center justify-between">
                  <span className="text-[11px] text-gray-500">
                    Общие расходы распределяются автоматически между залами салона на основе формулы: <code className="bg-gray-100 px-1 py-0.5 rounded text-indigo-700 font-mono">Общие_расходы × (Выручка_зала / Общая_выручка_салона)</code>
                  </span>
                  <button
                    onClick={() => setShowSuppliersRulesModal(false)}
                    className="text-xs text-blue-700 font-semibold hover:underline cursor-pointer"
                  >
                    Закрыть справочник ▲
                  </button>
                </div>
              </div>
            </div>
          )}
        </div>

        {/* Hall Area and Salaries Allocation Parameters Panel */}
        <div className="bg-white rounded-3xl border border-gray-150 p-6 shadow-sm space-y-5">
          <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3 border-b border-gray-100 pb-3">
            <div className="flex items-center gap-2.5">
              <Building2 className="w-5 h-5 text-indigo-600" />
              <div>
                <div className="flex items-center gap-2 flex-wrap">
                  <h3 className="font-display font-bold text-sm text-gray-900">
                    Ввод параметров залов: Площадь (м²) и Зарплата персонала (ФОТ)
                  </h3>
                  <span className="text-[10px] font-semibold bg-emerald-100 text-emerald-800 px-2 py-0.5 rounded-full border border-emerald-200">
                    ✓ Данные зарплаты заполняют статьи в Финансовом отчете
                  </span>
                </div>
                <p className="text-xs text-gray-500 mt-0.5">
                  Зарплата мастеров учитывается строго для каждого зала в введенной сумме и суммарно заполняет статью «Зарплата мастеров» финансового отчета. Зарплата административного персонала распределяется пропорционально выручке между всеми залами и заполняет статью «Зарплата административного персонала».
                </p>
              </div>
            </div>

            <div className="flex items-center gap-2 flex-wrap">
              <span className="text-[11px] font-mono font-bold text-sky-850 bg-sky-50 border border-sky-200 px-2.5 py-1 rounded-xl">
                Итого ЗП мастеров: {totalMastersSalaryFromServices.toLocaleString('ru-RU')} ₽
              </span>
              <span className="text-[11px] font-mono font-bold text-teal-850 bg-teal-50 border border-teal-200 px-2.5 py-1 rounded-xl">
                ЗП админов: {serviceAdminSalary.toLocaleString('ru-RU')} ₽
              </span>
              <span className="text-[11px] font-mono font-bold text-indigo-800 bg-indigo-50 border border-indigo-150 px-2.5 py-1 rounded-xl">
                Общая площадь: {servicesReport.summary.totalArea} м² • Всего ФОТ: {servicesReport.summary.totalSalaryCost.toLocaleString('ru-RU')} ₽
              </span>
            </div>
          </div>

          {/* Cost Allocation Method Selector */}
          <div className="bg-gradient-to-r from-indigo-50/70 via-slate-50 to-white rounded-2xl p-4 border border-indigo-200/90 shadow-2xs flex flex-col lg:flex-row lg:items-center justify-between gap-4">
            <div className="flex items-start gap-3">
              <div className="w-10 h-10 rounded-xl bg-indigo-100 text-indigo-700 flex items-center justify-center font-bold text-lg shrink-0 shadow-2xs">
                <SlidersHorizontal className="w-5 h-5 text-indigo-700" />
              </div>
              <div>
                <div className="flex items-center gap-2 flex-wrap">
                  <span className="text-xs font-bold text-gray-900 uppercase tracking-wide">
                    Способ распределения затрат по залам
                  </span>
                  <span className="text-[10px] font-semibold bg-indigo-100 text-indigo-800 px-2 py-0.5 rounded-full border border-indigo-200">
                    Переменные, постоянные расходы и ЗП админов
                  </span>
                </div>
                <p className="text-[11px] text-gray-600 mt-1 leading-relaxed">
                  {expenseAllocationMethod === 'revenue' 
                    ? 'Текущий режим: Пропорционально выручке. Все общие затраты распределяются между залами салона строго пропорционально их доле в общей выручке.'
                    : 'Текущий режим: Пропорционально площади залов. Затраты распределяются согласно площади залов (м²). Если по залу отсутствует выручка, затраты на него не распределяются (0 ₽).'}
                </p>
              </div>
            </div>

            <div className="flex items-center gap-1.5 bg-gray-150/80 p-1 rounded-xl shrink-0 border border-gray-250">
              <button
                type="button"
                onClick={() => setExpenseAllocationMethod('revenue')}
                className={`px-3.5 py-1.5 rounded-lg text-xs font-bold transition-all cursor-pointer flex items-center gap-1.5 ${
                  expenseAllocationMethod === 'revenue'
                    ? 'bg-white text-indigo-900 shadow-2xs border border-gray-200'
                    : 'text-gray-600 hover:text-gray-900'
                }`}
              >
                <span>Пропорционально выручке</span>
                {expenseAllocationMethod === 'revenue' && <Check className="w-3.5 h-3.5 text-indigo-600" />}
              </button>
              <button
                type="button"
                onClick={() => setExpenseAllocationMethod('area')}
                className={`px-3.5 py-1.5 rounded-lg text-xs font-bold transition-all cursor-pointer flex items-center gap-1.5 ${
                  expenseAllocationMethod === 'area'
                    ? 'bg-indigo-600 text-white shadow-2xs'
                    : 'text-gray-600 hover:text-gray-900'
                }`}
              >
                <span>Пропорционально площади</span>
                {expenseAllocationMethod === 'area' && <Check className="w-3.5 h-3.5 text-white" />}
              </button>
            </div>
          </div>

          {/* Admin Salary Distribution Banner */}
          <div className="bg-gradient-to-r from-teal-50/70 via-sky-50/40 to-white rounded-2xl p-4 border border-teal-200/80 flex flex-col md:flex-row md:items-center justify-between gap-4 shadow-2xs">
            <div className="flex items-start gap-3">
              <div className="w-10 h-10 rounded-xl bg-teal-100/90 text-teal-800 flex items-center justify-center font-bold text-lg shrink-0 shadow-2xs">
                👥
              </div>
              <div>
                <div className="flex items-center gap-2 flex-wrap">
                  <span className="text-xs font-bold text-teal-950 uppercase tracking-wide">
                    Зарплата административного персонала (Общий ФОТ админов)
                  </span>
                  <span className="text-[10px] font-mono bg-teal-100 text-teal-800 px-2 py-0.5 rounded-full font-bold">
                    Распределяется между залами
                  </span>
                  <span className="text-[10px] font-semibold bg-emerald-100 text-emerald-800 px-2 py-0.5 rounded-full border border-emerald-200">
                    Передается в Финотчет
                  </span>
                </div>
                <p className="text-[11px] text-teal-800/80 mt-1 leading-relaxed">
                  Введенная сумма автоматически заполняет статью «Зарплата административного персонала» в финансовом отчете и распределяется между залами согласно выбранному способу распределения ({expenseAllocationMethod === 'revenue' ? 'пропорционально выручке' : 'пропорционально площади залов'}).
                </p>
              </div>
            </div>

            <div className="flex items-center gap-2.5 shrink-0">
              <div className="relative">
                <input
                  type="number"
                  min="0"
                  step="1000"
                  value={serviceAdminSalary}
                  onChange={(e) => setServiceAdminSalary(Math.max(0, parseFloat(e.target.value) || 0))}
                  placeholder="Сумма ЗП админов"
                  className="w-48 text-right pr-8 text-base font-mono font-bold bg-white border border-teal-300 rounded-xl px-3 py-1.5 text-teal-950 focus:outline-none focus:ring-2 focus:ring-teal-400 shadow-2xs"
                />
                <span className="absolute right-3 top-2 text-xs font-mono font-bold text-teal-600">₽</span>
              </div>
              {serviceAdminSalary > 0 && (
                <button
                  type="button"
                  onClick={() => setServiceAdminSalary(0)}
                  className="text-[11px] font-semibold text-teal-700 hover:text-teal-950 underline cursor-pointer"
                  title="Очистить значение"
                >
                  Очистить
                </button>
              )}
            </div>
          </div>

          {/* 10 Halls Area & Master Salary Cards Grid */}
          <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 xl:grid-cols-5 gap-3.5">
            {HALL_NAMES.filter(h => h !== 'Прочие услуги').map((hName) => {
              const meta = HALL_METADATA[hName] || HALL_METADATA['Прочие услуги'];
              const currentArea = hallAreas[hName] ?? 0;
              const currentSalary = hallMastersSalaries[hName] ?? 0;
              const hallRep = servicesReport.halls.find(h => h.hallName === hName);

              return (
                <div 
                  key={`hall_param_${hName}`}
                  className="bg-white rounded-2xl border border-gray-200 p-4 shadow-2xs hover:border-indigo-300 transition-all flex flex-col justify-between space-y-3"
                >
                  <div>
                    <div className="flex items-center justify-between mb-1.5">
                      <span className={`text-[10px] font-bold px-2 py-0.5 rounded-md ${meta.badgeBg} ${meta.badgeText} border ${meta.badgeBorder}`}>
                        {meta.shortLabel}
                      </span>
                      <span className="text-[10px] font-mono text-gray-400 font-semibold">
                        {hallRep ? `${hallRep.summary.totalRevenueShare}% выручки` : '—'}
                      </span>
                    </div>
                    <h4 className="font-bold text-xs text-gray-900 line-clamp-1" title={hName}>
                      {hName}
                    </h4>
                  </div>

                  <div className="space-y-2 pt-1 border-t border-gray-100">
                    {/* Area Input */}
                    <div>
                      <div className="flex justify-between items-center mb-1">
                        <label className="text-[10px] uppercase font-bold text-gray-500">
                          Площадь зала:
                        </label>
                        {hallRep && currentArea > 0 && (
                          <span className="text-[9.5px] font-mono text-gray-400">
                            {hallRep.summary.areaShare}% пл.
                          </span>
                        )}
                      </div>
                      <div className="relative">
                        <input
                          type="number"
                          min="0"
                          step="0.5"
                          value={currentArea}
                          onChange={(e) => {
                            const val = parseFloat(e.target.value) || 0;
                            setHallAreas(prev => ({ ...prev, [hName]: val }));
                          }}
                          className="w-full text-sm font-mono font-bold bg-gray-50 hover:bg-white focus:bg-white border border-gray-250 focus:border-indigo-500 rounded-xl px-2.5 py-1 text-gray-900 pr-8 outline-none"
                        />
                        <span className="absolute right-2.5 top-1.5 text-xs font-mono text-gray-400 font-semibold">м²</span>
                      </div>
                    </div>

                    {/* Master Salary Input */}
                    <div>
                      <label className="text-[10px] uppercase font-bold text-sky-900 block mb-1">
                        ЗП мастеров зала:
                      </label>
                      <div className="relative">
                        <input
                          type="number"
                          min="0"
                          step="1000"
                          value={currentSalary}
                          onChange={(e) => {
                            const val = parseFloat(e.target.value) || 0;
                            setHallMastersSalaries(prev => ({ ...prev, [hName]: val }));
                          }}
                          className="w-full text-sm font-mono font-bold bg-sky-50/50 hover:bg-white focus:bg-white border border-sky-200 focus:border-sky-500 rounded-xl px-2.5 py-1 text-sky-950 pr-7 outline-none"
                        />
                        <span className="absolute right-2.5 top-1.5 text-xs font-mono text-sky-600 font-bold">₽</span>
                      </div>
                      <span className="text-[9px] text-sky-750 block mt-0.5 font-sans">
                        Строго на данный зал
                      </span>
                    </div>
                  </div>

                  {/* Calculated metrics */}
                  {hallRep && (
                    <div className="pt-2 border-t border-gray-100 space-y-1 text-[11px] text-gray-600 font-mono">
                      <div className="flex justify-between items-center text-[10px]">
                        <span className="text-gray-400 font-sans">Доля ЗП админов:</span>
                        <span className="text-teal-800 font-semibold">{hallRep.summary.adminSalary.toLocaleString('ru-RU')} ₽</span>
                      </div>
                      <div className="flex justify-between items-center text-[10px]">
                        <span className="text-gray-400 font-sans">Всего ФОТ зала:</span>
                        <span className="text-indigo-900 font-bold">{hallRep.summary.totalSalaryCost.toLocaleString('ru-RU')} ₽</span>
                      </div>
                      {currentArea > 0 && (
                        <div className="flex justify-between items-center text-[10px] pt-1 border-t border-dashed border-gray-150">
                          <span className="text-gray-400 font-sans">Выручка / м²:</span>
                          <span className="text-emerald-800 font-semibold">{hallRep.summary.revenuePerSqM.toLocaleString('ru-RU')} ₽</span>
                        </div>
                      )}
                      {expenseAllocationMethod === 'area' && hallRep.summary.totalRevenue === 0 && (
                        <div className="pt-1 text-[10px] text-amber-800 font-sans flex items-center gap-1 bg-amber-50 px-2 py-1 rounded-lg border border-amber-200">
                          <span>⚠️ Нет выручки — затраты не распределяются (0 ₽)</span>
                        </div>
                      )}
                    </div>
                  )}
                </div>
              );
            })}
          </div>
        </div>

        {/* KPI Bento Grid for Services */}
        <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-5 gap-4">
          <div className="bg-white rounded-2xl border border-gray-150 p-4.5 shadow-sm">
            <div className="text-[10px] uppercase font-bold text-gray-400 tracking-wider">Сумма выручки по услугам</div>
            <div className="font-display font-extrabold text-2xl text-emerald-600 mt-2">
              {servicesReport.summary.totalRevenue.toLocaleString('ru-RU')} ₽
            </div>
            <div className="text-xs text-gray-500 mt-1">
              {servicesReport.summary.totalCount} выполненных услуг
            </div>
          </div>

          <div className="bg-white rounded-2xl border border-gray-150 p-4.5 shadow-sm">
            <div className="text-[10px] uppercase font-bold text-gray-400 tracking-wider">Всего затрат по услугам</div>
            <div className="font-display font-extrabold text-2xl text-rose-500 mt-2">
              {servicesReport.summary.totalCost.toLocaleString('ru-RU')} ₽
            </div>
            <div className="text-[10.5px] text-gray-500 mt-1 flex flex-col gap-0.5">
              <span>Поставщики: {servicesReport.summary.totalSupplierCost.toLocaleString('ru-RU')} ₽</span>
              <span className="text-[10px] text-gray-400">ФОТ: {servicesReport.summary.totalSalaryCost.toLocaleString('ru-RU')} ₽ • Пер/Пост: {(servicesReport.summary.totalVariableCost + servicesReport.summary.totalFixedCost).toLocaleString('ru-RU')} ₽</span>
            </div>
          </div>

          <div className="bg-white rounded-2xl border border-gray-150 p-4.5 shadow-sm">
            <div className="text-[10px] uppercase font-bold text-sky-800 tracking-wider">Общий ФОТ персонала</div>
            <div className="font-display font-extrabold text-2xl text-sky-950 mt-2">
              {servicesReport.summary.totalSalaryCost.toLocaleString('ru-RU')} ₽
            </div>
            <div className="text-[10.5px] text-sky-800/80 mt-1 flex flex-col gap-0.5">
              <span>Мастера: {servicesReport.summary.totalMasterSalary.toLocaleString('ru-RU')} ₽ (по залам)</span>
              <span className="text-[10px] text-teal-700">Админы: {servicesReport.summary.totalAdminSalary.toLocaleString('ru-RU')} ₽ (между всеми)</span>
            </div>
          </div>

          <div className="bg-white rounded-2xl border border-gray-150 p-4.5 shadow-sm">
            <div className="text-[10px] uppercase font-bold text-indigo-800 tracking-wider">Площадь и удельные</div>
            <div className="font-display font-extrabold text-2xl text-indigo-950 mt-2">
              {servicesReport.summary.totalArea} м²
            </div>
            <div className="text-[10.5px] text-indigo-700 mt-1 flex flex-col gap-0.5">
              <span>Выручка: {(servicesReport.summary.totalArea > 0 ? (servicesReport.summary.totalRevenue / servicesReport.summary.totalArea).toFixed(0) : 0).toLocaleString('ru-RU')} ₽/м²</span>
              <span className="text-[10px] text-gray-400">Прибыль: {(servicesReport.summary.totalArea > 0 ? (servicesReport.summary.totalProfit / servicesReport.summary.totalArea).toFixed(0) : 0).toLocaleString('ru-RU')} ₽/м²</span>
            </div>
          </div>

          <div className="bg-white rounded-2xl border border-gray-150 p-4.5 shadow-sm">
            <div className="text-[10px] uppercase font-bold text-gray-400 tracking-wider">Прибыль / Рентабельность</div>
            <div className="font-display font-extrabold text-2xl text-indigo-700 mt-2">
              {servicesReport.summary.totalProfit.toLocaleString('ru-RU')} ₽
            </div>
            <div className="flex items-center gap-1.5 text-xs text-gray-500 mt-1">
              <span className="font-mono bg-indigo-50 text-indigo-700 font-bold px-1.5 py-0.5 rounded text-[10px]">
                {servicesReport.summary.totalProfitability.toFixed(1)}% рентабельность
              </span>
              <span>всего салона</span>
            </div>
          </div>
        </div>

        {/* Hall Performance Overview Bento Grid */}
        {servicesReport.items.length > 0 && (
          <div className="space-y-2.5">
            <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-1">
              <div className="flex items-center gap-2">
                <Building2 className="w-4 h-4 text-indigo-600" />
                <h3 className="font-display font-bold text-xs uppercase tracking-wider text-gray-700">
                  Группировка услуг по залам салона ({servicesReport.halls.filter(h => h.items.length > 0).length} активных залов)
                </h3>
              </div>
              <span className="text-[11px] text-gray-400">
                Кликните на зал для быстрой фильтрации или выберите в панели ниже
              </span>
            </div>

            <div className="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-3 lg:grid-cols-6 gap-3">
              {HALL_NAMES.map((hName) => {
                const meta = HALL_METADATA[hName];
                const hGroup = servicesReport.halls.find(h => h.hallName === hName);
                const count = hGroup ? hGroup.items.length : 0;
                const rev = hGroup ? hGroup.summary.totalRevenue : 0;
                const share = hGroup ? hGroup.summary.totalRevenueShare : 0;
                const profit = hGroup ? hGroup.summary.totalProfit : 0;
                const isSelected = selectedHallFilter === hName;

                if (count === 0 && hName === 'Прочие услуги') return null;

                return (
                  <div
                    key={`hall_card_${hName}`}
                    onClick={() => setSelectedHallFilter(isSelected ? 'all' : hName)}
                    className={`p-3.5 rounded-2xl border transition-all cursor-pointer flex flex-col justify-between ${
                      isSelected
                        ? 'bg-indigo-50/90 border-indigo-500 shadow-sm ring-2 ring-indigo-200'
                        : count > 0
                          ? 'bg-white hover:bg-gray-50 border-gray-200 shadow-2xs'
                          : 'bg-gray-50/50 border-gray-150 opacity-60'
                    }`}
                  >
                    <div>
                      <div className="flex items-center justify-between gap-1 mb-1.5">
                        <span className={`text-[10px] font-bold px-2 py-0.5 rounded-md ${meta.badgeBg} ${meta.badgeText} border ${meta.badgeBorder}`}>
                          {meta.shortLabel}
                        </span>
                        <span className="text-[10.5px] font-mono text-gray-400 font-semibold">
                          {count} усл.
                        </span>
                      </div>
                      <div className="font-display font-bold text-xs text-gray-900 leading-snug line-clamp-1" title={hName}>
                        {hName}
                      </div>
                      {hGroup && hGroup.summary.area > 0 && (
                        <div className="text-[10px] font-mono text-gray-400 mt-0.5">
                          {hGroup.summary.area} м² ({hGroup.summary.areaShare}%)
                        </div>
                      )}
                    </div>

                    <div className="mt-2.5 pt-2 border-t border-gray-100 space-y-1 text-xs">
                      <div className="flex justify-between items-baseline">
                        <span className="text-[10px] text-gray-400">Выручка:</span>
                        <span className="font-mono font-bold text-gray-900 text-[11px]">
                          {rev.toLocaleString('ru-RU')} ₽
                        </span>
                      </div>
                      <div className="flex justify-between items-baseline text-[10px]">
                        <span className="text-gray-400">Поставщики:</span>
                        <span className="font-mono font-semibold text-blue-700">
                          {(hGroup?.summary.totalSupplierCost || 0).toLocaleString('ru-RU')} ₽
                        </span>
                      </div>
                      <div className="flex justify-between items-baseline text-[10px]">
                        <span className="text-gray-400">ЗП мастеров:</span>
                        <span className="font-mono font-semibold text-sky-800">
                          {(hGroup?.summary.masterSalary || 0).toLocaleString('ru-RU')} ₽
                        </span>
                      </div>
                      <div className="flex justify-between items-baseline text-[10px]">
                        <span className="text-gray-400">ЗП админов:</span>
                        <span className="font-mono font-semibold text-teal-800">
                          {(hGroup?.summary.adminSalary || 0).toLocaleString('ru-RU')} ₽
                        </span>
                      </div>
                      <div className="flex justify-between items-baseline text-[10px] pt-1 border-t border-gray-100">
                        <span className="text-gray-400 font-semibold">Прибыль:</span>
                        <span className={`font-mono font-bold ${profit >= 0 ? 'text-emerald-700' : 'text-rose-600'}`}>
                          {profit.toLocaleString('ru-RU')} ₽
                        </span>
                      </div>
                    </div>
                  </div>
                );
              })}
            </div>
          </div>
        )}

        {/* "Прочие услуги" Allocation Callout Banner */}
        {otherServicesItems.length > 0 && (
          <div className="bg-gradient-to-r from-amber-50 to-orange-50 border-2 border-amber-300 rounded-3xl p-5 shadow-sm space-y-4">
            <div className="flex flex-col sm:flex-row items-start sm:items-center justify-between gap-3">
              <div className="flex items-start gap-3">
                <div className="p-2.5 bg-amber-500 text-white rounded-2xl shadow-sm shrink-0">
                  <AlertCircle className="w-5 h-5" />
                </div>
                <div>
                  <div className="flex items-center gap-2 flex-wrap">
                    <h3 className="font-display font-extrabold text-sm text-amber-950">
                      «Прочие услуги» требуют распределения по залам ({otherServicesItems.length} позиций)
                    </h3>
                    <span className="text-[11px] font-mono font-bold px-2 py-0.5 rounded-full bg-amber-200 text-amber-900">
                      {otherServicesItems.reduce((acc, it) => acc + it.revenue, 0).toLocaleString('ru-RU')} ₽ выручки
                    </span>
                  </div>
                  <p className="text-xs text-amber-900/80 mt-1 max-w-3xl leading-relaxed">
                    Услуги ниже не попали под стандартные правила автоопределения залов. 
                    Выберите подходящий зал салона для каждой услуги — <strong>система запомнит ваш выбор</strong> и автоматически применит его во всех отчетах, пересчитав выручку залов, затраты на материалы и заработную плату.
                  </p>
                </div>
              </div>

              <div className="flex items-center gap-2 shrink-0">
                <button
                  onClick={() => setShowHallOverridesModal(true)}
                  className="px-3.5 py-1.5 bg-white hover:bg-amber-100 text-amber-900 border border-amber-300 rounded-xl text-xs font-bold transition-all shadow-2xs flex items-center gap-1.5 cursor-pointer"
                >
                  <Settings2 className="w-3.5 h-3.5 text-amber-700" />
                  <span>Все назначения ({Object.keys(serviceHallOverrides).length})</span>
                </button>
              </div>
            </div>

            {/* Quick Assignment Items Cards */}
            <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-3 pt-2 border-t border-amber-200/80">
              {otherServicesItems.map((item) => (
                <div 
                  key={`other_assign_card_${item.service}`}
                  className="bg-white rounded-2xl border border-amber-200 p-3.5 shadow-2xs flex flex-col justify-between space-y-2.5 hover:border-amber-400 transition-all"
                >
                  <div>
                    <div className="font-semibold text-xs text-gray-900 leading-snug" title={item.service}>
                      {item.service}
                    </div>
                    <div className="flex items-center justify-between mt-1 text-[11px] font-mono">
                      <span className="text-gray-400">{item.count} операций</span>
                      <strong className="text-emerald-700 font-bold">{item.revenue.toLocaleString('ru-RU')} ₽</strong>
                    </div>
                  </div>

                  <div className="pt-2 border-t border-gray-100">
                    <label className="text-[10px] uppercase font-bold text-gray-400 block mb-1">
                      Отнести к залу:
                    </label>
                    <select
                      value=""
                      onChange={(e) => {
                        if (e.target.value) {
                          handleAssignServiceToHall(item.service, e.target.value as HallName);
                        }
                      }}
                      className="w-full text-xs font-semibold bg-amber-50/60 hover:bg-amber-50 focus:bg-white border border-amber-300 focus:border-indigo-500 rounded-xl px-2.5 py-1.5 text-gray-900 outline-none cursor-pointer"
                    >
                      <option value="">— Выберите зал салона —</option>
                      {HALL_NAMES.filter(h => h !== 'Прочие услуги').map((hName) => (
                        <option key={`opt_card_${item.service}_${hName}`} value={hName}>
                          {hName}
                        </option>
                      ))}
                    </select>
                  </div>
                </div>
              ))}
            </div>
          </div>
        )}

        {/* Success badge if all services are allocated and overrides exist */}
        {otherServicesItems.length === 0 && Object.keys(serviceHallOverrides).length > 0 && (
          <div className="bg-emerald-50 border border-emerald-200 rounded-2xl p-3 flex items-center justify-between gap-3 text-xs text-emerald-900 shadow-2xs">
            <div className="flex items-center gap-2">
              <CheckCircle className="w-4 h-4 text-emerald-600 shrink-0" />
              <span>Все услуги распределены по залам салона. Активно <strong>{Object.keys(serviceHallOverrides).length}</strong> сохраненных ручных соответствий.</span>
            </div>
            <button
              onClick={() => setShowHallOverridesModal(true)}
              className="px-3 py-1 bg-white hover:bg-emerald-100 text-emerald-800 border border-emerald-300 rounded-lg font-semibold text-[11px] cursor-pointer"
            >
              Управление соответствиями
            </button>
          </div>
        )}

        {/* Toolbar: Hall Filters, Search and Download */}
        <div className="bg-white rounded-2xl border border-gray-150 p-4 shadow-sm space-y-3">
          {/* Hall Filter Tabs */}
          <div className="flex items-center gap-1.5 overflow-x-auto pb-1 text-xs">
            <button
              onClick={() => setSelectedHallFilter('all')}
              className={`px-3 py-1.5 rounded-xl text-xs font-semibold whitespace-nowrap transition-all cursor-pointer ${
                selectedHallFilter === 'all'
                  ? 'bg-indigo-600 text-white shadow-xs'
                  : 'bg-gray-100 hover:bg-gray-200 text-gray-700'
              }`}
            >
              Все залы ({servicesReport.items.length})
            </button>
            {HALL_NAMES.map((hName) => {
              const meta = HALL_METADATA[hName];
              const hGroup = servicesReport.halls.find(h => h.hallName === hName);
              const count = hGroup ? hGroup.items.length : 0;
              if (count === 0 && hName === 'Прочие услуги') return null;

              return (
                <button
                  key={`filter_btn_${hName}`}
                  onClick={() => setSelectedHallFilter(selectedHallFilter === hName ? 'all' : hName)}
                  className={`px-3 py-1.5 rounded-xl text-xs font-semibold whitespace-nowrap transition-all cursor-pointer flex items-center gap-1.5 ${
                    selectedHallFilter === hName
                      ? 'bg-indigo-600 text-white shadow-xs'
                      : 'bg-gray-100 hover:bg-gray-200 text-gray-700'
                  }`}
                >
                  <span>{meta.shortLabel}</span>
                  <span className={`text-[10px] font-mono px-1.5 py-0.2 rounded-md ${
                    selectedHallFilter === hName ? 'bg-indigo-700 text-white' : 'bg-gray-200 text-gray-700'
                  }`}>
                    {count}
                  </span>
                </button>
              );
            })}
          </div>

          <div className="flex flex-col sm:flex-row gap-4 items-center justify-between pt-2 border-t border-gray-100">
            <div className="flex items-center gap-2 bg-gray-50 px-3.5 py-2 rounded-xl border border-gray-150 w-full sm:max-w-md">
              <Search className="w-4 h-4 text-gray-400" />
              <input
                type="text"
                value={servicesSearchQuery}
                onChange={(e) => setServicesSearchQuery(e.target.value)}
                placeholder="Поиск по наименованию услуги..."
                className="bg-transparent text-xs w-full outline-none text-gray-700 border-none inline-block focus:ring-0"
              />
              {servicesSearchQuery && (
                <button 
                  onClick={() => setServicesSearchQuery('')} 
                  className="text-gray-400 hover:text-gray-600 text-xs font-medium"
                >
                  Очистить
                </button>
              )}
            </div>

            <div className="flex items-center gap-3 w-full sm:w-auto justify-between sm:justify-end flex-wrap">
              {/* Hall Overrides manager button */}
              <button
                onClick={() => setShowHallOverridesModal(true)}
                className="flex items-center gap-1.5 px-3 py-1.5 bg-indigo-50 hover:bg-indigo-100 text-indigo-700 border border-indigo-200 rounded-xl text-xs font-semibold cursor-pointer shadow-2xs transition-all"
                title="Просмотр и управление сохраненными соответствиями услуг залам"
              >
                <Building2 className="w-3.5 h-3.5 text-indigo-600" />
                <span>Назначения залов</span>
                {Object.keys(serviceHallOverrides).length > 0 && (
                  <span className="px-1.5 py-0.2 rounded-md text-[10px] bg-indigo-600 text-white font-bold font-mono">
                    {Object.keys(serviceHallOverrides).length}
                  </span>
                )}
              </button>

              {/* Bulk Expand / Collapse Buttons for Halls and Services */}
              <div className="flex items-center gap-1.5 flex-wrap">
                <div className="flex items-center gap-1 bg-gray-50 p-1 rounded-xl border border-gray-150">
                  <button
                    onClick={() => toggleAllHalls(false)}
                    className="flex items-center gap-1 px-2.5 py-1 rounded-lg text-xs font-semibold text-gray-700 hover:text-indigo-700 hover:bg-white transition-all cursor-pointer shadow-2xs"
                    title="Развернуть все залы"
                  >
                    <ChevronsUpDown className="w-3.5 h-3.5 text-indigo-600" />
                    <span>Развернуть все залы</span>
                  </button>
                  <button
                    onClick={() => toggleAllHalls(true)}
                    className="flex items-center gap-1 px-2.5 py-1 rounded-lg text-xs font-semibold text-gray-700 hover:text-indigo-700 hover:bg-white transition-all cursor-pointer shadow-2xs"
                    title="Свернуть все залы"
                  >
                    <ChevronsDownUp className="w-3.5 h-3.5 text-gray-500" />
                    <span>Свернуть залы</span>
                  </button>
                </div>

                <div className="flex items-center gap-1 bg-gray-50 p-1 rounded-xl border border-gray-150">
                  <button
                    onClick={() => toggleAllServiceItems(true)}
                    className="flex items-center gap-1 px-2.5 py-1 rounded-lg text-xs font-semibold text-gray-700 hover:text-indigo-700 hover:bg-white transition-all cursor-pointer shadow-2xs"
                    title="Развернуть детализацию всех услуг (строки CRM)"
                  >
                    <ChevronsUpDown className="w-3.5 h-3.5 text-indigo-600" />
                    <span>Все детализации</span>
                  </button>
                  <button
                    onClick={() => toggleAllServiceItems(false)}
                    className="flex items-center gap-1 px-2.5 py-1 rounded-lg text-xs font-semibold text-gray-700 hover:text-indigo-700 hover:bg-white transition-all cursor-pointer shadow-2xs"
                    title="Свернуть все открытые детализации услуг"
                  >
                    <ChevronsDownUp className="w-3.5 h-3.5 text-gray-500" />
                    <span>Свернуть все услуги</span>
                  </button>
                </div>
              </div>

              <span className="text-xs text-gray-500 hidden xl:inline">
                Показано позиций: <strong className="text-gray-900">{filteredServiceItems.length}</strong> из {servicesReport.items.length}
              </span>

              {servicesReport.items.length > 0 && (
                <button
                  onClick={() => downloadServicesSpreadsheet(
                    servicesReport.items, 
                    servicesReport.summary, 
                    servicesFileName ? `Отчет_${servicesFileName}` : 'Отчет_по_услугам.xlsx',
                    servicesReport.halls,
                    servicesReport.suppliersBreakdown
                  )}
                  className="px-4 py-2 bg-emerald-600 hover:bg-emerald-700 text-white rounded-xl text-xs font-semibold shadow-sm transition-all cursor-pointer flex items-center gap-1.5 whitespace-nowrap"
                  title="Скачать Excel-таблицу со сводкой по залам, группировкой и подитогами"
                >
                  <Download className="w-3.5 h-3.5" />
                  <span>Скачать XLSX по залам</span>
                </button>
              )}
            </div>
          </div>
        </div>

        {/* Main Services Report Table Grouped by Hall */}
        <div className="bg-white rounded-3xl border border-gray-150 shadow-sm overflow-hidden">
          <div className="p-4 bg-gray-50/70 border-b border-gray-150 flex items-center justify-between select-none">
            <div className="flex items-center gap-2">
              <Table className="w-4 h-4 text-indigo-600" />
              <h3 className="font-display font-bold text-xs uppercase tracking-wider text-gray-700">
                Отчет по услугам с группировкой по залам и пропорциональным распределением затрат
              </h3>
            </div>
            <span className="text-[11px] text-indigo-700 bg-indigo-50 border border-indigo-150 px-2.5 py-0.5 rounded-md font-semibold">
              Залов: {filteredHalls.length} • Услуг: {filteredServiceItems.length}
            </span>
          </div>

          {filteredServiceItems.length === 0 ? (
            <div className="p-16 text-center text-gray-400 text-xs">
              {servicesSearchQuery 
                ? `По запросу «${servicesSearchQuery}» ничего не найдено.` 
                : 'Нет данных для отображения. Загрузите файл с данными по услугам выше или нажмите «Загрузить Демо Услуг».'}
            </div>
          ) : (
            <div className="overflow-x-auto">
              <table className="w-full text-left text-xs border-collapse">
                <thead>
                  <tr className="bg-gray-50 text-gray-600 font-bold select-none h-11 border-b border-gray-150 text-[10.5px]">
                    <th className="p-3 text-center w-12 font-mono text-[10px] text-gray-400">№</th>
                    <th className="p-3 min-w-[220px]">Услуга</th>
                    <th className="p-3 text-center w-16">Кол-во</th>
                    <th className="p-3 text-right w-24">Скидка</th>
                    <th className="p-3 text-right w-20">Баллы</th>
                    <th className="p-3 text-right w-24">Абонементы</th>
                    <th className="p-3 text-right w-24">Сертификаты</th>
                    <th className="p-3 text-right w-28">Со счетов, ₽</th>
                    <th className="p-3 text-right w-28">Деньгами, ₽</th>
                    <th className="p-3 text-right w-32 font-bold text-indigo-950 bg-indigo-50/30">Сумма выручки</th>
                    <th className="p-3 text-center w-20 bg-indigo-50/20">Доля</th>
                    <th className="p-3 text-right w-32 text-blue-900 bg-blue-50/25">Поставщики (материалы)</th>
                    <th className="p-3 text-right w-28 text-sky-900 bg-sky-50/25 font-bold">ЗП мастеров</th>
                    <th className="p-3 text-right w-28 text-teal-900 bg-teal-50/25 font-bold">ЗП админов</th>
                    <th className="p-3 text-right w-32 text-amber-800 bg-amber-50/20">Переменные затраты</th>
                    <th className="p-3 text-right w-32 text-purple-800 bg-purple-50/20">Постоянные затраты</th>
                    <th className="p-3 text-right w-32 text-rose-800 bg-rose-50/20">Всего затрат</th>
                    <th className="p-3 text-right w-32 text-emerald-800 bg-emerald-50/20 font-bold">Прибыль</th>
                    <th className="p-3 text-center w-24 pr-4">Рентабельность</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-gray-100 font-sans text-gray-800">
                  {filteredHalls.map((hallGroup) => {
                    const meta = HALL_METADATA[hallGroup.hallName as HallName] || HALL_METADATA['Прочие услуги'];
                    const isHallCollapsed = !!collapsedHalls[hallGroup.hallName];

                    return (
                      <React.Fragment key={`hall_section_${hallGroup.hallName}`}>
                        {/* Hall Section Header Row */}
                        <tr 
                          onClick={() => setCollapsedHalls(prev => ({ ...prev, [hallGroup.hallName]: !prev[hallGroup.hallName] }))}
                          className="bg-gray-100/90 hover:bg-gray-150 border-t-2 border-b border-gray-200 cursor-pointer select-none transition-colors"
                        >
                          <td colSpan={19} className="p-3">
                            <div className="flex flex-wrap items-center justify-between gap-2">
                              <div className="flex items-center gap-2.5">
                                <ChevronDown className={`w-4 h-4 text-gray-500 transition-transform ${isHallCollapsed ? '-rotate-90' : ''}`} />
                                <span className={`text-xs font-bold px-2.5 py-0.5 rounded-lg ${meta.badgeBg} ${meta.badgeText} border ${meta.badgeBorder}`}>
                                  {hallGroup.hallName}
                                </span>
                                {hallGroup.hallName === 'Прочие услуги' && hallGroup.items.length > 0 && (
                                  <span className="text-[10px] font-bold px-2 py-0.5 rounded-md bg-amber-100 text-amber-900 border border-amber-300">
                                    ⚠️ Требуется отнесение к залам
                                  </span>
                                )}
                                <span className="text-[11px] text-gray-500 font-medium">
                                  ({hallGroup.items.length} позиций • {hallGroup.summary.totalCount} операций)
                                </span>
                              </div>

                              <div className="flex items-center gap-3 sm:gap-4 text-xs font-mono">
                                <div>
                                  <span className="text-[10px] uppercase font-bold text-gray-400 mr-1.5 font-sans">Выручка зала:</span>
                                  <strong className="text-gray-900 font-extrabold">{hallGroup.summary.totalRevenue.toLocaleString('ru-RU')} ₽</strong>
                                  <span className="text-indigo-600 font-semibold ml-1">({hallGroup.summary.totalRevenueShare}%)</span>
                                </div>
                                {hallGroup.summary.area > 0 && (
                                  <div className="hidden xl:block">
                                    <span className="text-[10px] uppercase font-bold text-gray-400 mr-1.5 font-sans">Площадь:</span>
                                    <span className="text-indigo-900 font-semibold">{hallGroup.summary.area} м²</span>
                                  </div>
                                )}
                                <div className="hidden lg:block">
                                  <span className="text-[10px] uppercase font-bold text-gray-400 mr-1.5 font-sans">Поставщики:</span>
                                  <span className="text-blue-700 font-semibold">{hallGroup.summary.totalSupplierCost.toLocaleString('ru-RU')} ₽</span>
                                </div>
                                <div className="hidden lg:block">
                                  <span className="text-[10px] uppercase font-bold text-gray-400 mr-1.5 font-sans">ЗП мастеров:</span>
                                  <span className="text-sky-800 font-semibold">{hallGroup.summary.masterSalary.toLocaleString('ru-RU')} ₽</span>
                                </div>
                                <div className="hidden xl:block">
                                  <span className="text-[10px] uppercase font-bold text-gray-400 mr-1.5 font-sans">ЗП админов:</span>
                                  <span className="text-teal-800 font-semibold">{hallGroup.summary.adminSalary.toLocaleString('ru-RU')} ₽</span>
                                </div>
                                <div className="hidden md:block">
                                  <span className="text-[10px] uppercase font-bold text-gray-400 mr-1.5 font-sans">Затраты:</span>
                                  <span className="text-rose-700">{hallGroup.summary.totalCost.toLocaleString('ru-RU')} ₽</span>
                                </div>
                                <div>
                                  <span className="text-[10px] uppercase font-bold text-gray-400 mr-1.5 font-sans">Прибыль:</span>
                                  <strong className={hallGroup.summary.totalProfit >= 0 ? 'text-emerald-700' : 'text-rose-600'}>
                                    {hallGroup.summary.totalProfit.toLocaleString('ru-RU')} ₽
                                  </strong>
                                  <span className="text-[10px] text-gray-500 ml-1">({hallGroup.summary.totalProfitability.toFixed(1)}%)</span>
                                </div>

                                <button
                                  onClick={(e) => {
                                    e.stopPropagation();
                                    setCollapsedHalls(prev => ({ ...prev, [hallGroup.hallName]: !prev[hallGroup.hallName] }));
                                  }}
                                  className="px-2 py-0.5 text-[10.5px] font-semibold text-gray-600 hover:text-indigo-700 bg-white hover:bg-gray-50 rounded-lg border border-gray-250 transition-all flex items-center gap-1 cursor-pointer shadow-2xs ml-1"
                                  title={isHallCollapsed ? "Развернуть услуги этого зала" : "Свернуть услуги этого зала"}
                                >
                                  {isHallCollapsed ? (
                                    <>
                                      <ChevronDown className="w-3.5 h-3.5 text-indigo-600" />
                                      <span>Развернуть зал</span>
                                    </>
                                  ) : (
                                    <>
                                      <ChevronUp className="w-3.5 h-3.5 text-gray-500" />
                                      <span>Свернуть зал</span>
                                    </>
                                  )}
                                </button>
                              </div>
                            </div>
                          </td>
                        </tr>

                        {/* Services rows belonging to this Hall */}
                        {!isHallCollapsed && hallGroup.items.map((item, idx) => {
                          const isExpanded = !!expandedServiceItems[item.service];

                          return (
                            <React.Fragment key={`svc_row_${hallGroup.hallName}_${idx}`}>
                              <tr 
                                className="hover:bg-indigo-50/25 transition-colors h-11 cursor-pointer select-none"
                                onClick={() => setExpandedServiceItems(prev => ({ ...prev, [item.service]: !prev[item.service] }))}
                              >
                                <td className="p-3 text-center font-mono text-[11px] text-gray-400 border-r border-gray-50">
                                  {idx + 1}
                                </td>
                                <td className="p-3 font-medium text-gray-900">
                                  <div className="flex items-center gap-2 flex-wrap">
                                    <div className="flex items-center gap-1.5 shrink-0">
                                      <ChevronDown className={`w-3.5 h-3.5 text-indigo-500 transition-transform shrink-0 ${isExpanded ? 'rotate-180' : ''}`} />
                                      <span className="font-semibold text-gray-900">{item.service}</span>
                                    </div>

                                    {/* If this service is currently in "Прочие услуги", show quick hall assignment select */}
                                    {item.hall === 'Прочие услуги' && (
                                      <div className="inline-flex items-center gap-1" onClick={(e) => e.stopPropagation()}>
                                        <select
                                          value=""
                                          onChange={(e) => {
                                            if (e.target.value) {
                                              handleAssignServiceToHall(item.service, e.target.value as HallName);
                                            }
                                          }}
                                          className="text-[10.5px] font-bold bg-amber-50 hover:bg-amber-100 text-amber-900 border border-amber-300 rounded-lg px-2 py-0.5 outline-none cursor-pointer shadow-2xs focus:ring-2 focus:ring-amber-400"
                                          title="Отнести эту услугу к одному из залов салона с сохранением выбора"
                                        >
                                          <option value="">➡️ Отнести к залу...</option>
                                          {HALL_NAMES.filter(h => h !== 'Прочие услуги').map(h => (
                                            <option key={`row_opt_${item.service}_${h}`} value={h}>
                                              {h}
                                            </option>
                                          ))}
                                        </select>
                                      </div>
                                    )}

                                    {/* If this service was manually assigned via overrides */}
                                    {serviceHallOverrides[item.service] && (
                                      <div 
                                        className="inline-flex items-center gap-1 px-1.5 py-0.5 rounded-md text-[9.5px] font-bold bg-indigo-50 text-indigo-800 border border-indigo-200" 
                                        onClick={(e) => e.stopPropagation()}
                                        title={`Назначено пользователем в «${serviceHallOverrides[item.service]}». Сохранено в браузере.`}
                                      >
                                        <span>✏️ Назначено</span>
                                        <button
                                          type="button"
                                          onClick={() => handleResetServiceHall(item.service)}
                                          className="text-indigo-400 hover:text-rose-600 transition-colors ml-0.5 cursor-pointer font-bold"
                                          title="Сбросить ручное назначение"
                                        >
                                          ✕
                                        </button>
                                      </div>
                                    )}
                                  </div>
                                </td>
                                <td className="p-3 text-center font-mono font-semibold text-indigo-700">
                                  {item.count}
                                </td>
                                <td className="p-3 text-right font-mono text-gray-600">
                                  {item.discount > 0 ? `${item.discount.toLocaleString('ru-RU')} ₽` : '—'}
                                </td>
                                <td className="p-3 text-right font-mono text-gray-600">
                                  {item.points > 0 ? `${item.points.toLocaleString('ru-RU')} ₽` : '—'}
                                </td>
                                <td className="p-3 text-right font-mono text-gray-600">
                                  {item.subscriptions > 0 ? `${item.subscriptions.toLocaleString('ru-RU')} ₽` : '—'}
                                </td>
                                <td className="p-3 text-right font-mono text-gray-600">
                                  {item.certificates > 0 ? `${item.certificates.toLocaleString('ru-RU')} ₽` : '—'}
                                </td>
                                <td className="p-3 text-right font-mono text-gray-600">
                                  {item.clientAccounts > 0 ? `${item.clientAccounts.toLocaleString('ru-RU')} ₽` : '—'}
                                </td>
                                <td className="p-3 text-right font-mono text-gray-600">
                                  {item.money > 0 ? `${item.money.toLocaleString('ru-RU')} ₽` : '—'}
                                </td>
                                <td className="p-3 text-right font-mono font-bold text-indigo-950 bg-indigo-50/30">
                                  {item.revenue.toLocaleString('ru-RU', { minimumFractionDigits: 2, maximumFractionDigits: 2 })} ₽
                                </td>
                                <td className="p-3 text-center font-mono text-gray-600 bg-indigo-50/20 text-[11px]">
                                  {item.revenueShare.toFixed(1)}%
                                </td>
                                <td className="p-3 text-right font-mono text-blue-900 bg-blue-50/20 font-medium">
                                  {item.supplierCost.toLocaleString('ru-RU', { minimumFractionDigits: 2, maximumFractionDigits: 2 })} ₽
                                </td>
                                <td className="p-3 text-right font-mono text-sky-900 bg-sky-50/20 font-medium">
                                  {item.masterSalary.toLocaleString('ru-RU', { minimumFractionDigits: 2, maximumFractionDigits: 2 })} ₽
                                </td>
                                <td className="p-3 text-right font-mono text-teal-900 bg-teal-50/20 font-medium">
                                  {item.adminSalary.toLocaleString('ru-RU', { minimumFractionDigits: 2, maximumFractionDigits: 2 })} ₽
                                </td>
                                <td className="p-3 text-right font-mono text-amber-900 bg-amber-50/20 font-medium">
                                  {item.variableCost.toLocaleString('ru-RU', { minimumFractionDigits: 2, maximumFractionDigits: 2 })} ₽
                                </td>
                                <td className="p-3 text-right font-mono text-purple-900 bg-purple-50/20 font-medium">
                                  {item.fixedCost.toLocaleString('ru-RU', { minimumFractionDigits: 2, maximumFractionDigits: 2 })} ₽
                                </td>
                                <td className="p-3 text-right font-mono text-rose-700 bg-rose-50/20 font-semibold">
                                  {item.totalCost.toLocaleString('ru-RU', { minimumFractionDigits: 2, maximumFractionDigits: 2 })} ₽
                                </td>
                                <td className={`p-3 text-right font-mono font-bold bg-emerald-50/20 ${item.profit >= 0 ? 'text-emerald-700' : 'text-rose-600'}`}>
                                  {item.profit.toLocaleString('ru-RU', { minimumFractionDigits: 2, maximumFractionDigits: 2 })} ₽
                                </td>
                                <td className="p-3 text-center pr-4 font-mono text-[11px] font-semibold">
                                  <span className={`px-1.5 py-0.5 rounded ${item.profitability >= 0 ? 'bg-emerald-50 text-emerald-800' : 'bg-rose-50 text-rose-800'}`}>
                                    {item.profitability.toFixed(1)}%
                                  </span>
                                </td>
                              </tr>

                              {/* Expanded Drawer: All raw records for this grouped service */}
                              {isExpanded && item.rawRows && (
                                <tr className="bg-indigo-50/15">
                                  <td colSpan={19} className="p-4 pl-12 bg-indigo-50/25 border-b border-gray-150">
                                    <div className="bg-white rounded-2xl border border-indigo-150 shadow-sm overflow-hidden">
                                      <div className="p-3 bg-indigo-50/60 border-b border-indigo-100 flex items-center justify-between text-xs font-semibold text-indigo-950">
                                        <div className="flex items-center gap-2">
                                          <span>Все сгруппированные строки по услуге «{item.service}» ({item.rawRows.length} оп.):</span>
                                          <span className="font-mono text-indigo-700 font-bold">
                                            Итого выручка: {item.revenue.toLocaleString('ru-RU', { minimumFractionDigits: 2, maximumFractionDigits: 2 })} ₽
                                          </span>
                                        </div>
                                        <button
                                          onClick={(e) => {
                                            e.stopPropagation();
                                            setExpandedServiceItems(prev => ({ ...prev, [item.service]: false }));
                                          }}
                                          className="px-2.5 py-1 text-[11px] font-semibold text-indigo-700 hover:text-indigo-900 bg-white hover:bg-indigo-50 rounded-lg border border-indigo-200 transition-all flex items-center gap-1 cursor-pointer shadow-2xs"
                                          title="Свернуть детализацию этой услуги"
                                        >
                                          <ChevronUp className="w-3.5 h-3.5" />
                                          <span>Свернуть</span>
                                        </button>
                                      </div>
                                      <div className="overflow-x-auto">
                                        <table className="w-full text-left text-[11px] divide-y divide-gray-150">
                                          <thead className="bg-gray-50/80 font-bold text-gray-600 uppercase text-[9px] select-none">
                                            <tr>
                                              <th className="p-2.5 pl-4 w-16">Строка</th>
                                              <th className="p-2.5 text-right w-24">Скидка</th>
                                              <th className="p-2.5 text-right w-20">Баллы</th>
                                              <th className="p-2.5 text-right w-24">Абонементы</th>
                                              <th className="p-2.5 text-right w-24">Сертификаты</th>
                                              <th className="p-2.5 text-right w-28">Со счетов, ₽</th>
                                              <th className="p-2.5 text-right w-28">Деньгами, ₽</th>
                                              <th className="p-2.5 text-right pr-4 w-32 font-bold text-indigo-950">Сумма выручки</th>
                                            </tr>
                                          </thead>
                                          <tbody className="divide-y divide-gray-100 font-sans text-gray-700 font-mono">
                                            {item.rawRows.map((r, rIdx) => (
                                              <tr key={`raw_svc_${hallGroup.hallName}_${idx}_${rIdx}`} className="hover:bg-gray-50/60 transition-colors">
                                                <td className="p-2.5 pl-4 text-gray-400">
                                                  #{r.rawIndex}
                                                </td>
                                                <td className="p-2.5 text-right text-gray-600">
                                                  {r.discount > 0 ? `${r.discount.toLocaleString('ru-RU')} ₽` : '—'}
                                                </td>
                                                <td className="p-2.5 text-right text-gray-600">
                                                  {r.points > 0 ? `${r.points.toLocaleString('ru-RU')} ₽` : '—'}
                                                </td>
                                                <td className="p-2.5 text-right text-gray-600">
                                                  {r.subscriptions > 0 ? `${r.subscriptions.toLocaleString('ru-RU')} ₽` : '—'}
                                                </td>
                                                <td className="p-2.5 text-right text-gray-600">
                                                  {r.certificates > 0 ? `${r.certificates.toLocaleString('ru-RU')} ₽` : '—'}
                                                </td>
                                                <td className="p-2.5 text-right text-gray-600">
                                                  {r.clientAccounts > 0 ? `${r.clientAccounts.toLocaleString('ru-RU')} ₽` : '—'}
                                                </td>
                                                <td className="p-2.5 text-right text-gray-600">
                                                  {r.money > 0 ? `${r.money.toLocaleString('ru-RU')} ₽` : '—'}
                                                </td>
                                                <td className="p-2.5 text-right pr-4 font-bold text-indigo-950">
                                                  {r.revenue.toLocaleString('ru-RU', { minimumFractionDigits: 2, maximumFractionDigits: 2 })} ₽
                                                </td>
                                              </tr>
                                            ))}
                                          </tbody>
                                        </table>
                                      </div>
                                      {/* Footer Collapse Button */}
                                      <div className="p-2 bg-gray-50/60 border-t border-gray-100 flex justify-end">
                                        <button
                                          onClick={(e) => {
                                            e.stopPropagation();
                                            setExpandedServiceItems(prev => ({ ...prev, [item.service]: false }));
                                          }}
                                          className="px-2.5 py-1 text-[11px] font-semibold text-gray-600 hover:text-indigo-700 bg-white hover:bg-gray-100 rounded-lg border border-gray-200 transition-all flex items-center gap-1 cursor-pointer shadow-2xs"
                                        >
                                          <ChevronUp className="w-3.5 h-3.5 text-gray-400" />
                                          <span>Свернуть детализацию «{item.service}»</span>
                                        </button>
                                      </div>
                                    </div>
                                  </td>
                                </tr>
                              )}
                            </React.Fragment>
                          );
                        })}

                        {/* Hall Subtotal Row */}
                        {!isHallCollapsed && (
                          <tr className="bg-indigo-50/30 font-bold border-b-2 border-indigo-200/80 text-[11px] text-gray-900 select-none">
                            <td className="p-3 text-center font-mono text-indigo-700">∑</td>
                            <td className="p-3 font-extrabold text-indigo-950">
                              ИТОГО ПО ЗАЛУ: {hallGroup.hallName}
                            </td>
                            <td className="p-3 text-center font-mono font-bold text-indigo-700">
                              {hallGroup.summary.totalCount}
                            </td>
                            <td className="p-3 text-right font-mono">
                              {hallGroup.summary.totalDiscount.toLocaleString('ru-RU', { minimumFractionDigits: 2, maximumFractionDigits: 2 })} ₽
                            </td>
                            <td className="p-3 text-right font-mono">
                              {hallGroup.summary.totalPoints.toLocaleString('ru-RU', { minimumFractionDigits: 2, maximumFractionDigits: 2 })} ₽
                            </td>
                            <td className="p-3 text-right font-mono">
                              {hallGroup.summary.totalSubscriptions.toLocaleString('ru-RU', { minimumFractionDigits: 2, maximumFractionDigits: 2 })} ₽
                            </td>
                            <td className="p-3 text-right font-mono">
                              {hallGroup.summary.totalCertificates.toLocaleString('ru-RU', { minimumFractionDigits: 2, maximumFractionDigits: 2 })} ₽
                            </td>
                            <td className="p-3 text-right font-mono">
                              {hallGroup.summary.totalClientAccounts.toLocaleString('ru-RU', { minimumFractionDigits: 2, maximumFractionDigits: 2 })} ₽
                            </td>
                            <td className="p-3 text-right font-mono">
                              {hallGroup.summary.totalMoney.toLocaleString('ru-RU', { minimumFractionDigits: 2, maximumFractionDigits: 2 })} ₽
                            </td>
                            <td className="p-3 text-right font-mono font-extrabold text-indigo-950 bg-indigo-100/50">
                              {hallGroup.summary.totalRevenue.toLocaleString('ru-RU', { minimumFractionDigits: 2, maximumFractionDigits: 2 })} ₽
                            </td>
                            <td className="p-3 text-center font-mono font-bold text-indigo-700 bg-indigo-100/30">
                              {hallGroup.summary.totalRevenueShare}%
                            </td>
                            <td className="p-3 text-right font-mono text-blue-900 bg-blue-50/50 font-bold">
                              {hallGroup.summary.totalSupplierCost.toLocaleString('ru-RU', { minimumFractionDigits: 2, maximumFractionDigits: 2 })} ₽
                            </td>
                            <td className="p-3 text-right font-mono text-sky-900 bg-sky-50/50 font-bold">
                              {hallGroup.summary.masterSalary.toLocaleString('ru-RU', { minimumFractionDigits: 2, maximumFractionDigits: 2 })} ₽
                            </td>
                            <td className="p-3 text-right font-mono text-teal-900 bg-teal-50/50 font-bold">
                              {hallGroup.summary.adminSalary.toLocaleString('ru-RU', { minimumFractionDigits: 2, maximumFractionDigits: 2 })} ₽
                            </td>
                            <td className="p-3 text-right font-mono text-amber-900 bg-amber-50/50">
                              {hallGroup.summary.totalVariableCost.toLocaleString('ru-RU', { minimumFractionDigits: 2, maximumFractionDigits: 2 })} ₽
                            </td>
                            <td className="p-3 text-right font-mono text-purple-900 bg-purple-50/50">
                              {hallGroup.summary.totalFixedCost.toLocaleString('ru-RU', { minimumFractionDigits: 2, maximumFractionDigits: 2 })} ₽
                            </td>
                            <td className="p-3 text-right font-mono text-rose-800 bg-rose-50/50">
                              {hallGroup.summary.totalCost.toLocaleString('ru-RU', { minimumFractionDigits: 2, maximumFractionDigits: 2 })} ₽
                            </td>
                            <td className={`p-3 text-right font-mono font-extrabold bg-emerald-50/50 ${hallGroup.summary.totalProfit >= 0 ? 'text-emerald-800' : 'text-rose-700'}`}>
                              {hallGroup.summary.totalProfit.toLocaleString('ru-RU', { minimumFractionDigits: 2, maximumFractionDigits: 2 })} ₽
                            </td>
                            <td className="p-3 text-center pr-4 font-mono font-bold">
                              <span className={`px-1.5 py-0.5 rounded ${hallGroup.summary.totalProfitability >= 0 ? 'bg-emerald-100 text-emerald-900' : 'bg-rose-100 text-rose-900'}`}>
                                {hallGroup.summary.totalProfitability.toFixed(1)}%
                              </span>
                            </td>
                          </tr>
                        )}
                      </React.Fragment>
                    );
                  })}
                </tbody>

                {/* Grand Total Row (Общий итог по всем залам и услугам) */}
                <tfoot className="bg-gray-100 font-bold border-t-2 border-indigo-400 text-[11px] text-gray-900 select-none">
                  <tr className="h-12">
                    <td className="p-3 text-center font-mono text-gray-500">∑</td>
                    <td className="p-3 font-extrabold uppercase text-indigo-950">
                      ИТОГО ПО ВСЕМ ЗАЛАМ И УСЛУГАМ
                    </td>
                    <td className="p-3 text-center font-mono font-extrabold text-indigo-700">
                      {servicesReport.summary.totalCount}
                    </td>
                    <td className="p-3 text-right font-mono">
                      {servicesReport.summary.totalDiscount.toLocaleString('ru-RU', { minimumFractionDigits: 2, maximumFractionDigits: 2 })} ₽
                    </td>
                    <td className="p-3 text-right font-mono">
                      {servicesReport.summary.totalPoints.toLocaleString('ru-RU', { minimumFractionDigits: 2, maximumFractionDigits: 2 })} ₽
                    </td>
                    <td className="p-3 text-right font-mono">
                      {servicesReport.summary.totalSubscriptions.toLocaleString('ru-RU', { minimumFractionDigits: 2, maximumFractionDigits: 2 })} ₽
                    </td>
                    <td className="p-3 text-right font-mono">
                      {servicesReport.summary.totalCertificates.toLocaleString('ru-RU', { minimumFractionDigits: 2, maximumFractionDigits: 2 })} ₽
                    </td>
                    <td className="p-3 text-right font-mono">
                      {servicesReport.summary.totalClientAccounts.toLocaleString('ru-RU', { minimumFractionDigits: 2, maximumFractionDigits: 2 })} ₽
                    </td>
                    <td className="p-3 text-right font-mono">
                      {servicesReport.summary.totalMoney.toLocaleString('ru-RU', { minimumFractionDigits: 2, maximumFractionDigits: 2 })} ₽
                    </td>
                    <td className="p-3 text-right font-mono font-extrabold text-indigo-950 bg-indigo-150/40 text-xs">
                      {servicesReport.summary.totalRevenue.toLocaleString('ru-RU', { minimumFractionDigits: 2, maximumFractionDigits: 2 })} ₽
                    </td>
                    <td className="p-3 text-center font-mono font-bold bg-indigo-150/30">
                      100.0%
                    </td>
                    <td className="p-3 text-right font-mono text-blue-900 bg-blue-100/40 font-bold">
                      {servicesReport.summary.totalSupplierCost.toLocaleString('ru-RU', { minimumFractionDigits: 2, maximumFractionDigits: 2 })} ₽
                    </td>
                    <td className="p-3 text-right font-mono text-sky-900 bg-sky-100/40 font-bold">
                      {servicesReport.summary.totalMasterSalary.toLocaleString('ru-RU', { minimumFractionDigits: 2, maximumFractionDigits: 2 })} ₽
                    </td>
                    <td className="p-3 text-right font-mono text-teal-900 bg-teal-100/40 font-bold">
                      {servicesReport.summary.totalAdminSalary.toLocaleString('ru-RU', { minimumFractionDigits: 2, maximumFractionDigits: 2 })} ₽
                    </td>
                    <td className="p-3 text-right font-mono text-amber-900 bg-amber-100/40 font-bold">
                      {servicesReport.summary.totalVariableCost.toLocaleString('ru-RU', { minimumFractionDigits: 2, maximumFractionDigits: 2 })} ₽
                    </td>
                    <td className="p-3 text-right font-mono text-purple-900 bg-purple-100/40 font-bold">
                      {servicesReport.summary.totalFixedCost.toLocaleString('ru-RU', { minimumFractionDigits: 2, maximumFractionDigits: 2 })} ₽
                    </td>
                    <td className="p-3 text-right font-mono text-rose-800 bg-rose-100/40 font-bold">
                      {servicesReport.summary.totalCost.toLocaleString('ru-RU', { minimumFractionDigits: 2, maximumFractionDigits: 2 })} ₽
                    </td>
                    <td className={`p-3 text-right font-mono font-extrabold bg-emerald-100/40 text-xs ${servicesReport.summary.totalProfit >= 0 ? 'text-emerald-800' : 'text-rose-700'}`}>
                      {servicesReport.summary.totalProfit.toLocaleString('ru-RU', { minimumFractionDigits: 2, maximumFractionDigits: 2 })} ₽
                    </td>
                    <td className="p-3 text-center pr-4 font-mono font-extrabold">
                      <span className={`px-2 py-0.5 rounded ${servicesReport.summary.totalProfitability >= 0 ? 'bg-emerald-100 text-emerald-900' : 'bg-rose-100 text-rose-900'}`}>
                        {servicesReport.summary.totalProfitability.toFixed(1)}%
                      </span>
                    </td>
                  </tr>
                </tfoot>
              </table>
            </div>
          )}
        </div>
      </div>
    )}

      {/* Modal: Hall Overrides Management */}
      {showHallOverridesModal && (
        <div className="fixed inset-0 z-50 bg-black/50 backdrop-blur-xs flex items-center justify-center p-4">
          <div className="bg-white rounded-3xl max-w-2xl w-full border border-gray-200 shadow-2xl overflow-hidden flex flex-col max-h-[85vh]">
            {/* Modal Header */}
            <div className="p-5 bg-gray-50 border-b border-gray-150 flex items-center justify-between">
              <div className="flex items-center gap-2.5">
                <div className="p-2.5 bg-indigo-100 text-indigo-700 rounded-xl">
                  <Building2 className="w-5 h-5" />
                </div>
                <div>
                  <h3 className="font-display font-extrabold text-sm text-gray-900">
                    Соответствия услуг залам салона ({Object.keys(serviceHallOverrides).length})
                  </h3>
                  <p className="text-[11px] text-gray-500 mt-0.5">
                    Все сделанные выборы сохраняются в браузере и используются при формировании всех отчетов
                  </p>
                </div>
              </div>
              <button
                onClick={() => setShowHallOverridesModal(false)}
                className="p-1.5 text-gray-400 hover:text-gray-700 rounded-lg hover:bg-gray-200 transition-colors cursor-pointer"
              >
                <X className="w-5 h-5" />
              </button>
            </div>

            {/* Modal Body */}
            <div className="p-5 overflow-y-auto space-y-3 flex-1">
              {Object.keys(serviceHallOverrides).length === 0 ? (
                <div className="text-center py-12 text-gray-400 text-xs space-y-2">
                  <Building2 className="w-8 h-8 text-gray-300 mx-auto" />
                  <p className="font-semibold text-gray-600">Нет сохраненных ручных назначений</p>
                  <p className="text-[11px] max-w-sm mx-auto">
                    Когда вы относите услугу из «Прочие услуги» к одному из залов, она сохраняется здесь и автоматически учитывается при следующих расчетах.
                  </p>
                </div>
              ) : (
                <div className="divide-y divide-gray-100 border border-gray-150 rounded-2xl overflow-hidden shadow-2xs">
                  {Object.entries(serviceHallOverrides).map(([svcName, hall]) => {
                    const meta = HALL_METADATA[hall as HallName] || HALL_METADATA['Прочие услуги'];
                    return (
                      <div key={`override_item_${svcName}`} className="p-3.5 bg-white flex items-center justify-between gap-3 text-xs hover:bg-gray-50/50 transition-colors">
                        <div className="min-w-0 flex-1">
                          <div className="font-semibold text-gray-900 truncate" title={svcName}>
                            {svcName}
                          </div>
                          <div className="text-[10px] text-gray-400 mt-0.5">
                            Текущий зал: <span className={`font-semibold ${meta.badgeText}`}>{hall}</span>
                          </div>
                        </div>

                        <div className="flex items-center gap-2 shrink-0">
                          <select
                            value={hall}
                            onChange={(e) => handleAssignServiceToHall(svcName, e.target.value as HallName)}
                            className="text-xs font-semibold bg-gray-50 border border-gray-250 rounded-xl px-2.5 py-1.5 text-gray-800 outline-none cursor-pointer focus:ring-2 focus:ring-indigo-300"
                          >
                            {HALL_NAMES.filter(h => h !== 'Прочие услуги').map((hName) => (
                              <option key={hName} value={hName}>
                                {hName}
                              </option>
                            ))}
                          </select>

                          <button
                            onClick={() => handleResetServiceHall(svcName)}
                            className="p-1.5 text-gray-400 hover:text-rose-600 hover:bg-rose-50 rounded-lg transition-colors cursor-pointer"
                            title="Сбросить это назначение"
                          >
                            <Trash2 className="w-4 h-4" />
                          </button>
                        </div>
                      </div>
                    );
                  })}
                </div>
              )}
            </div>

            {/* Modal Footer */}
            <div className="p-4 bg-gray-50 border-t border-gray-150 flex items-center justify-between">
              {Object.keys(serviceHallOverrides).length > 0 ? (
                <button
                  onClick={handleClearAllHallOverrides}
                  className="px-3.5 py-1.5 text-rose-600 hover:text-rose-800 hover:bg-rose-50 rounded-xl text-xs font-semibold transition-colors cursor-pointer"
                >
                  Сбросить все назначения
                </button>
              ) : <div />}
              <button
                onClick={() => setShowHallOverridesModal(false)}
                className="px-4 py-2 bg-indigo-600 hover:bg-indigo-700 text-white rounded-xl text-xs font-bold transition-all cursor-pointer shadow-sm"
              >
                Закрыть
              </button>
            </div>
          </div>
        </div>
      )}
      {/* Modal: Detailed User Guide */}
      <UserGuideModal
        isOpen={showUserGuideModal}
        onClose={() => setShowUserGuideModal(false)}
        onOpenDemo={loadDemoData}
      />
      </main>
    </div>
  );
}
