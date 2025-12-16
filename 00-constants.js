// 00-constants.js — общие константы проекта
// Перенесено из Внесение.js для улучшения структуры кода.

const MOSCOW_TZ = 'Europe/Moscow';

// Цвета и стили
const COLOR_BG_FULL_GREEN  = '#C6E0B4'; // светло-зелёный фон "закрыто"
const COLOR_FONT_DARKGREEN = '#385723'; // тёмно-зелёный текст

const WALLET_COLORS = {
  'Р/С Строймат': '#2496dd',
  'Р/С Брендмар': '#EABB3D',
  'Наличные':     '#0dac50',
  'Карта':        '#17ddee',
  'Карта Артема': '#E6E0EC',
  'Карта Паши':   '#E6E0EC',
  'ИП Паши':      '#D9D9D9'
};

// Диапазоны листов
const SHT_IN        = '⏬ ВНЕСЕНИЕ';
const SHT_PROV      = '☑️ ПРОВОДКИ';
const SHT_ACTS      = 'РЕЕСТР АКТОВ';
const SHT_DICT      = 'Справочник';

// Входной блок B10:F40
const IN_START_ROW  = 10;
const IN_END_ROW    = 40;
const IN_HEIGHT     = IN_END_ROW - IN_START_ROW + 1; // 31
const IN_COL_B      = 2; // B
const IN_COL_F      = 6; // F
const IN_COL_J      = 10; // J

const DATE_FORMAT   = 'dd.MM.yyyy';
const NBSP_RE       = /\u00A0/g; // non-breaking space
