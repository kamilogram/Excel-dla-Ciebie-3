import { CellData, CellStyle } from '../types/excel';
import { colIndexToLetter, letterToColIndex, parseCellAddress } from './formulaEngine';

export interface SliceCell {
  cellId: string;
  raw: string;
  computed?: any;
  style?: CellStyle;
}

// -------------------------------------------------------------
// Predefined Known Series Lists
// -------------------------------------------------------------

const POLISH_DAYS_FULL = [
  'poniedziałek', 'wtorek', 'środa', 'czwartek', 'piątek', 'sobota', 'niedziela'
];
const POLISH_DAYS_FULL_NODIA = [
  'poniedzialek', 'wtorek', 'sroda', 'czwartek', 'piatek', 'sobota', 'niedziela'
];
const POLISH_DAYS_CANONICAL = [
  'Poniedziałek', 'Wtorek', 'Środa', 'Czwartek', 'Piątek', 'Sobota', 'Niedziela'
];

const POLISH_DAYS_SHORT = ['pn', 'wt', 'śr', 'cz', 'pt', 'sb', 'nd'];
const POLISH_DAYS_SHORT_NODIA = ['pn', 'wt', 'sr', 'cz', 'pt', 'sb', 'nd'];
const POLISH_DAYS_SHORT_CANONICAL = ['Pn', 'Wt', 'Śr', 'Cz', 'Pt', 'Sb', 'Nd'];

const POLISH_DAYS_SHORT3 = ['pon', 'wto', 'śro', 'czw', 'pią', 'sob', 'nie'];
const POLISH_DAYS_SHORT3_CANONICAL = ['Pon', 'Wto', 'Śro', 'Czw', 'Pią', 'Sob', 'Nie'];

const ENGLISH_DAYS_FULL = [
  'monday', 'tuesday', 'wednesday', 'thursday', 'friday', 'saturday', 'sunday'
];
const ENGLISH_DAYS_CANONICAL = [
  'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday'
];

const ENGLISH_DAYS_SHORT = ['mon', 'tue', 'wed', 'thu', 'fri', 'sat', 'sun'];
const ENGLISH_DAYS_SHORT_CANONICAL = ['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun'];

const POLISH_MONTHS_FULL = [
  'styczeń', 'luty', 'marzec', 'kwiecień', 'maj', 'czerwiec',
  'lipiec', 'sierpień', 'wrzesień', 'październik', 'listopad', 'grudzień'
];
const POLISH_MONTHS_FULL_NODIA = [
  'styczen', 'luty', 'marzec', 'kwiecien', 'maj', 'czerwiec',
  'lipiec', 'sierpien', 'wrzesien', 'pazdziernik', 'listopad', 'grudzien'
];
const POLISH_MONTHS_CANONICAL = [
  'Styczeń', 'Luty', 'Marzec', 'Kwiecień', 'Maj', 'Czerwiec',
  'Lipiec', 'Sierpień', 'Wrzesień', 'Październik', 'Listopad', 'Grudzień'
];

const POLISH_MONTHS_SHORT = [
  'sty', 'lut', 'mar', 'kwi', 'maj', 'cze', 'lip', 'sie', 'wrz', 'paź', 'lis', 'gru'
];
const POLISH_MONTHS_SHORT_NODIA = [
  'sty', 'lut', 'mar', 'kwi', 'maj', 'cze', 'lip', 'sie', 'wrz', 'paz', 'lis', 'gru'
];
const POLISH_MONTHS_SHORT_CANONICAL = [
  'Sty', 'Lut', 'Mar', 'Kwi', 'Maj', 'Cze', 'Lip', 'Sie', 'Wrz', 'Paź', 'Lis', 'Gru'
];

const ENGLISH_MONTHS_FULL = [
  'january', 'february', 'march', 'april', 'may', 'june',
  'july', 'august', 'september', 'october', 'november', 'december'
];
const ENGLISH_MONTHS_CANONICAL = [
  'January', 'February', 'March', 'April', 'May', 'June',
  'July', 'August', 'September', 'October', 'November', 'December'
];

const ENGLISH_MONTHS_SHORT = [
  'jan', 'feb', 'mar', 'apr', 'may', 'jun', 'jul', 'aug', 'sep', 'oct', 'nov', 'dec'
];
const ENGLISH_MONTHS_SHORT_CANONICAL = [
  'Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'
];

// Helper: match text casing style
function applyCasing(sample: string, target: string): string {
  if (sample === sample.toUpperCase()) {
    return target.toUpperCase();
  }
  if (sample === sample.toLowerCase()) {
    return target.toLowerCase();
  }
  // Capitalized: first letter uppercase, rest lowercase
  if (sample[0] === sample[0].toUpperCase()) {
    return target.charAt(0).toUpperCase() + target.slice(1).toLowerCase();
  }
  return target;
}

// -------------------------------------------------------------
// Formula Reference Shifter
// -------------------------------------------------------------
export function shiftFormula(formula: string, dCol: number, dRow: number): string {
  if (!formula.startsWith('=')) return formula;

  // Replace cell references, respecting $ locks
  return formula.replace(/(\$?)([A-Za-z]+)(\$?)(\d+)/g, (match, colLock, colLetters, rowLock, rowNumStr) => {
    let finalCol = colLetters.toUpperCase();
    let finalRow = parseInt(rowNumStr, 10);

    if (!colLock) {
      const origColIdx = letterToColIndex(finalCol);
      const newColIdx = Math.max(0, origColIdx + dCol);
      finalCol = colIndexToLetter(newColIdx);
    }

    if (!rowLock) {
      finalRow = Math.max(1, finalRow + dRow);
    }

    return `${colLock}${finalCol}${rowLock}${finalRow}`;
  });
}

// -------------------------------------------------------------
// Date Parsing & Stepping
// -------------------------------------------------------------
interface ParsedDate {
  date: Date;
  format: 'ISO' | 'DOT' | 'SLASH' | 'SLASH_US';
  separator: string;
  hasLeadingZeroDay: boolean;
  hasLeadingZeroMonth: boolean;
}

function parseDate(str: string): ParsedDate | null {
  const trimmed = str.trim();

  // YYYY-MM-DD
  const isoMatch = trimmed.match(/^(\d{4})([-/])(\d{1,2})\2(\d{1,2})$/);
  if (isoMatch) {
    const y = parseInt(isoMatch[1], 10);
    const m = parseInt(isoMatch[3], 10);
    const d = parseInt(isoMatch[4], 10);
    const date = new Date(y, m - 1, d);
    if (!isNaN(date.getTime()) && date.getDate() === d) {
      return {
        date,
        format: 'ISO',
        separator: isoMatch[2],
        hasLeadingZeroDay: isoMatch[4].length === 2,
        hasLeadingZeroMonth: isoMatch[3].length === 2,
      };
    }
  }

  // DD.MM.YYYY
  const dotMatch = trimmed.match(/^(\d{1,2})\.(\d{1,2})\.(\d{4})$/);
  if (dotMatch) {
    const d = parseInt(dotMatch[1], 10);
    const m = parseInt(dotMatch[2], 10);
    const y = parseInt(dotMatch[3], 10);
    const date = new Date(y, m - 1, d);
    if (!isNaN(date.getTime()) && date.getDate() === d) {
      return {
        date,
        format: 'DOT',
        separator: '.',
        hasLeadingZeroDay: dotMatch[1].length === 2,
        hasLeadingZeroMonth: dotMatch[2].length === 2,
      };
    }
  }

  // DD/MM/YYYY
  const slashMatch = trimmed.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (slashMatch) {
    const d = parseInt(slashMatch[1], 10);
    const m = parseInt(slashMatch[2], 10);
    const y = parseInt(slashMatch[3], 10);
    const date = new Date(y, m - 1, d);
    if (!isNaN(date.getTime()) && date.getDate() === d) {
      return {
        date,
        format: 'SLASH',
        separator: '/',
        hasLeadingZeroDay: slashMatch[1].length === 2,
        hasLeadingZeroMonth: slashMatch[2].length === 2,
      };
    }
  }

  return null;
}

function formatDate(date: Date, parsed: ParsedDate): string {
  const y = date.getFullYear();
  const m = date.getMonth() + 1;
  const d = date.getDate();

  const mStr = parsed.hasLeadingZeroMonth ? String(m).padStart(2, '0') : String(m);
  const dStr = parsed.hasLeadingZeroDay ? String(d).padStart(2, '0') : String(d);

  if (parsed.format === 'ISO') {
    return `${y}${parsed.separator}${mStr}${parsed.separator}${dStr}`;
  }
  if (parsed.format === 'DOT') {
    return `${dStr}.${mStr}.${y}`;
  }
  return `${dStr}/${mStr}/${y}`;
}

// -------------------------------------------------------------
// Alphanumeric with ending numbers: e.g. "Pozycja 1", "Dzień 01"
// -------------------------------------------------------------
interface AlphaNumericMatch {
  prefix: string;
  num: number;
  digitsLen: number;
}

function parseAlphaNumeric(str: string): AlphaNumericMatch | null {
  const match = str.match(/^(.*?)(\d+)$/);
  if (!match) return null;
  return {
    prefix: match[1],
    num: parseInt(match[2], 10),
    digitsLen: match[2].length,
  };
}

// -------------------------------------------------------------
// Quarters (Kwartały)
// -------------------------------------------------------------
function parseQuarter(str: string): { prefix: string; q: number; year?: number } | null {
  const trimmed = str.trim();
  // e.g. Q1, Q2 or K1, K2 or Kwartał 1, Kwartał 2 or Q1 2024
  const m1 = trimmed.match(/^([QqKk]|Kwartał\s+|Quarter\s+)([1-4])(?:\s+(\d{4}))?$/i);
  if (m1) {
    return {
      prefix: m1[1],
      q: parseInt(m1[2], 10),
      year: m1[3] ? parseInt(m1[3], 10) : undefined,
    };
  }
  return null;
}

// -------------------------------------------------------------
// Roman Numerals (I to L)
// -------------------------------------------------------------
const ROMAN_NUMERALS = [
  'I', 'II', 'III', 'IV', 'V', 'VI', 'VII', 'VIII', 'IX', 'X',
  'XI', 'XII', 'XIII', 'XIV', 'XV', 'XVI', 'XVII', 'XVIII', 'XIX', 'XX',
  'XXI', 'XXII', 'XXIII', 'XXIV', 'XXV', 'XXVI', 'XXVII', 'XXVIII', 'XXIX', 'XXX'
];

function parseRoman(str: string): number | null {
  const upper = str.trim().toUpperCase();
  const idx = ROMAN_NUMERALS.indexOf(upper);
  return idx !== -1 ? idx + 1 : null;
}

// -------------------------------------------------------------
// Core 1D Sequence Generator
// Given a slice of source cells in sequence order:
// Predict values for extension indices (stepOffset = 1, 2, 3...)
// -------------------------------------------------------------
export function predictNextValue(
  sourceSlice: SliceCell[],
  stepOffset: number, // positive for forward (down/right), negative for backward (up/left)
  dCol: number,
  dRow: number
): { raw: string; style?: CellStyle } {
  const n = sourceSlice.length;
  if (n === 0) return { raw: '' };

  const firstCell = sourceSlice[0];
  const lastCell = sourceSlice[n - 1];

  // Pattern cell to copy style from
  const styleSourceIdx = stepOffset > 0
    ? (n - 1 + stepOffset) % n
    : (((stepOffset % n) + n) % n);
  const matchedStyle = sourceSlice[styleSourceIdx]?.style;

  // 1. FORMULAS
  // If all cells (or first cell) are formulas, shift references!
  if (firstCell.raw.startsWith('=')) {
    // If multiple source formulas, pick corresponding cyclically and shift
    const baseCell = stepOffset > 0
      ? sourceSlice[(stepOffset - 1) % n]
      : sourceSlice[((stepOffset % n) + n) % n];
    const newFormula = shiftFormula(baseCell.raw, dCol, dRow);
    return { raw: newFormula, style: matchedStyle };
  }

  const rawValues = sourceSlice.map((c) => c.raw.trim());

  // 2. DAYS OF WEEK (Polish & English, Full & Short)
  const daySets = [
    { list: POLISH_DAYS_FULL, canon: POLISH_DAYS_CANONICAL },
    { list: POLISH_DAYS_FULL_NODIA, canon: POLISH_DAYS_CANONICAL },
    { list: POLISH_DAYS_SHORT, canon: POLISH_DAYS_SHORT_CANONICAL },
    { list: POLISH_DAYS_SHORT_NODIA, canon: POLISH_DAYS_SHORT_CANONICAL },
    { list: POLISH_DAYS_SHORT3, canon: POLISH_DAYS_SHORT3_CANONICAL },
    { list: ENGLISH_DAYS_FULL, canon: ENGLISH_DAYS_CANONICAL },
    { list: ENGLISH_DAYS_SHORT, canon: ENGLISH_DAYS_SHORT_CANONICAL },
  ];

  for (const daySet of daySets) {
    const indices = rawValues.map((v) => daySet.list.indexOf(v.toLowerCase()));
    if (indices.every((idx) => idx !== -1)) {
      let step = 1;
      if (n > 1) {
        // Calculate step between elements modulo 7
        const diff = (indices[n - 1] - indices[0] + 7) % 7;
        step = Math.round(diff / (n - 1)) || 1;
      }
      const baseIdx = stepOffset > 0 ? indices[n - 1] : indices[0];
      const nextIdx = (baseIdx + stepOffset * step + 70000) % 7;
      const targetCanon = daySet.canon[nextIdx];
      const resultText = applyCasing(rawValues[0], targetCanon);
      return { raw: resultText, style: matchedStyle };
    }
  }

  // 3. MONTHS OF YEAR (Polish & English, Full & Short)
  const monthSets = [
    { list: POLISH_MONTHS_FULL, canon: POLISH_MONTHS_CANONICAL },
    { list: POLISH_MONTHS_FULL_NODIA, canon: POLISH_MONTHS_CANONICAL },
    { list: POLISH_MONTHS_SHORT, canon: POLISH_MONTHS_SHORT_CANONICAL },
    { list: POLISH_MONTHS_SHORT_NODIA, canon: POLISH_MONTHS_SHORT_CANONICAL },
    { list: ENGLISH_MONTHS_FULL, canon: ENGLISH_MONTHS_CANONICAL },
    { list: ENGLISH_MONTHS_SHORT, canon: ENGLISH_MONTHS_SHORT_CANONICAL },
  ];

  for (const monthSet of monthSets) {
    const indices = rawValues.map((v) => monthSet.list.indexOf(v.toLowerCase()));
    if (indices.every((idx) => idx !== -1)) {
      let step = 1;
      if (n > 1) {
        const diff = (indices[n - 1] - indices[0] + 12) % 12;
        step = Math.round(diff / (n - 1)) || 1;
      }
      const baseIdx = stepOffset > 0 ? indices[n - 1] : indices[0];
      const nextIdx = (baseIdx + stepOffset * step + 120000) % 12;
      const targetCanon = monthSet.canon[nextIdx];
      const resultText = applyCasing(rawValues[0], targetCanon);
      return { raw: resultText, style: matchedStyle };
    }
  }

  // 4. DATES (ISO, DOT, SLASH)
  const parsedDates = rawValues.map((v) => parseDate(v));
  if (parsedDates.every((d): d is ParsedDate => d !== null)) {
    let dayStep = 1;
    if (n > 1) {
      const msDiff = parsedDates[n - 1].date.getTime() - parsedDates[0].date.getTime();
      const daysDiff = Math.round(msDiff / (1000 * 60 * 60 * 24));
      dayStep = Math.round(daysDiff / (n - 1)) || 1;
    }
    const baseDateObj = stepOffset > 0 ? parsedDates[n - 1] : parsedDates[0];
    const newDate = new Date(baseDateObj.date);
    newDate.setDate(newDate.getDate() + stepOffset * dayStep);
    return {
      raw: formatDate(newDate, baseDateObj),
      style: matchedStyle,
    };
  }

  // 5. QUARTERS (Q1..Q4, K1..K4, Kwartał 1..4)
  const quarters = rawValues.map((v) => parseQuarter(v));
  if (quarters.every((q): q is NonNullable<typeof q> => q !== null)) {
    const samePrefix = quarters.every((q) => q.prefix.toLowerCase() === quarters[0].prefix.toLowerCase());
    if (samePrefix) {
      let qStep = 1;
      if (n > 1) {
        qStep = Math.round((quarters[n - 1].q - quarters[0].q) / (n - 1)) || 1;
      }
      const baseQ = stepOffset > 0 ? quarters[n - 1] : quarters[0];
      let totalQ = (baseQ.year ? baseQ.year * 4 : 0) + (baseQ.q - 1) + stepOffset * qStep;
      let nextYear: number | undefined;
      if (baseQ.year) {
        nextYear = Math.floor(totalQ / 4);
      }
      let nextQ = ((((totalQ % 4) + 4) % 4) || 0) + 1;
      let res = `${baseQ.prefix}${nextQ}`;
      if (nextYear !== undefined) {
        res += ` ${nextYear}`;
      }
      return { raw: res, style: matchedStyle };
    }
  }

  // 6. NUMBERS & ARITHMETIC PROGRESSION (e.g. 3, 6, 9 -> 12, 15, 18 or 1 -> 2, 3)
  const numbers = rawValues.map((v) => {
    // Check if trimmed value is a valid number
    if (v === '') return NaN;
    // Handle comma as decimal separator in Polish locale if applicable
    const normalized = v.replace(',', '.');
    const num = Number(normalized);
    return isNaN(num) ? NaN : num;
  });

  if (numbers.every((num) => !isNaN(num))) {
    let step = 1;
    const isIntegerList = numbers.every((num) => Number.isInteger(num));
    const usesComma = rawValues.some((v) => v.includes(',') && !v.includes('.'));

    if (n > 1) {
      // Step between numbers in the series
      step = (numbers[n - 1] - numbers[0]) / (n - 1);
    }

    const baseVal = stepOffset > 0 ? numbers[n - 1] : numbers[0];
    const nextVal = baseVal + stepOffset * step;

    // Formatting: if integer, show integer without decimal point; otherwise preserve decimal precision
    let formatted: string;
    if (isIntegerList && Number.isInteger(step)) {
      formatted = String(Math.round(nextVal));
    } else {
      // Find maximum precision in source numbers
      let maxDecimals = 2;
      rawValues.forEach((v) => {
        const parts = v.split(/[.,]/);
        if (parts.length === 2) {
          maxDecimals = Math.max(maxDecimals, parts[1].length);
        }
      });
      formatted = nextVal.toFixed(maxDecimals);
      // Clean trailing zeros if appropriate
      if (formatted.includes('.')) {
        formatted = formatted.replace(/\.?0+$/, '');
      }
    }

    if (usesComma) {
      formatted = formatted.replace('.', ',');
    }

    return { raw: formatted, style: matchedStyle };
  }

  // 7. ALPHANUMERIC PATTERNS (Text ending with numbers: e.g. "Dzień 1", "Pozycja 01", "Projekt 2")
  const alphaMatches = rawValues.map((v) => parseAlphaNumeric(v));
  if (alphaMatches.every((m): m is AlphaNumericMatch => m !== null)) {
    const prefix0 = alphaMatches[0].prefix;
    const samePrefix = alphaMatches.every((m) => m.prefix === prefix0);
    if (samePrefix) {
      let step = 1;
      if (n > 1) {
        step = Math.round((alphaMatches[n - 1].num - alphaMatches[0].num) / (n - 1)) || 1;
      }
      const baseMatch = stepOffset > 0 ? alphaMatches[n - 1] : alphaMatches[0];
      const nextNum = baseMatch.num + stepOffset * step;
      const paddedNum = String(Math.max(0, nextNum)).padStart(baseMatch.digitsLen, '0');
      return {
        raw: `${prefix0}${paddedNum}`,
        style: matchedStyle,
      };
    }
  }

  // 8. ROMAN NUMERALS (I, II, III... or I -> II, III)
  const romans = rawValues.map((v) => parseRoman(v));
  if (romans.every((r): r is number => r !== null)) {
    let step = 1;
    if (n > 1) {
      step = Math.round((romans[n - 1] - romans[0]) / (n - 1)) || 1;
    }
    const baseRoman = stepOffset > 0 ? romans[n - 1] : romans[0];
    const nextIdx = baseRoman + stepOffset * step - 1;
    if (nextIdx >= 0 && nextIdx < ROMAN_NUMERALS.length) {
      const isLower = rawValues[0] === rawValues[0].toLowerCase();
      const resultRoman = ROMAN_NUMERALS[nextIdx];
      return {
        raw: isLower ? resultRoman.toLowerCase() : resultRoman,
        style: matchedStyle,
      };
    }
  }

  // 9. CYCLIC REPEAT (Excel default for non-sequential text patterns)
  // E.g. ['Tak', 'Nie'] -> ['Tak', 'Nie', 'Tak', 'Nie']
  // E.g. ['A', 'B', 'C'] -> ['A', 'B', 'C', 'A', 'B']
  const cycleIndex = stepOffset > 0
    ? (stepOffset - 1) % n
    : (((stepOffset % n) + n) % n);
  const src = sourceSlice[cycleIndex];
  return {
    raw: src?.raw ?? '',
    style: src?.style,
  };
}
