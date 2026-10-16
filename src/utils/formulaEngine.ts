import { SheetData } from '../types/excel';

export function colIndexToLetter(index: number): string {
  let temp = index;
  let letter = '';
  while (temp >= 0) {
    letter = String.fromCharCode((temp % 26) + 65) + letter;
    temp = Math.floor(temp / 26) - 1;
  }
  return letter;
}

export function letterToColIndex(col: string): number {
  let index = 0;
  const clean = col.toUpperCase().replace(/\$/g, '');
  for (let i = 0; i < clean.length; i++) {
    index = index * 26 + (clean.charCodeAt(i) - 64);
  }
  return index - 1;
}

export function parseCellAddress(addr: string): { col: string; row: number } | null {
  const match = addr.toUpperCase().replace(/\$/g, '').match(/^([A-Z]+)(\d+)$/);
  if (!match) return null;
  return {
    col: match[1],
    row: parseInt(match[2], 10),
  };
}

export function getRangeCells(rangeStr: string): string[] {
  const parts = rangeStr.split(':');
  if (parts.length === 1) return [parts[0].trim().toUpperCase()];
  if (parts.length !== 2) return [];

  const start = parseCellAddress(parts[0].trim());
  const end = parseCellAddress(parts[1].trim());
  if (!start || !end) return [];

  const startColIdx = letterToColIndex(start.col);
  const endColIdx = letterToColIndex(end.col);
  const minCol = Math.min(startColIdx, endColIdx);
  const maxCol = Math.max(startColIdx, endColIdx);

  const minRow = Math.min(start.row, end.row);
  const maxRow = Math.max(start.row, end.row);

  const cells: string[] = [];
  for (let r = minRow; r <= maxRow; r++) {
    for (let c = minCol; c <= maxCol; c++) {
      cells.push(`${colIndexToLetter(c)}${r}`);
    }
  }
  return cells;
}

/**
 * Inserts or replaces a cell or range reference in a formula when pointing to cells in Excel.
 */
export function insertOrReplaceCellRef(
  currentFormula: string,
  newRef: string,
  hasActiveRef: boolean
): string {
  if (!currentFormula.startsWith('=')) {
    return `=${newRef}`;
  }

  if (currentFormula === '=') {
    return `=${newRef}`;
  }

  const lastChar = currentFormula.trimEnd().slice(-1);
  const operatorChars = ['=', '(', '+', '-', '*', '/', ',', ';', ':', '>', '<', '&'];

  if (operatorChars.includes(lastChar) || !hasActiveRef) {
    return currentFormula + newRef;
  }

  const trailingRefRegex = /([$]?[A-Z]+[$]?[0-9]+(?::[$]?[A-Z]+[$]?[0-9]+)?)$/i;
  if (trailingRefRegex.test(currentFormula)) {
    return currentFormula.replace(trailingRefRegex, newRef);
  }

  return currentFormula + newRef;
}

/**
 * Clean numeric string from Polish comma decimals (e.g. "12,50" -> 12.50)
 */
export function parseNumericValue(val: any): number {
  if (typeof val === 'number') return val;
  if (val === null || val === undefined || val === '') return 0;
  if (typeof val === 'string') {
    const cleaned = val.replace(/\s+/g, '').replace(',', '.');
    const num = parseFloat(cleaned);
    return isNaN(num) ? 0 : num;
  }
  return 0;
}

/**
 * Format a cell's computed value according to its style
 */
export function formatCellValue(value: any, style?: any): string {
  if (value === null || value === undefined) return '';
  if (typeof value === 'boolean') return value ? 'PRAWDA' : 'FAŁSZ';
  if (typeof value === 'string' && isNaN(Number(value)) && !style?.format) {
    return value;
  }

  const num = parseNumericValue(value);
  const format = style?.format || 'general';
  const decimals = style?.decimals !== undefined ? style.decimals : 2;

  switch (format) {
    case 'currency':
      return new Intl.NumberFormat('pl-PL', {
        style: 'currency',
        currency: 'PLN',
        minimumFractionDigits: decimals,
        maximumFractionDigits: decimals,
      }).format(num);

    case 'percent':
      return new Intl.NumberFormat('pl-PL', {
        style: 'percent',
        minimumFractionDigits: decimals,
        maximumFractionDigits: decimals,
      }).format(num > 1 ? num / 100 : num);

    case 'number':
      return new Intl.NumberFormat('pl-PL', {
        minimumFractionDigits: decimals,
        maximumFractionDigits: decimals,
      }).format(num);

    case 'date':
      if (typeof value === 'string') return value;
      try {
        return new Date(value).toLocaleDateString('pl-PL');
      } catch {
        return String(value);
      }

    default:
      if (typeof value === 'number') {
        // If whole number, format without decimals, else with 2 decimals
        return Number.isInteger(value)
          ? String(value)
          : value.toLocaleString('pl-PL', { maximumFractionDigits: 4 });
      }
      return String(value);
  }
}

/**
 * Evaluates Excel formula for a given sheet.
 */
export function evaluateFormula(
  formula: string,
  sheet: SheetData,
  visited: Set<string> = new Set()
): { result: any; error?: string } {
  if (!formula.startsWith('=')) {
    const trimmed = formula.trim();
    if (!isNaN(Number(trimmed)) && trimmed !== '') {
      return { result: Number(trimmed) };
    }
    // Check if Polish decimal format like "12,5"
    if (/^-?\d+,\d+$/.test(trimmed)) {
      return { result: parseFloat(trimmed.replace(',', '.')) };
    }
    return { result: formula };
  }

  let expr = formula.slice(1).trim();

  // Helper to get raw/computed value of cell
  const getCellValue = (cellId: string): any => {
    const cleanId = cellId.toUpperCase().replace(/\$/g, '');
    if (visited.has(cleanId)) {
      return '#CYKL!';
    }
    const cell = sheet.cells[cleanId];
    if (!cell) return 0;
    if (cell.raw.startsWith('=')) {
      const nextVisited = new Set(visited).add(cleanId);
      const evaled = evaluateFormula(cell.raw, sheet, nextVisited);
      return evaled.result;
    }
    const num = Number(cell.raw.replace(',', '.'));
    return isNaN(num) ? cell.raw : num;
  };

  try {
    // 1. Process Polish / English uppercase functions
    // Support nested functions recursively
    const evaluatedExpr = evaluateFunctions(expr, getCellValue, sheet);
    if (evaluatedExpr.error) return { result: evaluatedExpr.error, error: evaluatedExpr.error };

    // 2. Evaluate remaining cell references (like A1 + B2, or $C$3 * $K$2)
    const resolvedMath = evaluatedExpr.result.replace(/(?:\$?[A-Z]+\$?[0-9]+)\b/gi, (match: string) => {
      const val = getCellValue(match);
      if (typeof val === 'number') return String(val);
      if (typeof val === 'string') {
        const n = Number(val.replace(',', '.'));
        return isNaN(n) ? `"${val}"` : String(n);
      }
      return '0';
    });

    // 3. Evaluate safe math expression
    const finalVal = safeMathEval(resolvedMath);
    return { result: finalVal };
  } catch (err: any) {
    return { result: '#ARG!', error: err.message || 'Błąd formuły' };
  }
}

/**
 * Evaluates functions like SUMA, AVERAGE, JEŻELI, WYSZUKAJ.PIONOWO
 */
function evaluateFunctions(
  expr: string,
  getCellValue: (id: string) => any,
  sheet: SheetData
): { result: string; error?: string } {
  // Regex to match FunctionName( ... )
  // We match innermost functions or process token by token
  let current = expr;

  // Replace Polish and English function names (both with and without Polish diacritics)
  current = current
    .replace(/(?:^|(?<=[^a-zA-Z0-9_\p{L}.]))(?:ŚREDNIA|SREDNIA)(?=\s*\()/giu, 'AVERAGE')
    .replace(/(?:^|(?<=[^a-zA-Z0-9_\p{L}.]))(?:SUMA|SUM)(?=\s*\()/giu, 'SUM')
    .replace(/(?:^|(?<=[^a-zA-Z0-9_\p{L}.]))(?:ILE\.LICZB|ILE_LICZB|COUNT)(?=\s*\()/giu, 'COUNT')
    .replace(/(?:^|(?<=[^a-zA-Z0-9_\p{L}.]))(?:JEŻELI|JEZELI|IF)(?=\s*\()/giu, 'IF')
    .replace(/(?:^|(?<=[^a-zA-Z0-9_\p{L}.]))(?:WYSZUKAJ\.PIONOWO|WYSZUKAJ_PIONOWO|VLOOKUP)(?=\s*\()/giu, 'VLOOKUP')
    .replace(/(?:^|(?<=[^a-zA-Z0-9_\p{L}.]))(?:LICZ\.JEŻELI|LICZ\.JEZELI|COUNTIF)(?=\s*\()/giu, 'COUNTIF')
    .replace(/(?:^|(?<=[^a-zA-Z0-9_\p{L}.]))(?:SUMA\.JEŻELI|SUMA\.JEZELI|SUMIF)(?=\s*\()/giu, 'SUMIF')
    .replace(/(?:^|(?<=[^a-zA-Z0-9_\p{L}.]))(?:ZŁĄCZ\.TEKSTY|ZLACZ\.TEKSTY|CONCATENATE|CONCAT)(?=\s*\()/giu, 'CONCAT')
    .replace(/(?:^|(?<=[^a-zA-Z0-9_\p{L}.]))(?:ZAOKR|ROUND)(?=\s*\()/giu, 'ROUND')
    .replace(/(?:^|(?<=[^a-zA-Z0-9_\p{L}.]))(?:LITERY\.WIELKIE|LITERY_WIELKIE|UPPER)(?=\s*\()/giu, 'UPPER')
    .replace(/(?:^|(?<=[^a-zA-Z0-9_\p{L}.]))(?:LITERY\.MAŁE|LITERY\.MALE|LITERY_MALE|LOWER)(?=\s*\()/giu, 'LOWER')
    .replace(/(?:^|(?<=[^a-zA-Z0-9_\p{L}.]))(?:DŁ|DL|LEN)(?=\s*\()/giu, 'LEN')
    .replace(/(?:^|(?<=[^a-zA-Z0-9_\p{L}.]))(?:DZIŚ|DZIS|TODAY)(?=\s*\()/giu, 'TODAY');

  let loopLimit = 30;
  const fnRegex = /\b(SUM|AVERAGE|COUNT|MIN|MAX|IF|VLOOKUP|COUNTIF|SUMIF|CONCAT|ROUND|UPPER|LOWER|LEN|TODAY)\s*\(([^()]*)\)/i;

  while (fnRegex.test(current) && loopLimit > 0) {
    loopLimit--;
    current = current.replace(fnRegex, (_full, fnName, argsStr) => {
      const fn = fnName.toUpperCase();
      const rawArgs = splitFunctionArgs(argsStr);

      switch (fn) {
        case 'SUM': {
          let sum = 0;
          for (const arg of rawArgs) {
            if (arg.includes(':')) {
              const cells = getRangeCells(arg);
              for (const c of cells) {
                sum += parseNumericValue(getCellValue(c));
              }
            } else {
              sum += parseNumericValue(evaluateSimpleArg(arg, getCellValue));
            }
          }
          return String(sum);
        }

        case 'AVERAGE': {
          let sum = 0;
          let count = 0;
          for (const arg of rawArgs) {
            if (arg.includes(':')) {
              const cells = getRangeCells(arg);
              for (const c of cells) {
                const cellObj = sheet.cells[c];
                // In Excel, blank cells in a range are NOT counted towards the average
                if (!cellObj || cellObj.raw === '' || cellObj.raw === undefined) continue;
                const val = getCellValue(c);
                const num = parseNumericValue(val);
                if (typeof val === 'number' || (typeof val === 'string' && val.trim() !== '' && !isNaN(Number(val.replace(',', '.'))))) {
                  sum += num;
                  count++;
                }
              }
            } else {
              const val = evaluateSimpleArg(arg, getCellValue);
              const num = parseNumericValue(val);
              sum += num;
              count++;
            }
          }
          return count > 0 ? String(sum / count) : '#DZIEL/0!';
        }

        case 'COUNT': {
          let count = 0;
          for (const arg of rawArgs) {
            if (arg.includes(':')) {
              const cells = getRangeCells(arg);
              for (const c of cells) {
                const val = getCellValue(c);
                if (val !== '' && val !== null && !isNaN(Number(String(val).replace(',', '.')))) {
                  count++;
                }
              }
            } else {
              const val = evaluateSimpleArg(arg, getCellValue);
              if (!isNaN(Number(val))) count++;
            }
          }
          return String(count);
        }

        case 'MIN': {
          let min = Infinity;
          for (const arg of rawArgs) {
            if (arg.includes(':')) {
              const cells = getRangeCells(arg);
              for (const c of cells) {
                const cellObj = sheet.cells[c];
                if (!cellObj || cellObj.raw === '' || cellObj.raw === undefined) continue;
                min = Math.min(min, parseNumericValue(getCellValue(c)));
              }
            } else {
              min = Math.min(min, parseNumericValue(evaluateSimpleArg(arg, getCellValue)));
            }
          }
          return String(min === Infinity ? 0 : min);
        }

        case 'MAX': {
          let max = -Infinity;
          for (const arg of rawArgs) {
            if (arg.includes(':')) {
              const cells = getRangeCells(arg);
              for (const c of cells) {
                const cellObj = sheet.cells[c];
                if (!cellObj || cellObj.raw === '' || cellObj.raw === undefined) continue;
                max = Math.max(max, parseNumericValue(getCellValue(c)));
              }
            } else {
              max = Math.max(max, parseNumericValue(evaluateSimpleArg(arg, getCellValue)));
            }
          }
          return String(max === -Infinity ? 0 : max);
        }

        case 'IF': {
          if (rawArgs.length < 2) return '#ARG!';
          const condition = rawArgs[0];
          const valTrue = rawArgs[1] ?? '""';
          const valFalse = rawArgs[2] ?? '""';

          const conditionResult = evaluateCondition(condition, getCellValue);
          const chosen = conditionResult ? valTrue : valFalse;
          return String(evaluateSimpleArg(chosen, getCellValue));
        }

        case 'VLOOKUP': {
          // VLOOKUP(lookup_value, table_array, col_index, [range_lookup])
          if (rawArgs.length < 3) return '#N/D!';
          const lookupVal = evaluateSimpleArg(rawArgs[0], getCellValue);
          const tableRange = rawArgs[1].trim();
          const colIndex = parseInt(evaluateSimpleArg(rawArgs[2], getCellValue), 10);

          if (!tableRange.includes(':') || isNaN(colIndex) || colIndex < 1) {
            return '#ARG!';
          }

          const parts = tableRange.split(':');
          const start = parseCellAddress(parts[0]);
          const end = parseCellAddress(parts[1]);
          if (!start || !end) return '#ARG!';

          const startColIdx = letterToColIndex(start.col);
          const minRow = Math.min(start.row, end.row);
          const maxRow = Math.max(start.row, end.row);
          const targetColIdx = startColIdx + colIndex - 1;

          for (let r = minRow; r <= maxRow; r++) {
            const firstColCell = `${colIndexToLetter(startColIdx)}${r}`;
            const firstVal = getCellValue(firstColCell);

            // Compare case-insensitively or numerically
            const match = String(firstVal).trim().toLowerCase() === String(lookupVal).trim().toLowerCase();
            if (match) {
              const targetCell = `${colIndexToLetter(targetColIdx)}${r}`;
              return String(getCellValue(targetCell));
            }
          }
          return '#N/D!';
        }

        case 'COUNTIF': {
          if (rawArgs.length < 2) return '#ARG!';
          const range = rawArgs[0];
          const criteria = evaluateSimpleArg(rawArgs[1], getCellValue);
          const cells = getRangeCells(range);
          let count = 0;
          for (const c of cells) {
            const val = getCellValue(c);
            if (checkCriteria(val, criteria)) count++;
          }
          return String(count);
        }

        case 'SUMIF': {
          if (rawArgs.length < 2) return '#ARG!';
          const range = rawArgs[0];
          const criteria = evaluateSimpleArg(rawArgs[1], getCellValue);
          const sumRange = rawArgs[2] ? rawArgs[2] : range;

          const testCells = getRangeCells(range);
          const sumCells = getRangeCells(sumRange);

          let sum = 0;
          for (let i = 0; i < testCells.length; i++) {
            const testVal = getCellValue(testCells[i]);
            if (checkCriteria(testVal, criteria)) {
              const sumCell = sumCells[i] || testCells[i];
              sum += parseNumericValue(getCellValue(sumCell));
            }
          }
          return String(sum);
        }

        case 'CONCAT': {
          let res = '';
          for (const arg of rawArgs) {
            if (arg.includes(':')) {
              const cells = getRangeCells(arg);
              for (const c of cells) res += String(getCellValue(c));
            } else {
              res += String(evaluateSimpleArg(arg, getCellValue));
            }
          }
          return `"${res}"`;
        }

        case 'ROUND': {
          if (rawArgs.length < 1) return '#ARG!';
          const num = parseNumericValue(evaluateSimpleArg(rawArgs[0], getCellValue));
          const decimals = rawArgs[1] ? parseInt(evaluateSimpleArg(rawArgs[1], getCellValue), 10) : 0;
          const factor = Math.pow(10, decimals);
          return String(Math.round(num * factor) / factor);
        }

        case 'UPPER': {
          const txt = String(evaluateSimpleArg(rawArgs[0] || '', getCellValue));
          return `"${txt.toUpperCase()}"`;
        }

        case 'LOWER': {
          const txt = String(evaluateSimpleArg(rawArgs[0] || '', getCellValue));
          return `"${txt.toLowerCase()}"`;
        }

        case 'LEN': {
          const txt = String(evaluateSimpleArg(rawArgs[0] || '', getCellValue));
          return String(txt.length);
        }

        case 'TODAY': {
          return `"${new Date().toLocaleDateString('pl-PL')}"`;
        }

        default:
          return '#NAZWA?';
      }
    });
  }

  return { result: current };
}

function splitFunctionArgs(argsStr: string): string[] {
  const result: string[] = [];
  let depth = 0;
  let current = '';
  let inQuotes = false;

  for (let i = 0; i < argsStr.length; i++) {
    const ch = argsStr[i];
    if (ch === '"') {
      inQuotes = !inQuotes;
      current += ch;
    } else if (ch === '(' && !inQuotes) {
      depth++;
      current += ch;
    } else if (ch === ')' && !inQuotes) {
      depth--;
      current += ch;
    } else if ((ch === ';' || ch === ',') && depth === 0 && !inQuotes) {
      result.push(current.trim());
      current = '';
    } else {
      current += ch;
    }
  }
  if (current.trim()) {
    result.push(current.trim());
  }
  return result;
}

function evaluateSimpleArg(arg: string, getCellValue: (id: string) => any): any {
  const clean = arg.trim();
  if (clean.startsWith('"') && clean.endsWith('"')) {
    return clean.slice(1, -1);
  }
  if (parseCellAddress(clean)) {
    return getCellValue(clean);
  }
  const num = Number(clean.replace(',', '.'));
  if (!isNaN(num)) return num;
  return clean;
}

function evaluateCondition(condStr: string, getCellValue: (id: string) => any): boolean {
  // Support operators: >=, <=, <>, !=, =, >, <
  const operators = ['>=', '<=', '<>', '!=', '=', '>', '<'];
  for (const op of operators) {
    const idx = condStr.indexOf(op);
    if (idx !== -1) {
      const leftRaw = condStr.substring(0, idx).trim();
      const rightRaw = condStr.substring(idx + op.length).trim();
      const leftVal = evaluateSimpleArg(leftRaw, getCellValue);
      const rightVal = evaluateSimpleArg(rightRaw, getCellValue);

      const numLeft = Number(leftVal);
      const numRight = Number(rightVal);
      const isNum = !isNaN(numLeft) && !isNaN(numRight);

      switch (op) {
        case '>=':
          return isNum ? numLeft >= numRight : String(leftVal) >= String(rightVal);
        case '<=':
          return isNum ? numLeft <= numRight : String(leftVal) <= String(rightVal);
        case '<>':
        case '!=':
          return String(leftVal).toLowerCase() !== String(rightVal).toLowerCase();
        case '=':
          return String(leftVal).toLowerCase() === String(rightVal).toLowerCase();
        case '>':
          return isNum ? numLeft > numRight : String(leftVal) > String(rightVal);
        case '<':
          return isNum ? numLeft < numRight : String(leftVal) < String(rightVal);
      }
    }
  }

  // Boolean value
  const val = evaluateSimpleArg(condStr, getCellValue);
  return Boolean(val);
}

function checkCriteria(val: any, criteria: any): boolean {
  const critStr = String(criteria).trim();
  if (critStr.startsWith('>=')) {
    return parseNumericValue(val) >= parseNumericValue(critStr.slice(2));
  }
  if (critStr.startsWith('<=')) {
    return parseNumericValue(val) <= parseNumericValue(critStr.slice(2));
  }
  if (critStr.startsWith('>')) {
    return parseNumericValue(val) > parseNumericValue(critStr.slice(1));
  }
  if (critStr.startsWith('<')) {
    return parseNumericValue(val) < parseNumericValue(critStr.slice(1));
  }
  if (critStr.startsWith('<>') || critStr.startsWith('!=')) {
    return String(val).toLowerCase() !== critStr.slice(2).trim().toLowerCase();
  }
  return String(val).trim().toLowerCase() === critStr.toLowerCase();
}

/**
 * Safe evaluator for standard arithmetic expressions
 */
function safeMathEval(expr: string): any {
  const clean = expr.trim();
  if (clean.startsWith('"') && clean.endsWith('"')) {
    return clean.slice(1, -1);
  }

  // String concatenation with &
  if (clean.includes('&')) {
    const parts = clean.split('&');
    return parts.map(p => {
      const trimmed = p.trim();
      if (trimmed.startsWith('"') && trimmed.endsWith('"')) return trimmed.slice(1, -1);
      return trimmed;
    }).join('');
  }

  // Support percentage notation e.g. * 23% -> * (23 * 0.01)
  const percentNormalized = clean.replace(/(\d+(?:\.\d+)?)%/g, '($1*0.01)');

  // Remove any dangerous characters, only allow digits, operators, parens, decimal dot
  const sanitized = percentNormalized.replace(/[^0-9+\-*/().\s]/g, '');
  if (!sanitized) return clean;

  try {
    // eslint-disable-next-line no-new-func
    const result = Function(`"use strict"; return (${sanitized});`)();
    if (typeof result === 'number' && !isFinite(result)) return '#DZIEL/0!';
    return result;
  } catch {
    return clean;
  }
}
