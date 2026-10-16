import { CellData, CellStyle, CellBorders, BorderSide, BorderActionType, BorderLineStyle, SelectionRange } from '../types/excel';
import { letterToColIndex, colIndexToLetter, parseCellAddress } from './formulaEngine';

export const DEFAULT_BORDER_COLOR = '#000000';
export const DEFAULT_BORDER_STYLE: BorderLineStyle = 'solid';
export const DEFAULT_BORDER_WIDTH = 1;

/**
 * Converts a BorderSide object to a CSS border shorthand string e.g. "2px solid #000000"
 */
export function borderSideToCss(side?: BorderSide | null): string | undefined {
  if (!side) return undefined;
  const width = side.width ?? DEFAULT_BORDER_WIDTH;
  const style = side.style ?? DEFAULT_BORDER_STYLE;
  const color = side.color ?? DEFAULT_BORDER_COLOR;
  return `${width}px ${style} ${color}`;
}

/**
 * Returns computed CSS border values (top, bottom, left, right) for a cell's style.
 * Supports both rich `borders` object and legacy `border` strings.
 */
export function getCellBorderCss(
  style?: CellStyle,
  isDarkMode: boolean = false
): {
  top?: string;
  bottom?: string;
  left?: string;
  right?: string;
  hasBorders: boolean;
} {
  if (!style) return { hasBorders: false };

  const defaultDarkColor = '#9ca3af';
  const defaultLightColor = '#374151';
  const fallbackColor = isDarkMode ? defaultDarkColor : defaultLightColor;

  // 1. Rich `borders` property
  if (style.borders) {
    const top = borderSideToCss(style.borders.top);
    const bottom = borderSideToCss(style.borders.bottom);
    const left = borderSideToCss(style.borders.left);
    const right = borderSideToCss(style.borders.right);
    const hasBorders = Boolean(top || bottom || left || right);
    return { top, bottom, left, right, hasBorders };
  }

  // 2. Legacy `border` string
  if (style.border && style.border !== 'none') {
    if (style.border === 'all') {
      const b = `1px solid ${fallbackColor}`;
      return { top: b, bottom: b, left: b, right: b, hasBorders: true };
    }
    if (style.border === 'outside' || style.border === 'thick-outside') {
      const b = `2px solid ${fallbackColor}`;
      return { top: b, bottom: b, left: b, right: b, hasBorders: true };
    }
    if (style.border === 'thick-bottom' || style.border === 'bottom-thick') {
      return { bottom: `2px solid ${fallbackColor}`, hasBorders: true };
    }
    if (style.border === 'bottom-double') {
      return { bottom: `3px double ${fallbackColor}`, hasBorders: true };
    }
    if (style.border === 'bottom') {
      return { bottom: `1px solid ${fallbackColor}`, hasBorders: true };
    }
    if (style.border === 'top') {
      return { top: `1px solid ${fallbackColor}`, hasBorders: true };
    }
    if (style.border === 'left') {
      return { left: `1px solid ${fallbackColor}`, hasBorders: true };
    }
    if (style.border === 'right') {
      return { right: `1px solid ${fallbackColor}`, hasBorders: true };
    }
    if (style.border === 'top-bottom') {
      const b = `1px solid ${fallbackColor}`;
      return { top: b, bottom: b, hasBorders: true };
    }
    if (style.border === 'top-bottom-double') {
      return { top: `1px solid ${fallbackColor}`, bottom: `3px double ${fallbackColor}`, hasBorders: true };
    }
    if (style.border === 'top-bottom-thick') {
      return { top: `1px solid ${fallbackColor}`, bottom: `2px solid ${fallbackColor}`, hasBorders: true };
    }
  }

  return { hasBorders: false };
}

export interface ApplyBorderParams {
  action: BorderActionType;
  range: SelectionRange | string; // range or single cell e.g. "B2"
  lineStyle?: BorderLineStyle;
  lineWidth?: number;
  lineColor?: string;
}

/**
 * Applies a border action across a range or single cell in sheet cells map,
 * returning a new record of updated cells.
 */
export function applyBordersToCells(
  currentCells: Record<string, CellData>,
  params: ApplyBorderParams
): Record<string, CellData> {
  const { action, range, lineStyle = 'solid', lineWidth = 1, lineColor = DEFAULT_BORDER_COLOR } = params;

  let minCol: number;
  let maxCol: number;
  let minRow: number;
  let maxRow: number;

  if (typeof range === 'string') {
    const coord = parseCellAddress(range);
    if (!coord) return currentCells;
    const c = letterToColIndex(coord.col);
    minCol = c;
    maxCol = c;
    minRow = coord.row;
    maxRow = coord.row;
  } else {
    const c1 = letterToColIndex(range.startCol);
    const c2 = letterToColIndex(range.endCol);
    minCol = Math.min(c1, c2);
    maxCol = Math.max(c1, c2);
    minRow = Math.min(range.startRow, range.endRow);
    maxRow = Math.max(range.startRow, range.endRow);
  }

  const updatedCells: Record<string, CellData> = { ...currentCells };

  const standardSide: BorderSide = {
    style: lineStyle,
    width: lineWidth,
    color: lineColor,
  };

  const thickSide: BorderSide = {
    style: 'solid',
    width: Math.max(2, lineWidth + 1),
    color: lineColor,
  };

  const doubleSide: BorderSide = {
    style: 'double',
    width: 3,
    color: lineColor,
  };

  for (let r = minRow; r <= maxRow; r++) {
    for (let c = minCol; c <= maxCol; c++) {
      const cellId = `${colIndexToLetter(c)}${r}`;
      const existing = updatedCells[cellId] || { raw: '' };
      const existingBorders: CellBorders = existing.style?.borders ? { ...existing.style.borders } : {};

      const isTopRow = r === minRow;
      const isBottomRow = r === maxRow;
      const isLeftCol = c === minCol;
      const isRightCol = c === maxCol;

      let nextBorders: CellBorders | undefined = { ...existingBorders };

      switch (action) {
        case 'none':
          nextBorders = undefined;
          break;

        case 'all':
          nextBorders = {
            top: standardSide,
            bottom: standardSide,
            left: standardSide,
            right: standardSide,
          };
          break;

        case 'outside':
          if (isTopRow) nextBorders.top = standardSide;
          if (isBottomRow) nextBorders.bottom = standardSide;
          if (isLeftCol) nextBorders.left = standardSide;
          if (isRightCol) nextBorders.right = standardSide;
          break;

        case 'thick-outside':
          if (isTopRow) nextBorders.top = thickSide;
          if (isBottomRow) nextBorders.bottom = thickSide;
          if (isLeftCol) nextBorders.left = thickSide;
          if (isRightCol) nextBorders.right = thickSide;
          break;

        case 'bottom':
          if (isBottomRow) nextBorders.bottom = standardSide;
          break;

        case 'top':
          if (isTopRow) nextBorders.top = standardSide;
          break;

        case 'left':
          if (isLeftCol) nextBorders.left = standardSide;
          break;

        case 'right':
          if (isRightCol) nextBorders.right = standardSide;
          break;

        case 'bottom-double':
          if (isBottomRow) nextBorders.bottom = doubleSide;
          break;

        case 'bottom-thick':
          if (isBottomRow) nextBorders.bottom = thickSide;
          break;

        case 'top-bottom':
          if (isTopRow) nextBorders.top = standardSide;
          if (isBottomRow) nextBorders.bottom = standardSide;
          break;

        case 'top-bottom-double':
          if (isTopRow) nextBorders.top = standardSide;
          if (isBottomRow) nextBorders.bottom = doubleSide;
          break;

        case 'top-bottom-thick':
          if (isTopRow) nextBorders.top = standardSide;
          if (isBottomRow) nextBorders.bottom = thickSide;
          break;

        case 'inside-horizontal':
          if (r < maxRow) nextBorders.bottom = standardSide;
          break;

        case 'inside-vertical':
          if (c < maxCol) nextBorders.right = standardSide;
          break;
      }

      updatedCells[cellId] = {
        ...existing,
        style: {
          ...(existing.style || {}),
          border: action === 'none' ? undefined : action,
          borders: nextBorders,
        },
      };
    }
  }

  return updatedCells;
}
