import React, { useState, useRef, useEffect, useCallback, useMemo } from 'react';
import {
  Bold,
  Italic,
  Underline,
  Scissors,
  Copy,
  Clipboard,
  Trash2,
  Plus,
  Sigma,
  BarChart2,
  DollarSign,
  Percent,
  PaintBucket,
  Type,
  AlignLeft,
  AlignCenter,
  AlignRight,
  Sparkles,
  Paintbrush,
  Move,
  AlertTriangle,
} from 'lucide-react';
import {
  SheetData,
  SelectionRange,
  CellStyle,
} from '../types/excel';
import {
  colIndexToLetter,
  letterToColIndex,
  parseCellAddress,
  formatCellValue,
} from '../utils/formulaEngine';
import { getCellBorderCss } from '../utils/borderUtils';
import { sounds } from '../utils/audio';

interface SpreadsheetGridProps {
  sheet: SheetData;
  activeCellId: string;
  selectionRange: SelectionRange | null;
  onSelectCell: (cellId: string) => void;
  onSelectRange: (range: SelectionRange | null) => void;
  onCellChange: (cellId: string, rawValue: string) => void;
  onAutoFill: (sourceRange: SelectionRange, targetRange: SelectionRange, copyOnly?: boolean) => void;
  targetLessonCell?: string;
  showGridlines: boolean;
  showHeaders: boolean;
  zoomLevel: number;
  isDarkMode: boolean;
  onShortcutUsed?: (name: string) => void;
  onApplyStyle?: (stylePatch: Partial<CellStyle>) => void;
  onInsertRow?: () => void;
  onDeleteRow?: () => void;
  onAutoSum?: (formulaType: 'SUMA') => void;
  onOpenChartModal?: () => void;
  isFormulaMode?: boolean;
  formulaPointedRange?: SelectionRange | null;
  onPointCell?: (cellId: string) => void;
  onPointRange?: (range: SelectionRange) => void;
  onSubmitFormula?: (moveRight?: boolean) => void;
  onCancelFormula?: () => void;
  onResizeColumn?: (col: string, width: number) => void;
  onResizeRow?: (row: number, height: number) => void;
  onUndo?: () => void;
  onRedo?: () => void;
  onCellTyping?: (val: string) => void;
  isFormatPainterActive?: boolean;
  isFormatPainterLocked?: boolean;
  formatPainterStyle?: CellStyle | null;
  onToggleFormatPainter?: (lock?: boolean) => void;
  onApplyFormatPainter?: (targetRange?: SelectionRange | string) => void;
  onCancelFormatPainter?: () => void;
  onClearSelection?: () => void;
  onCopy?: () => void;
  onCut?: () => void;
  onPaste?: () => void;
  onMoveCells?: (sourceRange: SelectionRange, targetRange: SelectionRange, newActiveCellId?: string) => void;
}

const DEFAULT_COLS = 26; // A to Z
const DEFAULT_ROWS = 60; // 1 to 60
const DEFAULT_COL_WIDTH = 100;
const DEFAULT_ROW_HEIGHT = 24;

export const SpreadsheetGrid: React.FC<SpreadsheetGridProps> = ({
  sheet,
  activeCellId,
  selectionRange,
  onSelectCell,
  onSelectRange,
  onCellChange,
  onAutoFill,
  targetLessonCell,
  showGridlines,
  showHeaders,
  zoomLevel,
  isDarkMode,
  onShortcutUsed,
  onApplyStyle,
  onInsertRow,
  onDeleteRow,
  onAutoSum,
  onOpenChartModal,
  isFormulaMode = false,
  formulaPointedRange = null,
  onPointCell,
  onPointRange,
  onSubmitFormula,
  onCancelFormula,
  onResizeColumn,
  onResizeRow,
  onUndo,
  onRedo,
  onCellTyping,
  isFormatPainterActive = false,
  isFormatPainterLocked = false,
  formatPainterStyle = null,
  onToggleFormatPainter,
  onApplyFormatPainter,
  onCancelFormatPainter,
  onClearSelection,
  onCopy,
  onCut,
  onPaste,
  onMoveCells,
}) => {
  const [isEditing, setIsEditing] = useState(false);
  const [editValue, setEditValue] = useState('');
  const [isSelecting, setIsSelecting] = useState(false);

  // Drag & Move Cells States (Excel Drag and Drop Border)
  const [isDraggingCells, setIsDraggingCells] = useState(false);
  const [dragOrigin, setDragOrigin] = useState<{ colIdx: number; row: number } | null>(null);
  const [dragDelta, setDragDelta] = useState<{ col: number; row: number }>({ col: 0, row: 0 });
  const [moveConfirmation, setMoveConfirmation] = useState<{
    sourceRange: SelectionRange;
    targetRange: SelectionRange;
    newActiveCellId: string;
    overwrittenCount: number;
  } | null>(null);

  // Column and Row Selection States (Whole col/row selection)
  const [isSelectingCols, setIsSelectingCols] = useState(false);
  const [colSelectStart, setColSelectStart] = useState<string | null>(null);

  const [isSelectingRows, setIsSelectingRows] = useState(false);
  const [rowSelectStart, setRowSelectStart] = useState<number | null>(null);

  const mouseDownCellIdRef = useRef<string | null>(null);

  // Column and Row Resize States
  const [resizingCol, setResizingCol] = useState<{
    col: string;
    startX: number;
    startWidth: number;
    currentWidth: number;
  } | null>(null);

  const [resizingRow, setResizingRow] = useState<{
    row: number;
    startY: number;
    startHeight: number;
    currentHeight: number;
  } | null>(null);

  // Formula Pointing Drag State
  const [isPointingFormula, setIsPointingFormula] = useState(false);
  const [formulaPointStart, setFormulaPointStart] = useState<string | null>(null);

  // Autofill Dragging States
  const [isDraggingFillHandle, setIsDraggingFillHandle] = useState(false);
  const [fillTargetCell, setFillTargetCell] = useState<string | null>(null);
  const [isFillModifierActive, setIsFillModifierActive] = useState(false);
  const [fillMousePos, setFillMousePos] = useState<{ x: number; y: number } | null>(null);

  // Right-Click Context Menu State
  const [contextMenu, setContextMenu] = useState<{
    x: number;
    y: number;
    cellId: string;
  } | null>(null);

  // Internal clipboard for copy/cut/paste
  const [clipboardData, setClipboardData] = useState<{
    raw: string;
    style?: CellStyle;
    isCut?: boolean;
    sourceCellId?: string;
  } | null>(null);

  const gridContainerRef = useRef<HTMLDivElement>(null);
  const editInputRef = useRef<HTMLInputElement>(null);
  const isNavigatingViaKeyRef = useRef(false);

  // Column and row sizes (responsive to active dragging)
  const getColWidth = useCallback(
    (col: string) => {
      if (resizingCol && resizingCol.col === col) {
        return resizingCol.currentWidth;
      }
      return sheet.colWidths?.[col] || DEFAULT_COL_WIDTH;
    },
    [sheet.colWidths, resizingCol]
  );

  const getRowHeight = useCallback(
    (row: number) => {
      if (resizingRow && resizingRow.row === row) {
        return resizingRow.currentHeight;
      }
      return sheet.rowHeights?.[row] || DEFAULT_ROW_HEIGHT;
    },
    [sheet.rowHeights, resizingRow]
  );

  // Column and Row Resize Handlers
  const handleColResizeStart = (col: string, e: React.MouseEvent) => {
    e.preventDefault();
    e.stopPropagation();
    const currentWidth = getColWidth(col);
    document.body.style.cursor = 'col-resize';
    document.body.style.userSelect = 'none';
    setResizingCol({
      col,
      startX: e.clientX,
      startWidth: currentWidth,
      currentWidth,
    });
  };

  const handleColAutoFit = (col: string, e: React.MouseEvent) => {
    e.preventDefault();
    e.stopPropagation();
    let maxLen = 0;
    for (let r = 1; r <= DEFAULT_ROWS; r++) {
      const cell = sheet.cells[`${col}${r}`];
      if (cell) {
        const text = cell.computed !== undefined ? String(cell.computed) : cell.raw || '';
        if (text.length > maxLen) maxLen = text.length;
      }
    }
    const autoWidth = Math.max(60, Math.min(450, Math.round(maxLen * 9 + 32)));
    onResizeColumn?.(col, autoWidth);
  };

  const handleRowResizeStart = (row: number, e: React.MouseEvent) => {
    e.preventDefault();
    e.stopPropagation();
    const currentHeight = getRowHeight(row);
    document.body.style.cursor = 'row-resize';
    document.body.style.userSelect = 'none';
    setResizingRow({
      row,
      startY: e.clientY,
      startHeight: currentHeight,
      currentHeight,
    });
  };

  const handleRowResetHeight = (row: number, e: React.MouseEvent) => {
    e.preventDefault();
    e.stopPropagation();
    onResizeRow?.(row, DEFAULT_ROW_HEIGHT);
  };

  // Window listeners for dragging column resize
  useEffect(() => {
    if (!resizingCol) return;

    const handleMouseMove = (e: MouseEvent) => {
      const scale = zoomLevel / 100;
      const deltaX = (e.clientX - resizingCol.startX) / scale;
      const newWidth = Math.max(8, Math.round(resizingCol.startWidth + deltaX));
      setResizingCol((prev) => (prev ? { ...prev, currentWidth: newWidth } : null));
    };

    const handleMouseUp = () => {
      if (resizingCol && onResizeColumn) {
        onResizeColumn(resizingCol.col, resizingCol.currentWidth);
      }
      document.body.style.cursor = '';
      document.body.style.userSelect = '';
      setResizingCol(null);
    };

    window.addEventListener('mousemove', handleMouseMove);
    window.addEventListener('mouseup', handleMouseUp);
    return () => {
      window.removeEventListener('mousemove', handleMouseMove);
      window.removeEventListener('mouseup', handleMouseUp);
    };
  }, [resizingCol, zoomLevel, onResizeColumn]);

  // Window listeners for dragging row resize
  useEffect(() => {
    if (!resizingRow) return;

    const handleMouseMove = (e: MouseEvent) => {
      const scale = zoomLevel / 100;
      const deltaY = (e.clientY - resizingRow.startY) / scale;
      const newHeight = Math.max(18, Math.round(resizingRow.startHeight + deltaY));
      setResizingRow((prev) => (prev ? { ...prev, currentHeight: newHeight } : null));
    };

    const handleMouseUp = () => {
      if (resizingRow && onResizeRow) {
        onResizeRow(resizingRow.row, resizingRow.currentHeight);
      }
      document.body.style.cursor = '';
      document.body.style.userSelect = '';
      setResizingRow(null);
    };

    window.addEventListener('mousemove', handleMouseMove);
    window.addEventListener('mouseup', handleMouseUp);
    return () => {
      window.removeEventListener('mousemove', handleMouseMove);
      window.removeEventListener('mouseup', handleMouseUp);
    };
  }, [resizingRow, zoomLevel, onResizeRow]);

  // Active cell coordinate
  const activeCoord = parseCellAddress(activeCellId) || { col: 'A', row: 1 };
  const activeColIdx = letterToColIndex(activeCoord.col);

  // Calculate regular selection bounds
  const getSelectionBounds = useCallback(() => {
    const range = selectionRange || {
      startCol: activeCoord.col,
      startRow: activeCoord.row,
      endCol: activeCoord.col,
      endRow: activeCoord.row,
    };

    const c1 = letterToColIndex(range.startCol);
    const c2 = letterToColIndex(range.endCol);
    const minColIdx = Math.min(c1, c2);
    const maxColIdx = Math.max(c1, c2);

    const minRow = Math.min(range.startRow, range.endRow);
    const maxRow = Math.max(range.startRow, range.endRow);

    return { minColIdx, maxColIdx, minRow, maxRow };
  }, [selectionRange, activeCoord]);

  const { minColIdx, maxColIdx, minRow, maxRow } = getSelectionBounds();

  const isEntireSheetSelected =
    minColIdx === 0 &&
    maxColIdx === DEFAULT_COLS - 1 &&
    minRow === 1 &&
    maxRow === DEFAULT_ROWS;

  const isEntireRowSelected = minColIdx === 0 && maxColIdx === DEFAULT_COLS - 1;
  const isEntireColSelected = minRow === 1 && maxRow === DEFAULT_ROWS;

  // Drag & Move border handler
  const handleDragBorderStart = (e: React.MouseEvent, colIdx: number, row: number) => {
    if (e.button !== 0) return; // Only left click
    if (isEditing) commitEdit();
    e.preventDefault();
    e.stopPropagation();

    setIsDraggingCells(true);
    setDragOrigin({ colIdx, row });
    setDragDelta({ col: 0, row: 0 });
    sounds.playCellClick();
  };

  const handleGlobalMouseMove = useCallback(
    (e: MouseEvent) => {
      if (isDraggingCells && dragOrigin) {
        const elem = document.elementFromPoint(e.clientX, e.clientY);
        const td = elem?.closest('td[data-cell-id]');
        if (td) {
          const cellId = td.getAttribute('data-cell-id');
          if (cellId) {
            const target = parseCellAddress(cellId);
            if (target) {
              const targetColIdx = letterToColIndex(target.col);
              const dCol = targetColIdx - dragOrigin.colIdx;
              const dRow = target.row - dragOrigin.row;

              const clampedDCol = Math.max(-minColIdx, Math.min(DEFAULT_COLS - 1 - maxColIdx, dCol));
              const clampedDRow = Math.max(1 - minRow, Math.min(DEFAULT_ROWS - maxRow, dRow));

              setDragDelta({ col: clampedDCol, row: clampedDRow });
            }
          }
        }
      }
    },
    [isDraggingCells, dragOrigin, minColIdx, maxColIdx, minRow, maxRow]
  );

  useEffect(() => {
    if (isDraggingCells) {
      window.addEventListener('mousemove', handleGlobalMouseMove);
      return () => window.removeEventListener('mousemove', handleGlobalMouseMove);
    }
  }, [isDraggingCells, handleGlobalMouseMove]);

  // Focus editing input when editing starts
  useEffect(() => {
    if (isEditing && editInputRef.current) {
      editInputRef.current.focus();
      editInputRef.current.select();
    }
  }, [isEditing, activeCellId]);

  // Close context menu on outside click
  useEffect(() => {
    const handleGlobalClick = () => {
      setContextMenu(null);
    };
    window.addEventListener('click', handleGlobalClick);
    return () => window.removeEventListener('click', handleGlobalClick);
  }, []);

  // Listen for Ctrl/Meta/Alt modifier key during autofill drag
  useEffect(() => {
    if (!isDraggingFillHandle) {
      setIsFillModifierActive(false);
      setFillMousePos(null);
      return;
    }

    const handleKeyMod = (e: KeyboardEvent) => {
      setIsFillModifierActive(e.ctrlKey || e.metaKey || e.altKey);
    };

    const handleMouseMove = (e: MouseEvent) => {
      setIsFillModifierActive(e.ctrlKey || e.metaKey || e.altKey);
      setFillMousePos({ x: e.clientX, y: e.clientY });
    };

    window.addEventListener('keydown', handleKeyMod);
    window.addEventListener('keyup', handleKeyMod);
    window.addEventListener('mousemove', handleMouseMove);
    return () => {
      window.removeEventListener('keydown', handleKeyMod);
      window.removeEventListener('keyup', handleKeyMod);
      window.removeEventListener('mousemove', handleMouseMove);
    };
  }, [isDraggingFillHandle]);

  // Enter edit mode
  const startEditing = useCallback((initialVal?: string, cellIdToEdit?: string) => {
    const targetCellId = cellIdToEdit || activeCellId;
    const currentRaw = sheet.cells[targetCellId]?.raw || '';
    const newVal = initialVal !== undefined ? initialVal : currentRaw;
    setEditValue(newVal);
    setIsEditing(true);
    setContextMenu(null);
    if (initialVal !== undefined) {
      onCellTyping?.(newVal);
    }
  }, [sheet.cells, activeCellId, onCellTyping]);

  const commitEdit = useCallback(() => {
    if (isEditing) {
      onCellChange(activeCellId, editValue);
      setIsEditing(false);
    }
  }, [isEditing, activeCellId, editValue, onCellChange]);

  const cancelEdit = useCallback(() => {
    setIsEditing(false);
  }, []);

  // Keyboard navigation & shortcuts
  const handleKeyDown = useCallback((e: React.KeyboardEvent) => {
    if (isEditing) {
      if (e.key === 'Enter') {
        e.preventDefault();
        const valToSave = editValue;
        const targetRow = e.shiftKey ? Math.max(1, activeCoord.row - 1) : activeCoord.row + 1;
        const nextCell = `${activeCoord.col}${targetRow}`;
        isNavigatingViaKeyRef.current = true;
        if (valToSave.startsWith('=')) {
          setIsEditing(false);
          onSubmitFormula?.(false);
        } else {
          onCellChange(activeCellId, valToSave);
        }
        onSelectCell(nextCell);
        const nextRaw = sheet.cells[nextCell]?.raw || '';
        setEditValue(nextRaw);
        setIsEditing(true);
        setTimeout(() => {
          isNavigatingViaKeyRef.current = false;
          if (editInputRef.current) {
            editInputRef.current.focus();
            editInputRef.current.select();
          }
        }, 40);
        return;
      } else if (e.key === 'Tab') {
        e.preventDefault();
        const valToSave = editValue;
        const targetColIdx = e.shiftKey ? Math.max(0, activeColIdx - 1) : activeColIdx + 1;
        const nextCol = colIndexToLetter(targetColIdx);
        const nextCell = `${nextCol}${activeCoord.row}`;
        isNavigatingViaKeyRef.current = true;
        if (valToSave.startsWith('=')) {
          setIsEditing(false);
          onSubmitFormula?.(true);
        } else {
          onCellChange(activeCellId, valToSave);
        }
        onSelectCell(nextCell);
        const nextRaw = sheet.cells[nextCell]?.raw || '';
        setEditValue(nextRaw);
        setIsEditing(true);
        setTimeout(() => {
          isNavigatingViaKeyRef.current = false;
          if (editInputRef.current) {
            editInputRef.current.focus();
            editInputRef.current.select();
          }
        }, 40);
        return;
      } else if (e.key === 'Escape') {
        e.preventDefault();
        if (editValue.startsWith('=')) {
          setIsEditing(false);
          onCancelFormula?.();
        } else {
          cancelEdit();
        }
        return;
      }
      return;
    }

    // Tryb "Wskaż" - zatwierdzanie Enterem i anulowanie Escapem w arkuszu
    if (isFormulaMode) {
      if (e.key === 'Enter') {
        e.preventDefault();
        onSubmitFormula?.(false);
        return;
      }
      if (e.key === 'Tab') {
        e.preventDefault();
        onSubmitFormula?.(true);
        return;
      }
      if (e.key === 'Escape') {
        e.preventDefault();
        if (isFormatPainterActive) {
          onCancelFormatPainter?.();
          return;
        }
        onCancelFormula?.();
        return;
      }
    }

    // Handle Move Confirmation Dialog Keyboard Shortcuts
    if (moveConfirmation) {
      if (e.key === 'Escape') {
        e.preventDefault();
        setMoveConfirmation(null);
        sounds.playCellClick();
        return;
      }
      if (e.key === 'Enter') {
        e.preventDefault();
        onMoveCells?.(
          moveConfirmation.sourceRange,
          moveConfirmation.targetRange,
          moveConfirmation.newActiveCellId
        );
        setMoveConfirmation(null);
        return;
      }
    }

    if (e.key === 'Escape') {
      e.preventDefault();
      if (isDraggingCells) {
        setIsDraggingCells(false);
        setDragDelta({ col: 0, row: 0 });
        setDragOrigin(null);
        return;
      }
      if (isFormatPainterActive) {
        onCancelFormatPainter?.();
        return;
      }
      if (contextMenu) {
        setContextMenu(null);
        return;
      }
      if (selectionRange) {
        onSelectRange(null);
        return;
      }
      return;
    }

    // Select All Shortcut (Ctrl+A)
    if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'a') {
      e.preventDefault();
      const maxColLetter = colIndexToLetter(DEFAULT_COLS - 1);
      const allRange = {
        startCol: 'A',
        startRow: 1,
        endCol: maxColLetter,
        endRow: DEFAULT_ROWS,
      };
      onSelectCell('A1');
      onSelectRange(allRange);
      if (isFormatPainterActive) {
        onApplyFormatPainter?.(allRange);
      }
      onShortcutUsed?.('Ctrl+A');
      return;
    }

    // Global Shortcuts
    if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'z') {
      e.preventDefault();
      if (e.shiftKey) {
        onRedo?.();
        onShortcutUsed?.('Ctrl+Y');
      } else {
        onUndo?.();
        onShortcutUsed?.('Ctrl+Z');
      }
      return;
    }
    if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'y') {
      e.preventDefault();
      onRedo?.();
      onShortcutUsed?.('Ctrl+Y');
      return;
    }

    // Copy / Paste Format (Ctrl+Shift+C / Ctrl+Shift+V)
    if ((e.ctrlKey || e.metaKey) && e.shiftKey && e.key.toLowerCase() === 'c') {
      e.preventDefault();
      onToggleFormatPainter?.(false);
      onShortcutUsed?.('Ctrl+Shift+C');
      return;
    }
    if ((e.ctrlKey || e.metaKey) && e.shiftKey && e.key.toLowerCase() === 'v') {
      e.preventDefault();
      onApplyFormatPainter?.(selectionRange || activeCellId);
      onShortcutUsed?.('Ctrl+Shift+V');
      return;
    }

    if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'b') {
      e.preventDefault();
      onShortcutUsed?.('Ctrl+B');
      return;
    }
    if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'i') {
      e.preventDefault();
      onShortcutUsed?.('Ctrl+I');
      return;
    }
    if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'u') {
      e.preventDefault();
      onShortcutUsed?.('Ctrl+U');
      return;
    }

    // Copy / Cut / Paste shortcuts
    if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'c') {
      e.preventDefault();
      if (onCopy) {
        onCopy();
      } else {
        const cell = sheet.cells[activeCellId];
        if (cell) {
          setClipboardData({ raw: cell.raw, style: cell.style });
        }
      }
      onShortcutUsed?.('Ctrl+C');
      return;
    }
    if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'x') {
      e.preventDefault();
      if (onCut) {
        onCut();
      } else {
        const cell = sheet.cells[activeCellId];
        if (cell) {
          setClipboardData({ raw: cell.raw, style: cell.style, isCut: true, sourceCellId: activeCellId });
        }
      }
      onShortcutUsed?.('Ctrl+X');
      return;
    }
    if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'v') {
      e.preventDefault();
      if (onPaste) {
        onPaste();
      } else if (clipboardData) {
        onCellChange(activeCellId, clipboardData.raw);
        if (clipboardData.isCut && clipboardData.sourceCellId) {
          onCellChange(clipboardData.sourceCellId, '');
          setClipboardData(null);
        }
      }
      onShortcutUsed?.('Ctrl+V');
      return;
    }

    if (e.key === 'F2') {
      e.preventDefault();
      startEditing();
      return;
    }

    if (e.key === 'Delete' || e.key === 'Backspace') {
      e.preventDefault();
      if (onClearSelection) {
        onClearSelection();
      } else {
        onCellChange(activeCellId, '');
      }
      return;
    }

    if (e.key === 'Enter') {
      e.preventDefault();
      if (e.shiftKey) {
        if (activeCoord.row > 1) {
          onSelectCell(`${activeCoord.col}${activeCoord.row - 1}`);
        }
      } else {
        onSelectCell(`${activeCoord.col}${activeCoord.row + 1}`);
      }
      return;
    }

    if (e.key === 'Tab') {
      e.preventDefault();
      if (e.shiftKey) {
        if (activeColIdx > 0) {
          onSelectCell(`${colIndexToLetter(activeColIdx - 1)}${activeCoord.row}`);
        }
      } else {
        onSelectCell(`${colIndexToLetter(activeColIdx + 1)}${activeCoord.row}`);
      }
      return;
    }

    if (e.key === 'ArrowUp') {
      e.preventDefault();
      if (activeCoord.row > 1) {
        onSelectCell(`${activeCoord.col}${activeCoord.row - 1}`);
      }
      return;
    }
    if (e.key === 'ArrowDown') {
      e.preventDefault();
      onSelectCell(`${activeCoord.col}${activeCoord.row + 1}`);
      return;
    }
    if (e.key === 'ArrowLeft') {
      e.preventDefault();
      if (activeColIdx > 0) {
        onSelectCell(`${colIndexToLetter(activeColIdx - 1)}${activeCoord.row}`);
      }
      return;
    }
    if (e.key === 'ArrowRight') {
      e.preventDefault();
      onSelectCell(`${colIndexToLetter(activeColIdx + 1)}${activeCoord.row}`);
      return;
    }

    // Direct character typing enters edit mode
    if (e.key.length === 1 && !e.ctrlKey && !e.metaKey && !e.altKey) {
      startEditing(e.key);
    }
  }, [
    isEditing,
    editValue,
    isFormulaMode,
    onSubmitFormula,
    onCancelFormula,
    commitEdit,
    cancelEdit,
    activeCoord,
    activeColIdx,
    onSelectCell,
    onShortcutUsed,
    startEditing,
    onCellChange,
    activeCellId,
    sheet.cells,
    clipboardData,
  ]);

  // Handle Mouse Selection
  const handleCellMouseDown = (cellId: string, e: React.MouseEvent) => {
    if (e.button !== 0) return; // Only left click

    // If formula mode is active:
    if (isFormulaMode) {
      if (cellId !== activeCellId) {
        e.preventDefault();
        sounds.playCellClick();
        setIsEditing(false);
        setIsPointingFormula(true);
        setFormulaPointStart(cellId);
        onPointCell?.(cellId);
        gridContainerRef.current?.focus();
        return;
      }
    }

    if (isEditing) commitEdit();
    setContextMenu(null);

    sounds.playCellClick();
    mouseDownCellIdRef.current = cellId;

    if (e.shiftKey) {
      // Expand selection range from active to clicked
      const clicked = parseCellAddress(cellId);
      if (clicked) {
        const newRange = {
          startCol: activeCoord.col,
          startRow: activeCoord.row,
          endCol: clicked.col,
          endRow: clicked.row,
        };
        onSelectRange(newRange);
        if (isFormatPainterActive) {
          onApplyFormatPainter?.(newRange);
        }
      }
    } else {
      onSelectCell(cellId);
      onSelectRange(null);
      setIsSelecting(true);
    }
  };

  const handleCellMouseEnter = (cellId: string) => {
    // If dragging while in formula pointing mode, expand range
    if (isPointingFormula && formulaPointStart) {
      const startCoord = parseCellAddress(formulaPointStart);
      const currCoord = parseCellAddress(cellId);
      if (startCoord && currCoord) {
        onPointRange?.({
          startCol: startCoord.col,
          startRow: startCoord.row,
          endCol: currCoord.col,
          endRow: currCoord.row,
        });
      }
      return;
    }

    if (isDraggingCells && dragOrigin) {
      const target = parseCellAddress(cellId);
      if (target) {
        const targetColIdx = letterToColIndex(target.col);
        const dCol = targetColIdx - dragOrigin.colIdx;
        const dRow = target.row - dragOrigin.row;

        const clampedDCol = Math.max(-minColIdx, Math.min(DEFAULT_COLS - 1 - maxColIdx, dCol));
        const clampedDRow = Math.max(1 - minRow, Math.min(DEFAULT_ROWS - maxRow, dRow));

        setDragDelta({ col: clampedDCol, row: clampedDRow });
      }
      return;
    }

    if (isSelecting) {
      const target = parseCellAddress(cellId);
      if (target) {
        onSelectRange({
          startCol: activeCoord.col,
          startRow: activeCoord.row,
          endCol: target.col,
          endRow: target.row,
        });
      }
    } else if (isDraggingFillHandle) {
      setFillTargetCell(cellId);
    }
  };

  // Column Header Selection Handlers (Excel Style)
  const handleColHeaderMouseDown = (colLetter: string, cIdx: number, e: React.MouseEvent) => {
    if (e.button !== 0) return;
    if (isEditing) commitEdit();
    setContextMenu(null);
    sounds.playCellClick();

    if (isFormulaMode) {
      onPointRange?.({
        startCol: colLetter,
        startRow: 1,
        endCol: colLetter,
        endRow: DEFAULT_ROWS,
      });
      return;
    }

    if (e.shiftKey && selectionRange) {
      const newRange: SelectionRange = {
        startCol: selectionRange.startCol,
        startRow: 1,
        endCol: colLetter,
        endRow: DEFAULT_ROWS,
      };
      onSelectRange(newRange);
      if (isFormatPainterActive) {
        onApplyFormatPainter?.(newRange);
      }
    } else {
      onSelectCell(`${colLetter}1`);
      const newRange: SelectionRange = {
        startCol: colLetter,
        startRow: 1,
        endCol: colLetter,
        endRow: DEFAULT_ROWS,
      };
      onSelectRange(newRange);
      setIsSelectingCols(true);
      setColSelectStart(colLetter);
    }
  };

  const handleColHeaderMouseEnter = (colLetter: string) => {
    if (isSelectingCols && colSelectStart) {
      onSelectRange({
        startCol: colSelectStart,
        startRow: 1,
        endCol: colLetter,
        endRow: DEFAULT_ROWS,
      });
    }
  };

  const handleColContextMenu = (colLetter: string, e: React.MouseEvent) => {
    e.preventDefault();
    e.stopPropagation();
    sounds.playCellClick();
    onSelectCell(`${colLetter}1`);
    onSelectRange({
      startCol: colLetter,
      startRow: 1,
      endCol: colLetter,
      endRow: DEFAULT_ROWS,
    });
    setContextMenu({
      x: Math.min(e.clientX, window.innerWidth - 280),
      y: Math.min(e.clientY, window.innerHeight - 380),
      cellId: `${colLetter}1`,
    });
  };

  // Row Header Selection Handlers (Excel Style)
  const handleRowHeaderMouseDown = (rowNumber: number, e: React.MouseEvent) => {
    if (e.button !== 0) return;
    if (isEditing) commitEdit();
    setContextMenu(null);
    sounds.playCellClick();

    const maxColLetter = colIndexToLetter(DEFAULT_COLS - 1);

    if (isFormulaMode) {
      onPointRange?.({
        startCol: 'A',
        startRow: rowNumber,
        endCol: maxColLetter,
        endRow: rowNumber,
      });
      return;
    }

    if (e.shiftKey && selectionRange) {
      const newRange: SelectionRange = {
        startCol: 'A',
        startRow: selectionRange.startRow,
        endCol: maxColLetter,
        endRow: rowNumber,
      };
      onSelectRange(newRange);
      if (isFormatPainterActive) {
        onApplyFormatPainter?.(newRange);
      }
    } else {
      onSelectCell(`A${rowNumber}`);
      const newRange: SelectionRange = {
        startCol: 'A',
        startRow: rowNumber,
        endCol: maxColLetter,
        endRow: rowNumber,
      };
      onSelectRange(newRange);
      setIsSelectingRows(true);
      setRowSelectStart(rowNumber);
    }
  };

  const handleRowHeaderMouseEnter = (rowNumber: number) => {
    if (isSelectingRows && rowSelectStart !== null) {
      const maxColLetter = colIndexToLetter(DEFAULT_COLS - 1);
      onSelectRange({
        startCol: 'A',
        startRow: rowSelectStart,
        endCol: maxColLetter,
        endRow: rowNumber,
      });
    }
  };

  const handleRowContextMenu = (rowNumber: number, e: React.MouseEvent) => {
    e.preventDefault();
    e.stopPropagation();
    sounds.playCellClick();
    const maxColLetter = colIndexToLetter(DEFAULT_COLS - 1);
    onSelectCell(`A${rowNumber}`);
    onSelectRange({
      startCol: 'A',
      startRow: rowNumber,
      endCol: maxColLetter,
      endRow: rowNumber,
    });
    setContextMenu({
      x: Math.min(e.clientX, window.innerWidth - 280),
      y: Math.min(e.clientY, window.innerHeight - 380),
      cellId: `A${rowNumber}`,
    });
  };

  // Select Entire Sheet (Corner Cell ◢)
  const handleSelectAll = (e: React.MouseEvent) => {
    e.preventDefault();
    if (isEditing) commitEdit();
    setContextMenu(null);
    sounds.playCellClick();
    const maxColLetter = colIndexToLetter(DEFAULT_COLS - 1);
    const allRange: SelectionRange = {
      startCol: 'A',
      startRow: 1,
      endCol: maxColLetter,
      endRow: DEFAULT_ROWS,
    };
    onSelectCell('A1');
    onSelectRange(allRange);
    if (isFormatPainterActive) {
      onApplyFormatPainter?.(allRange);
    }
  };

  // Right Click Context Menu Handler
  const handleCellContextMenu = (cellId: string, e: React.MouseEvent) => {
    e.preventDefault();
    e.stopPropagation();

    sounds.playCellClick();

    const clicked = parseCellAddress(cellId);
    if (!clicked) return;

    // If clicked cell is not within current selection, select it
    const cIdx = letterToColIndex(clicked.col);
    if (!isCellSelected(cIdx, clicked.row)) {
      onSelectCell(cellId);
      onSelectRange(null);
    }

    // Calculate smart menu coordinates so it doesn't clip offscreen
    const menuWidth = 270;
    const menuHeight = 360;
    const x = Math.min(e.clientX, window.innerWidth - menuWidth - 10);
    const y = Math.min(e.clientY, window.innerHeight - menuHeight - 10);

    setContextMenu({ x, y, cellId });
  };

  // Global Mouse Up for drag selection & autofill release
  const handleGlobalMouseUp = useCallback(() => {
    if (isPointingFormula) {
      setIsPointingFormula(false);
      setFormulaPointStart(null);
      gridContainerRef.current?.focus();
    }
    if (isSelectingCols) {
      setIsSelectingCols(false);
      setColSelectStart(null);
      if (isFormatPainterActive && selectionRange) {
        onApplyFormatPainter?.(selectionRange);
      }
    }
    if (isSelectingRows) {
      setIsSelectingRows(false);
      setRowSelectStart(null);
      if (isFormatPainterActive && selectionRange) {
        onApplyFormatPainter?.(selectionRange);
      }
    }
    if (isSelecting) {
      setIsSelecting(false);
      if (isFormatPainterActive) {
        const clickedId = mouseDownCellIdRef.current;
        const clickedCoord = clickedId ? parseCellAddress(clickedId) : null;
        const targetRange: SelectionRange = selectionRange || (clickedCoord ? {
          startCol: clickedCoord.col,
          startRow: clickedCoord.row,
          endCol: clickedCoord.col,
          endRow: clickedCoord.row,
        } : {
          startCol: activeCoord.col,
          startRow: activeCoord.row,
          endCol: activeCoord.col,
          endRow: activeCoord.row,
        });
        onApplyFormatPainter?.(targetRange);
      }
    }
    if (isDraggingCells) {
      setIsDraggingCells(false);
      const dCol = dragDelta.col;
      const dRow = dragDelta.row;
      setDragDelta({ col: 0, row: 0 });
      setDragOrigin(null);

      if (dCol !== 0 || dRow !== 0) {
        const targetMinCol = minColIdx + dCol;
        const targetMaxCol = maxColIdx + dCol;
        const targetMinRow = minRow + dRow;
        const targetMaxRow = maxRow + dRow;

        const targetRange: SelectionRange = {
          startCol: colIndexToLetter(targetMinCol),
          startRow: targetMinRow,
          endCol: colIndexToLetter(targetMaxCol),
          endRow: targetMaxRow,
        };

        const sourceRange: SelectionRange = {
          startCol: colIndexToLetter(minColIdx),
          startRow: minRow,
          endCol: colIndexToLetter(maxColIdx),
          endRow: maxRow,
        };

        const newActiveCellId = `${colIndexToLetter(activeColIdx + dCol)}${activeCoord.row + dRow}`;

        // Check if destination cells contain data
        let hasExistingData = false;
        let overwrittenCount = 0;

        for (let r = targetMinRow; r <= targetMaxRow; r++) {
          for (let c = targetMinCol; c <= targetMaxCol; c++) {
            const destId = `${colIndexToLetter(c)}${r}`;
            const cell = sheet.cells[destId];
            if (cell && cell.raw !== undefined && cell.raw !== null && cell.raw.trim() !== '') {
              hasExistingData = true;
              overwrittenCount++;
            }
          }
        }

        if (hasExistingData) {
          setMoveConfirmation({
            sourceRange,
            targetRange,
            newActiveCellId,
            overwrittenCount,
          });
        } else {
          onMoveCells?.(sourceRange, targetRange, newActiveCellId);
        }
      }
    }
    if (isDraggingFillHandle && fillTargetCell) {
      const copyOnly = isFillModifierActive;
      setIsDraggingFillHandle(false);
      setIsFillModifierActive(false);
      setFillMousePos(null);

      const normalizedSourceRange: SelectionRange = {
        startCol: colIndexToLetter(minColIdx),
        startRow: minRow,
        endCol: colIndexToLetter(maxColIdx),
        endRow: maxRow,
      };

      const targetCoord = parseCellAddress(fillTargetCell);
      if (targetCoord) {
        const targetColIdx = letterToColIndex(targetCoord.col);
        const targetRow = targetCoord.row;

        const dRow = targetRow > maxRow ? targetRow - maxRow : targetRow < minRow ? minRow - targetRow : 0;
        const dCol = targetColIdx > maxColIdx ? targetColIdx - maxColIdx : targetColIdx < minColIdx ? minColIdx - targetColIdx : 0;

        let fullMinCol = minColIdx;
        let fullMaxCol = maxColIdx;
        let fullMinRow = minRow;
        let fullMaxRow = maxRow;

        if (dRow >= dCol && dRow > 0) {
          fullMinRow = Math.min(minRow, targetRow);
          fullMaxRow = Math.max(maxRow, targetRow);
        } else if (dCol > 0) {
          fullMinCol = Math.min(minColIdx, targetColIdx);
          fullMaxCol = Math.max(maxColIdx, targetColIdx);
        }

        const fullRange: SelectionRange = {
          startCol: colIndexToLetter(fullMinCol),
          startRow: fullMinRow,
          endCol: colIndexToLetter(fullMaxCol),
          endRow: fullMaxRow,
        };
        onAutoFill(normalizedSourceRange, fullRange, copyOnly);
      }
      setFillTargetCell(null);
    } else if (isDraggingFillHandle) {
      setIsDraggingFillHandle(false);
      setIsFillModifierActive(false);
      setFillMousePos(null);
      setFillTargetCell(null);
    }
  }, [
    isPointingFormula,
    isSelectingCols,
    isSelectingRows,
    isSelecting,
    isDraggingCells,
    dragDelta,
    minColIdx,
    maxColIdx,
    minRow,
    maxRow,
    activeColIdx,
    activeCoord,
    sheet.cells,
    onMoveCells,
    isDraggingFillHandle,
    fillTargetCell,
    selectionRange,
    onAutoFill,
    isFormatPainterActive,
    onApplyFormatPainter,
  ]);

  useEffect(() => {
    window.addEventListener('mouseup', handleGlobalMouseUp);
    return () => window.removeEventListener('mouseup', handleGlobalMouseUp);
  }, [handleGlobalMouseUp]);

  // Calculate formula pointed bounds
  const getFormulaPointedBounds = useCallback(() => {
    if (!formulaPointedRange) return null;
    const c1 = letterToColIndex(formulaPointedRange.startCol);
    const c2 = letterToColIndex(formulaPointedRange.endCol);
    const minCol = Math.min(c1, c2);
    const maxCol = Math.max(c1, c2);
    const minRow = Math.min(formulaPointedRange.startRow, formulaPointedRange.endRow);
    const maxRow = Math.max(formulaPointedRange.startRow, formulaPointedRange.endRow);
    return { minCol, maxCol, minRow, maxRow };
  }, [formulaPointedRange]);

  const formulaBounds = getFormulaPointedBounds();

  const isCellInFormulaPoint = (colIdx: number, row: number) => {
    if (!formulaBounds) return false;
    return (
      colIdx >= formulaBounds.minCol &&
      colIdx <= formulaBounds.maxCol &&
      row >= formulaBounds.minRow &&
      row <= formulaBounds.maxRow
    );
  };

  // Calculate Autofill dragging bounds (for real Excel dashed marquee preview)
  const getFillBounds = useCallback(() => {
    if (!isDraggingFillHandle || !fillTargetCell) return null;

    const targetCoord = parseCellAddress(fillTargetCell);
    if (!targetCoord) return null;

    const targetColIdx = letterToColIndex(targetCoord.col);
    const targetRow = targetCoord.row;

    const dRow = targetRow > maxRow ? targetRow - maxRow : targetRow < minRow ? minRow - targetRow : 0;
    const dCol = targetColIdx > maxColIdx ? targetColIdx - maxColIdx : targetColIdx < minColIdx ? minColIdx - targetColIdx : 0;

    let fullMinCol = minColIdx;
    let fullMaxCol = maxColIdx;
    let fullMinRow = minRow;
    let fullMaxRow = maxRow;

    if (dRow >= dCol && dRow > 0) {
      fullMinRow = Math.min(minRow, targetRow);
      fullMaxRow = Math.max(maxRow, targetRow);
    } else if (dCol > 0) {
      fullMinCol = Math.min(minColIdx, targetColIdx);
      fullMaxCol = Math.max(maxColIdx, targetColIdx);
    }

    return { fullMinCol, fullMaxCol, fullMinRow, fullMaxRow };
  }, [isDraggingFillHandle, fillTargetCell, minColIdx, maxColIdx, minRow, maxRow]);

  const fillBounds = getFillBounds();

  // Helper to determine if a cell is inside the regular selection
  const isCellSelected = (colIdx: number, row: number) => {
    return colIdx >= minColIdx && colIdx <= maxColIdx && row >= minRow && row <= maxRow;
  };

  const isCellActive = (colIdx: number, row: number) => {
    return colIdx === activeColIdx && row === activeCoord.row;
  };

  // Helper to determine if a cell is inside the autofill drag projection
  const isCellInFillArea = (colIdx: number, row: number) => {
    if (!fillBounds) return false;
    return (
      colIdx >= fillBounds.fullMinCol &&
      colIdx <= fillBounds.fullMaxCol &&
      row >= fillBounds.fullMinRow &&
      row <= fillBounds.fullMaxRow
    );
  };

  // Helper to determine if a cell is in the extension area of autofill (excluding original source selection)
  const isCellInFillExtension = (colIdx: number, row: number) => {
    if (!fillBounds) return false;
    const inArea =
      colIdx >= fillBounds.fullMinCol &&
      colIdx <= fillBounds.fullMaxCol &&
      row >= fillBounds.fullMinRow &&
      row <= fillBounds.fullMaxRow;
    return inArea && !isCellSelected(colIdx, row);
  };

  // Border styling helper
  const getBorderStyle = (style?: CellStyle) => {
    if (!style?.border) return '';
    switch (style.border) {
      case 'all':
        return 'border border-neutral-400 dark:border-neutral-500';
      case 'outside':
        return 'border-2 border-neutral-800 dark:border-neutral-200';
      case 'thick-bottom':
        return 'border-b-2 border-b-neutral-900 dark:border-b-white';
      default:
        return '';
    }
  };

  // Total table width calculated strictly from all columns + header corner
  const totalTableWidth = useMemo(() => {
    let total = showHeaders ? 40 : 0;
    for (let c = 0; c < DEFAULT_COLS; c++) {
      total += getColWidth(colIndexToLetter(c));
    }
    return total;
  }, [showHeaders, getColWidth]);

  return (
    <div
      ref={gridContainerRef}
      tabIndex={0}
      onKeyDown={handleKeyDown}
      className={`relative flex-1 overflow-auto select-none outline-none excel-grid transition-colors ${
        isDarkMode ? 'bg-[#181818] text-neutral-200' : 'bg-white text-neutral-900'
      } ${isFormatPainterActive ? 'cursor-copy' : ''} ${isDraggingCells ? 'cursor-move' : ''}`}
      style={{
        transform: `scale(${zoomLevel / 100})`,
        transformOrigin: 'top left',
      }}
    >
      {/* Drag & Move Cells Preview Banner */}
      {isDraggingCells && (dragDelta.col !== 0 || dragDelta.row !== 0) && (
        <div className="sticky top-2 left-1/2 -translate-x-1/2 z-40 w-fit mx-auto my-1 px-3.5 py-1.5 bg-[#107c41] text-white font-semibold text-xs rounded-full shadow-2xl border border-emerald-300/40 flex items-center gap-2 select-none pointer-events-none animate-in fade-in duration-100">
          <Move size={14} className="text-white animate-pulse" />
          <span>
            Przenieś do: <strong>{colIndexToLetter(minColIdx + dragDelta.col)}{minRow + dragDelta.row}</strong>
            {minColIdx !== maxColIdx || minRow !== maxRow
              ? `:${colIndexToLetter(maxColIdx + dragDelta.col)}${maxRow + dragDelta.row}`
              : ''}
          </span>
          <span className="text-[10px] text-emerald-100 font-normal">
            (Upuść, aby przenieść • Esc: anuluj)
          </span>
        </div>
      )}

      {/* Format Painter Active Status Banner */}
      {isFormatPainterActive && (
        <div className="sticky top-2 left-1/2 -translate-x-1/2 z-40 w-fit mx-auto my-1 px-3 py-1 bg-amber-500 text-neutral-900 font-semibold text-xs rounded-full shadow-xl border border-amber-300 flex items-center gap-2 select-none pointer-events-none">
          <Paintbrush size={14} className="text-neutral-900 animate-pulse" />
          <span>Malarz formatów: Kliknij lub zaznacz komórkę, kolumnę lub wiersz, aby wkleić styl</span>
          {isFormatPainterLocked && (
            <span className="bg-neutral-900 text-amber-300 px-1.5 py-0.5 rounded text-[10px] uppercase font-bold">
              Tryb ciągły (PIN)
            </span>
          )}
          <span className="text-[10px] text-neutral-800 font-normal">(Esc: anuluj)</span>
        </div>
      )}

      <table
        className="border-collapse select-none"
        style={{
          tableLayout: 'fixed',
          width: totalTableWidth,
          minWidth: totalTableWidth,
        }}
      >
        {/* W3C Colgroup: Strictly enforces column widths on the entire table */}
        <colgroup>
          {showHeaders && <col style={{ width: 40 }} />}
          {Array.from({ length: DEFAULT_COLS }).map((_, cIdx) => {
            const colLetter = colIndexToLetter(cIdx);
            const w = getColWidth(colLetter);
            return <col key={colLetter} style={{ width: w }} />;
          })}
        </colgroup>

        {/* Table Header: Column letters */}
        {showHeaders && (
          <thead>
            <tr className="sticky top-0 z-20 shadow-xs">
              {/* Corner Cell (Select All) */}
              <th
                onClick={handleSelectAll}
                className={`w-10 h-6 border-b border-r text-center font-normal text-[11px] cursor-pointer transition-colors ${
                  isEntireSheetSelected
                    ? 'bg-[#107c41] text-white font-bold'
                    : isDarkMode
                    ? 'bg-[#2b2b2b] border-[#404040] text-neutral-400 hover:bg-[#383838]'
                    : 'bg-[#f3f4f6] border-[#d1d5db] text-neutral-500 hover:bg-[#e2e8f0]'
                }`}
                title="Zaznacz wszystkie komórki w arkuszu (Ctrl+A)"
              >
                ◢
              </th>
              {Array.from({ length: DEFAULT_COLS }).map((_, cIdx) => {
                const colLetter = colIndexToLetter(cIdx);
                const isColActive =
                  cIdx >= minColIdx &&
                  cIdx <= maxColIdx &&
                  (!isEntireRowSelected || isEntireSheetSelected);
                const width = getColWidth(colLetter);

                return (
                  <th
                    key={colLetter}
                    onMouseDown={(e) => handleColHeaderMouseDown(colLetter, cIdx, e)}
                    onMouseEnter={() => handleColHeaderMouseEnter(colLetter)}
                    onContextMenu={(e) => handleColContextMenu(colLetter, e)}
                    style={{ width, minWidth: width, maxWidth: width, overflow: 'hidden', boxSizing: 'border-box' }}
                    className={`relative h-6 border-b border-r text-center font-medium text-[11px] px-1 select-none font-mono cursor-pointer transition-colors ${
                      isColActive
                        ? 'bg-[#107c41] text-white font-bold'
                        : isDarkMode
                        ? 'bg-[#2b2b2b] border-[#404040] text-neutral-400 hover:bg-[#383838]'
                        : 'bg-[#f3f4f6] border-[#d1d5db] text-neutral-700 hover:bg-[#e2e8f0]'
                    }`}
                    title={`Zaznacz całą kolumnę ${colLetter} (kliknij i przeciągnij)`}
                  >
                    <span>{colLetter}</span>
                    {/* Column resize boundary handle */}
                    <div
                      onMouseDown={(e) => handleColResizeStart(colLetter, e)}
                      onDoubleClick={(e) => handleColAutoFit(colLetter, e)}
                      className="absolute top-0 -right-1 bottom-0 w-2.5 cursor-col-resize z-30 group/colres"
                      title={`Przeciągnij granicę, aby zmienić szerokość kolumny ${colLetter} (podwójny klik: dopasuj)`}
                    >
                      <div
                        className={`w-[2px] h-full mx-auto bg-[#107c41] transition-opacity ${
                          resizingCol?.col === colLetter ? 'opacity-100' : 'opacity-0 group-hover/colres:opacity-100'
                        }`}
                      />
                      {/* Live full-sheet guideline across all rows during column resize */}
                      {resizingCol?.col === colLetter && (
                        <div className="absolute top-0 left-[4px] w-[2px] h-[5000px] bg-[#107c41] pointer-events-none z-50 shadow-xs" />
                      )}
                    </div>
                  </th>
                );
              })}
            </tr>
          </thead>
        )}

        {/* Table Body: Rows & Cells */}
        <tbody>
          {Array.from({ length: DEFAULT_ROWS }).map((_, rIdx) => {
            const rowNumber = rIdx + 1;
            const isRowActive =
              rowNumber >= minRow &&
              rowNumber <= maxRow &&
              (!isEntireColSelected || isEntireSheetSelected);
            const rowHeight = getRowHeight(rowNumber);

            return (
              <tr key={rowNumber} style={{ height: rowHeight }}>
                {/* Row Number Header */}
                {showHeaders && (
                  <td
                    onMouseDown={(e) => handleRowHeaderMouseDown(rowNumber, e)}
                    onMouseEnter={() => handleRowHeaderMouseEnter(rowNumber)}
                    onContextMenu={(e) => handleRowContextMenu(rowNumber, e)}
                    className={`sticky left-0 z-10 w-10 border-b border-r text-center font-mono text-[11px] select-none relative cursor-pointer transition-colors ${
                      isRowActive
                        ? 'bg-[#107c41] text-white font-bold'
                        : isDarkMode
                        ? 'bg-[#2b2b2b] border-[#404040] text-neutral-400 hover:bg-[#383838]'
                        : 'bg-[#f3f4f6] border-[#d1d5db] text-neutral-600 hover:bg-[#e2e8f0]'
                    }`}
                    title={`Zaznacz cały wiersz ${rowNumber} (kliknij i przeciągnij)`}
                  >
                    <span>{rowNumber}</span>
                    {/* Row resize boundary handle - centered precisely on row border */}
                    <div
                      onMouseDown={(e) => handleRowResizeStart(rowNumber, e)}
                      onDoubleClick={(e) => handleRowResetHeight(rowNumber, e)}
                      className="absolute left-0 right-0 -bottom-[4px] h-[8px] cursor-row-resize z-30 group/rowres"
                      title={`Przeciągnij granicę, aby zmienić wysokość wiersza ${rowNumber} (podwójny klik: standardowa)`}
                    >
                      <div
                        className={`absolute left-0 right-0 top-[3px] h-[2px] bg-[#107c41] transition-opacity shadow-xs ${
                          resizingRow?.row === rowNumber ? 'opacity-100' : 'opacity-0 group-hover/rowres:opacity-100'
                        }`}
                      />
                      {/* Live full-sheet guideline across all columns during resize */}
                      {resizingRow?.row === rowNumber && (
                        <div className="absolute left-0 w-[5000px] top-[3px] h-[2px] bg-[#107c41] pointer-events-none z-50 shadow-xs" />
                      )}
                    </div>
                  </td>
                )}

                {/* Cells in Row */}
                {Array.from({ length: DEFAULT_COLS }).map((_, cIdx) => {
                  const colLetter = colIndexToLetter(cIdx);
                  const cellId = `${colLetter}${rowNumber}`;
                  const cell = sheet.cells[cellId];
                  const selected = isCellSelected(cIdx, rowNumber);
                  const active = isCellActive(cIdx, rowNumber);
                  const inFillArea = isCellInFillArea(cIdx, rowNumber);
                  const inFillExtension = isCellInFillExtension(cIdx, rowNumber);
                  const inFormulaPoint = isCellInFormulaPoint(cIdx, rowNumber);
                  const isTargetLesson = targetLessonCell === cellId;

                  // Formatting
                  const style = cell?.style || {};
                  const displayValue = cell
                    ? cell.computed !== undefined
                      ? formatCellValue(cell.computed, style)
                      : cell.raw
                    : '';

                  const colWidth = getColWidth(colLetter);
                  const borderCss = getCellBorderCss(style, isDarkMode);

                  // Is bottom-right corner of selection (for fill handle)
                  const isBottomRightCorner = cIdx === maxColIdx && rowNumber === maxRow;

                  // Selection perimeter edges
                  const isSelTopEdge = selected && rowNumber === minRow && !isEntireSheetSelected;
                  const isSelBottomEdge = selected && rowNumber === maxRow && !isEntireSheetSelected;
                  const isSelLeftEdge = selected && cIdx === minColIdx && !isEntireSheetSelected;
                  const isSelRightEdge = selected && cIdx === maxColIdx && !isEntireSheetSelected;

                  // Authentic Excel outer selection border (preserved during autofill!)
                  let selectionEdgeClasses = '';
                  if (selected && !inFormulaPoint && !isEntireSheetSelected) {
                    if (isSelTopEdge) selectionEdgeClasses += ' border-t-2 border-t-[#107c41]';
                    if (isSelBottomEdge) selectionEdgeClasses += ' border-b-2 border-b-[#107c41]';
                    if (isSelLeftEdge) selectionEdgeClasses += ' border-l-2 border-l-[#107c41]';
                    if (isSelRightEdge) selectionEdgeClasses += ' border-r-2 border-r-[#107c41]';
                  }

                  // Drag & Move Destination preview bounds
                  const targetMinCol = minColIdx + dragDelta.col;
                  const targetMaxCol = maxColIdx + dragDelta.col;
                  const targetMinRow = minRow + dragDelta.row;
                  const targetMaxRow = maxRow + dragDelta.row;

                  const isInMoveTarget =
                    isDraggingCells &&
                    (dragDelta.col !== 0 || dragDelta.row !== 0) &&
                    cIdx >= targetMinCol &&
                    cIdx <= targetMaxCol &&
                    rowNumber >= targetMinRow &&
                    rowNumber <= targetMaxRow;

                  const isMoveTargetTop = isInMoveTarget && rowNumber === targetMinRow;
                  const isMoveTargetBottom = isInMoveTarget && rowNumber === targetMaxRow;
                  const isMoveTargetLeft = isInMoveTarget && cIdx === targetMinCol;
                  const isMoveTargetRight = isInMoveTarget && cIdx === targetMaxCol;

                  let moveTargetClasses = '';
                  if (isInMoveTarget) {
                    if (isMoveTargetTop) moveTargetClasses += ' border-t-2 border-t-[#107c41] border-dashed';
                    if (isMoveTargetBottom) moveTargetClasses += ' border-b-2 border-b-[#107c41] border-dashed';
                    if (isMoveTargetLeft) moveTargetClasses += ' border-l-2 border-l-[#107c41] border-dashed';
                    if (isMoveTargetRight) moveTargetClasses += ' border-r-2 border-r-[#107c41] border-dashed';
                  }

                  // Authentic Excel Autofill Marquee Edges
                  let fillEdgeClasses = '';
                  if (inFillArea && fillBounds) {
                    const isTop = rowNumber === fillBounds.fullMinRow;
                    const isBottom = rowNumber === fillBounds.fullMaxRow;
                    const isLeft = cIdx === fillBounds.fullMinCol;
                    const isRight = cIdx === fillBounds.fullMaxCol;

                    if (isTop) fillEdgeClasses += ' border-t-2 border-t-neutral-800 dark:border-t-neutral-100 border-dashed';
                    if (isBottom) fillEdgeClasses += ' border-b-2 border-b-neutral-800 dark:border-b-neutral-100 border-dashed';
                    if (isLeft) fillEdgeClasses += ' border-l-2 border-l-neutral-800 dark:border-l-neutral-100 border-dashed';
                    if (isRight) fillEdgeClasses += ' border-r-2 border-r-neutral-800 dark:border-r-neutral-100 border-dashed';
                  }

                  return (
                    <td
                      key={cellId}
                      data-cell-id={cellId}
                      onMouseDown={(e) => handleCellMouseDown(cellId, e)}
                      onMouseEnter={() => handleCellMouseEnter(cellId)}
                      onContextMenu={(e) => handleCellContextMenu(cellId, e)}
                      onDoubleClick={() => startEditing()}
                      style={{
                        width: colWidth,
                        minWidth: colWidth,
                        maxWidth: colWidth,
                        overflow: 'hidden',
                        boxSizing: 'border-box',
                        backgroundColor: inFormulaPoint
                          ? 'rgba(15, 108, 189, 0.12)'
                          : inFillExtension
                          ? 'rgba(16, 124, 65, 0.18)'
                          : isInMoveTarget
                          ? 'rgba(16, 124, 65, 0.22)'
                          : style.bgColor
                          ? style.bgColor
                          : selected
                          ? 'rgba(16, 124, 65, 0.12)'
                          : undefined,
                        color: style.textColor,
                        fontWeight: style.bold ? 'bold' : 'normal',
                        fontStyle: style.italic ? 'italic' : 'normal',
                        textDecoration: style.underline ? 'underline' : 'none',
                        textAlign: style.align || (typeof cell?.computed === 'number' ? 'right' : 'left'),
                        fontSize: style.fontSize ? `${style.fontSize}px` : undefined,
                      }}
                      className={`relative px-1.5 py-0.5 text-xs truncate select-none tabular-numbers transition-colors ${
                        showGridlines && !inFillExtension && !inFormulaPoint
                          ? isDarkMode
                            ? 'border-b border-r border-[#303030]'
                            : 'border-b border-r border-[#e0e0e0]'
                          : ''
                      } ${
                        inFormulaPoint
                          ? 'excel-marching-ants ring-2 ring-[#0f6cbd] z-15'
                          : active && !inFillExtension && (!selectionRange || (minColIdx === maxColIdx && minRow === maxRow))
                          ? 'outline-2 outline-[#107c41] outline-offset-[-2px] z-10 bg-white dark:bg-[#222222]'
                          : selected && !inFillExtension
                          ? 'bg-[#107c41]/10'
                          : ''
                      } ${selectionEdgeClasses} ${moveTargetClasses} ${fillEdgeClasses} ${getBorderStyle(style)} ${
                        isTargetLesson ? 'pulse-target ring-2 ring-[#107c41] ring-inset' : ''
                      } ${isDraggingCells && selected ? 'opacity-65' : ''}`}
                    >
                      {/* Edge handles for Drag to Move (Excel 4-way arrow move cursor) */}
                      {!isEditing && !isDraggingFillHandle && !isEntireSheetSelected && selected && (
                        <>
                          {isSelTopEdge && (
                            <div
                              onMouseDown={(e) => handleDragBorderStart(e, cIdx, rowNumber)}
                              className="absolute -top-1 left-0 right-0 h-2 cursor-move z-25 group/topedge"
                              title="Przeciągnij krawędź zaznaczenia, aby przenieść komórki"
                            >
                              <div className="w-full h-full bg-[#107c41] opacity-0 group-hover/topedge:opacity-40 transition-opacity" />
                            </div>
                          )}
                          {isSelBottomEdge && (
                            <div
                              onMouseDown={(e) => handleDragBorderStart(e, cIdx, rowNumber)}
                              className={`absolute -bottom-1 left-0 ${isBottomRightCorner ? 'right-3' : 'right-0'} h-2 cursor-move z-25 group/botedge`}
                              title="Przeciągnij krawędź zaznaczenia, aby przenieść komórki"
                            >
                              <div className="w-full h-full bg-[#107c41] opacity-0 group-hover/botedge:opacity-40 transition-opacity" />
                            </div>
                          )}
                          {isSelLeftEdge && (
                            <div
                              onMouseDown={(e) => handleDragBorderStart(e, cIdx, rowNumber)}
                              className="absolute top-0 bottom-0 -left-1 w-2 cursor-move z-25 group/leftedge"
                              title="Przeciągnij krawędź zaznaczenia, aby przenieść komórki"
                            >
                              <div className="w-full h-full bg-[#107c41] opacity-0 group-hover/leftedge:opacity-40 transition-opacity" />
                            </div>
                          )}
                          {isSelRightEdge && (
                            <div
                              onMouseDown={(e) => handleDragBorderStart(e, cIdx, rowNumber)}
                              className={`absolute top-0 ${isBottomRightCorner ? 'bottom-3' : 'bottom-0'} -right-1 w-2 cursor-move z-25 group/rightedge`}
                              title="Przeciągnij krawędź zaznaczenia, aby przenieść komórki"
                            >
                              <div className="w-full h-full bg-[#107c41] opacity-0 group-hover/rightedge:opacity-40 transition-opacity" />
                            </div>
                          )}
                        </>
                      )}

                      {/* In-place Editing Input */}
                      {active && isEditing ? (
                        <input
                          ref={editInputRef}
                          type="text"
                          data-formula-input="true"
                          value={editValue}
                          onChange={(e) => {
                            setEditValue(e.target.value);
                            onCellTyping?.(e.target.value);
                          }}
                          onKeyDown={(e) => {
                            if (e.key === 'Enter') {
                              e.preventDefault();
                              e.stopPropagation();
                              const valToSave = editValue;
                              const targetRow = e.shiftKey ? Math.max(1, activeCoord.row - 1) : activeCoord.row + 1;
                              const nextCell = `${activeCoord.col}${targetRow}`;
                              isNavigatingViaKeyRef.current = true;
                              if (valToSave.startsWith('=')) {
                                setIsEditing(false);
                                onSubmitFormula?.(false);
                              } else {
                                onCellChange(activeCellId, valToSave);
                              }
                              onSelectCell(nextCell);
                              const nextRaw = sheet.cells[nextCell]?.raw || '';
                              setEditValue(nextRaw);
                              setIsEditing(true);
                              setTimeout(() => {
                                isNavigatingViaKeyRef.current = false;
                                if (editInputRef.current) {
                                  editInputRef.current.focus();
                                  editInputRef.current.select();
                                }
                              }, 40);
                            } else if (e.key === 'Tab') {
                              e.preventDefault();
                              e.stopPropagation();
                              const valToSave = editValue;
                              const targetColIdx = e.shiftKey ? Math.max(0, activeColIdx - 1) : activeColIdx + 1;
                              const nextCol = colIndexToLetter(targetColIdx);
                              const nextCell = `${nextCol}${activeCoord.row}`;
                              isNavigatingViaKeyRef.current = true;
                              if (valToSave.startsWith('=')) {
                                setIsEditing(false);
                                onSubmitFormula?.(true);
                              } else {
                                onCellChange(activeCellId, valToSave);
                              }
                              onSelectCell(nextCell);
                              const nextRaw = sheet.cells[nextCell]?.raw || '';
                              setEditValue(nextRaw);
                              setIsEditing(true);
                              setTimeout(() => {
                                isNavigatingViaKeyRef.current = false;
                                if (editInputRef.current) {
                                  editInputRef.current.focus();
                                  editInputRef.current.select();
                                }
                              }, 40);
                            } else if (e.key === 'Escape') {
                              e.preventDefault();
                              e.stopPropagation();
                              if (editValue.startsWith('=')) {
                                setIsEditing(false);
                                onCancelFormula?.();
                              } else {
                                cancelEdit();
                              }
                            }
                          }}
                          onBlur={() => {
                            if (isNavigatingViaKeyRef.current) return;
                            if (!isPointingFormula && !isFormulaMode) {
                              commitEdit();
                            }
                          }}
                          className={`absolute inset-0 w-full h-full px-1.5 text-xs font-mono outline-none border-2 border-[#107c41] z-30 ${
                            isDarkMode ? 'bg-[#1e1e1e] text-white' : 'bg-white text-black'
                          }`}
                        />
                      ) : (
                        <div
                          className="w-full h-full overflow-hidden text-ellipsis whitespace-nowrap"
                          title={displayValue}
                        >
                          {displayValue}
                        </div>
                      )}

                      {/* Crisp Excel Cell Borders Overlay (Guarantees all 4 borders top, bottom, left, right are drawn) */}
                      {borderCss.hasBorders && (
                        <div
                          className="pointer-events-none absolute -inset-[1px] z-5"
                          style={{
                            borderTop: borderCss.top || 'none',
                            borderBottom: borderCss.bottom || 'none',
                            borderLeft: borderCss.left || 'none',
                            borderRight: borderCss.right || 'none',
                          }}
                        />
                      )}

                      {/* Formula Point Reference Badge */}
                      {inFormulaPoint && formulaBounds && cIdx === formulaBounds.minCol && rowNumber === formulaBounds.minRow && (
                        <span className="absolute -top-3.5 right-1 z-30 px-1 py-0.2 bg-[#0f6cbd] text-white text-[9px] font-mono font-bold rounded shadow-xs pointer-events-none">
                          {formulaBounds.minCol === formulaBounds.maxCol && formulaBounds.minRow === formulaBounds.maxRow
                            ? cellId
                            : `${colIndexToLetter(formulaBounds.minCol)}${formulaBounds.minRow}:${colIndexToLetter(formulaBounds.maxCol)}${formulaBounds.maxRow}`}
                        </span>
                      )}

                      {/* Lesson Target Floating Pill */}
                      {isTargetLesson && !isEditing && (
                        <span className="absolute -top-3.5 left-1 z-30 px-1 py-0.2 bg-[#107c41] text-white text-[9px] font-bold rounded shadow-xs pointer-events-none">
                          🎯 Cel zadania
                        </span>
                      )}

                      {/* Fill Handle (Excel autofill small square at bottom-right of selection) */}
                      {isBottomRightCorner && !isEditing && !isDraggingFillHandle && !isDraggingCells && (
                        <div
                          onMouseDown={(e) => {
                            e.stopPropagation();
                            setIsDraggingFillHandle(true);
                            setFillTargetCell(cellId);
                          }}
                          title="Przeciągnij, aby automatycznie wypełnić lub skopiować formułę"
                          className="absolute -bottom-1 -right-1 w-2.5 h-2.5 bg-[#107c41] border border-white dark:border-black excel-fill-handle z-30 shadow-xs hover:scale-125 transition-transform cursor-crosshair"
                        />
                      )}
                    </td>
                  );
                })}
              </tr>
            );
          })}
        </tbody>
      </table>

      {/* --- AUTHENTIC EXCEL OFFICE 365 RIGHT-CLICK CONTEXT MENU --- */}
      {contextMenu && (
        <div
          style={{ top: contextMenu.y, left: contextMenu.x }}
          onClick={(e) => e.stopPropagation()}
          className="fixed z-50 w-64 bg-white dark:bg-[#2b2b2b] text-neutral-800 dark:text-neutral-100 rounded-lg shadow-2xl border border-neutral-300 dark:border-neutral-700 py-1.5 text-xs animate-in fade-in zoom-in-95 duration-100 select-none"
        >
          {/* Mini Formatting Toolbar at top of Context Menu */}
          <div className="flex items-center justify-between px-2.5 py-1.5 bg-neutral-100 dark:bg-[#202020] border-b border-neutral-200 dark:border-neutral-700 mb-1 rounded-t">
            <button
              onClick={() => {
                onApplyStyle?.({ bold: !sheet.cells[contextMenu.cellId]?.style?.bold });
                setContextMenu(null);
              }}
              title="Pogrubienie"
              className="p-1 hover:bg-neutral-200 dark:hover:bg-neutral-700 rounded"
            >
              <Bold size={13} />
            </button>
            <button
              onClick={() => {
                onApplyStyle?.({ italic: !sheet.cells[contextMenu.cellId]?.style?.italic });
                setContextMenu(null);
              }}
              title="Kursywa"
              className="p-1 hover:bg-neutral-200 dark:hover:bg-neutral-700 rounded"
            >
              <Italic size={13} />
            </button>
            <button
              onClick={() => {
                onApplyStyle?.({ underline: !sheet.cells[contextMenu.cellId]?.style?.underline });
                setContextMenu(null);
              }}
              title="Podkreślenie"
              className="p-1 hover:bg-neutral-200 dark:hover:bg-neutral-700 rounded"
            >
              <Underline size={13} />
            </button>
            <div className="h-3 w-px bg-neutral-300 dark:bg-neutral-600" />
            <button
              onClick={() => {
                onApplyStyle?.({ bgColor: '#fef3c7' });
                setContextMenu(null);
              }}
              title="Żółte wypełnienie tła"
              className="p-1 hover:bg-neutral-200 dark:hover:bg-neutral-700 rounded text-amber-500"
            >
              <PaintBucket size={13} />
            </button>
            <button
              onClick={() => {
                onApplyStyle?.({ textColor: '#dc2626' });
                setContextMenu(null);
              }}
              title="Czerwony kolor tekstu"
              className="p-1 hover:bg-neutral-200 dark:hover:bg-neutral-700 rounded text-rose-600 font-serif font-bold text-xs leading-none flex items-center justify-center w-5"
            >
              A
            </button>
            <button
              onClick={() => {
                onApplyStyle?.({ format: 'currency' });
                setContextMenu(null);
              }}
              title="Format waluty"
              className="p-1 hover:bg-neutral-200 dark:hover:bg-neutral-700 rounded font-bold text-[#107c41]"
            >
              <DollarSign size={13} />
            </button>
            <button
              onClick={() => {
                onApplyStyle?.({ format: 'percent' });
                setContextMenu(null);
              }}
              title="Format procentowy"
              className="p-1 hover:bg-neutral-200 dark:hover:bg-neutral-700 rounded font-bold"
            >
              <Percent size={13} />
            </button>
          </div>

          {/* Standard Office Context Menu Items */}
          <button
            onClick={() => {
              if (onCut) {
                onCut();
              } else {
                const cell = sheet.cells[contextMenu.cellId];
                if (cell) {
                  setClipboardData({ raw: cell.raw, style: cell.style, isCut: true, sourceCellId: contextMenu.cellId });
                }
              }
              setContextMenu(null);
            }}
            className="w-full flex items-center justify-between px-3 py-1.5 hover:bg-neutral-100 dark:hover:bg-[#383838] text-left transition-colors cursor-pointer"
          >
            <div className="flex items-center gap-2">
              <Scissors size={14} className="text-neutral-500" />
              <span>Wytnij</span>
            </div>
            <kbd className="text-[10px] text-neutral-400 font-mono">Ctrl+X</kbd>
          </button>

          <button
            onClick={() => {
              if (onCopy) {
                onCopy();
              } else {
                const cell = sheet.cells[contextMenu.cellId];
                if (cell) {
                  setClipboardData({ raw: cell.raw, style: cell.style });
                }
              }
              setContextMenu(null);
            }}
            className="w-full flex items-center justify-between px-3 py-1.5 hover:bg-neutral-100 dark:hover:bg-[#383838] text-left transition-colors cursor-pointer"
          >
            <div className="flex items-center gap-2">
              <Copy size={14} className="text-neutral-500" />
              <span>Kopiuj</span>
            </div>
            <kbd className="text-[10px] text-neutral-400 font-mono">Ctrl+C</kbd>
          </button>

          {/* Malarz formatów in Context Menu */}
          <button
            onClick={() => {
              onToggleFormatPainter?.(false);
              setContextMenu(null);
            }}
            className="w-full flex items-center justify-between px-3 py-1.5 hover:bg-neutral-100 dark:hover:bg-[#383838] text-left transition-colors cursor-pointer"
          >
            <div className="flex items-center gap-2">
              <Paintbrush size={14} className="text-amber-500" />
              <span>Malarz formatów</span>
            </div>
            <kbd className="text-[10px] text-neutral-400 font-mono">Ctrl+Shift+C</kbd>
          </button>

          <button
            onClick={() => {
              if (onPaste) {
                onPaste();
              } else if (clipboardData) {
                onCellChange(contextMenu.cellId, clipboardData.raw);
                if (clipboardData.isCut && clipboardData.sourceCellId) {
                  onCellChange(clipboardData.sourceCellId, '');
                  setClipboardData(null);
                }
              }
              setContextMenu(null);
            }}
            className="w-full flex items-center justify-between px-3 py-1.5 hover:bg-neutral-100 dark:hover:bg-[#383838] text-left transition-colors cursor-pointer"
          >
            <div className="flex items-center gap-2">
              <Clipboard size={14} className="text-neutral-500" />
              <span>Wklej</span>
            </div>
            <kbd className="text-[10px] text-neutral-400 font-mono">Ctrl+V</kbd>
          </button>

          <div className="my-1 border-t border-neutral-200 dark:border-neutral-700" />

          <button
            onClick={() => {
              onInsertRow?.();
              setContextMenu(null);
            }}
            className="w-full flex items-center gap-2 px-3 py-1.5 hover:bg-neutral-100 dark:hover:bg-[#383838] text-left transition-colors cursor-pointer"
          >
            <Plus size={14} className="text-[#107c41]" />
            <span>Wstaw wiersz</span>
          </button>

          <button
            onClick={() => {
              onDeleteRow?.();
              setContextMenu(null);
            }}
            className="w-full flex items-center gap-2 px-3 py-1.5 hover:bg-neutral-100 dark:hover:bg-[#383838] text-left text-rose-600 dark:text-rose-400 transition-colors cursor-pointer"
          >
            <Trash2 size={14} />
            <span>Usuń ten wiersz</span>
          </button>

          <button
            onClick={() => {
              if (onClearSelection) {
                onClearSelection();
              } else {
                onCellChange(contextMenu.cellId, '');
              }
              setContextMenu(null);
            }}
            className="w-full flex items-center justify-between px-3 py-1.5 hover:bg-neutral-100 dark:hover:bg-[#383838] text-left transition-colors cursor-pointer"
          >
            <span>Wyczyść zawartość</span>
            <kbd className="text-[10px] text-neutral-400 font-mono">Delete</kbd>
          </button>

          <div className="my-1 border-t border-neutral-200 dark:border-neutral-700" />

          <button
            onClick={() => {
              onAutoSum?.('SUMA');
              setContextMenu(null);
            }}
            className="w-full flex items-center gap-2 px-3 py-1.5 hover:bg-neutral-100 dark:hover:bg-[#383838] text-left text-[#107c41] dark:text-emerald-400 font-medium transition-colors"
          >
            <Sigma size={14} />
            <span>Autosumowanie (Σ)</span>
          </button>

          <button
            onClick={() => {
              onOpenChartModal?.();
              setContextMenu(null);
            }}
            className="w-full flex items-center gap-2 px-3 py-1.5 hover:bg-neutral-100 dark:hover:bg-[#383838] text-left transition-colors"
          >
            <BarChart2 size={14} className="text-[#0f6cbd]" />
            <span>Wstaw wykres z danych...</span>
          </button>
        </div>
      )}

      {/* Floating live size indicator while resizing column */}
      {resizingCol && (
        <div className="fixed top-24 left-1/2 -translate-x-1/2 z-50 bg-neutral-900/90 text-white text-xs font-mono px-3.5 py-1.5 rounded-md shadow-2xl border border-neutral-700 pointer-events-none flex items-center gap-2 backdrop-blur-xs select-none">
          <span>Kolumna <strong>{resizingCol.col}</strong></span>
          <span className="text-neutral-500">|</span>
          <span>Szerokość: <strong className="text-emerald-400">{resizingCol.currentWidth} px</strong></span>
        </div>
      )}

      {/* Floating live size indicator while resizing row */}
      {resizingRow && (
        <div className="fixed top-24 left-1/2 -translate-x-1/2 z-50 bg-neutral-900/90 text-white text-xs font-mono px-3.5 py-1.5 rounded-md shadow-2xl border border-neutral-700 pointer-events-none flex items-center gap-2 backdrop-blur-xs select-none">
          <span>Wiersz <strong>{resizingRow.row}</strong></span>
          <span className="text-neutral-500">|</span>
          <span>Wysokość: <strong className="text-emerald-400">{resizingRow.currentHeight} px</strong></span>
        </div>
      )}

      {/* Authentic Microsoft Excel Move Cells Confirmation Dialog */}
      {moveConfirmation && (
        <div
          className="fixed inset-0 z-50 flex items-center justify-center bg-black/40 backdrop-blur-xs"
          onClick={() => setMoveConfirmation(null)}
        >
          <div
            onClick={(e) => e.stopPropagation()}
            className="w-[420px] max-w-[92vw] bg-white dark:bg-[#2b2b2b] text-neutral-800 dark:text-neutral-100 rounded-lg shadow-2xl border border-neutral-300 dark:border-neutral-700 overflow-hidden animate-in fade-in zoom-in-95 duration-150 select-none"
          >
            {/* Title Bar */}
            <div className="flex items-center justify-between px-3.5 py-2 bg-neutral-100 dark:bg-[#202020] border-b border-neutral-200 dark:border-neutral-700">
              <div className="flex items-center gap-2">
                <div className="w-4 h-4 bg-[#107c41] rounded-xs flex items-center justify-center text-white text-[10px] font-bold">
                  X
                </div>
                <span className="font-semibold text-xs text-neutral-700 dark:text-neutral-200">
                  Microsoft Excel
                </span>
              </div>
              <button
                onClick={() => setMoveConfirmation(null)}
                className="text-neutral-500 hover:text-neutral-800 dark:hover:text-white p-0.5 rounded text-xs transition-colors cursor-pointer"
                title="Zamknij"
              >
                ✕
              </button>
            </div>

            {/* Dialog Content */}
            <div className="p-5 flex items-start gap-4">
              <div className="w-10 h-10 rounded-full bg-amber-100 dark:bg-amber-900/40 text-amber-600 dark:text-amber-400 flex items-center justify-center shrink-0">
                <AlertTriangle size={22} />
              </div>
              <div>
                <p className="text-sm font-semibold leading-snug">
                  W tym miejscu znajdują się już dane. Czy chcesz je zamienić?
                </p>
                <p className="text-xs text-neutral-500 dark:text-neutral-400 mt-2 font-mono">
                  Docelowy zakres:{' '}
                  <span className="font-bold text-neutral-700 dark:text-neutral-200">
                    {moveConfirmation.targetRange.startCol}{moveConfirmation.targetRange.startRow}
                    {moveConfirmation.targetRange.startCol !== moveConfirmation.targetRange.endCol ||
                    moveConfirmation.targetRange.startRow !== moveConfirmation.targetRange.endRow
                      ? `:${moveConfirmation.targetRange.endCol}${moveConfirmation.targetRange.endRow}`
                      : ''}
                  </span>{' '}
                  ({moveConfirmation.overwrittenCount}{' '}
                  {moveConfirmation.overwrittenCount === 1 ? 'komórka z danymi' : 'komórki z danymi'})
                </p>
              </div>
            </div>

            {/* Actions */}
            <div className="flex items-center justify-end gap-2.5 px-4 py-3 bg-neutral-50 dark:bg-[#242424] border-t border-neutral-200 dark:border-neutral-700">
              <button
                autoFocus
                onClick={() => {
                  onMoveCells?.(
                    moveConfirmation.sourceRange,
                    moveConfirmation.targetRange,
                    moveConfirmation.newActiveCellId
                  );
                  setMoveConfirmation(null);
                }}
                className="px-4 py-1.5 bg-[#107c41] hover:bg-[#0d6535] text-white text-xs font-semibold rounded shadow-xs transition-colors cursor-pointer"
              >
                OK
              </button>
              <button
                onClick={() => {
                  setMoveConfirmation(null);
                  sounds.playCellClick();
                }}
                className="px-4 py-1.5 bg-neutral-200 dark:bg-neutral-700 hover:bg-neutral-300 dark:hover:bg-neutral-600 text-neutral-800 dark:text-neutral-200 text-xs font-medium rounded transition-colors cursor-pointer"
              >
                Anuluj
              </button>
            </div>
          </div>
        </div>
      )}

      {/* AutoFill Copy Modifier Badge (+ indicator like Excel when Ctrl/Meta/Alt is held) */}
      {isDraggingFillHandle && isFillModifierActive && fillMousePos && (
        <div
          style={{
            top: fillMousePos.y + 14,
            left: fillMousePos.x + 14,
          }}
          className="fixed z-50 pointer-events-none px-2.5 py-1 bg-neutral-900/90 text-white text-[11px] font-sans rounded-md shadow-2xl border border-neutral-600 flex items-center gap-1.5 animate-in fade-in duration-75 select-none backdrop-blur-xs"
        >
          <span className="w-4 h-4 bg-emerald-600 text-white rounded flex items-center justify-center font-bold text-xs leading-none">+</span>
          <span className="font-medium">Kopiuj komórki (Ctrl)</span>
        </div>
      )}
    </div>
  );
};
