import React, { useState, useRef, useEffect, useCallback } from 'react';
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
import { sounds } from '../utils/audio';

interface SpreadsheetGridProps {
  sheet: SheetData;
  activeCellId: string;
  selectionRange: SelectionRange | null;
  onSelectCell: (cellId: string) => void;
  onSelectRange: (range: SelectionRange | null) => void;
  onCellChange: (cellId: string, rawValue: string) => void;
  onAutoFill: (sourceRange: SelectionRange, targetRange: SelectionRange) => void;
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
}) => {
  const [isEditing, setIsEditing] = useState(false);
  const [editValue, setEditValue] = useState('');
  const [isSelecting, setIsSelecting] = useState(false);

  // Autofill Dragging States
  const [isDraggingFillHandle, setIsDraggingFillHandle] = useState(false);
  const [fillTargetCell, setFillTargetCell] = useState<string | null>(null);

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

  // Column and row sizes
  const getColWidth = useCallback((col: string) => sheet.colWidths?.[col] || DEFAULT_COL_WIDTH, [sheet.colWidths]);
  const getRowHeight = useCallback((row: number) => sheet.rowHeights?.[row] || DEFAULT_ROW_HEIGHT, [sheet.rowHeights]);

  // Active cell coordinate
  const activeCoord = parseCellAddress(activeCellId) || { col: 'A', row: 1 };
  const activeColIdx = letterToColIndex(activeCoord.col);

  // Focus editing input when editing starts
  useEffect(() => {
    if (isEditing && editInputRef.current) {
      editInputRef.current.focus();
      editInputRef.current.select();
    }
  }, [isEditing]);

  // Close context menu on outside click
  useEffect(() => {
    const handleGlobalClick = () => {
      setContextMenu(null);
    };
    window.addEventListener('click', handleGlobalClick);
    return () => window.removeEventListener('click', handleGlobalClick);
  }, []);

  // Enter edit mode
  const startEditing = useCallback((initialVal?: string) => {
    const currentRaw = sheet.cells[activeCellId]?.raw || '';
    setEditValue(initialVal !== undefined ? initialVal : currentRaw);
    setIsEditing(true);
    setContextMenu(null);
  }, [sheet.cells, activeCellId]);

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
        commitEdit();
        // Move down
        const nextCell = `${activeCoord.col}${activeCoord.row + 1}`;
        onSelectCell(nextCell);
      } else if (e.key === 'Tab') {
        e.preventDefault();
        commitEdit();
        // Move right
        const nextCol = colIndexToLetter(activeColIdx + 1);
        onSelectCell(`${nextCol}${activeCoord.row}`);
      } else if (e.key === 'Escape') {
        e.preventDefault();
        cancelEdit();
      }
      return;
    }

    // Global Shortcuts
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
      const cell = sheet.cells[activeCellId];
      if (cell) {
        setClipboardData({ raw: cell.raw, style: cell.style });
        onShortcutUsed?.('Ctrl+C');
      }
      return;
    }
    if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'x') {
      e.preventDefault();
      const cell = sheet.cells[activeCellId];
      if (cell) {
        setClipboardData({ raw: cell.raw, style: cell.style, isCut: true, sourceCellId: activeCellId });
        onShortcutUsed?.('Ctrl+X');
      }
      return;
    }
    if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'v') {
      e.preventDefault();
      if (clipboardData) {
        onCellChange(activeCellId, clipboardData.raw);
        if (clipboardData.isCut && clipboardData.sourceCellId) {
          onCellChange(clipboardData.sourceCellId, '');
          setClipboardData(null);
        }
        onShortcutUsed?.('Ctrl+V');
      }
      return;
    }

    if (e.key === 'F2') {
      e.preventDefault();
      startEditing();
      return;
    }

    if (e.key === 'Delete' || e.key === 'Backspace') {
      e.preventDefault();
      onCellChange(activeCellId, '');
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
    if (isEditing) commitEdit();
    setContextMenu(null);

    sounds.playCellClick();

    if (e.shiftKey) {
      // Expand selection range from active to clicked
      const clicked = parseCellAddress(cellId);
      if (clicked) {
        onSelectRange({
          startCol: activeCoord.col,
          startRow: activeCoord.row,
          endCol: clicked.col,
          endRow: clicked.row,
        });
      }
    } else {
      onSelectCell(cellId);
      onSelectRange(null);
      setIsSelecting(true);
    }
  };

  const handleCellMouseEnter = (cellId: string) => {
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
    if (isSelecting) {
      setIsSelecting(false);
    }
    if (isDraggingFillHandle && fillTargetCell) {
      setIsDraggingFillHandle(false);
      const sourceRange: SelectionRange = selectionRange || {
        startCol: activeCoord.col,
        startRow: activeCoord.row,
        endCol: activeCoord.col,
        endRow: activeCoord.row,
      };

      const targetCoord = parseCellAddress(fillTargetCell);
      if (targetCoord) {
        const fullRange: SelectionRange = {
          startCol: sourceRange.startCol,
          startRow: sourceRange.startRow,
          endCol: targetCoord.col,
          endRow: targetCoord.row,
        };
        onAutoFill(sourceRange, fullRange);
      }
      setFillTargetCell(null);
    } else if (isDraggingFillHandle) {
      setIsDraggingFillHandle(false);
      setFillTargetCell(null);
    }
  }, [isSelecting, isDraggingFillHandle, fillTargetCell, selectionRange, activeCoord, onAutoFill]);

  useEffect(() => {
    window.addEventListener('mouseup', handleGlobalMouseUp);
    return () => window.removeEventListener('mouseup', handleGlobalMouseUp);
  }, [handleGlobalMouseUp]);

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

  // Calculate Autofill dragging bounds (for real Excel dashed marquee preview)
  const getFillBounds = useCallback(() => {
    if (!isDraggingFillHandle || !fillTargetCell) return null;

    const targetCoord = parseCellAddress(fillTargetCell);
    if (!targetCoord) return null;

    const targetColIdx = letterToColIndex(targetCoord.col);

    const fullMinCol = Math.min(minColIdx, targetColIdx);
    const fullMaxCol = Math.max(maxColIdx, targetColIdx);
    const fullMinRow = Math.min(minRow, targetCoord.row);
    const fullMaxRow = Math.max(maxRow, targetCoord.row);

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

  return (
    <div
      ref={gridContainerRef}
      tabIndex={0}
      onKeyDown={handleKeyDown}
      className={`relative flex-1 overflow-auto select-none outline-none excel-grid transition-colors ${
        isDarkMode ? 'bg-[#181818] text-neutral-200' : 'bg-white text-neutral-900'
      }`}
      style={{
        transform: `scale(${zoomLevel / 100})`,
        transformOrigin: 'top left',
      }}
    >
      <table className="border-collapse table-fixed w-max">
        {/* Table Header: Column letters */}
        {showHeaders && (
          <thead>
            <tr className="sticky top-0 z-20 shadow-xs">
              {/* Corner Cell */}
              <th className={`w-10 h-6 border-b border-r text-center font-normal text-[11px] ${
                isDarkMode ? 'bg-[#2b2b2b] border-[#404040] text-neutral-400' : 'bg-[#f3f4f6] border-[#d1d5db] text-neutral-500'
              }`}>
                ◢
              </th>
              {Array.from({ length: DEFAULT_COLS }).map((_, cIdx) => {
                const colLetter = colIndexToLetter(cIdx);
                const isColActive = cIdx >= minColIdx && cIdx <= maxColIdx;
                const width = getColWidth(colLetter);

                return (
                  <th
                    key={colLetter}
                    style={{ width, minWidth: width, maxWidth: width }}
                    className={`h-6 border-b border-r text-center font-medium text-[11px] px-1 select-none font-mono ${
                      isColActive
                        ? 'bg-[#107c41] text-white font-bold'
                        : isDarkMode
                        ? 'bg-[#2b2b2b] border-[#404040] text-neutral-400 hover:bg-[#333333]'
                        : 'bg-[#f3f4f6] border-[#d1d5db] text-neutral-700 hover:bg-[#e5e7eb]'
                    }`}
                  >
                    {colLetter}
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
            const isRowActive = rowNumber >= minRow && rowNumber <= maxRow;
            const rowHeight = getRowHeight(rowNumber);

            return (
              <tr key={rowNumber} style={{ height: rowHeight }}>
                {/* Row Number Header */}
                {showHeaders && (
                  <td
                    className={`sticky left-0 z-10 w-10 border-b border-r text-center font-mono text-[11px] select-none ${
                      isRowActive
                        ? 'bg-[#107c41] text-white font-bold'
                        : isDarkMode
                        ? 'bg-[#2b2b2b] border-[#404040] text-neutral-400 hover:bg-[#333333]'
                        : 'bg-[#f3f4f6] border-[#d1d5db] text-neutral-600 hover:bg-[#e5e7eb]'
                    }`}
                  >
                    {rowNumber}
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
                  const isTargetLesson = targetLessonCell === cellId;

                  // Formatting
                  const style = cell?.style || {};
                  const displayValue = cell
                    ? cell.computed !== undefined
                      ? formatCellValue(cell.computed, style)
                      : cell.raw
                    : '';

                  // Is bottom-right corner of selection (for fill handle)
                  const isBottomRightCorner = cIdx === maxColIdx && rowNumber === maxRow;

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
                        width: getColWidth(colLetter),
                        backgroundColor: inFillArea
                          ? 'rgba(16, 124, 65, 0.16)'
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
                        showGridlines && !inFillArea
                          ? isDarkMode
                            ? 'border-b border-r border-[#303030]'
                            : 'border-b border-r border-[#e0e0e0]'
                          : ''
                      } ${
                        active && !inFillArea
                          ? 'outline-2 outline-[#107c41] outline-offset-[-2px] z-10 bg-white dark:bg-[#222222]'
                          : selected && !inFillArea
                          ? 'bg-[#107c41]/10'
                          : ''
                      } ${fillEdgeClasses} ${getBorderStyle(style)} ${
                        isTargetLesson ? 'pulse-target ring-2 ring-[#107c41] ring-inset' : ''
                      }`}
                    >
                      {/* In-place Editing Input */}
                      {active && isEditing ? (
                        <input
                          ref={editInputRef}
                          type="text"
                          value={editValue}
                          onChange={(e) => setEditValue(e.target.value)}
                          onBlur={commitEdit}
                          className={`absolute inset-0 w-full h-full px-1.5 text-xs font-mono outline-none border-2 border-[#107c41] z-30 ${
                            isDarkMode ? 'bg-[#1e1e1e] text-white' : 'bg-white text-black'
                          }`}
                        />
                      ) : (
                        <span>{displayValue}</span>
                      )}

                      {/* Lesson Target Floating Pill */}
                      {isTargetLesson && !isEditing && (
                        <span className="absolute -top-3.5 left-1 z-30 px-1 py-0.2 bg-[#107c41] text-white text-[9px] font-bold rounded shadow-xs pointer-events-none">
                          🎯 Cel zadania
                        </span>
                      )}

                      {/* Fill Handle (Excel autofill small square at bottom-right of selection) */}
                      {isBottomRightCorner && !isEditing && !isDraggingFillHandle && (
                        <div
                          onMouseDown={(e) => {
                            e.stopPropagation();
                            setIsDraggingFillHandle(true);
                            setFillTargetCell(cellId);
                          }}
                          title="Przeciągnij, aby automatycznie wypełnić lub skopiować formułę"
                          className="absolute -bottom-1 -right-1 w-2.5 h-2.5 bg-[#107c41] border border-white dark:border-black excel-fill-handle z-20 shadow-xs hover:scale-125 transition-transform"
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
              title="Żółte wypełnienie"
              className="p-1 hover:bg-neutral-200 dark:hover:bg-neutral-700 rounded text-amber-500"
            >
              <PaintBucket size={13} />
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
              const cell = sheet.cells[contextMenu.cellId];
              if (cell) {
                setClipboardData({ raw: cell.raw, style: cell.style, isCut: true, sourceCellId: contextMenu.cellId });
              }
              setContextMenu(null);
            }}
            className="w-full flex items-center justify-between px-3 py-1.5 hover:bg-neutral-100 dark:hover:bg-[#383838] text-left transition-colors"
          >
            <div className="flex items-center gap-2">
              <Scissors size={14} className="text-neutral-500" />
              <span>Wytnij</span>
            </div>
            <kbd className="text-[10px] text-neutral-400 font-mono">Ctrl+X</kbd>
          </button>

          <button
            onClick={() => {
              const cell = sheet.cells[contextMenu.cellId];
              if (cell) {
                setClipboardData({ raw: cell.raw, style: cell.style });
              }
              setContextMenu(null);
            }}
            className="w-full flex items-center justify-between px-3 py-1.5 hover:bg-neutral-100 dark:hover:bg-[#383838] text-left transition-colors"
          >
            <div className="flex items-center gap-2">
              <Copy size={14} className="text-neutral-500" />
              <span>Kopiuj</span>
            </div>
            <kbd className="text-[10px] text-neutral-400 font-mono">Ctrl+C</kbd>
          </button>

          <button
            onClick={() => {
              if (clipboardData) {
                onCellChange(contextMenu.cellId, clipboardData.raw);
                if (clipboardData.isCut && clipboardData.sourceCellId) {
                  onCellChange(clipboardData.sourceCellId, '');
                  setClipboardData(null);
                }
              }
              setContextMenu(null);
            }}
            disabled={!clipboardData}
            className={`w-full flex items-center justify-between px-3 py-1.5 text-left transition-colors ${
              clipboardData
                ? 'hover:bg-neutral-100 dark:hover:bg-[#383838]'
                : 'opacity-40 cursor-not-allowed'
            }`}
          >
            <div className="flex items-center gap-2">
              <Clipboard size={14} className="text-neutral-500" />
              <span>Wklej wszystko</span>
            </div>
            <kbd className="text-[10px] text-neutral-400 font-mono">Ctrl+V</kbd>
          </button>

          <div className="my-1 border-t border-neutral-200 dark:border-neutral-700" />

          <button
            onClick={() => {
              onInsertRow?.();
              setContextMenu(null);
            }}
            className="w-full flex items-center gap-2 px-3 py-1.5 hover:bg-neutral-100 dark:hover:bg-[#383838] text-left transition-colors"
          >
            <Plus size={14} className="text-[#107c41]" />
            <span>Wstaw wiersz</span>
          </button>

          <button
            onClick={() => {
              onDeleteRow?.();
              setContextMenu(null);
            }}
            className="w-full flex items-center gap-2 px-3 py-1.5 hover:bg-neutral-100 dark:hover:bg-[#383838] text-left text-rose-600 dark:text-rose-400 transition-colors"
          >
            <Trash2 size={14} />
            <span>Usuń ten wiersz</span>
          </button>

          <button
            onClick={() => {
              onCellChange(contextMenu.cellId, '');
              setContextMenu(null);
            }}
            className="w-full flex items-center justify-between px-3 py-1.5 hover:bg-neutral-100 dark:hover:bg-[#383838] text-left transition-colors"
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
    </div>
  );
};
