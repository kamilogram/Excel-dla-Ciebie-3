import React, { useState } from 'react';
import { Plus, X, Edit2, ZoomIn, ZoomOut, Check, Wifi, WifiOff } from 'lucide-react';
import { SheetData, SelectionRange } from '../types/excel';
import { letterToColIndex, parseNumericValue } from '../utils/formulaEngine';

interface StatusBarProps {
  sheets: SheetData[];
  activeSheetId: string;
  onSelectSheet: (sheetId: string) => void;
  onAddSheet: () => void;
  onDeleteSheet: (sheetId: string) => void;
  onRenameSheet: (sheetId: string, newName: string) => void;
  activeSheet: SheetData;
  selectionRange: SelectionRange | null;
  activeCellId: string;
  zoomLevel: number;
  onSetZoom: (zoom: number) => void;
  isOnline: boolean;
  isDarkMode: boolean;
}

export const StatusBar: React.FC<StatusBarProps> = ({
  sheets,
  activeSheetId,
  onSelectSheet,
  onAddSheet,
  onDeleteSheet,
  onRenameSheet,
  activeSheet,
  selectionRange,
  activeCellId,
  zoomLevel,
  onSetZoom,
  isOnline,
  isDarkMode,
}) => {
  const [editingSheetId, setEditingSheetId] = useState<string | null>(null);
  const [editingName, setEditingName] = useState('');

  // Calculate live statistics for selected cells
  const getSelectionStats = () => {
    let cellsToCheck: string[] = [];

    if (!selectionRange) {
      cellsToCheck = [activeCellId];
    } else {
      const c1 = letterToColIndex(selectionRange.startCol);
      const c2 = letterToColIndex(selectionRange.endCol);
      const minC = Math.min(c1, c2);
      const maxC = Math.max(c1, c2);

      const minR = Math.min(selectionRange.startRow, selectionRange.endRow);
      const maxR = Math.max(selectionRange.startRow, selectionRange.endRow);

      for (let r = minR; r <= maxR; r++) {
        for (let c = minC; c <= maxC; c++) {
          const letter = String.fromCharCode(c + 65);
          cellsToCheck.push(`${letter}${r}`);
        }
      }
    }

    const numericValues: number[] = [];
    let countNonEmpty = 0;

    for (const id of cellsToCheck) {
      const cell = activeSheet.cells[id];
      if (cell && cell.raw !== '' && cell.raw !== undefined) {
        countNonEmpty++;
        const val = cell.computed !== undefined ? cell.computed : cell.raw;
        const num = parseNumericValue(val);
        if (num !== 0 || val === '0' || val === 0) {
          numericValues.push(num);
        }
      }
    }

    if (numericValues.length <= 1 && countNonEmpty <= 1) {
      return null;
    }

    const sum = numericValues.reduce((a, b) => a + b, 0);
    const avg = numericValues.length > 0 ? sum / numericValues.length : 0;
    const min = numericValues.length > 0 ? Math.min(...numericValues) : 0;
    const max = numericValues.length > 0 ? Math.max(...numericValues) : 0;

    return {
      count: countNonEmpty,
      numCount: numericValues.length,
      sum,
      avg,
      min,
      max,
    };
  };

  const stats = getSelectionStats();

  const handleStartRename = (sheet: SheetData) => {
    setEditingSheetId(sheet.id);
    setEditingName(sheet.name);
  };

  const handleFinishRename = (id: string) => {
    if (editingName.trim()) {
      onRenameSheet(id, editingName.trim());
    }
    setEditingSheetId(null);
  };

  return (
    <div className={`flex flex-col md:flex-row items-stretch md:items-center justify-between px-2 py-0.5 border-t text-xs select-none shrink-0 transition-colors ${
      isDarkMode ? 'bg-[#202020] border-[#333333] text-neutral-300' : 'bg-[#f3f4f6] border-[#d1d5db] text-neutral-700'
    }`}>
      {/* Left: Sheet Tabs & Add Button */}
      <div className="flex items-center overflow-x-auto gap-1 scrollbar-none py-1">
        <span className="hidden sm:inline px-2 text-[11px] font-semibold text-neutral-500">Gotowy</span>

        <div className="h-4 w-px bg-neutral-300 dark:bg-neutral-700 hidden sm:block" />

        {/* Sheet Tabs */}
        {sheets.map((sheet) => {
          const isActive = sheet.id === activeSheetId;
          return (
            <div
              key={sheet.id}
              className={`group flex items-center gap-1.5 px-3 py-1 rounded-t-sm border-b-2 text-xs font-medium cursor-pointer transition-colors ${
                isActive
                  ? isDarkMode
                    ? 'bg-[#2b2b2b] text-white border-[#107c41] shadow-xs'
                    : 'bg-white text-[#107c41] border-[#107c41] shadow-xs'
                  : isDarkMode
                  ? 'bg-[#202020] text-neutral-400 border-transparent hover:bg-[#2b2b2b]'
                  : 'bg-[#f3f4f6] text-neutral-600 border-transparent hover:bg-neutral-200'
              }`}
              onClick={() => onSelectSheet(sheet.id)}
              onDoubleClick={() => handleStartRename(sheet)}
            >
              {editingSheetId === sheet.id ? (
                <input
                  type="text"
                  value={editingName}
                  onChange={(e) => setEditingName(e.target.value)}
                  onBlur={() => handleFinishRename(sheet.id)}
                  onKeyDown={(e) => e.key === 'Enter' && handleFinishRename(sheet.id)}
                  autoFocus
                  className="w-24 px-1 py-0.5 text-xs bg-white text-black rounded outline-none border border-[#107c41]"
                />
              ) : (
                <span className="truncate max-w-[130px]">{sheet.name}</span>
              )}

              {/* Delete Sheet affordance */}
              {sheets.length > 1 && (
                <button
                  onClick={(e) => {
                    e.stopPropagation();
                    if (confirm(`Czy na pewno chcesz usunąć arkusz "${sheet.name}"?`)) {
                      onDeleteSheet(sheet.id);
                    }
                  }}
                  title="Usuń arkusz"
                  className="opacity-0 group-hover:opacity-100 hover:text-rose-500 p-0.5 rounded transition-opacity"
                >
                  <X size={11} />
                </button>
              )}
            </div>
          );
        })}

        {/* Add Sheet Button */}
        <button
          onClick={onAddSheet}
          title="Wstaw nowy arkusz (+)"
          className="p-1 rounded hover:bg-neutral-200 dark:hover:bg-neutral-700 text-neutral-600 dark:text-neutral-400 transition-colors ml-1"
        >
          <Plus size={14} />
        </button>
      </div>

      {/* Middle & Right: Real-time Selection Statistics & Zoom */}
      <div className="flex items-center justify-between md:justify-end gap-3 py-0.5 overflow-x-auto text-[11px] tabular-numbers">
        {/* Dynamic Statistics */}
        {stats && (
          <div className="flex items-center gap-2.5 text-neutral-600 dark:text-neutral-300 px-2 py-0.5 bg-neutral-200/50 dark:bg-neutral-800 rounded">
            <span>ŚREDNIA: <strong className="font-semibold text-neutral-900 dark:text-white">{stats.avg.toLocaleString('pl-PL', { maximumFractionDigits: 2 })}</strong></span>
            <span>LICZNIK: <strong className="font-semibold text-neutral-900 dark:text-white">{stats.count}</strong></span>
            {stats.numCount > 1 && (
              <>
                <span className="hidden lg:inline">MIN: <strong className="font-semibold text-neutral-900 dark:text-white">{stats.min.toLocaleString('pl-PL')}</strong></span>
                <span className="hidden lg:inline">MAKS: <strong className="font-semibold text-neutral-900 dark:text-white">{stats.max.toLocaleString('pl-PL')}</strong></span>
                <span>SUMA: <strong className="font-semibold text-[#107c41] dark:text-emerald-400">{stats.sum.toLocaleString('pl-PL', { maximumFractionDigits: 2 })} zł</strong></span>
              </>
            )}
          </div>
        )}

        {/* Zoom Controls */}
        <div className="flex items-center gap-1.5 shrink-0 pl-2">
          <button
            onClick={() => onSetZoom(Math.max(50, zoomLevel - 15))}
            title="Pomniejsz"
            className="p-1 hover:bg-neutral-200 dark:hover:bg-neutral-700 rounded"
          >
            <ZoomOut size={13} />
          </button>
          <span className="w-10 text-center font-medium">{zoomLevel}%</span>
          <button
            onClick={() => onSetZoom(Math.min(180, zoomLevel + 15))}
            title="Powiększ"
            className="p-1 hover:bg-neutral-200 dark:hover:bg-neutral-700 rounded"
          >
            <ZoomIn size={13} />
          </button>
        </div>
      </div>
    </div>
  );
};
