import React, { useState, useEffect } from 'react';
import { Plus, X, Edit2, ZoomIn, ZoomOut, Check, Wifi, WifiOff, Trash2 } from 'lucide-react';
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
  isFormulaMode?: boolean;
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
  isFormulaMode = false,
}) => {
  const [editingSheetId, setEditingSheetId] = useState<string | null>(null);
  const [editingName, setEditingName] = useState('');
  const [sheetToDelete, setSheetToDelete] = useState<SheetData | null>(null);
  const [tabContextMenu, setTabContextMenu] = useState<{ x: number; y: number; sheet: SheetData } | null>(null);

  // Close context menu on click outside
  useEffect(() => {
    const handleWindowClick = () => setTabContextMenu(null);
    window.addEventListener('click', handleWindowClick);
    return () => window.removeEventListener('click', handleWindowClick);
  }, []);

  const isUserSheet = (s: SheetData) => {
    return Boolean(s.isUserAdded || !['akademia-sheet', 'raport-sprzedazy', 'budzet-firmowy'].includes(s.id));
  };

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
        {isFormulaMode ? (
          <span
            title="Tryb wskazywania komórki (kliknij komórkę do formuły, Enter - zatwierdź, Escape - wyjdź)"
            className="hidden sm:inline-flex items-center gap-1.5 px-2 py-0.5 text-[11px] font-bold text-blue-700 dark:text-blue-300 bg-blue-100 dark:bg-blue-950/60 rounded border border-blue-300 dark:border-blue-700 shadow-xs"
          >
            <span className="w-1.5 h-1.5 rounded-full bg-blue-500 animate-pulse" />
            <span>Wskaż</span>
            <span className="text-[9px] font-normal opacity-75">(Enter / Esc)</span>
          </span>
        ) : (
          <span className="hidden sm:inline px-2 text-[11px] font-semibold text-neutral-500">Gotowy</span>
        )}

        <div className="h-4 w-px bg-neutral-300 dark:bg-neutral-700 hidden sm:block" />

        {/* Sheet Tabs */}
        {sheets.map((sheet) => {
          const isActive = sheet.id === activeSheetId;
          const deletable = isUserSheet(sheet) && sheets.length > 1;

          return (
            <div
              key={sheet.id}
              className={`group relative flex items-center gap-1.5 px-3 py-1 rounded-t-sm border-b-2 text-xs font-medium cursor-pointer transition-colors ${
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
              onContextMenu={(e) => {
                e.preventDefault();
                e.stopPropagation();
                setTabContextMenu({ x: e.clientX, y: e.clientY, sheet });
              }}
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

              {/* Delete Sheet affordance for user-added sheets */}
              {deletable && (
                <button
                  type="button"
                  onClick={(e) => {
                    e.stopPropagation();
                    setSheetToDelete(sheet);
                  }}
                  title="Usuń ten arkusz"
                  className="opacity-70 hover:opacity-100 hover:text-rose-500 hover:bg-rose-100 dark:hover:bg-rose-950/60 p-0.5 rounded transition-all"
                >
                  <X size={12} />
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

      {/* Tab Context Menu (Excel Style) */}
      {tabContextMenu && (
        <div
          className={`fixed z-50 py-1 rounded-md shadow-2xl border text-xs min-w-[170px] animate-in fade-in zoom-in-95 duration-100 ${
            isDarkMode ? 'bg-[#252525] border-[#404040] text-neutral-200' : 'bg-white border-neutral-300 text-neutral-800'
          }`}
          style={{
            top: Math.max(10, tabContextMenu.y - 110),
            left: Math.min(window.innerWidth - 180, tabContextMenu.x),
          }}
          onClick={(e) => e.stopPropagation()}
        >
          <button
            onClick={() => {
              handleStartRename(tabContextMenu.sheet);
              setTabContextMenu(null);
            }}
            className="w-full text-left px-3 py-1.5 hover:bg-neutral-100 dark:hover:bg-neutral-700/60 flex items-center gap-2 cursor-pointer"
          >
            <Edit2 size={12} />
            <span>Zmień nazwę</span>
          </button>

          {isUserSheet(tabContextMenu.sheet) && sheets.length > 1 ? (
            <button
              onClick={() => {
                setSheetToDelete(tabContextMenu.sheet);
                setTabContextMenu(null);
              }}
              className="w-full text-left px-3 py-1.5 hover:bg-rose-50 dark:hover:bg-rose-950/40 text-rose-600 dark:text-rose-400 flex items-center gap-2 cursor-pointer"
            >
              <Trash2 size={12} />
              <span>Usuń arkusz</span>
            </button>
          ) : (
            <div className="px-3 py-1.5 text-neutral-400 dark:text-neutral-500 text-[11px] italic cursor-not-allowed">
              Arkusz systemowy (zablokowany)
            </div>
          )}

          <div className="h-px bg-neutral-200 dark:bg-neutral-700 my-1" />

          <button
            onClick={() => {
              onAddSheet();
              setTabContextMenu(null);
            }}
            className="w-full text-left px-3 py-1.5 hover:bg-neutral-100 dark:hover:bg-neutral-700/60 flex items-center gap-2 cursor-pointer"
          >
            <Plus size={12} />
            <span>Nowy arkusz</span>
          </button>
        </div>
      )}

      {/* In-App Sheet Deletion Confirmation Modal */}
      {sheetToDelete && (
        <div
          className="fixed inset-0 z-50 flex items-center justify-center bg-black/50 backdrop-blur-xs p-4 animate-in fade-in duration-150"
          onClick={() => setSheetToDelete(null)}
        >
          <div
            className={`w-full max-w-sm rounded-lg shadow-2xl border p-5 ${
              isDarkMode ? 'bg-[#252525] border-[#3f3f3f] text-white' : 'bg-white border-neutral-300 text-neutral-900'
            }`}
            onClick={(e) => e.stopPropagation()}
          >
            <div className="flex items-center gap-2 text-rose-600 dark:text-rose-400 font-bold text-sm mb-2">
              <Trash2 size={18} />
              <span>Usunąć arkusz?</span>
            </div>
            <p className="text-xs text-neutral-600 dark:text-neutral-300 leading-relaxed mb-4">
              Czy na pewno chcesz usunąć arkusz <strong>„{sheetToDelete.name}”</strong>?
              Wszystkie wprowadzone w nim dane zostaną usunięte. Możesz cofnąć tę akcję za pomocą historii zmian.
            </p>
            <div className="flex items-center justify-end gap-2 text-xs font-medium">
              <button
                type="button"
                onClick={() => setSheetToDelete(null)}
                className={`px-3 py-1.5 rounded border transition-colors cursor-pointer ${
                  isDarkMode
                    ? 'border-neutral-600 hover:bg-neutral-800 text-neutral-300'
                    : 'border-neutral-300 hover:bg-neutral-100 text-neutral-700'
                }`}
              >
                Anuluj
              </button>
              <button
                type="button"
                onClick={() => {
                  onDeleteSheet(sheetToDelete.id);
                  setSheetToDelete(null);
                }}
                className="px-3.5 py-1.5 rounded bg-rose-600 hover:bg-rose-700 text-white shadow-xs transition-colors cursor-pointer"
              >
                Usuń arkusz
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
};
