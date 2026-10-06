import React, { useState, useEffect, useCallback, useMemo } from 'react';
import confetti from 'canvas-confetti';
import {
  SheetData,
  SelectionRange,
  CellStyle,
  LearningTask,
  UserBadge,
  LeaderboardUser,
  UserProfile,
} from './types/excel';
import {
  INITIAL_SHEETS,
  INITIAL_BADGES,
  INITIAL_LEADERBOARD,
  INITIAL_USER_PROFILE,
} from './data/initialSheets';
import { EXCEL_LESSONS } from './data/lessons';
import {
  colIndexToLetter,
  letterToColIndex,
  parseCellAddress,
  evaluateFormula,
  getRangeCells,
} from './utils/formulaEngine';
import { sounds } from './utils/audio';
import { ExcelHeader } from './components/ExcelHeader';
import { ExcelRibbon, RibbonTab } from './components/ExcelRibbon';
import { FormulaBar } from './components/FormulaBar';
import { SpreadsheetGrid } from './components/SpreadsheetGrid';
import { StatusBar } from './components/StatusBar';
import { LearningPanel } from './components/LearningPanel';
import { LeaderboardModal } from './components/LeaderboardModal';
import { ProfileModal } from './components/ProfileModal';
import { DailyMotivationModal } from './components/DailyMotivationModal';
import { ChartModal } from './components/ChartModal';
import { SpeedChallengeModal } from './components/SpeedChallengeModal';
import { ShortcutsModal } from './components/ShortcutsModal';

const STORAGE_KEY_SHEETS = 'excel_pro_365_sheets_v1';
const STORAGE_KEY_PROFILE = 'excel_pro_365_profile_v1';
const STORAGE_KEY_BADGES = 'excel_pro_365_badges_v1';
const STORAGE_KEY_DARK = 'excel_pro_365_dark_mode';

export default function App() {
  // --- Persistent State Initialization ---
  const [sheets, setSheets] = useState<SheetData[]>(() => {
    try {
      const saved = localStorage.getItem(STORAGE_KEY_SHEETS);
      if (saved) return JSON.parse(saved);
    } catch {}
    return INITIAL_SHEETS;
  });

  const [activeSheetId, setActiveSheetId] = useState<string>('akademia-sheet');

  const [userProfile, setUserProfile] = useState<UserProfile>(() => {
    try {
      const saved = localStorage.getItem(STORAGE_KEY_PROFILE);
      if (saved) return JSON.parse(saved);
    } catch {}
    return INITIAL_USER_PROFILE;
  });

  const [badges, setBadges] = useState<UserBadge[]>(() => {
    try {
      const saved = localStorage.getItem(STORAGE_KEY_BADGES);
      if (saved) return JSON.parse(saved);
    } catch {}
    return INITIAL_BADGES;
  });

  const [isDarkMode, setIsDarkMode] = useState<boolean>(() => {
    try {
      return localStorage.getItem(STORAGE_KEY_DARK) === 'true';
    } catch {
      return false;
    }
  });

  // --- UI / Workspace State ---
  const [activeTab, setActiveTab] = useState<RibbonTab>('home');
  const [activeCellId, setActiveCellId] = useState<string>('C8');
  const [selectionRange, setSelectionRange] = useState<SelectionRange | null>(null);
  const [formulaBarValue, setFormulaBarValue] = useState<string>('');
  const [isLearningPanelOpen, setIsLearningPanelOpen] = useState<boolean>(true);
  const [currentTaskId, setCurrentTaskId] = useState<string>('lesson-1');
  const [showGridlines, setShowGridlines] = useState<boolean>(true);
  const [showHeaders, setShowHeaders] = useState<boolean>(true);
  const [zoomLevel, setZoomLevel] = useState<number>(100);
  const [isOnline, setIsOnline] = useState<boolean>(navigator.onLine);
  const [lastSavedText, setLastSavedText] = useState<string>('Zsynchronizowano');
  const [hasClaimedDailyBonus, setHasClaimedDailyBonus] = useState<boolean>(false);

  // Undo / Redo stacks
  const [undoStack, setUndoStack] = useState<SheetData[][]>([]);
  const [redoStack, setRedoStack] = useState<SheetData[][]>([]);

  // Modals state
  const [isLeaderboardOpen, setIsLeaderboardOpen] = useState<boolean>(false);
  const [isProfileOpen, setIsProfileOpen] = useState<boolean>(false);
  const [isDailyModalOpen, setIsDailyModalOpen] = useState<boolean>(false);
  const [isChartModalOpen, setIsChartModalOpen] = useState<boolean>(false);
  const [chartModalInitialType, setChartModalInitialType] = useState<'bar' | 'line' | 'pie'>('bar');
  const [isSpeedChallengeOpen, setIsSpeedChallengeOpen] = useState<boolean>(false);
  const [isShortcutsOpen, setIsShortcutsOpen] = useState<boolean>(false);
  const [toastMessage, setToastMessage] = useState<{ title: string; desc: string } | null>(null);

  // Active sheet reference
  const activeSheet = useMemo(() => {
    return sheets.find((s) => s.id === activeSheetId) || sheets[0];
  }, [sheets, activeSheetId]);

  // Keep dark mode class in sync on <html>
  useEffect(() => {
    if (isDarkMode) {
      document.documentElement.classList.add('dark');
    } else {
      document.documentElement.classList.remove('dark');
    }
    localStorage.setItem(STORAGE_KEY_DARK, String(isDarkMode));
  }, [isDarkMode]);

  // Online / offline event listener
  useEffect(() => {
    const handleOnline = () => setIsOnline(true);
    const handleOffline = () => setIsOnline(false);
    window.addEventListener('online', handleOnline);
    window.addEventListener('offline', handleOffline);
    return () => {
      window.removeEventListener('online', handleOnline);
      window.removeEventListener('offline', handleOffline);
    };
  }, []);

  // Update formula bar input when active cell changes
  useEffect(() => {
    const cell = activeSheet.cells[activeCellId];
    setFormulaBarValue(cell?.raw || '');
  }, [activeCellId, activeSheet]);

  // Sync state to local storage (offline database)
  const saveToStorage = useCallback((newSheets: SheetData[], newProfile: UserProfile, newBadges: UserBadge[]) => {
    try {
      localStorage.setItem(STORAGE_KEY_SHEETS, JSON.stringify(newSheets));
      localStorage.setItem(STORAGE_KEY_PROFILE, JSON.stringify(newProfile));
      localStorage.setItem(STORAGE_KEY_BADGES, JSON.stringify(newBadges));
      const now = new Date();
      setLastSavedText(`Zapisano ${now.toLocaleTimeString([], { hour: '2-digit', minute: '2-digit', second: '2-digit' })}`);
    } catch {}
  }, []);

  // Push snapshot to undo stack
  const pushUndo = useCallback(() => {
    setUndoStack((prev) => [...prev.slice(-25), sheets]);
    setRedoStack([]);
  }, [sheets]);

  const handleUndo = useCallback(() => {
    if (undoStack.length === 0) return;
    const previous = undoStack[undoStack.length - 1];
    setUndoStack((prev) => prev.slice(0, -1));
    setRedoStack((prev) => [...prev, sheets]);
    setSheets(previous);
    saveToStorage(previous, userProfile, badges);
  }, [undoStack, sheets, saveToStorage, userProfile, badges]);

  const handleRedo = useCallback(() => {
    if (redoStack.length === 0) return;
    const next = redoStack[redoStack.length - 1];
    setRedoStack((prev) => prev.slice(0, -1));
    setUndoStack((prev) => [...prev, sheets]);
    setSheets(next);
    saveToStorage(next, userProfile, badges);
  }, [redoStack, sheets, saveToStorage, userProfile, badges]);

  // Trigger celebration banner / toast
  const triggerToast = (title: string, desc: string) => {
    setToastMessage({ title, desc });
    setTimeout(() => setToastMessage(null), 4500);
  };

  // Check and unlock badges
  const checkBadgeProgress = useCallback((
    profileUpdates: Partial<UserProfile>,
    sheetToCheck: SheetData
  ) => {
    setBadges((prevBadges) => {
      let hasChanges = false;
      const updated = prevBadges.map((badge) => {
        let newCurrent = badge.current;

        if (badge.id === 'badge-1') { // Pierwszy Krok
          newCurrent = (profileUpdates.completedTaskIds || userProfile.completedTaskIds).length;
        } else if (badge.id === 'badge-2') { // Mistrz Autosumowania
          // Check if SUMA or AVERAGE is present
          let hasSumOrAvg = 0;
          Object.values(sheetToCheck.cells).forEach((c) => {
            if (c.raw.toUpperCase().includes('SUMA') || c.raw.toUpperCase().includes('SUM')) hasSumOrAvg = Math.max(hasSumOrAvg, 1);
            if (c.raw.toUpperCase().includes('ŚREDNIA') || c.raw.toUpperCase().includes('AVERAGE')) hasSumOrAvg = 2;
          });
          newCurrent = Math.max(newCurrent, hasSumOrAvg);
        } else if (badge.id === 'badge-3') { // Logiczny Geniusz
          const hasIf = Object.values(sheetToCheck.cells).some((c) => c.raw.toUpperCase().includes('JEŻELI') || c.raw.toUpperCase().includes('IF'));
          if (hasIf) newCurrent = 1;
        } else if (badge.id === 'badge-4') { // Detektyw VLOOKUP
          const hasVlookup = Object.values(sheetToCheck.cells).some((c) => c.raw.toUpperCase().includes('WYSZUKAJ.PIONOWO') || c.raw.toUpperCase().includes('VLOOKUP'));
          if (hasVlookup) newCurrent = 1;
        } else if (badge.id === 'badge-5') { // Płomień passy
          newCurrent = userProfile.streakDays;
        } else if (badge.id === 'badge-6') { // Skróty klawiszowe
          newCurrent = (profileUpdates.shortcutsUsedCount ?? userProfile.shortcutsUsedCount);
        }

        const newProgress = Math.min(100, Math.round((newCurrent / badge.target) * 100));
        if (newProgress !== badge.progress || newCurrent !== badge.current) {
          hasChanges = true;
          if (newProgress >= 100 && badge.progress < 100) {
            sounds.playBadgeUnlocked();
            confetti({ particleCount: 70, spread: 60, origin: { y: 0.7 } });
            triggerToast(`🎖️ Zdobyto odznakę: ${badge.name}!`, badge.description);
          }
          return { ...badge, current: newCurrent, progress: newProgress, unlockedAt: newProgress >= 100 ? new Date().toISOString() : badge.unlockedAt };
        }
        return badge;
      });

      if (hasChanges) {
        localStorage.setItem(STORAGE_KEY_BADGES, JSON.stringify(updated));
      }
      return hasChanges ? updated : prevBadges;
    });
  }, [userProfile]);

  // Recalculate dependent cells across the sheet
  const recomputeSheet = useCallback((targetSheet: SheetData): SheetData => {
    const updatedCells = { ...targetSheet.cells };

    // Recompute all cells with formulas
    for (const [cellId, cell] of Object.entries(updatedCells)) {
      if (cell.raw && cell.raw.startsWith('=')) {
        const evalResult = evaluateFormula(cell.raw, { ...targetSheet, cells: updatedCells });
        updatedCells[cellId] = {
          ...cell,
          computed: evalResult.result,
          error: evalResult.error,
        };
      } else if (cell.raw && !isNaN(Number(cell.raw.replace(',', '.')))) {
        updatedCells[cellId] = {
          ...cell,
          computed: Number(cell.raw.replace(',', '.')),
        };
      } else {
        updatedCells[cellId] = {
          ...cell,
          computed: cell.raw,
        };
      }
    }

    return { ...targetSheet, cells: updatedCells };
  }, []);

  // Real-time task verification check
  const verifyCurrentTask = useCallback((updatedSheet: SheetData) => {
    const task = EXCEL_LESSONS.find((t) => t.id === currentTaskId);
    if (!task) return;

    if (updatedSheet.id !== task.targetSheetId) return;

    const targetCell = updatedSheet.cells[task.targetCell];
    const val = targetCell?.computed !== undefined ? targetCell.computed : targetCell?.raw || '';
    const raw = targetCell?.raw || '';

    try {
      const isCorrect = task.expectedResultCheck(val, raw, updatedSheet);
      if (isCorrect && !userProfile.completedTaskIds.includes(task.id)) {
        // SUCCESS CELEBRATION!
        sounds.playSuccess();
        confetti({
          particleCount: 100,
          spread: 70,
          origin: { y: 0.6 },
          colors: ['#107c41', '#f59e0b', '#3b82f6', '#10b981'],
        });

        const newCompleted = [...userProfile.completedTaskIds, task.id];
        const newXp = userProfile.xp + task.xpReward;
        const newLevel = Math.floor(newXp / 500) + 1;
        const newDailyDone = userProfile.dailyGoalTasksDone + 1;

        const updatedProfile: UserProfile = {
          ...userProfile,
          completedTaskIds: newCompleted,
          xp: newXp,
          level: newLevel,
          dailyGoalTasksDone: newDailyDone,
          dailyGoalCompleted: newDailyDone >= userProfile.dailyGoalTasksTarget,
        };

        setUserProfile(updatedProfile);
        triggerToast(
          `🎉 Zadanie "${task.title}" zaliczone!`,
          `Zdobyto +${task.xpReward} XP! Postęp ligowy wzrósł.`
        );

        checkBadgeProgress({ completedTaskIds: newCompleted, xp: newXp }, updatedSheet);
      }
    } catch {}
  }, [currentTaskId, userProfile, checkBadgeProgress]);

  // Update cell value
  const handleCellChange = useCallback((cellId: string, rawValue: string) => {
    pushUndo();

    setSheets((prevSheets) => {
      const updated = prevSheets.map((s) => {
        if (s.id !== activeSheetId) return s;

        const prevCell = s.cells[cellId] || { raw: '' };
        const updatedCells = {
          ...s.cells,
          [cellId]: {
            ...prevCell,
            raw: rawValue,
          },
        };

        const recomputed = recomputeSheet({ ...s, cells: updatedCells });
        verifyCurrentTask(recomputed);
        checkBadgeProgress({}, recomputed);
        return recomputed;
      });

      saveToStorage(updated, userProfile, badges);
      return updated;
    });

    setFormulaBarValue(rawValue);
  }, [activeSheetId, pushUndo, recomputeSheet, verifyCurrentTask, checkBadgeProgress, saveToStorage, userProfile, badges]);

  // Apply style to active cell or selection range
  const handleApplyStyle = useCallback((stylePatch: Partial<CellStyle>) => {
    pushUndo();

    const cellsToStyle: string[] = [];
    if (!selectionRange) {
      cellsToStyle.push(activeCellId);
    } else {
      const c1 = letterToColIndex(selectionRange.startCol);
      const c2 = letterToColIndex(selectionRange.endCol);
      const minC = Math.min(c1, c2);
      const maxC = Math.max(c1, c2);

      const minR = Math.min(selectionRange.startRow, selectionRange.endRow);
      const maxR = Math.max(selectionRange.startRow, selectionRange.endRow);

      for (let r = minR; r <= maxR; r++) {
        for (let c = minC; c <= maxC; c++) {
          cellsToStyle.push(`${String.fromCharCode(c + 65)}${r}`);
        }
      }
    }

    setSheets((prevSheets) => {
      const updated = prevSheets.map((s) => {
        if (s.id !== activeSheetId) return s;
        const newCells = { ...s.cells };

        cellsToStyle.forEach((id) => {
          const current = newCells[id] || { raw: '' };
          newCells[id] = {
            ...current,
            style: {
              ...(current.style || {}),
              ...stylePatch,
            },
          };
        });

        return { ...s, cells: newCells };
      });

      saveToStorage(updated, userProfile, badges);
      return updated;
    });
  }, [pushUndo, selectionRange, activeCellId, activeSheetId, saveToStorage, userProfile, badges]);

  // Intelligent AutoFill (Dragging Fill Handle)
  const handleAutoFill = useCallback((sourceRange: SelectionRange, targetRange: SelectionRange) => {
    pushUndo();

    const srcCells = getRangeCells(`${sourceRange.startCol}${sourceRange.startRow}:${sourceRange.endCol}${sourceRange.endRow}`);
    const tgtCells = getRangeCells(`${targetRange.startCol}${targetRange.startRow}:${targetRange.endCol}${targetRange.endRow}`);

    setSheets((prevSheets) => {
      const updated = prevSheets.map((s) => {
        if (s.id !== activeSheetId) return s;
        const newCells = { ...s.cells };

        // Determine direction and fill logic
        tgtCells.forEach((tgtCellId, idx) => {
          // Source pattern cell
          const srcCellId = srcCells[idx % srcCells.length];
          const srcCell = s.cells[srcCellId];

          if (!srcCell) return;

          const srcCoord = parseCellAddress(srcCellId);
          const tgtCoord = parseCellAddress(tgtCellId);
          if (!srcCoord || !tgtCoord) return;

          const rowOffset = tgtCoord.row - srcCoord.row;

          let newRaw = srcCell.raw;

          // If source is a formula (e.g. =C3*0.23 or =B3-C3), shift cell references down!
          if (newRaw.startsWith('=')) {
            newRaw = newRaw.replace(/([A-Z]+)(\d+)/g, (match, col, row) => {
              const shiftedRow = parseInt(row, 10) + rowOffset;
              return `${col}${shiftedRow}`;
            });
          } else if (!isNaN(Number(srcCell.raw)) && srcCell.raw.trim() !== '') {
            // Number sequence increment
            const baseNum = Number(srcCell.raw);
            const step = Math.floor(idx / srcCells.length) + 1;
            newRaw = String(baseNum + step);
          }

          newCells[tgtCellId] = {
            ...srcCell,
            raw: newRaw,
            style: { ...(srcCell.style || {}) },
          };
        });

        const recomputed = recomputeSheet({ ...s, cells: newCells });
        verifyCurrentTask(recomputed);
        return recomputed;
      });

      saveToStorage(updated, userProfile, badges);
      return updated;
    });
  }, [pushUndo, activeSheetId, recomputeSheet, verifyCurrentTask, saveToStorage, userProfile, badges]);

  // AutoSum Σ button handler
  const handleAutoSum = useCallback((formulaType: 'SUMA' | 'ŚREDNIA' | 'MAX' | 'MIN' | 'ILE.LICZB') => {
    const activeCoord = parseCellAddress(activeCellId);
    if (!activeCoord) return;

    // Detect range of numbers above active cell
    let startRow = activeCoord.row - 1;
    while (startRow > 0) {
      const aboveCell = activeSheet.cells[`${activeCoord.col}${startRow}`];
      if (!aboveCell || aboveCell.raw === '') {
        break;
      }
      startRow--;
    }
    const detectedStartRow = Math.max(1, startRow + 1);
    const detectedEndRow = Math.max(detectedStartRow, activeCoord.row - 1);

    const generatedFormula = `=${formulaType}(${activeCoord.col}${detectedStartRow}:${activeCoord.col}${detectedEndRow})`;
    handleCellChange(activeCellId, generatedFormula);
  }, [activeCellId, activeSheet.cells, handleCellChange]);

  // Sort rows in active sheet
  const handleSort = useCallback((direction: 'asc' | 'desc') => {
    pushUndo();
    const activeCoord = parseCellAddress(activeCellId);
    if (!activeCoord) return;

    const colLetter = activeCoord.col;

    // Sort rows 3 to 7 or all detected rows
    const rowsToSort = [3, 4, 5, 6, 7];

    setSheets((prevSheets) => {
      const updated = prevSheets.map((s) => {
        if (s.id !== activeSheetId) return s;

        const rowData = rowsToSort.map((r) => {
          const rowCells: Record<string, any> = {};
          for (let c = 0; c < 15; c++) {
            const letter = String.fromCharCode(c + 65);
            rowCells[letter] = s.cells[`${letter}${r}`];
          }
          const sortVal = s.cells[`${colLetter}${r}`]?.computed ?? s.cells[`${colLetter}${r}`]?.raw ?? '';
          return { rowNumber: r, cells: rowCells, sortVal };
        });

        rowData.sort((a, b) => {
          const numA = Number(a.sortVal);
          const numB = Number(b.sortVal);
          if (!isNaN(numA) && !isNaN(numB)) {
            return direction === 'asc' ? numA - numB : numB - numA;
          }
          return direction === 'asc'
            ? String(a.sortVal).localeCompare(String(b.sortVal))
            : String(b.sortVal).localeCompare(String(a.sortVal));
        });

        const newCells = { ...s.cells };
        rowsToSort.forEach((targetRow, idx) => {
          const sourceRowCells = rowData[idx].cells;
          for (let c = 0; c < 15; c++) {
            const letter = String.fromCharCode(c + 65);
            newCells[`${letter}${targetRow}`] = sourceRowCells[letter];
          }
        });

        return recomputeSheet({ ...s, cells: newCells });
      });

      saveToStorage(updated, userProfile, badges);
      return updated;
    });
  }, [pushUndo, activeCellId, activeSheetId, recomputeSheet, saveToStorage, userProfile, badges]);

  // Insert & Delete row
  const handleInsertRow = useCallback(() => {
    pushUndo();
    const activeCoord = parseCellAddress(activeCellId);
    if (!activeCoord) return;

    setSheets((prevSheets) => {
      const updated = prevSheets.map((s) => {
        if (s.id !== activeSheetId) return s;
        const newCells: Record<string, any> = {};

        Object.entries(s.cells).forEach(([key, cell]) => {
          const coord = parseCellAddress(key);
          if (!coord) return;
          if (coord.row >= activeCoord.row) {
            newCells[`${coord.col}${coord.row + 1}`] = cell;
          } else {
            newCells[key] = cell;
          }
        });

        return { ...s, cells: newCells };
      });
      return updated;
    });
  }, [pushUndo, activeCellId, activeSheetId]);

  const handleDeleteRow = useCallback(() => {
    pushUndo();
    const activeCoord = parseCellAddress(activeCellId);
    if (!activeCoord) return;

    setSheets((prevSheets) => {
      const updated = prevSheets.map((s) => {
        if (s.id !== activeSheetId) return s;
        const newCells: Record<string, any> = {};

        Object.entries(s.cells).forEach(([key, cell]) => {
          const coord = parseCellAddress(key);
          if (!coord) return;
          if (coord.row === activeCoord.row) {
            // Delete this row
          } else if (coord.row > activeCoord.row) {
            newCells[`${coord.col}${coord.row - 1}`] = cell;
          } else {
            newCells[key] = cell;
          }
        });

        return { ...s, cells: newCells };
      });
      return updated;
    });
  }, [pushUndo, activeCellId, activeSheetId]);

  // Export CSV
  const handleExportCSV = useCallback(() => {
    let csvContent = '\uFEFF'; // UTF-8 BOM for Excel Polish accents
    for (let r = 1; r <= 35; r++) {
      const rowVals: string[] = [];
      for (let c = 0; c < 12; c++) {
        const letter = String.fromCharCode(c + 65);
        const cell = activeSheet.cells[`${letter}${r}`];
        const val = cell?.computed !== undefined ? String(cell.computed) : cell?.raw || '';
        // Escape quotes
        rowVals.push(`"${val.replace(/"/g, '""')}"`);
      }
      csvContent += rowVals.join(';') + '\r\n';
    }

    const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = `${activeSheet.name.replace(/\s+/g, '_')}_365.csv`;
    link.click();
    URL.revokeObjectURL(url);
    triggerToast('Eksport CSV', `Pomyślnie wyeksportowano arkusz "${activeSheet.name}" do pliku CSV.`);
  }, [activeSheet]);

  // Import CSV
  const handleImportCSV = useCallback(() => {
    const input = document.createElement('input');
    input.type = 'file';
    input.accept = '.csv,text/csv';
    input.onchange = (e: any) => {
      const file = e.target.files?.[0];
      if (!file) return;
      const reader = new FileReader();
      reader.onload = (event) => {
        const text = event.target?.result as string;
        if (!text) return;

        pushUndo();
        const lines = text.split(/\r\n|\n/);
        const newCells: Record<string, any> = {};

        lines.forEach((line, rIdx) => {
          if (!line.trim()) return;
          const cols = line.split(/[;,]/);
          cols.forEach((colVal, cIdx) => {
            const letter = colIndexToLetter(cIdx);
            const cleanVal = colVal.replace(/^"|"$/g, '').trim();
            if (cleanVal) {
              newCells[`${letter}${rIdx + 1}`] = { raw: cleanVal };
            }
          });
        });

        setSheets((prevSheets) => {
          const updated = prevSheets.map((s) => {
            if (s.id !== activeSheetId) return s;
            return recomputeSheet({ ...s, cells: newCells });
          });
          saveToStorage(updated, userProfile, badges);
          return updated;
        });

        triggerToast('Import CSV', `Wczytano pomyślnie dane z pliku ${file.name}.`);
      };
      reader.readAsText(file);
    };
    input.click();
  }, [pushUndo, activeSheetId, recomputeSheet, saveToStorage, userProfile, badges]);

  // Keyboard shortcut tracking
  const handleShortcutUsed = useCallback((name: string) => {
    if (name === 'Ctrl+B') handleApplyStyle({ bold: !(activeSheet.cells[activeCellId]?.style?.bold) });
    if (name === 'Ctrl+I') handleApplyStyle({ italic: !(activeSheet.cells[activeCellId]?.style?.italic) });
    if (name === 'Ctrl+U') handleApplyStyle({ underline: !(activeSheet.cells[activeCellId]?.style?.underline) });

    const newCount = userProfile.shortcutsUsedCount + 1;
    setUserProfile((prev) => ({ ...prev, shortcutsUsedCount: newCount }));
    checkBadgeProgress({ shortcutsUsedCount: newCount }, activeSheet);
  }, [handleApplyStyle, activeSheet, activeCellId, userProfile.shortcutsUsedCount, checkBadgeProgress]);

  // Sheet Tabs Actions
  const handleAddSheet = useCallback(() => {
    const newId = `sheet-${Date.now()}`;
    const newName = `Arkusz${sheets.length + 1}`;
    const newSheet: SheetData = {
      id: newId,
      name: newName,
      colWidths: {},
      rowHeights: {},
      cells: {},
    };
    setSheets((prev) => [...prev, newSheet]);
    setActiveSheetId(newId);
  }, [sheets.length]);

  const handleDeleteSheet = useCallback((sheetId: string) => {
    if (sheets.length <= 1) return;
    setSheets((prev) => {
      const filtered = prev.filter((s) => s.id !== sheetId);
      if (activeSheetId === sheetId) {
        setActiveSheetId(filtered[0].id);
      }
      return filtered;
    });
  }, [sheets.length, activeSheetId]);

  const handleRenameSheet = useCallback((sheetId: string, newName: string) => {
    setSheets((prev) =>
      prev.map((s) => (s.id === sheetId ? { ...s, name: newName } : s))
    );
  }, []);

  // Daily Bonus XP claim
  const handleClaimDailyBonus = () => {
    if (hasClaimedDailyBonus) return;
    setHasClaimedDailyBonus(true);
    const newXp = userProfile.xp + 50;
    const newLevel = Math.floor(newXp / 500) + 1;
    setUserProfile((prev) => ({ ...prev, xp: newXp, level: newLevel }));
    sounds.playSuccess();
    confetti({ particleCount: 80, spread: 60, origin: { y: 0.6 } });
    triggerToast('🎉 Odebrano bonus!', '+50 XP dodane do Twojego profilu za dzisiejszą obecność!');
  };

  // Speed Challenge Complete
  const handleCompleteSpeedChallenge = (earnedXp: number) => {
    const newXp = userProfile.xp + earnedXp;
    const newLevel = Math.floor(newXp / 500) + 1;
    setUserProfile((prev) => ({ ...prev, xp: newXp, level: newLevel }));
    triggerToast('⏱️ Pojedynek zakończony!', `Zarobiono +${earnedXp} XP w ligowej rywalizacji!`);
  };

  // Current lesson target cell
  const currentLesson = EXCEL_LESSONS.find((l) => l.id === currentTaskId);
  const targetLessonCell = activeSheet.id === currentLesson?.targetSheetId ? currentLesson?.targetCell : undefined;

  return (
    <div className={`flex flex-col h-screen w-screen overflow-hidden select-none transition-colors ${
      isDarkMode ? 'bg-[#181818] text-white' : 'bg-[#f3f4f6] text-neutral-900'
    }`}>
      {/* Toast Notification */}
      {toastMessage && (
        <div className="fixed top-14 right-4 z-50 bg-[#107c41] text-white px-4 py-3 rounded-xl shadow-2xl flex items-start gap-3 border border-emerald-400/40 animate-in slide-in-from-top-4 duration-200 max-w-sm">
          <div className="text-xl">✨</div>
          <div>
            <div className="font-bold text-xs">{toastMessage.title}</div>
            <div className="text-[11px] text-emerald-100">{toastMessage.desc}</div>
          </div>
        </div>
      )}

      {/* Office 365 Header */}
      <ExcelHeader
        userProfile={userProfile}
        isDarkMode={isDarkMode}
        onToggleDarkMode={() => setIsDarkMode(!isDarkMode)}
        isOnline={isOnline}
        lastSavedText={lastSavedText}
        onSaveManual={() => saveToStorage(sheets, userProfile, badges)}
        canUndo={undoStack.length > 0}
        canRedo={redoStack.length > 0}
        onUndo={handleUndo}
        onRedo={handleRedo}
        onOpenProfile={() => setIsProfileOpen(true)}
        onOpenLeaderboard={() => setIsLeaderboardOpen(true)}
        onOpenDailyModal={() => setIsDailyModalOpen(true)}
        onOpenShortcuts={() => setIsShortcutsOpen(true)}
        onOpenChartModal={() => setIsChartModalOpen(true)}
        onToggleLearningPanel={() => setIsLearningPanelOpen(!isLearningPanelOpen)}
        isLearningPanelOpen={isLearningPanelOpen}
      />

      {/* Office 365 Ribbon */}
      <ExcelRibbon
        activeTab={activeTab}
        onTabChange={setActiveTab}
        currentStyle={activeSheet.cells[activeCellId]?.style}
        onApplyStyle={handleApplyStyle}
        onAutoSum={handleAutoSum}
        onSort={handleSort}
        onInsertRow={handleInsertRow}
        onDeleteRow={handleDeleteRow}
        onClearCell={() => handleCellChange(activeCellId, '')}
        onInsertFunctionTemplate={(tmpl) => handleCellChange(activeCellId, tmpl)}
        onExportCSV={handleExportCSV}
        onImportCSV={handleImportCSV}
        onOpenChartModal={(type) => {
          if (type) setChartModalInitialType(type);
          setIsChartModalOpen(true);
        }}
        onOpenLeaderboard={() => setIsLeaderboardOpen(true)}
        onOpenSpeedChallenge={() => setIsSpeedChallengeOpen(true)}
        onToggleLearningPanel={() => setIsLearningPanelOpen(!isLearningPanelOpen)}
        onShowLessonHint={() => {
          setIsLearningPanelOpen(true);
        }}
        onResetLessonData={() => {
          const fresh = INITIAL_SHEETS.find((s) => s.id === 'akademia-sheet');
          if (fresh) {
            setSheets((prev) => prev.map((s) => (s.id === 'akademia-sheet' ? fresh : s)));
            triggerToast('Reset Lekcji', 'Przywrócono początkowe dane w arkuszu Akademii.');
          }
        }}
        showGridlines={showGridlines}
        onToggleGridlines={() => setShowGridlines(!showGridlines)}
        showHeaders={showHeaders}
        onToggleHeaders={() => setShowHeaders(!showHeaders)}
        zoomLevel={zoomLevel}
        onSetZoom={setZoomLevel}
        isDarkMode={isDarkMode}
      />

      {/* Formula Bar */}
      <FormulaBar
        activeCellId={activeCellId}
        selectionRangeText={
          selectionRange
            ? `${selectionRange.startCol}${selectionRange.startRow}:${selectionRange.endCol}${selectionRange.endRow}`
            : ''
        }
        currentValue={formulaBarValue}
        onValueChange={(val) => {
          setFormulaBarValue(val);
          handleCellChange(activeCellId, val);
        }}
        onSubmit={() => {
          handleCellChange(activeCellId, formulaBarValue);
        }}
        onCancel={() => {
          setFormulaBarValue(activeSheet.cells[activeCellId]?.raw || '');
        }}
        onJumpToCell={(cellId) => setActiveCellId(cellId)}
        onOpenFunctionsMenu={() => setActiveTab('formulas')}
        isDarkMode={isDarkMode}
      />

      {/* Main Workspace Area (Grid + Learning Panel) */}
      <div className="flex-1 flex overflow-hidden relative">
        <SpreadsheetGrid
          sheet={activeSheet}
          activeCellId={activeCellId}
          selectionRange={selectionRange}
          onSelectCell={setActiveCellId}
          onSelectRange={setSelectionRange}
          onCellChange={handleCellChange}
          onAutoFill={handleAutoFill}
          targetLessonCell={targetLessonCell}
          showGridlines={showGridlines}
          showHeaders={showHeaders}
          zoomLevel={zoomLevel}
          isDarkMode={isDarkMode}
          onShortcutUsed={handleShortcutUsed}
          onApplyStyle={handleApplyStyle}
          onInsertRow={handleInsertRow}
          onDeleteRow={handleDeleteRow}
          onAutoSum={handleAutoSum}
          onOpenChartModal={() => setIsChartModalOpen(true)}
        />

        {/* Learning Academy Panel */}
        {isLearningPanelOpen && (
          <LearningPanel
            tasks={EXCEL_LESSONS}
            currentTaskId={currentTaskId}
            onSelectTask={(taskId) => {
              setCurrentTaskId(taskId);
              const task = EXCEL_LESSONS.find((t) => t.id === taskId);
              if (task) {
                if (activeSheetId !== task.targetSheetId) {
                  setActiveSheetId(task.targetSheetId);
                }
                setActiveCellId(task.targetCell);
              }
            }}
            completedTaskIds={userProfile.completedTaskIds}
            activeSheet={activeSheet}
            onNavigateToTaskSheet={(sheetId, cellId) => {
              setActiveSheetId(sheetId);
              setActiveCellId(cellId);
            }}
            onResetTaskSheet={() => {
              const fresh = INITIAL_SHEETS.find((s) => s.id === 'akademia-sheet');
              if (fresh) {
                setSheets((prev) => prev.map((s) => (s.id === 'akademia-sheet' ? fresh : s)));
              }
            }}
            isDarkMode={isDarkMode}
            onClose={() => setIsLearningPanelOpen(false)}
          />
        )}
      </div>

      {/* Bottom Status Bar */}
      <StatusBar
        sheets={sheets}
        activeSheetId={activeSheetId}
        onSelectSheet={setActiveSheetId}
        onAddSheet={handleAddSheet}
        onDeleteSheet={handleDeleteSheet}
        onRenameSheet={handleRenameSheet}
        activeSheet={activeSheet}
        selectionRange={selectionRange}
        activeCellId={activeCellId}
        zoomLevel={zoomLevel}
        onSetZoom={setZoomLevel}
        isOnline={isOnline}
        isDarkMode={isDarkMode}
      />

      {/* Modals */}
      <LeaderboardModal
        isOpen={isLeaderboardOpen}
        onClose={() => setIsLeaderboardOpen(false)}
        leaderboard={INITIAL_LEADERBOARD}
        userXp={userProfile.xp}
        userStreak={userProfile.streakDays}
        onOpenSpeedChallenge={() => setIsSpeedChallengeOpen(true)}
        isDarkMode={isDarkMode}
      />

      <ProfileModal
        isOpen={isProfileOpen}
        onClose={() => setIsProfileOpen(false)}
        profile={userProfile}
        badges={badges}
        isDarkMode={isDarkMode}
      />

      <DailyMotivationModal
        isOpen={isDailyModalOpen}
        onClose={() => setIsDailyModalOpen(false)}
        streakDays={userProfile.streakDays}
        dailyGoalTasksTarget={userProfile.dailyGoalTasksTarget}
        dailyGoalTasksDone={userProfile.dailyGoalTasksDone}
        onClaimDailyBonus={handleClaimDailyBonus}
        hasClaimedDailyBonus={hasClaimedDailyBonus}
        isDarkMode={isDarkMode}
      />

      <ChartModal
        isOpen={isChartModalOpen}
        onClose={() => setIsChartModalOpen(false)}
        activeSheet={activeSheet}
        selectionRange={selectionRange}
        initialType={chartModalInitialType}
        isDarkMode={isDarkMode}
      />

      <SpeedChallengeModal
        isOpen={isSpeedChallengeOpen}
        onClose={() => setIsSpeedChallengeOpen(false)}
        onCompleteChallenge={handleCompleteSpeedChallenge}
        isDarkMode={isDarkMode}
      />

      <ShortcutsModal
        isOpen={isShortcutsOpen}
        onClose={() => setIsShortcutsOpen(false)}
        isDarkMode={isDarkMode}
      />
    </div>
  );
}
