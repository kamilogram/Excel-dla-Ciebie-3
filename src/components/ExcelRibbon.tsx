import React, { useState } from 'react';
import {
  Bold,
  Italic,
  Underline,
  AlignLeft,
  AlignCenter,
  AlignRight,
  PaintBucket,
  Type,
  DollarSign,
  Percent,
  Sigma,
  SortAsc,
  SortDesc,
  Plus,
  Trash2,
  Table,
  BarChart2,
  LineChart,
  PieChart,
  FunctionSquare,
  FileSpreadsheet,
  Download,
  Upload,
  Trophy,
  Award,
  BookOpen,
  Grid,
  ChevronDown,
  Timer,
  Lightbulb,
  RotateCcw,
} from 'lucide-react';
import { CellStyle } from '../types/excel';

export type RibbonTab = 'home' | 'insert' | 'formulas' | 'data' | 'view' | 'learn' | 'compete';

interface ExcelRibbonProps {
  activeTab: RibbonTab;
  onTabChange: (tab: RibbonTab) => void;
  currentStyle?: CellStyle;
  onApplyStyle: (stylePatch: Partial<CellStyle>) => void;
  onAutoSum: (formulaType: 'SUMA' | 'ŚREDNIA' | 'MAX' | 'MIN' | 'ILE.LICZB') => void;
  onSort: (direction: 'asc' | 'desc') => void;
  onInsertRow: () => void;
  onDeleteRow: () => void;
  onClearCell: () => void;
  onInsertFunctionTemplate: (template: string) => void;
  onExportCSV: () => void;
  onImportCSV: () => void;
  onOpenChartModal: (type?: 'bar' | 'line' | 'pie') => void;
  onOpenLeaderboard: () => void;
  onOpenSpeedChallenge: () => void;
  onToggleLearningPanel: () => void;
  onShowLessonHint: () => void;
  onResetLessonData: () => void;
  showGridlines: boolean;
  onToggleGridlines: () => void;
  showHeaders: boolean;
  onToggleHeaders: () => void;
  zoomLevel: number;
  onSetZoom: (zoom: number) => void;
  isDarkMode: boolean;
}

export const ExcelRibbon: React.FC<ExcelRibbonProps> = ({
  activeTab,
  onTabChange,
  currentStyle = {},
  onApplyStyle,
  onAutoSum,
  onSort,
  onInsertRow,
  onDeleteRow,
  onClearCell,
  onInsertFunctionTemplate,
  onExportCSV,
  onImportCSV,
  onOpenChartModal,
  onOpenLeaderboard,
  onOpenSpeedChallenge,
  onToggleLearningPanel,
  onShowLessonHint,
  onResetLessonData,
  showGridlines,
  onToggleGridlines,
  showHeaders,
  onToggleHeaders,
  zoomLevel,
  onSetZoom,
  isDarkMode,
}) => {
  const [showAutoSumMenu, setShowAutoSumMenu] = useState(false);
  const [showBorderMenu, setShowBorderMenu] = useState(false);
  const [showFillColorPicker, setShowFillColorPicker] = useState(false);
  const [showTextColorPicker, setShowTextColorPicker] = useState(false);

  const colors = [
    '#ffffff', '#000000', '#107c41', '#0f6cbd', '#d83b01', '#881798',
    '#fef3c7', '#f0fdf4', '#eff6ff', '#fef2f2', '#f3f4f6', '#d1d5db',
  ];

  return (
    <div className={`flex flex-col border-b select-none transition-colors ${
      isDarkMode ? 'bg-[#202020] border-[#333333] text-neutral-200' : 'bg-[#f3f4f6] border-[#e5e7eb] text-neutral-800'
    }`}>
      {/* Ribbon Navigation Tabs */}
      <div className="flex items-center px-2 pt-1 gap-0.5 overflow-x-auto text-xs font-medium border-b border-black/5 dark:border-white/5 scrollbar-none">
        {[
          { id: 'home', label: 'Narzędzia główne' },
          { id: 'insert', label: 'Wstawianie' },
          { id: 'formulas', label: 'Formuły' },
          { id: 'data', label: 'Dane' },
          { id: 'learn', label: 'Akademia & Zadania' },
          { id: 'compete', label: 'Rywalizacja & Ligi' },
          { id: 'view', label: 'Widok' },
        ].map((tab) => {
          const isActive = activeTab === tab.id;
          return (
            <button
              key={tab.id}
              onClick={() => onTabChange(tab.id as RibbonTab)}
              className={`px-3 py-1.5 transition-colors whitespace-nowrap rounded-t-sm relative ${
                isActive
                  ? isDarkMode
                    ? 'bg-[#2b2b2b] text-white font-semibold after:absolute after:bottom-0 after:left-0 after:right-0 after:h-0.5 after:bg-[#107c41]'
                    : 'bg-white text-[#107c41] font-semibold shadow-xs after:absolute after:bottom-0 after:left-0 after:right-0 after:h-0.5 after:bg-[#107c41]'
                  : isDarkMode
                  ? 'text-neutral-400 hover:text-white hover:bg-[#2b2b2b]/50'
                  : 'text-neutral-600 hover:text-neutral-900 hover:bg-neutral-200/60'
              }`}
            >
              {tab.label}
            </button>
          );
        })}
      </div>

      {/* Ribbon Toolbar Content Area */}
      <div className={`p-1.5 min-h-[58px] flex items-center overflow-x-auto gap-3 text-xs scrollbar-none ${
        isDarkMode ? 'bg-[#2b2b2b]' : 'bg-white'
      }`}>
        {/* --- TAB 1: NARZĘDZIA GŁÓWNE (HOME) --- */}
        {activeTab === 'home' && (
          <div className="flex items-center gap-3">
            {/* Font Style Group */}
            <div className="flex items-center gap-1 pr-3 border-r border-neutral-300 dark:border-neutral-700">
              <button
                onClick={() => onApplyStyle({ bold: !currentStyle.bold })}
                title="Pogrubienie (Ctrl+B)"
                className={`p-1.5 rounded transition-colors ${
                  currentStyle.bold ? 'bg-[#107c41] text-white' : 'hover:bg-neutral-100 dark:hover:bg-neutral-700'
                }`}
              >
                <Bold size={14} />
              </button>
              <button
                onClick={() => onApplyStyle({ italic: !currentStyle.italic })}
                title="Kursywa (Ctrl+I)"
                className={`p-1.5 rounded transition-colors ${
                  currentStyle.italic ? 'bg-[#107c41] text-white' : 'hover:bg-neutral-100 dark:hover:bg-neutral-700'
                }`}
              >
                <Italic size={14} />
              </button>
              <button
                onClick={() => onApplyStyle({ underline: !currentStyle.underline })}
                title="Podkreślenie (Ctrl+U)"
                className={`p-1.5 rounded transition-colors ${
                  currentStyle.underline ? 'bg-[#107c41] text-white' : 'hover:bg-neutral-100 dark:hover:bg-neutral-700'
                }`}
              >
                <Underline size={14} />
              </button>

              {/* Fill Color Picker */}
              <div className="relative">
                <button
                  onClick={() => setShowFillColorPicker(!showFillColorPicker)}
                  title="Kolor wypełnienia komórki"
                  className="p-1.5 rounded hover:bg-neutral-100 dark:hover:bg-neutral-700 flex items-center gap-0.5"
                >
                  <PaintBucket size={14} style={{ color: currentStyle.bgColor || '#107c41' }} />
                  <ChevronDown size={10} />
                </button>
                {showFillColorPicker && (
                  <div className="absolute top-8 left-0 z-50 p-2 bg-white dark:bg-[#333333] shadow-lg rounded border border-neutral-200 dark:border-neutral-700 grid grid-cols-6 gap-1 w-36">
                    {colors.map((c) => (
                      <button
                        key={c}
                        onClick={() => {
                          onApplyStyle({ bgColor: c });
                          setShowFillColorPicker(false);
                        }}
                        style={{ backgroundColor: c }}
                        className="w-5 h-5 rounded border border-neutral-300 dark:border-neutral-600 hover:scale-110 transition-transform"
                      />
                    ))}
                    <button
                      onClick={() => {
                        onApplyStyle({ bgColor: undefined });
                        setShowFillColorPicker(false);
                      }}
                      className="col-span-6 text-[10px] text-center py-1 mt-1 hover:bg-neutral-100 dark:hover:bg-neutral-700 rounded"
                    >
                      Brak wypełnienia
                    </button>
                  </div>
                )}
              </div>

              {/* Text Color Picker */}
              <div className="relative">
                <button
                  onClick={() => setShowTextColorPicker(!showTextColorPicker)}
                  title="Kolor czcionki"
                  className="p-1.5 rounded hover:bg-neutral-100 dark:hover:bg-neutral-700 flex items-center gap-0.5"
                >
                  <Type size={14} style={{ color: currentStyle.textColor || '#107c41' }} />
                  <ChevronDown size={10} />
                </button>
                {showTextColorPicker && (
                  <div className="absolute top-8 left-0 z-50 p-2 bg-white dark:bg-[#333333] shadow-lg rounded border border-neutral-200 dark:border-neutral-700 grid grid-cols-6 gap-1 w-36">
                    {colors.map((c) => (
                      <button
                        key={c}
                        onClick={() => {
                          onApplyStyle({ textColor: c });
                          setShowTextColorPicker(false);
                        }}
                        style={{ backgroundColor: c }}
                        className="w-5 h-5 rounded border border-neutral-300 dark:border-neutral-600 hover:scale-110 transition-transform"
                      />
                    ))}
                  </div>
                )}
              </div>

              {/* Borders dropdown */}
              <div className="relative">
                <button
                  onClick={() => setShowBorderMenu(!showBorderMenu)}
                  title="Krawędzie komórki"
                  className="p-1.5 rounded hover:bg-neutral-100 dark:hover:bg-neutral-700 flex items-center gap-0.5"
                >
                  <Table size={14} />
                  <ChevronDown size={10} />
                </button>
                {showBorderMenu && (
                  <div className="absolute top-8 left-0 z-50 p-1.5 bg-white dark:bg-[#333333] shadow-lg rounded border border-neutral-200 dark:border-neutral-700 flex flex-col gap-1 w-44 text-[11px]">
                    <button
                      onClick={() => { onApplyStyle({ border: 'all' }); setShowBorderMenu(false); }}
                      className="px-2 py-1 text-left hover:bg-neutral-100 dark:hover:bg-neutral-700 rounded"
                    >
                      Wszystkie krawędzie
                    </button>
                    <button
                      onClick={() => { onApplyStyle({ border: 'outside' }); setShowBorderMenu(false); }}
                      className="px-2 py-1 text-left hover:bg-neutral-100 dark:hover:bg-neutral-700 rounded"
                    >
                      Krawędzie zewnętrzne
                    </button>
                    <button
                      onClick={() => { onApplyStyle({ border: 'thick-bottom' }); setShowBorderMenu(false); }}
                      className="px-2 py-1 text-left hover:bg-neutral-100 dark:hover:bg-neutral-700 rounded"
                    >
                      Gruba dolna krawędź sumy
                    </button>
                    <button
                      onClick={() => { onApplyStyle({ border: 'none' }); setShowBorderMenu(false); }}
                      className="px-2 py-1 text-left hover:bg-neutral-100 dark:hover:bg-neutral-700 rounded"
                    >
                      Brak krawędzi
                    </button>
                  </div>
                )}
              </div>
            </div>

            {/* Alignment Group */}
            <div className="flex items-center gap-1 pr-3 border-r border-neutral-300 dark:border-neutral-700">
              <button
                onClick={() => onApplyStyle({ align: 'left' })}
                title="Wyrównaj do lewej"
                className={`p-1.5 rounded transition-colors ${
                  currentStyle.align === 'left' ? 'bg-[#107c41] text-white' : 'hover:bg-neutral-100 dark:hover:bg-neutral-700'
                }`}
              >
                <AlignLeft size={14} />
              </button>
              <button
                onClick={() => onApplyStyle({ align: 'center' })}
                title="Wyśrodkuj"
                className={`p-1.5 rounded transition-colors ${
                  currentStyle.align === 'center' ? 'bg-[#107c41] text-white' : 'hover:bg-neutral-100 dark:hover:bg-neutral-700'
                }`}
              >
                <AlignCenter size={14} />
              </button>
              <button
                onClick={() => onApplyStyle({ align: 'right' })}
                title="Wyrównaj do prawej"
                className={`p-1.5 rounded transition-colors ${
                  currentStyle.align === 'right' ? 'bg-[#107c41] text-white' : 'hover:bg-neutral-100 dark:hover:bg-neutral-700'
                }`}
              >
                <AlignRight size={14} />
              </button>
            </div>

            {/* Number Formats Group */}
            <div className="flex items-center gap-1 pr-3 border-r border-neutral-300 dark:border-neutral-700">
              <button
                onClick={() => onApplyStyle({ format: currentStyle.format === 'currency' ? 'general' : 'currency' })}
                title="Format walutowy (PLN)"
                className={`flex items-center gap-1 px-2 py-1 rounded transition-colors font-medium ${
                  currentStyle.format === 'currency' ? 'bg-[#107c41] text-white' : 'hover:bg-neutral-100 dark:hover:bg-neutral-700'
                }`}
              >
                <DollarSign size={13} />
                <span>Waluta zł</span>
              </button>
              <button
                onClick={() => onApplyStyle({ format: currentStyle.format === 'percent' ? 'general' : 'percent' })}
                title="Format procentowy (%)"
                className={`flex items-center gap-1 px-2 py-1 rounded transition-colors font-medium ${
                  currentStyle.format === 'percent' ? 'bg-[#107c41] text-white' : 'hover:bg-neutral-100 dark:hover:bg-neutral-700'
                }`}
              >
                <Percent size={13} />
                <span>Procent %</span>
              </button>
            </div>

            {/* Editing: AutoSum & Sort Group */}
            <div className="flex items-center gap-1">
              {/* Autosum dropdown */}
              <div className="relative">
                <div className="flex items-center rounded hover:bg-neutral-100 dark:hover:bg-neutral-700 border border-neutral-300 dark:border-neutral-600">
                  <button
                    onClick={() => onAutoSum('SUMA')}
                    title="Autosumowanie (SUMA)"
                    className="flex items-center gap-1 px-2 py-1 font-semibold text-[#107c41] dark:text-emerald-400"
                  >
                    <Sigma size={14} />
                    <span>Autosumowanie</span>
                  </button>
                  <button
                    onClick={() => setShowAutoSumMenu(!showAutoSumMenu)}
                    className="px-1 py-1 border-l border-neutral-300 dark:border-neutral-600"
                  >
                    <ChevronDown size={11} />
                  </button>
                </div>
                {showAutoSumMenu && (
                  <div className="absolute top-8 left-0 z-50 p-1 bg-white dark:bg-[#333333] shadow-lg rounded border border-neutral-200 dark:border-neutral-700 flex flex-col gap-0.5 w-36 text-[11px]">
                    <button
                      onClick={() => { onAutoSum('SUMA'); setShowAutoSumMenu(false); }}
                      className="px-2 py-1 text-left hover:bg-neutral-100 dark:hover:bg-neutral-700 rounded"
                    >
                      SUMA
                    </button>
                    <button
                      onClick={() => { onAutoSum('ŚREDNIA'); setShowAutoSumMenu(false); }}
                      className="px-2 py-1 text-left hover:bg-neutral-100 dark:hover:bg-neutral-700 rounded"
                    >
                      ŚREDNIA
                    </button>
                    <button
                      onClick={() => { onAutoSum('MAX'); setShowAutoSumMenu(false); }}
                      className="px-2 py-1 text-left hover:bg-neutral-100 dark:hover:bg-neutral-700 rounded"
                    >
                      MAKS
                    </button>
                    <button
                      onClick={() => { onAutoSum('MIN'); setShowAutoSumMenu(false); }}
                      className="px-2 py-1 text-left hover:bg-neutral-100 dark:hover:bg-neutral-700 rounded"
                    >
                      MIN
                    </button>
                    <button
                      onClick={() => { onAutoSum('ILE.LICZB'); setShowAutoSumMenu(false); }}
                      className="px-2 py-1 text-left hover:bg-neutral-100 dark:hover:bg-neutral-700 rounded"
                    >
                      ILE.LICZB
                    </button>
                  </div>
                )}
              </div>

              {/* Sort buttons */}
              <button
                onClick={() => onSort('asc')}
                title="Sortuj od A do Z / rosnąco"
                className="p-1.5 rounded hover:bg-neutral-100 dark:hover:bg-neutral-700"
              >
                <SortAsc size={15} />
              </button>
              <button
                onClick={() => onSort('desc')}
                title="Sortuj od Z do A / malejąco"
                className="p-1.5 rounded hover:bg-neutral-100 dark:hover:bg-neutral-700"
              >
                <SortDesc size={15} />
              </button>

              {/* Row add / delete / clear */}
              <button
                onClick={onInsertRow}
                title="Wstaw wiersz poniżej"
                className="flex items-center gap-1 px-2 py-1 rounded hover:bg-neutral-100 dark:hover:bg-neutral-700"
              >
                <Plus size={13} />
                <span className="hidden sm:inline">Wiersz</span>
              </button>
              <button
                onClick={onDeleteRow}
                title="Usuń aktywny wiersz"
                className="flex items-center gap-1 px-2 py-1 rounded hover:bg-neutral-100 dark:hover:bg-neutral-700 text-rose-600 dark:text-rose-400"
              >
                <Trash2 size={13} />
              </button>
            </div>
          </div>
        )}

        {/* --- TAB 2: WSTAWIANIE (INSERT) --- */}
        {activeTab === 'insert' && (
          <div className="flex items-center gap-4">
            <div className="flex items-center gap-2 pr-4 border-r border-neutral-300 dark:border-neutral-700">
              <span className="text-[11px] font-semibold text-neutral-500 uppercase tracking-wide">Wykresy:</span>
              <button
                onClick={() => onOpenChartModal('bar')}
                className="flex items-center gap-1.5 px-2.5 py-1.5 rounded hover:bg-neutral-100 dark:hover:bg-neutral-700 border border-neutral-300 dark:border-neutral-600"
              >
                <BarChart2 size={16} className="text-[#107c41]" />
                <span>Kolumnowy</span>
              </button>
              <button
                onClick={() => onOpenChartModal('line')}
                className="flex items-center gap-1.5 px-2.5 py-1.5 rounded hover:bg-neutral-100 dark:hover:bg-neutral-700 border border-neutral-300 dark:border-neutral-600"
              >
                <LineChart size={16} className="text-[#0f6cbd]" />
                <span>Liniowy</span>
              </button>
              <button
                onClick={() => onOpenChartModal('pie')}
                className="flex items-center gap-1.5 px-2.5 py-1.5 rounded hover:bg-neutral-100 dark:hover:bg-neutral-700 border border-neutral-300 dark:border-neutral-600"
              >
                <PieChart size={16} className="text-amber-500" />
                <span>Kołowy</span>
              </button>
            </div>

            <div className="flex items-center gap-2">
              <button
                onClick={onInsertRow}
                className="flex items-center gap-1.5 px-2.5 py-1.5 rounded hover:bg-neutral-100 dark:hover:bg-neutral-700 border border-neutral-300 dark:border-neutral-600"
              >
                <Plus size={14} />
                <span>Wstaw wiersz danych</span>
              </button>
            </div>
          </div>
        )}

        {/* --- TAB 3: FORMUŁY (FORMULAS) --- */}
        {activeTab === 'formulas' && (
          <div className="flex items-center gap-2 overflow-x-auto">
            <div className="flex items-center gap-1 pr-3 border-r border-neutral-300 dark:border-neutral-700">
              <span className="text-[11px] font-semibold text-neutral-500 uppercase tracking-wide">Szybkie formuły:</span>
              <button
                onClick={() => onInsertFunctionTemplate('=SUMA()')}
                className="px-2 py-1 rounded bg-[#107c41]/10 text-[#107c41] dark:text-emerald-400 font-mono font-bold hover:bg-[#107c41]/20"
              >
                SUMA
              </button>
              <button
                onClick={() => onInsertFunctionTemplate('=ŚREDNIA()')}
                className="px-2 py-1 rounded bg-[#107c41]/10 text-[#107c41] dark:text-emerald-400 font-mono font-bold hover:bg-[#107c41]/20"
              >
                ŚREDNIA
              </button>
              <button
                onClick={() => onInsertFunctionTemplate('=JEŻELI(warunek; prawda; fałsz)')}
                className="px-2 py-1 rounded bg-amber-500/10 text-amber-700 dark:text-amber-400 font-mono font-bold hover:bg-amber-500/20"
              >
                JEŻELI
              </button>
              <button
                onClick={() => onInsertFunctionTemplate('=WYSZUKAJ.PIONOWO(szukaj; tabela; kolumna; 0)')}
                className="px-2 py-1 rounded bg-blue-500/10 text-blue-700 dark:text-blue-400 font-mono font-bold hover:bg-blue-500/20"
              >
                WYSZUKAJ.PIONOWO
              </button>
              <button
                onClick={() => onInsertFunctionTemplate('=LICZ.JEŻELI(zakres; kryterium)')}
                className="px-2 py-1 rounded bg-purple-500/10 text-purple-700 dark:text-purple-400 font-mono font-bold hover:bg-purple-500/20"
              >
                LICZ.JEŻELI
              </button>
              <button
                onClick={() => onInsertFunctionTemplate('=SUMA.JEŻELI(zakres_kryterium; kryterium; zakres_sumy)')}
                className="px-2 py-1 rounded bg-emerald-500/10 text-emerald-700 dark:text-emerald-400 font-mono font-bold hover:bg-emerald-500/20"
              >
                SUMA.JEŻELI
              </button>
              <button
                onClick={() => onInsertFunctionTemplate('=ZŁĄCZ.TEKSTY(tekst1; " "; tekst2)')}
                className="px-2 py-1 rounded bg-rose-500/10 text-rose-700 dark:text-rose-400 font-mono font-bold hover:bg-rose-500/20"
              >
                ZŁĄCZ.TEKSTY
              </button>
              <button
                onClick={() => onInsertFunctionTemplate('=ZAOKR(liczba; miejsca)')}
                className="px-2 py-1 rounded bg-neutral-200 dark:bg-neutral-700 font-mono font-bold hover:bg-neutral-300"
              >
                ZAOKR
              </button>
              <button
                onClick={() => onInsertFunctionTemplate('=DZIŚ()')}
                className="px-2 py-1 rounded bg-neutral-200 dark:bg-neutral-700 font-mono font-bold hover:bg-neutral-300"
              >
                DZIŚ()
              </button>
            </div>
          </div>
        )}

        {/* --- TAB 4: DANE (DATA) --- */}
        {activeTab === 'data' && (
          <div className="flex items-center gap-3">
            <button
              onClick={() => onSort('asc')}
              className="flex items-center gap-1.5 px-2.5 py-1.5 rounded hover:bg-neutral-100 dark:hover:bg-neutral-700 border border-neutral-300 dark:border-neutral-600"
            >
              <SortAsc size={14} />
              <span>Sortuj A do Z</span>
            </button>
            <button
              onClick={() => onSort('desc')}
              className="flex items-center gap-1.5 px-2.5 py-1.5 rounded hover:bg-neutral-100 dark:hover:bg-neutral-700 border border-neutral-300 dark:border-neutral-600"
            >
              <SortDesc size={14} />
              <span>Sortuj Z do A</span>
            </button>
            <div className="h-4 w-px bg-neutral-300 dark:bg-neutral-700" />
            <button
              onClick={onExportCSV}
              className="flex items-center gap-1.5 px-2.5 py-1.5 rounded hover:bg-neutral-100 dark:hover:bg-neutral-700 border border-neutral-300 dark:border-neutral-600"
            >
              <Download size={14} />
              <span>Eksportuj do CSV</span>
            </button>
            <button
              onClick={onImportCSV}
              className="flex items-center gap-1.5 px-2.5 py-1.5 rounded hover:bg-neutral-100 dark:hover:bg-neutral-700 border border-neutral-300 dark:border-neutral-600"
            >
              <Upload size={14} />
              <span>Importuj z CSV</span>
            </button>
          </div>
        )}

        {/* --- TAB 5: AKADEMIA & ZADANIA (LEARN) --- */}
        {activeTab === 'learn' && (
          <div className="flex items-center gap-3">
            <button
              onClick={onToggleLearningPanel}
              className="flex items-center gap-1.5 px-3 py-1.5 rounded bg-[#107c41] text-white font-semibold hover:bg-[#0d6434] shadow-xs"
            >
              <BookOpen size={14} />
              <span>Otwórz Panel Zadań</span>
            </button>
            <button
              onClick={onShowLessonHint}
              className="flex items-center gap-1.5 px-3 py-1.5 rounded bg-amber-500/10 text-amber-700 dark:text-amber-300 font-medium hover:bg-amber-500/20 border border-amber-500/30"
            >
              <Lightbulb size={14} className="text-amber-500" />
              <span>Pokaż podpowiedź do zadania</span>
            </button>
            <button
              onClick={onResetLessonData}
              className="flex items-center gap-1.5 px-2.5 py-1.5 rounded hover:bg-neutral-100 dark:hover:bg-neutral-700 border border-neutral-300 dark:border-neutral-600 text-neutral-600 dark:text-neutral-400"
            >
              <RotateCcw size={13} />
              <span>Zresetuj arkusz lekcji</span>
            </button>
          </div>
        )}

        {/* --- TAB 6: RYWALIZACJA & LIGI (COMPETE) --- */}
        {activeTab === 'compete' && (
          <div className="flex items-center gap-3">
            <button
              onClick={onOpenLeaderboard}
              className="flex items-center gap-1.5 px-3 py-1.5 rounded bg-amber-500 text-neutral-900 font-bold hover:bg-amber-400 shadow-xs"
            >
              <Trophy size={15} />
              <span>Tablica Wyników & Ligi</span>
            </button>
            <button
              onClick={onOpenSpeedChallenge}
              className="flex items-center gap-1.5 px-3 py-1.5 rounded bg-emerald-700 text-white font-semibold hover:bg-emerald-800 shadow-xs"
            >
              <Timer size={15} />
              <span>Szybki Pojedynek na Czas (60s)</span>
            </button>
          </div>
        )}

        {/* --- TAB 7: WIDOK (VIEW) --- */}
        {activeTab === 'view' && (
          <div className="flex items-center gap-3">
            <button
              onClick={onToggleGridlines}
              className={`flex items-center gap-1.5 px-2.5 py-1.5 rounded border transition-colors ${
                showGridlines
                  ? 'bg-[#107c41]/10 border-[#107c41] text-[#107c41] dark:text-emerald-400 font-semibold'
                  : 'border-neutral-300 dark:border-neutral-600'
              }`}
            >
              <Grid size={14} />
              <span>Linie siatki: {showGridlines ? 'Włączone' : 'Wyłączone'}</span>
            </button>

            <button
              onClick={onToggleHeaders}
              className={`flex items-center gap-1.5 px-2.5 py-1.5 rounded border transition-colors ${
                showHeaders
                  ? 'bg-[#107c41]/10 border-[#107c41] text-[#107c41] dark:text-emerald-400 font-semibold'
                  : 'border-neutral-300 dark:border-neutral-600'
              }`}
            >
              <Table size={14} />
              <span>Nagłówki wierszy i kolumn</span>
            </button>

            <div className="h-4 w-px bg-neutral-300 dark:bg-neutral-700" />

            <div className="flex items-center gap-1 text-xs">
              <span className="text-neutral-500">Skala:</span>
              {[75, 100, 125, 150].map((z) => (
                <button
                  key={z}
                  onClick={() => onSetZoom(z)}
                  className={`px-2 py-0.5 rounded text-[11px] ${
                    zoomLevel === z
                      ? 'bg-[#107c41] text-white font-bold'
                      : 'hover:bg-neutral-100 dark:hover:bg-neutral-700'
                  }`}
                >
                  {z}%
                </button>
              ))}
            </div>
          </div>
        )}
      </div>
    </div>
  );
};
