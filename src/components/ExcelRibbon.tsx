import React, { useState, useEffect } from 'react';
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
  Paintbrush,
  Copy,
  Scissors,
  ClipboardPaste,
} from 'lucide-react';
import { CellStyle, BorderActionType, BorderLineStyle } from '../types/excel';

export type RibbonTab = 'home' | 'insert' | 'formulas' | 'data' | 'view' | 'learn' | 'compete';

interface ExcelRibbonProps {
  activeTab: RibbonTab;
  onTabChange: (tab: RibbonTab) => void;
  currentStyle?: CellStyle;
  onApplyStyle: (stylePatch: Partial<CellStyle>) => void;
  onApplyBorder?: (
    action: BorderActionType,
    options?: { lineStyle?: BorderLineStyle; lineWidth?: number; lineColor?: string }
  ) => void;
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
  isFormatPainterActive?: boolean;
  isFormatPainterLocked?: boolean;
  onToggleFormatPainter?: (lock?: boolean) => void;
  onCopy?: () => void;
  onCut?: () => void;
  onPaste?: () => void;
}

export const ExcelRibbon: React.FC<ExcelRibbonProps> = ({
  activeTab,
  onTabChange,
  currentStyle = {},
  onApplyStyle,
  onApplyBorder,
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
  isFormatPainterActive = false,
  isFormatPainterLocked = false,
  onToggleFormatPainter,
  onCopy,
  onCut,
  onPaste,
}) => {
  const [showAutoSumMenu, setShowAutoSumMenu] = useState(false);
  const [showBorderMenu, setShowBorderMenu] = useState(false);
  const [showFillColorPicker, setShowFillColorPicker] = useState(false);
  const [showTextColorPicker, setShowTextColorPicker] = useState(false);

  const [selectedFillColor, setSelectedFillColor] = useState<string>('#fef3c7');
  const [selectedTextColor, setSelectedTextColor] = useState<string>('#dc2626');

  // Rich Excel Border Options State
  const [selectedBorderLineStyle, setSelectedBorderLineStyle] = useState<BorderLineStyle>('solid');
  const [selectedBorderLineWidth, setSelectedBorderLineWidth] = useState<number>(1);
  const [selectedBorderColor, setSelectedBorderColor] = useState<string>('#000000');
  const [lastBorderAction, setLastBorderAction] = useState<BorderActionType>('all');

  // Close all open dropdowns on outside click
  useEffect(() => {
    const handleGlobalClick = () => {
      setShowAutoSumMenu(false);
      setShowBorderMenu(false);
      setShowFillColorPicker(false);
      setShowTextColorPicker(false);
    };
    window.addEventListener('click', handleGlobalClick);
    return () => window.removeEventListener('click', handleGlobalClick);
  }, []);

  const THEME_COLORS = [
    '#000000', '#ffffff', '#e5e7eb', '#9ca3af', '#4b5563', '#1f2937',
    '#ef4444', '#dc2626', '#b91c1c', '#f97316', '#ea580c', '#d97706',
    '#eab308', '#ca8a04', '#84cc16', '#65a30d', '#22c55e', '#16a34a',
    '#107c41', '#059669', '#06b6d4', '#0284c7', '#0f6cbd', '#2563eb',
    '#4f46e5', '#7c3aed', '#9333ea', '#c026d3', '#db2777', '#e11d48',
  ];

  return (
    <div className={`flex flex-col border-b select-none transition-colors relative z-30 ${
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
      <div className={`p-1.5 min-h-[58px] flex items-center gap-3 text-xs overflow-visible ${
        isDarkMode ? 'bg-[#2b2b2b]' : 'bg-white'
      }`}>
        {/* --- TAB 1: NARZĘDZIA GŁÓWNE (HOME) --- */}
        {activeTab === 'home' && (
          <div className="flex items-center gap-3">
            {/* Clipboard (Schowek) & Format Painter Group */}
            <div className="flex items-center gap-1.5 pr-3 border-r border-neutral-300 dark:border-neutral-700">
              {/* Paste button */}
              <button
                type="button"
                onClick={onPaste}
                title="Wklej zawartość (Ctrl+V)"
                className="flex flex-col items-center justify-center px-2 py-1 rounded hover:bg-neutral-100 dark:hover:bg-neutral-700 text-neutral-700 dark:text-neutral-200 transition-colors cursor-pointer group"
              >
                <ClipboardPaste size={16} className="text-[#107c41]" />
                <span className="text-[10px] mt-0.5 font-medium leading-none">Wklej</span>
              </button>

              <div className="flex flex-col gap-0.5">
                <div className="flex items-center gap-0.5">
                  <button
                    type="button"
                    onClick={onCut}
                    title="Wytnij (Ctrl+X)"
                    className="p-1 rounded hover:bg-neutral-100 dark:hover:bg-neutral-700 text-neutral-700 dark:text-neutral-200 transition-colors cursor-pointer"
                  >
                    <Scissors size={12} />
                  </button>
                  <button
                    type="button"
                    onClick={onCopy}
                    title="Kopiuj (Ctrl+C)"
                    className="p-1 rounded hover:bg-neutral-100 dark:hover:bg-neutral-700 text-neutral-700 dark:text-neutral-200 transition-colors cursor-pointer"
                  >
                    <Copy size={12} />
                  </button>
                </div>

                {/* Malarz formatów (Format Painter) */}
                <button
                  type="button"
                  onClick={() => onToggleFormatPainter?.(false)}
                  onDoubleClick={() => onToggleFormatPainter?.(true)}
                  title="Malarz formatów (Kopiuje formatowanie z zaznaczonej komórki. Kliknij dwukrotnie, aby zastosować wielokrotnie)"
                  className={`flex items-center gap-1 px-1.5 py-0.5 rounded transition-all cursor-pointer select-none ${
                    isFormatPainterActive
                      ? 'bg-[#107c41] text-white font-semibold shadow-xs ring-2 ring-emerald-500 ring-offset-1'
                      : 'hover:bg-neutral-100 dark:hover:bg-neutral-700 text-neutral-700 dark:text-neutral-200'
                  }`}
                >
                  <Paintbrush
                    size={12}
                    className={isFormatPainterActive ? 'text-amber-300 animate-pulse' : 'text-amber-600 dark:text-amber-400'}
                  />
                  <span className="text-[11px] leading-tight">Malarz formatów</span>
                  {isFormatPainterLocked && (
                    <span className="text-[9px] bg-black/30 text-amber-200 px-1 py-0.2 rounded font-bold ml-0.5">
                      PIN
                    </span>
                  )}
                </button>
              </div>
            </div>

            {/* Font Style Group */}
            <div className="flex items-center gap-1.5 pr-3 border-r border-neutral-300 dark:border-neutral-700">
              <button
                type="button"
                onClick={() => onApplyStyle({ bold: !currentStyle.bold })}
                title="Pogrubienie (Ctrl+B)"
                className={`p-1.5 rounded transition-colors ${
                  currentStyle.bold ? 'bg-[#107c41] text-white' : 'hover:bg-neutral-100 dark:hover:bg-neutral-700'
                }`}
              >
                <Bold size={14} />
              </button>
              <button
                type="button"
                onClick={() => onApplyStyle({ italic: !currentStyle.italic })}
                title="Kursywa (Ctrl+I)"
                className={`p-1.5 rounded transition-colors ${
                  currentStyle.italic ? 'bg-[#107c41] text-white' : 'hover:bg-neutral-100 dark:hover:bg-neutral-700'
                }`}
              >
                <Italic size={14} />
              </button>
              <button
                type="button"
                onClick={() => onApplyStyle({ underline: !currentStyle.underline })}
                title="Podkreślenie (Ctrl+U)"
                className={`p-1.5 rounded transition-colors ${
                  currentStyle.underline ? 'bg-[#107c41] text-white' : 'hover:bg-neutral-100 dark:hover:bg-neutral-700'
                }`}
              >
                <Underline size={14} />
              </button>

              <div className="h-4 w-px bg-neutral-200 dark:bg-neutral-700 mx-0.5" />

              {/* Fill Color Picker (Split Button) */}
              <div className="relative flex items-center rounded border border-transparent hover:border-neutral-300 dark:hover:border-neutral-600 hover:bg-neutral-100 dark:hover:bg-neutral-700">
                <button
                  type="button"
                  onClick={() => onApplyStyle({ bgColor: currentStyle.bgColor || selectedFillColor || '#fef3c7' })}
                  title={`Zastosuj kolor wypełnienia (${selectedFillColor || '#fef3c7'})`}
                  className="p-1 rounded-l flex flex-col items-center justify-center min-w-[24px]"
                >
                  <PaintBucket size={13} className="text-neutral-700 dark:text-neutral-200" />
                  <span
                    className="w-3.5 h-[3px] rounded-xs mt-0.5"
                    style={{ backgroundColor: currentStyle.bgColor || selectedFillColor || '#fef3c7' }}
                  />
                </button>
                <button
                  type="button"
                  onClick={(e) => {
                    e.stopPropagation();
                    setShowFillColorPicker(!showFillColorPicker);
                    setShowTextColorPicker(false);
                    setShowBorderMenu(false);
                    setShowAutoSumMenu(false);
                  }}
                  title="Wybierz kolor wypełnienia"
                  className="px-0.5 py-2 rounded-r hover:bg-neutral-200 dark:hover:bg-neutral-600 cursor-pointer"
                >
                  <ChevronDown size={10} />
                </button>

                {showFillColorPicker && (
                  <div
                    onClick={(e) => e.stopPropagation()}
                    className="absolute top-9 left-0 z-50 p-2.5 bg-white dark:bg-[#2b2b2b] shadow-2xl rounded-md border border-neutral-300 dark:border-neutral-700 w-48 animate-in fade-in zoom-in-95 duration-100"
                  >
                    <div className="text-[11px] font-semibold text-neutral-500 dark:text-neutral-400 mb-1.5">
                      Kolor wypełnienia
                    </div>
                    <button
                      type="button"
                      onClick={() => {
                        onApplyStyle({ bgColor: undefined });
                        setShowFillColorPicker(false);
                      }}
                      className="w-full flex items-center gap-2 px-1.5 py-1 text-xs hover:bg-neutral-100 dark:hover:bg-neutral-700 rounded mb-2 border border-neutral-200 dark:border-neutral-700 cursor-pointer"
                    >
                      <span className="w-4 h-4 rounded-xs border border-dashed border-neutral-400 flex items-center justify-center text-[9px] text-neutral-500">✕</span>
                      <span>Brak wypełnienia</span>
                    </button>

                    <div className="text-[10px] uppercase font-semibold text-neutral-400 mb-1">Paleta kolorów</div>
                    <div className="grid grid-cols-6 gap-1">
                      {THEME_COLORS.map((c) => (
                        <button
                          key={c}
                          type="button"
                          onClick={() => {
                            onApplyStyle({ bgColor: c });
                            setSelectedFillColor(c);
                            setShowFillColorPicker(false);
                          }}
                          style={{ backgroundColor: c }}
                          title={c}
                          className="w-5 h-5 rounded-xs border border-neutral-300 dark:border-neutral-600 hover:scale-125 transition-transform cursor-pointer"
                        />
                      ))}
                    </div>
                  </div>
                )}
              </div>

              {/* Font Color Picker (Dedicated Split Button with 'A' underline) */}
              <div className="relative flex items-center rounded border border-transparent hover:border-neutral-300 dark:hover:border-neutral-600 hover:bg-neutral-100 dark:hover:bg-neutral-700">
                <button
                  type="button"
                  onClick={() => onApplyStyle({ textColor: currentStyle.textColor || selectedTextColor || '#dc2626' })}
                  title={`Zastosuj kolor czcionki (${selectedTextColor || 'Czerwony'})`}
                  className="p-1 rounded-l flex flex-col items-center justify-center min-w-[24px]"
                >
                  <span className="font-bold font-serif text-[13px] leading-none text-neutral-900 dark:text-white">A</span>
                  <span
                    className="w-3.5 h-[3px] rounded-xs mt-0.5"
                    style={{ backgroundColor: currentStyle.textColor || selectedTextColor || '#dc2626' }}
                  />
                </button>
                <button
                  type="button"
                  onClick={(e) => {
                    e.stopPropagation();
                    setShowTextColorPicker(!showTextColorPicker);
                    setShowFillColorPicker(false);
                    setShowBorderMenu(false);
                    setShowAutoSumMenu(false);
                  }}
                  title="Wybierz kolor czcionki"
                  className="px-0.5 py-2 rounded-r hover:bg-neutral-200 dark:hover:bg-neutral-600 cursor-pointer"
                >
                  <ChevronDown size={10} />
                </button>

                {showTextColorPicker && (
                  <div
                    onClick={(e) => e.stopPropagation()}
                    className="absolute top-9 left-0 z-50 p-2.5 bg-white dark:bg-[#2b2b2b] shadow-2xl rounded-md border border-neutral-300 dark:border-neutral-700 w-48 animate-in fade-in zoom-in-95 duration-100"
                  >
                    <div className="text-[11px] font-semibold text-neutral-500 dark:text-neutral-400 mb-1.5">
                      Kolor czcionki
                    </div>
                    <button
                      type="button"
                      onClick={() => {
                        onApplyStyle({ textColor: undefined });
                        setSelectedTextColor('#000000');
                        setShowTextColorPicker(false);
                      }}
                      className="w-full flex items-center gap-2 px-1.5 py-1 text-xs hover:bg-neutral-100 dark:hover:bg-neutral-700 rounded mb-2 border border-neutral-200 dark:border-neutral-700 cursor-pointer"
                    >
                      <span className="w-4 h-4 rounded-xs border border-neutral-300 bg-neutral-900 dark:bg-white" />
                      <span>Automatyczny (Domyślny)</span>
                    </button>

                    <div className="text-[10px] uppercase font-semibold text-neutral-400 mb-1">Paleta kolorów</div>
                    <div className="grid grid-cols-6 gap-1">
                      {THEME_COLORS.map((c) => (
                        <button
                          key={c}
                          type="button"
                          onClick={() => {
                            onApplyStyle({ textColor: c });
                            setSelectedTextColor(c);
                            setShowTextColorPicker(false);
                          }}
                          style={{ backgroundColor: c }}
                          title={c}
                          className="w-5 h-5 rounded-xs border border-neutral-300 dark:border-neutral-600 hover:scale-125 transition-transform cursor-pointer"
                        />
                      ))}
                    </div>
                  </div>
                )}
              </div>

              {/* Complete Office 365 Borders Suite */}
              <div className="relative flex items-center rounded border border-transparent hover:border-neutral-300 dark:hover:border-neutral-600 hover:bg-neutral-100 dark:hover:bg-neutral-700">
                {/* Main button: applies last selected border action */}
                <button
                  type="button"
                  onClick={() => {
                    if (onApplyBorder) {
                      onApplyBorder(lastBorderAction, {
                        lineStyle: selectedBorderLineStyle,
                        lineWidth: selectedBorderLineWidth,
                        lineColor: selectedBorderColor,
                      });
                    } else {
                      onApplyStyle({ border: lastBorderAction });
                    }
                  }}
                  title={`Zastosuj krawędzie: ${lastBorderAction} (${selectedBorderLineWidth}px ${selectedBorderLineStyle})`}
                  className="p-1 rounded-l flex items-center justify-center min-w-[24px] cursor-pointer"
                >
                  <Table size={13} className="text-neutral-700 dark:text-neutral-200" />
                </button>

                {/* Dropdown Chevron */}
                <button
                  type="button"
                  onClick={(e) => {
                    e.stopPropagation();
                    setShowBorderMenu(!showBorderMenu);
                    setShowFillColorPicker(false);
                    setShowTextColorPicker(false);
                    setShowAutoSumMenu(false);
                  }}
                  title="Wybierz styl i krawędzie komórki"
                  className="px-0.5 py-2 rounded-r hover:bg-neutral-200 dark:hover:bg-neutral-600 cursor-pointer"
                >
                  <ChevronDown size={10} />
                </button>

                {showBorderMenu && (
                  <div
                    onClick={(e) => e.stopPropagation()}
                    className="absolute top-9 left-0 z-50 p-2 bg-white dark:bg-[#282828] shadow-2xl rounded-lg border border-neutral-300 dark:border-neutral-700 flex flex-col gap-2 w-72 max-h-[480px] overflow-y-auto text-xs animate-in fade-in zoom-in-95 duration-100 select-none scrollbar-thin"
                  >
                    {/* Header */}
                    <div className="flex items-center justify-between pb-1.5 border-b border-neutral-200 dark:border-neutral-700">
                      <span className="font-semibold text-[11px] text-neutral-600 dark:text-neutral-300">
                        Krawędzie komórek
                      </span>
                      <div className="flex items-center gap-1.5 text-[10px] text-neutral-500 dark:text-neutral-400">
                        <span
                          className="w-3.5 h-3.5 rounded-full border border-neutral-400"
                          style={{ backgroundColor: selectedBorderColor }}
                        />
                        <span className="font-mono">{selectedBorderLineWidth}px</span>
                      </div>
                    </div>

                    {/* Standard Border Actions */}
                    <div className="flex flex-col gap-0.5">
                      <div className="text-[10px] uppercase font-semibold text-neutral-400 dark:text-neutral-500 px-1 mb-0.5">
                        Krawędzie podstawowe
                      </div>

                      {[
                        {
                          id: 'bottom' as BorderActionType,
                          label: 'Krawędź dolna',
                          preview: (
                            <div className="w-4 h-4 border border-dashed border-neutral-300 dark:border-neutral-600 relative shrink-0">
                              <div className="absolute bottom-0 left-0 right-0 h-[2px] bg-neutral-900 dark:bg-white" />
                            </div>
                          ),
                        },
                        {
                          id: 'top' as BorderActionType,
                          label: 'Krawędź górna',
                          preview: (
                            <div className="w-4 h-4 border border-dashed border-neutral-300 dark:border-neutral-600 relative shrink-0">
                              <div className="absolute top-0 left-0 right-0 h-[2px] bg-neutral-900 dark:bg-white" />
                            </div>
                          ),
                        },
                        {
                          id: 'left' as BorderActionType,
                          label: 'Krawędź lewa',
                          preview: (
                            <div className="w-4 h-4 border border-dashed border-neutral-300 dark:border-neutral-600 relative shrink-0">
                              <div className="absolute top-0 bottom-0 left-0 w-[2px] bg-neutral-900 dark:bg-white" />
                            </div>
                          ),
                        },
                        {
                          id: 'right' as BorderActionType,
                          label: 'Krawędź prawa',
                          preview: (
                            <div className="w-4 h-4 border border-dashed border-neutral-300 dark:border-neutral-600 relative shrink-0">
                              <div className="absolute top-0 bottom-0 right-0 w-[2px] bg-neutral-900 dark:bg-white" />
                            </div>
                          ),
                        },
                        {
                          id: 'none' as BorderActionType,
                          label: 'Brak krawędzi',
                          preview: (
                            <div className="w-4 h-4 border border-dashed border-neutral-300 dark:border-neutral-600 flex items-center justify-center text-rose-500 font-bold text-[9px] shrink-0">
                              ✕
                            </div>
                          ),
                        },
                        {
                          id: 'all' as BorderActionType,
                          label: 'Wszystkie krawędzie',
                          preview: (
                            <div className="w-4 h-4 grid grid-cols-2 grid-rows-2 border border-neutral-900 dark:border-white shrink-0">
                              <div className="border-r border-b border-neutral-900 dark:border-white" />
                              <div className="border-b border-neutral-900 dark:border-white" />
                              <div className="border-r border-neutral-900 dark:border-white" />
                              <div />
                            </div>
                          ),
                        },
                        {
                          id: 'outside' as BorderActionType,
                          label: 'Krawędzie zewnętrzne',
                          preview: (
                            <div className="w-4 h-4 border border-neutral-900 dark:border-white shrink-0" />
                          ),
                        },
                        {
                          id: 'thick-outside' as BorderActionType,
                          label: 'Gruba krawędź zewnętrzna',
                          preview: (
                            <div className="w-4 h-4 border-2 border-neutral-900 dark:border-white shrink-0" />
                          ),
                        },
                      ].map((item) => (
                        <button
                          key={item.id}
                          type="button"
                          onClick={() => {
                            setLastBorderAction(item.id);
                            if (onApplyBorder) {
                              onApplyBorder(item.id, {
                                lineStyle: selectedBorderLineStyle,
                                lineWidth: selectedBorderLineWidth,
                                lineColor: selectedBorderColor,
                              });
                            } else {
                              onApplyStyle({ border: item.id });
                            }
                            setShowBorderMenu(false);
                          }}
                          className="flex items-center gap-2 px-2 py-1.5 hover:bg-neutral-100 dark:hover:bg-neutral-700 rounded text-left transition-colors cursor-pointer"
                        >
                          {item.preview}
                          <span className="flex-1">{item.label}</span>
                        </button>
                      ))}
                    </div>

                    {/* Specialized Excel Borders */}
                    <div className="flex flex-col gap-0.5 pt-1.5 border-t border-neutral-200 dark:border-neutral-700">
                      <div className="text-[10px] uppercase font-semibold text-neutral-400 dark:text-neutral-500 px-1 mb-0.5">
                        Krawędzie zaawansowane
                      </div>

                      {[
                        {
                          id: 'bottom-double' as BorderActionType,
                          label: 'Podwójna dolna krawędź',
                          preview: (
                            <div className="w-4 h-4 border border-dashed border-neutral-300 dark:border-neutral-600 relative shrink-0">
                              <div className="absolute bottom-[2px] left-0 right-0 h-[1px] bg-neutral-900 dark:bg-white" />
                              <div className="absolute bottom-0 left-0 right-0 h-[1px] bg-neutral-900 dark:bg-white" />
                            </div>
                          ),
                        },
                        {
                          id: 'bottom-thick' as BorderActionType,
                          label: 'Gruba dolna krawędź',
                          preview: (
                            <div className="w-4 h-4 border border-dashed border-neutral-300 dark:border-neutral-600 relative shrink-0">
                              <div className="absolute bottom-0 left-0 right-0 h-[3px] bg-neutral-900 dark:bg-white" />
                            </div>
                          ),
                        },
                        {
                          id: 'top-bottom' as BorderActionType,
                          label: 'Górna i dolna krawędź',
                          preview: (
                            <div className="w-4 h-4 border-t border-b border-neutral-900 dark:border-white relative shrink-0" />
                          ),
                        },
                        {
                          id: 'top-bottom-double' as BorderActionType,
                          label: 'Górna i podwójna dolna krawędź',
                          preview: (
                            <div className="w-4 h-4 border-t border-neutral-900 dark:border-white relative shrink-0">
                              <div className="absolute bottom-[2px] left-0 right-0 h-[1px] bg-neutral-900 dark:bg-white" />
                              <div className="absolute bottom-0 left-0 right-0 h-[1px] bg-neutral-900 dark:bg-white" />
                            </div>
                          ),
                        },
                        {
                          id: 'top-bottom-thick' as BorderActionType,
                          label: 'Górna i gruba dolna krawędź',
                          preview: (
                            <div className="w-4 h-4 border-t border-neutral-900 dark:border-white relative shrink-0">
                              <div className="absolute bottom-0 left-0 right-0 h-[3px] bg-neutral-900 dark:bg-white" />
                            </div>
                          ),
                        },
                        {
                          id: 'inside-horizontal' as BorderActionType,
                          label: 'Krawędzie wewnętrzne poziome',
                          preview: (
                            <div className="w-4 h-4 border border-dashed border-neutral-300 dark:border-neutral-600 flex flex-col justify-center shrink-0">
                              <div className="w-full h-[1.5px] bg-neutral-900 dark:bg-white" />
                            </div>
                          ),
                        },
                        {
                          id: 'inside-vertical' as BorderActionType,
                          label: 'Krawędzie wewnętrzne pionowe',
                          preview: (
                            <div className="w-4 h-4 border border-dashed border-neutral-300 dark:border-neutral-600 flex justify-center shrink-0">
                              <div className="h-full w-[1.5px] bg-neutral-900 dark:bg-white" />
                            </div>
                          ),
                        },
                      ].map((item) => (
                        <button
                          key={item.id}
                          type="button"
                          onClick={() => {
                            setLastBorderAction(item.id);
                            if (onApplyBorder) {
                              onApplyBorder(item.id, {
                                lineStyle: selectedBorderLineStyle,
                                lineWidth: selectedBorderLineWidth,
                                lineColor: selectedBorderColor,
                              });
                            } else {
                              onApplyStyle({ border: item.id });
                            }
                            setShowBorderMenu(false);
                          }}
                          className="flex items-center gap-2 px-2 py-1.5 hover:bg-neutral-100 dark:hover:bg-neutral-700 rounded text-left transition-colors cursor-pointer"
                        >
                          {item.preview}
                          <span className="flex-1">{item.label}</span>
                        </button>
                      ))}
                    </div>

                    {/* Line Style Options */}
                    <div className="flex flex-col gap-1 pt-1.5 border-t border-neutral-200 dark:border-neutral-700">
                      <div className="text-[10px] uppercase font-semibold text-neutral-400 dark:text-neutral-500 px-1">
                        Styl linii (grubość i deseń)
                      </div>
                      <div className="grid grid-cols-2 gap-1 px-1">
                        {[
                          { style: 'solid' as BorderLineStyle, width: 1, label: 'Cienka (1px)' },
                          { style: 'solid' as BorderLineStyle, width: 2, label: 'Średnia (2px)' },
                          { style: 'solid' as BorderLineStyle, width: 3, label: 'Gruba (3px)' },
                          { style: 'dashed' as BorderLineStyle, width: 1, label: 'Kreskowana' },
                          { style: 'dotted' as BorderLineStyle, width: 1, label: 'Kropkowana' },
                          { style: 'double' as BorderLineStyle, width: 3, label: 'Podwójna' },
                        ].map((ls) => {
                          const isSelected =
                            selectedBorderLineStyle === ls.style && selectedBorderLineWidth === ls.width;
                          return (
                            <button
                              key={`${ls.style}-${ls.width}`}
                              type="button"
                              onClick={() => {
                                setSelectedBorderLineStyle(ls.style);
                                setSelectedBorderLineWidth(ls.width);
                              }}
                              className={`flex items-center justify-between px-2 py-1 rounded text-[11px] border cursor-pointer transition-all ${
                                isSelected
                                  ? 'bg-[#107c41]/15 border-[#107c41] text-[#107c41] dark:text-emerald-400 font-semibold'
                                  : 'border-neutral-200 dark:border-neutral-700 hover:bg-neutral-100 dark:hover:bg-neutral-700 text-neutral-700 dark:text-neutral-300'
                              }`}
                            >
                              <span>{ls.label}</span>
                              <div
                                className="w-6 h-0 shrink-0 ml-1"
                                style={{
                                  borderBottomWidth: `${ls.width}px`,
                                  borderBottomStyle: ls.style,
                                  borderBottomColor: 'currentColor',
                                }}
                              />
                            </button>
                          );
                        })}
                      </div>
                    </div>

                    {/* Line Color Picker */}
                    <div className="flex flex-col gap-1 pt-1.5 border-t border-neutral-200 dark:border-neutral-700">
                      <div className="flex items-center justify-between px-1">
                        <span className="text-[10px] uppercase font-semibold text-neutral-400 dark:text-neutral-500">
                          Kolor linii
                        </span>
                        <label className="flex items-center gap-1 text-[10px] text-neutral-500 hover:text-neutral-900 dark:hover:text-white cursor-pointer">
                          <span>Własny</span>
                          <input
                            type="color"
                            value={selectedBorderColor}
                            onChange={(e) => setSelectedBorderColor(e.target.value)}
                            className="w-4 h-4 rounded cursor-pointer border-0 p-0 bg-transparent"
                          />
                        </label>
                      </div>

                      <div className="grid grid-cols-8 gap-1 px-1">
                        {[
                          '#000000',
                          '#4b5563',
                          '#9ca3af',
                          '#107c41',
                          '#2563eb',
                          '#dc2626',
                          '#d97706',
                          '#7c3aed',
                        ].map((c) => (
                          <button
                            key={c}
                            type="button"
                            onClick={() => setSelectedBorderColor(c)}
                            style={{ backgroundColor: c }}
                            title={c}
                            className={`w-6 h-6 rounded-xs border transition-transform cursor-pointer flex items-center justify-center ${
                              selectedBorderColor.toLowerCase() === c.toLowerCase()
                                ? 'scale-110 ring-2 ring-[#107c41] ring-offset-1 border-white'
                                : 'border-neutral-300 dark:border-neutral-600 hover:scale-105'
                            }`}
                          >
                            {selectedBorderColor.toLowerCase() === c.toLowerCase() && (
                              <span className="text-[10px] text-white font-bold drop-shadow">✓</span>
                            )}
                          </button>
                        ))}
                      </div>
                    </div>
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
                    type="button"
                    onClick={(e) => {
                      e.stopPropagation();
                      setShowAutoSumMenu(!showAutoSumMenu);
                      setShowFillColorPicker(false);
                      setShowTextColorPicker(false);
                      setShowBorderMenu(false);
                    }}
                    className="px-1 py-1 border-l border-neutral-300 dark:border-neutral-600 hover:bg-neutral-200 dark:hover:bg-neutral-600 cursor-pointer"
                  >
                    <ChevronDown size={11} />
                  </button>
                </div>
                {showAutoSumMenu && (
                  <div
                    onClick={(e) => e.stopPropagation()}
                    className="absolute top-9 left-0 z-50 p-1.5 bg-white dark:bg-[#2b2b2b] shadow-2xl rounded-md border border-neutral-300 dark:border-neutral-700 flex flex-col gap-0.5 w-36 text-xs animate-in fade-in zoom-in-95 duration-100"
                  >
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
