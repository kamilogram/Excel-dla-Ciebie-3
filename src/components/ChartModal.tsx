import React, { useState } from 'react';
import { BarChart2, LineChart, PieChart, Download, X } from 'lucide-react';
import { SheetData, SelectionRange } from '../types/excel';
import { letterToColIndex, parseNumericValue } from '../utils/formulaEngine';

interface ChartModalProps {
  isOpen: boolean;
  onClose: () => void;
  activeSheet: SheetData;
  selectionRange: SelectionRange | null;
  initialType?: 'bar' | 'line' | 'pie';
  isDarkMode: boolean;
}

export const ChartModal: React.FC<ChartModalProps> = ({
  isOpen,
  onClose,
  activeSheet,
  selectionRange,
  initialType = 'bar',
  isDarkMode,
}) => {
  const [chartType, setChartType] = useState<'bar' | 'line' | 'pie'>(initialType);

  if (!isOpen) return null;

  // Extract chart data from active sheet selection or sensible defaults
  const getChartData = () => {
    let rows: number[] = [3, 4, 5, 6, 7];
    let labelCol = 'A';
    let valueCol = 'C';

    if (selectionRange) {
      const c1 = letterToColIndex(selectionRange.startCol);
      const c2 = letterToColIndex(selectionRange.endCol);
      const minColIdx = Math.min(c1, c2);
      const maxColIdx = Math.max(c1, c2);

      const minRow = Math.min(selectionRange.startRow, selectionRange.endRow);
      const maxRow = Math.max(selectionRange.startRow, selectionRange.endRow);

      rows = [];
      for (let r = minRow; r <= maxRow; r++) rows.push(r);

      labelCol = String.fromCharCode(minColIdx + 65);
      valueCol = String.fromCharCode(maxColIdx + 65);
    }

    const items: { label: string; value: number }[] = [];
    for (const r of rows) {
      const labelCell = activeSheet.cells[`${labelCol}${r}`];
      const valCell = activeSheet.cells[`${valueCol}${r}`];

      const rawVal = valCell?.computed !== undefined ? valCell.computed : valCell?.raw || '0';
      const num = parseNumericValue(rawVal);
      const label = labelCell?.raw || `Pozycja ${r}`;

      if (num > 0) {
        items.push({ label, value: num });
      }
    }

    if (items.length === 0) {
      return [
        { label: 'Północ', value: 5200 },
        { label: 'Południe', value: 4100 },
        { label: 'Wschód', value: 3200 },
        { label: 'Zachód', value: 6200 },
        { label: 'Centralny', value: 2900 },
      ];
    }
    return items;
  };

  const data = getChartData();
  const maxValue = Math.max(...data.map((d) => d.value), 1);
  const totalValue = data.reduce((acc, d) => acc + d.value, 0);

  const colors = ['#107c41', '#0f6cbd', '#d83b01', '#881798', '#00b7c3', '#ffaa44'];

  return (
    <div className="fixed inset-0 z-50 bg-black/60 backdrop-blur-xs flex items-center justify-center p-4">
      <div className={`w-full max-w-2xl rounded-xl shadow-2xl border flex flex-col overflow-hidden animate-in fade-in zoom-in-95 duration-150 ${
        isDarkMode ? 'bg-[#202020] border-[#383838] text-white' : 'bg-white border-neutral-200 text-neutral-900'
      }`}>
        {/* Header */}
        <div className="flex items-center justify-between px-6 py-4 border-b border-neutral-200 dark:border-neutral-800">
          <div className="flex items-center gap-2">
            <div className="p-2 bg-[#107c41]/10 text-[#107c41] dark:text-emerald-400 rounded-lg">
              <BarChart2 size={20} />
            </div>
            <div>
              <h2 className="text-base font-bold">Wstaw Wykres Excel 365</h2>
              <div className="text-xs text-neutral-500">Wizualizacja danych z aktywnego arkusza</div>
            </div>
          </div>
          <button onClick={onClose} className="p-1 rounded text-neutral-400 hover:text-neutral-700 dark:hover:text-white">
            <X size={18} />
          </button>
        </div>

        {/* Chart Type Selector */}
        <div className={`flex items-center px-6 py-2.5 gap-2 border-b ${
          isDarkMode ? 'bg-[#262626] border-[#383838]' : 'bg-neutral-50 border-neutral-200'
        }`}>
          <span className="text-xs font-semibold text-neutral-500 mr-2">Typ wykresu:</span>
          <button
            onClick={() => setChartType('bar')}
            className={`flex items-center gap-1.5 px-3 py-1.5 rounded-md text-xs font-semibold transition-colors ${
              chartType === 'bar'
                ? 'bg-[#107c41] text-white shadow-xs'
                : 'hover:bg-neutral-200 dark:hover:bg-neutral-700 text-neutral-600 dark:text-neutral-300'
            }`}
          >
            <BarChart2 size={14} />
            <span>Kolumnowy</span>
          </button>
          <button
            onClick={() => setChartType('line')}
            className={`flex items-center gap-1.5 px-3 py-1.5 rounded-md text-xs font-semibold transition-colors ${
              chartType === 'line'
                ? 'bg-[#107c41] text-white shadow-xs'
                : 'hover:bg-neutral-200 dark:hover:bg-neutral-700 text-neutral-600 dark:text-neutral-300'
            }`}
          >
            <LineChart size={14} />
            <span>Liniowy</span>
          </button>
          <button
            onClick={() => setChartType('pie')}
            className={`flex items-center gap-1.5 px-3 py-1.5 rounded-md text-xs font-semibold transition-colors ${
              chartType === 'pie'
                ? 'bg-[#107c41] text-white shadow-xs'
                : 'hover:bg-neutral-200 dark:hover:bg-neutral-700 text-neutral-600 dark:text-neutral-300'
            }`}
          >
            <PieChart size={14} />
            <span>Kołowy</span>
          </button>
        </div>

        {/* Chart Viewport */}
        <div className="p-6 flex flex-col items-center justify-center min-h-[300px]">
          {chartType === 'bar' && (
            <div className="w-full flex items-end justify-around gap-3 h-56 pt-6 border-b border-neutral-300 dark:border-neutral-700 pb-2">
              {data.map((item, idx) => {
                const heightPercent = Math.round((item.value / maxValue) * 100);
                return (
                  <div key={idx} className="flex-1 flex flex-col items-center h-full justify-end group">
                    <span className="text-[11px] font-mono font-bold mb-1 opacity-0 group-hover:opacity-100 transition-opacity">
                      {item.value.toLocaleString()} zł
                    </span>
                    <div
                      style={{
                        height: `${Math.max(10, heightPercent)}%`,
                        backgroundColor: colors[idx % colors.length],
                      }}
                      className="w-full max-w-[48px] rounded-t transition-all hover:brightness-110 shadow-sm"
                    />
                    <span className="text-[11px] font-medium text-neutral-600 dark:text-neutral-400 mt-2 truncate w-full text-center">
                      {item.label}
                    </span>
                  </div>
                );
              })}
            </div>
          )}

          {chartType === 'line' && (
            <div className="w-full h-56 relative flex items-center justify-center">
              <svg className="w-full h-44 overflow-visible" viewBox="0 0 500 160">
                {/* Horizontal reference lines */}
                <line x1="0" y1="20" x2="500" y2="20" stroke="#888" strokeOpacity="0.2" strokeDasharray="4 4" />
                <line x1="0" y1="80" x2="500" y2="80" stroke="#888" strokeOpacity="0.2" strokeDasharray="4 4" />
                <line x1="0" y1="140" x2="500" y2="140" stroke="#888" strokeOpacity="0.3" />

                {/* Polyline */}
                <polyline
                  fill="none"
                  stroke="#107c41"
                  strokeWidth="3.5"
                  strokeLinecap="round"
                  strokeLinejoin="round"
                  points={data
                    .map((d, i) => {
                      const x = (i / (data.length - 1 || 1)) * 460 + 20;
                      const y = 140 - (d.value / maxValue) * 110;
                      return `${x},${y}`;
                    })
                    .join(' ')}
                />

                {/* Data Points */}
                {data.map((d, i) => {
                  const x = (i / (data.length - 1 || 1)) * 460 + 20;
                  const y = 140 - (d.value / maxValue) * 110;
                  return (
                    <g key={i}>
                      <circle cx={x} cy={y} r="5" fill="#ffffff" stroke="#107c41" strokeWidth="3" />
                      <text x={x} y={y - 10} textAnchor="middle" fontSize="10" fill="currentColor" fontWeight="bold">
                        {d.value}
                      </text>
                      <text x={x} y={155} textAnchor="middle" fontSize="10" fill="#888">
                        {d.label}
                      </text>
                    </g>
                  );
                })}
              </svg>
            </div>
          )}

          {chartType === 'pie' && (
            <div className="flex flex-col sm:flex-row items-center justify-center gap-8 py-2">
              <div className="relative w-44 h-44">
                <svg viewBox="0 0 100 100" className="w-full h-full -rotate-90">
                  {(() => {
                    let accumulatedPercent = 0;
                    return data.map((d, idx) => {
                      const percent = d.value / totalValue;
                      const strokeDasharray = `${percent * 314.15} 314.15`;
                      const strokeDashoffset = `-${accumulatedPercent * 314.15}`;
                      accumulatedPercent += percent;

                      return (
                        <circle
                          key={idx}
                          r="25"
                          cx="50"
                          cy="50"
                          fill="transparent"
                          stroke={colors[idx % colors.length]}
                          strokeWidth="50"
                          strokeDasharray={strokeDasharray}
                          strokeDashoffset={strokeDashoffset}
                          className="hover:opacity-90 transition-opacity"
                        />
                      );
                    });
                  })()}
                </svg>
              </div>

              {/* Legend */}
              <div className="space-y-1.5 text-xs">
                {data.map((item, idx) => (
                  <div key={idx} className="flex items-center gap-2">
                    <span
                      className="w-3 h-3 rounded-sm shrink-0"
                      style={{ backgroundColor: colors[idx % colors.length] }}
                    />
                    <span className="font-medium">{item.label}:</span>
                    <strong className="font-mono">
                      {Math.round((item.value / totalValue) * 100)}% ({item.value} zł)
                    </strong>
                  </div>
                ))}
              </div>
            </div>
          )}
        </div>

        {/* Footer */}
        <div className={`px-6 py-3 border-t flex items-center justify-between ${
          isDarkMode ? 'bg-[#262626] border-[#383838]' : 'bg-neutral-50 border-neutral-200'
        }`}>
          <span className="text-xs text-neutral-500">
            Wykres dynamicznie odzwierciedla dane z arkusza Excela.
          </span>
          <button
            onClick={onClose}
            className="px-4 py-1.5 bg-[#107c41] text-white font-bold text-xs rounded-lg hover:bg-[#0d6434] transition-colors"
          >
            Gotowe
          </button>
        </div>
      </div>
    </div>
  );
};
