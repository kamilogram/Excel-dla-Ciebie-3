import React, { useState, useEffect } from 'react';
import {
  CheckCircle2,
  AlertCircle,
  HelpCircle,
  ChevronRight,
  ChevronLeft,
  Award,
  Zap,
  RotateCcw,
  BookOpen,
  Filter,
  Sparkles,
  ArrowRight,
  Eye,
  Lock,
} from 'lucide-react';
import { LearningTask, SheetData, TaskDifficulty } from '../types/excel';

interface LearningPanelProps {
  tasks: LearningTask[];
  currentTaskId: string;
  onSelectTask: (taskId: string) => void;
  completedTaskIds: string[];
  activeSheet: SheetData;
  onNavigateToTaskSheet: (sheetId: string, cellId: string) => void;
  onResetTaskSheet: () => void;
  isDarkMode: boolean;
  onClose: () => void;
}

export const LearningPanel: React.FC<LearningPanelProps> = ({
  tasks,
  currentTaskId,
  onSelectTask,
  completedTaskIds,
  activeSheet,
  onNavigateToTaskSheet,
  onResetTaskSheet,
  isDarkMode,
  onClose,
}) => {
  const [showHint, setShowHint] = useState(false);
  const [showSolution, setShowSolution] = useState(false);
  const [showInstructions, setShowInstructions] = useState(false);
  const [difficultyFilter, setDifficultyFilter] = useState<string>('all');

  useEffect(() => {
    setShowInstructions(false);
    setShowHint(false);
    setShowSolution(false);
  }, [currentTaskId]);

  const currentTask = tasks.find((t) => t.id === currentTaskId) || tasks[0];
  const currentIndex = tasks.findIndex((t) => t.id === currentTaskId);

  // Check validation for current task in active sheet
  const targetCellData = activeSheet.cells[currentTask.targetCell];
  const currentVal = targetCellData?.computed !== undefined ? targetCellData.computed : targetCellData?.raw || '';
  const currentRaw = targetCellData?.raw || '';

  const isCompleted = completedTaskIds.includes(currentTask.id);
  const isRightSheet = activeSheet.id === currentTask.targetSheetId;

  // Real-time check
  let isCurrentlyCorrect = false;
  if (isRightSheet) {
    try {
      isCurrentlyCorrect = currentTask.expectedResultCheck(currentVal, currentRaw, activeSheet);
    } catch {
      isCurrentlyCorrect = false;
    }
  }

  // Filter tasks
  const filteredTasks = tasks.filter((t) => {
    if (difficultyFilter === 'all') return true;
    return t.difficulty === difficultyFilter;
  });

  const getDifficultyBadge = (diff: TaskDifficulty) => {
    switch (diff) {
      case 'podstawowy':
        return <span className="text-[10px] font-bold text-emerald-600 dark:text-emerald-400">Podstawowy</span>;
      case 'średniozaawansowany':
        return <span className="text-[10px] font-bold text-amber-600 dark:text-amber-400">Średni</span>;
      case 'zaawansowany':
        return <span className="text-[10px] font-bold text-blue-600 dark:text-blue-400">Zaawansowany</span>;
      case 'ekspert':
        return <span className="text-[10px] font-bold text-purple-600 dark:text-purple-400">Ekspert</span>;
    }
  };

  return (
    <aside className={`w-80 md:w-96 flex flex-col border-l select-none shrink-0 transition-colors z-20 ${
      isDarkMode ? 'bg-[#202020] border-[#333333] text-neutral-200' : 'bg-[#fafafa] border-[#e5e7eb] text-neutral-800'
    }`}>
      {/* Panel Header */}
      <div className={`flex items-center justify-between px-4 py-3 border-b ${
        isDarkMode ? 'bg-[#262626] border-[#333333]' : 'bg-white border-[#e5e7eb]'
      }`}>
        <div className="flex items-center gap-2">
          <div className="p-1.5 rounded-md bg-[#107c41]/10 text-[#107c41] dark:text-emerald-400">
            <BookOpen size={16} />
          </div>
          <div>
            <h2 className="text-sm font-bold tracking-tight">Akademia Excela</h2>
            <div className="text-[11px] text-neutral-500">
              Ukończono: {completedTaskIds.length} z {tasks.length} zadań
            </div>
          </div>
        </div>
        <button
          onClick={onClose}
          className="text-xs text-neutral-400 hover:text-neutral-700 dark:hover:text-white px-2 py-1 rounded"
        >
          ✕
        </button>
      </div>

      {/* Main Task View */}
      <div className="flex-1 overflow-y-auto p-4 space-y-4">
        {/* Task Title & Metadata */}
        <div className={`p-3.5 rounded-lg border ${
          isCurrentlyCorrect || isCompleted
            ? 'bg-emerald-50 dark:bg-emerald-950/20 border-emerald-300 dark:border-emerald-800'
            : isDarkMode
            ? 'bg-[#262626] border-[#3a3a3a]'
            : 'bg-white border-neutral-200 shadow-xs'
        }`}>
          <div className="flex items-center justify-between gap-2 mb-1.5">
            <div className="flex items-center gap-2">
              {getDifficultyBadge(currentTask.difficulty)}
              <span className="text-[10px] text-neutral-400">·</span>
              <span className="text-[11px] text-neutral-500">{currentTask.category}</span>
            </div>
            <div className="flex items-center gap-1 text-xs font-bold text-amber-600 dark:text-amber-400">
              <Zap size={13} className="fill-amber-400 text-amber-500" />
              <span>+{currentTask.xpReward} XP</span>
            </div>
          </div>

          <h3 className="text-sm font-bold text-neutral-900 dark:text-white mb-1.5">
            {currentIndex + 1}. {currentTask.title}
          </h3>
          <p className="text-xs text-neutral-600 dark:text-neutral-300 leading-relaxed mb-3">
            {currentTask.description}
          </p>

          {/* Action Location Navigation */}
          {!isRightSheet && (
            <button
              onClick={() => onNavigateToTaskSheet(currentTask.targetSheetId, currentTask.targetCell)}
              className="w-full mb-3 flex items-center justify-center gap-1.5 py-1.5 px-3 bg-[#107c41] text-white text-xs font-semibold rounded hover:bg-[#0d6434] transition-colors"
            >
              <span>Przejdź do arkusza zadania:</span>
              <span className="underline">Akademia Excela ({currentTask.targetCell})</span>
              <ArrowRight size={13} />
            </button>
          )}

          {/* Instructions Section - Hidden by default */}
          <div className="space-y-2">
            {!showInstructions ? (
              <button
                onClick={() => setShowInstructions(true)}
                className="w-full flex items-center justify-between py-2 px-3 bg-neutral-100 hover:bg-neutral-200 dark:bg-neutral-800 dark:hover:bg-neutral-750 text-neutral-800 dark:text-neutral-200 rounded-lg text-xs font-semibold border border-neutral-200 dark:border-neutral-700 transition-colors group"
              >
                <div className="flex items-center gap-2">
                  <span className="p-1 rounded bg-[#107c41]/10 text-[#107c41] dark:text-emerald-400 group-hover:scale-110 transition-transform">
                    <Eye size={13} />
                  </span>
                  <span>Pokaż instrukcję krok po kroku</span>
                </div>
                <span className="text-[10px] text-neutral-400 font-normal">Kliknij, aby odsłonić</span>
              </button>
            ) : (
              <div className="space-y-1.5 text-xs bg-neutral-50 dark:bg-[#1f1f1f] p-3 rounded-lg border border-neutral-200 dark:border-[#333333] animate-in fade-in duration-150">
                <div className="flex items-center justify-between pb-1 border-b border-neutral-200 dark:border-neutral-800">
                  <span className="text-[11px] font-bold text-neutral-600 dark:text-neutral-300 uppercase tracking-wide">
                    Instrukcja wykonania zadania:
                  </span>
                  <button
                    onClick={() => setShowInstructions(false)}
                    className="text-[10px] text-neutral-400 hover:text-neutral-700 dark:hover:text-white"
                  >
                    Ukryj instrukcję
                  </button>
                </div>
                {currentTask.instructions.map((step, idx) => (
                  <div key={idx} className="flex items-start gap-2 pt-1">
                    <span className="w-4 h-4 rounded-full bg-[#107c41] text-white text-[10px] flex items-center justify-center shrink-0 mt-0.5 font-bold">
                      {idx + 1}
                    </span>
                    <span className="text-neutral-700 dark:text-neutral-300 leading-snug">{step}</span>
                  </div>
                ))}
              </div>
            )}
          </div>
        </div>

        {/* Real-time Verification Status Box */}
        <div className={`p-3 rounded-lg border ${
          isCurrentlyCorrect || isCompleted
            ? 'bg-emerald-500/10 border-emerald-500/30 text-emerald-900 dark:text-emerald-200'
            : 'bg-amber-500/10 border-amber-500/30 text-amber-900 dark:text-amber-200'
        }`}>
          <div className="flex items-center gap-2 mb-1.5">
            {isCurrentlyCorrect || isCompleted ? (
              <>
                <CheckCircle2 size={16} className="text-emerald-600 dark:text-emerald-400 shrink-0" />
                <span className="text-xs font-bold text-emerald-700 dark:text-emerald-300">
                  Zadanie rozwiązane poprawnie! 🎉
                </span>
              </>
            ) : (
              <>
                <AlertCircle size={16} className="text-amber-600 dark:text-amber-400 shrink-0" />
                <span className="text-xs font-bold text-amber-800 dark:text-amber-300">
                  Weryfikacja w czasie rzeczywistym
                </span>
              </>
            )}
          </div>

          <div className="text-xs space-y-1 text-neutral-600 dark:text-neutral-400">
            <div>
              Komórka docelowa: <strong className="font-mono text-neutral-900 dark:text-white">{currentTask.targetCell}</strong>
            </div>
            <div>
              Aktualna zawartość: <strong className="font-mono text-neutral-900 dark:text-white">{currentRaw || '(pusto)'}</strong>
            </div>
            {currentVal !== undefined && currentVal !== '' && (
              <div>
                Obliczony wynik: <strong className="font-mono text-neutral-900 dark:text-white">{String(currentVal)}</strong>
              </div>
            )}
          </div>
        </div>

        {/* Hint & Solution Area */}
        <div className="space-y-2">
          {!showHint ? (
            <button
              onClick={() => setShowHint(true)}
              className="w-full flex items-center justify-center gap-1.5 py-1.5 text-xs text-amber-700 dark:text-amber-300 bg-amber-500/10 hover:bg-amber-500/20 rounded border border-amber-500/30 transition-colors"
            >
              <HelpCircle size={13} />
              <span>Potrzebujesz podpowiedzi?</span>
            </button>
          ) : (
            <div className="p-3 text-xs bg-amber-50 dark:bg-amber-950/20 border border-amber-200 dark:border-amber-900/40 rounded text-amber-900 dark:text-amber-200">
              <div className="font-bold mb-1">💡 Podpowiedź:</div>
              <p>{currentTask.hint}</p>
            </div>
          )}

          {!showSolution ? (
            <button
              onClick={() => setShowSolution(true)}
              className="w-full flex items-center justify-center gap-1.5 py-1.5 text-xs text-neutral-600 dark:text-neutral-400 hover:text-neutral-900 dark:hover:text-white transition-colors"
            >
              <Eye size={13} />
              <span>Pokaż gotową formułę rozwiązania</span>
            </button>
          ) : (
            <div className="p-3 text-xs bg-blue-50 dark:bg-blue-950/20 border border-blue-200 dark:border-blue-900/40 rounded text-blue-900 dark:text-blue-200 font-mono">
              <div className="font-sans font-bold mb-1 text-blue-800 dark:text-blue-300">Wzorcowa formuła:</div>
              <div className="bg-white dark:bg-black/30 p-1.5 rounded select-all font-bold">
                {currentTask.solutionFormula}
              </div>
              <div className="mt-1 font-sans text-[11px] text-blue-700/80 dark:text-blue-300/80">
                {currentTask.explanation}
              </div>
            </div>
          )}
        </div>

        {/* Navigation buttons: Prev / Next */}
        <div className="flex items-center justify-between pt-2 border-t border-neutral-200 dark:border-neutral-700">
          <button
            onClick={() => {
              if (currentIndex > 0) {
                onSelectTask(tasks[currentIndex - 1].id);
                setShowHint(false);
                setShowSolution(false);
              }
            }}
            disabled={currentIndex === 0}
            className={`flex items-center gap-1 text-xs px-2.5 py-1.5 rounded ${
              currentIndex > 0
                ? 'hover:bg-neutral-200 dark:hover:bg-neutral-700 text-neutral-800 dark:text-neutral-200'
                : 'opacity-40 cursor-not-allowed'
            }`}
          >
            <ChevronLeft size={14} />
            <span>Poprzednie</span>
          </button>

          <button
            onClick={() => {
              if (currentIndex < tasks.length - 1) {
                onSelectTask(tasks[currentIndex + 1].id);
                setShowHint(false);
                setShowSolution(false);
              }
            }}
            disabled={currentIndex === tasks.length - 1}
            className={`flex items-center gap-1 text-xs px-3 py-1.5 rounded font-semibold ${
              currentIndex < tasks.length - 1
                ? 'bg-[#107c41] text-white hover:bg-[#0d6434]'
                : 'opacity-40 cursor-not-allowed'
            }`}
          >
            <span>Następne zadanie</span>
            <ChevronRight size={14} />
          </button>
        </div>

        {/* Full Lessons Curriculum Browser */}
        <div className="pt-3 border-t border-neutral-200 dark:border-neutral-700 space-y-2">
          <div className="flex items-center justify-between">
            <span className="text-xs font-bold text-neutral-700 dark:text-neutral-300">
              Wszystkie lekcje ({tasks.length})
            </span>

            {/* Filter */}
            <select
              value={difficultyFilter}
              onChange={(e) => setDifficultyFilter(e.target.value)}
              className="text-[11px] bg-neutral-100 dark:bg-neutral-800 border border-neutral-300 dark:border-neutral-700 rounded px-1.5 py-0.5 outline-none"
            >
              <option value="all">Wszystkie poziomy</option>
              <option value="podstawowy">Podstawowy</option>
              <option value="średniozaawansowany">Średni</option>
              <option value="zaawansowany">Zaawansowany</option>
              <option value="ekspert">Ekspert</option>
            </select>
          </div>

          <div className="space-y-1 max-h-56 overflow-y-auto pr-1">
            {filteredTasks.map((t, idx) => {
              const isTaskDone = completedTaskIds.includes(t.id);
              const isSelected = t.id === currentTaskId;

              return (
                <div
                  key={t.id}
                  onClick={() => {
                    onSelectTask(t.id);
                    setShowHint(false);
                    setShowSolution(false);
                  }}
                  className={`flex items-center justify-between p-2 rounded cursor-pointer text-xs transition-colors ${
                    isSelected
                      ? 'bg-[#107c41] text-white font-medium shadow-xs'
                      : 'hover:bg-neutral-100 dark:hover:bg-neutral-800 text-neutral-700 dark:text-neutral-300'
                  }`}
                >
                  <div className="flex items-center gap-2 truncate">
                    {isTaskDone ? (
                      <CheckCircle2 size={13} className={isSelected ? 'text-white' : 'text-emerald-500'} />
                    ) : (
                      <span className={`w-3.5 h-3.5 rounded-full border text-[9px] flex items-center justify-center font-mono ${
                        isSelected ? 'border-white text-white' : 'border-neutral-400 text-neutral-400'
                      }`}>
                        {idx + 1}
                      </span>
                    )}
                    <span className="truncate">{t.title}</span>
                  </div>

                  <span className={`text-[10px] shrink-0 font-bold ${
                    isSelected ? 'text-white' : 'text-amber-600 dark:text-amber-400'
                  }`}>
                    +{t.xpReward} XP
                  </span>
                </div>
              );
            })}
          </div>
        </div>
      </div>
    </aside>
  );
};
