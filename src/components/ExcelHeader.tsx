import React from 'react';
import {
  Save,
  Undo2,
  Redo2,
  Flame,
  Zap,
  Moon,
  Sun,
  Trophy,
  User,
  Wifi,
  WifiOff,
  CloudCheck,
  Search,
  Sparkles,
  Keyboard,
  BarChart2,
  BookOpen
} from 'lucide-react';
import { UserProfile } from '../types/excel';

interface ExcelHeaderProps {
  userProfile: UserProfile;
  isDarkMode: boolean;
  onToggleDarkMode: () => void;
  isOnline: boolean;
  lastSavedText: string;
  onSaveManual: () => void;
  canUndo: boolean;
  canRedo: boolean;
  onUndo: () => void;
  onRedo: () => void;
  onOpenProfile: () => void;
  onOpenLeaderboard: () => void;
  onOpenDailyModal: () => void;
  onOpenShortcuts: () => void;
  onOpenChartModal: () => void;
  onToggleLearningPanel: () => void;
  isLearningPanelOpen: boolean;
}

export const ExcelHeader: React.FC<ExcelHeaderProps> = ({
  userProfile,
  isDarkMode,
  onToggleDarkMode,
  isOnline,
  lastSavedText,
  onSaveManual,
  canUndo,
  canRedo,
  onUndo,
  onRedo,
  onOpenProfile,
  onOpenLeaderboard,
  onOpenDailyModal,
  onOpenShortcuts,
  onOpenChartModal,
  onToggleLearningPanel,
  isLearningPanelOpen,
}) => {
  return (
    <header className="flex flex-col bg-[#107c41] text-white select-none shrink-0 shadow-sm transition-colors">
      {/* Top Application Bar */}
      <div className="flex items-center justify-between px-3 py-1.5 h-11 text-xs gap-2 md:gap-4 border-b border-[#0d6434]">
        {/* Left: Office Excel App Brand & Quick Access */}
        <div className="flex items-center gap-2 md:gap-3 shrink-0">
          <div className="flex items-center gap-1.5 font-bold text-sm tracking-tight text-white pr-2 border-r border-[#15944f]">
            <div className="w-6 h-6 bg-white text-[#107c41] rounded flex items-center justify-center font-black text-sm shadow-sm">
              X
            </div>
            <span className="hidden sm:inline font-semibold">Excel</span>
            <span className="text-[11px] font-normal px-1.5 py-0.2 bg-[#094b27] rounded text-emerald-200">
              365
            </span>
          </div>

          {/* Quick Access Toolbar */}
          <div className="flex items-center gap-0.5">
            <button
              onClick={onSaveManual}
              title="Zapisz arkusz (Ctrl+S)"
              className="p-1 hover:bg-[#15944f] active:bg-[#0d6434] rounded transition-colors"
            >
              <Save size={15} />
            </button>
            <button
              onClick={onUndo}
              disabled={!canUndo}
              title="Cofnij (Ctrl+Z)"
              className={`p-1 rounded transition-colors ${canUndo ? 'hover:bg-[#15944f] active:bg-[#0d6434]' : 'opacity-40 cursor-not-allowed'}`}
            >
              <Undo2 size={15} />
            </button>
            <button
              onClick={onRedo}
              disabled={!canRedo}
              title="Ponów (Ctrl+Y)"
              className={`p-1 rounded transition-colors ${canRedo ? 'hover:bg-[#15944f] active:bg-[#0d6434]' : 'opacity-40 cursor-not-allowed'}`}
            >
              <Redo2 size={15} />
            </button>
          </div>

          {/* Cloud Sync State */}
          <div className="hidden lg:flex items-center gap-1.5 text-[11px] text-emerald-100 pl-2">
            {isOnline ? (
              <span className="flex items-center gap-1 text-emerald-200" title="Synchronizacja w czasie rzeczywistym">
                <CloudCheck size={13} className="text-emerald-300" />
                <span>Autozapis włączony</span>
              </span>
            ) : (
              <span className="flex items-center gap-1 text-amber-200 bg-amber-900/40 px-1.5 py-0.5 rounded">
                <WifiOff size={12} />
                <span>Tryb offline (lokalnie)</span>
              </span>
            )}
            <span className="text-emerald-300/70 text-[10px]">· {lastSavedText}</span>
          </div>
        </div>

        {/* Center: Search / Commands bar (Excel 365 Tell Me) */}
        <div className="flex-1 max-w-md hidden md:block">
          <div className="relative flex items-center">
            <Search size={14} className="absolute left-2.5 text-emerald-100/70 pointer-events-none" />
            <input
              type="text"
              placeholder="Wyszukaj funkcje (np. SUMA, JEŻELI, VLOOKUP) lub skróty..."
              onClick={onOpenShortcuts}
              readOnly
              className="w-full bg-[#0d6434]/80 hover:bg-[#0d6434] cursor-pointer text-white placeholder-emerald-100/70 text-xs rounded pl-8 pr-3 py-1 outline-none border border-transparent focus:border-white/40 transition-colors"
            />
          </div>
        </div>

        {/* Right: Gamification Badges, Streak, Mode & Profile */}
        <div className="flex items-center gap-1 md:gap-2">
          {/* Daily Streak Flame Button */}
          <button
            onClick={onOpenDailyModal}
            title="Kliknij, aby sprawdzić passę nauki i codzienne wyzwanie!"
            className="flex items-center gap-1 px-2 py-1 rounded bg-[#094b27] hover:bg-[#06351b] transition-all text-xs font-semibold text-amber-300 border border-amber-400/30 shadow-xs"
          >
            <Flame size={15} className="text-amber-400 animate-pulse fill-amber-400" />
            <span className="tabular-numbers">{userProfile.streakDays} dni</span>
          </button>

          {/* XP & Level Indicator */}
          <button
            onClick={onOpenLeaderboard}
            title="Twoje punkty XP i liga rywalizacji"
            className="hidden sm:flex items-center gap-1 px-2 py-1 rounded bg-[#0d6434] hover:bg-[#128546] transition-colors text-xs font-medium text-emerald-100"
          >
            <Zap size={14} className="text-amber-300 fill-amber-300" />
            <span className="tabular-numbers font-bold">{userProfile.xp} XP</span>
            <span className="text-[10px] text-emerald-200">({userProfile.league})</span>
          </button>

          {/* Academy Learning Toggle Button */}
          <button
            onClick={onToggleLearningPanel}
            title="Przełącz panel zadań Akademii Excela"
            className={`flex items-center gap-1 px-2.5 py-1 rounded font-medium text-xs transition-colors shadow-xs ${
              isLearningPanelOpen
                ? 'bg-amber-400 text-neutral-900 font-bold'
                : 'bg-[#15944f] hover:bg-[#1caa5c] text-white'
            }`}
          >
            <BookOpen size={14} />
            <span className="hidden sm:inline">Akademia</span>
          </button>

          {/* Chart quick modal */}
          <button
            onClick={onOpenChartModal}
            title="Generuj wykres z zaznaczonych danych"
            className="p-1.5 hover:bg-[#15944f] rounded transition-colors"
          >
            <BarChart2 size={16} />
          </button>

          {/* Shortcuts cheat sheet */}
          <button
            onClick={onOpenShortcuts}
            title="Skróty klawiszowe Excela"
            className="hidden sm:flex p-1.5 hover:bg-[#15944f] rounded transition-colors"
          >
            <Keyboard size={16} />
          </button>

          {/* Dark mode toggle */}
          <button
            onClick={onToggleDarkMode}
            title={isDarkMode ? 'Przełącz na motyw jasny' : 'Przełącz na motyw ciemny'}
            className="p-1.5 hover:bg-[#15944f] rounded transition-colors"
          >
            {isDarkMode ? <Sun size={16} className="text-amber-300" /> : <Moon size={16} />}
          </button>

          {/* User Profile & Leaderboard */}
          <button
            onClick={onOpenProfile}
            title="Profil użytkownika i zdobyte odznaki"
            className="flex items-center gap-1.5 pl-1.5 pr-2 py-0.5 rounded hover:bg-[#15944f] transition-colors"
          >
            <div className="w-6 h-6 rounded-full bg-emerald-800 border border-white/40 flex items-center justify-center font-bold text-[10px] text-white">
              {userProfile.name.slice(0, 2).toUpperCase()}
            </div>
            <span className="hidden xl:inline text-xs font-medium truncate max-w-[90px]">
              {userProfile.name}
            </span>
          </button>
        </div>
      </div>
    </header>
  );
};
