import React, { useState } from 'react';
import { Trophy, Flame, Zap, ArrowUp, ArrowDown, Minus, Shield, Timer, Award } from 'lucide-react';
import { LeaderboardUser } from '../types/excel';

interface LeaderboardModalProps {
  isOpen: boolean;
  onClose: () => void;
  leaderboard: LeaderboardUser[];
  userXp: number;
  userStreak: number;
  onOpenSpeedChallenge: () => void;
  isDarkMode: boolean;
}

const LEAGUES = ['Brązowa', 'Srebrna', 'Złota', 'Diamentowa', 'Mistrzowska'];

export const LeaderboardModal: React.FC<LeaderboardModalProps> = ({
  isOpen,
  onClose,
  leaderboard,
  userXp,
  userStreak,
  onOpenSpeedChallenge,
  isDarkMode,
}) => {
  const [selectedLeague, setSelectedLeague] = useState<string>('Złota');

  if (!isOpen) return null;

  return (
    <div className="fixed inset-0 z-50 bg-black/60 backdrop-blur-xs flex items-center justify-center p-4">
      <div className={`w-full max-w-2xl rounded-xl shadow-2xl border flex flex-col overflow-hidden max-h-[90vh] animate-in fade-in zoom-in-95 duration-150 ${
        isDarkMode ? 'bg-[#202020] border-[#383838] text-white' : 'bg-white border-neutral-200 text-neutral-900'
      }`}>
        {/* Header */}
        <div className="flex items-center justify-between px-6 py-4 bg-gradient-to-r from-[#107c41] to-[#0a522a] text-white">
          <div className="flex items-center gap-3">
            <div className="p-2.5 bg-white/10 rounded-lg">
              <Trophy size={24} className="text-amber-300" />
            </div>
            <div>
              <h2 className="text-lg font-bold">Biurowa Tablica Wyników & Ligi</h2>
              <p className="text-xs text-emerald-100">Rywalizuj z analitykami i kontrolerami finansowymi</p>
            </div>
          </div>
          <button
            onClick={onClose}
            className="text-white/80 hover:text-white p-1 rounded-md text-sm"
          >
            ✕
          </button>
        </div>

        {/* League Selector */}
        <div className={`flex items-center px-6 py-2.5 gap-2 border-b overflow-x-auto ${
          isDarkMode ? 'bg-[#262626] border-[#383838]' : 'bg-neutral-50 border-neutral-200'
        }`}>
          {LEAGUES.map((league) => (
            <button
              key={league}
              onClick={() => setSelectedLeague(league)}
              className={`px-3 py-1.5 rounded-full text-xs font-semibold whitespace-nowrap transition-all ${
                selectedLeague === league
                  ? 'bg-amber-400 text-neutral-900 shadow-xs'
                  : isDarkMode
                  ? 'text-neutral-400 hover:text-white hover:bg-neutral-800'
                  : 'text-neutral-600 hover:text-neutral-900 hover:bg-neutral-200'
              }`}
            >
              Liga {league}
            </button>
          ))}
        </div>

        {/* Speed Challenge Banner */}
        <div className="px-6 py-3 bg-amber-500/10 border-b border-amber-500/20 flex items-center justify-between">
          <div className="flex items-center gap-2">
            <Timer className="text-amber-500 shrink-0" size={18} />
            <div className="text-xs">
              <strong className="text-amber-700 dark:text-amber-300 font-bold">Szybki Pojedynek na Czas!</strong>
              <div className="text-neutral-600 dark:text-neutral-400">Rozwiąż 3 zadania w 60 sekund i zdobądź +150 XP do ligi.</div>
            </div>
          </div>
          <button
            onClick={() => {
              onClose();
              onOpenSpeedChallenge();
            }}
            className="px-3 py-1.5 bg-amber-500 hover:bg-amber-400 text-neutral-900 font-bold text-xs rounded-lg transition-colors shrink-0"
          >
            Graj teraz
          </button>
        </div>

        {/* Leaderboard Table */}
        <div className="flex-1 overflow-y-auto px-6 py-4 space-y-2">
          <div className="text-[11px] font-semibold text-neutral-500 uppercase tracking-wide flex justify-between px-3">
            <span>Pozycja & Użytkownik</span>
            <span>Passa / Punkty XP</span>
          </div>

          {leaderboard.map((user) => {
            const isTop3 = user.rank <= 3;
            const isUser = user.isCurrentUser;

            return (
              <div
                key={user.id}
                className={`flex items-center justify-between p-3 rounded-lg border transition-all ${
                  isUser
                    ? 'bg-[#107c41]/10 border-[#107c41] ring-1 ring-[#107c41]'
                    : isDarkMode
                    ? 'bg-[#262626] border-[#383838]'
                    : 'bg-white border-neutral-200 hover:border-neutral-300'
                }`}
              >
                {/* Left: Rank & Info */}
                <div className="flex items-center gap-3">
                  <div className="w-7 text-center font-bold text-sm font-mono flex items-center justify-center">
                    {user.rank === 1 && <span className="text-xl">🥇</span>}
                    {user.rank === 2 && <span className="text-xl">🥈</span>}
                    {user.rank === 3 && <span className="text-xl">🥉</span>}
                    {user.rank > 3 && <span className="text-neutral-400">#{user.rank}</span>}
                  </div>

                  {/* Trend */}
                  <div className="w-4">
                    {user.trend === 'up' && <ArrowUp size={13} className="text-emerald-500" />}
                    {user.trend === 'down' && <ArrowDown size={13} className="text-rose-500" />}
                    {user.trend === 'same' && <Minus size={13} className="text-neutral-400" />}
                  </div>

                  {/* Avatar */}
                  <div className={`w-8 h-8 rounded-full flex items-center justify-center font-bold text-xs text-white ${
                    isUser ? 'bg-[#107c41]' : isTop3 ? 'bg-amber-600' : 'bg-neutral-600'
                  }`}>
                    {user.avatar}
                  </div>

                  <div>
                    <div className="text-xs font-bold flex items-center gap-1.5">
                      <span>{user.name}</span>
                      {isUser && (
                        <span className="text-[10px] bg-[#107c41] text-white px-1.5 py-0.2 rounded font-medium">
                          Ty
                        </span>
                      )}
                    </div>
                    <div className="text-[11px] text-neutral-500">{user.role}</div>
                  </div>
                </div>

                {/* Right: Streak & XP */}
                <div className="flex items-center gap-4 text-xs">
                  <div className="flex items-center gap-1 text-amber-500 font-semibold">
                    <Flame size={14} className="fill-amber-400" />
                    <span>{isUser ? userStreak : user.streakDays}d</span>
                  </div>

                  <div className="flex items-center gap-1 font-mono font-bold text-sm tabular-numbers min-w-[70px] justify-end">
                    <Zap size={14} className="text-amber-500 fill-amber-400" />
                    <span>{isUser ? userXp : user.xp} XP</span>
                  </div>
                </div>
              </div>
            );
          })}
        </div>

        {/* Footer info */}
        <div className={`px-6 py-3 border-t text-xs flex items-center justify-between ${
          isDarkMode ? 'bg-[#262626] border-[#383838] text-neutral-400' : 'bg-neutral-50 border-neutral-200 text-neutral-600'
        }`}>
          <div className="flex items-center gap-2">
            <Shield size={14} className="text-[#107c41]" />
            <span>Koniec sezonu ligowego za <strong>4 dni 12 godz.</strong></span>
          </div>
          <span className="text-emerald-600 dark:text-emerald-400 font-semibold">Miejsca 1-3 awansują wyżej!</span>
        </div>
      </div>
    </div>
  );
};
