import React from 'react';
import { Award, Zap, Flame, Trophy, CheckCircle, Keyboard, Target, Calendar, Sparkles } from 'lucide-react';
import { UserProfile, UserBadge } from '../types/excel';

interface ProfileModalProps {
  isOpen: boolean;
  onClose: () => void;
  profile: UserProfile;
  badges: UserBadge[];
  isDarkMode: boolean;
}

export const ProfileModal: React.FC<ProfileModalProps> = ({
  isOpen,
  onClose,
  profile,
  badges,
  isDarkMode,
}) => {
  if (!isOpen) return null;

  const nextLevelXp = profile.level * 500;
  const currentLevelProgress = Math.min(100, Math.round((profile.xp % 500) / 5));

  return (
    <div className="fixed inset-0 z-50 bg-black/60 backdrop-blur-xs flex items-center justify-center p-4">
      <div className={`w-full max-w-3xl rounded-xl shadow-2xl border flex flex-col overflow-hidden max-h-[90vh] animate-in fade-in zoom-in-95 duration-150 ${
        isDarkMode ? 'bg-[#202020] border-[#383838] text-white' : 'bg-white border-neutral-200 text-neutral-900'
      }`}>
        {/* Profile Header */}
        <div className="p-6 bg-gradient-to-r from-[#107c41] via-[#0d6434] to-[#084222] text-white flex flex-col sm:flex-row items-center sm:items-start justify-between gap-4">
          <div className="flex flex-col sm:flex-row items-center sm:items-start gap-4">
            <div className="w-16 h-16 rounded-2xl bg-white/10 border-2 border-white/30 flex items-center justify-center text-2xl font-black shadow-lg">
              {profile.name.slice(0, 2).toUpperCase()}
            </div>
            <div className="text-center sm:text-left">
              <h2 className="text-xl font-black">{profile.name}</h2>
              <div className="text-xs text-emerald-200 font-medium">{profile.title} · {profile.email}</div>
              <div className="flex items-center gap-2 mt-2">
                <span className="px-2 py-0.5 rounded-full bg-white/20 text-xs font-semibold">
                  Liga {profile.league}
                </span>
                <span className="px-2 py-0.5 rounded-full bg-amber-400 text-neutral-900 text-xs font-bold flex items-center gap-1">
                  <Flame size={12} className="fill-neutral-900" /> {profile.streakDays} dni passy
                </span>
              </div>
            </div>
          </div>

          <button
            onClick={onClose}
            className="text-white/80 hover:text-white p-1 rounded-md text-sm self-start"
          >
            ✕
          </button>
        </div>

        {/* Level & XP Progress Bar */}
        <div className={`px-6 py-4 border-b ${
          isDarkMode ? 'bg-[#262626] border-[#383838]' : 'bg-neutral-50 border-neutral-200'
        }`}>
          <div className="flex items-center justify-between text-xs font-bold mb-1.5">
            <span className="flex items-center gap-1.5 text-[#107c41] dark:text-emerald-400">
              <Zap size={14} className="fill-[#107c41] dark:fill-emerald-400" />
              Poziom {profile.level}: Analityk Danych
            </span>
            <span className="font-mono text-neutral-500">
              {profile.xp} / {nextLevelXp} XP ({currentLevelProgress}%)
            </span>
          </div>
          <div className="w-full h-2.5 bg-neutral-200 dark:bg-neutral-700 rounded-full overflow-hidden">
            <div
              className="h-full bg-gradient-to-r from-[#107c41] to-emerald-400 rounded-full transition-all duration-500"
              style={{ width: `${currentLevelProgress}%` }}
            />
          </div>
        </div>

        {/* Key Stats Counter Grid */}
        <div className="grid grid-cols-2 sm:grid-cols-4 gap-3 p-6 border-b border-neutral-200 dark:border-neutral-800 text-center">
          <div className={`p-3 rounded-lg border ${isDarkMode ? 'bg-[#262626] border-[#383838]' : 'bg-neutral-50 border-neutral-200'}`}>
            <div className="text-xl font-bold font-mono text-[#107c41] dark:text-emerald-400">{profile.completedTaskIds.length}</div>
            <div className="text-[11px] text-neutral-500 mt-0.5">Ukończone lekcje</div>
          </div>
          <div className={`p-3 rounded-lg border ${isDarkMode ? 'bg-[#262626] border-[#383838]' : 'bg-neutral-50 border-neutral-200'}`}>
            <div className="text-xl font-bold font-mono text-amber-500">{profile.streakDays}</div>
            <div className="text-[11px] text-neutral-500 mt-0.5">Dni passy 🔥</div>
          </div>
          <div className={`p-3 rounded-lg border ${isDarkMode ? 'bg-[#262626] border-[#383838]' : 'bg-neutral-50 border-neutral-200'}`}>
            <div className="text-xl font-bold font-mono text-blue-500">{badges.filter(b => b.progress >= 100).length} / {badges.length}</div>
            <div className="text-[11px] text-neutral-500 mt-0.5">Zdobyte odznaki 🎖️</div>
          </div>
          <div className={`p-3 rounded-lg border ${isDarkMode ? 'bg-[#262626] border-[#383838]' : 'bg-neutral-50 border-neutral-200'}`}>
            <div className="text-xl font-bold font-mono text-purple-500">{profile.shortcutsUsedCount}</div>
            <div className="text-[11px] text-neutral-500 mt-0.5">Użyte skróty ⌨️</div>
          </div>
        </div>

        {/* Badges Section */}
        <div className="flex-1 overflow-y-auto px-6 py-4 space-y-4">
          <div>
            <h3 className="text-sm font-bold flex items-center gap-2 mb-1">
              <Award size={16} className="text-amber-500" />
              <span>Odznaki & Osiągnięcia w Excelu</span>
            </h3>
            <p className="text-xs text-neutral-500">
              Wykonuj zadania biurowe, używaj skrótów klawiszowych i utrzymuj passę, aby odblokowywać trofea.
            </p>
          </div>

          <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
            {badges.map((badge) => {
              const isUnlocked = badge.progress >= 100;

              return (
                <div
                  key={badge.id}
                  className={`p-3 rounded-xl border flex items-start gap-3 transition-all ${
                    isUnlocked
                      ? isDarkMode
                        ? 'bg-amber-950/20 border-amber-600/40 text-white'
                        : 'bg-amber-50/70 border-amber-200 text-neutral-900'
                      : isDarkMode
                      ? 'bg-[#262626] border-[#383838] opacity-60'
                      : 'bg-neutral-50 border-neutral-200 opacity-60'
                  }`}
                >
                  <div className="text-2xl p-2 rounded-lg bg-white/20 dark:bg-black/20 shrink-0">
                    {badge.icon}
                  </div>

                  <div className="flex-1 min-w-0">
                    <div className="flex items-center justify-between gap-1">
                      <h4 className="text-xs font-bold truncate">{badge.name}</h4>
                      {isUnlocked ? (
                        <span className="text-[10px] font-bold text-emerald-600 dark:text-emerald-400 flex items-center gap-0.5 shrink-0">
                          <CheckCircle size={11} /> Zdobyta
                        </span>
                      ) : (
                        <span className="text-[10px] font-mono text-neutral-400 shrink-0">
                          {badge.current}/{badge.target}
                        </span>
                      )}
                    </div>

                    <p className="text-[11px] text-neutral-500 dark:text-neutral-400 mt-0.5 line-clamp-2">
                      {badge.description}
                    </p>

                    {/* Progress bar */}
                    {!isUnlocked && (
                      <div className="w-full h-1.5 bg-neutral-200 dark:bg-neutral-700 rounded-full mt-2 overflow-hidden">
                        <div
                          className="h-full bg-amber-500 rounded-full"
                          style={{ width: `${badge.progress}%` }}
                        />
                      </div>
                    )}
                  </div>
                </div>
              );
            })}
          </div>
        </div>

        {/* Footer */}
        <div className={`px-6 py-3 border-t text-right ${
          isDarkMode ? 'bg-[#262626] border-[#383838]' : 'bg-neutral-50 border-neutral-200'
        }`}>
          <button
            onClick={onClose}
            className="px-4 py-1.5 bg-[#107c41] text-white text-xs font-bold rounded-lg hover:bg-[#0d6434] transition-colors"
          >
            Zamknij profil
          </button>
        </div>
      </div>
    </div>
  );
};
