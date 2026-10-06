import React, { useState } from 'react';
import { Flame, Bell, CheckCircle2, Zap, Calendar, Sparkles, Heart } from 'lucide-react';
import { sounds } from '../utils/audio';

interface DailyMotivationModalProps {
  isOpen: boolean;
  onClose: () => void;
  streakDays: number;
  dailyGoalTasksTarget: number;
  dailyGoalTasksDone: number;
  onClaimDailyBonus: () => void;
  hasClaimedDailyBonus: boolean;
  isDarkMode: boolean;
}

const MOTIVATIONAL_QUOTES = [
  '„Każda formuła w Excelu przybliża Cię do automatyzacji nudnych zadań w pracy biurowej!”',
  '„Zaawansowany użytkownik Excela oszczędza średnio 4 godziny w tygodniu na powtarzalnych raportach.”',
  '„Połączenie JEŻELI i WYSZUKAJ.PIONOWO to potężny duet w każdym dziale controllingu!”',
  '„Ciągłość nauki to klucz – nawet 5 minut dziennie buduje mistrzostwo analityka!”',
];

export const DailyMotivationModal: React.FC<DailyMotivationModalProps> = ({
  isOpen,
  onClose,
  streakDays,
  dailyGoalTasksTarget,
  dailyGoalTasksDone,
  onClaimDailyBonus,
  hasClaimedDailyBonus,
  isDarkMode,
}) => {
  const [reminderSet, setReminderSet] = useState(false);
  const randomQuote = MOTIVATIONAL_QUOTES[streakDays % MOTIVATIONAL_QUOTES.length];

  if (!isOpen) return null;

  const handleEnableNotifications = async () => {
    if ('Notification' in window) {
      try {
        const permission = await Notification.requestPermission();
        if (permission === 'granted') {
          setReminderSet(true);
          new Notification('Excel Pro 365: Przypomnienie aktywne! 🚀', {
            body: 'Będziemy motywować Cię do utrzymania passy nauki każdego dnia!',
          });
        }
      } catch {
        setReminderSet(true);
      }
    } else {
      setReminderSet(true);
    }
  };

  return (
    <div className="fixed inset-0 z-50 bg-black/60 backdrop-blur-xs flex items-center justify-center p-4">
      <div className={`w-full max-w-md rounded-xl shadow-2xl border flex flex-col overflow-hidden animate-in fade-in zoom-in-95 duration-150 ${
        isDarkMode ? 'bg-[#202020] border-[#383838] text-white' : 'bg-white border-neutral-200 text-neutral-900'
      }`}>
        {/* Fire Banner */}
        <div className="p-6 bg-gradient-to-b from-amber-500 via-orange-500 to-rose-600 text-white text-center relative overflow-hidden">
          <div className="absolute top-2 right-2">
            <button onClick={onClose} className="text-white/80 hover:text-white p-1 text-sm">
              ✕
            </button>
          </div>

          {/* Animated Flame Icon */}
          <div className="relative inline-block mb-2">
            <div className="w-20 h-20 mx-auto rounded-full bg-white/20 flex items-center justify-center animate-pulse">
              <Flame size={48} className="text-amber-200 fill-amber-300 animate-bounce" />
            </div>
            <div className="absolute -bottom-1 left-1/2 -translate-x-1/2 px-2.5 py-0.5 bg-black/40 backdrop-blur-xs rounded-full text-xs font-black tracking-wider text-amber-200">
              PASSA DNI
            </div>
          </div>

          <h2 className="text-2xl font-black">{streakDays} Dni w Ognistej Passie! 🔥</h2>
          <p className="text-xs text-amber-100 max-w-xs mx-auto mt-1">
            Nie przerywaj łańcucha! Każdy kolejny dzień zwiększa Twój mnożnik punktów w ligach.
          </p>
        </div>

        {/* Motivation Quote */}
        <div className={`p-4 border-b text-center italic text-xs leading-relaxed ${
          isDarkMode ? 'bg-[#262626] border-[#383838] text-neutral-300' : 'bg-amber-50/70 border-amber-200 text-amber-900'
        }`}>
          {randomQuote}
        </div>

        {/* Content: Daily Goals & Daily Bonus */}
        <div className="p-6 space-y-4">
          {/* Daily Goal Card */}
          <div className={`p-3.5 rounded-lg border ${
            isDarkMode ? 'bg-[#262626] border-[#383838]' : 'bg-neutral-50 border-neutral-200'
          }`}>
            <div className="flex items-center justify-between text-xs font-bold mb-2">
              <span className="flex items-center gap-1.5">
                <Calendar size={14} className="text-[#107c41]" />
                Dzienne Wyzwanie Nauki
              </span>
              <span className="text-neutral-500 font-mono">
                {dailyGoalTasksDone} / {dailyGoalTasksTarget} zadania
              </span>
            </div>

            <div className="w-full h-2 bg-neutral-200 dark:bg-neutral-700 rounded-full overflow-hidden">
              <div
                className="h-full bg-[#107c41] rounded-full transition-all duration-300"
                style={{ width: `${Math.min(100, (dailyGoalTasksDone / dailyGoalTasksTarget) * 100)}%` }}
              />
            </div>

            <div className="text-[11px] text-neutral-500 mt-2">
              {dailyGoalTasksDone >= dailyGoalTasksTarget
                ? '✅ Świetna robota! Twój dzisiejszy cel został osiągnięty.'
                : 'Rozwiąż jeszcze przynajmniej jedno zadanie w Akademii, aby zabezpieczyć dzisiejszą passę.'}
            </div>
          </div>

          {/* Daily XP Bonus button */}
          {!hasClaimedDailyBonus ? (
            <button
              onClick={onClaimDailyBonus}
              className="w-full py-2.5 px-4 bg-gradient-to-r from-amber-500 to-orange-500 hover:from-amber-400 hover:to-orange-400 text-neutral-900 font-black text-xs rounded-lg shadow-md flex items-center justify-center gap-2 transition-transform active:scale-95"
            >
              <Zap size={16} className="fill-neutral-900" />
              <span>Odbierz Codzienny Bonus: +50 XP!</span>
            </button>
          ) : (
            <div className="w-full py-2 px-4 bg-emerald-500/10 text-emerald-600 dark:text-emerald-400 border border-emerald-500/20 text-xs font-bold rounded-lg text-center flex items-center justify-center gap-1.5">
              <CheckCircle2 size={15} />
              <span>Codzienny bonus XP został już dziś odebrany!</span>
            </div>
          )}

          {/* Browser Notification Switch */}
          <div className="pt-2">
            {!reminderSet ? (
              <button
                onClick={handleEnableNotifications}
                className="w-full py-2 px-3 border border-neutral-300 dark:border-neutral-700 hover:border-neutral-400 rounded-lg text-xs flex items-center justify-center gap-2 text-neutral-700 dark:text-neutral-300 transition-colors"
              >
                <Bell size={14} className="text-amber-500" />
                <span>Włącz codzienne przypomnienie o Excelu</span>
              </button>
            ) : (
              <div className="text-center text-[11px] text-emerald-600 dark:text-emerald-400 flex items-center justify-center gap-1">
                <CheckCircle2 size={13} />
                <span>Przypomnienia o passie są włączone</span>
              </div>
            )}
          </div>
        </div>

        {/* Footer */}
        <div className={`p-4 border-t text-center ${
          isDarkMode ? 'bg-[#262626] border-[#383838]' : 'bg-neutral-50 border-neutral-200'
        }`}>
          <button
            onClick={onClose}
            className="w-full py-2 bg-[#107c41] text-white font-bold text-xs rounded-lg hover:bg-[#0d6434] transition-colors"
          >
            Przejdź do arkusza i działaj
          </button>
        </div>
      </div>
    </div>
  );
};
