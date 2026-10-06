import React from 'react';
import { Keyboard, X } from 'lucide-react';

interface ShortcutsModalProps {
  isOpen: boolean;
  onClose: () => void;
  isDarkMode: boolean;
}

const SHORTCUTS_DATA = [
  { key: 'Ctrl + B', desc: 'Pogrubienie tekstu w zaznaczonych komórkach' },
  { key: 'Ctrl + I', desc: 'Kursywa tekstu' },
  { key: 'Ctrl + U', desc: 'Podkreślenie tekstu' },
  { key: 'Ctrl + C', desc: 'Kopiowanie zaznaczenia do schowka' },
  { key: 'Ctrl + V', desc: 'Wklejanie ze schowka' },
  { key: 'Ctrl + Z', desc: 'Cofnij ostatnią operację (Undo)' },
  { key: 'Ctrl + Y', desc: 'Ponów operację (Redo)' },
  { key: 'F2', desc: 'Edycja aktywnej komórki w miejscu' },
  { key: 'Enter', desc: 'Zatwierdź i przejdź do komórki poniżej' },
  { key: 'Shift + Enter', desc: 'Zatwierdź i przejdź do komórki powyżej' },
  { key: 'Tab', desc: 'Zatwierdź i przejdź do komórki w prawo' },
  { key: 'Shift + Tab', desc: 'Zatwierdź i przejdź do komórki w lewo' },
  { key: 'Escape', desc: 'Anuluj edycję formuły' },
  { key: 'Delete / Backspace', desc: 'Wyczyść zawartość aktywnej komórki' },
  { key: 'Shift + Strzałki', desc: 'Rozszerz zaznaczenie zakresu komórek' },
];

export const ShortcutsModal: React.FC<ShortcutsModalProps> = ({ isOpen, onClose, isDarkMode }) => {
  if (!isOpen) return null;

  return (
    <div className="fixed inset-0 z-50 bg-black/60 backdrop-blur-xs flex items-center justify-center p-4">
      <div className={`w-full max-w-lg rounded-xl shadow-2xl border flex flex-col overflow-hidden animate-in fade-in zoom-in-95 duration-150 ${
        isDarkMode ? 'bg-[#202020] border-[#383838] text-white' : 'bg-white border-neutral-200 text-neutral-900'
      }`}>
        <div className="flex items-center justify-between px-6 py-4 border-b border-neutral-200 dark:border-neutral-800">
          <div className="flex items-center gap-2">
            <div className="p-2 bg-[#107c41]/10 text-[#107c41] dark:text-emerald-400 rounded-lg">
              <Keyboard size={20} />
            </div>
            <div>
              <h2 className="text-base font-bold">Skróty Klawiszowe Excel 365</h2>
              <div className="text-xs text-neutral-500">Najważniejsze skróty przyspieszające pracę z danymi</div>
            </div>
          </div>
          <button onClick={onClose} className="p-1 text-neutral-400 hover:text-neutral-700 dark:hover:text-white">
            <X size={18} />
          </button>
        </div>

        <div className="p-6 overflow-y-auto max-h-[60vh] space-y-2">
          {SHORTCUTS_DATA.map((item, idx) => (
            <div
              key={idx}
              className={`flex items-center justify-between p-2.5 rounded-lg border text-xs ${
                isDarkMode ? 'bg-[#262626] border-[#383838]' : 'bg-neutral-50 border-neutral-200'
              }`}
            >
              <span className="text-neutral-700 dark:text-neutral-300 font-medium">{item.desc}</span>
              <kbd className="px-2 py-1 bg-white dark:bg-[#1a1a1a] border border-neutral-300 dark:border-neutral-700 rounded shadow-xs font-mono font-bold text-[11px] text-[#107c41] dark:text-emerald-400">
                {item.key}
              </kbd>
            </div>
          ))}
        </div>

        <div className={`px-6 py-3 border-t text-right ${
          isDarkMode ? 'bg-[#262626] border-[#383838]' : 'bg-neutral-50 border-neutral-200'
        }`}>
          <button
            onClick={onClose}
            className="px-4 py-1.5 bg-[#107c41] text-white font-bold text-xs rounded-lg hover:bg-[#0d6434]"
          >
            Zamknij
          </button>
        </div>
      </div>
    </div>
  );
};
