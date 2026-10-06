import React, { useState, useEffect, useRef } from 'react';
import { Check, X, FunctionSquare } from 'lucide-react';

interface FormulaBarProps {
  activeCellId: string;
  selectionRangeText: string;
  currentValue: string;
  onValueChange: (val: string) => void;
  onSubmit: () => void;
  onCancel: () => void;
  onJumpToCell: (cellId: string) => void;
  onOpenFunctionsMenu: () => void;
  isDarkMode: boolean;
}

const COMMON_FUNCTIONS = [
  { name: 'SUMA', syntax: '=SUMA(liczba1; [liczba2]; ...)', desc: 'Dodaje wszystkie liczby w zakresie komórek' },
  { name: 'ŚREDNIA', syntax: '=ŚREDNIA(liczba1; [liczba2]; ...)', desc: 'Zwraca średnią arytmetyczną argumentów' },
  { name: 'JEŻELI', syntax: '=JEŻELI(warunek; wartość_gdy_prawda; wartość_gdy_fałsz)', desc: 'Sprawdza czy warunek jest spełniony' },
  { name: 'WYSZUKAJ.PIONOWO', syntax: '=WYSZUKAJ.PIONOWO(szukana; tabela; kolumna; 0)', desc: 'Wyszukuje wartość w pierwszej kolumnie tabeli' },
  { name: 'LICZ.JEŻELI', syntax: '=LICZ.JEŻELI(zakres; kryteria)', desc: 'Zlicza komórki spełniające podany warunek' },
  { name: 'SUMA.JEŻELI', syntax: '=SUMA.JEŻELI(zakres; kryteria; [zakres_sumy])', desc: 'Dodaje komórki określone podanym kryterium' },
  { name: 'MAX', syntax: '=MAX(liczba1; [liczba2]; ...)', desc: 'Zwraca największą wartość' },
  { name: 'MIN', syntax: '=MIN(liczba1; [liczba2]; ...)', desc: 'Zwraca najmniejszą wartość' },
  { name: 'ZŁĄCZ.TEKSTY', syntax: '=ZŁĄCZ.TEKSTY(tekst1; [tekst2]; ...)', desc: 'Łączy kilka ciągów tekstowych w jeden' },
  { name: 'ZAOKR', syntax: '=ZAOKR(liczba; cyfry)', desc: 'Zaokrągla liczbę do określonej liczby cyfr' },
];

export const FormulaBar: React.FC<FormulaBarProps> = ({
  activeCellId,
  selectionRangeText,
  currentValue,
  onValueChange,
  onSubmit,
  onCancel,
  onJumpToCell,
  onOpenFunctionsMenu,
  isDarkMode,
}) => {
  const [addressInput, setAddressInput] = useState(activeCellId);
  const [showSuggestions, setShowSuggestions] = useState(false);
  const [activeSuggestionIndex, setActiveSuggestionIndex] = useState(0);
  const inputRef = useRef<HTMLInputElement>(null);

  useEffect(() => {
    setAddressInput(selectionRangeText || activeCellId);
  }, [activeCellId, selectionRangeText]);

  // Determine formula hints
  const isTypingFormula = currentValue.startsWith('=');
  const formulaKeyword = isTypingFormula ? currentValue.slice(1).toUpperCase() : '';
  const matchedFunctions = isTypingFormula && formulaKeyword.length > 0 && !formulaKeyword.includes('(')
    ? COMMON_FUNCTIONS.filter((fn) => fn.name.startsWith(formulaKeyword))
    : [];

  const handleKeyDown = (e: React.KeyboardEvent<HTMLInputElement>) => {
    if (e.key === 'Enter') {
      if (showSuggestions && matchedFunctions[activeSuggestionIndex]) {
        e.preventDefault();
        onValueChange(`=${matchedFunctions[activeSuggestionIndex].name}(`);
        setShowSuggestions(false);
      } else {
        onSubmit();
      }
    } else if (e.key === 'Escape') {
      setShowSuggestions(false);
      onCancel();
    } else if (e.key === 'ArrowDown' && showSuggestions && matchedFunctions.length > 0) {
      e.preventDefault();
      setActiveSuggestionIndex((prev) => (prev + 1) % matchedFunctions.length);
    } else if (e.key === 'ArrowUp' && showSuggestions && matchedFunctions.length > 0) {
      e.preventDefault();
      setActiveSuggestionIndex((prev) => (prev - 1 + matchedFunctions.length) % matchedFunctions.length);
    } else if (e.key === 'Tab' && showSuggestions && matchedFunctions[activeSuggestionIndex]) {
      e.preventDefault();
      onValueChange(`=${matchedFunctions[activeSuggestionIndex].name}(`);
      setShowSuggestions(false);
    }
  };

  const handleAddressSubmit = (e: React.KeyboardEvent<HTMLInputElement>) => {
    if (e.key === 'Enter') {
      const clean = addressInput.trim().toUpperCase();
      if (/^[A-Z]+[0-9]+$/.test(clean)) {
        onJumpToCell(clean);
      }
    }
  };

  return (
    <div className={`relative flex items-center px-2 py-1 gap-1 border-b text-xs select-none transition-colors ${
      isDarkMode ? 'bg-[#252525] border-[#383838] text-neutral-200' : 'bg-white border-[#e5e7eb] text-neutral-800'
    }`}>
      {/* Name Box (Pole Nazwy) */}
      <div className="shrink-0 w-20 md:w-24">
        <input
          type="text"
          value={addressInput}
          onChange={(e) => setAddressInput(e.target.value)}
          onKeyDown={handleAddressSubmit}
          title="Pole nazwy - naciśnij Enter, aby przejść do komórki"
          className={`w-full px-2 py-1 text-center font-semibold rounded border outline-none font-mono text-[11px] ${
            isDarkMode
              ? 'bg-[#1e1e1e] border-[#404040] text-white focus:border-[#107c41]'
              : 'bg-white border-neutral-300 text-neutral-900 focus:border-[#107c41]'
          }`}
        />
      </div>

      {/* Formula Buttons: Cancel, Confirm, Insert Function */}
      <div className="flex items-center gap-0.5 shrink-0 px-1 border-r border-neutral-300 dark:border-neutral-700">
        <button
          onClick={onCancel}
          title="Anuluj (Esc)"
          className="p-1 hover:bg-rose-100 dark:hover:bg-rose-950/40 text-neutral-500 hover:text-rose-600 rounded transition-colors"
        >
          <X size={13} />
        </button>
        <button
          onClick={onSubmit}
          title="Zatwierdź (Enter)"
          className="p-1 hover:bg-emerald-100 dark:hover:bg-emerald-950/40 text-neutral-500 hover:text-[#107c41] rounded transition-colors"
        >
          <Check size={13} />
        </button>
        <button
          onClick={onOpenFunctionsMenu}
          title="Wstaw funkcję (fx)"
          className="px-1.5 py-1 hover:bg-neutral-100 dark:hover:bg-neutral-700 text-[#107c41] dark:text-emerald-400 font-serif italic font-bold rounded transition-colors"
        >
          fx
        </button>
      </div>

      {/* Formula Input Box */}
      <div className="flex-1 relative flex items-center">
        <input
          ref={inputRef}
          type="text"
          value={currentValue}
          onChange={(e) => {
            onValueChange(e.target.value);
            setShowSuggestions(e.target.value.startsWith('='));
          }}
          onKeyDown={handleKeyDown}
          placeholder="Wpisz wartość lub formułę (np. =SUMA(C3:C7), =JEŻELI(C3>5000; 500; 0))..."
          className={`w-full px-2.5 py-1 text-xs outline-none rounded font-mono ${
            isDarkMode
              ? 'bg-[#1e1e1e] text-neutral-100 placeholder-neutral-500 focus:ring-1 focus:ring-[#107c41]'
              : 'bg-white text-neutral-900 placeholder-neutral-400 focus:ring-1 focus:ring-[#107c41]'
          }`}
        />

        {/* Function Autocomplete Suggestions Popup */}
        {showSuggestions && matchedFunctions.length > 0 && (
          <div className="absolute top-8 left-0 z-50 w-80 bg-white dark:bg-[#333333] shadow-xl rounded-md border border-neutral-200 dark:border-neutral-700 overflow-hidden text-xs">
            <div className="px-2 py-1 bg-neutral-100 dark:bg-neutral-800 text-[10px] font-semibold text-neutral-500 uppercase tracking-wide">
              Dostępne funkcje Excela
            </div>
            {matchedFunctions.map((fn, idx) => (
              <div
                key={fn.name}
                onClick={() => {
                  onValueChange(`=${fn.name}(`);
                  setShowSuggestions(false);
                  inputRef.current?.focus();
                }}
                className={`p-2 cursor-pointer transition-colors ${
                  idx === activeSuggestionIndex
                    ? 'bg-[#107c41]/10 text-[#107c41] dark:text-emerald-400 font-medium'
                    : 'hover:bg-neutral-50 dark:hover:bg-neutral-700/50'
                }`}
              >
                <div className="font-mono font-bold text-xs">{fn.name}</div>
                <div className="text-[11px] text-neutral-500 dark:text-neutral-400 font-mono truncate">{fn.syntax}</div>
                <div className="text-[10px] text-neutral-400 dark:text-neutral-500">{fn.desc}</div>
              </div>
            ))}
          </div>
        )}
      </div>
    </div>
  );
};
