import React, { useState, useEffect } from 'react';
import { Timer, CheckCircle, XCircle, Zap, Trophy, Play, RotateCcw } from 'lucide-react';
import { sounds } from '../utils/audio';

interface SpeedChallengeModalProps {
  isOpen: boolean;
  onClose: () => void;
  onCompleteChallenge: (earnedXp: number) => void;
  isDarkMode: boolean;
}

type Question = {
  id: number;
  question: string;
  context: string;
  expectedAnswers: string[];
  placeholder: string;
};

const CHALLENGE_QUESTIONS: Question[] = [
  {
    id: 1,
    question: 'Jaką formułą podliczysz sumę komórek z zakresu B2 do B10 w języku polskim?',
    context: 'Kolumna B zawiera kwoty faktur.',
    expectedAnswers: ['=SUMA(B2:B10)', 'SUMA(B2:B10)', '=SUM(B2:B10)', 'SUM(B2:B10)'],
    placeholder: 'Wpisz formułę, np. =SUMA(...)',
  },
  {
    id: 2,
    question: 'Zapisz funkcję logiczną: jeśli wartość A1 > 100 wypisz "OK", w przeciwnym razie "NIE".',
    context: 'Sprawdzenie limitu zamówienia.',
    expectedAnswers: [
      '=JEŻELI(A1>100; "OK"; "NIE")',
      '=JEŻELI(A1>100;"OK";"NIE")',
      '=IF(A1>100, "OK", "NIE")',
      '=IF(A1>100,"OK","NIE")',
    ],
    placeholder: '=JEŻELI(A1>100; "OK"; "NIE")',
  },
  {
    id: 3,
    question: 'Jaką funkcją znajdziesz najwyższą liczbę w zakresie C1:C20?',
    context: 'Poszukiwanie maksymalnej sprzedaży.',
    expectedAnswers: ['=MAX(C1:C20)', 'MAX(C1:C20)', '=MAX(C1:C20);'],
    placeholder: '=MAX(...)',
  },
];

export const SpeedChallengeModal: React.FC<SpeedChallengeModalProps> = ({
  isOpen,
  onClose,
  onCompleteChallenge,
  isDarkMode,
}) => {
  const [isPlaying, setIsPlaying] = useState(false);
  const [timeLeft, setTimeLeft] = useState(60);
  const [currentQuestionIdx, setCurrentQuestionIdx] = useState(0);
  const [userAnswer, setUserAnswer] = useState('');
  const [correctCount, setCorrectCount] = useState(0);
  const [isFinished, setIsFinished] = useState(false);
  const [feedback, setFeedback] = useState<'correct' | 'wrong' | null>(null);

  useEffect(() => {
    let timer: any;
    if (isPlaying && timeLeft > 0 && !isFinished) {
      timer = setInterval(() => {
        setTimeLeft((prev) => {
          if (prev <= 1) {
            setIsFinished(true);
            setIsPlaying(false);
            return 0;
          }
          return prev - 1;
        });
      }, 1000);
    }
    return () => clearInterval(timer);
  }, [isPlaying, timeLeft, isFinished]);

  if (!isOpen) return null;

  const handleStart = () => {
    setIsPlaying(true);
    setTimeLeft(60);
    setCurrentQuestionIdx(0);
    setUserAnswer('');
    setCorrectCount(0);
    setIsFinished(false);
    setFeedback(null);
  };

  const handleAnswerSubmit = (e: React.FormEvent) => {
    e.preventDefault();
    if (!userAnswer.trim()) return;

    const currentQ = CHALLENGE_QUESTIONS[currentQuestionIdx];
    const cleanUser = userAnswer.trim().replace(/\s+/g, ' ').replace(/"/g, '"');
    const isCorrect = currentQ.expectedAnswers.some((ans) => {
      const cleanAns = ans.replace(/\s+/g, ' ');
      return cleanUser.toLowerCase() === cleanAns.toLowerCase();
    });

    if (isCorrect) {
      sounds.playSuccess();
      setCorrectCount((prev) => prev + 1);
      setFeedback('correct');
    } else {
      setFeedback('wrong');
    }

    setTimeout(() => {
      setFeedback(null);
      setUserAnswer('');
      if (currentQuestionIdx < CHALLENGE_QUESTIONS.length - 1) {
        setCurrentQuestionIdx((prev) => prev + 1);
      } else {
        setIsFinished(true);
        setIsPlaying(false);
        const earned = (correctCount + (isCorrect ? 1 : 0)) * 50;
        onCompleteChallenge(earned);
      }
    }, 600);
  };

  return (
    <div className="fixed inset-0 z-50 bg-black/60 backdrop-blur-xs flex items-center justify-center p-4">
      <div className={`w-full max-w-lg rounded-xl shadow-2xl border flex flex-col overflow-hidden animate-in fade-in zoom-in-95 duration-150 ${
        isDarkMode ? 'bg-[#202020] border-[#383838] text-white' : 'bg-white border-neutral-200 text-neutral-900'
      }`}>
        {/* Header */}
        <div className="p-5 bg-gradient-to-r from-amber-500 to-orange-600 text-neutral-900 flex items-center justify-between">
          <div className="flex items-center gap-2.5">
            <div className="p-2 bg-neutral-900/10 rounded-lg">
              <Timer size={22} className="text-neutral-900" />
            </div>
            <div>
              <h2 className="text-base font-black">Szybki Pojedynek na Czas (60s)</h2>
              <div className="text-xs font-medium text-neutral-800">
                Pokaż biegłość w formułach i zgarnij punkty do ligi!
              </div>
            </div>
          </div>
          <button onClick={onClose} className="p-1 rounded text-neutral-900/80 hover:text-neutral-900">
            ✕
          </button>
        </div>

        {/* Content */}
        <div className="p-6">
          {!isPlaying && !isFinished && (
            <div className="text-center space-y-4 py-4">
              <div className="text-4xl">⏱️</div>
              <h3 className="text-base font-bold">Zasady pojedynku:</h3>
              <p className="text-xs text-neutral-500 max-w-sm mx-auto">
                Masz dokładnie 60 sekund na wpisanie 3 poprawnych formuł Excela. Za każdą poprawną odpowiedź zyskujesz <strong>+50 XP</strong> do rankingu!
              </p>
              <button
                onClick={handleStart}
                className="px-6 py-2.5 bg-[#107c41] text-white font-bold text-xs rounded-lg hover:bg-[#0d6434] shadow-md flex items-center gap-2 mx-auto"
              >
                <Play size={16} />
                <span>Rozpocznij wyzwanie (Start)</span>
              </button>
            </div>
          )}

          {isPlaying && !isFinished && (
            <div className="space-y-4">
              {/* Timer Bar */}
              <div className="flex items-center justify-between text-xs font-bold">
                <span className="flex items-center gap-1 text-amber-500">
                  <Timer size={15} />
                  <span>Czas: {timeLeft}s</span>
                </span>
                <span className="text-neutral-500">
                  Pytanie {currentQuestionIdx + 1} z {CHALLENGE_QUESTIONS.length}
                </span>
              </div>

              <div className="w-full h-2 bg-neutral-200 dark:bg-neutral-700 rounded-full overflow-hidden">
                <div
                  className={`h-full transition-all duration-1000 ${
                    timeLeft < 15 ? 'bg-rose-500' : 'bg-amber-500'
                  }`}
                  style={{ width: `${(timeLeft / 60) * 100}%` }}
                />
              </div>

              {/* Current Question */}
              <div className={`p-4 rounded-xl border ${
                feedback === 'correct'
                  ? 'bg-emerald-500/10 border-emerald-500'
                  : feedback === 'wrong'
                  ? 'bg-rose-500/10 border-rose-500'
                  : isDarkMode
                  ? 'bg-[#262626] border-[#383838]'
                  : 'bg-neutral-50 border-neutral-200'
              }`}>
                <div className="text-[11px] text-neutral-500 uppercase font-semibold">
                  {CHALLENGE_QUESTIONS[currentQuestionIdx].context}
                </div>
                <h4 className="text-sm font-bold mt-1 text-neutral-900 dark:text-white">
                  {CHALLENGE_QUESTIONS[currentQuestionIdx].question}
                </h4>
              </div>

              {/* Form Input */}
              <form onSubmit={handleAnswerSubmit} className="space-y-3">
                <input
                  type="text"
                  value={userAnswer}
                  onChange={(e) => setUserAnswer(e.target.value)}
                  placeholder={CHALLENGE_QUESTIONS[currentQuestionIdx].placeholder}
                  autoFocus
                  className={`w-full p-2.5 text-xs font-mono rounded-lg border outline-none ${
                    isDarkMode
                      ? 'bg-[#1e1e1e] border-neutral-700 text-white focus:border-[#107c41]'
                      : 'bg-white border-neutral-300 text-black focus:border-[#107c41]'
                  }`}
                />
                <button
                  type="submit"
                  className="w-full py-2 bg-[#107c41] text-white font-bold text-xs rounded-lg hover:bg-[#0d6434] transition-colors"
                >
                  Zatwierdź odpowiedź (Enter)
                </button>
              </form>
            </div>
          )}

          {isFinished && (
            <div className="text-center space-y-4 py-4">
              <div className="w-16 h-16 rounded-full bg-amber-500/10 text-amber-500 mx-auto flex items-center justify-center">
                <Trophy size={32} />
              </div>
              <h3 className="text-lg font-bold">Pojedynek zakończony!</h3>
              <p className="text-xs text-neutral-500">
                Poprawne odpowiedzi: <strong>{correctCount} z {CHALLENGE_QUESTIONS.length}</strong>
              </p>
              <div className="p-3 bg-emerald-500/10 border border-emerald-500/20 rounded-lg text-emerald-600 dark:text-emerald-400 font-bold text-sm">
                +{correctCount * 50} XP dodane do Twojego konta!
              </div>
              <div className="flex gap-2 justify-center">
                <button
                  onClick={handleStart}
                  className="px-4 py-2 border border-neutral-300 dark:border-neutral-700 rounded-lg text-xs font-semibold flex items-center gap-1.5"
                >
                  <RotateCcw size={14} />
                  <span>Zagraj ponownie</span>
                </button>
                <button
                  onClick={onClose}
                  className="px-4 py-2 bg-[#107c41] text-white rounded-lg text-xs font-bold"
                >
                  Wróć do arkusza
                </button>
              </div>
            </div>
          )}
        </div>
      </div>
    </div>
  );
};
