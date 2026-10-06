export type CellStyle = {
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  align?: 'left' | 'center' | 'right';
  fontSize?: number;
  textColor?: string;
  bgColor?: string;
  format?: 'general' | 'currency' | 'percent' | 'number' | 'date';
  decimals?: number;
  border?: 'none' | 'all' | 'outside' | 'thick-bottom';
};

export type CellData = {
  raw: string; // The formula or typed string e.g. "=SUMA(A1:A5)" or "450"
  computed?: string | number | boolean | null;
  error?: string;
  style?: CellStyle;
};

export type SheetData = {
  id: string;
  name: string;
  cells: Record<string, CellData>; // key: "A1", "B2", etc.
  colWidths: Record<string, number>; // key: "A", "B", etc.
  rowHeights: Record<number, number>; // key: 1, 2, etc.
};

export type SelectionRange = {
  startCol: string;
  startRow: number;
  endCol: string;
  endRow: number;
};

export type TaskDifficulty = 'podstawowy' | 'średniozaawansowany' | 'zaawansowany' | 'ekspert';

export type LearningTask = {
  id: string;
  title: string;
  difficulty: TaskDifficulty;
  category: 'Wprowadzanie & Styl' | 'Formuły Podstawowe' | 'Funkcje Logiczne' | 'Wyszukiwanie & Bazy' | 'Analiza Biznesowa';
  xpReward: number;
  description: string;
  instructions: string[];
  targetSheetId: string;
  targetCell: string; // e.g. "D8" or range "D8:D12"
  expectedFormulaKeywords?: string[]; // e.g. ["SUMA", "SUM"] or ["JEŻELI", "IF"]
  expectedResultCheck: (val: any, raw: string, sheet: SheetData) => boolean;
  hint: string;
  solutionFormula: string;
  explanation: string;
};

export type UserBadge = {
  id: string;
  name: string;
  icon: string;
  description: string;
  category: 'nauka' | 'passa' | 'rywalizacja' | 'skróty';
  unlockedAt?: string;
  xpBonus: number;
  progress: number; // 0 to 100
  target: number;
  current: number;
};

export type LeaderboardUser = {
  id: string;
  name: string;
  avatar: string;
  role: string;
  league: 'Brązowa' | 'Srebrna' | 'Złota' | 'Diamentowa' | 'Mistrzowska';
  xp: number;
  streakDays: number;
  rank: number;
  trend: 'up' | 'down' | 'same';
  isCurrentUser?: boolean;
};

export type UserProfile = {
  name: string;
  email: string;
  title: string;
  xp: number;
  level: number;
  streakDays: number;
  lastActiveDate: string;
  completedTaskIds: string[];
  unlockedBadgeIds: string[];
  dailyGoalCompleted: boolean;
  dailyGoalTasksTarget: number;
  dailyGoalTasksDone: number;
  league: 'Brązowa' | 'Srebrna' | 'Złota' | 'Diamentowa' | 'Mistrzowska';
  shortcutsUsedCount: number;
};
