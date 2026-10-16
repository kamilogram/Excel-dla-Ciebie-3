import { LearningTask } from '../types/excel';

export const EXCEL_LESSONS: LearningTask[] = [
  // 1. Wprowadzanie tekstu i poruszanie się
  {
    id: 'lesson-1',
    title: 'Wprowadzenie tekstu do tabeli',
    difficulty: 'podstawowy',
    category: 'Wprowadzanie Danych',
    xpReward: 50,
    description: 'W komórce B3 wpisz nazwisko nowego kierownika regionu Północ: "Marek Nowak".',
    instructions: [
      'Kliknij komórkę B3.',
      'Wpisz z klawiatury: Marek Nowak.',
      'Zatwierdź klawiszem Enter lub Tab.',
    ],
    targetSheetId: 'akademia-sheet',
    targetCell: 'B3',
    expectedResultCheck: (val) => String(val).trim().toLowerCase() === 'marek nowak',
    hint: 'Kliknij komórkę B3 i wprowadź tekst: Marek Nowak',
    solutionFormula: 'Marek Nowak',
    explanation: 'W Excelu wprowadzanie danych polega na zaznaczeniu komórki i wpisaniu wartości, a następnie zatwierdzeniu klawiszem Enter lub Tab.',
  },

  // 2. Wprowadzanie daty
  {
    id: 'lesson-2',
    title: 'Wprowadzenie daty raportu',
    difficulty: 'podstawowy',
    category: 'Wprowadzanie Danych',
    xpReward: 50,
    description: 'W komórce E1 wprowadź datę sporządzenia raportu w formacie: 2026-01-15 (lub 15.01.2026).',
    instructions: [
      'Kliknij komórkę E1.',
      'Wpisz datę: 2026-01-15 lub 15.01.2026.',
      'Zatwierdź klawiszem Enter.',
    ],
    targetSheetId: 'akademia-sheet',
    targetCell: 'E1',
    expectedResultCheck: (val) => /2026-01-15|15\.01\.2026/.test(String(val).trim()),
    hint: 'W komórce E1 wpisz wartość 2026-01-15',
    solutionFormula: '2026-01-15',
    explanation: 'Excel automatycznie rozpoznaje formaty dat, takie jak RRRR-MM-DD lub DD.MM.RRRR, i traktuje je jako wartości kalendarzowe.',
  },

  // 3. Formatowanie i stylowanie komórki
  {
    id: 'lesson-3',
    title: 'Wyróżnienie tytułu raportu',
    difficulty: 'podstawowy',
    category: 'Formatowanie & Styl',
    xpReward: 50,
    description: 'W komórce A1 sformatuj nagłówek raportu: nadaj komórce pogrubienie tekstu (skrót Ctrl+B lub ikona B) albo kolor wypełnienia tła.',
    instructions: [
      'Zaznacz komórkę A1.',
      'Na Wstążce kliknij ikonę Pogrubienie (B) lub naciśnij skrót Ctrl+B. Możesz także wybrać kolor wypełnienia.',
    ],
    targetSheetId: 'akademia-sheet',
    targetCell: 'A1',
    expectedResultCheck: (_val, _raw, sheet) => Boolean(sheet.cells['A1']?.style?.bold || sheet.cells['A1']?.style?.bgColor),
    hint: 'Zaznacz komórkę A1 i wciśnij skrót Ctrl+B na klawiaturze.',
    solutionFormula: 'Pogrubienie (Ctrl+B)',
    explanation: 'Formatowanie komórek zwiększa czytelność arkusza kalkulacyjnego i pozwala szybko wyróżnić kluczowe nagłówki.',
  },

  // 4. Przeciąganie serii danych (dni tygodnia)
  {
    id: 'lesson-4',
    title: 'Autouzupełnianie serii dni roboczych',
    difficulty: 'podstawowy',
    category: 'Przeciąganie & Serie',
    xpReward: 60,
    description: 'W komórce E15 znajduje się "Poniedziałek". Chwyć zielony kwadracik w prawym dolnym rogu ramki komórki E15 (uchwyt wypełnienia) i przeciągnij go w dół do komórki E19, aby automatycznie dopełnić kolejne dni tygodnia aż do Piątku.',
    instructions: [
      'Kliknij komórkę E15.',
      'Najedź kursorem na prawy dolny róg zaznaczenia (kursor zmieni się w czarny krzyżyk).',
      'Przytrzymaj lewy przycisk myszy i przeciągnij w dół do komórki E19.',
    ],
    targetSheetId: 'akademia-sheet',
    targetCell: 'E19',
    expectedResultCheck: (_val, _raw, sheet) => {
      const e19 = String(sheet.cells['E19']?.raw || '').trim().toLowerCase();
      const e16 = String(sheet.cells['E16']?.raw || '').trim().toLowerCase();
      return (e19.includes('piątek') || e19.includes('piatek')) && e16.includes('wtorek');
    },
    hint: 'Chwyć uchwyt wypełnienia w prawym dolnym rogu komórki E15 i przeciągnij w dół do E19.',
    solutionFormula: 'Przeciągnięcie serii do E19',
    explanation: 'Inteligentny mechanizm autouzupełniania w Excelu potrafi rozpoznawać logiczne ciągi, takie jak dni tygodnia, miesiące czy ciągi liczbowe.',
  },

  // 5. Działanie arytmetyczne – proste mnożenie
  {
    id: 'lesson-5',
    title: 'Kalkulacja podatku dla regionu Północ',
    difficulty: 'podstawowy',
    category: 'Działania Arytmetyczne',
    xpReward: 60,
    description: 'W komórce D3 oblicz kwotę podatku VAT (23%), mnożąc wartość sprzedaży netto z komórki C3 przez stawkę 0.23.',
    instructions: [
      'Kliknij komórkę D3.',
      'Wpisz znak równości =, wskaż komórkę C3, wpisz znak mnożenia * oraz 0.23.',
      'Naciśnij klawisz Enter.',
    ],
    targetSheetId: 'akademia-sheet',
    targetCell: 'D3',
    expectedFormulaKeywords: ['C3', '*'],
    expectedResultCheck: (val) => Math.round(Number(val)) === 1196,
    hint: 'Wprowadź formułę: =C3*0.23 lub =C3*23%',
    solutionFormula: '=C3*0.23',
    explanation: 'Działania arytmetyczne w arkuszu zawsze rozpoczynamy od znaku równości =. Znak gwiazdki * oznacza mnożenie.',
  },

  // 6. Blokowanie komórki znakiem $ (Adresowanie bezwzględne)
  {
    id: 'lesson-6',
    title: 'Przeliczenie waluty z blokowaniem komórki kursu ($)',
    difficulty: 'podstawowy',
    category: 'Adresowanie Bezwzględne',
    xpReward: 70,
    description: 'W komórce F3 przelicz sprzedaż z komórki C3 na EUR, dzieląc ją przez kurs znajdujący się w stałej komórce I1 (4.30). Zablokuj adres komórki I1 za pomocą znaku $ (np. $I$1), aby po przeciągnięciu formuły w dół adres kursu pozostał stały.',
    instructions: [
      'Kliknij komórkę F3.',
      'Wpisz formułę dzielącą komórkę C3 przez zablokowaną komórkę I1 ze znakami dolara: =C3/$I$1.',
      'Zatwierdź klawiszem Enter (wynik to ok. 1209,30).',
    ],
    targetSheetId: 'akademia-sheet',
    targetCell: 'F3',
    expectedResultCheck: (val, raw) => {
      const num = Math.round(Number(val));
      const hasLock = /\$I\$1|I\$1|\$I1/i.test(raw);
      return (num === 1209 || Math.abs(Number(val) - (5200 / 4.3)) < 1) && hasLock;
    },
    hint: 'Znak dolara $ blokuje adres komórki. Wpisz w F3: =C3/$I$1',
    solutionFormula: '=C3/$I$1',
    explanation: 'Znak dolara $ tworzy adres bezwzględny. Wpisanie $I$1 oznacza, że kolumna I oraz wiersz 1 nie zmienią się podczas kopiowania i przeciągania formuły.',
  },

  // 7. Przeciąganie formuły z blokadą $ w dół kolumny
  {
    id: 'lesson-7',
    title: 'Kopiowanie formuły z blokadą $ w dół kolumny',
    difficulty: 'podstawowy',
    category: 'Adresowanie Bezwzględne',
    xpReward: 70,
    description: 'Przeciągnij komórkę F3 za uchwyt wypełnienia w dół do komórki F7. Zwróć uwagę, jak dzięki blokadzie $I$1 komórka C3 zmieniła się na C4, C5 itd., a kurs waluty $I$1 pozostał nienaruszony!',
    instructions: [
      'Kliknij komórkę F3.',
      'Chwyć uchwyt wypełnienia w prawym dolnym rogu komórki F3.',
      'Przeciągnij w dół aż do komórki F7.',
    ],
    targetSheetId: 'akademia-sheet',
    targetCell: 'F7',
    expectedResultCheck: (_val, _raw, sheet) => {
      const f7 = sheet.cells['F7'];
      return Boolean(f7 && Math.abs(Number(f7.computed) - (2900 / 4.3)) < 1 && /\$I\$1|I\$1|\$I1/i.test(f7.raw));
    },
    hint: 'Przeciągnij uchwyt wypełnienia od komórki F3 w dół do F7.',
    solutionFormula: '=C7/$I$1',
    explanation: 'Przy przeciąganiu formuły adresy względne (jak C3) przesuwają się wraz z ruchem, a adresy zablokowane znakiem $ (jak $I$1) pozostają stałe.',
  },

  // 8. Łączna suma sprzedaży
  {
    id: 'lesson-8',
    title: 'Łączna suma sprzedaży',
    difficulty: 'podstawowy',
    category: 'Formuły Podstawowe',
    xpReward: 60,
    description: 'W komórce C8 oblicz łączną wartość sprzedaży ze wszystkich regionów z tabeli (zakres C3:C7).',
    instructions: [
      'Kliknij komórkę C8.',
      'Wprowadź odpowiednią formułę sumującą wartości z zakresu C3:C7 (możesz użyć przycisku Autosumowanie Σ na Wstążce).',
      'Naciśnij Enter.',
    ],
    targetSheetId: 'akademia-sheet',
    targetCell: 'C8',
    expectedFormulaKeywords: ['SUMA', 'SUM'],
    expectedResultCheck: (val) => Number(val) === 21600,
    hint: 'Wprowadź formułę: =SUMA(C3:C7)',
    solutionFormula: '=SUMA(C3:C7)',
    explanation: 'Funkcja SUMA dodaje wszystkie liczby w zadanym zakresie komórek.',
  },

  // 9. Średnia sprzedaż na region
  {
    id: 'lesson-9',
    title: 'Średnia sprzedaż na region',
    difficulty: 'podstawowy',
    category: 'Formuły Podstawowe',
    xpReward: 60,
    description: 'W komórce C9 wyznacz średnią wielkość sprzedaży przypadającą na jeden region handlowy.',
    instructions: [
      'Kliknij komórkę C9.',
      'Wprowadź formułę obliczającą średnią arytmetyczną z zakresu C3:C7.',
      'Zatwierdź klawiszem Enter.',
    ],
    targetSheetId: 'akademia-sheet',
    targetCell: 'C9',
    expectedFormulaKeywords: ['ŚREDNIA', 'SREDNIA', 'AVERAGE'],
    expectedResultCheck: (val) => Math.round(Number(val)) === 4320,
    hint: 'Wprowadź formułę: =ŚREDNIA(C3:C7)',
    solutionFormula: '=ŚREDNIA(C3:C7)',
    explanation: 'Funkcja ŚREDNIA dzieli sumę liczb przez ich liczbę w podanym zakresie.',
  },

  // 10. Najwyższy wynik w kwartale
  {
    id: 'lesson-10',
    title: 'Najwyższy wynik sprzedaży',
    difficulty: 'podstawowy',
    category: 'Formuły Podstawowe',
    xpReward: 60,
    description: 'W komórce C10 wskaż najwyższą wartość sprzedaży osiągniętą w pierwszym kwartale.',
    instructions: [
      'Wybierz komórkę C10.',
      'Wprowadź formułę odnajdującą największą liczbę w zakresie C3:C7.',
      'Zatwierdź klawiszem Enter.',
    ],
    targetSheetId: 'akademia-sheet',
    targetCell: 'C10',
    expectedFormulaKeywords: ['MAX'],
    expectedResultCheck: (val) => Number(val) === 6200,
    hint: 'Wprowadź formułę: =MAX(C3:C7)',
    solutionFormula: '=MAX(C3:C7)',
    explanation: 'Funkcja MAX wyszukuje największą wartość liczbową w zbiorze danych.',
  },

  // 11. Najniższy wynik w kwartale
  {
    id: 'lesson-11',
    title: 'Najniższy wynik sprzedaży',
    difficulty: 'podstawowy',
    category: 'Formuły Podstawowe',
    xpReward: 60,
    description: 'W komórce C11 wyznacz najniższą kwotę sprzedaży spośród wszystkich regionów.',
    instructions: [
      'Wybierz komórkę C11.',
      'Wprowadź formułę wyznaczającą najmniejszą wartość w zakresie C3:C7.',
      'Zatwierdź klawiszem Enter.',
    ],
    targetSheetId: 'akademia-sheet',
    targetCell: 'C11',
    expectedFormulaKeywords: ['MIN'],
    expectedResultCheck: (val) => Number(val) === 2900,
    hint: 'Wprowadź formułę: =MIN(C3:C7)',
    solutionFormula: '=MIN(C3:C7)',
    explanation: 'Funkcja MIN wyszukuje najmniejszą wartość liczbową w zbiorze komórek.',
  },

  // 12. Liczba celów zrealizowanych
  {
    id: 'lesson-12',
    title: 'Liczba regionów z realizacją celu',
    difficulty: 'średniozaawansowany',
    category: 'Funkcje Logiczne',
    xpReward: 80,
    description: 'W komórce C12 policz, ile regionów zrealizowało plan przekraczając kwotę 4000 zł sprzedaży (zakres C3:C7).',
    instructions: [
      'Kliknij komórkę C12.',
      'Wprowadź formułę zliczającą komórki w zakresie C3:C7 spełniające warunek ">4000".',
      'Zatwierdź klawiszem Enter.',
    ],
    targetSheetId: 'akademia-sheet',
    targetCell: 'C12',
    expectedFormulaKeywords: ['LICZ.JEŻELI', 'COUNTIF'],
    expectedResultCheck: (val) => Number(val) === 3,
    hint: 'Pamiętaj o cudzysłowie wokół kryterium: =LICZ.JEŻELI(C3:C7; ">4000")',
    solutionFormula: '=LICZ.JEŻELI(C3:C7; ">4000")',
    explanation: 'LICZ.JEŻELI zlicza komórki w wyznaczonym zakresie, które spełniają zadany warunek logiczny lub tekstowy.',
  },

  // 13. Premia warunkowa pracownika
  {
    id: 'lesson-13',
    title: 'Przyznanie premii kwartalnej',
    difficulty: 'średniozaawansowany',
    category: 'Funkcje Logiczne',
    xpReward: 90,
    description: 'W komórce E3 wyznacz premię dla regionu Północ: jeżeli sprzedaż w C3 przekracza 5000 zł, przyznaj premię 500 zł, w przeciwnym razie 0 zł.',
    instructions: [
      'Przejdź do komórki E3.',
      'Wprowadź formułę sprawdzającą warunek: jeżeli C3 > 5000, zwróć 500, a w przeciwnym razie 0.',
      'Zatwierdź klawiszem Enter.',
    ],
    targetSheetId: 'akademia-sheet',
    targetCell: 'E3',
    expectedFormulaKeywords: ['JEŻELI', 'IF'],
    expectedResultCheck: (val) => Number(val) === 500,
    hint: 'Użyj średników jako separatorów argumentów: =JEŻELI(C3>5000; 500; 0)',
    solutionFormula: '=JEŻELI(C3>5000; 500; 0)',
    explanation: 'Funkcja JEŻELI sprawdza warunek i zwraca jedną wartość gdy warunek jest prawdziwy, a drugą gdy jest fałszywy.',
  },

  // 14. Pobieranie ceny z cennika
  {
    id: 'lesson-14',
    title: 'Pobieranie ceny produktu z cennika',
    difficulty: 'zaawansowany',
    category: 'Wyszukiwanie & Bazy',
    xpReward: 100,
    description: 'W komórce D16 znajdź cenę katalogową produktu o kodzie z komórki B16 w tabeli cennika (zakres G3:I7).',
    instructions: [
      'Przejdź do komórki D16.',
      'Wprowadź formułę odszukującą kod z B16 w pierwszej kolumnie tabeli G3:I7 i zwracającą wartość z 3. kolumny (cena). Użyj 0 jako ostatniego argumentu dla dokładnego dopasowania.',
      'Zatwierdź klawiszem Enter.',
    ],
    targetSheetId: 'akademia-sheet',
    targetCell: 'D16',
    expectedFormulaKeywords: ['WYSZUKAJ.PIONOWO', 'VLOOKUP'],
    expectedResultCheck: (val) => Number(val) === 350,
    hint: 'Składnia: =WYSZUKAJ.PIONOWO(szukana_wartość; tabela; nr_indeksu_kolumny; 0). Wpisz: =WYSZUKAJ.PIONOWO(B16; G3:I7; 3; 0)',
    solutionFormula: '=WYSZUKAJ.PIONOWO(B16; G3:I7; 3; 0)',
    explanation: 'WYSZUKAJ.PIONOWO przeszukuje pionowo pierwszą kolumnę tabeli i pobiera wartość ze wskazanego numeru kolumny w tym samym wierszu.',
  },

  // 15. Łączenie tekstów
  {
    id: 'lesson-15',
    title: 'Utworzenie pełnego imienia i nazwiska',
    difficulty: 'zaawansowany',
    category: 'Operacje na Tekstach',
    xpReward: 90,
    description: 'W komórce C21 połącz imię z komórki A21 i nazwisko z komórki B21, wstawiając pomiędzy nimi pojedynczą spację.',
    instructions: [
      'Kliknij komórkę C21.',
      'Połącz imię z A21, spację " " oraz nazwisko z B21 (możesz użyć dedykowanej funkcji lub operatora &).',
      'Zatwierdź klawiszem Enter.',
    ],
    targetSheetId: 'akademia-sheet',
    targetCell: 'C21',
    expectedResultCheck: (val) => String(val).trim().toLowerCase() === 'anna kowalska',
    hint: 'Wpisz =A21 & " " & B21 lub =ZŁĄCZ.TEKSTY(A21; " "; B21)',
    solutionFormula: '=A21 & " " & B21',
    explanation: 'Łączenie ciągów tekstowych można wykonać za pomocą funkcji ZŁĄCZ.TEKSTY lub operatora ampersand &.',
  },

  // 16. Suma zamówień dla wybranej kategorii
  {
    id: 'lesson-16',
    title: 'Suma wydatków dla wybranej kategorii',
    difficulty: 'ekspert',
    category: 'Analiza Biznesowa',
    xpReward: 120,
    description: 'W komórce C28 wylicz łączną kwotę zamówień należących wyłącznie do kategorii "Biuro" (kategorie w B25:B27, kwoty w C25:C27).',
    instructions: [
      'Kliknij komórkę C28.',
      'Wprowadź formułę sumującą kwoty z zakresu C25:C27 tylko dla tych wierszy, w których kategoria w B25:B27 to "Biuro".',
      'Zatwierdź klawiszem Enter.',
    ],
    targetSheetId: 'akademia-sheet',
    targetCell: 'C28',
    expectedFormulaKeywords: ['SUMA.JEŻELI', 'SUMIF'],
    expectedResultCheck: (val) => Number(val) === 7200,
    hint: 'Wpisz: =SUMA.JEŻELI(B25:B27; "Biuro"; C25:C27)',
    solutionFormula: '=SUMA.JEŻELI(B25:B27; "Biuro"; C25:C27)',
    explanation: 'SUMA.JEŻELI sumuje komórki spełniające określone kryterium, filtrując wiersze w czasie rzeczywistym.',
  },
];
