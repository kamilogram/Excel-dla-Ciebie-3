export function extractFormulaCellReferences(formula: string): string[] {
  const clean = formula.trim();
  if (!clean.startsWith('=')) return [];

  const expr = clean.slice(1);
  const refs: string[] = [];
  const refPattern = /(?:\$?[A-Z]+\$?[0-9]+(?:\:\$?[A-Z]+\$?[0-9]+)?)/gi;

  for (const match of expr.matchAll(refPattern)) {
    const normalized = match[0].replace(/\$/g, '').toUpperCase();
    refs.push(normalized);
  }

  return refs.filter((ref, index, arr) => arr.indexOf(ref) === index);
}

export function getFormulaReferenceBounds(formula: string): { minCol: number; maxCol: number; minRow: number; maxRow: number } | null {
  const refs = extractFormulaCellReferences(formula);
  if (refs.length === 0) return null;

  const cells: string[] = [];
  for (const ref of refs) {
    if (ref.includes(':')) {
      cells.push(...ref.split(':'));
    } else {
      cells.push(ref);
    }
  }

  const coords = cells
    .map((cell) => {
      const parsed = /^(\$?[A-Z]+)\$?([0-9]+)$/i.exec(cell);
      if (!parsed) return null;
      return {
        col: parsed[1].replace(/\$/g, '').toUpperCase(),
        row: parseInt(parsed[2], 10),
      };
    })
    .filter((coord): coord is { col: string; row: number } => !!coord);

  if (coords.length === 0) return null;

  const colIndexes = coords.map((coord) => {
    let index = 0;
    for (let i = 0; i < coord.col.length; i++) {
      index = index * 26 + (coord.col.charCodeAt(i) - 64);
    }
    return index - 1;
  });

  const minCol = Math.min(...colIndexes);
  const maxCol = Math.max(...colIndexes);
  const minRow = Math.min(...coords.map((coord) => coord.row));
  const maxRow = Math.max(...coords.map((coord) => coord.row));

  return { minCol, maxCol, minRow, maxRow };
}
