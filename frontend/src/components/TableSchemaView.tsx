import type { TableSchema } from "../types";

interface Props {
  table: TableSchema;
  onCellClick?: (tableIdx: number, row: number, col: number) => void;
}

const CELL_STYLES: Record<string, string> = {
  label: "bg-indigo-900/30 text-indigo-300 font-medium",
  empty: "bg-red-900/20 border-dashed border-red-700/50 text-red-400 cursor-pointer hover:bg-red-900/40",
  prefilled: "bg-slate-800 text-slate-400",
  placeholder: "bg-amber-900/20 border-dashed border-amber-700/50 text-amber-400 cursor-pointer hover:bg-amber-900/40",
  unknown: "bg-slate-800/50 text-slate-500",
};

export default function TableSchemaView({ table, onCellClick }: Props) {
  // 그리드 구성
  const grid: (typeof table.cells[0] | null)[][] = Array.from(
    { length: table.rows },
    () => Array(table.cols).fill(null),
  );
  for (const cell of table.cells) {
    if (cell.row < table.rows && cell.col < table.cols) {
      grid[cell.row][cell.col] = cell;
    }
  }

  return (
    <div className="overflow-auto">
      <table className="border-collapse text-[11px] w-full">
        <tbody>
          {grid.map((row, ri) => (
            <tr key={ri}>
              {row.map((cell, ci) => {
                const type = cell?.cell_type ?? "unknown";
                const clickable = type === "empty" || type === "placeholder";
                return (
                  <td
                    key={ci}
                    className={`border border-slate-700 px-1.5 py-1 max-w-[120px] truncate ${CELL_STYLES[type] ?? CELL_STYLES.unknown}`}
                    onClick={() => clickable && onCellClick?.(table.table_idx, ri, ci)}
                    title={cell?.text || (cell?.needs_fill ? "(빈 셀 - 클릭하여 채우기)" : "")}
                  >
                    {cell?.text || (cell?.needs_fill ? "..." : "")}
                  </td>
                );
              })}
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
}
