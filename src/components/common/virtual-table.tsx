import type { ReactNode } from "react";
import { TableVirtuoso } from "react-virtuoso";

export type VirtualColumn<T> = {
  key: string;
  label: string;
  headerClassName?: string;
  cellClassName?: string;
  render: (row: T, index: number) => ReactNode;
};

export function VirtualTable<T>({
  rows,
  columns,
  height = 460,
  getRowKey,
  emptyText = "No data"
}: {
  rows: T[];
  columns: Array<VirtualColumn<T>>;
  height?: number;
  getRowKey: (row: T, index: number) => string;
  emptyText?: string;
}) {
  if (!rows.length) {
    return (
      <div className="rounded-xl border border-white/70 bg-white shadow-sm">
        <div className="p-4 text-center text-sm text-muted-foreground">{emptyText}</div>
      </div>
    );
  }

  return (
    <div className="rounded-xl border border-white/70 bg-white shadow-sm">
      <TableVirtuoso
        style={{ height }}
        data={rows}
        computeItemKey={(index, row) => getRowKey(row, index)}
        fixedHeaderContent={() => (
          <tr>
            {columns.map((col) => (
              <th
                key={col.key}
                className={col.headerClassName ?? "bg-slate-100 px-2 py-2 text-left font-medium text-slate-700"}
              >
                {col.label}
              </th>
            ))}
          </tr>
        )}
        itemContent={(index, row) =>
          columns.map((col) => (
            <td key={col.key} className={col.cellClassName ?? "p-2"}>
              {col.render(row, index)}
            </td>
          ))
        }
        components={{
          Table: (props) => <table {...props} className="w-full border-collapse text-sm" />,
          TableRow: (props) => <tr {...props} className="border-b transition-colors hover:bg-slate-50" />
        }}
      />
    </div>
  );
}
