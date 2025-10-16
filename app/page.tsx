'use client';

import { useMemo, useState } from 'react';
import { Columns2, Download, Plus, RotateCcw } from 'lucide-react';
import { buildWorkbook, createDefaultRows, downloadWorkbook } from '../lib/excel';

const DEFAULT_COLUMNS = ['Title', 'Owner', 'Status'];
const DEFAULT_ROW_COUNT = 4;

function generateFilename() {
  const timestamp = new Date()
    .toISOString()
    .replace(/[-:]/g, '')
    .replace('T', '_')
    .slice(0, 15);
  return `spreadsheet_${timestamp}.xlsx`;
}

export default function HomePage() {
  const [columns, setColumns] = useState<string[]>(DEFAULT_COLUMNS);
  const [rows, setRows] = useState<string[][]>(
    () => createDefaultRows(DEFAULT_COLUMNS, DEFAULT_ROW_COUNT)
  );

  const filledCellCount = useMemo(
    () => rows.flat().filter((value) => value.trim().length > 0).length,
    [rows]
  );

  const handleCellChange = (rowIndex: number, columnIndex: number, value: string) => {
    setRows((currentRows) =>
      currentRows.map((row, rIdx) =>
        rIdx === rowIndex
          ? row.map((cell, cIdx) => (cIdx === columnIndex ? value : cell))
          : row
      )
    );
  };

  const handleColumnRename = (columnIndex: number, value: string) => {
    setColumns((currentColumns) =>
      currentColumns.map((column, idx) => (idx === columnIndex ? value : column))
    );
  };

  const addRow = () => {
    setRows((currentRows) => [...currentRows, columns.map(() => '')]);
  };

  const addColumn = () => {
    setColumns((currentColumns) => {
      const nextColumn = `Column ${currentColumns.length + 1}`;
      const updatedColumns = [...currentColumns, nextColumn];
      setRows((currentRows) =>
        currentRows.map((row) => [...row, ''])
      );
      return updatedColumns;
    });
  };

  const reset = () => {
    setColumns(DEFAULT_COLUMNS);
    setRows(createDefaultRows(DEFAULT_COLUMNS, DEFAULT_ROW_COUNT));
  };

  const download = () => {
    const workbook = buildWorkbook({ columns, rows, sheetName: 'Data' });
    downloadWorkbook(workbook, generateFilename());
  };

  return (
    <main>
      <div className="card">
        <header>
          <h1>Excel Sheet Builder</h1>
          <p className="description">
            Rapidly capture ideas, plans, and checklists, then export a polished Excel
            workbook in one click. Rename columns, add rows, and populate data directly in
            the grid.
          </p>
        </header>

        <section className="toolbar">
          <button type="button" onClick={addRow}>
            <Plus size={18} /> New row
          </button>
          <button type="button" onClick={addColumn}>
            <Columns2 size={18} /> New column
          </button>
          <button type="button" className="secondary" onClick={reset}>
            <RotateCcw size={18} /> Reset grid
          </button>
          <button type="button" onClick={download}>
            <Download size={18} /> Download Excel
          </button>
          <span className="info">
            {columns.length} columns • {rows.length} rows • {filledCellCount} filled cells
          </span>
        </section>

        <section className="table-container" role="region" aria-live="polite">
          <table>
            <thead>
              <tr>
                {columns.map((column, columnIndex) => (
                  <th key={`header-${columnIndex}`} scope="col">
                    <input
                      value={column}
                      onChange={(event) => handleColumnRename(columnIndex, event.target.value)}
                      aria-label={`Column ${columnIndex + 1} name`}
                    />
                  </th>
                ))}
              </tr>
            </thead>
            <tbody>
              {rows.map((row, rowIndex) => (
                <tr key={`row-${rowIndex}`}>
                  {row.map((value, columnIndex) => (
                    <td key={`cell-${rowIndex}-${columnIndex}`}>
                      <input
                        value={value}
                        onChange={(event) =>
                          handleCellChange(rowIndex, columnIndex, event.target.value)
                        }
                        aria-label={`Row ${rowIndex + 1}, column ${columnIndex + 1}`}
                      />
                    </td>
                  ))}
                </tr>
              ))}
            </tbody>
          </table>
        </section>

        <footer className="footer">
          <strong>Need inspiration?</strong>
          <p>
            Start by renaming the headers to match your use case (project plan, inventory,
            contact list, etc.), populate the rows, then export an Excel workbook ready for
            sharing or further analysis.
          </p>
        </footer>
      </div>
    </main>
  );
}
