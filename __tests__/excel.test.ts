import * as XLSX from 'xlsx';
import { buildWorkbook } from '../lib/excel';

describe('buildWorkbook', () => {
  it('creates a worksheet that preserves column order and cell values', () => {
    const workbook = buildWorkbook({
      columns: ['Item', 'Quantity', 'Notes'],
      rows: [
        ['Apples', '10', 'Organic'],
        ['Bananas', '5', ''],
        ['Grapes', '42', 'Seedless']
      ],
      sheetName: 'Inventory'
    });

    const worksheet = workbook.Sheets.Inventory;
    const extracted = XLSX.utils.sheet_to_json<Record<string, string>>(worksheet, {
      header: 1
    });

    expect(extracted[0]).toEqual(['Item', 'Quantity', 'Notes']);
    expect(extracted.slice(1)).toEqual([
      ['Apples', '10', 'Organic'],
      ['Bananas', '5', ''],
      ['Grapes', '42', 'Seedless']
    ]);
  });

  it('fills missing cells with empty strings to maintain sheet shape', () => {
    const workbook = buildWorkbook({
      columns: ['Name', 'Email'],
      rows: [['Jane Doe'], ['John Doe', 'john@example.com']],
      sheetName: 'Contacts'
    });

    const worksheet = workbook.Sheets.Contacts;
    const extracted = XLSX.utils.sheet_to_json<string[]>(worksheet, { header: 1 });

    expect(extracted).toEqual([
      ['Name', 'Email'],
      ['Jane Doe', ''],
      ['John Doe', 'john@example.com']
    ]);
  });
});
