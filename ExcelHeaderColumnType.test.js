import { describe, expect } from '@jest/globals';
import ExcelHeaderColumnType from './ExcelHeaderColumnType';

describe('ExcelHeaderColumnType', () => {
  const range = 'A1:D10';
  const worksheet = {
    A1: { t: 'n', v: 78 },
    B1: { t: 'n', v: 680 },
    C1: { t: 's', v: 'Mark', r: '<t>Mark</t>', h: 'Mark' },
    D1: { t: 's', v: 'Steele', r: '<t>Steele</t>', h: 'Steele' },
    A2: { t: 'n', v: 265 },
    B2: { t: 'n', v: 279 },
  };
  const columnType = new ExcelHeaderColumnType({ range, worksheet });
  it('should return header column name', () => {
    expect(columnType.getHeaderColumnName()).toEqual(['A1', 'B1', 'C1', 'D1']);
  });

  describe('getHeaderColumnName', () => {
    it('should return header column type', () => {
      expect(columnType.getHeaderColumnType()).toEqual([
        'number',
        'number',
        'string',
        'string',
      ]);
    });
    it('should return an array of column names from startCol to endCol', () => {
      columnType.startCol = 'C';
      columnType.endCol = 'F';
      columnType.startRow = 1;
      const expected = ['C1', 'D1', 'E1', 'F1'];
      expect(columnType.getHeaderColumnName()).toEqual(expected);
    });
  });

  describe('getNextColumn', () => {
    it('should return A if column is empty', () => {
      expect(columnType.getNextColumn('')).toBe('A');
    });

    it('should return B if column is A', () => {
      expect(columnType.getNextColumn('A')).toBe('B');
    });

    it('should return AA if column is Z', () => {
      expect(columnType.getNextColumn('Z')).toBe('AA');
    });

    it('should return AB if column is AA', () => {
      expect(columnType.getNextColumn('AA')).toBe('AB');
    });

    it('should return AZ if column is AY', () => {
      expect(columnType.getNextColumn('AY')).toBe('AZ');
    });
  });
});
