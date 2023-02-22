import _pick from 'lodash/pick';

export default class ExcelHeaderColumnType {
  constructor(data) {
    const q = ExcelHeaderColumnType.sanitizeInput(data);
    this.startCol = q.startCol;
    this.startRow = q.startRow;
    this.endCol = q.endCol;
    this.endRow = q.endRow;
    this.worksheet = q.worksheet;
    this.range = q.range;
  }

  static sanitizeInput(data) {
    const { range, worksheet = {} } = data || {};
    const [startCol, startRow, endCol, endRow] = range
      ? range.match(/[A-Z]+|[0-9]+/g)
      : [];
    return {
      startCol,
      startRow,
      endCol,
      endRow,
      worksheet,
      range,
    };
  }

  getNextColumn(column) {
    if (column === '') return 'A';
    const lastChar = column.slice(-1);
    const rest = column.slice(0, -1);
    if (lastChar === 'Z') {
      return `${this.getNextColumn(rest)}A`;
    }
    return rest + String.fromCharCode(lastChar.charCodeAt(0) + 1);
  }

  getHeaderColumnName() {
    if (!this.range) return [];

    const columns = [];
    let colName = this.startCol;
    while (colName <= this.endCol) {
      columns.push(colName + this.startRow);
      colName = this.getNextColumn(colName);
    }
    return columns;
  }

  getHeaderColumnType() {
    if (!this.range) return [];

    const filterData = _pick(this.worksheet, this.getHeaderColumnName());
    const mappingType = {
      s: 'string',
      n: 'number',
      d: 'date',
      f: 'formula',
    };

    const columnType = Object.values(filterData).map(
      (cell) => mappingType[cell.t]
    );
    return columnType;
  }

}
