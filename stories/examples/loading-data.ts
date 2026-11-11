/**
 * The grid stores its cells in a sparse `Map`: the key is `"row:col"`, counted
 * from zero (`"0:0"` is A1), and the value is the cell. Only cells that hold
 * something are in the map.
 *
 * `rawValue` is what the user typed, `displayValue` is what the cell shows and
 * `type` says how to treat it. `setData()` replaces everything in the grid.
 */
import 'y11n-spreadsheet';
import { cellKey } from 'y11n-spreadsheet';
import type { GridData } from 'y11n-spreadsheet';

export default function loadingData() {
  const data: GridData = new Map();
  data.set(cellKey(0, 0), { rawValue: 'Name', displayValue: 'Name', type: 'text' });
  data.set(cellKey(0, 1), { rawValue: 'Score', displayValue: 'Score', type: 'text' });
  data.set(cellKey(1, 0), { rawValue: 'Alice', displayValue: 'Alice', type: 'text' });
  data.set(cellKey(1, 1), { rawValue: '95', displayValue: '95', type: 'number' });
  data.set(cellKey(2, 0), { rawValue: 'Bob', displayValue: 'Bob', type: 'text' });
  data.set(cellKey(2, 1), { rawValue: '82', displayValue: '82', type: 'number' });
  // A formula: what it shows is worked out when the data is loaded
  data.set('3:0', { rawValue: 'Average', displayValue: 'Average', type: 'text' });
  data.set('3:1', { rawValue: '=AVERAGE(B2:B3)', displayValue: '', type: 'text' });

  const sheet = document.createElement('y11n-spreadsheet');
  sheet.rows = 20;
  sheet.cols = 8;
  sheet.setData(data);
  return sheet;
}
