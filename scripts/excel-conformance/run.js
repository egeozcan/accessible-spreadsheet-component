ObjC.import('Foundation');
const read = p => ObjC.unwrap($.NSString.stringWithContentsOfFileEncodingError(p, $.NSUTF8StringEncoding, null));
const cases = JSON.parse(read('cases.json'));
const xl = Application('Microsoft Excel');
const ws = xl.activeSheet;
const R = a => ws.ranges[a];
const DEF = {A1:1,A2:2,A3:3,A4:4,A5:5,B1:'a',B2:'b',B3:'c',C1:10,C2:20,C3:30,D1:'=1/0',E1:'apple',E2:'Banana',E3:'cherry'};
const put = (a, v) => { if (v === null) { R(a).clearContents(); } else if (typeof v === 'string' && v.startsWith('=')) R(a).formula = v; else if (typeof v === 'string' && v.trim() !== '' && !isNaN(Number(v))) R(a).value = Number(v); else R(a).value = v; };
const NEWER = new Set('IFS SWITCH MAXIFS MINIFS TEXTJOIN CONCAT XLOOKUP XMATCH SEQUENCE UNIQUE COT COTH ACOT ACOTH CSC CSCH SEC SECH GAMMA GAUSS PHI DAYS ISOWEEKNUM NUMBERVALUE UNICHAR UNICODE BITAND BITOR BITXOR BITLSHIFT BITRSHIFT IFNA XOR RRI PDURATION BASE DECIMAL ARABIC ISFORMULA SHEET SHEETS FORMULATEXT RANDARRAY TEXTBEFORE TEXTAFTER TEXTSPLIT VSTACK HSTACK TAKE DROP EXPAND TOCOL TOROW WRAPROWS WRAPCOLS CHOOSECOLS CHOOSEROWS VALUETOTEXT ARRAYTOTEXT FLOOR.MATH PERMUTATIONA BINOM.DIST.RANGE FILTER SORT SORTBY'.split(' '));
const needs = n => n.indexOf('.') >= 0 || NEWER.has(n);
const WS = new Set(['FILTER','SORT','SORTBY','UNIQUE','SEQUENCE','RANDARRAY','XLOOKUP','XMATCH']);
const pref = f => f.replace(/([A-Z][A-Z0-9.]*)\(/g, (m, n) => needs(n) ? (WS.has(n) ? '_xlfn._xlws.' : '_xlfn.') + n + '(' : m).replace(/_xlfn\._xlws\.(SEQUENCE|RANDARRAY|XLOOKUP|XMATCH)/g, '_xlfn.$1');
const colL = n => String.fromCharCode(64 + n);
R('A1:AZ60').clearContents(); // includes the result area, so a rejected formula cannot report a stale value
for (const a in DEF) put(a, DEF[a]);
const results = [];
const dump = () => $.NSString.alloc.initWithUTF8String(JSON.stringify(results)).writeToFileAtomicallyEncodingError('results.json', true, $.NSUTF8StringEncoding, null);
const cellGot = a => { const v = R(a).value2(); return v === null ? R(a).stringValue() : v; };
for (let k = 0; k < cases.length; k++) {
  const c = cases[k]; let got;
  try {
    const touched = [];
    if (c.grid) for (const a in c.grid) { put(a, c.grid[a]); touched.push(a); }
    if (c.rows) {
      const rng = 'AA1:' + 'A' + colL(c.cols) + c.rows;
      R(rng).formulaArray = pref(c.f);
      got = []; for (let r = 1; r <= c.rows; r++) { const row = []; for (let q = 1; q <= c.cols; q++) row.push(cellGot('A' + colL(q) + r)); got.push(row); }
    } else {
      R('AA1').formula2 = pref(c.f);
      got = cellGot('AA1');
    }
    R('AA1:AZ60').clearContents();
    for (const a of touched) put(a, a in DEF ? DEF[a] : null);
  } catch (e) { got = {err: String(e)}; try { R('AA1:AZ60').clearContents(); if (c.grid) for (const a in c.grid) put(a, a in DEF ? DEF[a] : null); } catch (e2) {} }
  results.push(got);
  if (k % 100 === 0) dump();
}
dump();
'done ' + results.length;
