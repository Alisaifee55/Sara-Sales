/*! xls-writer.js v1.0 — writes an ExcelJS workbook as a real Excel 97-2003 (.xls, BIFF8) file.
 *
 * Needs the SheetJS "xlsx.full.min.js" bundle loaded first (only for its CFB container writer).
 * Public API:  XlsWriter.exportWorkbook(exceljsWorkbook) -> Uint8Array (.xls file bytes)
 *
 * Supported: values, strings (incl. rich text runs), formulas (+ - * / = <> < > <= >=, SUM, SUMIF, parentheses),
 *            fonts, solid fills, borders, alignment, number formats, merged cells, column widths/hidden,
 *            row heights/hidden/outline groups, frozen rows/cols, custom colour palette.
 * Not supported (by design, explained in the app): live conditional formatting (resolved to static fills),
 *            theme colours, images, charts, data validation.
 * Excel 97 limits enforced: 65,536 rows, 256 columns, 31-char sheet name.
 */
(function (global) {
  'use strict';

  // ======================================================================
  // Byte helpers
  // ======================================================================
  function u8(a, v) { a.push(v & 0xFF); }
  function u16(a, v) { a.push(v & 0xFF, (v >>> 8) & 0xFF); }
  function u32(a, v) { a.push(v & 0xFF, (v >>> 8) & 0xFF, (v >>> 16) & 0xFF, (v >>> 24) & 0xFF); }
  function f64(a, v) {
    var dv = new DataView(new ArrayBuffer(8));
    dv.setFloat64(0, v, true);
    for (var i = 0; i < 8; i++) a.push(dv.getUint8(i));
  }
  function append(dst, src) { for (var i = 0; i < src.length; i++) dst.push(src[i]); }

  function rec(id, data) {
    if (data.length > 8224) throw new Error('XLS record 0x' + id.toString(16) + ' is too large (' + data.length + ' bytes).');
    var out = [id & 0xFF, (id >>> 8) & 0xFF, data.length & 0xFF, (data.length >>> 8) & 0xFF];
    append(out, data);
    return out;
  }

  // Unicode string body: flags byte + chars (compressed to 1 byte/char when every code unit < 256).
  function strBody(s) {
    var wide = false, i;
    for (i = 0; i < s.length; i++) if (s.charCodeAt(i) > 255) { wide = true; break; }
    var out = [wide ? 1 : 0];
    for (i = 0; i < s.length; i++) {
      var c = s.charCodeAt(i);
      out.push(c & 0xFF);
      if (wide) out.push((c >>> 8) & 0xFF);
    }
    return out;
  }
  // [cch(1)] + body  (BIFF8 short unicode string, e.g. sheet name, font name)
  function shortStr(s) {
    if (s.length > 255) throw new Error('String too long for XLS short-string field.');
    var out = [s.length];
    append(out, strBody(s));
    return out;
  }
  // [cch(2)] + body
  function longStr(s) {
    var out = [];
    u16(out, s.length);
    append(out, strBody(s));
    return out;
  }

  // ======================================================================
  // Colours / palette
  // ======================================================================
  var DEFAULT_PALETTE = [
    0x000000, 0xFFFFFF, 0xFF0000, 0x00FF00, 0x0000FF, 0xFFFF00, 0xFF00FF, 0x00FFFF,
    0x800000, 0x008000, 0x000080, 0x808000, 0x800080, 0x008080, 0xC0C0C0, 0x808080,
    0x9999FF, 0x993366, 0xFFFFCC, 0xCCFFFF, 0x660066, 0xFF8080, 0x0066CC, 0xCCCCFF,
    0x000080, 0xFF00FF, 0xFFFF00, 0x00FFFF, 0x800080, 0x800000, 0x008080, 0x0000FF,
    0x00CCFF, 0xCCFFFF, 0xCCFFCC, 0xFFFF99, 0x99CCFF, 0xFF99CC, 0xCC99FF, 0xFFCC99,
    0x3366FF, 0x33CCCC, 0x99CC00, 0xFFCC00, 0xFF9900, 0xFF6600, 0x666699, 0x969696,
    0x003366, 0x339966, 0x003300, 0x333300, 0x993300, 0x993366, 0x333399, 0x333333
  ]; // palette slots 8..63

  function argbToRgb(argb) {
    if (argb === undefined || argb === null) return null;
    var s = String(argb).replace('#', '');
    if (s.length === 8) s = s.slice(2);
    if (!/^[0-9a-fA-F]{6}$/.test(s)) return null;
    return parseInt(s, 16);
  }
  function colourDist(a, b) {
    var dr = ((a >> 16) & 255) - ((b >> 16) & 255), dg = ((a >> 8) & 255) - ((b >> 8) & 255), db = (a & 255) - (b & 255);
    return dr * dr + dg * dg + db * db;
  }

  function Palette() { this.custom = []; this.index = {}; this.approximated = 0; }
  // returns an icv (colour index) for an RGB number
  Palette.prototype.icv = function (rgb) {
    if (rgb === 0x000000) return 0;
    if (rgb === 0xFFFFFF) return 1;
    if (this.index[rgb] !== undefined) return this.index[rgb];
    var slot;
    if (this.custom.length < 56) {
      this.custom.push(rgb);
      slot = 8 + this.custom.length - 1;
    } else {
      // palette full: use the nearest colour that is already in the palette
      var best = 8, bestD = Infinity, all = this.paletteColours(), i;
      for (i = 0; i < 56; i++) { var d = colourDist(rgb, all[i]); if (d < bestD) { bestD = d; best = 8 + i; } }
      slot = best;
      this.approximated++;
    }
    this.index[rgb] = slot;
    return slot;
  };
  Palette.prototype.paletteColours = function () {
    var out = [], i;
    for (i = 0; i < 56; i++) out.push(i < this.custom.length ? this.custom[i] : DEFAULT_PALETTE[i]);
    return out;
  };
  Palette.prototype.record = function () {
    var d = [], all = this.paletteColours(), i;
    u16(d, 56);
    for (i = 0; i < 56; i++) { u8(d, (all[i] >> 16) & 255); u8(d, (all[i] >> 8) & 255); u8(d, all[i] & 255); u8(d, 0); }
    return rec(0x0092, d);
  };

  // ======================================================================
  // Formula parser (AST) -> BIFF8 tokens + evaluator
  // ======================================================================
  function colToNum(letters) {
    var n = 0;
    for (var i = 0; i < letters.length; i++) n = n * 26 + (letters.charCodeAt(i) - 64);
    return n; // 1-based
  }
  function numToCol(n) {
    var s = '';
    while (n > 0) { var m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); }
    return s;
  }
  function parseA1(ref) { // 'B12' -> {r:12,c:2}
    var m = /^\$?([A-Z]{1,3})\$?(\d+)$/.exec(ref);
    if (!m) throw new Error('Bad cell reference: ' + ref);
    return { r: parseInt(m[2], 10), c: colToNum(m[1]) };
  }

  function tokenize(src) {
    var toks = [], i = 0, m;
    while (i < src.length) {
      var ch = src[i];
      if (ch === ' ') { i++; continue; }
      if (ch === '"') {
        var j = i + 1, str = '';
        while (j < src.length) {
          if (src[j] === '"') { if (src[j + 1] === '"') { str += '"'; j += 2; continue; } break; }
          str += src[j]; j++;
        }
        if (j >= src.length) throw new Error('Unterminated string in formula: ' + src);
        toks.push({ t: 'str', v: str }); i = j + 1; continue;
      }
      var rest = src.slice(i);
      if ((m = /^\d+(\.\d+)?([eE][+-]?\d+)?/.exec(rest)) || (m = /^\.\d+/.exec(rest))) {
        toks.push({ t: 'num', v: parseFloat(m[0]) }); i += m[0].length; continue;
      }
      if ((m = /^([A-Z][A-Z0-9.]*)\(/.exec(rest))) { toks.push({ t: 'func', v: m[1] }); i += m[1].length + 1; continue; }
      if ((m = /^(\$?)([A-Z]{1,3})(\$?)(\d+)/.exec(rest))) {
        toks.push({ t: 'ref', colAbs: !!m[1], col: colToNum(m[2]), rowAbs: !!m[3], row: parseInt(m[4], 10) });
        i += m[0].length; continue;
      }
      if ((m = /^(<>|<=|>=|[+\-*\/=<>(),:])/.exec(rest))) { toks.push({ t: 'op', v: m[1] }); i += m[1].length; continue; }
      throw new Error('Unsupported character in formula "' + src + '" near: ' + rest.slice(0, 12));
    }
    return toks;
  }

  function parseFormula(src) {
    var toks = tokenize(src), p = 0;
    function peek() { return toks[p]; }
    function isOp(v) { return toks[p] && toks[p].t === 'op' && toks[p].v === v; }
    function expectOp(v) { if (!isOp(v)) throw new Error('Formula parse error, expected "' + v + '" in: ' + src); p++; }

    function parseCompare() {
      var left = parseAdd();
      while (toks[p] && toks[p].t === 'op' && ['=', '<>', '<', '>', '<=', '>='].indexOf(toks[p].v) >= 0) {
        var op = toks[p++].v;
        left = { k: 'bin', op: op, a: left, b: parseAdd() };
      }
      return left;
    }
    function parseAdd() {
      var left = parseMul();
      while (isOp('+') || isOp('-')) { var op = toks[p++].v; left = { k: 'bin', op: op, a: left, b: parseMul() }; }
      return left;
    }
    function parseMul() {
      var left = parseUnary();
      while (isOp('*') || isOp('/')) { var op = toks[p++].v; left = { k: 'bin', op: op, a: left, b: parseUnary() }; }
      return left;
    }
    function parseUnary() {
      if (isOp('-')) { p++; return { k: 'neg', a: parseUnary() }; }
      if (isOp('+')) { p++; return parseUnary(); }
      return parsePrimary();
    }
    function parsePrimary() {
      var t = peek();
      if (!t) throw new Error('Unexpected end of formula: ' + src);
      if (t.t === 'num') { p++; return { k: 'num', v: t.v }; }
      if (t.t === 'str') { p++; return { k: 'str', v: t.v }; }
      if (t.t === 'op' && t.v === '(') { p++; var inner = parseCompare(); expectOp(')'); return { k: 'paren', a: inner }; }
      if (t.t === 'func') {
        p++;
        var args = [];
        if (!isOp(')')) {
          args.push(parseCompare());
          while (isOp(',')) { p++; args.push(parseCompare()); }
        }
        expectOp(')');
        return { k: 'func', name: t.v, args: args };
      }
      if (t.t === 'ref') {
        p++;
        var a = { k: 'ref', row: t.row, col: t.col, rowAbs: t.rowAbs, colAbs: t.colAbs };
        if (isOp(':')) {
          p++;
          var t2 = peek();
          if (!t2 || t2.t !== 'ref') throw new Error('Bad range in formula: ' + src);
          p++;
          return { k: 'area', r1: a.row, c1: a.col, r1Abs: a.rowAbs, c1Abs: a.colAbs,
                   r2: t2.row, c2: t2.col, r2Abs: t2.rowAbs, c2Abs: t2.colAbs };
        }
        return a;
      }
      throw new Error('Unexpected token in formula "' + src + '"');
    }
    var ast = parseCompare();
    if (p < toks.length) throw new Error('Unexpected trailing text in formula: ' + src);
    return ast;
  }

  // Excel 97 has no IFERROR. The report only ever uses IFERROR(<cell>/<cell>,0); rewrite that to an
  // equivalent expression made only of Excel-97 operators:  a/(b+(b=0))*(b<>0)
  function compatFormula(f) {
    return f.replace(/IFERROR\(([^,()]+),\s*0\)/g, function (all, expr) {
      var m = /^\s*(\$?[A-Z]{1,3}\$?\d+)\s*\/\s*(\$?[A-Z]{1,3}\$?\d+)\s*$/.exec(expr);
      if (!m) throw new Error('Cannot convert IFERROR formula for Excel 97: ' + all);
      return m[1] + '/(' + m[2] + '+(' + m[2] + '=0))*(' + m[2] + '<>0)';
    });
  }

  var FUNC_TAB = { SUM: { id: 4, min: 1 }, SUMIF: { id: 345, argc: 3 } };
  var BIN_PTG = { '+': 0x03, '-': 0x04, '*': 0x05, '/': 0x06, '<': 0x09, '<=': 0x0A, '=': 0x0B, '>=': 0x0C, '>': 0x0D, '<>': 0x0E };

  function refWord(col, colAbs, rowAbs) { // col is 1-based
    return (col - 1) | (colAbs ? 0 : 0x4000) | (rowAbs ? 0 : 0x8000);
  }
  function encodeAst(n, out, asRef) {
    var k = n.k;
    if (k === 'num') {
      if (n.v >= 0 && n.v <= 65535 && Math.floor(n.v) === n.v) { u8(out, 0x1E); u16(out, n.v); }
      else { u8(out, 0x1F); f64(out, n.v); }
    } else if (k === 'str') {
      if (n.v.length > 255) throw new Error('Formula string literal too long.');
      u8(out, 0x17); u8(out, n.v.length); append(out, strBody(n.v));
    } else if (k === 'ref') {
      u8(out, asRef ? 0x24 : 0x44); u16(out, n.row - 1); u16(out, refWord(n.col, n.colAbs, n.rowAbs));
    } else if (k === 'area') {
      u8(out, asRef ? 0x25 : 0x45);
      u16(out, n.r1 - 1); u16(out, n.r2 - 1);
      u16(out, refWord(n.c1, n.c1Abs, n.r1Abs)); u16(out, refWord(n.c2, n.c2Abs, n.r2Abs));
    } else if (k === 'paren') {
      encodeAst(n.a, out, false); u8(out, 0x15);
    } else if (k === 'neg') {
      encodeAst(n.a, out, false); u8(out, 0x13);
    } else if (k === 'bin') {
      encodeAst(n.a, out, false); encodeAst(n.b, out, false); u8(out, BIN_PTG[n.op]);
    } else if (k === 'func') {
      var def = FUNC_TAB[n.name];
      if (!def) throw new Error('Function ' + n.name + ' is not supported by the Excel 97 exporter.');
      if (def.argc && n.args.length !== def.argc) throw new Error(n.name + ' expects ' + def.argc + ' arguments.');
      for (var i = 0; i < n.args.length; i++) encodeAst(n.args[i], out, true);
      u8(out, 0x42); u8(out, n.args.length); u16(out, def.id);
    } else throw new Error('Cannot encode formula node ' + k);
  }

  // Evaluates an AST. env.cell(r,c) returns the cell's value (number|string|null); rowShift moves relative rows.
  function evalAst(n, env, rowShift) {
    var k = n.k;
    function R(x) { return x.rowAbs ? x.row : x.row + rowShift; }
    if (k === 'num' || k === 'str') return n.v;
    if (k === 'paren') return evalAst(n.a, env, rowShift);
    if (k === 'neg') return -num(evalAst(n.a, env, rowShift));
    if (k === 'ref') return env.cell(R(n), n.col);
    if (k === 'area') throw new Error('Range used outside a function.');
    if (k === 'bin') {
      var a = evalAst(n.a, env, rowShift), b = evalAst(n.b, env, rowShift);
      switch (n.op) {
        case '+': return num(a) + num(b);
        case '-': return num(a) - num(b);
        case '*': return num(a) * num(b);
        case '/': return num(a) / num(b);
        case '=': return cmpEq(a, b);
        case '<>': return !cmpEq(a, b);
        case '<': return num(a) < num(b);
        case '>': return num(a) > num(b);
        case '<=': return num(a) <= num(b);
        case '>=': return num(a) >= num(b);
      }
    }
    if (k === 'func') {
      if (n.name === 'SUM') {
        var s = 0;
        n.args.forEach(function (arg) {
          if (arg.k === 'area') s += sumArea(arg, env, rowShift);
          else s += num(evalAst(arg, env, rowShift));
        });
        return s;
      }
      if (n.name === 'SUMIF') {
        var rng = n.args[0], crit = evalAst(n.args[1], env, rowShift), sr = n.args[2];
        if (rng.k !== 'area' || sr.k !== 'area') throw new Error('SUMIF needs ranges.');
        var total = 0, rows = rng.r2 - rng.r1, cols = rng.c2 - rng.c1, ri, ci;
        for (ri = 0; ri <= rows; ri++) for (ci = 0; ci <= cols; ci++) {
          var cv = env.cell(rng.r1 + ri + (rng.r1Abs ? 0 : rowShift), rng.c1 + ci);
          if (cmpEq(cv, crit)) total += num(env.cell(sr.r1 + ri + (sr.r1Abs ? 0 : rowShift), sr.c1 + ci));
        }
        return total;
      }
    }
    throw new Error('Cannot evaluate formula node ' + k);
  }
  function sumArea(a, env, rowShift) {
    var s = 0, r, c, v, r1 = a.r1 + (a.r1Abs ? 0 : rowShift), r2 = a.r2 + (a.r2Abs ? 0 : rowShift);
    for (r = r1; r <= r2; r++) for (c = a.c1; c <= a.c2; c++) { v = env.cell(r, c); if (typeof v === 'number') s += v; }
    return s;
  }
  function num(v) {
    if (typeof v === 'number') return v;
    if (typeof v === 'boolean') return v ? 1 : 0;
    if (v === null || v === undefined || v === '') return 0;
    var x = Number(v); return isNaN(x) ? 0 : x;
  }
  function cmpEq(a, b) {
    if (typeof a === 'string' || typeof b === 'string') return String(a === null || a === undefined ? '' : a).toLowerCase() === String(b === null || b === undefined ? '' : b).toLowerCase();
    return num(a) === num(b);
  }

  // ======================================================================
  // Style registries
  // ======================================================================
  var BORDER_CODE = { thin: 1, medium: 2, dashed: 3, dotted: 4, thick: 5, 'double': 6, hair: 7,
    mediumDashed: 8, dashDot: 9, mediumDashDot: 10, dashDotDot: 11, mediumDashDotDot: 12, slantDashDot: 13 };
  var H_CODE = { general: 0, left: 1, center: 2, right: 3, fill: 4, justify: 5, centerContinuous: 6, distributed: 7 };
  var V_CODE = { top: 0, middle: 1, center: 1, bottom: 2, justify: 3, distributed: 4 };
  var BUILTIN_FMT = { 'General': 0, '0': 1, '0.00': 2, '#,##0': 3, '#,##0.00': 4, '0%': 9, '0.00%': 10 };

  function Styles(palette) {
    this.pal = palette;
    this.fonts = [];       // logical fonts
    this.fontIndex = {};
    this.fmts = [];        // custom formats (id = 164 + i)
    this.fmtIndex = {};
    this.xfs = [];         // cell xfs (file index = 16 + i)
    this.xfIndex = {};
    // Excel always starts with 4 standard fonts (+ a skipped slot 4)
    this.addFont({ name: 'Calibri', size: 11 });
    this.addFont({ name: 'Calibri', size: 11, bold: true });
    this.addFont({ name: 'Calibri', size: 11, italic: true });
    this.addFont({ name: 'Calibri', size: 11, bold: true, italic: true });
    this.addFont({ name: 'Calibri', size: 11 }); // logical 4 -> file index 5
  }
  Styles.prototype.addFont = function (f) {
    f = f || {};
    var rgb = f.color ? argbToRgb(f.color.argb) : null;
    var norm = { name: f.name || 'Calibri', size: f.size || 11, bold: !!f.bold, italic: !!f.italic,
                 underline: !!f.underline, strike: !!f.strike, rgb: rgb };
    var key = JSON.stringify(norm);
    if (this.fontIndex[key] === undefined) {
      this.fontIndex[key] = this.fonts.length;
      norm.icv = rgb === null ? 0x7FFF : this.pal.icv(rgb);
      this.fonts.push(norm);
    }
    return this.fontIndex[key];
  };
  Styles.fileFontIndex = function (logical) { return logical < 4 ? logical : logical + 1; };
  Styles.prototype.fmtId = function (code) {
    if (!code || code === 'General') return 0;
    if (BUILTIN_FMT[code] !== undefined) return BUILTIN_FMT[code];
    if (this.fmtIndex[code] === undefined) { this.fmtIndex[code] = 164 + this.fmts.length; this.fmts.push(code); }
    return this.fmtIndex[code];
  };
  Styles.prototype.xf = function (st) {
    st = st || {};
    var font = this.addFont(st.font);
    var fmt = this.fmtId(st.numFmt);
    var pal = this.pal;
    var fill = null;
    if (st.fill && st.fill.type === 'pattern' && st.fill.pattern === 'solid') {
      var frgb = st.fill.fgColor ? argbToRgb(st.fill.fgColor.argb) : null;
      if (frgb !== null) fill = pal.icv(frgb);
    }
    var b = st.border || {}, borders = {}, any = false;
    ['left', 'right', 'top', 'bottom'].forEach(function (side) {
      var e = b[side];
      if (e && e.style) {
        var code = BORDER_CODE[e.style];
        if (code === undefined) code = 1;
        var crgb = e.color ? argbToRgb(e.color.argb) : null;
        borders[side] = [code, crgb === null ? 64 : pal.icv(crgb)];
        any = true;
      } else borders[side] = [0, 64];
    });
    var al = st.alignment || {};
    var desc = {
      font: font, fmt: fmt, fill: fill, borders: borders, anyBorder: any,
      h: al.horizontal && H_CODE[al.horizontal] !== undefined ? H_CODE[al.horizontal] : 0,
      v: al.vertical && V_CODE[al.vertical] !== undefined ? V_CODE[al.vertical] : 2,
      wrap: !!al.wrapText
    };
    var key = JSON.stringify(desc);
    if (this.xfIndex[key] === undefined) { this.xfIndex[key] = this.xfs.length; this.xfs.push(desc); }
    return 16 + this.xfIndex[key];
  };

  function fontRecord(f) {
    var d = [];
    u16(d, Math.round(f.size * 20));
    u16(d, (f.bold ? 0x01 : 0) | (f.italic ? 0x02 : 0) | (f.strike ? 0x08 : 0));
    u16(d, f.icv);
    u16(d, f.bold ? 700 : 400);
    u16(d, 0);                       // no super/subscript
    u8(d, f.underline ? 1 : 0);
    u8(d, 0); u8(d, 0); u8(d, 0);    // family, charset, reserved
    append(d, shortStr(f.name));
    return rec(0x0031, d);
  }
  function styleXfRecord(index) {
    var d = [];
    u16(d, 0); u16(d, 0);
    u16(d, 0xFFF5);                  // style XF, parent none
    u8(d, 0x20); u8(d, 0); u8(d, 0); u8(d, 0);
    u32(d, 0); u32(d, 0);
    u16(d, 0x20C0);
    return rec(0x00E0, d);
  }
  function defaultCellXfRecord() {
    var d = [];
    u16(d, 0); u16(d, 0);
    u16(d, 0x0001);
    u8(d, 0x20); u8(d, 0); u8(d, 0); u8(d, 0);
    u32(d, 0); u32(d, 0);
    u16(d, 0x20C0);
    return rec(0x00E0, d);
  }
  function cellXfRecord(x) {
    var d = [], B = x.borders;
    u16(d, Styles.fileFontIndex(x.font));
    u16(d, x.fmt);
    u16(d, 0x0001);                  // locked, parent style XF 0
    u8(d, (x.h & 7) | (x.wrap ? 0x08 : 0) | ((x.v & 7) << 4));
    u8(d, 0);
    u8(d, 0);
    u8(d, 0xFC);                     // all attribute groups are used
    u32(d, (B.left[0] & 15) | ((B.right[0] & 15) << 4) | ((B.top[0] & 15) << 8) | ((B.bottom[0] & 15) << 12) |
           ((B.left[1] & 127) << 16) | ((B.right[1] & 127) << 23));
    var fillPattern = x.fill === null ? 0 : 1;
    u32(d, ((B.top[1] & 127) | ((B.bottom[1] & 127) << 7) | (64 << 14) | (fillPattern << 26)) >>> 0);
    u16(d, (x.fill === null ? 64 : x.fill) | ((x.fill === null ? 65 : 65) << 7));
    return rec(0x00E0, d);
  }

  // ======================================================================
  // Shared string table
  // ======================================================================
  function Sst() { this.list = []; this.index = {}; this.refs = 0; }
  Sst.prototype.add = function (text, runs) {
    this.refs++;
    var key = JSON.stringify([text, runs || null]);
    if (this.index[key] === undefined) { this.index[key] = this.list.length; this.list.push({ text: text, runs: runs || null }); }
    return this.index[key];
  };
  function sstItemBytes(it) {
    var out = [], body = strBody(it.text), rich = it.runs && it.runs.length > 0;
    u16(out, it.text.length);
    u8(out, body[0] | (rich ? 0x08 : 0));
    if (rich) u16(out, it.runs.length);
    for (var i = 1; i < body.length; i++) out.push(body[i]);
    if (rich) it.runs.forEach(function (r) { u16(out, r[0]); u16(out, r[1]); });
    return out;
  }
  Sst.prototype.records = function () {
    var out = [], chunk = [], first = true, self = this;
    function flush() {
      var data = [];
      if (first) { u32(data, self.refs); u32(data, self.list.length); append(data, chunk); out.push.apply(out, rec(0x00FC, data)); first = false; }
      else out.push.apply(out, rec(0x003C, chunk));
      chunk = [];
    }
    var limit = 8224 - 8;
    this.list.forEach(function (it) {
      var b = sstItemBytes(it);
      if (b.length > limit) throw new Error('A text cell is too long for the Excel 97 format.');
      if (chunk.length + b.length > (first ? limit : 8224)) flush();
      append(chunk, b);
    });
    if (chunk.length || first) flush();
    return out;
  };

  // ======================================================================
  // ExcelJS -> intermediate model
  // ======================================================================
  function rangeOf(str) {
    var p = str.split(':'), a = parseA1(p[0]), b = parseA1(p[1] || p[0]);
    return { r1: a.r, c1: a.c, r2: b.r, c2: b.c };
  }

  function convertSheet(ws, ExcelJS) {
    var cells = {}, maxR = 0, maxC = 0, key;
    function put(r, c, obj) { cells[r + ',' + c] = obj; if (r > maxR) maxR = r; if (c > maxC) maxC = c; }

    ws.eachRow({ includeEmpty: true }, function (row, rn) {
      row.eachCell({ includeEmpty: true }, function (cell, cn) {
        var st = cell.style || {};
        var isSlave = cell.isMerged && cell.master && cell.master.address !== cell.address;
        var obj = { style: JSON.parse(JSON.stringify(st)), kind: 'blank' };
        if (!isSlave) {
          var v = cell.value;
          if (v === null || v === undefined) obj.kind = 'blank';
          else if (typeof v === 'number') { obj.kind = 'n'; obj.v = v; }
          else if (typeof v === 'string') { obj.kind = 's'; obj.v = v; }
          else if (typeof v === 'boolean') { obj.kind = 's'; obj.v = v ? 'TRUE' : 'FALSE'; }
          else if (v instanceof Date) { obj.kind = 'n'; obj.v = v.getTime() / 86400000 + 25569; }
          else if (v.richText) { obj.kind = 'rich'; obj.rich = v.richText; }
          else if (typeof v.formula === 'string') { obj.kind = 'f'; obj.formula = v.formula; }
          else throw new Error('Unsupported value in cell ' + cell.address + ' for the Excel 97 format.');
        }
        var hasStyle = Object.keys(obj.style).length > 0;
        if (obj.kind !== 'blank' || hasStyle) put(rn, cn, obj);
      });
    });

    if (maxR > 65536) throw new Error('This report has more than 65,536 rows — the Excel 97 format cannot hold it.');
    if (maxC > 256) throw new Error('This report has more than 256 columns — the Excel 97 format cannot hold it.');

    // Conditional formatting -> static fills (live rules are not written to .xls)
    var cfResolved = 0;
    (ws.conditionalFormattings || []).forEach(function (cf) {
      var rg = rangeOf(cf.ref);
      var sortedRules = (cf.rules || []).slice().sort(function (a, b) { return (a.priority || 0) - (b.priority || 0); });
      var env = makeEnv(cells);
      for (var r = rg.r1; r <= rg.r2; r++) for (var c = rg.c1; c <= rg.c2; c++) {
        var cell = cells[r + ',' + c]; if (!cell) continue;
        for (var i = 0; i < sortedRules.length; i++) {
          var rule = sortedRules[i];
          if (rule.type !== 'expression' || !rule.formulae || !rule.formulae[0]) continue;
          var ast = parseFormula(compatFormula(rule.formulae[0]));
          var truth;
          try { truth = evalAst(ast, env, r - rg.r1); } catch (e) { throw new Error('Conditional format could not be resolved: ' + e.message); }
          if (truth === true) {
            var ruleStyle = rule.style || {};
            if (ruleStyle.fill) {
              var colour = (ruleStyle.fill.bgColor && ruleStyle.fill.bgColor.argb) || (ruleStyle.fill.fgColor && ruleStyle.fill.fgColor.argb);
              if (colour) cell.style.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: colour } };
            }
            if (ruleStyle.font) cell.style.font = Object.assign({}, cell.style.font || {}, ruleStyle.font);
            cfResolved++;
            break;
          }
        }
      }
    });

    var colInfo = [], ci;
    var colCount = Math.max(maxC, ws.columnCount || 0);
    for (ci = 1; ci <= colCount; ci++) {
      var col = ws.getColumn(ci);
      colInfo.push({ c: ci, width: col.width, hidden: !!col.hidden });
    }

    var rows = {};
    var lastRow = Math.max(maxR, ws.rowCount || 0);
    for (var rr = 1; rr <= lastRow; rr++) {
      var row = ws.getRow(rr);
      rows[rr] = { height: row.height || null, hidden: !!row.hidden, level: row.outlineLevel || 0, collapsed: false };
    }
    // summary row sitting right below a hidden outline group is shown as collapsed
    for (rr = 2; rr <= lastRow; rr++) {
      if (rows[rr].level < rows[rr - 1].level && rows[rr - 1].hidden) rows[rr].collapsed = true;
    }

    var view = (ws.views && ws.views[0]) || {};
    return {
      name: String(ws.name || 'Sheet1').replace(/[\\\/\?\*\[\]:]/g, '_').slice(0, 31),
      cells: cells, maxR: maxR, maxC: maxC, cols: colInfo, rows: rows,
      merges: ((ws.model && ws.model.merges) || []).map(rangeOf),
      frozen: view.state === 'frozen' ? { x: view.xSplit || 0, y: view.ySplit || 0 } : null,
      cfResolved: cfResolved
    };
  }

  // cached-value evaluation environment (formulas are evaluated here for the stored results)
  function makeEnv(cells) {
    var memo = {}, busy = {};
    var env = {
      cell: function (r, c) {
        var k = r + ',' + c, o = cells[k];
        if (!o) return null;
        if (o.kind === 'n' || o.kind === 's') return o.v;
        if (o.kind === 'rich') return o.rich.map(function (x) { return x.text; }).join('');
        if (o.kind === 'f') {
          if (memo[k] !== undefined) return memo[k];
          if (busy[k]) throw new Error('Circular reference at ' + numToCol(c) + r);
          busy[k] = true;
          var ast = o.ast || (o.ast = parseFormula(compatFormula(o.formula)));
          var val = evalAst(ast, env, 0);
          busy[k] = false;
          memo[k] = val;
          return val;
        }
        return null;
      }
    };
    return env;
  }

  // ======================================================================
  // Sheet + workbook assembly
  // ======================================================================
  function bofRecord(type) {
    var d = [];
    u16(d, 0x0600); u16(d, type); u16(d, 0x0DBB); u16(d, 0x07CC); u32(d, 0); u32(d, 6);
    return rec(0x0809, d);
  }

  function buildSheetStream(sh, styles, sst) {
    var out = [], r, c;
    var env = makeEnv(sh.cells);

    // --- pre-compute XF ids / SST ids for every cell, row by row ---
    var byRow = {};
    Object.keys(sh.cells).forEach(function (k) {
      var p = k.split(','), rr = +p[0], cc = +p[1];
      (byRow[rr] = byRow[rr] || []).push({ r: rr, c: cc, o: sh.cells[k] });
    });

    var maxLevel = 0;
    Object.keys(sh.rows).forEach(function (k) { if (sh.rows[k].level > maxLevel) maxLevel = sh.rows[k].level; });

    append(out, bofRecord(0x0010));

    // GUTS (outline gutters)
    var g = []; u16(g, maxLevel ? 5 + 12 * maxLevel : 0); u16(g, 0); u16(g, maxLevel ? maxLevel + 1 : 0); u16(g, 0);
    append(out, rec(0x0080, g));
    // DEFAULTROWHEIGHT 15pt, WSBOOL (summary rows below, outline), DEFCOLWIDTH
    var dr = []; u16(dr, 0); u16(dr, 300); append(out, rec(0x0225, dr));
    var wb = []; u16(wb, 0x0441); append(out, rec(0x0081, wb));
    var dc = []; u16(dc, 8); append(out, rec(0x0055, dc));

    // COLINFO (one record per column)
    sh.cols.forEach(function (ci) {
      if (ci.width === undefined || ci.width === null) { if (!ci.hidden) return; }
      var d = [], w = (ci.width === undefined || ci.width === null) ? 8.43 : ci.width;
      u16(d, ci.c - 1); u16(d, ci.c - 1);
      u16(d, Math.min(65535, Math.round(w * 256)));
      u16(d, 15);
      u16(d, ci.hidden ? 1 : 0);
      u16(d, 0);
      append(out, rec(0x007D, d));
    });

    // DIMENSIONS
    var lastRowNo = Math.max(sh.maxR, 1), lastColNo = Math.max(sh.maxC, 1);
    var dim = []; u32(dim, 0); u32(dim, lastRowNo); u16(dim, 0); u16(dim, lastColNo); u16(dim, 0);
    append(out, rec(0x0200, dim));

    // ROW records
    var rowNums = Object.keys(sh.rows).map(Number).sort(function (a, b) { return a - b; });
    rowNums.forEach(function (rn) {
      var ri = sh.rows[rn], has = byRow[rn];
      var custom = ri.height !== null || ri.hidden || ri.level > 0 || ri.collapsed;
      if (!has && !custom) return;
      var d = [], minC = 0, maxCc = 0;
      if (has) {
        has.sort(function (a, b) { return a.c - b.c; });
        minC = has[0].c - 1; maxCc = has[has.length - 1].c;
      }
      u16(d, rn - 1); u16(d, minC); u16(d, maxCc);
      u16(d, ri.height !== null ? Math.round(ri.height * 20) : 300);
      u16(d, 0); u16(d, 0);
      u16(d, (ri.level & 7) | (ri.collapsed ? 0x10 : 0) | (ri.hidden ? 0x20 : 0) | (ri.height !== null ? 0x40 : 0) | 0x100);
      u16(d, 15);
      append(out, rec(0x0208, d));
    });

    // Cell records
    var cellOrder = Object.keys(byRow).map(Number).sort(function (a, b) { return a - b; });
    cellOrder.forEach(function (rn) {
      byRow[rn].sort(function (a, b) { return a.c - b.c; }).forEach(function (cell) {
        var o = cell.o, xf = styles.xf(o.style), d = [];
        var row0 = cell.r - 1, col0 = cell.c - 1;
        if (o.kind === 'blank') {
          u16(d, row0); u16(d, col0); u16(d, xf); append(out, rec(0x0201, d));
        } else if (o.kind === 'n') {
          u16(d, row0); u16(d, col0); u16(d, xf); f64(d, o.v); append(out, rec(0x0203, d));
        } else if (o.kind === 's') {
          u16(d, row0); u16(d, col0); u16(d, xf); u32(d, sst.add(o.v, null)); append(out, rec(0x00FD, d));
        } else if (o.kind === 'rich') {
          var text = '', runs = [];
          o.rich.forEach(function (seg) {
            runs.push([text.length, Styles.fileFontIndex(styles.addFont(Object.assign({}, (o.style && o.style.font) || {}, seg.font || {})))]);
            text += seg.text;
          });
          u16(d, row0); u16(d, col0); u16(d, xf); u32(d, sst.add(text, runs)); append(out, rec(0x00FD, d));
        } else if (o.kind === 'f') {
          var ast = o.ast || (o.ast = parseFormula(compatFormula(o.formula)));
          var rgce = []; encodeAst(ast, rgce, false);
          var val = evalAst(ast, env, 0);
          if (typeof val === 'boolean') val = val ? 1 : 0;
          if (typeof val !== 'number' || !isFinite(val)) val = 0;
          u16(d, row0); u16(d, col0); u16(d, xf); f64(d, val);
          u16(d, 0x0001); u32(d, 0); u16(d, rgce.length); append(d, rgce);
          append(out, rec(0x0006, d));
        }
      });
    });

    // WINDOW2
    var w2 = []; u16(w2, sh.frozen ? 0x06BE : 0x06B6); u16(w2, 0); u16(w2, 0); u32(w2, 0x40); u16(w2, 0); u16(w2, 0); u32(w2, 0);
    append(out, rec(0x023E, w2));
    if (sh.frozen && (sh.frozen.x || sh.frozen.y)) {
      var fx = sh.frozen.x || 0, fy = sh.frozen.y || 0;
      var active = fx && fy ? 0 : (fy ? 2 : 1);
      var pn = []; u16(pn, fx); u16(pn, fy); u16(pn, fy); u16(pn, fx); u8(pn, active); u8(pn, 0);
      append(out, rec(0x0041, pn));
      var panes = fx && fy ? [3, 1, 2, 0] : (fy ? [3, 2] : [1, 0]);
      panes.forEach(function (pnn) {
        var s = [];
        u8(s, pnn);
        var ar = (pnn === 2 || pnn === 0) ? fy : 0, ac = (pnn === 1 || pnn === 0) ? fx : 0;
        u16(s, ar); u16(s, ac); u16(s, 0); u16(s, 1);
        u16(s, ar); u16(s, ar); u8(s, ac); u8(s, ac);
        append(out, rec(0x001D, s));
      });
    }

    // MERGEDCELLS (max 1026 per record)
    for (var i = 0; i < sh.merges.length; i += 1026) {
      var part = sh.merges.slice(i, i + 1026), md = [];
      u16(md, part.length);
      part.forEach(function (m) { u16(md, m.r1 - 1); u16(md, m.r2 - 1); u16(md, m.c1 - 1); u16(md, m.c2 - 1); });
      append(out, rec(0x00E5, md));
    }

    append(out, rec(0x000A, []));
    return out;
  }

  function buildWorkbookStream(ejsWb) {
    var sheets = ejsWb.worksheets;
    if (!sheets || sheets.length === 0) throw new Error('The workbook has no worksheet.');
    if (sheets.length > 1) throw new Error('The Excel 97 exporter handles one worksheet per workbook.');
    var palette = new Palette();
    var styles = new Styles(palette);
    var sst = new Sst();
    var model = convertSheet(sheets[0]);
    var sheetBytes = buildSheetStream(model, styles, sst);

    // --- globals ---
    var g = [];
    append(g, bofRecord(0x0005));
    var cp = []; u16(cp, 0x04B0); append(g, rec(0x0042, cp));
    var dm = []; u16(dm, 0); append(g, rec(0x0022, dm));
    var w1 = []; u16(w1, 0x0168); u16(w1, 0x0087); u16(w1, 0x3A5E); u16(w1, 0x23BE); u16(w1, 0x0038); u16(w1, 0); u16(w1, 0); u16(w1, 1); u16(w1, 0x0258);
    append(g, rec(0x003D, w1));
    styles.fonts.forEach(function (f) { append(g, fontRecord(f)); });
    styles.fmts.forEach(function (code, i) {
      var d = []; u16(d, 164 + i); append(d, longStr(code)); append(g, rec(0x041E, d));
    });
    var n;
    for (n = 0; n < 15; n++) append(g, styleXfRecord(n));
    append(g, defaultCellXfRecord());
    styles.xfs.forEach(function (x) { append(g, cellXfRecord(x)); });
    if (styles.xfs.length + 16 > 4000) throw new Error('Too many distinct cell formats for the Excel 97 format.');
    var st = []; u16(st, 0x8000); u8(st, 0); u8(st, 0xFF); append(g, rec(0x0293, st));
    if (palette.custom.length) append(g, palette.record());

    // BOUNDSHEET (offset patched below)
    var bsPos = g.length;
    var bs = []; u32(bs, 0); u16(bs, 0); append(bs, shortStr(model.name));
    append(g, rec(0x0085, bs));
    append(g, sst.records());
    append(g, rec(0x000A, []));

    var offset = g.length;
    // lbPlyPos sits after the 4-byte record header
    g[bsPos + 4] = offset & 0xFF; g[bsPos + 5] = (offset >>> 8) & 0xFF; g[bsPos + 6] = (offset >>> 16) & 0xFF; g[bsPos + 7] = (offset >>> 24) & 0xFF;

    var all = g.concat(sheetBytes);
    while (all.length < 4096) all.push(0); // keep the stream out of the CFB mini-stream
    return { bytes: all, info: { cfResolved: model.cfResolved, approximatedColours: palette.approximated } };
  }

  function exportWorkbook(ejsWb, opts) {
    var XLSX = (opts && opts.XLSX) || global.XLSX;
    if (!XLSX || !XLSX.CFB) throw new Error('SheetJS (xlsx.full.min.js) must be loaded before the Excel 97 exporter can run.');
    var built = buildWorkbookStream(ejsWb);
    var cfb = XLSX.CFB.utils.cfb_new({ root: 'R' });
    XLSX.CFB.utils.cfb_add(cfb, '/Workbook', built.bytes);
    var out = XLSX.CFB.write(cfb, { type: 'array' });
    var u8out = out instanceof Uint8Array ? out : new Uint8Array(out);
    u8out.xlsInfo = built.info;
    return u8out;
  }

  var api = { exportWorkbook: exportWorkbook, _internals: { parseFormula: parseFormula, compatFormula: compatFormula, evalAst: evalAst, encodeAst: encodeAst, buildWorkbookStream: buildWorkbookStream } };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  global.XlsWriter = api;
})(typeof window !== 'undefined' ? window : globalThis);
