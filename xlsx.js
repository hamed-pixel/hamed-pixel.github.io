/* Smartech Stock — styled .xlsx writer (raw OOXML via JSZip). Works in browser (window.buildStockXlsx) and Node. */
(function (root) {
  var NAVY = '1F2A44', GOLD = 'C9A227', GOLDLT = 'F7EDC8', BAND = 'EEF1F6', LINE = 'D5DAE3';

  function esc(s) { return String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;'); }
  function col(i) { var s = ''; i++; while (i > 0) { var m = (i - 1) % 26; s = String.fromCharCode(65 + m) + s; i = Math.floor((i - 1) / 26); } return s; }
  function serial(iso) { var p = iso.split('-'); return (Date.UTC(+p[0], p[1] - 1, +p[2]) - Date.UTC(1899, 11, 30)) / 86400000; }

  // cellXfs indexes
  var S = { title: 1, sub: 2, head: 3, txt: 4, txtB: 5, date: 6, dateB: 7, num: 8, numB: 9, gTxt: 10, gNum: 11, tTxt: 12, tNum: 13, note: 14, sec: 15 };

  var styles = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
    '<numFmts count="2"><numFmt numFmtId="164" formatCode="dd/mm/yyyy"/><numFmt numFmtId="165" formatCode="#,##0.##"/></numFmts>' +
    '<fonts count="7">' +
    '<font><sz val="11"/><color rgb="FF222222"/><name val="Calibri"/></font>' +
    '<font><b/><sz val="11"/><color rgb="FF' + GOLD + '"/><name val="Calibri"/></font>' +
    '<font><b/><sz val="16"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font>' +
    '<font><i/><sz val="10"/><color rgb="FF5B6475"/><name val="Calibri"/></font>' +
    '<font><b/><sz val="11"/><color rgb="FF' + NAVY + '"/><name val="Calibri"/></font>' +
    '<font><b/><sz val="13"/><color rgb="FF' + NAVY + '"/><name val="Calibri"/></font>' +
    '<font><sz val="10"/><color rgb="FFC9D2E3"/><name val="Calibri"/></font>' +
    '</fonts>' +
    '<fills count="6"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill>' +
    '<fill><patternFill patternType="solid"><fgColor rgb="FF' + NAVY + '"/></patternFill></fill>' +
    '<fill><patternFill patternType="solid"><fgColor rgb="FF' + BAND + '"/></patternFill></fill>' +
    '<fill><patternFill patternType="solid"><fgColor rgb="FF' + GOLDLT + '"/></patternFill></fill>' +
    '<fill><patternFill patternType="solid"><fgColor rgb="FFFFFFFF"/></patternFill></fill></fills>' +
    '<borders count="3"><border/>' +
    '<border><left style="thin"><color rgb="FF' + LINE + '"/></left><right style="thin"><color rgb="FF' + LINE + '"/></right><top style="thin"><color rgb="FF' + LINE + '"/></top><bottom style="thin"><color rgb="FF' + LINE + '"/></bottom></border>' +
    '<border><top style="medium"><color rgb="FF' + GOLD + '"/></top><bottom style="medium"><color rgb="FF' + GOLD + '"/></bottom></border>' +
    '</borders>' +
    '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
    '<cellXfs count="16">' +
    '<xf xfId="0" numFmtId="0" fontId="0" fillId="0" borderId="0"/>' +
    '<xf xfId="0" numFmtId="0" fontId="2" fillId="2" borderId="0" applyFont="1" applyFill="1"><alignment vertical="center" indent="1"/></xf>' +
    '<xf xfId="0" numFmtId="0" fontId="3" fillId="0" borderId="0" applyFont="1"><alignment indent="1"/></xf>' +
    '<xf xfId="0" numFmtId="0" fontId="1" fillId="2" borderId="1" applyFont="1" applyFill="1" applyBorder="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>' +
    '<xf xfId="0" numFmtId="0" fontId="0" fillId="5" borderId="1" applyFill="1" applyBorder="1"><alignment vertical="center" indent="1"/></xf>' +
    '<xf xfId="0" numFmtId="0" fontId="0" fillId="3" borderId="1" applyFill="1" applyBorder="1"><alignment vertical="center" indent="1"/></xf>' +
    '<xf xfId="0" numFmtId="164" fontId="0" fillId="5" borderId="1" applyNumberFormat="1" applyFill="1" applyBorder="1"><alignment horizontal="center"/></xf>' +
    '<xf xfId="0" numFmtId="164" fontId="0" fillId="3" borderId="1" applyNumberFormat="1" applyFill="1" applyBorder="1"><alignment horizontal="center"/></xf>' +
    '<xf xfId="0" numFmtId="165" fontId="0" fillId="5" borderId="1" applyNumberFormat="1" applyFill="1" applyBorder="1"><alignment horizontal="right"/></xf>' +
    '<xf xfId="0" numFmtId="165" fontId="0" fillId="3" borderId="1" applyNumberFormat="1" applyFill="1" applyBorder="1"><alignment horizontal="right"/></xf>' +
    '<xf xfId="0" numFmtId="0" fontId="4" fillId="4" borderId="1" applyFont="1" applyFill="1" applyBorder="1"/>' +
    '<xf xfId="0" numFmtId="165" fontId="4" fillId="4" borderId="1" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1"><alignment horizontal="right"/></xf>' +
    '<xf xfId="0" numFmtId="0" fontId="1" fillId="2" borderId="2" applyFont="1" applyFill="1" applyBorder="1"/>' +
    '<xf xfId="0" numFmtId="165" fontId="1" fillId="2" borderId="2" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1"><alignment horizontal="right"/></xf>' +
    '<xf xfId="0" numFmtId="0" fontId="3" fillId="0" borderId="0" applyFont="1"><alignment wrapText="1" vertical="top"/></xf>' +
    '<xf xfId="0" numFmtId="0" fontId="5" fillId="0" borderId="0" applyFont="1"/>' +
    '</cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>';

  // Sheet builder: rows = [{h:height, cells:[[value, style, kind]]}]  kind: 's' text, 'n' number, 'd' iso date
  function sheetXml(rows, cols, merges, freezeRow, extra) {
    var x = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">';
    x += '<sheetViews><sheetView workbookViewId="0" showGridLines="0">' +
      (freezeRow ? '<pane ySplit="' + freezeRow + '" topLeftCell="A' + (freezeRow + 1) + '" activePane="bottomLeft" state="frozen"/>' : '') +
      '</sheetView></sheetViews><sheetFormatPr defaultRowHeight="18"/>';
    x += '<cols>' + cols.map(function (c, i) { return '<col min="' + (i + 1) + '" max="' + (i + 1) + '" width="' + c.w + '" customWidth="1"' + (c.hidden ? ' hidden="1"' : '') + '/>'; }).join('') + '</cols><sheetData>';
    rows.forEach(function (r, ri) {
      x += '<row r="' + (ri + 1) + '"' + (r.h ? ' ht="' + r.h + '" customHeight="1"' : '') + '>';
      (r.cells || []).forEach(function (c, ci) {
        if (!c) return;
        var ref = col(ci) + (ri + 1), v = c[0], st = c[1] || 0, k = c[2] || (typeof v === 'number' ? 'n' : 's');
        if (v === '' || v == null) x += '<c r="' + ref + '" s="' + st + '"/>';
        else if (k === 'n') x += '<c r="' + ref + '" s="' + st + '"><v>' + v + '</v></c>';
        else if (k === 'd') x += '<c r="' + ref + '" s="' + st + '"><v>' + serial(v) + '</v></c>';
        else x += '<c r="' + ref + '" s="' + st + '" t="inlineStr"><is><t xml:space="preserve">' + esc(v) + '</t></is></c>';
      });
      x += '</row>';
    });
    x += '</sheetData>';
    if (merges.length) x += '<mergeCells count="' + merges.length + '">' + merges.map(function (m) { return '<mergeCell ref="' + m + '"/>'; }).join('') + '</mergeCells>';
    x += '<pageMargins left="0.5" right="0.5" top="0.6" bottom="0.6" header="0.3" footer="0.3"/><pageSetup orientation="portrait" fitToHeight="0"/>';
    return x + (extra || '') + '</worksheet>';
  }

  function chartXml(title, dir, catRef, cats, series) {
    var A = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"';
    function strCache(a) { return '<c:strCache><c:ptCount val="' + a.length + '"/>' + a.map(function (v, i) { return '<c:pt idx="' + i + '"><c:v>' + esc(v) + '</c:v></c:pt>'; }).join('') + '</c:strCache>'; }
    function numCache(a) { return '<c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="' + a.length + '"/>' + a.map(function (v, i) { return '<c:pt idx="' + i + '"><c:v>' + v + '</c:v></c:pt>'; }).join('') + '</c:numCache>'; }
    var txt = '<c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="900"><a:solidFill><a:srgbClr val="3A4357"/></a:solidFill></a:defRPr></a:pPr><a:endParaRPr lang="en-US"/></a:p></c:txPr>';
    var x = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" ' + A + ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><c:roundedCorners val="0"/><c:chart>' +
      '<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="1200" b="1"/></a:pPr><a:r><a:rPr lang="en-US" sz="1200" b="1"><a:solidFill><a:srgbClr val="' + NAVY + '"/></a:solidFill></a:rPr><a:t>' + esc(title) + '</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title>' +
      '<c:autoTitleDeleted val="0"/><c:plotArea><c:layout/><c:barChart><c:barDir val="' + dir + '"/><c:grouping val="clustered"/><c:varyColors val="0"/>';
    series.forEach(function (s, i) {
      x += '<c:ser><c:idx val="' + i + '"/><c:order val="' + i + '"/><c:tx><c:v>' + esc(s.name) + '</c:v></c:tx>' +
        '<c:spPr><a:solidFill><a:srgbClr val="' + s.color + '"/></a:solidFill></c:spPr><c:invertIfNegative val="0"/>' +
        '<c:cat><c:strRef><c:f>' + catRef + '</c:f>' + strCache(cats) + '</c:strRef></c:cat>' +
        '<c:val><c:numRef><c:f>' + s.ref + '</c:f>' + numCache(s.values) + '</c:numRef></c:val></c:ser>';
    });
    var catPos = dir === 'col' ? 'b' : 'l', valPos = dir === 'col' ? 'l' : 'b';
    x += '<c:gapWidth val="70"/><c:overlap val="-10"/><c:axId val="111"/><c:axId val="222"/></c:barChart>' +
      '<c:catAx><c:axId val="111"/><c:scaling><c:orientation val="' + (dir === 'bar' ? 'maxMin' : 'minMax') + '"/></c:scaling><c:delete val="0"/><c:axPos val="' + catPos + '"/><c:numFmt formatCode="General" sourceLinked="0"/><c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/>' + txt + '<c:crossAx val="222"/><c:crosses val="autoZero"/><c:auto val="1"/><c:lblAlgn val="ctr"/><c:lblOffset val="100"/><c:noMultiLvlLbl val="0"/></c:catAx>' +
      '<c:valAx><c:axId val="222"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="' + valPos + '"/><c:majorGridlines><c:spPr><a:ln w="6350"><a:solidFill><a:srgbClr val="E1E5EC"/></a:solidFill></a:ln></c:spPr></c:majorGridlines><c:numFmt formatCode="General" sourceLinked="0"/><c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/>' + txt + '<c:crossAx val="111"/><c:crosses val="' + (dir === 'bar' ? 'max' : 'autoZero') + '"/><c:crossBetween val="between"/></c:valAx>' +
      '</c:plotArea><c:legend><c:legendPos val="b"/><c:overlay val="0"/>' + txt + '</c:legend><c:plotVisOnly val="0"/><c:dispBlanksAs val="gap"/></c:chart></c:chartSpace>';
    return x;
  }

  function anchor(id, c1, r1, c2, r2) {
    return '<xdr:twoCellAnchor editAs="oneCell"><xdr:from><xdr:col>' + c1 + '</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>' + r1 + '</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>' + c2 + '</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>' + r2 + '</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>' +
      '<xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="' + (id + 1) + '" name="Chart ' + id + '"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm>' +
      '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" r:id="rId' + id + '"/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:twoCellAnchor>';
  }

  function unitLabel(units) { var k = Object.keys(units); return k.length === 1 ? k[0] : (k.length ? 'mixed' : ''); }

  /* model: {company, title, periodLabel, groupBy ('Name'|'Day'|'Month'), groups:[{label, rows:[{date, product, unit, ordered, received, receiver}]}],
            byName:[{name, items, ordered, received}], byProduct:[{product, unit, ordered, received}], generated} */
  function build(JSZip, m) {
    var zip = new JSZip();
    // ---- Report sheet
    var R = [], merges = ['A1:F1', 'A2:F2'];
    R.push({ h: 30, cells: [[m.company + ' — ' + m.title, S.title], ['', S.title], ['', S.title], ['', S.title], ['', S.title], ['', S.title]] });
    R.push({ cells: [[m.periodLabel + '  ·  ' + m.groupBy + '-wise  ·  Generated ' + m.generated, S.sub]] });
    R.push({ h: 8, cells: [] });
    R.push({ h: 24, cells: ['Date', 'Product', 'Unit', 'Order Qty', 'Received Qty', 'Receiver'].map(function (h) { return [h, S.head]; }) });
    var gO = 0, gR = 0, gU = {};
    m.groups.forEach(function (g) {
      var o = 0, r = 0, u = {};
      g.rows.forEach(function (row, i) {
        var b = i % 2 === 1;
        R.push({ cells: [[row.date, b ? S.dateB : S.date, 'd'], [row.product, b ? S.txtB : S.txt], [row.unit, b ? S.txtB : S.txt], [row.ordered, b ? S.numB : S.num, 'n'], [row.received, b ? S.numB : S.num, 'n'], [row.receiver || 'Pending', b ? S.txtB : S.txt]] });
        o += row.ordered; r += row.received; u[row.unit] = 1;
      });
      R.push({ cells: [['Total — ' + g.label, S.gTxt], ['', S.gTxt], [unitLabel(u), S.gTxt], [o, S.gNum, 'n'], [r, S.gNum, 'n'], [g.rows.length + ' line' + (g.rows.length === 1 ? '' : 's'), S.gTxt]] });
      merges.push('A' + R.length + ':B' + R.length);
      gO += o; gR += r; Object.assign(gU, u);
    });
    var lines = m.groups.reduce(function (a, g) { return a + g.rows.length; }, 0);
    R.push({ h: 22, cells: [['GRAND TOTAL', S.tTxt], ['', S.tTxt], [unitLabel(gU), S.tTxt], [gO, S.tNum, 'n'], [gR, S.tNum, 'n'], [lines + ' lines', S.tTxt]] });
    merges.push('A' + R.length + ':B' + R.length);
    R.push({ h: 8, cells: [] });
    R.push({ h: 30, cells: [['Note: total rows add quantities across different units (pcs, box, can, packet, roll) when the Unit column says "mixed". See the Summary sheet for product totals by unit.', S.note]] });
    merges.push('A' + R.length + ':F' + R.length);
    var s1 = sheetXml(R, [{ w: 13 }, { w: 26 }, { w: 9 }, { w: 12 }, { w: 14 }, { w: 22 }], merges, 4);

    // ---- Summary sheet
    var Q = [], mg = ['A1:D1', 'A2:D2'];
    Q.push({ h: 30, cells: [['Summary — ' + m.periodLabel, S.title], ['', S.title], ['', S.title], ['', S.title]] });
    Q.push({ cells: [[m.company + '  ·  Generated ' + m.generated, S.sub]] });
    Q.push({ h: 8, cells: [] });
    Q.push({ cells: [['By receiver name', S.sec]] });
    Q.push({ h: 22, cells: ['Name', 'Lines', 'Ordered', 'Received'].map(function (h) { return [h, S.head]; }) });
    var n0 = Q.length + 1;
    m.byName.forEach(function (n, i) { var b = i % 2; Q.push({ cells: [[n.name, b ? S.txtB : S.txt], [n.items, b ? S.numB : S.num, 'n'], [n.ordered, b ? S.numB : S.num, 'n'], [n.received, b ? S.numB : S.num, 'n']] }); });
    var n1 = Q.length;
    Q.push({ cells: [['Total (mixed units)', S.tTxt], [m.byName.reduce(function (a, n) { return a + n.items; }, 0), S.tNum, 'n'], [m.byName.reduce(function (a, n) { return a + n.ordered; }, 0), S.tNum, 'n'], [m.byName.reduce(function (a, n) { return a + n.received; }, 0), S.tNum, 'n']] });
    Q.push({ h: 12, cells: [] });
    Q.push({ cells: [['By product (totals per unit)', S.sec]] });
    Q.push({ h: 22, cells: ['Product', 'Unit', 'Ordered', 'Received'].map(function (h) { return [h, S.head]; }) });
    var p0 = Q.length + 1;
    m.byProduct.forEach(function (p, i) { var b = i % 2; Q.push({ cells: [[p.product, b ? S.txtB : S.txt], [p.unit, b ? S.txtB : S.txt], [p.ordered, b ? S.numB : S.num, 'n'], [p.received, b ? S.numB : S.num, 'n'], [p.product + ' (' + p.unit + ')', 0]] }); });
    var p1 = Q.length;
    // per-unit totals
    var byU = {};
    m.byProduct.forEach(function (p) { var t = byU[p.unit] || (byU[p.unit] = { o: 0, r: 0 }); t.o += p.ordered; t.r += p.received; });
    Object.keys(byU).sort().forEach(function (u) { Q.push({ cells: [['Total', S.gTxt], [u, S.gTxt], [byU[u].o, S.gNum, 'n'], [byU[u].r, S.gNum, 'n']] }); });
    Q.push({ h: 8, cells: [] });
    Q.push({ h: 44, cells: [['Units are never added together in the product table: each unit gets its own total row. The name table and its chart count every quantity regardless of unit, so read them as activity, not stock volume.', S.note]] });
    mg.push('A' + Q.length + ':D' + Q.length);

    var hasN = m.byName.length > 0, hasP = m.byProduct.length > 0;
    var pEnd = 3 + Math.max(16, Math.round(m.byProduct.length * 1.6) + 6);
    var drawing = '', charts = [];
    if (hasN || hasP) {
      var anchors = '', rels = '', id = 0;
      if (hasN) {
        id++;
        charts.push(chartXml('Usage by receiver', 'col', 'Summary!$A$' + n0 + ':$A$' + n1, m.byName.map(function (n) { return n.name; }), [
          { name: 'Ordered', color: GOLD, ref: 'Summary!$C$' + n0 + ':$C$' + n1, values: m.byName.map(function (n) { return n.ordered; }) },
          { name: 'Received', color: NAVY, ref: 'Summary!$D$' + n0 + ':$D$' + n1, values: m.byName.map(function (n) { return n.received; }) }]));
        anchors += anchor(id, 6, 3, 15, 21);
        rels += '<Relationship Id="rId' + id + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart' + id + '.xml"/>';
      }
      if (hasP) {
        id++;
        var top = hasN ? 23 : 3;
        charts.push(chartXml('Usage by product', 'bar', 'Summary!$E$' + p0 + ':$E$' + p1, m.byProduct.map(function (p) { return p.product + ' (' + p.unit + ')'; }), [
          { name: 'Ordered', color: GOLD, ref: 'Summary!$C$' + p0 + ':$C$' + p1, values: m.byProduct.map(function (p) { return p.ordered; }) },
          { name: 'Received', color: NAVY, ref: 'Summary!$D$' + p0 + ':$D$' + p1, values: m.byProduct.map(function (p) { return p.received; }) }]));
        anchors += anchor(id, 6, top, 15, top + pEnd - 3);
        rels += '<Relationship Id="rId' + id + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart' + id + '.xml"/>';
      }
      zip.file('xl/drawings/drawing1.xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' + anchors + '</xdr:wsDr>');
      zip.file('xl/drawings/_rels/drawing1.xml.rels', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' + rels + '</Relationships>');
      zip.file('xl/worksheets/_rels/sheet2.xml.rels', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/></Relationships>');
      charts.forEach(function (c, i) { zip.file('xl/charts/chart' + (i + 1) + '.xml', c); });
      drawing = '<drawing r:id="rId1"/>';
    }
    var s2 = sheetXml(Q, [{ w: 26 }, { w: 9 }, { w: 11 }, { w: 11 }, { w: 30, hidden: true }, { w: 3 }], mg, 0, drawing);

    zip.file('[Content_Types].xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>' +
      '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
      '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' +
      '<Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' +
      '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>' +
      (charts.length ? '<Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>' : '') +
      charts.map(function (c, i) { return '<Override PartName="/xl/charts/chart' + (i + 1) + '.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>'; }).join('') +
      '</Types>');
    zip.file('_rels/.rels', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>');
    zip.file('xl/workbook.xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><bookViews><workbookView/></bookViews><sheets><sheet name="Stock Report" sheetId="1" r:id="rId1"/><sheet name="Summary" sheetId="2" r:id="rId2"/></sheets></workbook>');
    zip.file('xl/_rels/workbook.xml.rels', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>');
    zip.file('xl/styles.xml', styles);
    zip.file('xl/worksheets/sheet1.xml', s1);
    zip.file('xl/worksheets/sheet2.xml', s2);
    return zip;
  }

  if (typeof module !== 'undefined') module.exports = build; else root.buildStockXlsx = build;
})(this);
