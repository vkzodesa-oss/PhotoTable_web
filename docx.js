/* Фототаблиця (огляд останків): формування .docx (ZIP + OOXML) без сторонніх бібліотек.
   Оформлення — за зразком «фототаблиця_…docx». */
(function (root) {
  'use strict';

  var NS_W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  var NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  var NS_WP = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing';
  var NS_A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
  var NS_PIC = 'http://schemas.openxmlformats.org/drawingml/2006/picture';
  var REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  var DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';

  var CM = 360000;             // EMU в 1 см
  var TEXT_W_CM = 16.5;        // ширина тексту на A4 (поля 3 см / 1,5 см)
  var PAGE_BODY_CM = 25.0;     // робоча висота сторінки (A4, поля 2,25 / 2 см) з невеликим запасом
  var LINE_CM = 0.54;          // рядок тексту 12 пт
  var SPARE_CM = 0.8;          // запас, щоб підпис гарантовано лишався на сторінці
  var CHARS_PER_LINE = 74;     // обережна оцінка символів у рядку підпису

  var encoder = new TextEncoder();

  /* ---------- текст ---------- */
  function clean(s) {
    return String(s == null ? '' : s)
      .replace(/[\x00-\x08\x0B\x0C\x0E-\x1F]/g, '')
      .replace(/[\r\n]/g, ' ')
      .replace(/\s{2,}/g, ' ')
      .trim();
  }
  function esc(s) {
    return s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  }
  function endWithPeriod(s) { return /[.!?…]$/.test(s) ? s : s + '.'; }
  function capWord(w) {
    return w.split('-').map(function (p) {
      return p ? p.charAt(0).toUpperCase() + p.slice(1).toLowerCase() : p;
    }).join('-');
  }
  /* «родзін анатолій валерійович» -> «РОДЗІН Анатолій Валерійович» (прізвище першим) */
  function formatSurnameFirst(input) {
    var words = String(input == null ? '' : input).trim().split(/\s+/).filter(Boolean);
    if (!words.length) return '';
    return words.map(function (w, i) { return i === 0 ? w.toUpperCase() : capWord(w); }).join(' ');
  }
  /* '2026-09-30' -> '30.09.2026' */
  function formatDate(iso) {
    var m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(iso || '');
    return m ? m[3] + '.' + m[2] + '.' + m[1] : String(iso || '');
  }
  function cleanBodyNo(s) { return clean(s).replace(/^№\s*/, ''); }
  /* «Фототаблиця 20260916010004067.docx» */
  function makeFileName(bodyNo) {
    var no = cleanBodyNo(bodyNo).replace(/[\\\/:*?"<>|]/g, '-');
    return 'Фототаблиця' + (no ? ' ' + no : '') + '.docx';
  }
  function countLines(text) {
    var n = Array.from(text).length;
    return n === 0 ? 1 : Math.ceil(n / CHARS_PER_LINE);
  }
  /* «3» / «3, 4» / «7, 8, 9» */
  function photoLabel(start, count) {
    var a = [];
    for (var i = 0; i < count; i++) a.push(start + i);
    return a.join(', ');
  }

  /* ---------- OOXML: абзаци та runs ---------- */
  function run(text, o) {
    o = o || {};
    var size = o.size || 24;
    return '<w:r><w:rPr>' + (o.bold ? '<w:b/><w:bCs/>' : '') + (o.color ? '<w:color w:val="' + o.color + '"/>' : '') +
      '<w:sz w:val="' + size + '"/><w:szCs w:val="' + size + '"/>' + (o.u ? '<w:u w:val="single"/>' : '') +
      '</w:rPr><w:t xml:space="preserve">' + esc(text) + '</w:t></w:r>';
  }
  function para(runs, o) {
    o = o || {};
    return '<w:p><w:pPr>' + (o.keepNext ? '<w:keepNext/>' : '') + (o.pageBreak ? '<w:pageBreakBefore/>' : '') +
      '<w:spacing w:after="0"' + (o.single ? ' w:line="240" w:lineRule="auto"' : '') + '/>' +
      (o.align ? '<w:jc w:val="' + o.align + '"/>' : '') + '</w:pPr>' + (runs || '') + '</w:p>';
  }
  function emptyPara() { return para('', {}); }

  function imagePara(id, rId, ext, cx, cy, o) {
    return '<w:p><w:pPr><w:keepNext/>' + (o.pageBreak ? '<w:pageBreakBefore/>' : '') +
      '<w:spacing w:after="0"/><w:jc w:val="center"/></w:pPr>' +
      '<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0">' +
      '<wp:extent cx="' + cx + '" cy="' + cy + '"/>' +
      '<wp:docPr id="' + id + '" name="Рисунок ' + id + '"/>' +
      '<wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr>' +
      '<a:graphic><a:graphicData uri="' + NS_PIC + '"><pic:pic>' +
      '<pic:nvPicPr><pic:cNvPr id="' + id + '" name="image' + id + '.' + ext + '"/><pic:cNvPicPr/></pic:nvPicPr>' +
      '<pic:blipFill><a:blip r:embed="' + rId + '"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>' +
      '<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="' + cx + '" cy="' + cy + '"/></a:xfrm>' +
      '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>' +
      '</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>';
  }

  /* ---------- OOXML: таблиці ---------- */
  var BORDERS = '<w:tblBorders>' +
    '<w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/>' +
    '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/>' +
    '<w:insideH w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:insideV w:val="single" w:sz="4" w:space="0" w:color="auto"/>' +
    '</w:tblBorders>';
  function cell(width, paras) {
    return '<w:tc><w:tcPr><w:tcW w:w="' + width + '" w:type="dxa"/></w:tcPr>' + paras + '</w:tc>';
  }
  function row(height, cells) {
    return '<w:tr><w:trPr><w:cantSplit/><w:trHeight w:val="' + height + '"/></w:trPr>' + cells + '</w:tr>';
  }
  function cp(runs, align, keep) { return para(runs, { single: true, align: align, keepNext: keep }); }  // абзац у клітинці

  /* Шапка: місце дослідження / № тіла / дата / ФОТОТАБЛИЦЯ */
  function headerTable(info) {
    var inst = clean(info.institution), dept = clean(info.department);
    var no = cleanBodyNo(info.bodyNo), date = formatDate(info.date);
    var c1 = cp(run('Місце дослідження ', { size: 28 }) + (inst ? run(inst, { size: 28, u: true }) : '') +
      run('___________________________', { size: 24, u: true })) +
      cp('') + (dept ? cp(run(dept, { size: 28, u: true })) : cp('')) + cp('');
    var c2 = cp('') + cp(run('№ тіла (останків) ', { size: 28 })) + cp('') +
      cp(run('№ ', { size: 28 }) + run(no, { size: 28, bold: true, color: 'FF0000' }));
    var c3 = cp('') + cp(run('Дата дослідження ', { size: 28 }) + run(date, { size: 28, bold: true }));
    var c4 = cp('') + cp(run('ФОТОТАБЛИЦЯ', { size: 28, bold: true, u: true }) + run('_', { size: 24, bold: true, u: true }) +
      run('__________________________', { size: 22, bold: true })) + cp('');
    return '<w:tbl><w:tblPr><w:tblW w:w="10485" w:type="dxa"/><w:tblInd w:w="-1131" w:type="dxa"/>' + BORDERS +
      '<w:tblLook w:val="04A0"/></w:tblPr><w:tblGrid><w:gridCol w:w="5005"/><w:gridCol w:w="5480"/></w:tblGrid>' +
      row(1294, cell(5005, c1) + cell(5480, c2)) + row(628, cell(5005, c3) + cell(5480, c4)) + '</w:tbl>';
  }

  /* Оцінка висоти таблиці підписів, см */
  function signTableHeight(specs) {
    var h = 0.65;
    specs.forEach(function (sp, i) {
      var nameL = Math.max(1, Math.ceil(clean(sp.name).length / 30));
      var pos = clean(sp.position);
      var posL = pos ? Math.ceil(pos.length / 38) : 0;
      var rowMin = i === 0 ? 2.41 : 2.84;
      h += Math.max(rowMin, nameL * 0.62 + (pos ? 0.5 + posL * 0.47 : 0) + 0.15);
    });
    return h;
  }

  /* Таблиця підписів: № / ПІБ / Підпис (рядки не розриваються між сторінками) */
  function signTable(specialists) {
    var W = [2694, 4531, 2551];
    var n = specialists.length;
    var r = row(296, cell(W[0], cp(run('№', { size: 28, bold: true }), 'center', true)) +
      cell(W[1], cp(run('ПІБ', { size: 28, bold: true }), 'center', true)) +
      cell(W[2], cp(run('Підпис', { size: 28, bold: true }), 'center', true)));
    specialists.forEach(function (sp, i) {
      var keep = i < n - 1;
      var name = formatSurnameFirst(clean(sp.name));
      var words = name.split(' ');
      var nameRuns = name ? run(words[0], { size: 28, bold: true }) + (words.length > 1 ? run(' ' + words.slice(1).join(' '), { size: 28 }) : '') : '';
      var pos = clean(sp.position);
      var c2 = cp(nameRuns, null, keep) + (pos ? cp('', null, keep) + cp(run(pos, { size: 24 }), null, keep) : '');
      r += row(i === 0 ? 1366 : 1611,
        cell(W[0], cp(run('Спеціаліст №' + (i + 1), { size: 28, bold: true }), null, keep)) + cell(W[1], c2) + cell(W[2], cp('', null, keep)));
    });
    return '<w:tbl><w:tblPr><w:tblW w:w="9776" w:type="dxa"/>' + BORDERS + '<w:tblLook w:val="04A0"/></w:tblPr>' +
      '<w:tblGrid><w:gridCol w:w="2694"/><w:gridCol w:w="4531"/><w:gridCol w:w="2551"/></w:tblGrid>' + r + '</w:tbl>';
  }

  /* ---------- службові частини ---------- */
  var STYLES = DECL + '<w:styles xmlns:w="' + NS_W + '"><w:docDefaults><w:rPrDefault><w:rPr>' +
    '<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:eastAsia="Times New Roman" w:cs="Times New Roman"/>' +
    '<w:sz w:val="24"/><w:szCs w:val="24"/><w:lang w:val="uk-UA" w:eastAsia="en-US" w:bidi="ar-SA"/></w:rPr></w:rPrDefault>' +
    '<w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="259" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>' +
    '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style></w:styles>';
  var CT = DECL + '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
    '<Default Extension="xml" ContentType="application/xml"/><Default Extension="jpeg" ContentType="image/jpeg"/>' +
    '<Default Extension="png" ContentType="image/png"/>' +
    '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
    '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>';
  var ROOT_RELS = DECL + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    '<Relationship Id="rId1" Type="' + REL + '/officeDocument" Target="word/document.xml"/></Relationships>';
  var SECT = '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>' +
    '<w:pgMar w:top="1276" w:right="850" w:bottom="1134" w:left="1701" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr>';

  /* ---------- ZIP (без стиснення) ---------- */
  var CRC_TABLE = (function () {
    var t = new Uint32Array(256);
    for (var n = 0; n < 256; n++) {
      var c = n;
      for (var k = 0; k < 8; k++) c = (c & 1) ? (0xEDB88320 ^ (c >>> 1)) : (c >>> 1);
      t[n] = c >>> 0;
    }
    return t;
  })();
  function crc32(data) {
    var c = 0xFFFFFFFF;
    for (var i = 0; i < data.length; i++) c = CRC_TABLE[(c ^ data[i]) & 0xFF] ^ (c >>> 8);
    return (c ^ 0xFFFFFFFF) >>> 0;
  }
  function buildZip(entries) {
    var chunks = [], central = [], offset = 0, centralSize = 0;
    entries.forEach(function (e) {
      var name = encoder.encode(e.name), crc = crc32(e.data), size = e.data.length;
      var lh = new DataView(new ArrayBuffer(30));
      lh.setUint32(0, 0x04034b50, true); lh.setUint16(4, 20, true); lh.setUint16(6, 0x0800, true);
      lh.setUint16(8, 0, true); lh.setUint16(10, 0, true); lh.setUint16(12, 0x0021, true);
      lh.setUint32(14, crc, true); lh.setUint32(18, size, true); lh.setUint32(22, size, true);
      lh.setUint16(26, name.length, true); lh.setUint16(28, 0, true);
      chunks.push(new Uint8Array(lh.buffer), name, e.data);
      var cd = new DataView(new ArrayBuffer(46));
      cd.setUint32(0, 0x02014b50, true); cd.setUint16(4, 20, true); cd.setUint16(6, 20, true);
      cd.setUint16(8, 0x0800, true); cd.setUint16(10, 0, true); cd.setUint16(12, 0, true);
      cd.setUint16(14, 0x0021, true); cd.setUint32(16, crc, true); cd.setUint32(20, size, true);
      cd.setUint32(24, size, true); cd.setUint16(28, name.length, true); cd.setUint16(30, 0, true);
      cd.setUint16(32, 0, true); cd.setUint16(34, 0, true); cd.setUint16(36, 0, true);
      cd.setUint32(38, 0, true); cd.setUint32(42, offset, true);
      central.push(new Uint8Array(cd.buffer), name);
      centralSize += 46 + name.length;
      offset += 30 + name.length + size;
    });
    var end = new DataView(new ArrayBuffer(22));
    end.setUint32(0, 0x06054b50, true); end.setUint16(4, 0, true); end.setUint16(6, 0, true);
    end.setUint16(8, entries.length, true); end.setUint16(10, entries.length, true);
    end.setUint32(12, centralSize, true); end.setUint32(16, offset, true); end.setUint16(20, 0, true);
    var all = chunks.concat(central, [new Uint8Array(end.buffer)]);
    var total = 0; all.forEach(function (c) { total += c.length; });
    var out = new Uint8Array(total), pos = 0;
    all.forEach(function (c) { out.set(c, pos); pos += c.length; });
    return out;
  }

  /* ---------- розкладка по сторінках ----------
     Одиночні фото — по 2 на сторінці; кожна група (2+ фото під одним підписом) — на окремій сторінці. */
  function firstPageOverhead(info) {
    var inst = clean(info.institution), dept = clean(info.department);
    var l1 = Math.ceil((('Місце дослідження ' + inst).length + 27) / 36);
    var l2 = Math.max(1, Math.ceil(dept.length / 40));
    var row1 = Math.max(2.3, (l1 + 1 + l2 + 1) * 0.62);
    var row2 = 1.3;
    return LINE_CM * 2 + row1 + row2 + LINE_CM + 0.4;   // «Додаток…», порожній рядок, таблиця, порожній рядок
  }

  function buildPages(units) {
    var pages = [], cur = [];
    units.forEach(function (u) {
      if (u.photos.length === 1) {
        cur.push(u);
        if (cur.length === 2) { pages.push({ units: cur }); cur = []; }
      } else {
        if (cur.length) { pages.push({ units: cur }); cur = []; }
        pages.push({ units: [u] });
      }
    });
    if (cur.length) pages.push({ units: cur });
    return pages;
  }

  /* Висоти (см) для кожного фото на сторінці */
  function layoutHeights(page, isFirst, overhead0) {
    var overhead = isFirst ? overhead0 : 0;
    var us = page.units;
    var nat = us.map(function (u) {
      return u.photos.map(function (p) {
        var w = p.w, h = p.h;
        if (!(w > 0) || !(h > 0)) { w = 4; h = 3; }
        return TEXT_W_CM * h / w;
      });
    });
    if (us.length === 1 && us[0].photos.length > 1) {                    // група: фото одне під одним
      var avail = PAGE_BODY_CM - overhead - countLines(us[0].captionFull) * LINE_CM - SPARE_CM;
      var sum = nat[0].reduce(function (a, b) { return a + b; }, 0);
      var s = Math.min(1, avail / sum);
      return [nat[0].map(function (h) { return Math.max(2.5, h * s); })];
    }
    var cap = us.map(function (u) { return countLines(u.captionFull) * LINE_CM; });
    if (us.length === 1) {                                              // одне фото на сторінці
      var av1 = PAGE_BODY_CM - overhead - cap[0] - SPARE_CM;
      return [[Math.max(3, Math.min(nat[0][0], av1))]];
    }
    var av = PAGE_BODY_CM - overhead - cap[0] - cap[1] - LINE_CM - SPARE_CM;   // два фото + порожній рядок
    var half = av / 2, a = nat[0][0], b = nat[1][0], ha, hb;
    if (a <= half) { ha = a; hb = Math.min(b, av - a); }
    else if (b <= half) { hb = b; ha = Math.min(a, av - b); }
    else { ha = half; hb = half; }
    return [[Math.max(3, ha)], [Math.max(3, hb)]];
  }

  function usedHeight(page, heights, isFirst, overhead0) {
    var h = isFirst ? overhead0 : 0;
    page.units.forEach(function (u, ui) {
      heights[ui].forEach(function (x) { h += x; });
      h += countLines(u.captionFull) * LINE_CM;
    });
    if (page.units.length > 1) h += LINE_CM * (page.units.length - 1);
    return h;
  }

  /* ---------- збирання документа ----------
     info:  { bodyNo, date:'РРРР-ММ-ДД', institution, department, specialists:[{name, position}] }
     units: [{ caption, photos:[{ bytes:Uint8Array, ext, w, h }] }] */
  function buildDocx(info, units) {
    // нумерація фото та повні підписи
    var counter = 1;
    units.forEach(function (u) {
      var label = photoLabel(counter, u.photos.length);
      counter += u.photos.length;
      var text = clean(u.caption);
      u.captionFull = text ? endWithPeriod('Фото №' + label + ': ' + text) : 'Фото №' + label + '.';
    });

    var overhead0 = firstPageOverhead(info);
    var pages = buildPages(units);
    var body = [];
    body.push(para(run('Додаток до протоколу огляду'), { align: 'right' }));
    body.push(emptyPara());
    body.push(headerTable(info));
    body.push(emptyPara());

    var rels = [], media = [], imgId = 0, textW = Math.round(TEXT_W_CM * CM), lastUsed = 0;
    pages.forEach(function (page, pi) {
      var heights = layoutHeights(page, pi === 0, overhead0);
      lastUsed = usedHeight(page, heights, pi === 0, overhead0);
      page.units.forEach(function (u, ui) {
        u.photos.forEach(function (p, k) {
          imgId++;
          var rId = 'rId' + (10 + imgId), fileName = 'image' + imgId + '.' + p.ext;
          var w = p.w, h = p.h;
          if (!(w > 0) || !(h > 0)) { w = 4; h = 3; }
          var cy = Math.round(heights[ui][k] * CM), cx = Math.round(cy * w / h);
          if (cx > textW) { cx = textW; cy = Math.round(cx * h / w); }
          body.push(imagePara(imgId, rId, p.ext, cx, cy, { pageBreak: pi > 0 && ui === 0 && k === 0 }));
          rels.push('<Relationship Id="' + rId + '" Type="' + REL + '/image" Target="media/' + fileName + '"/>');
          media.push({ name: 'word/media/' + fileName, data: p.bytes });
        });
        body.push(para(run(u.captionFull), { align: 'center' }));
        if (ui < page.units.length - 1) body.push(emptyPara());
      });
    });

    var specs = (info.specialists || []).filter(function (s) { return clean(s.name) || clean(s.position); });
    if (specs.length) {
      // якщо таблиця підписів не вміщується під останньою сторінкою — переносимо її цілком
      var needBreak = pages.length > 0 && (lastUsed + LINE_CM + signTableHeight(specs) + SPARE_CM > PAGE_BODY_CM);
      body.push(needBreak ? para('', { pageBreak: true }) : emptyPara());
      body.push(signTable(specs));
    }
    body.push(emptyPara());

    var documentXml = DECL + '<w:document xmlns:w="' + NS_W + '" xmlns:r="' + NS_R + '" xmlns:wp="' + NS_WP +
      '" xmlns:a="' + NS_A + '" xmlns:pic="' + NS_PIC + '"><w:body>' + body.join('') + SECT + '</w:body></w:document>';
    var docRels = DECL + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="' + REL + '/styles" Target="styles.xml"/>' + rels.join('') + '</Relationships>';
    var B = function (s) { return encoder.encode(s); };
    return buildZip([
      { name: '[Content_Types].xml', data: B(CT) },
      { name: '_rels/.rels', data: B(ROOT_RELS) },
      { name: 'word/document.xml', data: B(documentXml) },
      { name: 'word/_rels/document.xml.rels', data: B(docRels) },
      { name: 'word/styles.xml', data: B(STYLES) }
    ].concat(media));
  }

  var api = {
    buildDocx: buildDocx, formatSurnameFirst: formatSurnameFirst, formatDate: formatDate,
    makeFileName: makeFileName, photoLabel: photoLabel, cleanBodyNo: cleanBodyNo
  };
  root.RemainsDocx = api;
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
})(typeof window !== 'undefined' ? window : globalThis);
