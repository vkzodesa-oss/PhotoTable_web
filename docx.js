/* PhotoTable (огляд останків): формування .docx (ZIP + OOXML) без сторонніх бібліотек.
   Формат відповідає зразку «Фототаблиця № тіла»: шапка з таблицею, фото з підписами
   «Фото №1: …» / «Фото №3, 4: …», таблиця підписів спеціалістів. */
(function (root) {
  'use strict';

  var NS_W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  var NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  var NS_WP = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing';
  var NS_A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
  var NS_PIC = 'http://schemas.openxmlformats.org/drawingml/2006/picture';
  var REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  var DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';

  var CM = 360000;            // EMU в 1 см
  var TEXT_W_CM = 16.5;       // ширина тексту на A4 (11906-1701-850 twips)
  var MAX_IMG_H_CM = 20.5;    // обмеження висоти фото

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
  function endWithPeriod(s) {
    return /[.!?…]$/.test(s) ? s : s + '.';
  }
  function formatDate(iso) {
    var m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(iso || '');
    return m ? m[3] + '.' + m[2] + '.' + m[1] : String(iso || '');
  }

  /* ---------- OOXML ---------- */
  function run(text, bold, size, underline) {
    size = size || 24;
    var b = bold ? '<w:b/><w:bCs/>' : '';
    var u = underline ? '<w:u w:val="single"/>' : '';
    return '<w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>' +
      b + u + '<w:sz w:val="' + size + '"/><w:szCs w:val="' + size + '"/></w:rPr>' +
      '<w:t xml:space="preserve">' + esc(text) + '</w:t></w:r>';
  }
  function para(runs, align, keepNext) {
    return '<w:p><w:pPr>' + (keepNext ? '<w:keepNext/>' : '') +
      '<w:spacing w:after="0"/><w:jc w:val="' + (align || 'both') + '"/></w:pPr>' + runs + '</w:p>';
  }
  function emptyPara() { return '<w:p><w:pPr><w:spacing w:after="0"/></w:pPr></w:p>'; }

  function imagePara(index, rId, ext, cx, cy) {
    var id = index + 1;
    return '<w:p><w:pPr><w:keepNext/><w:spacing w:after="0"/><w:jc w:val="center"/></w:pPr>' +
      '<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0">' +
      '<wp:extent cx="' + cx + '" cy="' + cy + '"/>' +
      '<wp:docPr id="' + id + '" name="Фото ' + id + '"/>' +
      '<wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr>' +
      '<a:graphic><a:graphicData uri="' + NS_PIC + '"><pic:pic>' +
      '<pic:nvPicPr><pic:cNvPr id="' + id + '" name="image' + id + '.' + ext + '"/><pic:cNvPicPr/></pic:nvPicPr>' +
      '<pic:blipFill><a:blip r:embed="' + rId + '"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>' +
      '<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="' + cx + '" cy="' + cy + '"/></a:xfrm>' +
      '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>' +
      '</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>';
  }

  var STYLES = DECL + '<w:styles xmlns:w="' + NS_W + '"><w:docDefaults><w:rPrDefault><w:rPr>' +
    '<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:eastAsia="Times New Roman" w:cs="Times New Roman"/>' +
    '<w:sz w:val="24"/><w:szCs w:val="24"/><w:lang w:val="uk-UA" w:eastAsia="en-US" w:bidi="ar-SA"/></w:rPr></w:rPrDefault>' +
    '<w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>' +
    '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style></w:styles>';

  var CT = DECL + '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
    '<Default Extension="xml" ContentType="application/xml"/>' +
    '<Default Extension="jpeg" ContentType="application/vnd.openxmlformats-officedocument.image/jpeg"/>' +
    '<Default Extension="png" ContentType="application/vnd.openxmlformats-officedocument.image/png"/>' +
    '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
    '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>';

  var ROOT_RELS = DECL + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    '<Relationship Id="rId1" Type="' + REL + '/officeDocument" Target="word/document.xml"/></Relationships>';

  /* Поля сторінки як у зразку: верх 1276, право 850, низ 1134, ліво 1701 */
  var SECT = '<w:sectPr>' +
    '<w:pgSz w:w="11906" w:h="16838"/>' +
    '<w:pgMar w:top="1276" w:right="850" w:bottom="1134" w:left="1701" w:header="708" w:footer="708" w:gutter="0"/>' +
    '<w:cols w:space="708"/></w:sectPr>';

  var TBL_BORDERS = '<w:tblBorders>' +
    ['top', 'left', 'bottom', 'right', 'insideH', 'insideV'].map(function (e) {
      return '<w:' + e + ' w:val="single" w:sz="4" w:space="0" w:color="auto"/>';
    }).join('') + '</w:tblBorders>';
  function tblPr(tblW, floating, tblpY) {
    var p = '<w:tblPr>';
    if (floating) {
      p += '<w:tblpPr w:leftFromText="180" w:rightFromText="180" w:vertAnchor="' + (tblpY === 1636 ? 'page' : 'text') +
        '" w:horzAnchor="margin"' + (tblpY === 1636 ? ' w:tblpXSpec="right"' : '') + ' w:tblpY="' + tblpY + '"/>';
    }
    p += '<w:tblW w:w="' + tblW + '" w:type="dxa"/>' + TBL_BORDERS +
      '<w:tblLook w:val="04A0" w:firstRow="1" w:lastRow="0" w:firstColumn="1" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/></w:tblPr>';
    return p;
  }
  function tc(w, inner) { return '<w:tc><w:tcPr><w:tcW w:w="' + w + '" w:type="dxa"/></w:tcPr>' + inner + '</w:tc>'; }
  function tr(h, cells) {
    return '<w:tr><w:trPr><w:trHeight w:val="' + h + '"/></w:trPr>' + cells + '</w:tr>';
  }

  /* Плаваюча таблиця шапки (правий верхній кут сторінки), як у зразку. */
  function headerTable(info) {
    var placeLines = String(info.place || '').split('\n').map(function (s) { return clean(s); }).filter(Boolean);
    if (!placeLines.length) placeLines = [''];

    var leftCell = tc(5005,
      para(run('Місце дослідження ', false, 28) + run(placeLines[0], false, 28, true) +
        run(' ___________________________', false, 24, true)) +
      emptyPara() +
      placeLines.slice(1).map(function (l) { return para(run(l, false, 28, true)); }).join('') +
      (placeLines.length > 1 ? emptyPara() : '')
    );

    var bodyNo = clean(info.bodyNo);
    var rightCell = tc(5480,
      emptyPara() +
      para(run('№ тіла (останків)', false, 28)) +
      para(run('№ ', false, 28) + run(bodyNo, true, 28)) +
      emptyPara()
    );

    var row2 = tr(628,
      tc(5005, para(run('Дата дослідження ', false, 28) + run(formatDate(info.date), true, 28))) +
      tc(5480, para(run('ФОТОТАБЛИЦЯ___________________________', false, 28, true)))
    );

    return '<w:tbl>' + tblPr(10485, true, 1636) +
      '<w:tblGrid><w:gridCol w:w="5005"/><w:gridCol w:w="5480"/></w:tblGrid>' +
      tr(1294, leftCell + rightCell) + row2 + '</w:tbl>';
  }

  /* Таблиця підписів наприкінці документа. */
  function signTable(specialists) {
    var head = tr(296,
      tc(2694, para(run('№', true, 28), 'center')) +
      tc(4531, para(run('ПІБ', true, 28), 'center')) +
      tc(2551, para(run('Підпис', true, 28), 'center'))
    );
    var rows = specialists.map(function (sp, i) {
      var words = clean(sp.name).split(/\s+/).filter(Boolean);
      var surname = words.length ? words[0] : '';
      var rest = words.slice(1).join(' ');
      var nameRuns = run(surname, true, 28) + (rest ? run(' ' + rest, false, 28) : '');
      var position = clean(sp.position);
      return tr(1366,
        tc(2694, para(run('Спеціаліст №' + (i + 1), true, 28))) +
        tc(4531, para(nameRuns) + emptyPara() + (position ? para(run(position, false, 24)) : '')) +
        tc(2551, emptyPara())
      );
    }).join('');
    return '<w:tbl>' + tblPr(9776, true, -99) +
      '<w:tblGrid><w:gridCol w:w="2694"/><w:gridCol w:w="4531"/><w:gridCol w:w="2551"/></w:tblGrid>' +
      head + rows + '</w:tbl>';
  }

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
      var name = encoder.encode(e.name);
      var crc = crc32(e.data), size = e.data.length;

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

  /* ---------- збирання документа ----------
     info:         { place, bodyNo, date: 'РРРР-ММ-ДД' }
     images:       [{ bytes: Uint8Array, ext: 'jpeg'|'png', w, h, caption, group }] — group=true: продовжує
                   попередню фотографію, спільний підпис («Фото №3, 4: …»)
     specialists:  [{ name, position }] */
  function buildDocx(info, images, specialists) {
    var body = [];
    body.push(para(run('Додаток до протоколу огляду', false, 24), 'right'));
    body.push(emptyPara());
    body.push(headerTable(info));

    /* нумерація фото та підписи груп */
    var n = images.length;
    var capLabel = new Array(n);
    for (var i = 0; i < n; i++) {
      if (images[i].group && i > 0) capLabel[i] = null; // у групі
      else {
        var nums = [String(i + 1)];
        var j = i + 1;
        while (j < n && images[j].group) { nums.push(String(j + 1)); j++; }
        capLabel[i] = 'Фото №' + nums.join(', ') + ':';
      }
    }

    var rels = [], media = [];
    var textW = Math.round(TEXT_W_CM * CM);
    var maxH = Math.round(MAX_IMG_H_CM * CM);
    images.forEach(function (im, i) {
      var w = im.w, h = im.h;
      if (!(w > 0) || !(h > 0)) { w = 4; h = 3; }
      var cy = Math.round(TEXT_W_CM * CM * h / w);
      var cx = textW;
      if (cy > maxH) { cy = maxH; cx = Math.round(cy * w / h); }
      var rId = 'rId' + (10 + i);
      var fileName = 'image' + (i + 1) + '.' + im.ext;
      var lastInGroup = (i === n - 1) || !images[i + 1].group;
      var hasCaption = !im.group;

      body.push(imagePara(i, rId, im.ext, cx, cy));
      if (hasCaption) {
        var text = clean(im.caption);
        var cap = capLabel[i] || ('Фото №' + (i + 1) + ':');
        cap = text ? endWithPeriod(cap + ' ' + text) : endWithPeriod(cap);
        body.push(para(run(cap, false, 24), 'center', !lastInGroup));
      }
      rels.push('<Relationship Id="' + rId + '" Type="' + REL + '/image" Target="media/' + fileName + '"/>');
      media.push({ name: 'word/media/' + fileName, data: im.bytes });
    });

    body.push(signTable(specialists || []));

    var documentXml = DECL + '<w:document xmlns:w="' + NS_W + '" xmlns:r="' + NS_R + '" xmlns:wp="' + NS_WP +
      '" xmlns:a="' + NS_A + '" xmlns:pic="' + NS_PIC + '"><w:body>' + body.join('') + SECT + '</w:body></w:document>';

    var docRels = DECL + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="' + REL + '/styles" Target="styles.xml"/>' +
      rels.join('') + '</Relationships>';

    var B = function (str) { return encoder.encode(str); };
    var entries = [
      { name: '[Content_Types].xml', data: B(CT) },
      { name: '_rels/.rels', data: B(ROOT_RELS) },
      { name: 'word/document.xml', data: B(documentXml) },
      { name: 'word/_rels/document.xml.rels', data: B(docRels) },
      { name: 'word/styles.xml', data: B(STYLES) }
    ].concat(media);
    return buildZip(entries);
  }

  var api = { buildDocx: buildDocx, formatDate: formatDate, clean: clean };
  root.PhotoTableDocx = api;
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
})(typeof window !== 'undefined' ? window : globalThis);
