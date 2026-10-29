(function (root) {
'use strict';
var NS_W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
var REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
var DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
var CM = 360000; var TEXT_W_CM = 16.5;
var encoder = new TextEncoder();

function esc(s) { return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;'); }
function run(text, bold, size) { size=size||24; return `<w:r><w:rPr>${bold?'<w:b/>':''}<w:sz w:val="${size}"/></w:rPr><w:t xml:space="preserve">${esc(text)}</w:t></w:r>`; }
function para(runs, align) { return `<w:p><w:pPr><w:jc w:val="${align||'both'}"/></w:pPr>${runs}</w:p>`; }

function buildRemainsDocx(info, images) {
  var body = [];
  // Шапка
  body.push(para(run('ДОДАТОК ДО ПРОТОКОЛУ ОГЛЯДУ МІСЦЯ ПОДІЇ (ОСТАНКИ)', true, 28), 'center'));
  body.push(para(run(`Місце дослідження: ${info.place || '_______________'}   № Тіла: ${info.bodyNumber}`, false, 24), 'left'));
  body.push(para(run(`Дата дослідження: ${info.date || '_______________'}`, false, 24), 'left'));
  body.push(para(run('ФОТОТАБЛИЦЯ', true, 28), 'center'));
  body.push('<w:p/>');

  // Обробка фотографій та груп
  var currentGroup = null;
  var groupIndices = [];
  
  images.forEach(function(im, idx) {
    // Якщо це кінець групи або одинарне фото
    var isLastInGroup = (idx === images.length - 1) || (images[idx+1].group !== im.group);
    var isFirstInGroup = (idx === 0) || (images[idx-1].group !== im.group);

    // Додаємо картинку
    var rId = 'rId' + (idx + 10);
    var cx = Math.round(TEXT_W_CM * CM * 0.45); // Ширина 45% для двох в ряд
    var cy = Math.round(cx * im.h / im.w);
    if (im.group && !isFirstInGroup) { /* Друге фото в парі - вирівнювання праворуч спрощене */ }
    
    body.push(`<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:drawing><wp:inline><wp:extent cx="${cx}" cy="${cy}"/><wp:docPr id="${idx+1}" name="Рис${idx+1}"/><a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:nvPicPr><pic:cNvPr id="${idx+1}" name="img${idx+1}"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="${rId}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>`);

    // Підпис з урахуванням груп
    if (isLastInGroup) {
      var capNums = "";
      if (im.group) {
        var startIdx = idx; while(startIdx > 0 && images[startIdx-1].group === im.group) startIdx--;
        var endIdx = idx; while(endIdx < images.length-1 && images[endIdx+1].group === im.group) endIdx++;
        if (startIdx === endIdx) capNums = `Фото № ${startIdx + 1}:`;
        else capNums = `Фото № ${startIdx + 1}-${endIdx + 1}:`;
      } else {
        capNums = `Фото № ${idx + 1}:`;
      }
      var fullCaption = capNums + " " + (im.caption || "");
      body.push(para(run(fullCaption, false, 24), 'left'));
    }
  });

  // Підписи спеціалістів
  body.push('<w:p/>');
  body.push(para(run(`Спеціаліст №1: _________________________ ${info.spec1 || ''}`, false, 24), 'left'));
  body.push(para(run(`Спеціаліст №2: _________________________ ${info.spec2 || ''}`, false, 24), 'left'));

  var documentXml = DECL + `<w:document xmlns:w="${NS_W}" xmlns:r="${REL}" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><w:body>${body.join('')}</w:body></w:document>`;
  
  var rels = ['<Relationship Id="rId1" Type="'+REL+'/styles" Target="styles.xml"/>'];
  var media = [];
  images.forEach(function(im, i) {
    rels.push(`<Relationship Id="rId${10+i}" Type="${REL}/image" Target="media/image${i+1}.${im.ext}"/>`);
    media.push({ name: 'word/media/image'+(i+1)+'.'+im.ext, data: im.bytes });
  });
  var docRels = DECL + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' + rels.join('') + '</Relationships>';
  
  var entries = [
    { name: '[Content_Types].xml', data: encoder.encode(DECL+'<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="jpeg" ContentType="image/jpeg"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>') },
    { name: '_rels/.rels', data: encoder.encode(DECL+'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="'+REL+'/officeDocument" Target="word/document.xml"/></Relationships>') },
    { name: 'word/document.xml', data: encoder.encode(documentXml) },
    { name: 'word/_rels/document.xml.rels', data: encoder.encode(docRels) },
    { name: 'word/styles.xml', data: encoder.encode(DECL+'<w:styles xmlns:w="'+NS_W+'"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="24"/></w:rPr></w:rPrDefault></w:docDefaults></w:styles>') }
  ].concat(media);
  
  return buildZip(entries);
}

/* ZIP Builder (Стиснення не використовується для швидкодії) */
var CRC_TABLE = new Uint32Array(256); for(var n=0;n<256;n++){var c=n;for(var k=0;k<8;k++)c=(c&1)?(0xEDB88320^(c>>>1)):(c>>>1);CRC_TABLE[n]=c>>>0;}
function crc32(data){var c=0xFFFFFFFF;for(var i=0;i<data.length;i++)c=CRC_TABLE[(c^data[i])&0xFF]^(c>>>8);return(c^0xFFFFFFFF)>>>0;}
function buildZip(entries){var chunks=[],central=[],offset=0,centralSize=0;entries.forEach(function(e){var name=encoder.encode(e.name);var crc=crc32(e.data),size=e.data.length;var lh=new DataView(new ArrayBuffer(30));lh.setUint32(0,0x04034b50,true);lh.setUint16(4,20,true);lh.setUint16(6,0x0800,true);lh.setUint16(8,0,true);lh.setUint16(10,0,true);lh.setUint16(12,0x0021,true);lh.setUint32(14,crc,true);lh.setUint32(18,size,true);lh.setUint32(22,size,true);lh.setUint16(26,name.length,true);lh.setUint16(28,0,true);chunks.push(new Uint8Array(lh.buffer),name,e.data);var cd=new DataView(new ArrayBuffer(46));cd.setUint32(0,0x02014b50,true);cd.setUint16(4,20,true);cd.setUint16(6,20,true);cd.setUint16(8,0x0800,true);cd.setUint16(10,0,true);cd.setUint16(12,0,true);cd.setUint16(14,0x0021,true);cd.setUint32(16,crc,true);cd.setUint32(20,size,true);cd.setUint32(24,size,true);cd.setUint16(28,name.length,true);cd.setUint16(30,0,true);cd.setUint16(32,0,true);cd.setUint16(34,0,true);cd.setUint16(36,0,true);cd.setUint32(38,0,true);cd.setUint32(42,offset,true);central.push(new Uint8Array(cd.buffer),name);centralSize+=46+name.length;offset+=30+name.length+size;});var end=new DataView(new ArrayBuffer(22));end.setUint32(0,0x06054b50,true);end.setUint16(4,0,true);end.setUint16(6,0,true);end.setUint16(8,entries.length,true);end.setUint16(10,entries.length,true);end.setUint32(12,centralSize,true);end.setUint32(16,offset,true);end.setUint16(20,0,true);var all=chunks.concat(central,[new Uint8Array(end.buffer)]);var total=0;all.forEach(function(c){total+=c.length;});var out=new Uint8Array(total),pos=0;all.forEach(function(c){out.set(c,pos);pos+=c.length;});return out;}

root.CrimDocx = { buildRemainsDocx: buildRemainsDocx };
})(typeof window !== 'undefined' ? window : globalThis);