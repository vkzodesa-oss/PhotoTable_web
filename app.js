/* CrimTable · Останки: фототаблиця до протоколу огляду. Усе працює на пристрої, без інтернету. */
(function () {
  'use strict';

  var D = window.RemainsDocx;
  var MIME = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';
  var MAX_SIDE = 2200;
  var QUALITY = 0.88;
  var SETTINGS_KEY = 'remains.settings.v1';
  var DB_NAME = 'remains-table';
  var DEFAULT_INST = 'ДСУ «Одеське ОБ СМЕ»';
  var DEFAULT_DEPT = 'Центральне міжрайонне відділення №3';

  function svg(path) { return '<svg viewBox="0 0 24 24" width="22" height="22"><path fill="currentColor" d="' + path + '"/></svg>'; }
  var ICON_UP = svg('M4 12l1.4 1.4L11 7.8V20h2V7.8l5.6 5.6L20 12l-8-8-8 8z');
  var ICON_DOWN = svg('M20 12l-1.4-1.4L13 16.2V4h-2v12.2l-5.6-5.6L4 12l8 8 8-8z');
  var ICON_TRASH = svg('M6 19a2 2 0 0 0 2 2h8a2 2 0 0 0 2-2V7H6v12zM19 4h-3.5l-1-1h-5l-1 1H5v2h14V4z');
  var ICON_CLOSE = svg('M19 6.4 17.6 5 12 10.6 6.4 5 5 6.4 10.6 12 5 17.6 6.4 19 12 13.4 17.6 19 19 17.6 13.4 12z');
  var ICON_LEFT = svg('M15.41 7.41 14 6l-6 6 6 6 1.41-1.41L10.83 12z');
  var ICON_RIGHT = svg('M10 6 8.59 7.41 13.17 12l-4.58 4.59L10 18l6-6z');

  function $(id) { return document.getElementById(id); }
  var els = {
    bodyNo: $('bodyNo'), date: $('date'), institution: $('institution'), department: $('department'),
    specList: $('specList'), count: $('count'), unitList: $('unitList'), cam: $('camInput'), gal: $('galInput')
  };
  var state = { units: [], specialists: [] };

  function el(tag, cls, text) {
    var e = document.createElement(tag);
    if (cls) e.className = cls;
    if (text != null) e.textContent = text;
    return e;
  }
  function uid() { return Date.now().toString(36) + Math.random().toString(36).slice(2, 7); }
  function pad(n) { return n < 10 ? '0' + n : '' + n; }
  function todayISO() { var d = new Date(); return d.getFullYear() + '-' + pad(d.getMonth() + 1) + '-' + pad(d.getDate()); }
  function totalPhotos() { return state.units.reduce(function (s, u) { return s + u.photos.length; }, 0); }

  /* ---------- сховище ---------- */
  var idb = null;
  function openDb() {
    return new Promise(function (resolve) {
      if (!window.indexedDB) { resolve(null); return; }
      try {
        var req = indexedDB.open(DB_NAME, 1);
        req.onupgradeneeded = function () { req.result.createObjectStore('kv'); req.result.createObjectStore('photos'); };
        req.onsuccess = function () { resolve(req.result); };
        req.onerror = function () { resolve(null); };
      } catch (e) { resolve(null); }
    });
  }
  function dbReq(store, mode, fn) {
    return new Promise(function (resolve, reject) {
      if (!idb) { resolve(undefined); return; }
      try {
        var tx = idb.transaction(store, mode);
        var r = fn(tx.objectStore(store));
        tx.oncomplete = function () { resolve(r ? r.result : undefined); };
        tx.onerror = tx.onabort = function () { reject(tx.error); };
      } catch (e) { reject(e); }
    });
  }
  function putPhoto(p) { return dbReq('photos', 'readwrite', function (os) { return os.put({ buf: p.buf, w: p.w, h: p.h }, p.id); }).catch(function () {}); }
  function delPhoto(p) { return dbReq('photos', 'readwrite', function (os) { return os.delete(p.id); }).catch(function () {}); }

  function loadSettings() {
    var s = null;
    try { s = JSON.parse(localStorage.getItem(SETTINGS_KEY) || 'null'); } catch (e) { s = null; }
    els.institution.value = s && typeof s.institution === 'string' ? s.institution : DEFAULT_INST;
    els.department.value = s && typeof s.department === 'string' ? s.department : DEFAULT_DEPT;
    var list = s && Array.isArray(s.specialists) && s.specialists.length ? s.specialists : [{}, {}];
    state.specialists = list.map(function (x) { return { name: String((x && x.name) || ''), position: String((x && x.position) || '') }; });
  }
  function saveSettings() {
    try {
      localStorage.setItem(SETTINGS_KEY, JSON.stringify({
        institution: els.institution.value, department: els.department.value, specialists: state.specialists
      }));
    } catch (e) { /* ignore */ }
  }
  var saveTimer = null;
  function saveSoon() { clearTimeout(saveTimer); saveTimer = setTimeout(saveNow, 300); }
  function saveNow() {
    clearTimeout(saveTimer);
    saveSettings();
    var meta = {
      bodyNo: els.bodyNo.value, date: els.date.value,
      units: state.units.map(function (u) {
        return { id: u.id, caption: u.caption, layout: u.layout || 'auto', photoIds: u.photos.map(function (p) { return p.id; }) };
      })
    };
    return dbReq('kv', 'readwrite', function (os) { return os.put(meta, 'draft'); }).catch(function () {});
  }

  /* ---------- допоміжне UI ---------- */
  function setBusy(on, text) {
    $('overlay').hidden = !on;
    $('overlayText').textContent = text || '';
    $('btnFinish').disabled = on;
  }
  function closeModal() { $('modal').hidden = true; }
  function openModal(title, text, buttons) {
    $('modalTitle').textContent = title;
    $('modalText').textContent = text;
    var box = $('modalBtns');
    box.innerHTML = '';
    buttons.forEach(function (b) {
      var btn = document.createElement('button');
      btn.type = 'button';
      btn.className = 'btn ' + (b.cls || 'outline');
      btn.textContent = b.label;
      btn.onclick = function () { if (!b.keep) closeModal(); if (b.onClick) b.onClick(); };
      box.appendChild(btn);
    });
    $('modal').hidden = false;
  }
  function showMessage(title, text) { openModal(title, text, [{ label: 'OK', cls: 'primary' }]); }
  function confirmBox(text) {
    return new Promise(function (resolve) {
      openModal('', text, [
        { label: 'Так', cls: 'primary', onClick: function () { resolve(true); } },
        { label: 'Ні', onClick: function () { resolve(false); } }
      ]);
    });
  }
  function iconBtn(svgHtml, label, disabled, onClick) {
    var b = document.createElement('button');
    b.type = 'button';
    b.className = 'icon-btn';
    b.innerHTML = svgHtml;
    b.setAttribute('aria-label', label);
    b.title = label;
    b.disabled = !!disabled;
    b.onclick = onClick;
    return b;
  }

  /* ---------- фото ---------- */
  function loadImage(url) {
    return new Promise(function (resolve, reject) {
      var img = new Image();
      img.onload = function () { resolve(img); };
      img.onerror = function () { reject(new Error('не вдалося прочитати зображення')); };
      img.src = url;
    });
  }
  function blobToBuf(blob) {
    if (blob.arrayBuffer) return blob.arrayBuffer();
    return new Promise(function (resolve, reject) {
      var fr = new FileReader();
      fr.onload = function () { resolve(fr.result); };
      fr.onerror = function () { reject(fr.error); };
      fr.readAsArrayBuffer(blob);
    });
  }
  async function processImage(file) {
    var url = URL.createObjectURL(file);
    try {
      var img = await loadImage(url);
      var w0 = img.naturalWidth, h0 = img.naturalHeight;
      if (!w0 || !h0) throw new Error('порожнє зображення');
      var k = Math.min(1, MAX_SIDE / Math.max(w0, h0));
      var w = Math.max(1, Math.round(w0 * k)), h = Math.max(1, Math.round(h0 * k));
      var c = document.createElement('canvas');
      c.width = w; c.height = h;
      var ctx = c.getContext('2d');
      ctx.fillStyle = '#fff';
      ctx.fillRect(0, 0, w, h);
      ctx.drawImage(img, 0, 0, w, h);
      var blob = await new Promise(function (r) { c.toBlob(r, 'image/jpeg', QUALITY); });
      c.width = 1; c.height = 1;
      if (!blob) throw new Error('не вдалося стиснути зображення');
      return { buf: await blobToBuf(blob), w: w, h: h };
    } finally { URL.revokeObjectURL(url); }
  }
  function photoUrl(buf) { return URL.createObjectURL(new Blob([buf], { type: 'image/jpeg' })); }

  /* Екран «Порядок фото»: торкайтесь фото в потрібному порядку. Повертає масив або null. */
  /* Екран «Порядок фото»: великі превью (цілі фото), прокрутка, збільшення, вибір розміру сітки.
     Повертає масив у вибраному порядку або null. */
  function chooseOrder(pending) {
    return new Promise(function (resolve) {
      var order = [];
      var cols = 2;
      try { cols = parseInt(localStorage.getItem('orderCols'), 10) || 2; } catch (e) { cols = 2; }
      if (cols < 1 || cols > 3) cols = 2;
      var ZOOM_SVG = '<svg viewBox="0 0 24 24" width="20" height="20"><path fill="currentColor" d="M15.5 14h-.79l-.28-.27A6.47 6.47 0 0 0 16 9.5 6.5 6.5 0 1 0 9.5 16c1.61 0 3.09-.59 4.23-1.57l.27.28v.79l5 4.99L20.49 19l-4.99-5zm-6 0C7.01 14 5 11.99 5 9.5S7.01 5 9.5 5 14 7.01 14 9.5 11.99 14 9.5 14zm.5-7H9v2H7v1h2v2h1v-2h2V9h-2V7z"/></svg>';

      var page = document.createElement('div');
      page.className = 'order-page';

      var head = document.createElement('div');
      head.className = 'order-head';
      var title = document.createElement('h3');
      title.textContent = 'Порядок фото';
      var size = document.createElement('button');
      size.type = 'button';
      size.className = 'link-btn';
      var reset = document.createElement('button');
      reset.type = 'button';
      reset.className = 'link-btn';
      reset.textContent = 'Скинути';

      var hint = document.createElement('p');
      hint.className = 'order-hint';
      hint.textContent = 'Торкайтесь фото в тому порядку, у якому вони мають бути в таблиці. Повторне торкання скасовує вибір. Лупа збільшує фото, кнопка «в ряд» змінює розмір.';

      var grid = document.createElement('div');
      grid.className = 'order-grid';
      var cells = [];

      function toggle(i) {
        var k = order.indexOf(i);
        if (k >= 0) order.splice(k, 1); else order.push(i);
        update();
      }

      /* повноекранний перегляд одного фото */
      function openPreview(start) {
        var idx = start;
        var ov = document.createElement('div');
        ov.className = 'order-preview';
        var img = document.createElement('img');
        img.alt = '';
        var cap = document.createElement('div');
        cap.className = 'order-preview-cap';
        var bar = document.createElement('div');
        bar.className = 'order-preview-bar';
        var prev = document.createElement('button'); prev.type = 'button'; prev.className = 'btn outline'; prev.textContent = '‹ Назад';
        var next = document.createElement('button'); next.type = 'button'; next.className = 'btn outline'; next.textContent = 'Далі ›';
        var pick = document.createElement('button'); pick.type = 'button'; pick.className = 'btn primary';
        var close = document.createElement('button'); close.type = 'button'; close.className = 'order-preview-close'; close.setAttribute('aria-label', 'Закрити'); close.textContent = '✕';
        function show() {
          img.src = pending[idx].url;
          var k = order.indexOf(idx);
          cap.textContent = 'Фото ' + (idx + 1) + ' з ' + pending.length + (k >= 0 ? '  ·  обрано під № ' + (k + 1) : '');
          pick.textContent = k >= 0 ? 'Зняти вибір' : 'Обрати';
          prev.disabled = idx === 0;
          next.disabled = idx === pending.length - 1;
        }
        prev.addEventListener('click', function () { if (idx > 0) { idx--; show(); } });
        next.addEventListener('click', function () { if (idx < pending.length - 1) { idx++; show(); } });
        pick.addEventListener('click', function () { toggle(idx); show(); });
        close.addEventListener('click', function () { if (ov.parentNode) ov.parentNode.removeChild(ov); });
        bar.appendChild(prev); bar.appendChild(pick); bar.appendChild(next);
        ov.appendChild(close); ov.appendChild(img); ov.appendChild(cap); ov.appendChild(bar);
        page.appendChild(ov);
        show();
      }

      pending.forEach(function (p, i) {
        var cell = document.createElement('div');
        cell.className = 'order-cell';
        cell.setAttribute('role', 'button');
        cell.tabIndex = 0;
        var im = document.createElement('img');
        im.src = p.url;
        im.alt = '';
        var badge = document.createElement('span');
        badge.className = 'order-badge';
        var num = document.createElement('span');
        num.className = 'order-num';
        num.textContent = String(i + 1);
        var zoom = document.createElement('button');
        zoom.type = 'button';
        zoom.className = 'order-zoom';
        zoom.setAttribute('aria-label', 'Збільшити');
        zoom.innerHTML = ZOOM_SVG;
        zoom.addEventListener('click', function (e) { e.stopPropagation(); openPreview(i); });
        cell.addEventListener('click', function () { toggle(i); });
        cell.addEventListener('keydown', function (e) { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); toggle(i); } });
        cell.appendChild(im); cell.appendChild(badge); cell.appendChild(num); cell.appendChild(zoom);
        grid.appendChild(cell);
        cells.push({ cell: cell, badge: badge });
      });

      var foot = document.createElement('div');
      foot.className = 'order-foot';
      var ok = document.createElement('button');
      ok.type = 'button';
      ok.className = 'btn primary';
      foot.appendChild(ok);

      function applyCols() {
        grid.style.setProperty('--cols', String(cols));
        size.textContent = cols === 1 ? '1 в ряд' : cols + ' в ряд';
      }
      function update() {
        cells.forEach(function (c, i) {
          var k = order.indexOf(i);
          c.cell.classList.toggle('sel', k >= 0);
          c.badge.textContent = k >= 0 ? String(k + 1) : '';
        });
        ok.textContent = 'Додати (' + order.length + ' з ' + pending.length + ' обрано)';
        reset.disabled = !order.length;
      }
      function done(result) {
        if (page.parentNode) page.parentNode.removeChild(page);
        resolve(result);
      }

      head.appendChild(iconBtn(ICON_CLOSE, 'Скасувати', false, function () { done(null); }));
      head.appendChild(title);
      head.appendChild(size);
      head.appendChild(reset);
      size.addEventListener('click', function () {
        cols = cols === 2 ? 1 : (cols === 1 ? 3 : 2);
        try { localStorage.setItem('orderCols', String(cols)); } catch (e) { /* ignore */ }
        applyCols();
      });
      reset.addEventListener('click', function () { order = []; update(); });
      ok.addEventListener('click', function () {
        var rest = [];
        pending.forEach(function (_, i) { if (order.indexOf(i) < 0) rest.push(i); });
        done(order.concat(rest).map(function (i) { return pending[i]; }));
      });

      page.appendChild(head);
      page.appendChild(hint);
      page.appendChild(grid);
      page.appendChild(foot);
      document.body.appendChild(page);
      applyCols();
      update();
    });
  }

  async function addFiles(files, focusLast, askOrder) {
    if (!files.length) return;
    setBusy(true, 'Обробка фото…');
    var pending = [], errors = [];
    for (var i = 0; i < files.length; i++) {
      try {
        var r = await processImage(files[i]);
        pending.push({ id: uid(), buf: r.buf, w: r.w, h: r.h, url: photoUrl(r.buf) });
      } catch (e) {
        errors.push((files[i].name || 'фото') + ': ' + (e && e.message ? e.message : e));
      }
    }
    setBusy(false);
    var ordered = pending;
    if (askOrder && pending.length > 1) {
      ordered = await chooseOrder(pending);
      if (!ordered) {
        pending.forEach(function (p) { URL.revokeObjectURL(p.url); });
        if (errors.length) showMessage('Помилка', 'Не вдалося обробити: ' + errors.join('; '));
        return;
      }
    }
    for (var j = 0; j < ordered.length; j++) {
      await putPhoto(ordered[j]);
      state.units.push({ id: uid(), caption: '', photos: [ordered[j]] });
    }
    saveNow();
    render();
    if (errors.length) showMessage('Помилка', 'Не вдалося обробити: ' + errors.join('; '));
    if (focusLast && ordered.length) {
      var last = els.unitList.querySelector('.unit:last-child textarea');
      if (last) { last.focus(); last.scrollIntoView({ block: 'center', behavior: 'smooth' }); }
    }
  }

  /* ---------- дії над одиницями (фото / група) ---------- */
  function moveUnit(ui, d) {
    var t = ui + d;
    if (t < 0 || t >= state.units.length) return;
    var u = state.units.splice(ui, 1)[0];
    state.units.splice(t, 0, u);
    saveNow(); render();
  }
  async function removeUnit(ui) {
    var u = state.units[ui];
    var q = u.photos.length > 1 ? 'Видалити групу з ' + u.photos.length + ' фото?' : 'Видалити це фото?';
    if (!(await confirmBox(q))) return;
    state.units.splice(ui, 1);
    u.photos.forEach(function (p) { URL.revokeObjectURL(p.url); delPhoto(p); });
    saveNow(); render();
  }
  function mergeNext(ui) {
    var a = state.units[ui], b = state.units[ui + 1];
    if (!a || !b) return;
    a.photos = a.photos.concat(b.photos);
    if (!a.caption.trim() && b.caption) a.caption = b.caption;
    state.units.splice(ui + 1, 1);
    saveNow(); render();
  }
  function splitUnit(ui) {
    var u = state.units[ui];
    if (!u || u.photos.length < 2) return;
    var singles = u.photos.map(function (p, i) { return { id: i === 0 ? u.id : uid(), caption: i === 0 ? u.caption : '', photos: [p] }; });
    Array.prototype.splice.apply(state.units, [ui, 1].concat(singles));
    saveNow(); render();
  }
  function movePhotoIn(ui, pi, d) {
    var arr = state.units[ui].photos, t = pi + d;
    if (t < 0 || t >= arr.length) return;
    var p = arr.splice(pi, 1)[0];
    arr.splice(t, 0, p);
    saveNow(); render();
  }
  async function removePhotoIn(ui, pi) {
    if (!(await confirmBox('Видалити це фото з групи?'))) return;
    var u = state.units[ui], p = u.photos.splice(pi, 1)[0];
    URL.revokeObjectURL(p.url); delPhoto(p);
    if (!u.photos.length) state.units.splice(ui, 1);
    saveNow(); render();
  }
  function insertBodyNo(ta, u) {
    var no = D.cleanBodyNo(els.bodyNo.value);
    if (!no) { showMessage('Помилка', 'Спочатку введіть № тіла (останків) вгорі форми.'); return; }
    var ins = '№' + no, s = ta.selectionStart, e = ta.selectionEnd;
    if (typeof s !== 'number') { s = e = ta.value.length; }
    ta.value = ta.value.slice(0, s) + ins + ta.value.slice(e);
    var pos = s + ins.length;
    ta.focus();
    try { ta.setSelectionRange(pos, pos); } catch (x) { /* ignore */ }
    u.caption = ta.value;
    saveSoon();
  }

  function render() {
    els.count.textContent = totalPhotos();
    els.unitList.innerHTML = '';
    var counter = 1;
    state.units.forEach(function (u, ui) {
      var n = u.photos.length;
      var label = D.photoLabel(counter, n);
      counter += n;

      var card = el('div', 'unit');
      card.appendChild(el('div', 'utitle', 'Фото №' + label));

      var thumbs = el('div', 'uthumbs');
      u.photos.forEach(function (p, pi) {
        var g = el('div', 'gthumb');
        var img = document.createElement('img');
        img.src = p.url; img.alt = '';
        g.appendChild(img);
        if (n > 1) {
          var tools = el('div', 'gtools');
          tools.appendChild(iconBtn(ICON_LEFT, 'Раніше', pi === 0, function () { movePhotoIn(ui, pi, -1); }));
          tools.appendChild(iconBtn(ICON_CLOSE, 'Видалити фото', false, function () { removePhotoIn(ui, pi); }));
          tools.appendChild(iconBtn(ICON_RIGHT, 'Пізніше', pi === n - 1, function () { movePhotoIn(ui, pi, 1); }));
          g.appendChild(tools);
        }
        thumbs.appendChild(g);
      });
      card.appendChild(thumbs);

      var ctrls = el('div', 'ucontrols');
      ctrls.appendChild(iconBtn(ICON_UP, 'Вище', ui === 0, function () { moveUnit(ui, -1); }));
      ctrls.appendChild(iconBtn(ICON_DOWN, 'Нижче', ui === state.units.length - 1, function () { moveUnit(ui, 1); }));
      ctrls.appendChild(iconBtn(ICON_TRASH, n > 1 ? 'Видалити групу' : 'Видалити', false, function () { removeUnit(ui); }));
      card.appendChild(ctrls);

      var pills = el('div', 'upills');
      if (ui < state.units.length - 1) {
        var m = el('button', 'pill', 'Об’єднати з наступним');
        m.type = 'button';
        m.onclick = function () { mergeNext(ui); };
        pills.appendChild(m);
      }
      if (n > 1) {
        var s = el('button', 'pill', 'Роз’єднати');
        s.type = 'button';
        s.onclick = function () { splitUnit(ui); };
        pills.appendChild(s);
      }
      if (pills.childNodes.length) card.appendChild(pills);

      if (n > 1) {
        var lf = el('label', 'field');
        lf.appendChild(el('span', null, 'Розташування фото в групі'));
        var sel = document.createElement('select');
        [['auto', 'Разом (колаж, автоматично)'], ['row', 'В один ряд'], ['column', 'Стовпчиком']].forEach(function (o) {
          var op = document.createElement('option');
          op.value = o[0];
          op.textContent = o[1];
          sel.appendChild(op);
        });
        sel.value = u.layout || 'auto';
        sel.addEventListener('change', function () { u.layout = sel.value; saveSoon(); });
        lf.appendChild(sel);
        card.appendChild(lf);
      }

      var head = el('div', 'cap-head');
      head.appendChild(el('span', null, 'Підпис до фото №' + label));
      var chip = el('button', 'pill', 'Вставити № тіла');
      chip.type = 'button';
      head.appendChild(chip);
      var ta = document.createElement('textarea');
      ta.rows = 3;
      ta.value = u.caption;
      ta.placeholder = 'напр. Загальний вигляд першого мішка для транспортування тіл';
      ta.addEventListener('input', function () { u.caption = ta.value; saveSoon(); });
      chip.onclick = function () { insertBodyNo(ta, u); };
      card.appendChild(head);
      card.appendChild(ta);

      els.unitList.appendChild(card);
    });
  }

  /* ---------- спеціалісти ---------- */
  function renderSpecs() {
    els.specList.innerHTML = '';
    state.specialists.forEach(function (sp, i) {
      var card = el('div', 'spec');
      var head = el('div', 'spec-head');
      head.appendChild(el('b', null, 'Спеціаліст №' + (i + 1)));
      if (state.specialists.length > 1) {
        head.appendChild(iconBtn(ICON_TRASH, 'Видалити спеціаліста', false, async function () {
          if (!(await confirmBox('Видалити спеціаліста №' + (i + 1) + '?'))) return;
          state.specialists.splice(i, 1); saveNow(); renderSpecs();
        }));
      }
      card.appendChild(head);

      var l1 = el('label', 'field');
      l1.appendChild(el('span', null, 'ПІБ (Прізвище Ім’я По батькові)'));
      var name = document.createElement('input');
      name.type = 'text'; name.autocomplete = 'off'; name.value = sp.name;
      name.setAttribute('autocapitalize', 'words');
      name.placeholder = 'напр. Родзін Анатолій Валерійович';
      name.addEventListener('input', function () { sp.name = name.value; saveSoon(); });
      name.addEventListener('blur', function () { sp.name = D.formatSurnameFirst(name.value); name.value = sp.name; saveSoon(); });
      l1.appendChild(name);

      var l2 = el('label', 'field');
      l2.appendChild(el('span', null, 'Посада'));
      var pos = document.createElement('textarea');
      pos.rows = 2; pos.value = sp.position;
      pos.placeholder = 'напр. Спеціаліст-криміналіст ВКЗ СУ ГУНП в Одеській області, капітан поліції';
      pos.addEventListener('input', function () { sp.position = pos.value; saveSoon(); });
      l2.appendChild(pos);

      card.appendChild(l1); card.appendChild(l2);
      els.specList.appendChild(card);
    });
  }

  /* ---------- створення Word ---------- */
  async function finish() {
    if (document.activeElement && document.activeElement.blur) document.activeElement.blur();
    state.specialists.forEach(function (sp) { sp.name = D.formatSurnameFirst(sp.name); });
    if (!state.units.length) { showMessage('Помилка', 'Додайте хоча б одне фото.'); return; }
    setBusy(true, 'Створення Word-файлу…');
    try {
      await new Promise(function (r) { setTimeout(r, 60); });
      await saveNow();
      var iso = els.date.value || todayISO();
      var units = state.units.map(function (u) {
        return {
          caption: u.caption,
          layout: u.layout || 'auto',
          photos: u.photos.map(function (p) { return { bytes: new Uint8Array(p.buf), ext: 'jpeg', w: p.w, h: p.h }; })
        };
      });
      var bytes = D.buildDocx({
        bodyNo: els.bodyNo.value, date: iso, institution: els.institution.value, department: els.department.value,
        specialists: state.specialists
      }, units);
      var file = new File([bytes], D.makeFileName(els.bodyNo.value), { type: MIME });
      setBusy(false);
      showResult(file);
    } catch (e) {
      setBusy(false);
      showMessage('Помилка', 'Не вдалося створити файл: ' + (e && e.message ? e.message : e));
    }
  }

  function download(file) {
    var url = URL.createObjectURL(file);
    var a = document.createElement('a');
    a.href = url; a.download = file.name;
    document.body.appendChild(a); a.click(); a.remove();
    setTimeout(function () { URL.revokeObjectURL(url); }, 60000);
  }

  function showResult(file) {
    var buttons = [], canShare = false;
    try { canShare = !!(navigator.share && navigator.canShare && navigator.canShare({ files: [file] })); } catch (e) { canShare = false; }
    var isIOS = /iPad|iPhone|iPod/.test(navigator.userAgent) || (navigator.platform === 'MacIntel' && navigator.maxTouchPoints > 1);
    var text = file.name;
    if (canShare) {
      buttons.push({
        label: 'Поділитись', cls: 'primary', keep: true,
        onClick: function () {
          navigator.share({ files: [file], title: file.name }).catch(function (e) {
            if (e && e.name === 'AbortError') return;
            $('modalText').textContent = file.name + '\n\nНе вдалося поділитись' + (e && e.name ? ' (' + e.name + ')' : '') +
              '. Натисніть «Зберегти файл» і відкрийте його з «Файлів» або зі сповіщення про завантаження.';
          });
        }
      });
    } else {
      text += '\n\nЦей браузер не дозволяє надсилати файли .docx через «Поділитись». Натисніть «Зберегти файл», ' +
        'потім відкрийте його зі сповіщення про завантаження або в застосунку «Файли» (чи в Word) і поділіться звідти.';
    }
    buttons.push({ label: 'Зберегти файл', cls: canShare ? 'outline' : 'primary', keep: true, onClick: function () { download(file); } });
    if (isIOS && !canShare) {
      buttons.push({ label: 'Відкрити', cls: 'outline', keep: true, onClick: function () { window.open(URL.createObjectURL(file), '_blank'); } });
    }
    buttons.push({ label: 'Закрити' });
    openModal('Файл готовий', text, buttons);
  }

  /* ---------- нова таблиця ---------- */
  async function newTable() {
    if (!(await confirmBox('Почати нову таблицю? Усі фото та підписи поточної таблиці буде видалено.'))) return;
    state.units.forEach(function (u) { u.photos.forEach(function (p) { URL.revokeObjectURL(p.url); }); });
    state.units = [];
    els.bodyNo.value = '';
    els.date.value = todayISO();
    dbReq('photos', 'readwrite', function (os) { return os.clear(); }).catch(function () {});
    saveNow(); render();
    window.scrollTo(0, 0);
  }

  /* ---------- завантаження чернетки ---------- */
  async function load() {
    loadSettings();
    renderSpecs();
    idb = await openDb();
    els.date.value = todayISO();
    var meta = null;
    try { meta = await dbReq('kv', 'readonly', function (os) { return os.get('draft'); }); } catch (e) { meta = null; }
    if (meta) {
      els.bodyNo.value = meta.bodyNo || '';
      if (meta.date) els.date.value = meta.date;
      var list = meta.units || [];
      for (var i = 0; i < list.length; i++) {
        var u = { id: list[i].id, caption: list[i].caption || '', layout: list[i].layout || 'auto', photos: [] };
        var ids = list[i].photoIds || [];
        for (var k = 0; k < ids.length; k++) {
          var rec = null;
          try { rec = await dbReq('photos', 'readonly', function (os) { return os.get(ids[k]); }); } catch (e) { rec = null; }
          if (rec && rec.buf) u.photos.push({ id: ids[k], buf: rec.buf, w: rec.w, h: rec.h, url: photoUrl(rec.buf) });
        }
        if (u.photos.length) state.units.push(u);
      }
    }
    render();
  }

  /* ---------- події ---------- */
  ['bodyNo', 'date', 'institution', 'department'].forEach(function (k) {
    els[k].addEventListener('input', saveSoon);
    els[k].addEventListener('change', saveSoon);
  });
  $('btnAddSpec').addEventListener('click', function () {
    state.specialists.push({ name: '', position: '' });
    saveNow(); renderSpecs();
    var inputs = els.specList.querySelectorAll('.spec:last-child input');
    if (inputs.length) inputs[0].focus();
  });
  $('btnCamera').addEventListener('click', function () { els.cam.click(); });
  $('btnGallery').addEventListener('click', function () { els.gal.click(); });
  els.cam.addEventListener('change', function () {
    var files = Array.prototype.slice.call(els.cam.files || []);
    els.cam.value = '';
    addFiles(files, true, false);
  });
  els.gal.addEventListener('change', function () {
    var files = Array.prototype.slice.call(els.gal.files || []);
    els.gal.value = '';
    addFiles(files, files.length === 1, true);
  });
  $('btnFinish').addEventListener('click', finish);
  $('btnNew').addEventListener('click', newTable);
  document.addEventListener('visibilitychange', function () { if (document.hidden) saveNow(); });
  window.addEventListener('pagehide', saveNow);

  /* Поки відкрита клавіатура, нижню панель ховаємо, щоб вона не з’їдала екран. */
  function isTextField(e) { return e && (e.tagName === 'INPUT' || e.tagName === 'TEXTAREA') && e.type !== 'file'; }
  /* Панель ховається лише коли клавіатура справді відкрита (висота вікна помітно зменшилась). */
  function viewportH() { return window.visualViewport ? window.visualViewport.height : window.innerHeight; }
  var baseVH = viewportH();
  function updateTyping() {
    var h = viewportH();
    if (h > baseVH) baseVH = h;
    var keyboardOpen = h < baseVH - 120;
    document.body.classList.toggle('typing', keyboardOpen && isTextField(document.activeElement));
  }
  document.addEventListener('focusin', function () { setTimeout(updateTyping, 350); });
  document.addEventListener('focusout', function () { setTimeout(updateTyping, 150); });
  if (window.visualViewport) window.visualViewport.addEventListener('resize', updateTyping);
  window.addEventListener('resize', updateTyping);
  window.addEventListener('orientationchange', function () { setTimeout(function () { baseVH = viewportH(); updateTyping(); }, 500); });

  /* Торкання поза полем вводу закриває клавіатуру (і повертає нижню панель). */
  function dismissKeyboardOnTap(e) {
    var a = document.activeElement;
    if (!isTextField(a)) return;
    var t = e.target;
    if (t && t.closest && t.closest('input,textarea,select,button,label,a')) return;
    a.blur();
  }
  document.addEventListener('click', dismissKeyboardOnTap);
  document.addEventListener('touchend', dismissKeyboardOnTap, { passive: true });

  load();

  if (navigator.storage && navigator.storage.persist) { navigator.storage.persist().catch(function () {}); }
  if ('serviceWorker' in navigator) {
    window.addEventListener('load', function () { navigator.serviceWorker.register('./sw.js').catch(function () {}); });
  }
})();
