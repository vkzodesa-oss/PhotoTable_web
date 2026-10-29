/* PhotoTable (PWA): фототаблиця до протоколу огляду останків -> Word.
   Усе працює на пристрої. */
(function () {
  'use strict';

  var D = window.PhotoTableDocx;
  var MIME = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';
  var MAX_SIDE = 1800;
  var QUALITY = 0.85;

  var ICON_UP = '<svg viewBox="0 0 24 24" width="22" height="22"><path fill="currentColor" d="M4 12l1.4 1.4L11 7.8V20h2V7.8l5.6 5.6L20 12l-8-8-8 8z"/></svg>';
  var ICON_DOWN = '<svg viewBox="0 0 24 24" width="22" height="22"><path fill="currentColor" d="M20 12l-1.4-1.4L13 16.2V4h-2v12.2l-5.6-5.6L4 12l8 8 8-8z"/></svg>';
  var ICON_CLOSE = '<svg viewBox="0 0 24 24" width="24" height="24"><path fill="currentColor" d="M19 6.4 17.6 5 12 10.6 6.4 5 5 6.4 10.6 12 5 17.6 6.4 19 12 13.4 17.6 19 19 17.6 13.4 12z"/></svg>';
  var ICON_TRASH = '<svg viewBox="0 0 24 24" width="22" height="22"><path fill="currentColor" d="M6 19a2 2 0 0 0 2 2h8a2 2 0 0 0 2-2V7H6v12zM19 4h-3.5l-1-1h-5l-1 1H5v2h14V4z"/></svg>';

  function $(id) { return document.getElementById(id); }
  var els = {
    place: $('place'), bodyNo: $('bodyNo'), date: $('date'),
    spCount: $('spCount'), specList: $('specList'), btnAddSpec: $('btnAddSpec'),
    count: $('count'), list: $('photoList'), cam: $('camInput'), gal: $('galInput')
  };
  var state = { photos: [], specialists: [{ name: '', position: '' }, { name: '', position: '' }] };

  /* ---------- сховище (IndexedDB) ---------- */
  var idb = null;
  function openDb() {
    return new Promise(function (resolve) {
      if (!window.indexedDB) { resolve(null); return; }
      try {
        var req = indexedDB.open('phototable', 1);
        req.onupgradeneeded = function () {
          req.result.createObjectStore('kv');
          req.result.createObjectStore('photos');
        };
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

  function pad(n) { return n < 10 ? '0' + n : '' + n; }
  function todayISO() {
    var d = new Date();
    return d.getFullYear() + '-' + pad(d.getMonth() + 1) + '-' + pad(d.getDate());
  }

  var saveTimer = null;
  function saveSoon() { clearTimeout(saveTimer); saveTimer = setTimeout(saveNow, 300); }
  function saveNow() {
    clearTimeout(saveTimer);
    var meta = {
      place: els.place.value, bodyNo: els.bodyNo.value, date: els.date.value,
      specialists: state.specialists.map(function (s) { return { name: s.name, position: s.position }; }),
      order: state.photos.map(function (p) { return { id: p.id, caption: p.caption, group: !!p.group }; })
    };
    return dbReq('kv', 'readwrite', function (os) { return os.put(meta, 'draft'); }).catch(function () { /* ignore */ });
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
      btn.onclick = function () {
        if (!b.keep) closeModal();
        if (b.onClick) b.onClick();
      };
      box.appendChild(btn);
    });
    $('modal').hidden = false;
  }
  function showMessage(title, text) {
    openModal(title, text, [{ label: 'OK', cls: 'primary' }]);
  }
  function confirmBox(text) {
    return new Promise(function (resolve) {
      openModal('', text, [
        { label: 'Так', cls: 'primary', onClick: function () { resolve(true); } },
        { label: 'Ні', onClick: function () { resolve(false); } }
      ]);
    });
  }

  function iconBtn(svg, label, disabled, onClick) {
    var b = document.createElement('button');
    b.type = 'button';
    b.className = 'icon-btn';
    b.innerHTML = svg;
    b.setAttribute('aria-label', label);
    b.title = label;
    b.disabled = !!disabled;
    b.onclick = onClick;
    return b;
  }

  /* ---------- спеціалісти ---------- */
  function renderSpecs() {
    els.spCount.textContent = state.specialists.length;
    els.specList.innerHTML = '';
    state.specialists.forEach(function (sp, i) {
      var card = document.createElement('div');
      card.className = 'spec';

      var head = document.createElement('div');
      head.className = 'spec-head';
      var title = document.createElement('div');
      title.className = 'ptitle';
      title.textContent = 'Спеціаліст №' + (i + 1);
      head.appendChild(title);
      if (state.specialists.length > 1) {
        head.appendChild(iconBtn(ICON_TRASH, 'Видалити спеціаліста', false, function () { removeSpec(i); }));
      }
      card.appendChild(head);

      var fName = document.createElement('label');
      fName.className = 'field';
      var sName = document.createElement('span');
      sName.textContent = 'ПІБ (прізвище — великими літерами)';
      var iName = document.createElement('input');
      iName.type = 'text';
      iName.autocomplete = 'off';
      iName.autocapitalize = 'words';
      iName.enterkeyhint = 'next';
      iName.value = sp.name;
      iName.placeholder = 'напр. РОДЗІН Анатолій Валерійович';
      iName.addEventListener('input', function () { sp.name = iName.value; saveSoon(); });
      fName.appendChild(sName); fName.appendChild(iName);
      card.appendChild(fName);

      var fPos = document.createElement('label');
      fPos.className = 'field';
      var sPos = document.createElement('span');
      sPos.textContent = 'Посада / звання (другий рядок у графі «ПІБ»)';
      var iPos = document.createElement('input');
      iPos.type = 'text';
      iPos.autocomplete = 'off';
      iPos.enterkeyhint = 'next';
      iPos.value = sp.position;
      iPos.placeholder = 'напр. спеціаліст-криміналіст ВКЗ СУ ГУНП в Одеській області, капітан поліції';
      iPos.addEventListener('input', function () { sp.position = iPos.value; saveSoon(); });
      fPos.appendChild(sPos); fPos.appendChild(iPos);
      card.appendChild(fPos);

      els.specList.appendChild(card);
    });
  }
  async function removeSpec(i) {
    if (!(await confirmBox('Видалити спеціаліста №' + (i + 1) + '?'))) return;
    state.specialists.splice(i, 1);
    saveNow();
    renderSpecs();
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
    } finally {
      URL.revokeObjectURL(url);
    }
  }
  function photoUrl(buf) { return URL.createObjectURL(new Blob([buf], { type: 'image/jpeg' })); }

  /* Номери фото з урахуванням груп: повертає [{num, groupStart, groupEnd, inGroup}] */
  function photoNumbers() {
    var out = [];
    var i = 0, n = state.photos.length;
    while (i < n) {
      var j = i + 1;
      while (j < n && state.photos[j].group) j++;
      for (var k = i; k < j; k++) {
        out.push({ num: k + 1, start: i + 1, end: j, inGroup: j - i > 1 });
      }
      i = j;
    }
    return out;
  }

  /* Екран «Порядок фото»: торкайтесь фото в потрібному порядку. */
  function chooseOrder(pending) {
    return new Promise(function (resolve) {
      var order = [];
      var page = document.createElement('div');
      page.className = 'order-page';

      var head = document.createElement('div');
      head.className = 'order-head';
      var title = document.createElement('h3');
      title.textContent = 'Порядок фото';
      var reset = document.createElement('button');
      reset.type = 'button';
      reset.className = 'link-btn';
      reset.textContent = 'Скинути';

      var hint = document.createElement('p');
      hint.className = 'order-hint';
      hint.textContent = 'Торкайтесь фото в тому порядку, у якому вони мають бути у фототаблиці. Повторне торкання скасовує вибір.';

      var grid = document.createElement('div');
      grid.className = 'order-grid';
      var cells = [];
      pending.forEach(function (p, i) {
        var cell = document.createElement('button');
        cell.type = 'button';
        cell.className = 'order-cell';
        var im = document.createElement('img');
        im.src = p.url;
        im.alt = '';
        var badge = document.createElement('span');
        badge.className = 'order-badge';
        cell.appendChild(im);
        cell.appendChild(badge);
        cell.addEventListener('click', function () {
          var k = order.indexOf(i);
          if (k >= 0) order.splice(k, 1); else order.push(i);
          update();
        });
        grid.appendChild(cell);
        cells.push({ cell: cell, badge: badge });
      });

      var foot = document.createElement('div');
      foot.className = 'order-foot';
      var ok = document.createElement('button');
      ok.type = 'button';
      ok.className = 'btn primary';
      foot.appendChild(ok);

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
      head.appendChild(reset);
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
        var id = Date.now().toString(36) + Math.random().toString(36).slice(2, 7);
        pending.push({ id: id, buf: r.buf, w: r.w, h: r.h, caption: '', group: false, url: photoUrl(r.buf) });
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
      var ph = ordered[j];
      state.photos.push(ph);
      await dbReq('photos', 'readwrite', function (os) { return os.put({ buf: ph.buf, w: ph.w, h: ph.h }, ph.id); })
        .catch(function () { /* без збереження чернетки працюємо далі */ });
    }
    saveNow();
    render();
    if (errors.length) showMessage('Помилка', 'Не вдалося обробити: ' + errors.join('; '));
    if (focusLast && ordered.length) {
      var last = els.list.querySelector('.photo:last-child textarea');
      if (last) { last.focus(); last.scrollIntoView({ block: 'center', behavior: 'smooth' }); }
    }
  }

  function movePhoto(i, d) {
    var t = i + d;
    if (t < 0 || t >= state.photos.length) return;
    var p = state.photos.splice(i, 1)[0];
    state.photos.splice(t, 0, p);
    saveNow();
    render();
  }
  async function removePhoto(i) {
    if (!(await confirmBox('Видалити фото №' + (i + 1) + '?'))) return;
    var p = state.photos.splice(i, 1)[0];
    if (p) {
      URL.revokeObjectURL(p.url);
      dbReq('photos', 'readwrite', function (os) { return os.delete(p.id); }).catch(function () { /* ignore */ });
    }
    saveNow();
    render();
  }

  function render() {
    els.count.textContent = state.photos.length;
    els.list.innerHTML = '';
    var nums = photoNumbers();
    state.photos.forEach(function (p, i) {
      var nm = nums[i];
      var card = document.createElement('div');
      card.className = 'photo' + (p.group ? ' grouped' : '');

      var top = document.createElement('div');
      top.className = 'photo-top';
      var img = document.createElement('img');
      img.className = 'thumb';
      img.alt = '';
      img.src = p.url;
      var meta = document.createElement('div');
      meta.className = 'pmeta';
      var title = document.createElement('div');
      title.className = 'ptitle';
      title.textContent = nm.inGroup
        ? 'Фото №' + nm.num + ' (група №' + nm.start + '–' + nm.end + ')'
        : 'Фото №' + nm.num;
      var ctrls = document.createElement('div');
      ctrls.className = 'ctrls';
      ctrls.appendChild(iconBtn(ICON_UP, 'Вище', i === 0, function () { movePhoto(i, -1); }));
      ctrls.appendChild(iconBtn(ICON_DOWN, 'Нижче', i === state.photos.length - 1, function () { movePhoto(i, 1); }));
      ctrls.appendChild(iconBtn(ICON_TRASH, 'Видалити', false, function () { removePhoto(i); }));
      meta.appendChild(title);
      meta.appendChild(ctrls);
      top.appendChild(img);
      top.appendChild(meta);
      card.appendChild(top);

      /* перемикач групування */
      var swRow = document.createElement('label');
      swRow.className = 'switch-row';
      var cb = document.createElement('input');
      cb.type = 'checkbox';
      cb.checked = !!p.group;
      cb.disabled = i === 0;
      var track = document.createElement('span');
      track.className = 'track';
      var swLabel = document.createElement('span');
      swLabel.className = 'sw-label';
      swLabel.textContent = 'Об’єднати з попереднім фото (спільний підпис)';
      cb.addEventListener('change', function () { p.group = cb.checked; saveSoon(); render(); });
      swRow.appendChild(cb);
      swRow.appendChild(track);
      swRow.appendChild(swLabel);
      card.appendChild(swRow);

      /* підпис тільки для першого фото групи */
      if (!p.group) {
        var label = document.createElement('label');
        label.className = 'field';
        var span = document.createElement('span');
        span.textContent = 'Підпис (текст після «Фото №' + nm.num + (nm.end > nm.num ? '–' + nm.end : '') + '::»)';
        var ta = document.createElement('textarea');
        ta.rows = 2;
        ta.value = p.caption;
        ta.placeholder = 'напр. загальний вигляд мішка для транспортування тіла №20260916010004067';
        ta.addEventListener('input', function () { p.caption = ta.value; saveSoon(); });
        label.appendChild(span);
        label.appendChild(ta);
        card.appendChild(label);
      }

      els.list.appendChild(card);
    });
  }

  /* ---------- створення Word ---------- */
  function fileName() {
    var no = D.clean(els.bodyNo.value).replace(/[\\/:*?"<>|]/g, '');
    return 'Фототаблиця ' + (no || D.formatDate(els.date.value || todayISO())) + '.docx';
  }

  async function finish() {
    if (document.activeElement && document.activeElement.blur) document.activeElement.blur();
    if (!state.photos.length) { showMessage('Помилка', 'Додайте хоча б одну фотографію.'); return; }
    var unnamed = state.specialists.filter(function (s) { return !D.clean(s.name); }).length;
    if (unnamed) { showMessage('Помилка', 'Заповніть ПІБ усіх спеціалістів (або видаліть зайвих).'); return; }
    setBusy(true, 'Створення Word-файлу…');
    try {
      await new Promise(function (r) { setTimeout(r, 60); });
      await saveNow();
      var images = state.photos.map(function (p) {
        return { bytes: new Uint8Array(p.buf), ext: 'jpeg', w: p.w, h: p.h, caption: p.caption, group: !!p.group };
      });
      var bytes = D.buildDocx(
        { place: els.place.value, bodyNo: els.bodyNo.value, date: els.date.value || todayISO() },
        images,
        state.specialists
      );
      var file = new File([bytes], fileName(), { type: MIME });
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
    a.href = url;
    a.download = file.name;
    document.body.appendChild(a);
    a.click();
    a.remove();
    setTimeout(function () { URL.revokeObjectURL(url); }, 60000);
  }

  function showResult(file) {
    var buttons = [];
    var canShare = false;
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
    if (!(await confirmBox('Почати нову фототаблицю? Усі фото, підписи та дані поточної таблиці буде видалено.'))) return;
    state.photos.forEach(function (p) { URL.revokeObjectURL(p.url); });
    state.photos = [];
    state.specialists = [{ name: '', position: '' }, { name: '', position: '' }];
    els.place.value = '';
    els.bodyNo.value = '';
    els.date.value = todayISO();
    dbReq('photos', 'readwrite', function (os) { return os.clear(); }).catch(function () { /* ignore */ });
    saveNow();
    renderSpecs();
    render();
    window.scrollTo(0, 0);
  }

  /* ---------- завантаження чернетки ---------- */
  async function load() {
    idb = await openDb();
    els.date.value = todayISO();
    var meta = null;
    try { meta = await dbReq('kv', 'readonly', function (os) { return os.get('draft'); }); } catch (e) { meta = null; }
    if (meta) {
      els.place.value = meta.place || '';
      els.bodyNo.value = meta.bodyNo || '';
      if (meta.date) els.date.value = meta.date;
      if (Array.isArray(meta.specialists) && meta.specialists.length) {
        state.specialists = meta.specialists.map(function (s) { return { name: s.name || '', position: s.position || '' }; });
      }
      var order = meta.order || [];
      for (var i = 0; i < order.length; i++) {
        var rec = null;
        try { rec = await dbReq('photos', 'readonly', function (os) { return os.get(order[i].id); }); } catch (e) { rec = null; }
        if (rec && rec.buf) {
          state.photos.push({
            id: order[i].id, buf: rec.buf, w: rec.w, h: rec.h,
            caption: order[i].caption || '', group: !!order[i].group, url: photoUrl(rec.buf)
          });
        }
      }
    }
    renderSpecs();
    render();
  }

  /* ---------- події ---------- */
  ['place', 'bodyNo', 'date'].forEach(function (k) {
    els[k].addEventListener('input', saveSoon);
    els[k].addEventListener('change', saveSoon);
  });

  /* Поки відкрита клавіатура, нижню панель ховаємо, щоб вона не з’їдала екран. */
  function isTextField(el) { return el && (el.tagName === 'INPUT' || el.tagName === 'TEXTAREA') && el.type !== 'file'; }
  document.addEventListener('focusin', function (e) { if (isTextField(e.target)) document.body.classList.add('typing'); });
  document.addEventListener('focusout', function () {
    setTimeout(function () { if (!isTextField(document.activeElement)) document.body.classList.remove('typing'); }, 120);
  });

  els.btnAddSpec.addEventListener('click', function () {
    state.specialists.push({ name: '', position: '' });
    saveSoon();
    renderSpecs();
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

  load();

  if (navigator.storage && navigator.storage.persist) { navigator.storage.persist().catch(function () { /* ignore */ }); }
  if ('serviceWorker' in navigator) {
    window.addEventListener('load', function () {
      navigator.serviceWorker.register('./sw.js').catch(function () { /* ignore */ });
    });
  }
})();
