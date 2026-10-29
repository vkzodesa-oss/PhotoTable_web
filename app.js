(function () {
'use strict';
var D = window.CrimDocx;
var MIME = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';
var MAX_SIDE = 1800;
var QUALITY = 0.85;

function $(id) { return document.getElementById(id); }
var els = { bodyNumber: $('bodyNumber'), place: $('place'), date: $('date'), spec1: $('spec1'), spec2: $('spec2'), count: $('count'), list: $('photoList'), cam: $('camInput'), gal: $('galInput') };
var state = { photos: [] };

/* ---------- Допоміжні UI ---------- */
function setBusy(on, text) { $('overlay').hidden = !on; $('btnFinish').disabled = on; }
function closeModal() { $('modal').hidden = true; }
function openModal(title, text, buttons) {
  $('modalTitle').textContent = title; $('modalText').textContent = text;
  var box = $('modalBtns'); box.innerHTML = '';
  buttons.forEach(function (b) {
    var btn = document.createElement('button'); btn.type = 'button'; btn.className = 'btn ' + (b.cls || 'outline'); btn.textContent = b.label;
    btn.onclick = function () { if (!b.keep) closeModal(); if (b.onClick) b.onClick(); };
    box.appendChild(btn);
  });
  $('modal').hidden = false;
}
function showMessage(t, m) { openModal(t, m, [{ label: 'OK', cls: 'primary' }]); }
function confirmBox(text) {
  return new Promise(function (resolve) {
    openModal('', text, [{ label: 'Так', cls: 'primary', onClick: function () { resolve(true); } }, { label: 'Ні', onClick: function () { resolve(false); } }]);
  });
}

/* ---------- Фото та Групи ---------- */
async function addFiles(files) {
  if (!files.length) return;
  for (var i = 0; i < files.length; i++) {
    try {
      var url = URL.createObjectURL(files[i]);
      var img = await loadImage(url);
      var w = Math.min(img.naturalWidth, MAX_SIDE), h = Math.round(w * img.naturalHeight / img.naturalWidth);
      var c = document.createElement('canvas'); c.width = w; c.height = h;
      c.getContext('2d').drawImage(img, 0, 0, w, h);
      var blob = await new Promise(r => c.toBlob(r, 'image/jpeg', QUALITY));
      var buf = await blobToBuf(blob);
      state.photos.push({ id: Date.now().toString(36)+Math.random().toString(36).slice(2,7), buf: buf, w: w, h: h, caption: '', group: null, url: photoUrl(buf) });
    } catch(e) {}
  }
  render(); saveNow();
}
function render() {
  els.count.textContent = state.photos.length; els.list.innerHTML = '';
  state.photos.forEach(function (p, i) {
    var card = document.createElement('div'); card.className = 'photo';
    card.innerHTML = `<div class="photo-top"><img class="thumb" src="${p.url}"><div class="pmeta"><div class="ptitle">Фото ${i+1}</div></div></div>`;
    
    // Кнопки керування та Групою
    var ctrls = document.createElement('div'); ctrls.className = 'ctrls';
    ctrls.innerHTML = `
      <button class="icon-btn" onclick="window.appMove(${i}, -1)">⬆️</button>
      <button class="icon-btn" onclick="window.appMove(${i}, 1)">⬇️</button>
      <button class="icon-btn" onclick="window.appRemove(${i})">🗑️</button>
      <button class="btn-group ${p.group?'active':''}" onclick="window.appToggleGroup(${i})">${p.group ? '🔗 У групі' : '➕ До групи'}</button>
    `;
    card.querySelector('.pmeta').appendChild(ctrls);

    var ta = document.createElement('textarea'); ta.rows = 2; ta.value = p.caption; ta.placeholder = "Підпис до фото";
    ta.oninput = function() { p.caption = ta.value; saveSoon(); };
    card.appendChild(ta); els.list.appendChild(card);
  });
  $('btnFinish').disabled = !state.photos.length || !els.bodyNumber.value.trim();
}
window.appMove = function(i, d) { var t=i+d; if(t<0||t>=state.photos.length)return; var p=state.photos.splice(i,1)[0]; state.photos.splice(t,0,p); render(); saveNow(); };
window.appRemove = async function(i) { if(await confirmBox('Видалити?')) { state.photos.splice(i,1); render(); saveNow(); } };
window.appToggleGroup = function(i) { 
  var g = state.photos[i].group;
  if (g === null) { // Створити нову групу або приєднатись
    var lastGroup = state.photos.filter(p=>p.group!==null).sort((a,b)=>b.group-a.group)[0];
    state.photos[i].group = lastGroup ? lastGroup.group + 1 : 1;
  } else { // Розгрупувати
    state.photos[i].group = null;
  }
  render(); saveNow();
};

/* ---------- Завантаження та Word ---------- */
async function finish() {
  normalizeNames();
  var bn = els.bodyNumber.value.trim();
  if(!bn) { showMessage("Помилка","Введіть номер тіла!"); return; }
  if(!state.photos.length) { showMessage("Помилка","Додайте фото."); return; }
  
  setBusy(true, "Збірка...");
  try {
    var images = state.photos.map(function(p){ return { bytes:new Uint8Array(p.buf), ext:'jpeg', w:p.w, h:p.h, caption:p.caption, group:p.group }; });
    var info = { bodyNumber: bn, place: els.place.value, date: els.date.value, spec1: els.spec1.value, spec2: els.spec2.value };
    var bytes = D.buildRemainsDocx(info, images);
    var file = new File([bytes], `Фототаблиця_${bn}.docx`, { type: MIME });
    showResult(file);
  } catch(e) { showMessage("Помилка", e.message); }
  finally { setBusy(false); }
}
function showResult(file) {
  openModal('Готово', `Файл: ${file.name}`, [
    { label: 'Поділитись', cls: 'primary', keep:true, onClick: function(){ navigator.share({files:[file]}).catch(()=>{}); } },
    { label: 'Зберегти файл', cls: 'outline', keep:true, onClick: function(){ download(file); } },
    { label: 'Закрити' }
  ]);
}
function download(file) { var a=document.createElement('a'); a.href=URL.createObjectURL(file); a.download=file.name; a.click(); }
function normalizeNames() { /* Проста валідація */ }

/* ---------- Події ---------- */
$('btnCamera').onclick = () => els.cam.click();
$('btnGallery').onclick = () => els.gal.click();
els.cam.onchange = e => addFiles(Array.from(e.target.files));
els.gal.onchange = e => addFiles(Array.from(e.target.files));
$('btnFinish').onclick = finish;
$('btnNew').onclick = async () => { if(await confirmBox('Очистити все?')) { state.photos=[]; render(); localStorage.clear(); } };
['bodyNumber','place','date','spec1','spec2'].forEach(k => els[k].addEventListener('input', saveSoon));

/* IndexedDB та Service Worker (спрощено) */
function saveSoon() { setTimeout(saveNow, 500); }
function saveNow() { localStorage.setItem('remains_draft', JSON.stringify(state)); }
function load() { var d=localStorage.getItem('remains_draft'); if(d) { try{ state=JSON.parse(d); render(); }catch(e){} } }
load();
})();