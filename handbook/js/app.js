"use strict";

const state = {
  view: "catalog",
  id: null,
  q: "",
  category: "all",
  status: "all",
  compare: [],
};

const app = document.getElementById("app");

function esc(s) {
  return String(s ?? "").replace(/[&<>"']/g, (ch) => ({
    "&": "&amp;",
    "<": "&lt;",
    ">": "&gt;",
    '"': "&quot;",
    "'": "&#39;",
  }[ch]));
}

function byId(id) {
  return AIRCRAFT.find((a) => a.id === id);
}

function parseHash() {
  const raw = location.hash.replace(/^#\/?/, "");
  const [view, id] = raw.split("/");
  if (!view) return { view: "catalog", id: null };
  if (view === "type" && byId(id)) return { view: "type", id };
  if (view === "sections" || view === "about" || view === "compare") return { view, id: null };
  return { view: "catalog", id: null };
}

function go(view, id) {
  if (view === "type") location.hash = `#/type/${id}`;
  else if (view === "catalog") location.hash = "#/";
  else location.hash = `#/${view}`;
}

function syncNav() {
  document.querySelectorAll(".nav button").forEach((btn) => {
    const dest = btn.dataset.go;
    const on = (dest === "catalog" && state.view === "catalog") || dest === state.view;
    btn.setAttribute("aria-current", on ? "page" : "false");
  });
}

function matches(ac) {
  if (state.category !== "all" && ac.category !== state.category) return false;
  if (state.status !== "all" && ac.status !== state.status) return false;
  if (!state.q.trim()) return true;
  const blob = [
    ac.name, ac.family, ac.manufacturer, ac.icao, ac.search, ac.role,
    ...ac.engines.map((e) => e.name),
    ...ac.apus.map((e) => e.name),
  ].join(" ").toLowerCase();
  return state.q.trim().toLowerCase().split(/\s+/).every((part) => blob.includes(part));
}

function render() {
  const parsed = parseHash();
  state.view = parsed.view;
  state.id = parsed.id;
  syncNav();
  if (state.view === "type") renderType(byId(state.id));
  else if (state.view === "sections") renderSections();
  else if (state.view === "about") renderAbout();
  else if (state.view === "compare") renderCompare();
  else renderCatalog();
  const bar = document.getElementById("compare-bar");
  if (bar) bar.onclick = onCompareClick;
}

function renderCatalog() {
  const groups = CATEGORIES.map((cat) => ({
    cat,
    items: AIRCRAFT.filter((a) => a.category === cat.id && matches(a)),
  })).filter((g) => g.items.length);
  app.innerHTML = `
    <p class="banner">${esc(HANDBOOK_NOTE)}</p>
    <div class="filters">
      <input id="q" type="search" placeholder="Тип, ICAO, двигатель, ВСУ" value="${esc(state.q)}" aria-label="Поиск">
      <select id="category" aria-label="Класс">
        <option value="all">Все классы</option>
        ${CATEGORIES.map((c) => `<option value="${c.id}" ${state.category === c.id ? "selected" : ""}>${esc(c.title)}</option>`).join("")}
      </select>
      <select id="status" aria-label="Статус выпуска">
        <option value="all">Любой статус</option>
        ${Object.entries(STATUSES).map(([id, label]) => `<option value="${id}" ${state.status === id ? "selected" : ""}>${esc(label)}</option>`).join("")}
      </select>
    </div>
    ${groups.length ? groups.map(groupHtml).join("") : "<p>Ничего не найдено. Сбросьте фильтр или поищите по ICAO, например A320 или ПС-90.</p>"}
    ${compareBar()}
  `;
  document.getElementById("q").addEventListener("input", (e) => {
    state.q = e.target.value;
    renderCatalog();
    const input = document.getElementById("q");
    input.focus();
    input.setSelectionRange(input.value.length, input.value.length);
  });
  document.getElementById("category").onchange = (e) => { state.category = e.target.value; render(); };
  document.getElementById("status").onchange = (e) => { state.status = e.target.value; render(); };
  app.querySelectorAll("[data-open]").forEach((btn) => {
    btn.onclick = () => go("type", btn.dataset.open);
  });
  const bar = document.getElementById("compare-bar");
  if (bar) bar.onclick = onCompareClick;
}

function groupHtml(group) {
  return `
    <h2 class="group-title">${esc(group.cat.title)}</h2>
    <div class="cards">
      ${group.items.map(cardHtml).join("")}
    </div>
  `;
}

function cardHtml(ac) {
  return `
    <button class="card" type="button" data-open="${ac.id}">
      <div class="thumb">${crossSectionSVG(ac, { mode: "mini", ppm: 26 })}</div>
      <h3>${esc(ac.name)}</h3>
      <div class="meta">${esc(ac.icao)} · ${esc(ac.years)}</div>
      <div><span class="pill ${ac.status}">${esc(STATUSES[ac.status])}</span></div>
      <div class="meta">${esc(abreastLabel(ac))}</div>
    </button>
  `;
}

function renderType(ac) {
  if (!ac) { go("catalog"); return; }
  document.title = `${ac.name} — справочник типов ВС`;
  app.innerHTML = `
    <p class="banner">${esc(HANDBOOK_NOTE)}</p>
    <div class="back"><button class="text-btn" type="button" id="back">К каталогу</button>
      <button class="text-btn" type="button" id="add-compare">${state.compare.includes(ac.id) ? "Убрать из сравнения" : "К сравнению сечений"}</button>
    </div>
    <div class="layout">
      <article class="panel">
        <p class="meta">${esc(ac.family)} · ${esc(ac.manufacturer)}</p>
        <h2>${esc(ac.name)}</h2>
        <p><span class="pill ${ac.status}">${esc(STATUSES[ac.status])}</span> <span class="meta">${esc(ac.icao)} · ${esc(ac.years)}</span></p>
        <p>${esc(ac.role)}</p>
        ${ac.insight.map((p) => `<p>${esc(p)}</p>`).join("")}
        <h3 class="subhead">Характеристики</h3>
        <table>${ac.specs.map((row) => `<tr><th>${esc(row[0])}</th><td>${esc(row[1])}</td></tr>`).join("")}</table>
        <p class="caption">${esc(ac.yearsNote)}</p>
        <h3 class="subhead">Модификации</h3>
        <ul class="list">${ac.variants.map((v) => `<li><strong>${esc(v.t)}.</strong> ${esc(v.d)}</li>`).join("")}</ul>
        <h3 class="subhead">Двигатели</h3>
        <ul class="list">${ac.engines.map((v) => `<li><strong>${esc(v.name)}.</strong> ${esc(v.detail)}</li>`).join("")}</ul>
        <h3 class="subhead">ВСУ</h3>
        <ul class="list">${ac.apus.map((v) => `<li><strong>${esc(v.name)}.</strong> ${esc(v.detail)}</li>`).join("")}</ul>
        <h3 class="subhead">Шасси</h3>
        <ul class="list">${ac.gear.points.map((p) => `<li>${esc(p)}</li>`).join("")}</ul>
      </article>
      <aside>
        <div class="panel drawing">
          <h3 class="subhead" style="margin-top:0">Поперечное сечение</h3>
          ${crossSectionSVG(ac, { mode: "detail" })}
          <p class="caption">Фигура стоит в проходе и везде имеет рост 175 см. Сплошной красный размер без пометы — опубликованная величина. Помета «схема» — высота прохода, принятая для чертежа, если в открытых источниках нет одной официальной цифры. «Допущение» — величина, которой нет у производителя и которая выбрана, чтобы сечение было физически согласовано.</p>
        </div>
        <div class="panel drawing" style="margin-top:14px">
          <h3 class="subhead" style="margin-top:0">След шасси</h3>
          ${gearSVG(ac)}
        </div>
      </aside>
    </div>
    ${compareBar()}
  `;
  document.getElementById("back").onclick = () => go("catalog");
  document.getElementById("add-compare").onclick = () => toggleCompare(ac.id);
}

function renderSections() {
  document.title = "Сечения в одном масштабе — справочник типов ВС";
  app.innerHTML = `
    <p class="banner">Все контуры ниже нарисованы с одним масштабом: 48 пикселей на метр. Рост фигуры везде 175 см. Широкий фюзеляж больше на экране, а не «растянут под карточку».</p>
    <div class="gallery">
      ${AIRCRAFT.map((ac) => `
        <figure>
          <button class="text-btn" type="button" data-open="${ac.id}" style="padding:0;border:0;background:transparent">
            ${crossSectionSVG(ac, { mode: "gallery", ppm: 48 })}
          </button>
          <figcaption><strong>${esc(ac.name)}</strong>${esc(abreastLabel(ac))}</figcaption>
        </figure>
      `).join("")}
    </div>
    <h2 class="group-title">Ширина салона</h2>
    <div class="panel">
      <table>
        <tr><th>Тип</th><th>Фюзеляж</th><th>Салон</th><th>Ряд</th></tr>
        ${AIRCRAFT.map((ac) => {
          const d = ac.section.decks[0];
          return `<tr><td>${esc(ac.name)}</td><td>${esc(fmtM(ac.section.dimW))}${ac.section.dimWKind === "published" ? "" : " · схема"}</td><td>${esc(fmtM(d.cabinW || d.span))}</td><td>${esc(d.groups.join("+"))}</td></tr>`;
        }).join("")}
      </table>
    </div>
    ${compareBar()}
  `;
  app.querySelectorAll("[data-open]").forEach((btn) => {
    btn.onclick = () => go("type", btn.dataset.open);
  });
}

function renderCompare() {
  const items = state.compare.map(byId).filter(Boolean);
  document.title = "Сравнение сечений — справочник типов ВС";
  app.innerHTML = `
    <div class="back"><button class="text-btn" type="button" id="back">К каталогу</button></div>
    <h2 class="group-title">Сравнение</h2>
    <p>До трёх типов, один масштаб. Добавляйте карточки кнопкой «К сравнению сечений».</p>
    ${items.length ? `<div class="gallery">${items.map((ac) => `
      <figure>
        ${crossSectionSVG(ac, { mode: "compare", ppm: 58 })}
        <figcaption><strong>${esc(ac.name)}</strong>${esc(abreastLabel(ac))}</figcaption>
      </figure>`).join("")}</div>` : "<p>Пока пусто.</p>"}
    ${compareBar()}
  `;
  document.getElementById("back").onclick = () => go("catalog");
}

function renderAbout() {
  document.title = "Как читать справочник — типы ВС";
  app.innerHTML = `
    <article class="prose panel">
      <h2>Как читать</h2>
      <p>Справочник собран для занятий и для полки инженерно-авиационной службы, не для расчёта рейса. В каждой карточке — одна представительная модификация и семейство вокруг неё. Тяга двигателя зависит от суффикса на табличке; здесь указан порядок величины.</p>
      <h3 class="subhead">Сечение</h3>
      <p>Рисунок — вид в хвост, на гермофюзеляж в миделе. Кресла показаны спереди, фигура стоит в проходе. Её рост 175 см — ориентир среднего взрослого, не сертификационный манекен и не 95-й перцентиль.</p>
      <p>Опубликованные размеры (диаметр по листу типа, ширина салона из буклета производителя) подписаны без оговорок. Если открытые источники не дают одной высоты прохода, размер помечен «схема». У C919 рекламные 3,9 м салона не сходятся с внешней шириной 3,96 м, поэтому на чертеже стоит помеченное допущение.</p>
      <p>Кессон крыла низкоплана проходит через нижнюю часть фюзеляжа и делит багажники на передний и задний. На схемах он не выделен. У высокопланов (ATR, Q400, Ан-148) крыло лежит над потолком салона.</p>
      <h3 class="subhead">ВСУ</h3>
      <p>Вспомогательная силовая установка — небольшой газотурбинный двигатель, почти всегда в хвостовом конусе. На земле она даёт электричество и, на большинстве типов, сжатый воздух для кондиционирования и запуска маршевых двигателей. В полёте это резервный источник.</p>
      <p>Исключения в этом справочнике стоит запомнить отдельно. У ATR штатной ВСУ нет: работает режим Hotel правого двигателя при заторможенном винте. У Boeing 787 ВСУ APS5000 электрическая, без привычного отбора воздуха; компрессоры салона тоже электрические. Остальные типы в подборке — с воздушной ВСУ.</p>
      <h3 class="subhead">Шасси</h3>
      <p>Почти все машины в справочнике — убирающиеся трёхопорные. Число колёс определяет масса, а не число кресел в ряду. A321 при массе около 97 т остаётся на шести колёсах. Ту-214 и Boeing 757 при массе за 100 т переходят на четырёхколёсные тележки, хотя салон у них узкий. A350-1000 меняет тележку на шестиколёсную при том же сечении, что у A350-900.</p>
      <p>Пневматики наполняют азотом. В колёсах стоят легкоплавкие пробки: перегретый тормоз стравливает давление, а не рвёт шину. Носовая стойка управляемая. Тормоза — на основных колёсах.</p>
      ${primerGearSVG()}
      <h3 class="subhead">Откуда цифры</h3>
      <p>Листы данных типа EASA и FAA, документы airport planning Airbus, Boeing и COMAC, буклеты ОАК, ATR, Embraer и De Havilland, обзоры парка ВСУ. Годы выпуска и статус программ сверены с открытыми сообщениями на конец 2025 — ноябрь 2026. SJ-100 и МС-21 описаны как программы испытаний и постановки в серию, а не как уже массово переданные в авиакомпании борта.</p>
      <p>В подборку не вошли деловые самолёты, вертолёты, Ан-24, Ил-114, грузовые рамповые типы и военные машины. Их сечения — отдельная тетрадь.</p>
    </article>
  `;
}

function compareBar() {
  if (!state.compare.length) return "";
  const names = state.compare.map((id) => byId(id)).filter(Boolean);
  return `
    <div class="compare-bar no-print" id="compare-bar">
      <span>В сравнении: ${names.map((a) => esc(a.name)).join(", ")}</span>
      <button type="button" data-act="open">Открыть</button>
      <button type="button" data-act="clear">Очистить</button>
    </div>
  `;
}

function onCompareClick(e) {
  const btn = e.target.closest("button");
  if (!btn) return;
  if (btn.dataset.act === "open") go("compare");
  if (btn.dataset.act === "clear") {
    state.compare = [];
    render();
  }
}

function toggleCompare(id) {
  const i = state.compare.indexOf(id);
  if (i >= 0) state.compare.splice(i, 1);
  else if (state.compare.length < 3) state.compare.push(id);
  else state.compare.splice(0, 1, id);
  render();
}

document.querySelectorAll(".nav button").forEach((btn) => {
  btn.onclick = () => go(btn.dataset.go);
});
window.addEventListener("hashchange", render);
render();
