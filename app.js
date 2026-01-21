/*
  Школьное расписание
  - Роли: admin, deputy, teacher, user.
*/

const ROLE_LABEL = {
  admin: "Администратор",
  deputy: "Завуч",
  teacher: "Учитель",
  user: "Пользователь",
  public: "Публичный просмотр",
};

const CAN_EDIT_SCHEDULE = new Set(["admin", "deputy"]);
const CAN_EDIT_SETTINGS = new Set(["admin"]);

// Supabase авторизация

function getPasswords() {
  try {
    const raw = localStorage.getItem("sched_passwords");
    const obj = raw ? JSON.parse(raw) : null;
    return {
      ...PASSWORDS_DEFAULT,
      ...(obj && typeof obj === "object" ? obj : {}),
    };
  } catch {
    return { ...PASSWORDS_DEFAULT };
  }
}

function setPasswords(pwObj) {
  const safe = {
    admin: String(pwObj?.admin ?? PASSWORDS_DEFAULT.admin),
    deputy: String(pwObj?.deputy ?? PASSWORDS_DEFAULT.deputy),
    teacher: String(pwObj?.teacher ?? PASSWORDS_DEFAULT.teacher),
  };
  localStorage.setItem("sched_passwords", JSON.stringify(safe));
}

function $(sel, root = document) {
  return root.querySelector(sel);
}
function $all(sel, root = document) {
  return Array.from(root.querySelectorAll(sel));
}
function toast(msg, type = "info") {
  const el = $("#toast");
  if (!el) return;
  el.textContent = msg;
  el.classList.remove("ok", "warn", "err");
  el.classList.add(type === "ok" ? "ok" : type === "err" ? "err" : "warn");
  el.classList.add("show");
  window.clearTimeout(toast._t);
  toast._t = window.setTimeout(() => el.classList.remove("show"), 2800);
}

// ------------------------------
// IndexedDB: хранение fileHandle
// ------------------------------
const IDB_NAME = "school_timetable";
const IDB_STORE = "handles";

function idbOpen() {
  return new Promise((resolve, reject) => {
    const req = indexedDB.open(IDB_NAME, 1);
    req.onupgradeneeded = () => {
      const db = req.result;
      if (!db.objectStoreNames.contains(IDB_STORE)) db.createObjectStore(IDB_STORE);
    };
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error);
  });
}
async function idbSet(key, val) {
  const db = await idbOpen();
  return new Promise((resolve, reject) => {
    const tx = db.transaction(IDB_STORE, "readwrite");
    tx.objectStore(IDB_STORE).put(val, key);
    tx.oncomplete = () => resolve();
    tx.onerror = () => reject(tx.error);
  });
}
async function idbGet(key) {
  const db = await idbOpen();
  return new Promise((resolve, reject) => {
    const tx = db.transaction(IDB_STORE, "readonly");
    const req = tx.objectStore(IDB_STORE).get(key);
    req.onsuccess = () => resolve(req.result ?? null);
    req.onerror = () => reject(req.error);
  });
}

// ------------------------------
// Auth
// ------------------------------
function getSession() {
  try {
    return JSON.parse(localStorage.getItem("sched_session") || "null");
  } catch {
    return null;
  }
}
function setSession(role) {
  localStorage.setItem("sched_session", JSON.stringify({ role, ts: Date.now() }));
}
function clearSession() {
  localStorage.removeItem("sched_session");
}

// ------------------------------
// DB IO
// ------------------------------
let dbData = null;
let dbHandle = null;

// Supabase helpers
function sbEnabled() {
  return !!window.sbApi?.getSbConfig?.();
}

async function loadData() {
  if (sbEnabled()) {
    try {
      const data = await window.sbApi.sbLoadSchoolData();
      dbData = data ? data : structuredClone(DB_DEFAULT);
      return;
    } catch (e) {
      console.error(e);
      dbData = structuredClone(DB_DEFAULT);
      toast("Supabase недоступен — загружены демо-данные", "warn");
      return;
    }
  }
  await loadDbAuto();
}

async function saveData() {
  if (sbEnabled()) {
    await window.sbApi.sbSaveSchoolData(dbData);
    toast("Сохранено в Supabase", "ok");
    return;
  }
  await saveDb();
}

async function openDbPicker() {
  if (!window.showOpenFilePicker) {
    toast("Этот браузер не поддерживает запись в JSON-файл. Используйте Edge/Chrome.", "err");
    throw new Error("File System Access API not supported");
  }
  const [handle] = await window.showOpenFilePicker({
    types: [{ description: "JSON", accept: { "application/json": [".json"] } }],
    multiple: false,
  });
  const perm = await handle.requestPermission({ mode: "readwrite" });
  if (perm !== "granted") throw new Error("No permission");
  await idbSet("dbHandle", handle);
  dbHandle = handle;
  await loadDbFromHandle();
  return handle;
}

async function loadDbFromHandle() {
  if (!dbHandle) throw new Error("No handle");
  const file = await dbHandle.getFile();
  const text = await file.text();
  try {
    dbData = JSON.parse(text);
  } catch {
    dbData = structuredClone(DB_DEFAULT);
  }

  // миграция/синхронизация паролей
  dbData.meta ??= {};
  dbData.meta.passwords ??= null;
  const pwLocal = getPasswords();
  const pwDb = dbData.meta.passwords && typeof dbData.meta.passwords === "object" ? dbData.meta.passwords : null;
  if (pwDb) {
    // база — источник истины
    setPasswords({ ...pwLocal, ...pwDb });
  } else {
    // в базе нет — записываем из localStorage/дефолтов
    dbData.meta.passwords = pwLocal;
  }
}

async function loadDbAuto() {
  // Если Supabase настроен — пробуем грузить оттуда.
  try {
    const sbApi = window.sbApi;
    if (sbApi?.getSbConfig?.()) {
      const data = await sbApi.sbLoadSchoolData();
      if (data) {
        dbData = data;
        dbHandle = null;
        return;
      }
    }
  } catch (e) {
    console.warn("Supabase load failed, fallback to local", e);
  }

  dbHandle = await idbGet("dbHandle");
  if (!dbHandle) {
    dbData = structuredClone(DB_DEFAULT);
    return;
  }
  // В некоторых случаях после перезапуска потребуется заново запросить права.
  const perm = await dbHandle.queryPermission({ mode: "readwrite" });
  if (perm !== "granted") {
    const perm2 = await dbHandle.requestPermission({ mode: "readwrite" });
    if (perm2 !== "granted") {
      dbData = structuredClone(DB_DEFAULT);
      dbHandle = null;
      return;
    }
  }
  await loadDbFromHandle();
}

async function saveDb() {
  // Приоритет — Supabase
  try {
    const sbApi = window.sbApi;
    if (sbApi?.getSbConfig?.()) {
      await sbApi.sbSaveSchoolData(dbData);
      toast("Сохранено в Supabase", "ok");
      return;
    }
  } catch (e) {
    console.error(e);
    toast("Ошибка сохранения в Supabase", "err");
    return;
  }

  // Фолбэк: локальный JSON (старый режим)
  if (!dbHandle) {
    toast("Supabase не настроен. Для локального режима сначала выберите db.json (кнопка «Импорт JSON»)", "warn");
    return;
  }
  const writable = await dbHandle.createWritable();
  await writable.write(JSON.stringify(dbData, null, 2));
  await writable.close();
  toast("Сохранено в db.json", "ok");
}

// ------------------------------
// Helpers: data access
// ------------------------------
function getActiveTermId() {
  return dbData?.meta?.activeTermId || dbData?.meta?.terms?.[0]?.id;
}
function getTermById(termId) {
  return (dbData?.meta?.terms || []).find((t) => t.id === termId) || null;
}
function ensurePath(termId, classId, day) {
  dbData.timetable ??= {};
  dbData.timetable[termId] ??= {};
  dbData.timetable[termId][classId] ??= {};
  dbData.timetable[termId][classId][day] ??= {};
  return dbData.timetable[termId][classId][day];
}
function getCell(termId, classId, day, lessonNo) {
  const t = dbData.timetable?.[termId]?.[classId]?.[day]?.[String(lessonNo)];
  return t || null;
}

// Ячейка может быть:
// 1) объектом {subject, teacherId, room, note} (старый формат)
// 2) массивом таких объектов (подгруппы в одном уроке)
function normalizeCell(cell) {
  if (!cell) return [];
  if (Array.isArray(cell)) return cell.filter(Boolean);
  if (Array.isArray(cell.items)) return cell.items.filter(Boolean);
  if (typeof cell === "object") return [cell];
  return [];
}

function denormalizeCell(items) {
  const clean = (items || [])
    .map((x) => ({
      subject: (x?.subject || "").trim(),
      teacherId: (x?.teacherId || "").trim(),
      room: (x?.room || "").trim(),
      note: (x?.note || "").trim(),
    }))
    .filter((x) => x.subject);
  if (!clean.length) return null;
  if (clean.length === 1) return clean[0];
  return clean;
}
function setCell(termId, classId, day, lessonNo, valOrNull) {
  const dayObj = ensurePath(termId, classId, day);
  const key = String(lessonNo);
  if (!valOrNull) {
    delete dayObj[key];
  } else {
    dayObj[key] = valOrNull;
  }
}
function teacherNameById(id) {
  return dbData.settings.teachers.find((t) => t.id === id)?.name || "";
}

function fmtSubjectNote(x) {
  const s = (x?.subject || "").trim();
  const n = (x?.note || "").trim();
  return n ? `${s} (${n})` : s;
}

// ------------------------------
// Sorting helpers (RU locale)
// ------------------------------
function ruSort(a, b) {
  return String(a || "").localeCompare(String(b || ""), "ru", { sensitivity: "base" });
}

// Сортировка классов: сначала номер (1..11), затем буква (А, Б, В...), затем остаток.
function parseClassName(name) {
  const s = String(name || "").trim().toUpperCase();
  const m = s.match(/^(\d+)\s*([А-ЯA-Z]?)(.*)$/u);
  if (!m) return { n: 999, l: s, rest: "" };
  return { n: parseInt(m[1], 10), l: m[2] || "", rest: m[3] || "" };
}
function classSort(a, b) {
  const A = parseClassName(a?.name ?? a);
  const B = parseClassName(b?.name ?? b);
  if (A.n !== B.n) return A.n - B.n;
  const l = A.l.localeCompare(B.l, "ru", { sensitivity: "base" });
  if (l !== 0) return l;
  return (A.rest || "").localeCompare(B.rest || "", "ru", { sensitivity: "base" });
}

// ------------------------------
// Login page
// ------------------------------
function initLogin() {
  const emailInp = $("#emailInput");
  const passInp = $("#passwordInput");
  const btn = $("#btnLogin");
  const btnGuest = $("#btnGuest");

  btn?.addEventListener("click", async () => {
    const sbApi = window.sbApi;
    const cfgOk = sbApi?.getSbConfig?.();
    if (!cfgOk) {
      toast("Supabase не настроен: создайте config.js (см. config.example.js)", "err");
      return;
    }
    const email = (emailInp?.value || "").trim();
    const password = passInp?.value || "";
    if (!email || !password) {
      toast("Введите email и пароль", "warn");
      return;
    }
    try {
      const { error } = await sbApi.sbSignIn(email, password);
      if (error) {
        toast("Не удалось войти: " + (error.message || "ошибка"), "err");
        return;
      }
      location.href = "dashboard.html";
    } catch (e) {
      console.error(e);
      toast("Ошибка входа", "err");
    }
  });

  btnGuest?.addEventListener("click", () => {
    // Публичный просмотр без входа
    location.href = "dashboard.html";
  });
}

// ------------------------------
// Dashboard page
// ------------------------------
let state = {
  role: "user",
  termId: null,
  classId: null,
  viewMode: "byClass", // byClass | byTeacher
  teacherId: null,
};

function initDashboard() {
  const sbApi = window.sbApi;
  const cfgOk = sbApi?.getSbConfig?.();
  const sbStatus = $("#sbStatus");
  if (sbStatus) sbStatus.textContent = cfgOk ? "· Supabase" : "· локальный режим";

  // По умолчанию: публичный просмотр
  state.role = "public";
  $("#roleBadge").textContent = ROLE_LABEL[state.role] || state.role;

  $("#btnLogin")?.addEventListener("click", () => {
    location.href = "login.html";
  });

  $("#btnLogout")?.addEventListener("click", async () => {
    try {
      await sbApi?.sbSignOut?.();
    } finally {
      location.reload();
    }
  });


  // Bootstrap: роль + загрузка данных
  (async () => {
    let loggedIn = false;
    let role = "public";

    if (cfgOk) {
      const sess = await sbApi.sbGetSession();
      loggedIn = !!sess?.session;
      if (loggedIn) {
        const prof = await sbApi.sbGetMyProfile();
        // если профиль не заведен, считаем ролью teacher (просмотр)
        role = prof?.role || "teacher";
      }
    }

    state.role = role;
    $("#roleBadge").textContent = ROLE_LABEL[state.role] || state.role;

    const isEditor = CAN_EDIT_SCHEDULE.has(state.role);
    $("#editHint").textContent = isEditor ? "Режим: редактирование" : "Режим: просмотр";

    // teacher: по умолчанию показываем расписание учителя
    if (state.role === "teacher") {
      state.viewMode = "byTeacher";
      const vm = $("#viewMode");
      if (vm) vm.value = "byTeacher";
    }

    updateEditorUi();

    await loadDbAuto();
    hydrateControls();
    renderSchedule();
    updateEditorUi();
  })();

  $("#btnOpenDb").addEventListener("click", async () => {
    // Импорт JSON разрешён только редакторам (admin/deputy)
    if (!CAN_EDIT_SCHEDULE.has(state.role) && !CAN_EDIT_SETTINGS.has(state.role)) {
      toast("Импорт доступен только после входа (завуч/админ)", "err");
      return;
    }
    try {
      await openDbPicker();
      // Если Supabase настроен — сразу пушим импортированную базу в БД.
      const cfgOk = window.sbApi?.getSbConfig?.();
      if (cfgOk) {
        await saveDb();
      }
      hydrateControls();
      renderSchedule();
      updateEditorUi();
      toast(cfgOk ? "Импортировано в Supabase" : "База открыта", "ok");
    } catch (e) {
      console.error(e);
      toast("Не удалось открыть файл", "err");
    }
  });

  $("#btnSave").addEventListener("click", async () => {
    if (!CAN_EDIT_SCHEDULE.has(state.role) && !CAN_EDIT_SETTINGS.has(state.role)) {
      toast("Нет прав на сохранение", "err");
      return;
    }
    try {
      await saveDb();
    } catch (e) {
      console.error(e);
      toast("Ошибка сохранения", "err");
    }
  });

  $("#btnExport").addEventListener("click", exportExcel);

  $("#viewMode").addEventListener("change", (e) => {
    state.viewMode = e.target.value;
    renderSchedule();
    updateEditorUi();
  });

  $("#termSelect").addEventListener("change", (e) => {
    state.termId = e.target.value;
    renderSchedule();
  });

  $("#classSelect").addEventListener("change", (e) => {
    state.classId = e.target.value;
    renderSchedule();
  });

  $("#teacherSelect").addEventListener("change", (e) => {
    state.teacherId = e.target.value;
    renderSchedule();
  });

  $("#btnSettings").addEventListener("click", () => openSettingsModal());
  $("#btnManage").addEventListener("click", () => openManageModal());

  // modal close
  $("#modal").addEventListener("click", (e) => {
    if (e.target.id === "modal") closeModal();
  });

  // кнопка ✕ в заголовке модалки (если есть)
  $("#modalX")?.addEventListener("click", closeModal);
}


function updateEditorUi() {
  const isEditor = CAN_EDIT_SCHEDULE.has(state.role);
  const canSettings = CAN_EDIT_SETTINGS.has(state.role);
  $("#btnSave").style.display = (isEditor || canSettings) ? "inline-flex" : "none";
  $("#btnManage").style.display = isEditor ? "inline-flex" : "none";
  $("#btnSettings").style.display = canSettings ? "inline-flex" : "none";

  // Импорт JSON — только редактор (и только как ручной импорт/миграция)
  const imp = $("#btnOpenDb");
  if (imp) imp.style.display = (isEditor || canSettings) ? "inline-flex" : "none";

  // Вход/выход
  const btnLogin = $("#btnLogin");
  const btnLogout = $("#btnLogout");
  const isPublic = state.role === "public";
  if (btnLogin) btnLogin.style.display = isPublic ? "inline-flex" : "none";
  if (btnLogout) btnLogout.style.display = isPublic ? "none" : "inline-flex";

  // view controls
  const byClass = state.viewMode === "byClass";
  $("#classSelect").closest(".field").style.display = byClass ? "block" : "none";
  $("#teacherSelect").closest(".field").style.display = byClass ? "none" : "block";
}

function hydrateControls() {
  // school title
  const st = $("#schoolTitle");
  if (st) st.textContent = (dbData?.meta?.schoolName || "Школа") + " — расписание";

  // terms (sorted A–Я)
  const termSelect = $("#termSelect");
  termSelect.innerHTML = "";
  const termsSorted = [...(dbData.meta.terms || [])].sort((a, b) => ruSort(a?.name, b?.name));
  for (const t of termsSorted) {
    const opt = document.createElement("option");
    opt.value = t.id;
    opt.textContent = t.name || t.id;
    termSelect.appendChild(opt);
  }
  state.termId = state.termId || getActiveTermId();
  termSelect.value = state.termId;
  // classes (sorted: number then letter)
  const classSelect = $("#classSelect");
  classSelect.innerHTML = "";
  const classesSorted = [...(dbData.settings.classes || [])].sort(classSort);
  for (const c of classesSorted) {
    const opt = document.createElement("option");
    opt.value = c.id;
    opt.textContent = c.name || c.id;
    classSelect.appendChild(opt);
  }
  state.classId = state.classId || classesSorted[0]?.id || null;
  if (state.classId) classSelect.value = state.classId;
  // teachers (sorted A–Я)
  const teacherSelect = $("#teacherSelect");
  teacherSelect.innerHTML = "";
  const optAll = document.createElement("option");
  optAll.value = "";
  optAll.textContent = "— Выберите учителя —";
  teacherSelect.appendChild(optAll);

  const teachersSorted = [...(dbData.settings.teachers || [])].sort((a, b) => ruSort(a?.name, b?.name));
  for (const t of teachersSorted) {
    const opt = document.createElement("option");
    opt.value = t.id;
    opt.textContent = t.name || t.id;
    teacherSelect.appendChild(opt);
  }
  // default view mode for teacher role
  if (state.role === "teacher") {
    state.viewMode = "byTeacher";
    $("#viewMode").value = "byTeacher";
  }

  updateEditorUi();
}

function renderSchedule() {
  const mount = $("#scheduleMount");
  mount.innerHTML = "";

  if (!dbData) {
    mount.innerHTML = `<div class="card"><div class="muted">Нет данных</div></div>`;
    return;
  }

  const termId = state.termId || getActiveTermId();
  const settings = dbData.settings;
  const days = settings.days;
  const lessonNos = settings.bellSchedule.map((b) => b.lesson);

  if (state.viewMode === "byTeacher") {
    if (!state.teacherId) {
      mount.innerHTML = `<div class="card"><div class="muted">Выберите учителя, чтобы увидеть его расписание.</div></div>`;
      return;
    }
    mount.appendChild(renderTeacherTable(termId, state.teacherId, days, lessonNos));
    return;
  }

  if (!state.classId) {
    mount.innerHTML = `<div class="card"><div class="muted">Добавьте классы в разделе «Классы/учителя».</div></div>`;
    return;
  }
  mount.appendChild(renderClassTable(termId, state.classId, days, lessonNos));
}

function renderClassTable(termId, classId, days, lessonNos) {
  const c = dbData.settings.classes.find((x) => x.id === classId);
  const isEditor = CAN_EDIT_SCHEDULE.has(state.role);

  const card = document.createElement("div");
  card.className = "card";

  const header = document.createElement("div");
  header.className = "row";
  header.innerHTML = `
    <div>
      <div class="h">Расписание: ${escapeHtml(c?.name || classId)}</div>
      <div class="muted">Полугодие: ${escapeHtml(getTermById(termId)?.name || termId)}</div>
    </div>
    <div class="pill">Каб: ${escapeHtml(c?.room ?? "—")}, Кл.рук.: ${escapeHtml(teacherNameById(c?.classTeacherId) || "—")}</div>
  `;
  card.appendChild(header);

  const tableWrap = document.createElement("div");
  tableWrap.className = "tableWrap";

  const table = document.createElement("table");
  table.className = "sched";

  const thead = document.createElement("thead");
  thead.innerHTML = `
    <tr>
      <th class="col-lesson">Урок</th>
      <th class="col-time">Время</th>
      ${days.map((d) => `<th>${escapeHtml(d)}</th>`).join("")}
    </tr>
  `;
  table.appendChild(thead);

  const tbody = document.createElement("tbody");
  for (const lessonNo of lessonNos) {
    const bell = dbData.settings.bellSchedule.find((b) => b.lesson === lessonNo);
    const tr = document.createElement("tr");

    const tdN = document.createElement("td");
    tdN.className = "col-lesson";
    tdN.textContent = String(lessonNo);

    const tdT = document.createElement("td");
    tdT.className = "col-time";
    const label = bell?.label ? ` (${bell.label})` : "";
    tdT.textContent = bell ? `${bell.start}–${bell.end}${label}` : "";

    tr.appendChild(tdN);
    tr.appendChild(tdT);

    for (const day of days) {
      const td = document.createElement("td");
      const cell = getCell(termId, classId, day, lessonNo);
      td.innerHTML = renderClassCellHtml(cell);
      const tip = buildClassTooltip(cell);
      if (tip) td.title = tip;
      td.dataset.term = termId;
      td.dataset.class = classId;
      td.dataset.day = day;
      td.dataset.lesson = String(lessonNo);
      td.classList.toggle("editable", isEditor);
      if (isEditor) {
        td.addEventListener("click", () => openEditCellModal(td));
      }
      tr.appendChild(td);
    }

    tbody.appendChild(tr);
  }

  table.appendChild(tbody);
  tableWrap.appendChild(table);
  card.appendChild(tableWrap);
  return card;
}

function renderTeacherTable(termId, teacherId, days, lessonNos) {
  const isEditor = false; // учитель всегда только просмотр

  const card = document.createElement("div");
  card.className = "card";
  card.innerHTML = `
    <div class="row">
      <div>
        <div class="h">Расписание учителя: ${escapeHtml(teacherNameById(teacherId) || teacherId)}</div>
        <div class="muted">Полугодие: ${escapeHtml(getTermById(termId)?.name || termId)}</div>
      </div>
      <div class="pill">Экспорт доступен</div>
    </div>
  `;

  // Собираем матрицу: день x урок -> список занятий (учитываем подгруппы)
  const items = {};
  for (const day of days) items[day] = {};

  const classes = dbData.settings.classes;
  const classesSorted = [...classes].sort(classSort);
  for (const c of classes) {
    for (const day of days) {
      for (const lessonNo of lessonNos) {
        const raw = getCell(termId, c.id, day, lessonNo);
        const parts = normalizeCell(raw);
        for (const p of parts) {
          if ((p?.teacherId || "") !== teacherId) continue;
          items[day][lessonNo] ??= [];
          items[day][lessonNo].push({
            subjectNote: fmtSubjectNote(p),
            className: c.name || c.id,
            room: (p?.room || "").trim(),
          });
        }
      }
    }
  }

  const tableWrap = document.createElement("div");
  tableWrap.className = "tableWrap";

  const table = document.createElement("table");
  table.className = "sched";

  table.innerHTML = `
    <thead>
      <tr>
        <th class="col-lesson">Урок</th>
        <th class="col-time">Время</th>
        ${days.map((d) => `<th>${escapeHtml(d)}</th>`).join("")}
      </tr>
    </thead>
  `;

  const tbody = document.createElement("tbody");
  for (const lessonNo of lessonNos) {
    const bell = dbData.settings.bellSchedule.find((b) => b.lesson === lessonNo);
    const tr = document.createElement("tr");

    const tdN = document.createElement("td");
    tdN.className = "col-lesson";
    tdN.textContent = String(lessonNo);

    const tdT = document.createElement("td");
    tdT.className = "col-time";
    const label = bell?.label ? ` (${bell.label})` : "";
    tdT.textContent = bell ? `${bell.start}–${bell.end}${label}` : "";

    tr.appendChild(tdN);
    tr.appendChild(tdT);

    for (const day of days) {
      const td = document.createElement("td");
      const list = items[day][lessonNo] || [];
      td.innerHTML = list.length
        ? list
            .map(
              (x) =>
                `<div class="tcell-line"><span class="tcell-subj">${escapeHtml(x.subjectNote)}</span> <span class="tcell-class">${escapeHtml(x.className)}</span></div>`
            )
            .join("")
        : "";
      const tip = buildTeacherTooltip(list);
      if (tip) td.title = tip;
      td.classList.toggle("editable", isEditor);
      tr.appendChild(td);
    }

    tbody.appendChild(tr);
  }
  table.appendChild(tbody);

  tableWrap.appendChild(table);
  card.appendChild(tableWrap);
  return card;
}

function renderClassCellHtml(raw) {
  const parts = normalizeCell(raw);
  if (!parts.length) return "";
  const main = parts.map((x) => fmtSubjectNote(x)).join(" / ");
  return `<div class="cell"><div class="subject">${escapeHtml(main)}</div></div>`;
}

function buildClassTooltip(raw) {
  const parts = normalizeCell(raw);
  if (!parts.length) return "";
  return parts
    .map((x) => {
      const sn = fmtSubjectNote(x);
      const tn = (x?.teacherId || "") ? teacherNameById(x.teacherId) : "";
      const rm = (x?.room || "").trim();
      const tail = [tn, rm ? `каб. ${rm}` : ""].filter(Boolean).join(" · ");
      return tail ? `${sn} — ${tail}` : sn;
    })
    .join("\n");
}

function buildTeacherTooltip(list) {
  if (!list?.length) return "";
  return list
    .map((x) => {
      const lines = [x.subjectNote, x.className];
      if (x.room) lines.push(`каб. ${x.room}`);
      return lines.join("\n");
    })
    .join("\n\n");
}

// ------------------------------
// Modal system
// ------------------------------
function openModal(title, bodyHtml, footerHtml) {
  $("#modalTitle").textContent = title;
  $("#modalBody").innerHTML = bodyHtml;
  $("#modalFooter").innerHTML = footerHtml || "";
  // styles.css toggles visibility via `.modal.show`
  $("#modal").classList.add("show");
}
function closeModal() {
  $("#modal").classList.remove("show");
}

function openEditCellModal(td) {
  const termId = td.dataset.term;
  const classId = td.dataset.class;
  const day = td.dataset.day;
  const lessonNo = parseInt(td.dataset.lesson, 10);
  const parts0 = normalizeCell(getCell(termId, classId, day, lessonNo));
  const parts = parts0.length ? structuredClone(parts0) : [{ subject: "", teacherId: "", room: "", note: "" }];

  const subjSorted = [...(dbData.settings.subjects || [])].sort(ruSort);
  const subjOpts = ["", ...subjSorted]
    .map((s) => `<option value="${escapeAttr(s)}">${escapeHtml(s || "—")}</option>`)
    .join("");
  const teachersSorted = [...(dbData.settings.teachers || [])].sort((a, b) => ruSort(a?.name, b?.name));
  const teachOpts = [{ id: "", name: "—" }, ...teachersSorted]
    .map((t) => `<option value="${escapeAttr(t.id)}">${escapeHtml(t.name || t.id || "—")}</option>`)
    .join("");
  const rowHtml = (idx, cur) => `
    <div class="grp" data-idx="${idx}">
      <div class="grid2" style="gap:10px">
        <div>
          <div class="label">Предмет</div>
          <select class="gSubject">${subjOpts}</select>
        </div>
        <div>
          <div class="label">Учитель</div>
          <select class="gTeacher">${teachOpts}</select>
        </div>
        <div>
          <div class="label">Кабинет</div>
          <input class="gRoom" type="text" placeholder="например 12"/>
        </div>
        <div>
          <div class="label">Примечание (в скобках)</div>
          <input class="gNote" type="text" placeholder="например: подгруппа"/>
        </div>
      </div>
      <div style="margin-top:8px; display:flex; gap:8px; justify-content:flex-end">
        <button class="btn ghost sm" data-act="delGrp">Удалить подгруппу</button>
      </div>
      <div class="hr"></div>
    </div>
  `;

  openModal(
    `Редактирование: ${escapeHtml(day)} · урок ${lessonNo}`,
    `
      <div class="muted" style="margin-bottom:10px">
        Можно добавить 2–4 подгруппы в один урок. В расписании класса будет показано: <b>Предмет (примечание) / …</b>.
        Наведите на ячейку — увидите подробности.
      </div>
      <div id="mGroups">${parts.map((p, i) => rowHtml(i, p)).join("")}</div>
      <div style="display:flex; gap:10px; align-items:center; justify-content:space-between; margin-top:8px">
        <button class="btn sm" id="addGrp">+ Подгруппа</button>
        <div class="muted">Пустой предмет в подгруппе игнорируется.</div>
      </div>
    `,
    `
      <button class="btn ghost" id="mCancel">Отмена</button>
      <button class="btn danger" id="mClear">Очистить</button>
      <button class="btn" id="mApply">Применить</button>
    `
  );

  function syncRowValues() {
    $all("#mGroups .grp").forEach((grp, idx) => {
      const cur = parts[idx] || (parts[idx] = { subject: "", teacherId: "", room: "", note: "" });
      const subj = $(".gSubject", grp);
      const tch = $(".gTeacher", grp);
      const rm = $(".gRoom", grp);
      const nt = $(".gNote", grp);
      subj.value = cur.subject || "";
      tch.value = cur.teacherId || "";
      rm.value = cur.room || "";
      nt.value = cur.note || "";
    });
  }

  function readRows() {
    const out = [];
    $all("#mGroups .grp").forEach((grp) => {
      out.push({
        subject: $(".gSubject", grp).value,
        teacherId: $(".gTeacher", grp).value,
        room: $(".gRoom", grp).value.trim(),
        note: $(".gNote", grp).value.trim(),
      });
    });
    return out;
  }

  function rebuild() {
    const mount = $("#mGroups");
    mount.innerHTML = parts.map((p, i) => rowHtml(i, p)).join("");
    syncRowValues();
    updateDeleteButtons();
  }

  function updateDeleteButtons() {
    const canDel = $all("#mGroups .grp").length > 1;
    $all("#mGroups button[data-act='delGrp']").forEach((b) => (b.disabled = !canDel));
    $("#addGrp").disabled = $all("#mGroups .grp").length >= 4;
  }

  $("#mGroups").addEventListener("click", (e) => {
    const btn = e.target.closest("button[data-act]");
    if (!btn) return;
    if (btn.dataset.act === "delGrp") {
      const grp = btn.closest(".grp");
      const idx = parseInt(grp.dataset.idx, 10);
      parts.splice(idx, 1);
      rebuild();
    }
  });

  $("#addGrp").addEventListener("click", () => {
    if (parts.length >= 4) return;
    parts.push({ subject: "", teacherId: "", room: "", note: "" });
    rebuild();
  });

  $("#mCancel").addEventListener("click", closeModal);
  $("#mClear").addEventListener("click", () => {
    setCell(termId, classId, day, lessonNo, null);
    td.innerHTML = "";
    td.title = "";
    closeModal();
  });
  $("#mApply").addEventListener("click", async () => {
    const val = denormalizeCell(readRows());
    setCell(termId, classId, day, lessonNo, val);
    td.innerHTML = renderClassCellHtml(val);
    td.title = buildClassTooltip(val);
    closeModal();
  });

  // init
  rebuild();
}


function openSettingsModal() {
  if (!CAN_EDIT_SETTINGS.has(state.role)) {
    toast("Нет прав", "err");
    return;
  }

  const s = dbData.settings;
  const sbApi = window.sbApi;
  const hasSupabase = !!sbApi?.getSbConfig?.();

  const bellRows = (s.bellSchedule || [])
    .sort((a, b) => a.lesson - b.lesson)
    .map(
      (b, idx) => `
      <tr>
        <td>${escapeHtml(String(b.lesson))}</td>
        <td><input data-k="start" data-idx="${idx}" value="${escapeAttr(b.start)}"/></td>
        <td><input data-k="end" data-idx="${idx}" value="${escapeAttr(b.end)}"/></td>
        <td><input data-k="label" data-idx="${idx}" value="${escapeAttr(b.label || "")}" placeholder="опционально"/></td>
      </tr>
    `
    )
    .join("");

  const subjRows = (s.subjects || [])
    .map(
      (name, i) => `
      <tr>
        <td><input class="subjInp" data-i="${i}" value="${escapeAttr(String(name))}"/></td>
        <td style="width:1%"><button class="btn danger sm" data-act="delSubj" data-i="${i}">Удалить</button></td>
      </tr>
    `
    )
    .join("");

  const terms = dbData.meta.terms || [];
  const termRows = terms
    .map(
      (t, i) => `
      <tr data-i="${i}">
        <td><input class="termId" value="${escapeAttr(t.id || "")}" /></td>
        <td><input class="termName" value="${escapeAttr(t.name || "")}" /></td>
        <td><input class="termStart" value="${escapeAttr(t.start || "")}" placeholder="YYYY-MM-DD" /></td>
        <td><input class="termEnd" value="${escapeAttr(t.end || "")}" placeholder="YYYY-MM-DD" /></td>
        <td style="width:1%"><button class="btn danger sm" data-act="delTerm" data-i="${i}">Удалить</button></td>
      </tr>
    `
    )
    .join("");

  openModal(
    "Настройки (только администратор)",
    `
      <div class="tabs">
        <button class="tab active" data-tab="general">Общее</button>
        <button class="tab" data-tab="terms">Полугодия</button>
        <button class="tab" data-tab="subjects">Предметы</button>
        <button class="tab" data-tab="passwords">Пароль</button>
      </div>

      <div class="tabPanel" id="tab-general">
        <div class="grid2">
          <div>
            <div class="label">Название школы</div>
            <input id="setSchool" value="${escapeAttr(dbData.meta.schoolName || "")}"/>
          </div>
          <div>
            <div class="label">Кол-во уроков в день (без нулевого)</div>
            <input id="setLessons" type="number" min="1" max="12" value="${escapeAttr(String(s.lessonsPerDay || 7))}"/>
          </div>
        </div>

        <div style="margin-top:12px" class="label">Время звонков</div>
        <div class="tableWrap" style="max-height:340px">
          <table class="mini">
            <thead><tr><th>Урок</th><th>Начало</th><th>Конец</th><th>Метка</th></tr></thead>
            <tbody>${bellRows}</tbody>
          </table>
        </div>
        <div class="muted" style="margin-top:10px">После изменения нажмите «Применить», затем «Сохранить базу».</div>
      </div>

      <div class="tabPanel" id="tab-terms" style="display:none">
        <div class="grid2">
          <div>
            <div class="label">Полугодие по умолчанию</div>
            <select id="activeTermSelect"></select>
            <div class="muted" style="margin-top:6px">
              Это полугодие будет открываться по умолчанию на главной странице.
            </div>
          </div>
          <div></div>
        </div>

        <div class="label" style="margin-top:12px">Список полугодий</div>
        <div class="tableWrap" style="max-height:340px">
          <table class="mini" id="termTable">
            <thead><tr><th style="width:18%">ID</th><th>Название</th><th style="width:18%">Начало</th><th style="width:18%">Конец</th><th></th></tr></thead>
            <tbody>${termRows}</tbody>
          </table>
        </div>
        <div style="margin-top:10px"><button class="btn sm" id="addTerm">+ Добавить полугодие</button></div>
      </div>

      <div class="tabPanel" id="tab-subjects" style="display:none">
        <div class="muted" style="margin:6px 0 10px">Список предметов используется в редакторе расписания.</div>
        <div class="tableWrap" style="max-height:340px">
          <table class="mini" id="subjTable">
            <thead><tr><th>Предмет</th><th></th></tr></thead>
            <tbody>${subjRows}</tbody>
          </table>
        </div>
        <div style="margin-top:10px"><button class="btn sm" id="addSubj">+ Добавить предмет</button></div>
      </div>

      <div class="tabPanel" id="tab-passwords" style="display:none">
        <div class="muted" style="margin:6px 0 10px">
          Если используется Supabase, здесь можно сменить <b>свой</b> пароль (вход по email).
        </div>

        ${hasSupabase ? `
          <div class="grid2">
            <div>
              <div class="label">Новый пароль</div>
              <input id="pwNew" type="password" placeholder="Новый пароль"/>
            </div>
            <div>
              <div class="label">Повторите пароль</div>
              <input id="pwNew2" type="password" placeholder="Повтор"/>
            </div>
          </div>
          <div class="muted" style="margin-top:10px">Чтобы применить, нажмите «Применить».</div>
        ` : `
          <div class="muted">Supabase не настроен (нет config.js). Смена пароля недоступна.</div>
        `}
      </div>
    `,
    `
      <button class="btn ghost" id="mClose">Закрыть</button>
      <button class="btn" id="mApply">Применить</button>
    `
  );

  // tabs
  $all(".tab").forEach((b) => {
    b.addEventListener("click", () => {
      $all(".tab").forEach((x) => x.classList.remove("active"));
      b.classList.add("active");
      const tab = b.dataset.tab;
      $("#tab-general").style.display = tab === "general" ? "block" : "none";
      $("#tab-terms").style.display = tab === "terms" ? "block" : "none";
      $("#tab-subjects").style.display = tab === "subjects" ? "block" : "none";
      $("#tab-passwords").style.display = tab === "passwords" ? "block" : "none";
    });
  });

  // init active term select (sorted A–Я)
  const activeSel = $("#activeTermSelect");
  activeSel.innerHTML = "";
  const termsSorted = [...(dbData.meta.terms || [])].sort((a, b) => ruSort(a?.name, b?.name));
  for (const t of termsSorted) {
    const opt = document.createElement("option");
    opt.value = t.id;
    opt.textContent = t.name || t.id;
    activeSel.appendChild(opt);
  }
  activeSel.value = dbData.meta.activeTermId || termsSorted[0]?.id || "";

  // add/delete subjects (reopen to keep indices simple)
  $("#addSubj").addEventListener("click", () => {
    s.subjects ??= [];
    s.subjects.push("Новый предмет");
    closeModal();
    openSettingsModal();
  });

  $("#addTerm").addEventListener("click", () => {
    dbData.meta.terms ??= [];
    const id = `term_${Date.now()}`;
    dbData.meta.terms.push({ id, name: "Новое полугодие", start: "", end: "" });
    dbData.meta.activeTermId = id;
    closeModal();
    openSettingsModal();
  });

  $("#modalBody").addEventListener("click", (e) => {
    const btn = e.target.closest("button[data-act]");
    if (!btn) return;

    if (btn.dataset.act === "delSubj") {
      const i = parseInt(btn.dataset.i, 10);
      s.subjects.splice(i, 1);
      closeModal();
      openSettingsModal();
      return;
    }

    if (btn.dataset.act === "delTerm") {
      const i = parseInt(btn.dataset.i, 10);
      const removed = (dbData.meta.terms || []).splice(i, 1)[0];
      // если удалили активное — переключаем на первое доступное
      if (removed?.id && dbData.meta.activeTermId === removed.id) {
        dbData.meta.activeTermId = (dbData.meta.terms || [])[0]?.id || "";
      }
      closeModal();
      openSettingsModal();
      return;
    }
  });

  $("#mClose").addEventListener("click", closeModal);

  $("#mApply").addEventListener("click", async () => {
    // general
    dbData.meta.schoolName = $("#setSchool").value.trim() || "Школа";
    const lessons = Math.max(1, Math.min(12, parseInt($("#setLessons").value || "7", 10)));
    s.lessonsPerDay = lessons;

    // ensure bellSchedule has 0..lessons
    const existing = new Map((s.bellSchedule || []).map((b) => [b.lesson, b]));
    const next = [];
    for (let i = 0; i <= lessons; i++) {
      next.push(existing.get(i) || { lesson: i, start: "00:00", end: "00:00", label: i === 0 ? "Кл. час" : "" });
    }
    s.bellSchedule = next;

    $all("#tab-general table.mini input").forEach((inp) => {
      const idx = parseInt(inp.dataset.idx, 10);
      const k = inp.dataset.k;
      s.bellSchedule[idx][k] = inp.value.trim();
    });

    // subjects (unique + sorted A–Я)
    const newSubjects = $all("#tab-subjects .subjInp")
      .map((inp) => inp.value.trim())
      .filter(Boolean);
    if (newSubjects.length) {
      s.subjects = Array.from(new Set(newSubjects)).sort(ruSort);
    }

    // terms
    const termTrs = $all("#termTable tbody tr");
    const newTerms = [];
    for (const tr of termTrs) {
      const id = $(".termId", tr)?.value?.trim() || "";
      const name = $(".termName", tr)?.value?.trim() || "";
      const start = $(".termStart", tr)?.value?.trim() || "";
      const end = $(".termEnd", tr)?.value?.trim() || "";
      if (!id) continue;
      newTerms.push({ id, name: name || id, start, end });
    }

    if (!newTerms.length) {
      toast("Должно быть хотя бы одно полугодие", "err");
      return;
    }

    dbData.meta.terms = newTerms;

    const active = $("#activeTermSelect")?.value || "";
    dbData.meta.activeTermId = newTerms.some((t) => t.id === active) ? active : newTerms[0].id;

    // password (Supabase only): change my password if fields filled
    if (hasSupabase) {
      const p1 = ($("#pwNew")?.value || "").trim();
      const p2 = ($("#pwNew2")?.value || "").trim();
      if (p1 || p2) {
        if (!p1 || p1 !== p2) {
          toast("Пароли не совпадают", "err");
          return;
        }
        try {
          const { error } = await sbApi.sbUpdateMyPassword(p1);
          if (error) {
            toast("Ошибка смены пароля: " + (error.message || "ошибка"), "err");
            return;
          }
          toast("Пароль успешно изменён", "ok");
        } catch (e) {
          console.error(e);
          toast("Ошибка смены пароля", "err");
          return;
        }
      }
    }

    closeModal();
    hydrateControls();
    renderSchedule();
    toast("Настройки применены", "ok");
  });
}


function openManageModal() {
  if (!CAN_EDIT_SCHEDULE.has(state.role)) {
    toast("Нет прав", "err");
    return;
  }

  const classes = dbData.settings.classes;
  const teachers = dbData.settings.teachers;

  const classRows = classes
    .map(
      (c, i) => `
      <tr>
        <td><input data-type="class" data-i="${i}" data-k="name" value="${escapeAttr(c.name)}"/></td>
        <td><input data-type="class" data-i="${i}" data-k="id" value="${escapeAttr(c.id)}"/></td>
        <td><input data-type="class" data-i="${i}" data-k="room" value="${escapeAttr(c.room || "")}"/></td>
        <td>
          <select data-type="class" data-i="${i}" data-k="classTeacherId">
            <option value="">—</option>
            ${teachers
              .map((t) => `<option value="${escapeAttr(t.id)}" ${t.id === c.classTeacherId ? "selected" : ""}>${escapeHtml(t.name)}</option>`)
              .join("")}
          </select>
        </td>
        <td><button class="btn danger sm" data-act="delClass" data-i="${i}">Удалить</button></td>
      </tr>
    `
    )
    .join("");

  const teacherRows = teachers
    .map(
      (t, i) => `
      <tr>
        <td><input data-type="teacher" data-i="${i}" data-k="name" value="${escapeAttr(t.name)}"/></td>
        <td><input data-type="teacher" data-i="${i}" data-k="id" value="${escapeAttr(t.id)}"/></td>
        <td><button class="btn danger sm" data-act="delTeacher" data-i="${i}">Удалить</button></td>
      </tr>
    `
    )
    .join("");

  openModal(
    "Классы и учителя (Администратор/Завуч)",
    `
      <div class="tabs">
        <button class="tab active" data-tab="classes">Классы</button>
        <button class="tab" data-tab="teachers">Учителя</button>
      </div>

      <div class="tabPanel" id="tab-classes">
        <div class="muted" style="margin:6px 0 10px">Изменения применяются после «Применить».</div>
        <div class="tableWrap" style="max-height:340px">
          <table class="mini">
            <thead><tr><th>Название</th><th>ID</th><th>Каб</th><th>Кл.рук.</th><th></th></tr></thead>
            <tbody>${classRows}</tbody>
          </table>
        </div>
        <div style="margin-top:10px"><button class="btn sm" id="addClass">+ Добавить класс</button></div>
      </div>

      <div class="tabPanel" id="tab-teachers" style="display:none">
        <div class="muted" style="margin:6px 0 10px">ID должен быть уникальным (например t12).</div>
        <div class="tableWrap" style="max-height:340px">
          <table class="mini">
            <thead><tr><th>ФИО</th><th>ID</th><th></th></tr></thead>
            <tbody>${teacherRows}</tbody>
          </table>
        </div>
        <div style="margin-top:10px"><button class="btn sm" id="addTeacher">+ Добавить учителя</button></div>
      </div>
    `,
    `
      <button class="btn ghost" id="mClose">Закрыть</button>
      <button class="btn" id="mApply">Применить</button>
    `
  );

  // tabs
  $all(".tab").forEach((b) => {
    b.addEventListener("click", () => {
      $all(".tab").forEach((x) => x.classList.remove("active"));
      b.classList.add("active");
      const tab = b.dataset.tab;
      $("#tab-classes").style.display = tab === "classes" ? "block" : "none";
      $("#tab-teachers").style.display = tab === "teachers" ? "block" : "none";
    });
  });

  // add row
  $("#addClass").addEventListener("click", () => {
    dbData.settings.classes.push({ id: `C${Date.now()}`, name: "Новый класс", room: "", classTeacherId: "" });
    closeModal();
    openManageModal();
  });
  $("#addTeacher").addEventListener("click", () => {
    dbData.settings.teachers.push({ id: `t${Date.now()}`, name: "Новый учитель" });
    closeModal();
    openManageModal();
  });

  // delete row
  $("#modalBody").addEventListener("click", (e) => {
    const btn = e.target.closest("button[data-act]");
    if (!btn) return;
    const i = parseInt(btn.dataset.i, 10);
    if (btn.dataset.act === "delClass") {
      dbData.settings.classes.splice(i, 1);
      closeModal();
      openManageModal();
    }
    if (btn.dataset.act === "delTeacher") {
      const removed = dbData.settings.teachers.splice(i, 1)[0];
      // убрать ссылки на учителя из классов и расписания
      for (const c of dbData.settings.classes) {
        if (c.classTeacherId === removed.id) c.classTeacherId = "";
      }
      // в расписании: teacherId = "" where matched
      for (const term of Object.values(dbData.timetable || {})) {
        for (const classObj of Object.values(term || {})) {
          for (const dayObj of Object.values(classObj || {})) {
            for (const cell of Object.values(dayObj || {})) {
              if (cell?.teacherId === removed.id) cell.teacherId = "";
            }
          }
        }
      }
      closeModal();
      openManageModal();
    }
  });

  $("#mClose").addEventListener("click", closeModal);
  $("#mApply").addEventListener("click", () => {
    // apply edits from inputs
    $all("#modalBody input[data-type], #modalBody select[data-type]").forEach((el) => {
      const type = el.dataset.type;
      const i = parseInt(el.dataset.i, 10);
      const k = el.dataset.k;
      const v = el.value.trim();
      if (type === "class") dbData.settings.classes[i][k] = v;
      if (type === "teacher") dbData.settings.teachers[i][k] = v;
    });

    // de-duplicate IDs
    dbData.settings.classes = uniqById(dbData.settings.classes, "id");
    dbData.settings.teachers = uniqById(dbData.settings.teachers, "id");

    closeModal();
    hydrateControls();
    renderSchedule();
    toast("Изменения применены", "ok");
  });
}

function uniqById(arr, key) {
  const seen = new Set();
  const out = [];
  for (const item of arr) {
    const id = (item[key] || "").trim();
    if (!id || seen.has(id)) continue;
    seen.add(id);
    out.push(item);
  }
  return out;
}

// ------------------------------
// Export to Excel (ExcelJS)
// ------------------------------
async function exportExcel() {
  if (!window.ExcelJS) {
    toast("ExcelJS не загружен", "err");
    return;
  }
  if (!dbData) {
    toast("Нет данных", "err");
    return;
  }

  const termId = state.termId || getActiveTermId();
  const term = getTermById(termId);
  const s = dbData.settings;

  // Экспорт делаем в "формате как пример":
  // - строка 14: названия классов по колонкам D..
  // - строка 15: "Каб" и кабинеты
  // - строка 16: "К.р." и классные руководители
  // - блоки дней: в колонке B (merged), в колонке C номера уроков (0..N)

  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet("Расписание");

  const classes = s.classes;
  const days = s.days;
  const lessonNos = s.bellSchedule.map((b) => b.lesson);

  // Column widths approximated from the sample
  ws.getColumn(1).width = 13; // A
  ws.getColumn(2).width = 25; // B
  ws.getColumn(3).width = 22; // C
  // classes columns D..
  for (let i = 0; i < classes.length; i++) {
    ws.getColumn(4 + i).width = 18;
  }

  // Header rows start at 14
  const rClass = 14;
  const rRoom = 15;
  const rCT = 16;

  // classes names
  for (let i = 0; i < classes.length; i++) {
    ws.getCell(rClass, 4 + i).value = classes[i].name;
    ws.getCell(rRoom, 4 + i).value = classes[i].room || "";
    ws.getCell(rCT, 4 + i).value = teacherNameById(classes[i].classTeacherId) || "";
  }

  ws.getCell(rRoom, 3).value = "Каб";
  ws.getCell(rCT, 3).value = "К.р.";

  // Basic styles
  const border = {
    top: { style: "thin" },
    left: { style: "thin" },
    bottom: { style: "thin" },
    right: { style: "thin" },
  };
  const headerFill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF1B2A4A" } };
  const dayFill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF13213D" } };

  function applyCellStyle(cell, opts = {}) {
    cell.border = border;
    cell.alignment = { vertical: "middle", horizontal: "center", wrapText: true, ...opts.alignment };
    if (opts.fill) cell.fill = opts.fill;
    if (opts.font) cell.font = opts.font;
  }

  // Style header
  for (let col = 3; col < 4 + classes.length; col++) {
    for (let row of [rClass, rRoom, rCT]) {
      const cell = ws.getCell(row, col);
      applyCellStyle(cell, {
        fill: headerFill,
        font: { color: { argb: "FFFFFFFF" }, bold: true },
      });
    }
  }
  // class name row a bit taller
  ws.getRow(rClass).height = 60;
  ws.getRow(rRoom).height = 40;
  ws.getRow(rCT).height = 45;

  // Day blocks
  let row = 17;
  for (const day of days) {
    const startRow = row;
    // merge B for day over number of lessons
    const blockRows = lessonNos.length;
    ws.mergeCells(startRow, 2, startRow + blockRows - 1, 2);
    ws.getCell(startRow, 2).value = day;
    applyCellStyle(ws.getCell(startRow, 2), {
      fill: dayFill,
      font: { color: { argb: "FFFFFFFF" }, bold: true, size: 12 },
      alignment: { horizontal: "center" },
    });

    for (let i = 0; i < lessonNos.length; i++) {
      const lessonNo = lessonNos[i];
      const r = startRow + i;
      ws.getRow(r).height = 45;

      // lesson number in C
      const cLesson = ws.getCell(r, 3);
      cLesson.value = lessonNo;
      applyCellStyle(cLesson, { fill: dayFill, font: { color: { argb: "FFFFFFFF" }, bold: true } });

      // cells per class
      for (let ci = 0; ci < classes.length; ci++) {
        const cl = classes[ci];
        const raw = getCell(termId, cl.id, day, lessonNo);
        const parts = normalizeCell(raw);
        const text = parts.length
          ? parts
              .map((p) => {
                const a = [
                  fmtSubjectNote(p),
                  (p?.teacherId || "") ? teacherNameById(p.teacherId) : "",
                  (p?.room || "").trim() ? `каб. ${(p.room || "").trim()}` : "",
                ].filter(Boolean);
                return a.join("\n");
              })
              .join("\n\n/\n\n")
          : "";
        const xc = ws.getCell(r, 4 + ci);
        xc.value = text;
        applyCellStyle(xc, { alignment: { horizontal: "center" } });
      }
    }

    // border & fill for merged B cells below start
    for (let r = startRow; r < startRow + blockRows; r++) {
      const cb = ws.getCell(r, 2);
      cb.border = border;
      cb.alignment = { vertical: "middle", horizontal: "center", wrapText: true };
      cb.fill = dayFill;
      cb.font = { color: { argb: "FFFFFFFF" }, bold: true };
    }

    row += blockRows;
  }

  // Add title (optional) at top
  ws.getCell(1, 1).value = `${dbData.meta.schoolName || "Школа"} — расписание (${term?.name || termId})`;
  ws.mergeCells(1, 1, 1, Math.max(4, 3 + classes.length));
  ws.getRow(1).height = 24;
  ws.getCell(1, 1).font = { bold: true, size: 14, color: { argb: "FFFFFFFF" } };
  ws.getCell(1, 1).alignment = { vertical: "middle", horizontal: "left" };

  // Freeze panes like Excel
  ws.views = [{ state: "frozen", xSplit: 3, ySplit: 16 }];

  const buf = await wb.xlsx.writeBuffer();
  const blob = new Blob([buf], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
  const a = document.createElement("a");
  const fname = `Расписание_${termId}.xlsx`;
  a.href = URL.createObjectURL(blob);
  a.download = fname;
  document.body.appendChild(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(a.href), 1500);
}

// ------------------------------
// Utils
// ------------------------------
function escapeHtml(s) {
  return String(s ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#39;");
}
function escapeAttr(s) {
  return escapeHtml(s).replaceAll("\n", " ");
}

// Boot
document.addEventListener("DOMContentLoaded", () => {
  if (document.body.dataset.page === "login") initLogin();
  if (document.body.dataset.page === "dashboard") initDashboard();
});
