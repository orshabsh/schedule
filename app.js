/*
  Школьное расписание (локально, без сервера)
  - Данные в JSON-файле (File System Access API).
  - Роли: admin, deputy, teacher, user.

  ВАЖНО:
  В браузере нельзя "тихо" записать файл на диск без явного доступа пользователя.
  Поэтому при первом запуске нужно выбрать db.json через кнопку "Открыть базу".
*/

const ROLE_LABEL = {
  admin: "Администратор",
  deputy: "Завуч",
  teacher: "Учитель",
  user: "Пользователь",
};

const CAN_EDIT_SCHEDULE = new Set(["admin", "deputy"]);
const CAN_EDIT_SETTINGS = new Set(["admin"]);

// Демо-пароли (для локального развёртывания).
// Если вы хотите вынести их в отдельный файл — можно.
const PASSWORDS = {
  admin: "1779",
  deputy: "1346",
  teacher: "1234",
};

const DB_DEFAULT = {
  meta: {
    schoolName: "Школа",
    activeTermId: "2025_H1",
    terms: [
      { id: "2025_H1", name: "1 полугодие 2025/26", start: "2025-09-01", end: "2025-12-31" },
      { id: "2025_H2", name: "2 полугодие 2025/26", start: "2026-01-09", end: "2026-05-31" },
    ],
  },
  settings: {
    lessonsPerDay: 7,
    days: ["Понедельник", "Вторник", "Среда", "Четверг", "Пятница"],
    bellSchedule: [
      { lesson: 0, start: "08:00", end: "08:30", label: "Кл. час" },
      { lesson: 1, start: "08:35", end: "09:20" },
      { lesson: 2, start: "09:30", end: "10:15" },
      { lesson: 3, start: "10:25", end: "11:10" },
      { lesson: 4, start: "11:20", end: "12:05" },
      { lesson: 5, start: "12:25", end: "13:10" },
      { lesson: 6, start: "13:20", end: "14:05" },
      { lesson: 7, start: "14:15", end: "15:00" },
    ],
    classes: [
      { id: "5A", name: "5 «А»", room: "5", classTeacherId: "t1" },
      { id: "5B", name: "5 «Б»", room: "3", classTeacherId: "t2" },
    ],
    teachers: [
      { id: "t1", name: "Орлова И.Н." },
      { id: "t2", name: "Мильто Ю.П." },
      { id: "t3", name: "Шукалович С.И." },
    ],
    subjects: [
      "Математика",
      "Русский язык",
      "Белорусский язык",
      "Английский язык",
      "ФК и З",
      "Информатика",
      "ОБЖ",
      "Кл. час",
      "Трудовое обучение",
      "История",
      "Биология",
    ],
  },
  timetable: {
    
  },
};

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
}

async function loadDbAuto() {
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
  if (!dbHandle) {
    toast("Сначала откройте db.json (кнопка «Открыть базу»)", "warn");
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

// ------------------------------
// Login page
// ------------------------------
function initLogin() {
  const roleSel = $("#roleSelect");
  const passWrap = $("#passwordWrap");
  const passInp = $("#passwordInput");
  const btn = $("#btnLogin");

  const update = () => {
    const role = roleSel.value;
    const needPass = role !== "user";
    passWrap.style.display = needPass ? "block" : "none";
    if (!needPass) passInp.value = "";
  };

  roleSel.addEventListener("change", update);
  update();

  btn.addEventListener("click", () => {
    const role = roleSel.value;
    if (role === "user") {
      setSession("user");
      location.href = "dashboard.html";
      return;
    }
    const p = passInp.value || "";
    if (p !== PASSWORDS[role]) {
      toast("Неверный пароль", "err");
      return;
    }
    setSession(role);
    location.href = "dashboard.html";
  });

  $("#btnGuest")?.addEventListener("click", () => {
    setSession("user");
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
  const session = getSession();
  if (!session?.role) {
    location.href = "login.html";
    return;
  }
  state.role = session.role;
  $("#roleBadge").textContent = ROLE_LABEL[state.role] || state.role;

  $("#btnLogout").addEventListener("click", () => {
    clearSession();
    location.href = "login.html";
  });

  // teacher/user: скрываем элементы управления редактированием
  const isEditor = CAN_EDIT_SCHEDULE.has(state.role);
  $("#editHint").textContent = isEditor
    ? "Режим: редактирование разрешено"
    : "Режим: только просмотр";

  // загрузка базы
  loadDbAuto().then(() => {
    hydrateControls();
    renderSchedule();
    updateEditorUi();
  });

  $("#btnOpenDb").addEventListener("click", async () => {
    try {
      await openDbPicker();
      hydrateControls();
      renderSchedule();
      updateEditorUi();
      toast("База открыта", "ok");
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
  
  autoLoadDb();
}

async function autoLoadDb() {
  try {
    const res = await fetch("db.json", { cache: "no-store" });
    if (!res.ok) throw new Error("db.json not found");

    const data = await res.json();
    window.DB = data;
    renderAll();
    toast("База данных загружена автоматически");
  } catch (e) {
    toast("db.json не найден — используйте «Открыть базу»", true);
  }
}


function updateEditorUi() {
  const isEditor = CAN_EDIT_SCHEDULE.has(state.role);
  const canSettings = CAN_EDIT_SETTINGS.has(state.role);
  $("#btnSave").style.display = (isEditor || canSettings) ? "inline-flex" : "none";
  $("#btnManage").style.display = isEditor ? "inline-flex" : "none";
  $("#btnSettings").style.display = canSettings ? "inline-flex" : "none";

  // view controls
  const byClass = state.viewMode === "byClass";
  $("#classSelect").closest(".field").style.display = byClass ? "block" : "none";
  $("#teacherSelect").closest(".field").style.display = byClass ? "none" : "block";
}

function hydrateControls() {
  // school title
  const st = $("#schoolTitle");
  if (st) st.textContent = (dbData?.meta?.schoolName || "Школа") + " — расписание";

  // terms
  const termSelect = $("#termSelect");
  termSelect.innerHTML = "";
  for (const t of dbData.meta.terms) {
    const opt = document.createElement("option");
    opt.value = t.id;
    opt.textContent = t.name;
    termSelect.appendChild(opt);
  }
  state.termId = state.termId || getActiveTermId();
  termSelect.value = state.termId;

  // classes
  const classSelect = $("#classSelect");
  classSelect.innerHTML = "";
  for (const c of dbData.settings.classes) {
    const opt = document.createElement("option");
    opt.value = c.id;
    opt.textContent = c.name;
    classSelect.appendChild(opt);
  }
  state.classId = state.classId || dbData.settings.classes[0]?.id || null;
  if (state.classId) classSelect.value = state.classId;

  // teachers
  const teacherSelect = $("#teacherSelect");
  teacherSelect.innerHTML = "";
  const optAll = document.createElement("option");
  optAll.value = "";
  optAll.textContent = "— Выберите учителя —";
  teacherSelect.appendChild(optAll);
  for (const t of dbData.settings.teachers) {
    const opt = document.createElement("option");
    opt.value = t.id;
    opt.textContent = t.name;
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
      td.innerHTML = renderCellText(cell);
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

  // Собираем матрицу: день x урок -> список занятий
  const items = {};
  for (const day of days) items[day] = {};

  const classes = dbData.settings.classes;
  for (const c of classes) {
    for (const day of days) {
      for (const lessonNo of lessonNos) {
        const cell = getCell(termId, c.id, day, lessonNo);
        if (cell?.teacherId === teacherId) {
          const line = `${c.name}: ${cell.subject}${cell.room ? `, каб. ${cell.room}` : ""}`;
          items[day][lessonNo] ??= [];
          items[day][lessonNo].push(line);
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
      td.innerHTML = list.length ? list.map(escapeHtml).join("<br>") : "";
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

function renderCellText(cell) {
  if (!cell) return "";
  const subject = escapeHtml(cell.subject || "");
  const t = cell.teacherId ? escapeHtml(teacherNameById(cell.teacherId)) : "";
  const room = cell.room ? escapeHtml(String(cell.room)) : "";
  const note = cell.note ? escapeHtml(String(cell.note)) : "";
  return `
    <div class="cell">
      <div class="cell-main">${subject || "—"}</div>
      <div class="cell-sub">${[t, room ? `каб. ${room}` : "", note].filter(Boolean).join(" · ")}</div>
    </div>
  `;
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
  const cur = getCell(termId, classId, day, lessonNo) || { subject: "", teacherId: "", room: "", note: "" };

  const subjOpts = ["", ...dbData.settings.subjects]
    .map((s) => `<option value="${escapeAttr(s)}" ${s === cur.subject ? "selected" : ""}>${escapeHtml(s || "—")}</option>`)
    .join("");
  const teachOpts = ["", ...dbData.settings.teachers.map((t) => t.id)]
    .map((id) => {
      const name = id ? teacherNameById(id) : "—";
      return `<option value="${escapeAttr(id)}" ${id === cur.teacherId ? "selected" : ""}>${escapeHtml(name)}</option>`;
    })
    .join("");

  openModal(
    `Редактирование: ${escapeHtml(day)} · урок ${lessonNo}`,
    `
      <div class="grid2">
        <div>
          <div class="label">Предмет</div>
          <select id="mSubject">${subjOpts}</select>
        </div>
        <div>
          <div class="label">Учитель</div>
          <select id="mTeacher">${teachOpts}</select>
        </div>
        <div>
          <div class="label">Кабинет</div>
          <input id="mRoom" type="text" value="${escapeAttr(cur.room || "")}" placeholder="например 12"/>
        </div>
        <div>
          <div class="label">Примечание (опционально)</div>
          <input id="mNote" type="text" value="${escapeAttr(cur.note || "")}" placeholder="например: подгруппа"/>
        </div>
      </div>
      <div class="muted" style="margin-top:10px">Подсказка: пустой предмет = очистить ячейку.</div>
    `,
    `
      <button class="btn ghost" id="mCancel">Отмена</button>
      <button class="btn danger" id="mClear">Очистить</button>
      <button class="btn" id="mApply">Применить</button>
    `
  );

  $("#mCancel").addEventListener("click", closeModal);
  $("#mClear").addEventListener("click", () => {
    setCell(termId, classId, day, lessonNo, null);
    td.innerHTML = "";
    closeModal();
  });
  $("#mApply").addEventListener("click", () => {
    const subject = $("#mSubject").value;
    const teacherId = $("#mTeacher").value;
    const room = $("#mRoom").value.trim();
    const note = $("#mNote").value.trim();

    if (!subject) {
      setCell(termId, classId, day, lessonNo, null);
      td.innerHTML = "";
      closeModal();
      return;
    }

    const val = {
      subject,
      teacherId: teacherId || "",
      room: room || "",
      note: note || "",
    };
    setCell(termId, classId, day, lessonNo, val);
    td.innerHTML = renderCellText(val);
    closeModal();
  });
}

function openSettingsModal() {
  if (!CAN_EDIT_SETTINGS.has(state.role)) {
    toast("Нет прав", "err");
    return;
  }
  const s = dbData.settings;
  const rows = s.bellSchedule
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

  openModal(
    "Настройки (только администратор)",
    `
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
          <tbody>${rows}</tbody>
        </table>
      </div>
      <div class="muted" style="margin-top:10px">После изменения нажмите «Сохранить».</div>
    `,
    `
      <button class="btn ghost" id="mClose">Закрыть</button>
      <button class="btn" id="mApply">Применить</button>
    `
  );

  $("#mClose").addEventListener("click", closeModal);
  $("#mApply").addEventListener("click", () => {
    dbData.meta.schoolName = $("#setSchool").value.trim() || "Школа";
    const lessons = Math.max(1, Math.min(12, parseInt($("#setLessons").value || "7", 10)));
    s.lessonsPerDay = lessons;

    // ensure bellSchedule has 0..lessons
    const existing = new Map(s.bellSchedule.map((b) => [b.lesson, b]));
    const next = [];
    for (let i = 0; i <= lessons; i++) {
      next.push(existing.get(i) || { lesson: i, start: "00:00", end: "00:00", label: i === 0 ? "Кл. час" : "" });
    }
    s.bellSchedule = next;

    $all("table.mini input").forEach((inp) => {
      const idx = parseInt(inp.dataset.idx, 10);
      const k = inp.dataset.k;
      s.bellSchedule[idx][k] = inp.value.trim();
    });

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
        const cell = getCell(termId, cl.id, day, lessonNo);
        const text = cell ? [cell.subject, teacherNameById(cell.teacherId), cell.room ? `каб. ${cell.room}` : ""]
          .filter(Boolean)
          .join("\n") : "";
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
