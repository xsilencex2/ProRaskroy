import * as XLSX from "xlsx";

// --- Material thicknesses ---
const MATERIAL_THICKNESSES = {
  "ДСП": 16,
  "МДФ": 18,
  "Фанера": 18,
  "ХДФ": 3
};

// --- Utils ---
function formatNumber(v, digits=1) {
  return (Math.round(v * Math.pow(10, digits)) / Math.pow(10, digits)).toLocaleString('ru-RU');
}

function validateDimension(n, min=5, max=5000) {
  return Number.isFinite(n) && n >= min && n <= max;
}

function validateThickness(t, min=1, max=40) {
  return Number.isFinite(t) && t >= min && t <= max;
}

// --- DOM helpers ---
function el(sel) { return document.querySelector(sel); }

// Helper to create edge select (with current 'Кромка' value)
function createEdgeSelect(currentValue, edgeOptionsValue) {
  const select = document.createElement('select');
  select.name = 'edge-side';
  const val = typeof currentValue === "string" ? currentValue : "";
  // Если нет кромки — одно "Нет" (пустое) значение
  if (!edgeOptionsValue) {
    // "Нет" (пусто)
    const optNone = document.createElement('option');
    optNone.value = '';
    optNone.textContent = 'Нет';
    select.appendChild(optNone);
    select.value = '';
    return select;
  }
  // Есть кромка → "Нет" и кромка
  [
    {value: '', label: 'Нет'},
    {value: edgeOptionsValue, label: edgeOptionsValue}
  ].forEach(optData => {
    const opt = document.createElement('option');
    opt.value = optData.value;
    opt.textContent = optData.label;
    select.appendChild(opt);
  });
  select.value = val;
  return select;
}

// Modified createPartRow to correctly pass Кромка value from current sheet settings
function createPartRow(idx, part = {}, edgeOptionsValue = "") {
  const tr = document.createElement('tr');
  tr.dataset.idx = idx;

  function input(name, value, type="number", extra={}) {
    const inp = document.createElement('input');
    inp.type = type;
    inp.name = name;
    inp.value = value ?? '';
    if (type === "number") {
      inp.min = extra.min ?? 5;
      inp.max = extra.max ?? 5000;
      inp.step = extra.step ?? 1;
    }
    inp.required = true;
    if (extra.placeholder) inp.placeholder = extra.placeholder;
    if (extra.maxLength) inp.maxLength = extra.maxLength;
    return inp;
  }

  tr.innerHTML = `
    <td>${idx + 1}</td>
    <td></td><td></td><td></td><td></td><td></td><td></td><td></td><td></td>
    <td>
      <button type="button" class="delete-part-btn" title="Удалить">🗑️</button>
    </td>
  `;
  // Длина
  tr.children[1].appendChild(input('length', part.length, 'number', { min:5, max:5000 }));
  // Ширина
  tr.children[2].appendChild(input('width', part.width, 'number', { min:5, max:5000 }));
  // Кол-во
  tr.children[3].appendChild(input('count', part.count ?? 1, 'number', { min:1, max:500 }));

  // TXT (текстура) — чекбокс
  const txtBox = document.createElement('input');
  txtBox.type = 'checkbox';
  txtBox.name = 'texture';
  txtBox.checked = !!part.texture;
  txtBox.style.transform = 'scale(1.2)';
  txtBox.style.cursor = 'pointer';
  tr.children[4].appendChild(txtBox);

  // Вверх/Вниз/Лево/Право — селект с вариантами: нет, кромка
  ["top", "bottom", "left", "right"].forEach((side, i) => {
    const sideValue = part[side] || '';
    const select = createEdgeSelect(sideValue, edgeOptionsValue);
    select.dataset.side = side;
    tr.children[5 + i].appendChild(select);
  });

  return tr;
}

// --- State ---
const state = {
  user: {
    firstname: "",
    lastname: "",
    phone: ""
  },
  sheet: {
    width: 2800,
    height: 2070,
    material: "",
    thickness: 16,
    edge: "",
    edgeThickness: "1 мм"
  },
  parts: []
};

// --- Sheet Param Controls ---
const sheetSel = el('#sheet-size-select');
const customSheetDiv = el('#custom-sheet-size');
const sheetWidthInp = el('#sheet-width');
const sheetHeightInp = el('#sheet-height');
const sheetMaterialInp = el('#sheet-material');
const sheetThicknessInp = el('#sheet-thickness');
const sheetEdgeInp = el('#sheet-edge');
const edgeThicknessSel = el('#edge-thickness');

function syncSheetFromInputs() {
  state.sheet.width = parseInt(sheetWidthInp.value) || 2800;
  state.sheet.height = parseInt(sheetHeightInp.value) || 2070;
  state.sheet.material = sheetMaterialInp.value.trim();
  let th = parseFloat(sheetThicknessInp.value);
  if (!validateThickness(th)) th = 16;
  state.sheet.thickness = th;
  state.sheet.edge = sheetEdgeInp.value.trim();
  state.sheet.edgeThickness = edgeThicknessSel.value;
}

sheetSel.addEventListener('change', () => {
  if (sheetSel.value === "custom") {
    customSheetDiv.style.display = "";
    sheetWidthInp.disabled = false; sheetHeightInp.disabled = false;
  } else {
    let w = 2800, h = 2070;
    if (sheetSel.value === "2500x1830") { w = 2500; h = 1830; }
    customSheetDiv.style.display = "none";
    sheetWidthInp.value = w;
    sheetHeightInp.value = h;
    sheetWidthInp.disabled = true; sheetHeightInp.disabled = true;
  }
  syncSheetFromInputs();
  repaintEverything();
  repaintPartsTable(); // update edge selects in parts
});
[sheetWidthInp, sheetHeightInp, sheetMaterialInp, sheetThicknessInp, sheetEdgeInp, edgeThicknessSel].forEach(inp=>
  inp.addEventListener('input', () => {
    syncSheetFromInputs();
    repaintEverything();
    repaintPartsTable(); // update edge selects in parts if "Кромка" changed
  })
);
edgeThicknessSel.addEventListener('change', () => {
  syncSheetFromInputs();
  repaintEverything();
  repaintPartsTable();
});

// --- Parts Table Logic ---
const partsTbody = el('#parts-tbody');
function makePartObj(tr) {
  const length = parseInt(tr.querySelector('[name=length]').value);
  const width = parseInt(tr.querySelector('[name=width]').value);
  const count = parseInt(tr.querySelector('[name=count]').value);
  const texture = tr.querySelector('[name=texture]').checked ? "да" : "";
  // Вверх/Вниз/Лево/Право теперь из селектов edge-side
  const selects = tr.querySelectorAll('select[name="edge-side"]');
  const sides = {};
  selects.forEach(sel => {
    const side = sel.dataset.side;
    sides[side] = sel.value ?? '';
  });
  return {
    length,
    width,
    count,
    texture,
    top: sides.top,
    bottom: sides.bottom,
    left: sides.left,
    right: sides.right
  };
}

function repaintPartsTable() {
  partsTbody.innerHTML = '';
  // Передаем текущую кромку
  const edgeOpt = (state.sheet && state.sheet.edge) ? state.sheet.edge : "";
  state.parts.forEach((part, idx) => {
    const tr = createPartRow(idx, part, edgeOpt);
    partsTbody.appendChild(tr);
  });
}
function updatePartsFromTable() {
  const trs = partsTbody.querySelectorAll('tr');
  state.parts = [];
  trs.forEach(tr=>{
    state.parts.push(makePartObj(tr));
  });
}

function validateParts() {
  let ok = true, msg = '';
  state.parts.forEach((p, idx) => {
    if (!validateDimension(p.length) || !validateDimension(p.width)) {
      ok = false;
      msg = `Неверный размер детали №${idx+1}: длина и ширина должны быть 5–5000 мм.`;
    }
    if (!Number.isFinite(p.count) || p.count < 1) {
      ok = false;
      msg = `Неверное количество у детали №${idx+1}.`;
    }
  });
  return ok ? null : msg;
}

function addPart(part) {
  if (!part) {
    part = { length: 300, width: 300, count: 1, texture: "", top: "", bottom: "", left: "", right: "" };
  }
  state.parts.push(part);
  repaintPartsTable();
  repaintEverything();
}
el('#add-part-btn').addEventListener('click', () => addPart());

// Delegate change events for inputs in parts table:
partsTbody.addEventListener('input', ()=>{
  updatePartsFromTable();
  repaintEverything();
});
partsTbody.addEventListener('change', ()=>{
  updatePartsFromTable();
  repaintEverything();
});
partsTbody.addEventListener('click', (e)=>{
  if (e.target.classList.contains('delete-part-btn')) {
    const idx = parseInt(e.target.closest('tr').dataset.idx,10);
    state.parts.splice(idx,1);
    repaintPartsTable();
    repaintEverything();
  }
});

// --- Calculation ---
function getAllUsedArea() {
  // в mm^2
  return state.parts.reduce((acc,p)=>acc+(p.length*p.width*p.count),0);
}
function getSheetArea() {
  return state.sheet.width * state.sheet.height;
}
function wastePercent() {
  const used = getAllUsedArea();
  const total = getSheetArea();
  if (!used || !total) return 0;
  return Math.max(0, 100 - (used/total)*100);
}

function repaintEverything() {
  // Площадь и процент отходов
  const mm2toM2 = 1e-6;
  const usedArea = getAllUsedArea()*mm2toM2;
  el('#used-area').textContent = formatNumber(usedArea, 2);

  const waste = wastePercent();
  el('#waste-percent').textContent = formatNumber(waste, 1);
}

// Формирование Excel
function exportToExcel() {
  // Сбор данных пользователя
  const fname = (state.user.firstname || '').trim();
  const lname = (state.user.lastname || '').trim();
  const phone = (state.user.phone || '').trim();

  // Заголовок, подзаголовок, служебная строка
  const mainTitle = "Шаблон для покраски";
  const subTitle = "Текстура: X - вдоль, Y - поперёк. Положение кромки указывается для каждой стороны.";
  const contactInfo = [
    `Имя: ${fname}  Фамилия: ${lname}  Телефон: ${phone}`
  ];

  // Столбцы шаблона (по фото + текстура, положение кромки)
  const header = [
    "№",
    "Длина (мм)",
    "Ширина (мм)",
    "Кол-во",
    "TXT (текстура)",
    "Вверх",
    "Вниз",
    "Лево",
    "Право"
  ];

  // Заполнение строк шаблона для деталей
  const rows = state.parts.map((p, idx) => [
    idx+1,
    p.length || "",
    p.width || "",
    p.count || "",
    p.texture ? '✔' : "",
    p.top || "",
    p.bottom || "",
    p.left || "",
    p.right || ""
  ]);
  
  const sheetInfo = [
    [
      `Размер листа: ${state.sheet.width} x ${state.sheet.height} мм, материал: ${state.sheet.material}, кромка: ${state.sheet.edge} (${state.sheet.edgeThickness}), толщина: ${state.sheet.thickness} мм`
    ]
  ];

  // Сбор всех данных
  const ws_data = [];
  ws_data.push([mainTitle]);
  ws_data.push([subTitle]);
  ws_data.push(contactInfo);
  ws_data.push([]); // пустая строка
  ws_data.push(header);
  rows.forEach(row => ws_data.push(row));
  ws_data.push([]);
  ws_data.push(...sheetInfo);

  // формирование листа Excel
  const ws = XLSX.utils.aoa_to_sheet(ws_data);

  // Стилизация шаблона: четкие "квадратики"
  if (!ws['!merges']) ws['!merges'] = [];
  // A1:I1, A2:I2, A3:I3 (название и заголовки объединить)
  ws["!merges"].push({s: {r: 0, c: 0}, e: {r: 0, c: header.length}});
  ws["!merges"].push({s: {r: 1, c: 0}, e: {r: 1, c: header.length}});
  ws["!merges"].push({s: {r: 2, c: 0}, e: {r: 2, c: header.length}});

  const borderStyle = {
    top: {style: "thin", color: {rgb:"000000"}},
    bottom: {style: "thin", color: {rgb:"000000"}},
    left: {style: "thin", color: {rgb:"000000"}},
    right: {style: "thin", color: {rgb:"000000"}}
  };

  // Обвожу рамками строки таблицы данных
  const firstDataRow = 4, lastDataRow = 4+rows.length;
  for (let r = firstDataRow; r <= lastDataRow; r++) {
    for (let c = 0; c < header.length; c++) {
      const cellRef = XLSX.utils.encode_cell({r, c});
      if (!ws[cellRef]) continue;
      ws[cellRef].s = {
        border: borderStyle,
        alignment: {horizontal: "center", vertical:"center", wrapText: true},
        font: r===firstDataRow ? {bold:true} : {}
      };
    }
  }

  // Размер строк
  ws['!rows'] = [];
  ws['!rows'][0] = {hpt: 22};
  ws['!rows'][1] = {hpt: 18};
  ws['!rows'][2] = {hpt: 15};
  ws['!rows'][firstDataRow] = {hpt: 23}; // header
  for(let r=firstDataRow+1; r<=lastDataRow; r++) ws['!rows'][r] = {hpt:20};

  // Ширины столбцов по виду
  ws['!cols'] = [
    {wch:6},   // номер
    {wch:12},  // длина
    {wch:12},  // ширина
    {wch:7},   // кол-во
    {wch:16},  // текстура
    {wch:10},  // вверх
    {wch:10},  // вниз
    {wch:10},  // лево
    {wch:10}   // право
  ];

  // Оформление крупного заголовка
  const mainTitleCell = ws["A1"];
  if (mainTitleCell) {
    mainTitleCell.s = {
      font: {bold: true, sz: 15},
      alignment: {horizontal:"center", vertical:"center"}
    };
  }
  const subTitleCell = ws["A2"];
  if (subTitleCell) {
    subTitleCell.s = {
      font: {italic: true, sz: 10},
      alignment: {wrapText: true, horizontal:"center", vertical:"center"}
    };
  }
  const contactCell = ws["A3"];
  if (contactCell) {
    contactCell.s = {
      font: {sz: 11},
      alignment: {horizontal: "left", vertical:"center"}
    };
  }

  // Обработка инфо строки о листе
  const summaryRowIdx = lastDataRow + 2;
  for (let c = 0; c <= header.length; c++) {
    const cellRef = XLSX.utils.encode_cell({r: summaryRowIdx, c});
    if (ws[cellRef]) {
      ws[cellRef].s = {
        border: borderStyle,
        font: {italic: true, color:{rgb:"365881"}},
        alignment: {wrapText:true}
      };
    }
  }

  // Сохраняем файл с именем пользователя (translit for filename, fallback if not found)
  function filenameSafe(s) {
    return (s || "")
      .toLowerCase()
      .replace(/[а-яё]/g, ch =>
        ("abvgdezijklmnoprstufhcy".split('')[
          "абвгдеёжзийклмнопрстуфхцч".indexOf(ch)
        ] || "x")
      )
      .replace(/\s+/g,"_").replace(/[^a-z0-9_]/g,"");
  }
  let fnameOut = (fname && lname)
    ? `zakaz-${filenameSafe(lname)}-${filenameSafe(fname)}`
    : "zakaz";
  XLSX.utils.book_new();
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Шаблон_покраски");
  XLSX.writeFile(wb, `${fnameOut}.xlsx`);
}

// -- Export section --
el('#export-excel-btn').addEventListener('click', e=>{
  // validate user
  state.user.firstname = el('#user-firstname').value.trim();
  state.user.lastname = el('#user-lastname').value.trim();
  state.user.phone = el('#user-phone').value.trim();
  if (!state.user.firstname || !state.user.lastname || !state.user.phone) {
    alert("Пожалуйста, заполните имя, фамилию и телефон!");
    return;
  }
  updatePartsFromTable();
  if (state.parts.length<1) {
    alert("Добавьте хотя бы одну деталь!");
    return;
  }
  syncSheetFromInputs();
  const err = validateParts();
  if (err) { alert(err); return; }
  exportToExcel();
});

// --- Init ---
window.addEventListener('DOMContentLoaded', ()=>{
  state.user.firstname = "";
  state.user.lastname = "";
  state.user.phone = "";
  syncSheetFromInputs();
  // Добавим заготовку детали для удобства
  addPart({length:600, width:400, count:2, texture:"", top:"", bottom:"", left:"", right:""});
  repaintPartsTable();
  repaintEverything();
});