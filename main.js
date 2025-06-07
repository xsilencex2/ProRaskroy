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
function createPartRow(idx, part={}) {
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
    return inp;
  }

  tr.innerHTML = `
    <td>${idx + 1}</td>
    <td></td><td></td><td></td><td></td><td></td><td></td>
    <td>
      <button type="button" class="delete-part-btn" title="Удалить">🗑️</button>
    </td>
  `;
  // Width
  const tdW = tr.children[1];
  tdW.appendChild(input('width', part.width, 'number', { min:5, max:5000 }));
  // Height
  const tdH = tr.children[2];
  tdH.appendChild(input('height', part.height, 'number', { min:5, max:5000 }));
  // Count
  const tdC = tr.children[3];
  tdC.appendChild(input('count', part.count ?? 1, 'number', { min:1, max:500 }));
  // Edge
  const tdEdge = tr.children[4];
  const edgeInp = document.createElement('input');
  edgeInp.type = 'text';
  edgeInp.name = 'edge';
  edgeInp.value = part.edge ?? '';
  edgeInp.placeholder = 'Тип/код кромки';
  edgeInp.required = false;
  edgeInp.maxLength = 50;
  tdEdge.appendChild(edgeInp);
  // Material
  const tdMat = tr.children[5];
  const matSel = document.createElement('select');
  ['ДСП','МДФ','Фанера','ХДФ','Другое'].forEach(e => {
    const opt = document.createElement('option');
    opt.value = e; opt.textContent = e;
    matSel.appendChild(opt);
  });
  matSel.value = part.material ?? '';
  matSel.name = 'material';
  tdMat.appendChild(matSel);
  // Thickness
  const tdThk = tr.children[6];
  const mat = part.material ?? 'ДСП';
  let thVal;
  if (typeof part.thickness === "number" && part.thickness > 0) {
    thVal = part.thickness;
  } else {
    thVal = MATERIAL_THICKNESSES[mat] || 16;
  }
  tdThk.appendChild(input('thickness', thVal, 'number', { min:1, max:40, step:0.1 }));
  return tr;
}

// --- State ---
const state = {
  sheet: {
    width: 2800,
    height: 2070,
    material: 'ДСП',
    thickness: 16
  },
  parts: []
};

// --- Sheet Param Controls ---
const sheetSel = el('#sheet-size-select');
const customSheetDiv = el('#custom-sheet-size');
const sheetWidthInp = el('#sheet-width');
const sheetHeightInp = el('#sheet-height');
const sheetMatSel = el('#sheet-material');
const sheetThicknessLabel = el('#default-thickness-label');
const sheetThicknessValue = el('#sheet-thickness-value');
const sheetThicknessInp = el('#sheet-thickness');
const customThicknessLabel = el('#custom-thickness-label');

function updateSheetThicknessUI(material, thickness) {
  if (material !== "custom" && MATERIAL_THICKNESSES[material]) {
    sheetThicknessLabel.style.display = "";
    sheetThicknessValue.textContent = MATERIAL_THICKNESSES[material];
    customThicknessLabel.style.display = "none";
    sheetThicknessInp.value = MATERIAL_THICKNESSES[material];
  } else {
    sheetThicknessLabel.style.display = "none";
    customThicknessLabel.style.display = "";
    sheetThicknessInp.value = thickness || 16;
  }
}

function syncSheetFromInputs() {
  state.sheet.width = parseInt(sheetWidthInp.value) || 2800;
  state.sheet.height = parseInt(sheetHeightInp.value) || 2070;
  const matVal = sheetMatSel.value;
  state.sheet.material = matVal !== "custom" ? matVal : (sheetMatSel.selectedOptions[0].textContent || "Другое");
  if (matVal !== "custom" && MATERIAL_THICKNESSES[matVal]) {
    state.sheet.thickness = MATERIAL_THICKNESSES[matVal];
  } else {
    let th = parseFloat(sheetThicknessInp.value);
    if (!validateThickness(th)) th = 16;
    state.sheet.thickness = th;
  }
  updateSheetThicknessUI(matVal, state.sheet.thickness);
}

sheetSel.addEventListener('change', (e) => {
  if (sheetSel.value === "custom") {
    customSheetDiv.style.display = "";
    sheetWidthInp.disabled = false; 
    sheetHeightInp.disabled = false;
  } else {
    customSheetDiv.style.display = "none";
    sheetWidthInp.value = 2800;
    sheetHeightInp.value = 2070;
    sheetWidthInp.disabled = true; 
    sheetHeightInp.disabled = true;
    syncSheetFromInputs();
    repaintEverything();
  }
});

[sheetWidthInp, sheetHeightInp, sheetMatSel, sheetThicknessInp].forEach(inp =>
  inp.addEventListener('input', () => {
    syncSheetFromInputs();
    repaintEverything();
  })
);

sheetMatSel.addEventListener('change', () => {
  syncSheetFromInputs();
  repaintEverything();
});

// --- Parts Table Logic ---
const partsTbody = el('#parts-tbody');
function makePartObj(tr) {
  const width = parseInt(tr.querySelector('[name=width]').value);
  const height = parseInt(tr.querySelector('[name=height]').value);
  const count = parseInt(tr.querySelector('[name=count]').value);
  const edge = tr.querySelector('[name=edge]').value; // Исправлено: убрано лишнее пространство
  const material = tr.querySelector('[name=material]').value;
  let thickness = parseFloat(tr.querySelector('[name=thickness]').value);
  if (!validateThickness(thickness)) thickness = MATERIAL_THICKNESSES[material] || 16;
  return { width, height, count, edge, material, thickness };
}

function repaintPartsTable() {
  partsTbody.innerHTML = '';
  state.parts.forEach((part, idx) => {
    const tr = createPartRow(idx, part);
    partsTbody.appendChild(tr);
  });
}

function updatePartsFromTable() {
  const trs = partsTbody.querySelectorAll('tr');
  state.parts = [];
  trs.forEach(tr => {
    const p = makePartObj(tr);
    state.parts.push(p);
  });
}

function validateParts() {
  let ok = true, msg = '';
  state.parts.forEach((p, idx) => {
    if (!validateDimension(p.width) || !validateDimension(p.height)) {
      ok = false;
      msg = `Неверный размер детали №${idx+1}: ширина и высота должны быть 5–5000 мм.`;
    }
    if (!Number.isFinite(p.count) || p.count < 1) {
      ok = false;
      msg = `Неверное количество у детали №${idx+1}.`;
    }
    if (!validateThickness(p.thickness)) {
      ok = false;
      msg = `Укажите корректную толщину (1-40 мм) детали №${idx+1}.`;
    }
  });
  return ok ? null : msg;
}

function addPart(part) {
  if (!part) {
    const material = state.sheet.material;
    const thickness = state.sheet.thickness;
    part = { width: 300, height: 300, count: 1, edge: "", material, thickness };
  }
  if (!validateThickness(part.thickness)) {
    part.thickness = MATERIAL_THICKNESSES[part.material] || 16;
  }
  state.parts.push(part);
  repaintPartsTable();
  repaintEverything();
}

el('#add-part-btn').addEventListener('click', () => addPart());

partsTbody.addEventListener('input', (e) => {
  updatePartsFromTable();
  repaintEverything();
});

partsTbody.addEventListener('change', (e) => {
  updatePartsFromTable();
  repaintEverything();
});

partsTbody.addEventListener('click', (e) => {
  if (e.target.classList.contains('delete-part-btn')) {
    const idx = parseInt(e.target.closest('tr').dataset.idx, 10);
    state.parts.splice(idx, 1);
    repaintPartsTable();
    repaintEverything();
  }
});

// --- Calculation & Visualization ---
function getAllUsedArea() {
  return state.parts.reduce((acc, p) => acc + (p.width * p.height * p.count), 0);
}

function getSheetArea() {
  return state.sheet.width * state.sheet.height;
}

function wastePercent() {
  const used = getAllUsedArea();
  const total = getSheetArea();
  if (!used || !total) return 0;
  return Math.max(0, 100 - (used / total) * 100);
}

function drawPreview() {
  const c = el('#sheet-preview');
  const ctx = c.getContext('2d');
  ctx.clearRect(0, 0, c.width, c.height);

  ctx.strokeStyle = "#365881";
  ctx.lineWidth = 2;
  ctx.strokeRect(0, 0, c.width - 1, c.height - 1);

  const pad = 10;
  const sw = state.sheet.width, sh = state.sheet.height;
  const scale = Math.min((c.width - 2 * pad) / sw, (c.height - 2 * pad) / sh);

  let x = pad, y = pad;
  let rowH = 0;

  let partsList = [];
  state.parts.forEach(p => {
    for (let i = 0; i < p.count; i++) partsList.push({ ...p });
  });
  partsList.sort((a, b) => (b.height * b.width) - (a.height * a.width));

  ctx.font = "10px Segoe UI";
  let colorIdx = 0, colors = ["#33a7ee", "#efbb00", "#a2d541", "#fc6a5e", "#a968c6"];

  for (let i = 0; i < partsList.length; i++) {
    const p = partsList[i];
    let w = p.width * scale, h = p.height * scale;
    if (x + w > c.width - pad) { x = pad; y += rowH + 4; rowH = 0; }
    if (y + h > c.height - pad) break;
    ctx.fillStyle = colors[colorIdx++ % colors.length] + "88";
    ctx.fillRect(x, y, w, h);
    ctx.strokeStyle = "#222";
    ctx.strokeRect(x, y, w, h);
    ctx.fillStyle = "#222";
    ctx.fillText(`${p.width}×${p.height}`, x + 4, y + 14);
    x += w + 4;
    if (h > rowH) rowH = h;
  }

  ctx.save();
  ctx.globalAlpha = 0.65;
  ctx.fillStyle = "#fff";
  ctx.fillRect(0, c.height - 24, 120, 24);
  ctx.fillStyle = "#365881";
  ctx.font = "bold 13px Segoe UI";
  ctx.fillText(`${sw}×${sh} мм`, 5, c.height - 8);
  ctx.restore();
}

function repaintEverything() {
  const mm2toM2 = 1e-6;
  const usedArea = getAllUsedArea() * mm2toM2;
  el('#used-area').textContent = formatNumber(usedArea, 2);

  const waste = wastePercent();
  el('#waste-percent').textContent = formatNumber(waste, 1);

  drawPreview();
}

function exportToExcel() {
  // Формируем данные под шаблон (шаблон.xlsx)
  const mainTitle = "ШАБЛОН ДЛЯ КРАСКРОЯ";
  const subTitle = "текстура 1 - не поворот 0 - поворот кромка 1 - кромка 2 -";

  // Заголовки таблицы
  const header = [
    "Длина",
    "Ширина",
    "Количество",
    "текстура",
    "кр длина",
    "кр длина",
    "кр ширина",
    "кр ширина",
    "наименование детали"
  ];

  // Данные для таблицы
  const rows = state.parts.map((p, idx) => {
    const width = Number.isFinite(p.width) ? p.width : "";
    const height = Number.isFinite(p.height) ? p.height : "";
    const count = Number.isFinite(p.count) ? p.count : "";
    const texture = ""; // Пустое
    const edgeLen = validateDimension(width) && validateDimension(height) && Number.isFinite(count) && count >= 1
      ? (2 * (Number(width) + Number(height)) * Number(count))
      : "";
    const edgeLen2 = edgeLen; // Дублируем
    const edgeWidth = p.edge || ""; // Используем поле edge
    const edgeWidth2 = p.edge || ""; // Дублируем
    const partName = `Деталь №${idx + 1}`;
    return [
      width,
      height,
      count,
      texture,
      edgeLen,
      edgeLen2,
      edgeWidth,
      edgeWidth2,
      partName
    ];
  });

  // Собираем данные для листа
  const ws_data = [];
  ws_data.push([mainTitle]);
  ws_data.push([subTitle]);
  ws_data.push([]); // Пустая строка
  ws_data.push(header);
  rows.forEach(row => ws_data.push(row));

  // Добавляем 30 пустых строк для "квадратиков"
  for (let i = 0; i < 30; i++) {
    ws_data.push(new Array(9).fill("")); // 9 пустых ячеек
  }

  // Создаем лист Excel
  const ws = XLSX.utils.aoa_to_sheet(ws_data);

  // --- Форматирование ---
  const colCnt = 9; // Количество столбцов

  // Объединения ячеек
  if (!ws["!merges"]) ws["!merges"] = [];
  ws["!merges"].push({ s: { r: 0, c: 0 }, e: { r: 0, c: colCnt - 1 } }); // Заголовок
  ws["!merges"].push({ s: { r: 1, c: 0 }, e: { r: 1, c: colCnt - 1 } }); // Подзаголовок

  // Стили границ
  const borderStyle = {
    top: { style: "thin", color: { rgb: "000000" } },
    bottom: { style: "thin", color: { rgb: "000000" } },
    left: { style: "thin", color: { rgb: "000000" } },
    right: { style: "thin", color: { rgb: "000000" } }
  };

  // Применение стилей
  function applyCellStyle(range, styleObj) {
    for (let r = range.s.r; r <= range.e.r; r++) {
      for (let c = range.s.c; c <= range.e.c; c++) {
        const cellRef = XLSX.utils.encode_cell({ r, c });
        if (!ws[cellRef]) {
          ws[cellRef] = { t: "s", v: "" }; // Создаем пустую ячейку
        }
        if (!ws[cellRef].s) ws[cellRef].s = {};
        Object.assign(ws[cellRef].s, styleObj);
      }
    }
  }

  // Стили для таблицы и пустых строк
  const firstDataRow = 3, lastDataRow = 3 + rows.length;
  const totalRows = ws_data.length; // Все строки, включая пустые
  for (let r = firstDataRow; r < totalRows; r++) {
    for (let c = 0; c < colCnt; c++) {
      const cellRef = XLSX.utils.encode_cell({ r, c });
      if (!ws[cellRef]) {
        ws[cellRef] = { t: "s", v: "" }; // Создаем пустую ячейку
      }
      ws[cellRef].s = {
        border: borderStyle,
        alignment: { horizontal: "center", vertical: "center", wrapText: true },
        font: r === firstDataRow ? { name: "Arial", bold: true, sz: 10 } : { name: "Arial", sz: 10 }
      };
    }
  }

  // Стили для заголовка
  const mainTitleCell = ws["A1"];
  if (mainTitleCell) {
    mainTitleCell.s = {
      font: { name: "Arial", bold: true, sz: 14 },
      alignment: { horizontal: "center", vertical: "center" }
    };
  }

  // Стили для подзаголовка
  const subTitleCell = ws["A2"];
  if (subTitleCell) {
    subTitleCell.s = {
      font: { name: "Arial", italic: true, sz: 10 },
      alignment: { wrapText: true, horizontal: "center", vertical: "center" }
    };
  }

  // Установка высоты строк
  ws['!rows'] = [];
  ws['!rows'][0] = { hpt: 30 }; // Заголовок
  ws['!rows'][1] = { hpt: 15 }; // Подзаголовок
  ws['!rows'][2] = { hpt: 12.75 }; // Пустая строка
  ws['!rows'][3] = { hpt: 18 }; // Заголовки таблицы
  for (let r = 4; r < totalRows; r++) ws['!rows'][r] = { hpt: 15 }; // Данные и пустые строки

  // Установка ширины столбцов
  ws['!cols'] = [
    { wch: 12 }, // Длина
    { wch: 12 }, // Ширина
    { wch: 10 }, // Количество
    { wch: 12 }, // текстура
    { wch: 12 }, // кр длина
    { wch: 12 }, // кр длина
    { wch: 12 }, // кр ширина
    { wch: 12 }, // кр ширина
    { wch: 20 }  // наименование детали
  ];

  // Создание книги и экспорт
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Лист1");
  XLSX.writeFile(wb, `шаблон-для-краскроя-${Date.now()}.xlsx`);
}

// -- Export section --
el('#export-excel-btn').addEventListener('click', e => {
  updatePartsFromTable();
  if (state.parts.length < 1) {
    alert("Добавьте хотя бы одну деталь!");
    return;
  }
  syncSheetFromInputs();
  const err = validateParts();
  if (err) { alert(err); return; }
  exportToExcel();
});

// --- Init ---
window.addEventListener('DOMContentLoaded', () => {
  syncSheetFromInputs();
  addPart({ width: 600, height: 400, count: 2, edge: '', material: 'ДСП', thickness: MATERIAL_THICKNESSES['ДСП'] });
  repaintPartsTable();
  repaintEverything();
});