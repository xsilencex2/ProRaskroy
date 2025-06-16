// Parts table management
import { state } from './state.js';
import { el, validateDimension } from './utils.js';

function createEdgeSelect(currentValue, edgeOptionsValue) {
  const select = document.createElement('select');
  select.name = 'edge-side';
  const val = typeof currentValue === "string" ? currentValue : "";
  if (!edgeOptionsValue) {
    const optNone = document.createElement('option');
    optNone.value = '';
    optNone.textContent = 'Нет';
    select.appendChild(optNone);
    select.value = '';
    return select;
  }
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
  tr.children[1].appendChild(input('length', part.length, 'number', { min:5, max:5000 }));
  tr.children[2].appendChild(input('width', part.width, 'number', { min:5, max:5000 }));
  tr.children[3].appendChild(input('count', part.count ?? 1, 'number', { min:1, max:500 }));

  // TXT (текстура)
  const txtBox = document.createElement('input');
  txtBox.type = 'checkbox';
  txtBox.name = 'texture';
  txtBox.checked = !!part.texture;
  txtBox.style.transform = 'scale(1.2)';
  txtBox.style.cursor = 'pointer';
  if (state.sheet.hasTexture === false) {
    txtBox.disabled = true;
    txtBox.checked = false;
  }
  tr.children[4].appendChild(txtBox);

  // Вверх/Вниз/Лево/Право — селект
  ["top", "bottom", "left", "right"].forEach((side, i) => {
    const sideValue = part[side] || '';
    const select = createEdgeSelect(sideValue, edgeOptionsValue);
    select.dataset.side = side;
    tr.children[5 + i].appendChild(select);
  });

  return tr;
}

function makePartObj(tr) {
  const length = parseInt(tr.querySelector('[name=length]').value);
  const width = parseInt(tr.querySelector('[name=width]').value);
  const count = parseInt(tr.querySelector('[name=count]').value);
  let texture = "";
  if(state.sheet.hasTexture){
    texture = tr.querySelector('[name=texture]').checked ? "да" : "";
  }
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

export function repaintPartsTable() {
  const partsTbody = el('#parts-tbody');
  partsTbody.innerHTML = '';
  const edgeOpt = (state.sheet && state.sheet.edge) ? state.sheet.edge : "";
  state.parts.forEach((part, idx) => {
    const tr = createPartRow(idx, part, edgeOpt);
    partsTbody.appendChild(tr);
  });
}

export function updatePartsFromTable() {
  const partsTbody = el('#parts-tbody');
  const trs = partsTbody.querySelectorAll('tr');
  state.parts = [];
  trs.forEach(tr=>{
    state.parts.push(makePartObj(tr));
  });
}

export function validateParts() {
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

export function addPart(part) {
  if (!part) {
    part = { length: 300, width: 300, count: 1, texture: "", top: "", bottom: "", left: "", right: "" };
  }
  state.parts.push(part);
  repaintPartsTable();
}