// Excel export functionality
import * as XLSX from "xlsx";
import { state } from './state.js';
import { el, filenameSafe } from './utils.js';

export function exportToExcel() {
  const fname = (state.user.firstname || '').trim();
  const lname = (state.user.lastname || '').trim();
  const phone = (state.user.phone || '').trim();

  const mainTitle = "Шаблон для покраски";
  const subTitle = "Текстура: X - вдоль, Y - поперёк. Положение кромки указывается для каждой стороны.";
  const contactInfo = [
    `Имя: ${fname}  Фамилия: ${lname}  Телефон: ${phone}`
  ];

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
  
  let matStr = state.sheet.materialCustomName || state.sheet.material;

  const sheetInfo = [
    [
      `Размер листа: ${state.sheet.width} x ${state.sheet.height} мм, материал: ${matStr}, кромка: ${state.sheet.edge} (${state.sheet.edgeThickness}), толщина: ${state.sheet.thickness} мм, текстура: ${state.sheet.hasTexture ? 'да' : 'нет'}`
    ]
  ];

  const ws_data = [];
  ws_data.push([mainTitle]);
  ws_data.push([subTitle]);
  ws_data.push(contactInfo);
  ws_data.push([]);
  ws_data.push(header);
  rows.forEach(row => ws_data.push(row));
  ws_data.push([]);
  ws_data.push(...sheetInfo);

  const ws = XLSX.utils.aoa_to_sheet(ws_data);

  if (!ws['!merges']) ws['!merges'] = [];
  ws["!merges"].push({s: {r: 0, c: 0}, e: {r: 0, c: header.length}});
  ws["!merges"].push({s: {r: 1, c: 0}, e: {r: 1, c: header.length}});
  ws["!merges"].push({s: {r: 2, c: 0}, e: {r: 2, c: header.length}});

  const borderStyle = {
    top: {style: "thin", color: {rgb:"000000"}},
    bottom: {style: "thin", color: {rgb:"000000"}},
    left: {style: "thin", color: {rgb:"000000"}},
    right: {style: "thin", color: {rgb:"000000"}}
  };

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
  ws['!rows'] = [];
  ws['!rows'][0] = {hpt: 22};
  ws['!rows'][1] = {hpt: 18};
  ws['!rows'][2] = {hpt: 15};
  ws['!rows'][firstDataRow] = {hpt: 23};
  for(let r=firstDataRow+1; r<=lastDataRow; r++) ws['!rows'][r] = {hpt:20};

  ws['!cols'] = [
    {wch:6},
    {wch:12},
    {wch:12},
    {wch:7},
    {wch:16},
    {wch:10},
    {wch:10},
    {wch:10},
    {wch:10}
  ];

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

  let fnameOut = (fname && lname)
    ? `zakaz-${filenameSafe(lname)}-${filenameSafe(fname)}`
    : "zakaz";
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Шаблон_покраски");
  XLSX.writeFile(wb, `${fnameOut}.xlsx`);
}