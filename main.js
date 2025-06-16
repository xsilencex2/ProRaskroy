// Main application entry point
import { state } from './state.js';
import { el, formatNumber } from './utils.js';
import { getAllUsedArea, wastePercent } from './calculations.js';
import { repaintPartsTable, updatePartsFromTable, validateParts, addPart } from './parts-manager.js';
import { syncSheetFromInputs, setupSheetUI, handleMaterialTypeChange } from './sheet-manager.js';
import { exportToExcel } from './excel-export.js';

function repaintEverything() {
  // Update summary calculations
  const usedArea = getAllUsedArea() / 1_000_000; // Convert mm² to m²
  const waste = wastePercent();
  
  el('#used-area').textContent = formatNumber(usedArea, 2);
  el('#waste-percent').textContent = formatNumber(waste, 1);
}

function setupPartsDelegation() {
  const partsTable = el('#parts-table');
  
  // Delegation for delete buttons
  partsTable.addEventListener('click', (e) => {
    if (e.target.classList.contains('delete-part-btn')) {
      const row = e.target.closest('tr');
      const idx = parseInt(row.dataset.idx);
      state.parts.splice(idx, 1);
      repaintPartsTable();
      repaintEverything();
    }
  });
  
  // Delegation for input changes
  partsTable.addEventListener('input', (e) => {
    updatePartsFromTable();
    repaintEverything();
  });
  
  partsTable.addEventListener('change', (e) => {
    updatePartsFromTable();
    repaintEverything();
  });
  
  // Add part button
  el('#add-part-btn').addEventListener('click', () => {
    addPart();
    repaintEverything();
  });
}

function setupExportButton() {
  el('#export-excel-btn').addEventListener('click', () => {
    // Validate user data
    const firstname = el('#user-firstname').value.trim();
    const lastname = el('#user-lastname').value.trim();
    const phone = el('#user-phone').value.trim();
    
    if (!firstname || !lastname || !phone) {
      alert('Пожалуйста, заполните все поля с контактными данными.');
      return;
    }
    
    // Update state with user data
    state.user.firstname = firstname;
    state.user.lastname = lastname;
    state.user.phone = phone;
    
    // Validate parts
    updatePartsFromTable();
    const partsError = validateParts();
    if (partsError) {
      alert(partsError);
      return;
    }
    
    if (state.parts.length === 0) {
      alert('Добавьте хотя бы одну деталь для раскроя.');
      return;
    }
    
    try {
      exportToExcel();
      el('#export-output').innerHTML = `<strong>Заказ успешно экспортирован!</strong><br>Файл Excel скачан на ваше устройство.`;
    } catch (error) {
      console.error('Export error:', error);
      el('#export-output').innerHTML = `<strong>Ошибка экспорта:</strong> ${error.message}`;
    }
  });
}

// Init
window.addEventListener('DOMContentLoaded', ()=>{
  state.user.firstname = "";
  state.user.lastname = "";
  state.user.phone = "";

  handleMaterialTypeChange();
  el('#sheet-texture').checked = false;
  state.sheet.hasTexture = false;

  syncSheetFromInputs();
  addPart({length:600, width:400, count:2, texture:"", top:"", bottom:"", left:"", right:""});
  repaintPartsTable();
  repaintEverything();

  setupSheetUI(syncSheetFromInputs, repaintEverything, repaintPartsTable);
  setupPartsDelegation();
  setupExportButton();
});