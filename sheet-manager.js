// Sheet parameters management
import { state, MATERIAL_THICKNESSES, MATERIAL_SIZES } from './state.js';
import { el, validateThickness } from './utils.js';

const sheetSel = el('#sheet-size-select');
const customSheetDiv = el('#custom-sheet-size');
const sheetWidthInp = el('#sheet-width');
const sheetHeightInp = el('#sheet-height');
const sheetMaterialTypeSel = el('#sheet-material-type');
const sheetMaterialInp = el('#sheet-material');
const materialCustomNameSpan = el('#material-custom-name');
const sheetTextureCheckbox = el('#sheet-texture');
const sheetThicknessInp = el('#sheet-thickness');
const sheetEdgeInp = el('#sheet-edge');
const edgeThicknessSel = el('#edge-thickness');
const sheetThicknessValueLabel = el('#sheet-thickness-value');

function updateMaterialThicknessLabel() {
  let currentMat = sheetMaterialTypeSel.value;
  let t = MATERIAL_THICKNESSES[currentMat];
  if (!t) t = 16;
  sheetThicknessValueLabel.textContent = t;
}

function updateSheetSizeOptions() {
  const materialType = sheetMaterialTypeSel.value;
  const sheetSel = el('#sheet-size-select');
  
  // Clear existing options
  sheetSel.innerHTML = '';
  
  // Get available sizes for current material
  const availableSizes = MATERIAL_SIZES[materialType] || MATERIAL_SIZES["ДСП"];
  
  // Add options for current material
  Object.keys(availableSizes).forEach(sizeKey => {
    const option = document.createElement('option');
    option.value = sizeKey;
    const [width, height] = availableSizes[sizeKey];
    option.textContent = `${width} × ${height} мм`;
    sheetSel.appendChild(option);
  });
  
  // Add custom option
  const customOption = document.createElement('option');
  customOption.value = 'custom';
  customOption.textContent = 'Другое (свой размер)';
  sheetSel.appendChild(customOption);
  
  // Set first option as selected and update dimensions
  sheetSel.selectedIndex = 0;
  const firstSize = Object.values(availableSizes)[0];
  if (firstSize) {
    sheetWidthInp.value = firstSize[0];
    sheetHeightInp.value = firstSize[1];
    state.sheet.width = firstSize[0];
    state.sheet.height = firstSize[1];
  }
  
  // Hide custom size inputs
  customSheetDiv.style.display = "none";
  sheetWidthInp.disabled = true;
  sheetHeightInp.disabled = true;
}

function handleMaterialTypeChange() {
  let currentMat = sheetMaterialTypeSel.value;
  // Always show material name field
  materialCustomNameSpan.style.display = '';
  sheetMaterialInp.required = true;
  
  if (currentMat === 'custom') {
    state.sheet.material = sheetMaterialInp.value.trim() || '';
    state.sheet.materialCustomName = sheetMaterialInp.value.trim() || '';
    state.sheet.materialType = 'custom';
  } else {
    // Set default material name but keep field editable
    if (!sheetMaterialInp.value.trim()) {
      sheetMaterialInp.value = currentMat;
    }
    state.sheet.material = sheetMaterialInp.value.trim() || currentMat;
    state.sheet.materialCustomName = sheetMaterialInp.value.trim() || currentMat;
    state.sheet.materialType = currentMat;
    // Set default thickness for standard materials
    if (MATERIAL_THICKNESSES[currentMat]) {
      sheetThicknessInp.value = MATERIAL_THICKNESSES[currentMat];
    }
  }
  
  // Update available sheet sizes based on material
  updateSheetSizeOptions();
  updateMaterialThicknessLabel();
}

export function syncSheetFromInputs() {
  // Размер
  state.sheet.width = parseInt(sheetWidthInp.value) || 2800;
  state.sheet.height = parseInt(sheetHeightInp.value) || 2070;
  // Тип материала и толщина
  state.sheet.materialType = sheetMaterialTypeSel.value;
  state.sheet.material = sheetMaterialInp.value.trim() || state.sheet.materialType;
  state.sheet.materialCustomName = sheetMaterialInp.value.trim() || state.sheet.materialType;
  
  // Толщина
  let th = parseFloat(sheetThicknessInp.value);
  if (!validateThickness(th)) {
    let dflt = MATERIAL_THICKNESSES[state.sheet.materialType] || 16;
    sheetThicknessInp.value = dflt;
    th = dflt;
  }
  state.sheet.thickness = th;

  // Учет текстуры
  state.sheet.hasTexture = sheetTextureCheckbox.checked;

  state.sheet.edge = sheetEdgeInp.value.trim();
  state.sheet.edgeThickness = edgeThicknessSel.value;

  updateMaterialThicknessLabel();
}

export function setupSheetUI(syncCallback, repaintEverything, repaintPartsTable) {
  sheetMaterialTypeSel.addEventListener('change', () => {
    handleMaterialTypeChange();
    syncCallback();
    repaintEverything();
    repaintPartsTable();
  });
  sheetMaterialInp.addEventListener('input', () => {
    state.sheet.materialCustomName = sheetMaterialInp.value.trim() || '';
    state.sheet.material = sheetMaterialInp.value.trim() || '';
    syncCallback();
    repaintEverything();
    repaintPartsTable();
  });
  sheetTextureCheckbox.addEventListener('change', () => {
    state.sheet.hasTexture = sheetTextureCheckbox.checked;
    repaintPartsTable();
    syncCallback();
    repaintEverything();
  });
  sheetSel.addEventListener('change', () => {
    const materialType = sheetMaterialTypeSel.value;
    const availableSizes = MATERIAL_SIZES[materialType] || MATERIAL_SIZES["ДСП"];
    
    let w = 2800, h = 2070;
    if (sheetSel.value === "custom") {
      customSheetDiv.style.display = "";
      sheetWidthInp.disabled = false; sheetHeightInp.disabled = false;
    } else {
      if (sheetSel.value in availableSizes) {
        w = availableSizes[sheetSel.value][0];
        h = availableSizes[sheetSel.value][1];
      }
      customSheetDiv.style.display = "none";
      sheetWidthInp.value = w;
      sheetHeightInp.value = h;
      sheetWidthInp.disabled = true; sheetHeightInp.disabled = true;
    }
    syncCallback();
    repaintEverything();
    repaintPartsTable();
  });
  [sheetWidthInp, sheetHeightInp, sheetThicknessInp, sheetEdgeInp, edgeThicknessSel].forEach(inp =>
    inp.addEventListener('input', () => {
      syncCallback();
      repaintEverything();
      repaintPartsTable();
    })
  );
  edgeThicknessSel.addEventListener('change', () => {
    syncCallback();
    repaintEverything();
    repaintPartsTable();
  });
}

export { handleMaterialTypeChange };