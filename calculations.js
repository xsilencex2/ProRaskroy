// Calculation functions
import { state } from './state.js';

export function getAllUsedArea() {
  return state.parts.reduce((acc,p)=>acc+(p.length*p.width*p.count),0);
}

export function getSheetArea() {
  return state.sheet.width * state.sheet.height;
}

export function wastePercent() {
  const used = getAllUsedArea();
  const total = getSheetArea();
  if (!used || !total) return 0;
  return Math.max(0, 100 - (used/total)*100);
}

