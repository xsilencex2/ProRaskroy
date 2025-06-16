// Utility functions
export function formatNumber(v, digits=1) {
  return (Math.round(v * Math.pow(10, digits)) / Math.pow(10, digits)).toLocaleString('ru-RU');
}

export function validateDimension(n, min=5, max=5000) {
  return Number.isFinite(n) && n >= min && n <= max;
}

export function validateThickness(t, min=1, max=40) {
  return Number.isFinite(t) && t >= min && t <= max;
}

export function el(sel) { 
  return document.querySelector(sel); 
}

export function filenameSafe(s) {
  const map = {
    а:"a", б:"b", в:"v", г:"g", д:"d", е:"e", ё:"e", ж:"zh", з:"z", и:"i",
    й:"y", к:"k", л:"l", м:"m", н:"n", о:"o", п:"p", р:"r", с:"s", т:"t",
    у:"u", ф:"f", х:"h", ц:"c", ч:"ch", ш:"sh", щ:"sh", ъ:"", ы:"y", ь:"",
    э:"e", ю:"yu", я:"ya"
  };
  return (s || "")
    .toLowerCase()
    .replace(/[а-яё]/g, ch => map[ch] || "x")
    .replace(/\s+/g,"_").replace(/[^a-z0-9_]/g,"");
}

