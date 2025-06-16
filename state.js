// State management module
export const MATERIAL_THICKNESSES = {
  "ДСП": 16,
  "МДФ": 18,
  "Фанера": 18,
  "ХДФ": 3
};

export const MATERIAL_SIZES = {
  "ДСП": {
    "2800x2070": [2800, 2070],
    "2500x1830": [2500, 1830],
    "2750x1830": [2750, 1830]
  },
  "МДФ": {
    "2800x2070": [2800, 2070],
    "2800x1220": [2800, 1220]
  },
  "ХДФ": {
    "2800x2070": [2800, 2070],
    "2500x2070": [2500, 2070]
  }
};

export const state = {
  user: {
    firstname: "",
    lastname: "",
    phone: ""
  },
  sheet: {
    width: 2800,
    height: 2070,
    material: "ДСП",
    materialCustomName: "",
    materialType: "ДСП",
    hasTexture: false,
    thickness: 16,
    edge: "",
    edgeThickness: "1 мм"
  },
  parts: []
};

