// ====================================================================
// 📦  1. CONSTANTES & VARIABLES GLOBALES
// ====================================================================

const defaultMatosList = [
  "6x4",
  "8x4",
  "6x2 GRUE",
  "AC3",
  "PB",
  "PC",
  "8x2 AMP GRUE",
];
const defaultChauffeursList = [];
const baseAssocies = ["ALLARD", "LANDAIS", "STD"];
const pilierCommunKey = "Piliers Commun";
const date = document.getElementById("current-date-picker")?.value;
const saved = localStorage.getItem("planning_" + date);
const tableBody = document.getElementById("table-body");
const active = document.activeElement;
const annee = getSelectedYear();
const moisAnnee = getSelectedMonth();

let isMouseDown = false;
let estEnTrainDeColler = false;
let pendingTimeouts = [];

let transportConfigClient =
  JSON.parse(localStorage.getItem("transport_config_client")) || {};
let transportConfigAssocie =
  JSON.parse(localStorage.getItem("transport_config_associe")) || {};
let transportConfigChauffeur =
  JSON.parse(localStorage.getItem("transport_config_chauffeur")) || {};
let transportConfigTelephone =
  JSON.parse(localStorage.getItem("transport_config_telephone")) || {};
let transportConfigGroup =
  JSON.parse(localStorage.getItem("transport_config_group")) || {};
let transportConfigMateriel =
  JSON.parse(localStorage.getItem("transport_config_materiel")) || {};
let transportConfigCategorie =
  JSON.parse(localStorage.getItem("transport_config_categorie")) || {};
let transportConfigGroupCat =
  JSON.parse(localStorage.getItem("transportConfigGroupCat")) || {};
let transportConfigGroupInfo =
  JSON.parse(localStorage.getItem("transportConfigGroupInfo")) || {};

const UNDO_MAX = 50;
let undoStack = JSON.parse(localStorage.getItem("undoStack") || "[]");
let redoStack = JSON.parse(localStorage.getItem("redoStack") || "[]");
let undoDebounceTimer = null;
let isUndoRedoAction = false; // évite de capturer pendant undo/redo
let lastCapturedData = JSON.parse(
  localStorage.getItem("lastCapturedData") || "{}",
);
let isInitialLoading = true;

// ====================================================================
// 🔧  2. UTILITAIRES GÉNÉRAUX
// ====================================================================

function getSelectedYear() {
  const dateValeur = document.getElementById("current-date-picker").value;
  return dateValeur ? dateValeur.substring(0, 4) : "";
}

function getSelectedMonth() {
  const dateValeur = document.getElementById("current-date-picker").value;
  return dateValeur ? dateValeur.substring(0, 7) : "";
}

function rgbToHex(rgb) {
  if (!rgb || rgb === "" || rgb === "transparent") return "";
  if (rgb.startsWith("#")) return rgb;
  const result = rgb.match(/\d+/g);
  if (!result || result.length < 3) return rgb;
  const r = parseInt(result[0]);
  const g = parseInt(result[1]);
  const b = parseInt(result[2]);
  return "#" + ((1 << 24) + (r << 16) + (g << 8) + b).toString(16).slice(1);
}

function formatPhoneNumber(value) {
  const numbers = value.replace(/\D/g, "");
  return numbers.match(/.{1,2}/g)?.join(" ") || numbers;
}

function calculateIndice(k, cm1, c0) {
  if (!k || !cm1 || !c0) return 0;
  const result = (1 - k + (k * cm1) / c0 - 1) * 100;
  return Math.round(result * 100) / 100;
}

function afficherNotification(message, couleur = "#ff9800") {
  const notif = document.getElementById("notification-doublon");
  if (!notif) return;
  notif.textContent = message;
  notif.style.backgroundColor = couleur;
  notif.style.display = "block";
  setTimeout(() => {
    notif.style.display = "none";
  }, 3000);
}

function toggleHelp() {
  const modal = document.getElementById("help-modal");
  const overlay = document.getElementById("help-overlay");
  const isVisible = modal.style.display === "block";
  modal.style.display = isVisible ? "none" : "block";
  overlay.style.display = isVisible ? "none" : "block";
}

function toggleMenu(id, event) {
  if (event) event.stopPropagation();
  const menus = [
    "sous-menu-add",
    "sous-menu-copy",
    "sous-menu-photos",
    "sous-menu-config",
  ];
  menus.forEach((menuId) => {
    const el = document.getElementById(menuId);
    if (menuId === id) {
      el.style.display =
        el.style.display === "none" || el.style.display === ""
          ? "block"
          : "none";
    } else {
      el.style.display = "none";
    }
  });
}

function closeConfigs(id) {
  document.getElementById(id).style.display = "none";
  document.body.style.overflow = "";
}

function trierDatalist(id) {
  const list = document.getElementById(id);
  if (!list) return;
  const options = Array.from(list.options);
  options.sort((a, b) => {
    const valA = a.value.toUpperCase();
    const valB = b.value.toUpperCase();
    if (id === "list-associe") {
      const isABase = baseAssocies.includes(valA);
      const isBBase = baseAssocies.includes(valB);
      if (isABase && !isBBase) return -1;
      if (!isABase && isBBase) return 1;
    }
    return valA.localeCompare(valB);
  });
  list.innerHTML = "";
  options.forEach((opt) => list.appendChild(opt));
}

function filterTable() {
  const associeQuery = document
    .getElementById("filter-associe")
    .value.toUpperCase();
  const clientQuery = document
    .getElementById("filter-client")
    .value.toUpperCase();
  const chantierQuery = document
    .getElementById("filter-chantier")
    .value.toUpperCase();
  const materielQuery = document
    .getElementById("filter-materiel")
    .value.toUpperCase();
  const rows = document.querySelectorAll("#table-body tr");
  rows.forEach((tr) => {
    const associeVal = tr.querySelector(".select-associe")
      ? tr.querySelector(".select-associe").value.toUpperCase()
      : "";
    const clientVal = tr.querySelector(".select-client")
      ? tr.querySelector(".select-client").value.toUpperCase()
      : "";
    const materielVal = tr.querySelector(".select-vehicule")
      ? tr.querySelector(".select-vehicule").value.toUpperCase()
      : "";
    const chantierCell = tr.cells[3];
    let chantierVal = chantierCell
      ? chantierCell.textContent.toUpperCase()
      : "";
    if (chantierVal === "CHANTIER...") chantierVal = "";
    if (
      associeVal.includes(associeQuery) &&
      clientVal.includes(clientQuery) &&
      chantierVal.includes(chantierQuery) &&
      materielVal.includes(materielQuery)
    ) {
      tr.style.display = "";
    } else {
      tr.style.display = "none";
    }
  });
}

function resetFilters() {
  document.getElementById("filter-associe").value = "";
  document.getElementById("filter-client").value = "";
  document.getElementById("filter-chantier").value = "";
  document.getElementById("filter-materiel").value = "";
  filterTable();
}

function trierPlanningParAssocie() {
  const rows = Array.from(tableBody.querySelectorAll("tr"));
  const rowsToSort = rows.filter(
    (tr) => !tr.classList.contains("group-separator"),
  );
  rowsToSort.sort((a, b) => {
    const valA = a.querySelector(".select-associe")?.value.toUpperCase() || "";
    const valB = b.querySelector(".select-associe")?.value.toUpperCase() || "";
    if (valA === "" && valB === "") return 0;
    if (valA === "") return 1;
    if (valB === "") return -1;
    const isABase = baseAssocies.includes(valA);
    const isBBase = baseAssocies.includes(valB);
    if (isABase && !isBBase) return -1;
    if (!isABase && isBBase) return 1;
    const compareAssocie = valA.localeCompare(valB);
    if (compareAssocie !== 0) return compareAssocie;
    const matosA =
      a.querySelector(".select-vehicule")?.value.toUpperCase() || "";
    const matosB =
      b.querySelector(".select-vehicule")?.value.toUpperCase() || "";
    if (matosA === "" && matosB === "") return 0;
    if (matosA === "") return 1;
    if (matosB === "") return -1;
    return matosA.localeCompare(matosB);
  });
  rowsToSort.forEach((tr) => tableBody.appendChild(tr));
}

function updateAllAfterDateChange() {
  renderConfigUI();
  if (typeof renderConfigUI2 === "function") renderConfigUI2();
  if (typeof refreshAllPlanningTarifs === "function")
    refreshAllPlanningTarifs();
  else if (typeof renderPlanning === "function") renderPlanning();
}

// ====================================================================
// ↩️  3. UNDO / REDO
// ====================================================================

function sauvegarderUndoRedoEnLocal() {
  try {
    localStorage.setItem("undoStack", JSON.stringify(undoStack));
    localStorage.setItem("redoStack", JSON.stringify(redoStack));
    localStorage.setItem("lastCapturedData", JSON.stringify(lastCapturedData));
  } catch (e) {} // si localStorage plein, on ignore silencieusement
}

function captureUndoStateBefore() {
  if (isUndoRedoAction || estEnTrainDeColler || isInitialLoading) return;
  if (!date) return;
  if (!saved) return;
  const hadRedo = redoStack.some((s) => s.date === date);
  redoStack = redoStack.filter((s) => s.date !== date);
  if (hadRedo) updateUndoRedoButtons();
  if (lastCapturedData[date] === saved) return;
  undoStack.push({ date, data: saved });
  if (undoStack.length > UNDO_MAX) undoStack.shift();
  lastCapturedData[date] = saved;
  updateUndoRedoButtons();
}

function undoPlanning() {
  if (!date || undoStack.length === 0) return;
  let idx = undoStack.length - 1;
  while (idx >= 0 && undoStack[idx].date !== date) idx--;
  if (idx < 0) return;
  if (saved) {
    redoStack.push({ date, data: saved });
  }
  const stateToRestore = undoStack[idx];
  undoStack.splice(idx, 1);
  isUndoRedoAction = true;
  localStorage.setItem("planning_" + date, stateToRestore.data);
  lastCapturedData[date] = stateToRestore.data;
  rechargerPlanningDepuisLocalStorage(date, () => {
    isUndoRedoAction = false;
    updateUndoRedoButtons();
  });
  afficherNotification("↩ Annulation", "#546e7a");
}

function redoPlanning() {
  if (!date || redoStack.length === 0) return;
  let idx = redoStack.length - 1;
  while (idx >= 0 && redoStack[idx].date !== date) idx--;
  if (idx < 0) return;
  if (saved) {
    undoStack.push({ date, data: saved });
    lastCapturedData[date] = saved;
  }
  const stateToRestore = redoStack[idx];
  redoStack.splice(idx, 1);
  isUndoRedoAction = true;
  localStorage.setItem("planning_" + date, stateToRestore.data);
  lastCapturedData[date] = stateToRestore.data;
  rechargerPlanningDepuisLocalStorage(date, () => {
    isUndoRedoAction = false;
    updateUndoRedoButtons();
  });
  afficherNotification("↪ Rétablissement", "#546e7a");
}

function rechargerPlanningDepuisLocalStorage(date, onDone) {
  pendingTimeouts.forEach((t) => clearTimeout(t));
  pendingTimeouts = [];
  tableBody.innerHTML = "";
  if (saved) {
    try {
      const rows = JSON.parse(saved);
      if (rows.length > 0) {
        rows.forEach(ajouterLigne);
        setTimeout(() => {
          forceRecalculTousLesTarifs();
          lastCapturedData[date] = localStorage.getItem("planning_" + date);
          if (onDone) onDone();
        }, 150);
        return;
      }
    } catch (e) {}
  }
  ajouterLigne();
  lastCapturedData[date] = localStorage.getItem("planning_" + date);
  if (onDone) onDone();
}

function updateUndoRedoButtons() {
  const btnUndo = document.getElementById("btn-undo");
  const btnRedo = document.getElementById("btn-redo");
  if (!btnUndo || !btnRedo) return;
  const hasUndo = date && undoStack.some((s) => s.date === date);
  const hasRedo = date && redoStack.some((s) => s.date === date);
  btnUndo.disabled = !hasUndo;
  btnUndo.style.opacity = hasUndo ? "1" : "0.35";
  btnRedo.disabled = !hasRedo;
  btnRedo.style.opacity = hasRedo ? "1" : "0.35";
}

document.addEventListener("keydown", function (e) {
  if ((e.ctrlKey || e.metaKey) && e.key === "z" && !e.shiftKey) {
    if (
      active &&
      (active.isContentEditable ||
        active.tagName === "INPUT" ||
        active.tagName === "TEXTAREA")
    )
      return;
    e.preventDefault();
    undoPlanning();
  }
  if (
    (e.ctrlKey || e.metaKey) &&
    (e.key === "y" || (e.key === "z" && e.shiftKey))
  ) {
    if (
      active &&
      (active.isContentEditable ||
        active.tagName === "INPUT" ||
        active.tagName === "TEXTAREA")
    )
      return;
    e.preventDefault();
    redoPlanning();
  }
});

// ====================================================================
// 💾  4. SAUVEGARDE & LECTURE DU PLANNING
// ====================================================================

function saveCurrentDay() {
  if (estEnTrainDeColler) return;
  if (!date) return;
  const rows = [];
  document.querySelectorAll("#table-body tr").forEach((tr) => {
    const dateLigne = tr.cells[2]?.innerText || "";
    const estMemeMois = estMemeMoisQueSelection(dateLigne);
    let snapshot = null;
    if (tr.dataset.snapshot) snapshot = JSON.parse(tr.dataset.snapshot);
    if (estMemeMois) {
      const client = tr.querySelector(".select-client").value;
      const vehicule = tr.querySelector(".select-vehicule").value;
      const associe = tr.querySelector(".select-associe").value;
      if ((client && vehicule) || (associe && vehicule)) {
        snapshot = captureSnapshot(client, vehicule, associe);
        tr.dataset.snapshot = JSON.stringify(snapshot);
      }
    }
    const cellules = tr.querySelectorAll("td");
    const couleursParCellule = {};
    cellules.forEach((td, index) => {
      if (index === 0) return;
      const bgColor = td.style.backgroundColor;
      if (bgColor && bgColor !== "" && bgColor !== "transparent") {
        const hexColor = rgbToHex(bgColor);
        if (hexColor) couleursParCellule[index] = hexColor;
      }
    });
    rows.push({
      associe: tr.querySelector(".select-associe").value,
      chantier: tr.cells[1].innerText,
      date: tr.cells[2].innerText,
      vehicule: tr.querySelector(".select-vehicule").value,
      client: tr.querySelector(".select-client").value,
      telephone:
        tr
          .querySelector(".col-telephone")
          ?.getAttribute("data-telephone-full") ||
        tr.querySelector(".select-telephone")?.value ||
        "",
      consigne: tr.cells[6].innerText,
      chauffeur: tr.querySelector(".select-chauffeur").value,
      n_chantier: tr.cells[8].innerText,
      n_commande: tr.cells[9].innerText,
      n_bon: tr.cells[10].innerText,
      tarif: tr.cells[11].innerText,
      supplement_client: tr.cells[12].innerText,
      IC_client: tr.cells[13].innerText,
      tarif_affrete: tr.cells[14].innerText,
      supplement_affrete: tr.cells[15].innerText,
      IC_affrete: tr.cells[16].innerText,
      commentaire: tr.cells[17].innerText,
      snapshot: snapshot,
      couleurs_cellules:
        Object.keys(couleursParCellule).length > 0 ? couleursParCellule : null,
    });
  });
  localStorage.setItem("planning_" + date, JSON.stringify(rows));
  calculerTousLesTotaux();
}

function chargerCouleursDate(date) {}

function estMemeMoisQueSelection(dateStr) {
  if (!dateStr) return false;
  if (!date) return false;
  const parties = dateStr.split("/");
  if (parties.length !== 3) return false;
  const moisLigne = `20${parties[2]}-${parties[1]}`;
  const moisSelectionne = date.substring(0, 7);
  return moisLigne === moisSelectionne;
}

// ====================================================================
// 💰  5. TARIFICATION — LECTURE DES CONFIGS
// ====================================================================

function getCatValues(groupOrAssocie, cat) {
  const base = transportConfigGroupInfo[groupOrAssocie]?.[cat] || {};
  const k = parseFloat(base["k_" + annee]) || 0;
  const c0 = parseFloat(base["c0_" + annee]) || 0;
  const usesCM =
    transportConfigGroupInfo[groupOrAssocie]?.["_useCM_" + cat] === true;
  let cm1;
  if (usesCM) {
    cm1 = parseFloat(base["cm1_" + moisAnnee]) || 0;
  } else {
    const cm1Global =
      transportConfigGroupInfo["GLOBAL_SETTINGS"]?.["ALL"]?.[
        "cm1_" + moisAnnee
      ];
    cm1 =
      cm1Global !== undefined
        ? parseFloat(cm1Global)
        : parseFloat(base["cm1_" + moisAnnee]) || 0;
  }
  return { k, cm1, c0, usesCM };
}

function setCatValue(groupOrAssocie, cat, field, value) {
  if (!transportConfigGroupInfo[groupOrAssocie])
    transportConfigGroupInfo[groupOrAssocie] = {};
  if (!transportConfigGroupInfo[groupOrAssocie][cat])
    transportConfigGroupInfo[groupOrAssocie][cat] = {};
  const usesCM =
    transportConfigGroupInfo[groupOrAssocie]?.["_useCM_" + cat] === true;
  if (field === "cm1") {
    if (usesCM) {
      transportConfigGroupInfo[groupOrAssocie][cat]["cm1_" + moisAnnee] =
        parseFloat(value) || 0;
    } else {
      if (!transportConfigGroupInfo["GLOBAL_SETTINGS"])
        transportConfigGroupInfo["GLOBAL_SETTINGS"] = {};
      if (!transportConfigGroupInfo["GLOBAL_SETTINGS"]["ALL"])
        transportConfigGroupInfo["GLOBAL_SETTINGS"]["ALL"] = {};
      transportConfigGroupInfo["GLOBAL_SETTINGS"]["ALL"]["cm1_" + moisAnnee] =
        parseFloat(value) || 0;
    }
  } else if (field === "useCM") {
    transportConfigGroupInfo[groupOrAssocie]["_useCM_" + cat] = value === true;
  } else {
    let storageKey = field === "k" ? "k_" + annee : "c0_" + annee;
    transportConfigGroupInfo[groupOrAssocie][cat][storageKey] =
      parseFloat(value) || 0;
  }
  localStorage.setItem(
    "transportConfigGroupInfo",
    JSON.stringify(transportConfigGroupInfo),
  );
}

function getTarifClient(clientNom, matosNom) {
  const groupeNom = transportConfigClient[clientNom];
  if (!groupeNom) return 0;
  const groupeData = transportConfigGroup[groupeNom];
  if (groupeData && groupeData[matosNom]) {
    const tarifData = groupeData[matosNom].tarif;
    if (typeof tarifData === "object" && tarifData !== null) {
      return parseFloat(tarifData[annee]) || 0;
    }
    return parseFloat(tarifData) || 0;
  }
  return 0;
}

function getICClient(clientNom, matosNom) {
  const groupeNom = transportConfigClient[clientNom];
  if (!groupeNom) return 0;
  const groupeData = transportConfigGroup[groupeNom];
  if (groupeData && groupeData[matosNom]) {
    const tarif = getTarifClient(clientNom, matosNom);
    let indice = 0;
    const indiceData = groupeData[matosNom].indice;
    if (typeof indiceData === "object")
      indice = parseFloat(indiceData[moisAnnee]) || 0;
    else indice = parseFloat(indiceData) || 0;
    if (indice === 0) {
      const catDuMatos = transportConfigMateriel[matosNom];
      if (catDuMatos && transportConfigGroupInfo[groupeNom]) {
        const { k, cm1, c0 } = getCatValues(groupeNom, catDuMatos);
        indice = calculateIndice(k, cm1, c0);
      }
    }
    return tarif * (indice / 100);
  }
  return 0;
}

function getTarifAffrete(associeNom, matosNom) {
  if (!associeNom || !matosNom || !transportConfigAssocie[associeNom]) return 0;
  const data = transportConfigAssocie[associeNom][matosNom];
  if (!data) return 0;
  if (typeof data.tarif === "object" && data.tarif !== null) {
    return parseFloat(data.tarif[annee]) || 0;
  }
  return parseFloat(data.tarif) || 0;
}

function getICAffrete(associeNom, matosNom) {
  if (!associeNom || !matosNom || !transportConfigAssocie[associeNom]) return 0;
  const data = transportConfigAssocie[associeNom][matosNom];
  if (!data) return 0;
  const tarif = getTarifAffrete(associeNom, matosNom);
  let indice = 0;
  const indiceData = data.indice;
  if (typeof indiceData === "object")
    indice = parseFloat(indiceData[moisAnnee]) || 0;
  else indice = parseFloat(indiceData) || 0;
  if (indice === 0) {
    const catDuMatos = transportConfigMateriel[matosNom];
    if (catDuMatos && transportConfigGroupInfo[associeNom]) {
      const { k, cm1, c0 } = getCatValues(associeNom, catDuMatos);
      indice = calculateIndice(k, cm1, c0);
    }
  }
  return tarif * (indice / 100);
}

// ====================================================================
// 🔗  6. PILIERS COMMUNS (ALLARD / LANDAIS / STD)
// ====================================================================

function getPilierCommunTarif(matos) {
  const data = transportConfigAssocie[pilierCommunKey]?.[matos];
  if (!data) return 0;
  if (typeof data.tarif === "object" && data.tarif !== null)
    return parseFloat(data.tarif[annee]) || 0;
  return parseFloat(data.tarif) || 0;
}

function getPilierCommunIndice(matos) {
  const data = transportConfigAssocie[pilierCommunKey]?.[matos];
  if (!data) return 0;
  if (typeof data.indice === "object")
    return parseFloat(data.indice[moisAnnee]) || 0;
  return parseFloat(data.indice) || 0;
}

function isPilierOverridden(pilierName, matos) {
  if (!baseAssocies.includes(pilierName.toUpperCase())) return false;
  return transportConfigAssocie[pilierName]?.[matos]?.["_override"] === true;
}

function setPilierOverride(pilierName, matos, isOverride) {
  if (!transportConfigAssocie[pilierName])
    transportConfigAssocie[pilierName] = {};
  if (!transportConfigAssocie[pilierName][matos])
    transportConfigAssocie[pilierName][matos] = { tarif: 0, indice: 0 };
  transportConfigAssocie[pilierName][matos]["_override"] = isOverride;
  localStorage.setItem(
    "transport_config_associe",
    JSON.stringify(transportConfigAssocie),
  );
}

function syncPiliersFromCommon(matos, forceAll = false) {
  const communData = transportConfigAssocie[pilierCommunKey]?.[matos];
  if (!communData) return;
  const communTarif =
    typeof communData.tarif === "object"
      ? communData.tarif[annee] || 0
      : communData.tarif || 0;
  const communIndice =
    typeof communData.indice === "object"
      ? communData.indice[moisAnnee] || 0
      : communData.indice || 0;

  baseAssocies.forEach((pilier) => {
    if (!forceAll && isPilierOverridden(pilier, matos)) return; // skip if custom override
    if (!transportConfigAssocie[pilier]) transportConfigAssocie[pilier] = {};
    if (!transportConfigAssocie[pilier][matos])
      transportConfigAssocie[pilier][matos] = { tarif: {}, indice: {} };

    if (
      typeof transportConfigAssocie[pilier][matos].tarif !== "object" ||
      transportConfigAssocie[pilier][matos].tarif === null
    ) {
      transportConfigAssocie[pilier][matos].tarif = {};
    }
    transportConfigAssocie[pilier][matos].tarif[annee] = communTarif;

    if (
      typeof transportConfigAssocie[pilier][matos].indice !== "object" ||
      transportConfigAssocie[pilier][matos].indice === null
    ) {
      transportConfigAssocie[pilier][matos].indice = {};
    }
    transportConfigAssocie[pilier][matos].indice[moisAnnee] = communIndice;
    transportConfigAssocie[pilier][matos]["_override"] = false;
  });
  localStorage.setItem(
    "transport_config_associe",
    JSON.stringify(transportConfigAssocie),
  );
}

function syncCatFromCommon(cat, forceAll = false) {
  const communCatInfo = transportConfigGroupInfo[pilierCommunKey]?.[cat];
  if (!communCatInfo) return;

  baseAssocies.forEach((pilier) => {
    const overrideKey = "_catOverride_" + cat;
    if (!forceAll && transportConfigGroupInfo[pilier]?.[overrideKey] === true)
      return;
    if (!transportConfigGroupInfo[pilier])
      transportConfigGroupInfo[pilier] = {};
    if (!transportConfigGroupInfo[pilier][cat])
      transportConfigGroupInfo[pilier][cat] = {};
    Object.assign(transportConfigGroupInfo[pilier][cat], communCatInfo);
    const useCMKey = "_useCM_" + cat;
    transportConfigGroupInfo[pilier][useCMKey] =
      transportConfigGroupInfo[pilierCommunKey]?.[useCMKey] || false;
    transportConfigGroupInfo[pilier][overrideKey] = false;
  });
  localStorage.setItem(
    "transportConfigGroupInfo",
    JSON.stringify(transportConfigGroupInfo),
  );
}

function updatePilierCommunTarif(matos, value) {
  if (!transportConfigAssocie[pilierCommunKey])
    transportConfigAssocie[pilierCommunKey] = {};
  if (!transportConfigAssocie[pilierCommunKey][matos])
    transportConfigAssocie[pilierCommunKey][matos] = {
      tarif: {},
      indice: {},
    };
  if (typeof transportConfigAssocie[pilierCommunKey][matos].tarif !== "object")
    transportConfigAssocie[pilierCommunKey][matos].tarif = {};
  transportConfigAssocie[pilierCommunKey][matos].tarif[annee] =
    parseFloat(value) || 0;
  localStorage.setItem(
    "transport_config_associe",
    JSON.stringify(transportConfigAssocie),
  );
  syncPiliersFromCommon(matos);
  refreshAllPlanningTarifs();
  renderConfigUI2();
}

function updatePilierCommunIndice(matos, value) {
  if (!transportConfigAssocie[pilierCommunKey])
    transportConfigAssocie[pilierCommunKey] = {};
  if (!transportConfigAssocie[pilierCommunKey][matos])
    transportConfigAssocie[pilierCommunKey][matos] = {
      tarif: {},
      indice: {},
    };
  if (typeof transportConfigAssocie[pilierCommunKey][matos].indice !== "object")
    transportConfigAssocie[pilierCommunKey][matos].indice = {};
  transportConfigAssocie[pilierCommunKey][matos].indice[moisAnnee] =
    parseFloat(value) || 0;
  localStorage.setItem(
    "transport_config_associe",
    JSON.stringify(transportConfigAssocie),
  );
  syncPiliersFromCommon(matos);
  refreshAllPlanningTarifs();
  renderConfigUI2();
}

function resetPilierToCommon(pilierName, matos) {
  setPilierOverride(pilierName, matos, false);
  syncPiliersFromCommon(matos, false);
  refreshAllPlanningTarifs();
  renderConfigUI2();
  afficherNotification(
    `${pilierName} / ${matos} réinitialisé aux valeurs communes`,
    "#007bff",
  );
}

function updatePilierIndividuelTarif(pilierName, matos, value) {
  if (!transportConfigAssocie[pilierName])
    transportConfigAssocie[pilierName] = {};
  if (!transportConfigAssocie[pilierName][matos])
    transportConfigAssocie[pilierName][matos] = { tarif: {}, indice: {} };
  if (
    typeof transportConfigAssocie[pilierName][matos].tarif !== "object" ||
    transportConfigAssocie[pilierName][matos].tarif === null
  ) {
    transportConfigAssocie[pilierName][matos].tarif = {};
  }
  transportConfigAssocie[pilierName][matos].tarif[annee] =
    parseFloat(value) || 0;
  transportConfigAssocie[pilierName][matos]["_override"] = true;
  localStorage.setItem(
    "transport_config_associe",
    JSON.stringify(transportConfigAssocie),
  );
  refreshAllPlanningTarifs();
}

function updatePilierIndividuelIndice(pilierName, matos, value) {
  if (!transportConfigAssocie[pilierName])
    transportConfigAssocie[pilierName] = {};
  if (!transportConfigAssocie[pilierName][matos])
    transportConfigAssocie[pilierName][matos] = { tarif: {}, indice: {} };
  if (typeof transportConfigAssocie[pilierName][matos].indice !== "object")
    transportConfigAssocie[pilierName][matos].indice = {};
  transportConfigAssocie[pilierName][matos].indice[moisAnnee] =
    parseFloat(value) || 0;
  transportConfigAssocie[pilierName][matos]["_override"] = true;
  localStorage.setItem(
    "transport_config_associe",
    JSON.stringify(transportConfigAssocie),
  );
  refreshAllPlanningTarifs();
}

function updatePilierCommunCatValue(cat, field, value) {
  setCatValue(pilierCommunKey, cat, field, value);
  syncCatFromCommon(cat);
  renderConfigUI2();
  refreshAllPlanningTarifs();
}

function togglePilierCommunUseCM(cat, checked) {
  setCatValue(pilierCommunKey, cat, "useCM", checked);
  syncCatFromCommon(cat);
  renderConfigUI2();
  refreshAllPlanningTarifs();
}

function ensurePiliersCommunsExists() {
  if (!transportConfigAssocie[pilierCommunKey]) {
    transportConfigAssocie[pilierCommunKey] = {};
  }
  baseAssocies.forEach((pilier) => {
    if (transportConfigAssocie[pilier]) {
      Object.keys(transportConfigAssocie[pilier]).forEach((matos) => {
        if (!transportConfigAssocie[pilierCommunKey][matos]) {
          transportConfigAssocie[pilierCommunKey][matos] = {
            tarif: {},
            indice: {},
          };
        }
      });
    }
  });
  defaultMatosList.forEach((m) => {
    if (!transportConfigAssocie[pilierCommunKey][m]) {
      transportConfigAssocie[pilierCommunKey][m] = {
        tarif: {},
        indice: {},
      };
    }
  });
  localStorage.setItem(
    "transport_config_associe",
    JSON.stringify(transportConfigAssocie),
  );
}

// ====================================================================
// 🔄  7. TARIFICATION — MISE À JOUR EN TEMPS RÉEL
// ====================================================================

function updateRowTarif(tr) {
  const dateLigne = tr.cells[2]?.innerText || "";
  const partiesDate = dateLigne.split("/");
  const anneeLigne = partiesDate.length === 3 ? "20" + partiesDate[2] : "";
  if (anneeLigne && annee && anneeLigne !== annee) {
    return;
  }

  const eventTrigger = window.event ? window.event.target : null;
  const isComboChange =
    eventTrigger && eventTrigger.classList.contains("combo-input");
  const client = tr.querySelector(".select-client").value;
  const associe = tr.querySelector(".select-associe").value;
  const matos = tr.querySelector(".select-vehicule").value;
  const cellTarifClient = tr.querySelector(".cell-tarif");
  const cellTarifAffrete = tr.querySelector(".cell-tarif-affrete");
  const cellICClient = tr.querySelector(".cell-IC-client");
  const cellICAffrete = tr.querySelector(".cell-IC-affrete");

  if (client && matos) {
    let tarifC =
      parseFloat(
        cellTarifClient.textContent.replace(/[€\s]/g, "").replace(",", "."),
      ) || 0;
    if (isComboChange || tarifC === 0) tarifC = getTarifClient(client, matos);
    if (document.activeElement !== cellTarifClient)
      cellTarifClient.textContent = tarifC.toFixed(2) + " €";
    const icC = getICClient(client, matos);
    cellICClient.textContent = icC.toFixed(2) + " €";
    cellICClient.style.color =
      icC > 0 ? "#28a745" : icC < 0 ? "#dc3545" : "#95a5a6";
    cellICClient.style.fontWeight = icC !== 0 ? "bold" : "normal";
  } else {
    if (document.activeElement !== cellTarifClient)
      cellTarifClient.textContent = "0.00 €";
    cellICClient.textContent = "0.00 €";
  }

  if (associe && matos && client) {
    let tarifA =
      parseFloat(
        cellTarifAffrete.textContent.replace(/[€\s]/g, "").replace(",", "."),
      ) || 0;
    if (isComboChange || tarifA === 0) {
      tarifA = getTarifAffrete(associe, matos);
    }
    if (document.activeElement !== cellTarifAffrete)
      cellTarifAffrete.textContent = tarifA.toFixed(2) + " €";
    const icA = getICAffrete(associe, matos);
    cellICAffrete.textContent = icA.toFixed(2) + " €";
    cellICAffrete.style.color =
      icA > 0 ? "#28a745" : icA < 0 ? "#dc3545" : "#95a5a6";
    cellICAffrete.style.fontWeight = icA !== 0 ? "bold" : "normal";
  } else {
    if (document.activeElement !== cellTarifAffrete)
      cellTarifAffrete.textContent = "0.00 €";
    cellICAffrete.textContent = "0.00 €";
  }

  saveCurrentDay();
  if ((client && matos) || (associe && matos)) {
    const snapshot = captureSnapshot(client, matos, associe);
    tr.dataset.snapshot = JSON.stringify(snapshot);
  }
  calculerTousLesTotaux();
}

function updateRowHighlight(tr) {
  const bonCell = tr.querySelector(".cell-bon");
  if (!bonCell) return;
  const texte = bonCell.textContent.trim();
  const hasBon = texte !== "" && texte !== "N° Bon...";
  if (!hasBon) return;
  const cells = tr.querySelectorAll("td");
  cells.forEach((td, index) => {
    if (index === 0) return; // jamais toucher la couleur associé
    const currentBg = td.style.backgroundColor;
    if (currentBg && currentBg !== "" && currentBg !== "transparent") return;
    td.style.backgroundColor = "#d4edda";
  });
}

function refreshAllPlanningTarifs() {
  const rows = document.querySelectorAll("#table-body tr");
  rows.forEach((tr) => {
    const dateLigne = tr.cells[2]?.innerText || ""; // format DD/MM/YY
    const partiesDate = dateLigne.split("/");
    const anneeLigne = partiesDate.length === 3 ? "20" + partiesDate[2] : "";

    if (anneeLigne && anneeLigne !== annee) return;

    const client = tr.querySelector(".select-client")?.value;
    const matos = tr.querySelector(".select-vehicule")?.value;
    const associe = tr.querySelector(".select-associe")?.value;

    if (client && matos) {
      const tarifC = getTarifClient(client, matos);
      const icC = getICClient(client, matos);
      tr.querySelector(".cell-tarif").textContent = tarifC.toFixed(2) + " €";
      tr.querySelector(".cell-IC-client").textContent = icC.toFixed(2) + " €";
    }

    if (associe && matos) {
      const tarifA = getTarifAffrete(associe, matos);
      const icA = getICAffrete(associe, matos);
      tr.querySelector(".cell-tarif-affrete").textContent =
        tarifA.toFixed(2) + " €";
      tr.querySelector(".cell-IC-affrete").textContent = icA.toFixed(2) + " €";
    }

    if ((client && matos) || (associe && matos)) {
      const snapshot = captureSnapshot(client, matos, associe);
      tr.dataset.snapshot = JSON.stringify(snapshot);
    }
  });

  const keysToUpdate = [];
  for (let i = 0; i < localStorage.length; i++) {
    const key = localStorage.key(i);
    if (key && key.startsWith("planning_") && key.substring(9, 13) === annee) {
      keysToUpdate.push(key);
    }
  }

  keysToUpdate.forEach((key) => {
    try {
      const dayData = JSON.parse(localStorage.getItem(key));
      if (!Array.isArray(dayData)) return;

      let modifie = false;

      dayData.forEach((row) => {
        const client =
          row.client && row.client !== "Client..." ? row.client : "";
        const matos =
          row.vehicule && row.vehicule !== "Véhicule..." ? row.vehicule : "";
        const associe =
          row.associe && row.associe !== "Associé..." ? row.associe : "";

        if (!matos) return;

        if (client && matos) {
          const nouveauTarifC = getTarifClient(client, matos);
          const nouveauICClient = getICClient(client, matos);
          const ancienTarifC =
            parseFloat(
              (row.tarif || "0").replace(/[€\s]/g, "").replace(",", "."),
            ) || 0;
          const ancienICClient =
            parseFloat(
              (row.IC_client || "0").replace(/[€\s]/g, "").replace(",", "."),
            ) || 0;

          if (Math.abs(ancienTarifC - nouveauTarifC) > 0.001) {
            row.tarif = nouveauTarifC.toFixed(2) + " €";
            modifie = true;
          }
          if (Math.abs(ancienICClient - nouveauICClient) > 0.001) {
            row.IC_client = nouveauICClient.toFixed(2) + " €";
            modifie = true;
          }
        }

        if (associe && matos) {
          const nouveauTarifA = getTarifAffrete(associe, matos);
          const nouveauICAffrete = getICAffrete(associe, matos);
          const ancienTarifA =
            parseFloat(
              (row.tarif_affrete || "0")
                .replace(/[€\s]/g, "")
                .replace(",", "."),
            ) || 0;
          const ancienICAffrete =
            parseFloat(
              (row.IC_affrete || "0").replace(/[€\s]/g, "").replace(",", "."),
            ) || 0;

          if (Math.abs(ancienTarifA - nouveauTarifA) > 0.001) {
            row.tarif_affrete = nouveauTarifA.toFixed(2) + " €";
            modifie = true;
          }
          if (Math.abs(ancienICAffrete - nouveauICAffrete) > 0.001) {
            row.IC_affrete = nouveauICAffrete.toFixed(2) + " €";
            modifie = true;
          }
        }

        if ((client && matos) || (associe && matos)) {
          row.snapshot = captureSnapshot(client, matos, associe);
          modifie = true;
        }
      });

      if (modifie) {
        localStorage.setItem(key, JSON.stringify(dayData));
      }
    } catch (e) {
      console.error("Erreur lors de la mise à jour du jour " + key, e);
    }
  });

  calculerTousLesTotaux();
}

function forceRecalculTousLesTarifs() {
  const rows = document.querySelectorAll("#table-body tr");
  rows.forEach((tr) => {
    updateRowTarif(tr);
  });
  calculerTousLesTotaux();
}

// ====================================================================
// 📋  8. LIGNES DU PLANNING
// ====================================================================

function appliquerCouleurAssocie(input) {
  const cell = input.parentElement;
  cell.className = "matos-cell";
  const val = input.value.toUpperCase();
  if (val) {
    if (baseAssocies.includes(val))
      cell.classList.add("color-" + val.toLowerCase());
    else {
      cell.style.backgroundColor = "#e9ecef";
      cell.style.color = "#000";
    }
  } else {
    cell.style.backgroundColor = "";
  }
}

function verifierDoublon(inputModifie) {
  const nomSaisi = inputModifie.value.trim();
  if (nomSaisi === "") return;
  const toutesLesLignes = document.querySelectorAll("#table-body tr");
  let existeDeja = false;
  toutesLesLignes.forEach((tr) => {
    const inputChauffeur = tr.querySelector(".select-chauffeur");
    if (inputChauffeur && inputChauffeur !== inputModifie) {
      if (inputChauffeur.value.trim() === nomSaisi) existeDeja = true;
    }
  });
  if (existeDeja) {
    if (
      !confirm(
        `ATTENTION : ${nomSaisi} est déjà sur le planning.\n\nVoulez-vous quand même l'ajouter ?`,
      )
    ) {
      inputModifie.value = "";
      saveCurrentDay();
    }
  }
}

function extractContactName(fullValue) {
  if (!fullValue) return "";
  if (fullValue.includes(" - ")) return fullValue.split(" - ")[0].trim();
  return fullValue;
}

function extractPhoneNumber(fullValue) {
  if (!fullValue) return "";
  if (fullValue.includes(" - ")) {
    const parts = fullValue.split(" - ");
    return parts[1] ? parts[1].trim() : "";
  }
  const contactName = fullValue.trim();
  if (transportConfigTelephone[contactName])
    return transportConfigTelephone[contactName].numero || "";
  return "";
}

function displayTelephoneInCell(td, fullValue) {
  const contactName = extractContactName(fullValue);
  if (!contactName) {
    td.innerHTML = `<input class="combo-input select-telephone" list="list-telephone" value="" placeholder="Contact...">`;
    const newInput = td.querySelector(".select-telephone");
    if (newInput) {
      newInput.addEventListener("input", () => {
        if (newInput.value.trim() === "") {
          td.setAttribute("data-telephone-full", "");
          displayTelephoneInCell(td, "");
          saveCurrentDay();
        }
      });
      newInput.addEventListener("change", () => {
        const newContactName = newInput.value.trim();
        if (!newContactName) {
          td.setAttribute("data-telephone-full", "");
          displayTelephoneInCell(td, "");
        } else if (transportConfigTelephone[newContactName]) {
          const phoneNumber =
            transportConfigTelephone[newContactName].numero || "";
          const fullValue = phoneNumber
            ? `${newContactName} - ${phoneNumber}`
            : newContactName;
          td.setAttribute("data-telephone-full", fullValue);
          displayTelephoneInCell(td, fullValue);
        } else {
          td.setAttribute("data-telephone-full", newContactName);
          displayTelephoneInCell(td, newContactName);
        }
        saveCurrentDay();
      });
    }
    return;
  }
  let phoneNumber = "";
  if (transportConfigTelephone[contactName])
    phoneNumber = transportConfigTelephone[contactName].numero || "";
  else if (fullValue.includes(" - "))
    phoneNumber = extractPhoneNumber(fullValue);
  td.innerHTML = `<div class="telephone-wrapper"><input class="combo-input select-telephone" list="list-telephone" value="${contactName}" placeholder="Contact...">${phoneNumber ? `<div class="telephone-number">${phoneNumber}</div>` : ""}</div>`;
  const newInput = td.querySelector(".select-telephone");
  if (newInput) {
    newInput.addEventListener("input", () => {
      if (newInput.value.trim() === "") {
        td.setAttribute("data-telephone-full", "");
        displayTelephoneInCell(td, "");
        saveCurrentDay();
      }
    });
    newInput.addEventListener("change", () => {
      const newContactName = newInput.value.trim();
      if (!newContactName) {
        td.setAttribute("data-telephone-full", "");
        displayTelephoneInCell(td, "");
      } else if (transportConfigTelephone[newContactName]) {
        const newPhoneNumber =
          transportConfigTelephone[newContactName].numero || "";
        const newFullValue = newPhoneNumber
          ? `${newContactName} - ${newPhoneNumber}`
          : newContactName;
        td.setAttribute("data-telephone-full", newFullValue);
        displayTelephoneInCell(td, newFullValue);
      } else {
        td.setAttribute("data-telephone-full", newContactName);
        displayTelephoneInCell(td, newContactName);
      }
      saveCurrentDay();
    });
  }
}

function ajouterLigne(data = null) {
  const datePicker = document.getElementById("current-date-picker");
  const dateSelectionnee = new Date(datePicker.value);
  const dateFormatee = dateSelectionnee.toLocaleDateString("fr-FR", {
    day: "2-digit",
    month: "2-digit",
    year: "2-digit",
  });
  const tr = document.createElement("tr");
  tr.style.backgroundColor = "";
  tr.removeAttribute("style");
  const getPlaceholderClass = (val, defaultText) =>
    !val || val === defaultText ? "placeholder-style" : "";
  const valChantier = data ? data.chantier : "Chantier...";
  const valConsigne = data ? data.consigne : "Consigne...";
  const valChantierNum = data ? data.n_chantier : "N° Chantier...";
  const valCommande = data ? data.n_commande : "N° Commande...";
  const valBon = data ? data.n_bon : "N° Bon...";
  const valComm = data ? data.commentaire : "Commentaire...";

  tr.innerHTML = `
          <td class="matos-cell"><input class="combo-input select-associe" list="list-associe" value="${data ? data.associe : ""}" placeholder="Associé..."></td>
          <td contenteditable="true" class="${getPlaceholderClass(valChantier, "Chantier...")}">${valChantier}</td>
          <td contenteditable="true">${data && data.date && data.date.includes("/20") ? data.date.replace("/20", "/") : data ? data.date : dateFormatee}</td>
          <td class="matos-cell"><input class="combo-input select-vehicule" list="list-vehicule" value="${data ? data.vehicule : ""}" placeholder="Véhicule..."></td>
          <td class="matos-cell"><input class="combo-input select-client" list="list-client" value="${data ? data.client : ""}" placeholder="Client..."></td>
          <td class="matos-cell col-telephone" data-telephone-full="${data ? data.telephone : ""}"><input class="combo-input select-telephone" list="list-telephone" value="${data ? extractContactName(data.telephone) : ""}" placeholder="Contact..."></td>
          <td contenteditable="true" class="col-consigne ${getPlaceholderClass(valConsigne, "Consigne...")}">${valConsigne}</td>
          <td class="matos-cell"><input class="combo-input select-chauffeur" list="list-chauffeur" value="${data ? data.chauffeur : ""}" placeholder="Chauffeur..."></td>
          <td contenteditable="true" class="${getPlaceholderClass(valChantierNum, "N° Chantier...")}">${valChantierNum}</td>
          <td contenteditable="true" class="${getPlaceholderClass(valCommande, "N° Commande...")}">${valCommande}</td>
          <td contenteditable="true" class="cell-bon ${getPlaceholderClass(valBon, "N° Bon...")}">${valBon}</td>
          <td class="cell-tarif" contenteditable="true" style="padding:4px">${data && data.tarif ? data.tarif : "0.00 €"}</td>
          <td class="cell-supplement-client" contenteditable="true" style="padding:4px">${data && data.supplement_client ? data.supplement_client : "0.00 €"}</td>
          <td class="cell-IC-client" contenteditable="true" style="padding:4px">${data && data.IC_client ? data.IC_client : "0.00 €"}</td>
          <td class="cell-tarif-affrete" contenteditable="true" style="padding:4px">${data && data.tarif_affrete ? data.tarif_affrete : "0.00 €"}</td>
          <td class="cell-supplement-affrete" contenteditable="true" style="padding:4px">${data && data.supplement_affrete ? data.supplement_affrete : "0.00 €"}</td>
          <td class="cell-IC-affrete" contenteditable="true" style="padding:4px">${data && data.IC_affrete ? data.IC_affrete : "0.00 €"}</td>
          <td contenteditable="true" class="${getPlaceholderClass(valComm, "Commentaire...")}">${valComm}</td>
          <td class="hide-export" style="text-align:center; white-space:nowrap;">
            <button style="color:red; background:none; border:none; cursor:pointer; font-size:16px;" onclick="captureUndoStateBefore(); if(confirm('Supprimer ?')){this.closest('tr').remove(); saveCurrentDay(); calculerTousLesTotaux();}">×</button>
            <button style="color:orange; background:none; border:none; cursor:pointer; font-size:16px;" onclick="captureUndoStateBefore(); dupliquerLigne(this)">📑</button>
          </td>
        `;

  tr.querySelectorAll('[contenteditable="true"]').forEach((cell) => {
    cell.addEventListener("input", () => {
      if (
        cell.classList.contains("cell-tarif") ||
        cell.classList.contains("cell-tarif-affrete")
      )
        updateRowTarif(tr);
      if (cell.classList.contains("cell-bon")) updateRowHighlight(tr);
      saveCurrentDay();
    });
    cell.addEventListener("blur", () => {
      const contenuSaisi = cell.textContent.trim();
      const placeholders = [
        "Chantier...",
        "Consigne...",
        "N° Chantier...",
        "N° Commande...",
        "N° Bon...",
        "Commentaire...",
      ];
      const estDerniereLigne = !tr.nextElementSibling;
      if (
        estDerniereLigne &&
        contenuSaisi !== "" &&
        !placeholders.includes(contenuSaisi)
      )
        ajouterLigne();
      const isNumericCell =
        cell.classList.contains("cell-tarif") ||
        cell.classList.contains("cell-tarif-affrete") ||
        cell.classList.contains("cell-supplement-client") ||
        cell.classList.contains("cell-supplement-affrete") ||
        cell.classList.contains("cell-IC-client") ||
        cell.classList.contains("cell-IC-affrete");
      if (isNumericCell) {
        let val = parseFloat(contenuSaisi.replace(",", ".")) || 0;
        cell.textContent = val.toFixed(2) + " €";
        calculerTousLesTotaux();
      } else if (contenuSaisi === "") {
        const index = cell.cellIndex;
        const map = {
          1: "Chantier...",
          6: "Consigne...",
          8: "N° Chantier...",
          9: "N° Commande...",
          10: "N° Bon...",
          17: "Commentaire...",
        };
        if (map[index]) {
          cell.textContent = map[index];
          cell.classList.add("placeholder-style");
        }
      }
      saveCurrentDay();
    });
    cell.addEventListener("focus", () => {
      captureUndoStateBefore();
      const placeholders = [
        "Chantier...",
        "Consigne...",
        "N° Chantier...",
        "N° Commande...",
        "N° Bon...",
        "Commentaire...",
      ];
      const isNumericCell =
        cell.classList.contains("cell-tarif") ||
        cell.classList.contains("cell-tarif-affrete") ||
        cell.classList.contains("cell-supplement-client") ||
        cell.classList.contains("cell-supplement-affrete") ||
        cell.classList.contains("cell-IC-client") ||
        cell.classList.contains("cell-IC-affrete");
      if (!isNumericCell) {
        if (placeholders.includes(cell.textContent.trim())) {
          cell.textContent = "";
          cell.classList.remove("placeholder-style");
        }
      } else {
        let val = cell.textContent.replace(" €", "").trim();
        if (val === "0.00") cell.textContent = "";
        else cell.textContent = val;
      }
    });
  });

  tr.querySelectorAll(".combo-input").forEach((input) => {
    input.addEventListener("focus", () => {
      captureUndoStateBefore();
    });
    input.addEventListener("change", () => {
      const estDerniereLigne = !tr.nextElementSibling;
      if (estDerniereLigne && input.value.trim() !== "") ajouterLigne();
      if (input.classList.contains("select-vehicule")) {
        const vehicule = input.value.trim();
        if (!vehicule) {
          tr.querySelector(".cell-tarif").textContent = "0.00 €";
          tr.querySelector(".cell-IC-client").textContent = "0.00 €";
          tr.querySelector(".cell-tarif-affrete").textContent = "0.00 €";
          tr.querySelector(".cell-IC-affrete").textContent = "0.00 €";
          delete tr.dataset.snapshot;
          saveCurrentDay();
          calculerTousLesTotaux();
          return;
        }
      }
      if (input.classList.contains("select-telephone")) {
        const td = input.closest("td");
        const contactName = input.value.trim();
        if (transportConfigTelephone[contactName]) {
          const phoneNumber =
            transportConfigTelephone[contactName].numero || "";
          const fullValue = phoneNumber
            ? `${contactName} - ${phoneNumber}`
            : contactName;
          td.setAttribute("data-telephone-full", fullValue);
          displayTelephoneInCell(td, fullValue);
          const newInput = td.querySelector(".select-telephone");
          if (newInput) {
            newInput.addEventListener("change", () => {
              const newContactName = newInput.value.trim();
              if (transportConfigTelephone[newContactName]) {
                const newPhoneNumber =
                  transportConfigTelephone[newContactName].numero || "";
                const newFullValue = newPhoneNumber
                  ? `${newContactName} - ${newPhoneNumber}`
                  : newContactName;
                td.setAttribute("data-telephone-full", newFullValue);
                displayTelephoneInCell(td, newFullValue);
              }
              saveCurrentDay();
            });
          }
        } else {
          td.setAttribute("data-telephone-full", contactName);
          displayTelephoneInCell(td, contactName);
        }
      }
      if (
        input.classList.contains("select-client") ||
        input.classList.contains("select-vehicule") ||
        input.classList.contains("select-associe")
      ) {
        const client = tr.querySelector(".select-client").value;
        const vehicule = tr.querySelector(".select-vehicule").value;
        const associe = tr.querySelector(".select-associe").value;
        if (vehicule && ((client && vehicule) || (associe && vehicule))) {
          const snapshot = captureSnapshot(client, vehicule, associe);
          tr.dataset.snapshot = JSON.stringify(snapshot);
          applySnapshotToRow(tr, snapshot);
        }
      }
      saveCurrentDay();
      if (input.classList.contains("select-associe"))
        appliquerCouleurAssocie(input);
    });
    if (input.classList.contains("select-associe")) {
      input.addEventListener("blur", () => {
        trierPlanningParAssocie();
      });
    }
  });

  tableBody.appendChild(tr);
  const inputAssocie = tr.querySelector(".select-associe");
  const cells = tr.querySelectorAll("td");

  if (data) {
    appliquerCouleurAssocie(inputAssocie);
    if (data.couleurs_cellules) {
      Object.keys(data.couleurs_cellules).forEach((index) => {
        const td = cells[parseInt(index)];
        if (td) td.style.backgroundColor = data.couleurs_cellules[index];
      });
    } else if (data.couleur) {
      cells.forEach((td) => {
        if (!td.classList.contains("matos-cell"))
          td.style.backgroundColor = data.couleur;
      });
    }
    if (data.snapshot) tr.dataset.snapshot = JSON.stringify(data.snapshot);
    if (data.telephone) {
      const telTd = tr.querySelector(".col-telephone");
      if (telTd) {
        displayTelephoneInCell(telTd, data.telephone);
        const newInput = telTd.querySelector(".select-telephone");
        if (newInput) {
          newInput.addEventListener("change", () => {
            const contactName = newInput.value.trim();
            if (transportConfigTelephone[contactName]) {
              const phoneNumber =
                transportConfigTelephone[contactName].numero || "";
              const fullValue = phoneNumber
                ? `${contactName} - ${phoneNumber}`
                : contactName;
              telTd.setAttribute("data-telephone-full", fullValue);
              displayTelephoneInCell(telTd, fullValue);
            }
            saveCurrentDay();
          });
        }
      }
    }
  } else {
    tr.style.backgroundColor = "";
    tr.removeAttribute("style");
    cells.forEach((td) => {
      td.style.backgroundColor = "";
      td.style.color = "";
    });
    appliquerCouleurAssocie(inputAssocie);
  }

  calculerTousLesTotaux();
  updateRowHighlight(tr);
  filterTable();
  trierPlanningParAssocie();
  const tid = setTimeout(() => {
    pendingTimeouts = pendingTimeouts.filter((t) => t !== tid);
    updateRowTarif(tr);
  }, 50);
  pendingTimeouts.push(tid);
}

function supprimerLigne(btn) {
  if (confirm("Voulez-vous vraiment supprimer cette ligne ?")) {
    btn.closest("tr").remove();
    if (tableBody.querySelectorAll("tr").length === 0) ajouterLigne();
    saveCurrentDay();
    calculerTousLesTotaux();
  }
}

function dupliquerLigne(bouton) {
  const ligne = bouton.closest("tr");
  const data = {
    associe: ligne.querySelector(".select-associe").value,
    chantier: ligne.cells[1].innerText,
    date: ligne.cells[2].innerText,
    vehicule: ligne.querySelector(".select-vehicule").value,
    client: ligne.querySelector(".select-client").value,
    telephone: ligne.querySelector(".select-telephone").value,
    consigne: ligne.cells[6].innerText,
    chauffeur: ligne.querySelector(".select-chauffeur").value,
    n_chantier: ligne.cells[8].innerText,
    n_commande: ligne.cells[9].innerText,
    n_bon: ligne.cells[10].innerText,
    tarif: ligne.cells[11].innerText,
    supplement_client: ligne.cells[12].innerText,
    IC_client: ligne.cells[13].innerText,
    tarif_affrete: ligne.cells[14].innerText,
    supplement_affrete: ligne.cells[15].innerText,
    IC_affrete: ligne.cells[16].innerText,
    commentaire: ligne.cells[17].innerText,
  };
  ajouterLigne(data);
  const nouvelleLigne = tableBody.lastElementChild;
  updateRowTarif(nouvelleLigne);
  calculerTousLesTotaux();
  saveCurrentDay();
}

function setColorToSelectedRow(color) {
  const rows = document.querySelectorAll("tr:has(.cell-selected), .row-marked");
  const currentDate = document.getElementById("current-date-picker").value;
  const allSavedColors = JSON.parse(
    localStorage.getItem("planningColors") || "{}",
  );
  if (!allSavedColors[currentDate]) allSavedColors[currentDate] = {};
  const savedColors = allSavedColors[currentDate];
  if (rows.length === 0) {
    alert("Veuillez d'abord sélectionner au moins une cellule ou une ligne.");
    return;
  }
  captureUndoStateBefore();
  rows.forEach((row) => {
    const rowId = row.innerText.trim().split("\t")[0] + row.rowIndex;
    const cells = row.querySelectorAll("td");
    if (color === "#ffee00") {
      const selectedCellsInRow = row.querySelectorAll(".cell-selected");
      if (!savedColors[rowId] || typeof savedColors[rowId] !== "object")
        savedColors[rowId] = { type: "partial", cells: {} };
      selectedCellsInRow.forEach((td) => {
        const cellIndex = td.cellIndex;
        td.style.backgroundColor = color;
        savedColors[rowId].cells[cellIndex] = color;
      });
    } else if (color === "#ffffff") {
      const existingData = savedColors[rowId];
      if (
        existingData &&
        typeof existingData === "object" &&
        existingData.type === "partial"
      ) {
        const selectedCellsInRow = row.querySelectorAll(".cell-selected");
        selectedCellsInRow.forEach((td) => {
          const cellIndex = td.cellIndex;
          td.style.backgroundColor = "";
          delete savedColors[rowId].cells[cellIndex];
        });
        if (Object.keys(savedColors[rowId].cells).length === 0)
          delete savedColors[rowId];
      } else {
        cells.forEach((td, index) => {
          const hasSpecificColor =
            td.style.backgroundColor !== "" &&
            td.style.backgroundColor !== "transparent";
          if (index === 0 && hasSpecificColor) return;
          td.style.backgroundColor = "";
        });
        delete savedColors[rowId];
      }
    } else {
      cells.forEach((td, index) => {
        const hasSpecificColor =
          td.style.backgroundColor !== "" &&
          td.style.backgroundColor !== "transparent";
        if (index === 0 && hasSpecificColor) return;
        td.style.backgroundColor = color;
      });
      savedColors[rowId] = color;
    }
  });
  saveCurrentDay();
}

let startRowIndex = null;

document.getElementById("my-table").addEventListener("mousedown", function (e) {
  const td = e.target.closest("td");
  if (!td) return;
  if (e.button === 2) return;
  isMouseDown = true;
  startRowIndex = td.parentElement.rowIndex;
  if (!e.ctrlKey)
    document
      .querySelectorAll(".cell-selected")
      .forEach((el) => el.classList.remove("cell-selected"));
  td.classList.add("cell-selected");
});

document.getElementById("my-table").addEventListener("mouseover", function (e) {
  if (isMouseDown) {
    const td = e.target.closest("td");
    if (td && td.parentElement.rowIndex === startRowIndex)
      td.classList.add("cell-selected");
  }
});

window.addEventListener("mouseup", () => {
  isMouseDown = false;
  startRowIndex = null;
});

document.addEventListener("click", function (e) {
  if (e.button === 2) return;
  if (!e.target.closest("#my-table"))
    document
      .querySelectorAll(".cell-selected")
      .forEach((el) => el.classList.remove("cell-selected"));
});

document
  .getElementById("my-table")
  .addEventListener("contextmenu", function (e) {
    const td = e.target.closest("td");
    if (td && !td.classList.contains("cell-selected")) {
      document
        .querySelectorAll(".cell-selected")
        .forEach((el) => el.classList.remove("cell-selected"));
      td.classList.add("cell-selected");
    }
    if (td) {
      const selection = window.getSelection();
      const range = document.createRange();
      const input = td.querySelector("input");
      try {
        if (input) {
          input.select();
          input.focus();
        } else {
          range.selectNodeContents(td);
          selection.removeAllRanges();
          selection.addRange(range);
        }
      } catch (err) {}
    }
  });

document.addEventListener("copy", function (e) {
  const selectedCells = document.querySelectorAll(".cell-selected");
  if (selectedCells.length === 0) return;
  let rows = {};
  const sortedCells = Array.from(selectedCells).sort((a, b) => {
    if (a.parentElement.rowIndex !== b.parentElement.rowIndex)
      return a.parentElement.rowIndex - b.parentElement.rowIndex;
    return a.cellIndex - b.cellIndex;
  });
  sortedCells.forEach((td) => {
    const rowIndex = td.parentElement.rowIndex;
    if (!rows[rowIndex]) rows[rowIndex] = [];
    const input = td.querySelector("input");
    let val = input ? input.value.trim() : td.innerText.trim();
    const placeholders = [
      "Chantier...",
      "Véhicule...",
      "Client...",
      "Contact...",
      "Consigne...",
      "Chauffeur...",
      "N° Chantier...",
      "N° Commande...",
      "N° Bon...",
      "Commentaire...",
    ];
    if (placeholders.includes(val)) val = "";
    if (val !== "") {
      if (input && input.classList.contains("select-telephone")) {
        const contactName = val;
        let phoneNumber = "";
        if (transportConfigTelephone[contactName])
          phoneNumber = transportConfigTelephone[contactName].numero || "";
        val = phoneNumber
          ? "Contact : " + contactName + " " + phoneNumber
          : "Contact : " + contactName;
      } else if (input) {
        if (input.classList.contains("select-vehicule"))
          val = "Véhicule : " + val;
        else if (input.classList.contains("select-client"))
          val = "Client : " + val;
        else if (input.classList.contains("select-chauffeur"))
          val = "Chauffeur : " + val;
      } else {
        const cellIndex = td.cellIndex;
        if (td.classList.contains("col-consigne")) val = "Consigne : " + val;
        else if (cellIndex === 8) val = "N° Chantier : " + val;
        else if (cellIndex === 9) val = "N° Commande : " + val;
        else if (cellIndex === 10) val = "N° Bon : " + val;
      }
    }
    rows[rowIndex].push(val);
  });
  const textResult = Object.values(rows)
    .map((rowCells) => rowCells.filter((c) => c !== "").join(" | "))
    .join("\n");
  if (textResult) {
    e.clipboardData.setData("text/plain", textResult);
    e.preventDefault();
    afficherNotification(
      `${selectedCells.length} cellule(s) copiée(s) !`,
      "#4caf50",
    );
  }
});

window.onclick = function (event) {
  if (!event.target.matches(".btn")) {
    const menus = [
      "sous-menu-add",
      "sous-menu-copy",
      "sous-menu-photos",
      "sous-menu-config",
    ];
    menus.forEach((menuId) => {
      const el = document.getElementById(menuId);
      if (el) el.style.display = "none";
    });
  }
};

// ====================================================================
// 📄  9. COPIER / COLLER JOURNÉE
// ====================================================================

function captureSnapshot(client, vehicule, associe) {
  const snapshot = {
    tarif_client_snapshot: 0,
    indice_client_snapshot: 0,
    tarif_affrete_snapshot: 0,
    indice_affrete_snapshot: 0,
    date_capture: new Date().toISOString(),
  };
  if (client && vehicule) {
    const groupeNom = transportConfigClient[client];
    if (
      groupeNom &&
      transportConfigGroup[groupeNom] &&
      transportConfigGroup[groupeNom][vehicule]
    ) {
      snapshot.tarif_client_snapshot = getTarifClient(client, vehicule);
      let indice =
        parseFloat(transportConfigGroup[groupeNom][vehicule].indice) || 0;
      if (indice === 0) {
        const cat = transportConfigMateriel[vehicule];
        if (cat) {
          const { k, cm1, c0 } = getCatValues(groupeNom, cat);
          indice = calculateIndice(k, cm1, c0);
        }
      }
      snapshot.indice_client_snapshot = indice;
    }
  }
  if (associe && vehicule) {
    if (
      transportConfigAssocie[associe] &&
      transportConfigAssocie[associe][vehicule]
    ) {
      snapshot.tarif_affrete_snapshot = getTarifAffrete(associe, vehicule);
      let indice =
        parseFloat(transportConfigAssocie[associe][vehicule].indice) || 0;
      if (indice === 0) {
        const cat = transportConfigMateriel[vehicule];
        if (cat) {
          const { k, cm1, c0 } = getCatValues(associe, cat);
          indice = calculateIndice(k, cm1, c0);
        }
      }
      snapshot.indice_affrete_snapshot = indice;
    }
  }
  return snapshot;
}

function applySnapshotToRow(tr, snapshot) {
  if (!snapshot) return;
  const cellTarifClient = tr.querySelector(".cell-tarif");
  const cellTarifAffrete = tr.querySelector(".cell-tarif-affrete");
  const cellICClient = tr.querySelector(".cell-IC-client");
  const cellICAffrete = tr.querySelector(".cell-IC-affrete");
  if (snapshot.tarif_client_snapshot !== undefined) {
    const tarifC = snapshot.tarif_client_snapshot;
    const indiceC = snapshot.indice_client_snapshot || 0;
    const icC = tarifC * (indiceC / 100);
    cellTarifClient.textContent = tarifC.toFixed(2) + " €";
    cellICClient.textContent = icC.toFixed(2) + " €";
    cellICClient.style.color =
      icC > 0 ? "#28a745" : icC < 0 ? "#dc3545" : "#95a5a6";
    cellICClient.style.fontWeight = icC !== 0 ? "bold" : "normal";
  }
  if (snapshot.tarif_affrete_snapshot !== undefined) {
    const tarifA = snapshot.tarif_affrete_snapshot;
    const indiceA = snapshot.indice_affrete_snapshot || 0;
    const icA = tarifA * (indiceA / 100);
    cellTarifAffrete.textContent = tarifA.toFixed(2) + " €";
    cellICAffrete.textContent = icA.toFixed(2) + " €";
    cellICAffrete.style.color =
      icA > 0 ? "#28a745" : icA < 0 ? "#dc3545" : "#95a5a6";
    cellICAffrete.style.fontWeight = icA !== 0 ? "bold" : "normal";
  }
}

function copierJournee() {
  const rows = [];
  document.querySelectorAll("#table-body tr").forEach((tr) => {
    rows.push({
      associe: tr.querySelector(".select-associe").value,
      chantier: tr.cells[1].innerText,
      date: tr.cells[2].innerText,
      vehicule: tr.querySelector(".select-vehicule").value,
      client: tr.querySelector(".select-client").value,
      telephone: tr.querySelector(".select-telephone").value,
      consigne: tr.cells[6].innerText,
      chauffeur: tr.querySelector(".select-chauffeur").value,
      n_chantier: tr.cells[8].innerText,
      n_commande: tr.cells[9].innerText,
      n_bon: tr.cells[10].innerText,
      tarif: tr.cells[11].innerText,
      supplement_client: tr.cells[12].innerText,
      IC_client: tr.cells[13].innerText,
      tarif_affrete: tr.cells[14].innerText,
      supplement_affrete: tr.cells[15].innerText,
      IC_affrete: tr.cells[16].innerText,
      commentaire: tr.cells[17].innerText,
    });
  });
  if (rows.length === 0) return;
  localStorage.setItem("planning_copy_buffer", JSON.stringify(rows));
  afficherNotification("Journée copiée", "#2196F3");
}

function collerJournee() {
  const buffer = localStorage.getItem("planning_copy_buffer");
  if (!buffer) {
    alert("Aucune journée copiée.");
    return;
  }
  if (!confirm("Remplacer la journée actuelle par celle copiée ?")) return;
  captureUndoStateBefore();
  const data = JSON.parse(buffer);

  pendingTimeouts.forEach((t) => clearTimeout(t));
  pendingTimeouts = [];
  tableBody.innerHTML = "";
  const datePicker = document.getElementById("current-date-picker");
  const dateFormatee = new Date(datePicker.value).toLocaleDateString("fr-FR");
  estEnTrainDeColler = true;
  data.forEach((row) => {
    row.date = dateFormatee;
    ajouterLigne(row);
  });
  saveCurrentDay();
  setTimeout(() => {
    estEnTrainDeColler = false;
    saveCurrentDay();
    afficherNotification("Journée collée", "#2196F3");
  }, 150);
}

// ====================================================================
// 📊  10. TOTAUX & MARGES
// ====================================================================

function calculerTotal() {
  let total = 0;
  document.querySelectorAll("#table-body tr").forEach((tr) => {
    const content = tr.cells[11].textContent
      .replace(/[€\s]/g, "")
      .replace(",", ".");
    const val = parseFloat(content);
    if (!isNaN(val)) total += val;
  });
  document.getElementById("total-price").textContent = total.toFixed(2) + " €";
}

function calculerTotalSupplement_client() {
  let total = 0;
  document.querySelectorAll("#table-body tr").forEach((tr) => {
    const content = tr.cells[12].textContent
      .replace(/[€\s]/g, "")
      .replace(",", ".");
    const val = parseFloat(content);
    if (!isNaN(val)) total += val;
  });
  document.getElementById("total-price-supplement-client").textContent =
    total.toFixed(2) + " €";
}

function calculerTotalIC_client() {
  let total = 0;
  document.querySelectorAll("#table-body tr").forEach((tr) => {
    const content = tr.cells[13].textContent
      .replace(/[€\s]/g, "")
      .replace(",", ".");
    const val = parseFloat(content);
    if (!isNaN(val)) total += val;
  });
  document.getElementById("total-price-IC-client").textContent =
    total.toFixed(2) + " €";
}

function calculerMarge_Client() {
  let totalMarge = 0;
  document.querySelectorAll("#table-body tr").forEach((tr) => {
    const tarif = parseFloat(tr.querySelector(".cell-tarif")?.textContent) || 0;
    const supp =
      parseFloat(tr.querySelector(".cell-supplement-client")?.textContent) || 0;
    const ic =
      parseFloat(tr.querySelector(".cell-IC-client")?.textContent) || 0;
    totalMarge += tarif + supp + ic;
  });
  const displayMarge = document.getElementById("total-marge-client");
  if (displayMarge) displayMarge.textContent = totalMarge.toFixed(2) + " €";
  return totalMarge;
}

function calculerTotalAffrete() {
  let total = 0;
  document.querySelectorAll("#table-body tr").forEach((tr) => {
    const content = tr.cells[14].textContent
      .replace(/[€\s]/g, "")
      .replace(",", ".");
    const val = parseFloat(content);
    if (!isNaN(val)) total += val;
  });
  document.getElementById("total-price-affrete").textContent =
    total.toFixed(2) + " €";
}

function calculerTotalSupplement_affrete() {
  let total = 0;
  document.querySelectorAll("#table-body tr").forEach((tr) => {
    const content = tr.cells[15].textContent
      .replace(/[€\s]/g, "")
      .replace(",", ".");
    const val = parseFloat(content);
    if (!isNaN(val)) total += val;
  });
  document.getElementById("total-price-supplement-affrete").textContent =
    total.toFixed(2) + " €";
}

function calculerTotalIC_affrete() {
  let total = 0;
  document.querySelectorAll("#table-body tr").forEach((tr) => {
    const content = tr.cells[16].textContent
      .replace(/[€\s]/g, "")
      .replace(",", ".");
    const val = parseFloat(content);
    if (!isNaN(val)) total += val;
  });
  document.getElementById("total-price-IC-affrete").textContent =
    total.toFixed(2) + " €";
}

function calculerMarge_Affrete() {
  let totalMarge = 0;
  document.querySelectorAll("#table-body tr").forEach((tr) => {
    const tarif =
      parseFloat(tr.querySelector(".cell-tarif-affrete")?.textContent) || 0;
    const supp =
      parseFloat(tr.querySelector(".cell-supplement-affrete")?.textContent) ||
      0;
    const ic =
      parseFloat(tr.querySelector(".cell-IC-affrete")?.textContent) || 0;
    totalMarge += tarif + supp + ic;
  });
  const displayMarge = document.getElementById("total-marge-affrete");
  if (displayMarge) displayMarge.textContent = totalMarge.toFixed(2) + " €";
  return totalMarge;
}

function calculerTousLesTotaux() {
  if (typeof calculerTotal === "function") calculerTotal();
  if (typeof calculerTotalSupplement_client === "function")
    calculerTotalSupplement_client();
  if (typeof calculerTotalIC_client === "function") calculerTotalIC_client();
  if (typeof calculerTotalAffrete === "function") calculerTotalAffrete();
  if (typeof calculerTotalSupplement_affrete === "function")
    calculerTotalSupplement_affrete();
  if (typeof calculerTotalIC_affrete === "function") calculerTotalIC_affrete();
  const totalCA = calculerMarge_Client();
  const totalCout = calculerMarge_Affrete();
  const margeNette = totalCA - totalCout;
  const displayMarge = document.getElementById("marge-nette");
  if (displayMarge) {
    displayMarge.textContent = margeNette.toFixed(2) + " €";
    displayMarge.style.color = margeNette >= 0 ? "#28a745" : "#dc3545";
  }
}

// ====================================================================
// 👥  11. CONFIG — CLIENTS & GROUPES
// ====================================================================

function addGlobalMatos(type) {
  let inputId =
    type === "client" ? "global-matos-client" : "global-matos-associe";
  let val = document.getElementById(inputId).value.trim().toUpperCase();
  if (!val) return;
  for (let groupName in transportConfigGroup) {
    if (!transportConfigGroup[groupName][val])
      transportConfigGroup[groupName][val] = { tarif: 0, indice: 0 };
  }
  for (let associeName in transportConfigAssocie) {
    if (!transportConfigAssocie[associeName][val])
      transportConfigAssocie[associeName][val] = { tarif: 0, indice: 0 };
  }
  if (!transportConfigAssocie[pilierCommunKey])
    transportConfigAssocie[pilierCommunKey] = {};
  if (!transportConfigAssocie[pilierCommunKey][val])
    transportConfigAssocie[pilierCommunKey][val] = {
      tarif: {},
      indice: {},
    };
  localStorage.setItem(
    "transport_config_group",
    JSON.stringify(transportConfigGroup),
  );
  localStorage.setItem(
    "transport_config_associe",
    JSON.stringify(transportConfigAssocie),
  );
  document.getElementById(inputId).value = "";
  renderConfigUI();
  renderConfigUI2();
  if (typeof updateMatosDatalist === "function") updateMatosDatalist();
}

function deleteGlobalMatos(type, matosNom) {
  if (
    confirm(
      `Supprimer définitivement "${matosNom}" de TOUS les Groupes (Clients) et de TOUS les Associés ?`,
    )
  ) {
    for (let g in transportConfigGroup) {
      if (transportConfigGroup[g][matosNom] !== undefined)
        delete transportConfigGroup[g][matosNom];
    }
    for (let a in transportConfigAssocie) {
      if (transportConfigAssocie[a][matosNom] !== undefined)
        delete transportConfigAssocie[a][matosNom];
    }
    if (transportConfigMateriel[matosNom])
      delete transportConfigMateriel[matosNom];
    localStorage.setItem(
      "transport_config_group",
      JSON.stringify(transportConfigGroup),
    );
    localStorage.setItem(
      "transport_config_associe",
      JSON.stringify(transportConfigAssocie),
    );
    localStorage.setItem(
      "transport_config_materiel",
      JSON.stringify(transportConfigMateriel),
    );
    renderConfigUI();
    renderConfigUI2();
    if (typeof updateMatosDatalist === "function") updateMatosDatalist();
    const sn = document.getElementById("save-notif");
    if (sn) {
      sn.textContent = `"${matosNom}" supprimé partout.`;
      sn.style.display = "inline";
      setTimeout(() => (sn.style.display = "none"), 2000);
    }
  }
}

function updateCategoryValue(group, cat, field, value) {
  setCatValue(group, cat, field, value);
  localStorage.setItem(
    "transportConfigGroupInfo",
    JSON.stringify(transportConfigGroupInfo),
  );
  renderConfigUI();
  renderConfigUI2();
  refreshAllPlanningTarifs();
}

function toggleUseCM(groupOrAssocie, cat, checked) {
  setCatValue(groupOrAssocie, cat, "useCM", checked);
  localStorage.setItem(
    "transportConfigGroupInfo",
    JSON.stringify(transportConfigGroupInfo),
  );
  renderConfigUI();
  renderConfigUI2();
  refreshAllPlanningTarifs();
}

function buildCatBlock(
  groupOrAssocie,
  cat,
  moisAnnee,
  annee,
  isPilierCommun = false,
) {
  const catInfo = getCatValues(groupOrAssocie, cat);
  const calculatedIndice = calculateIndice(catInfo.k, catInfo.cm1, catInfo.c0);
  const usesCM = catInfo.usesCM;
  const cmLabel = usesCM ? "CM" : "CM-1";
  const cmText = usesCM
    ? "Utilise CM (mois actuel)"
    : "Utilise CM-1 (mois précédent)";
  const cmColor = usesCM ? "#7b2d8b" : "#e67e00";
  const cmBg = usesCM ? "#f9f0ff" : "#fff8ee";
  const cmBorder = usesCM ? "#7b2d8b" : "#e67e00";
  if (isPilierCommun) {
    return `
            <div style="margin-bottom:12px;text-align:left;border-bottom:1px solid #ddd;padding-bottom:8px;background:#f0f4ff;padding:8px;border-radius:4px;border-left:3px solid #0f3460;">
              <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:8px;">
                <label style="font-size:0.85em;color:#0f3460;font-weight:bold;">${cat}</label>
                <button onclick="deleteCategorie('${cat}')" style="background:none;border:none;color:red;cursor:pointer;font-size:12px;">✕</button>
              </div>
              <div style="background:#e8efff;border-radius:4px;padding:6px;margin-bottom:6px;">
                <div style="font-size:10px;color:#555;font-weight:bold;margin-bottom:4px;">📅 Valeurs annuelles (${annee})</div>
                <div style="display:grid;grid-template-columns:1fr 1fr;gap:8px;">
                  <div><label style="font-size:11px;color:#666;display:block;margin-bottom:3px;">K</label><input type="number" step="0.001" value="${catInfo.k !== 0 ? catInfo.k : ""}" onchange="updatePilierCommunCatValue('${cat}','k',this.value)" style="width:100%;padding:5px;font-size:0.9em;text-align:center;border:1px solid #ddd;border-radius:3px;"></div>
                  <div><label style="font-size:11px;color:#666;display:block;margin-bottom:3px;">C0</label><input type="number" step="0.01" value="${catInfo.c0 !== 0 ? catInfo.c0 : ""}" onchange="updatePilierCommunCatValue('${cat}','c0',this.value)" style="width:100%;padding:5px;font-size:0.9em;text-align:center;border:1px solid #ddd;border-radius:3px;"></div>
                </div>
              </div>
              <div style="background:${cmBg};border-radius:4px;padding:6px;margin-bottom:6px;">
                <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:4px;">
                  <div style="font-size:10px;color:#555;font-weight:bold;">📆 Valeur mensuelle (${moisAnnee})</div>
                  <label style="font-size:10px;cursor:pointer;display:flex;align-items:center;gap:4px;color:${cmColor};font-weight:bold;" title="Cochez si ce groupe utilise le CM du mois actuel">
                    <input type="checkbox" ${usesCM ? "checked" : ""} onchange="togglePilierCommunUseCM('${cat}',this.checked)" style="cursor:pointer;">
                    ${cmText}
                  </label>
                </div>
                <div style="display:grid;grid-template-columns:1fr 1fr;gap:8px;">
                  <div><label style="font-size:11px;color:${cmColor};display:block;margin-bottom:3px;font-weight:bold;">${cmLabel}</label><input type="number" step="0.01" value="${catInfo.cm1 !== 0 ? catInfo.cm1 : ""}" onchange="updatePilierCommunCatValue('${cat}','cm1',this.value)" style="width:100%;padding:5px;font-size:0.9em;text-align:center;border:2px solid ${cmBorder};border-radius:3px;background:${cmBg};"></div>
                  <div><label style="font-size:11px;color:#666;display:block;margin-bottom:3px;">📊 Indice (%)</label><input type="text" value="${calculatedIndice !== 0 ? calculatedIndice.toFixed(2) + " %" : ""}" readonly style="width:100%;padding:5px;font-size:1em;text-align:center;background:#e8f5e9;font-weight:bold;color:#27ae60;border:2px solid #27ae60;border-radius:3px;"></div>
                </div>
              </div>
            </div>`;
  }

  return `
          <div style="margin-bottom:12px;text-align:left;border-bottom:1px solid #ddd;padding-bottom:8px;background:#f9f9f9;padding:8px;border-radius:4px;">
            <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:8px;">
              <label style="font-size:0.85em;color:#333;font-weight:bold;">${cat}</label>
              <button onclick="deleteCategorie('${cat}')" style="background:none;border:none;color:red;cursor:pointer;font-size:12px;">✕</button>
            </div>
            <div style="background:#eef6ff;border-radius:4px;padding:6px;margin-bottom:6px;">
              <div style="font-size:10px;color:#555;font-weight:bold;margin-bottom:4px;">📅 Valeurs annuelles (${annee})</div>
              <div style="display:grid;grid-template-columns:1fr 1fr;gap:8px;">
                <div><label style="font-size:11px;color:#666;display:block;margin-bottom:3px;">K</label><input type="number" step="0.001" value="${catInfo.k !== 0 ? catInfo.k : ""}" onchange="updateCategoryValue('${groupOrAssocie}','${cat}','k',this.value)" style="width:100%;padding:5px;font-size:0.9em;text-align:center;border:1px solid #ddd;border-radius:3px;"></div>
                <div><label style="font-size:11px;color:#666;display:block;margin-bottom:3px;">C0</label><input type="number" step="0.01" value="${catInfo.c0 !== 0 ? catInfo.c0 : ""}" onchange="updateCategoryValue('${groupOrAssocie}','${cat}','c0',this.value)" style="width:100%;padding:5px;font-size:0.9em;text-align:center;border:1px solid #ddd;border-radius:3px;"></div>
              </div>
            </div>
            <div style="background:${cmBg};border-radius:4px;padding:6px;margin-bottom:6px;">
              <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:4px;">
                <div style="font-size:10px;color:#555;font-weight:bold;">📆 Valeur mensuelle (${moisAnnee})</div>
                <label style="font-size:10px;cursor:pointer;display:flex;align-items:center;gap:4px;color:${cmColor};font-weight:bold;" title="Cochez si ce groupe utilise le CM du mois actuel">
                  <input type="checkbox" ${usesCM ? "checked" : ""} onchange="toggleUseCM('${groupOrAssocie}','${cat}',this.checked)" style="cursor:pointer;">
                  ${cmText}
                </label>
              </div>
              <div style="display:grid;grid-template-columns:1fr 1fr;gap:8px;">
                <div><label style="font-size:11px;color:${cmColor};display:block;margin-bottom:3px;font-weight:bold;">${cmLabel}</label><input type="number" step="0.01" value="${catInfo.cm1 !== 0 ? catInfo.cm1 : ""}" onchange="updateCategoryValue('${groupOrAssocie}','${cat}','cm1',this.value)" style="width:100%;padding:5px;font-size:0.9em;text-align:center;border:2px solid ${cmBorder};border-radius:3px;background:${cmBg};"></div>
                <div><label style="font-size:11px;color:#666;display:block;margin-bottom:3px;">📊 Indice (%)</label><input type="text" value="${calculatedIndice !== 0 ? calculatedIndice.toFixed(2) + " %" : ""}" readonly style="width:100%;padding:5px;font-size:1em;text-align:center;background:#e8f5e9;font-weight:bold;color:#27ae60;border:2px solid #27ae60;border-radius:3px;"></div>
              </div>
            </div>
          </div>`;
}

function renderConfigUI() {
  const container = document.getElementById("config-container");
  let html = `<h3>Tarifs par Groupe</h3><table class="config-table"><thead><tr><th>Groupe</th><th style="width: 20%;">Catégorie</th><th style="width: 25%;">Matériel</th><th style="width: 13%;">Tarif (€)</th><th style="width: 10%;">Indice (%)</th><th>Action</th></tr></thead><tbody>`;
  const groups = Object.keys(transportConfigGroup).sort();
  const categorie = Object.keys(transportConfigCategorie).sort();
  if (groups.length === 0) {
    html += `<tr><td colspan="6"><i>Aucun groupe configuré...</i></td></tr>`;
  } else {
    groups.forEach((group, index) => {
      if (index > 0)
        html += `<tr style="background-color: white;"><td colspan="6" style="height: 25px; border: none;"></td></tr>`;
      let matosKeys = Object.keys(transportConfigGroup[group]).sort();
      if (matosKeys.length === 0) {
        html += `<tr><td class="group-header"><strong>${group}</strong></td><td colspan="5">Aucun matériel - <button onclick="deleteGroup('${group}')" class="btn-delete">Supprimer</button></td></tr>`;
      } else {
        matosKeys.forEach((matos, matosIdx) => {
          const val = transportConfigGroup[group][matos];
          const currentCat = transportConfigMateriel[matos] || "";
          let indiceAafficher = 0;
          let isUsingCategory = false;
          const indiceData = val.indice;
          if (typeof indiceData === "object")
            indiceAafficher = indiceData[moisAnnee] || 0;
          else indiceAafficher = indiceData || 0;
          if (indiceAafficher === 0 && currentCat) {
            const { k, cm1, c0 } = getCatValues(group, currentCat);
            indiceAafficher = calculateIndice(k, cm1, c0);
            isUsingCategory = true;
          }
          html += `<tr>
                  ${matosIdx === 0 ? `<td rowspan="${matosKeys.length}" class="group-header" style="vertical-align: top;"><strong>${group}</strong><br><button onclick="deleteGroup('${group}')" class="btn-delete" style="margin-top:5px;">Supprimer</button><div style="margin-top:15px; border-top: 1px solid #ccc; padding-top:10px;"><p style="font-weight:bold; font-size:0.85em; margin-bottom:10px; color:#555;">Indices Carburant par Catégorie :</p>${categorie.map((cat) => buildCatBlock(group, cat, moisAnnee, annee)).join("")}</div></td>` : ""}
                  <td style="text-align:left; padding-left:10px;"><select onchange="setMaterielCategorie('${matos}', this.value)" style="width:100%; padding: 4px;"><option value="">-- Sans --</option>${categorie.map((c) => `<option value="${c}" ${currentCat === c ? "selected" : ""}>${c}</option>`).join("")}</select></td>
                  <td>${matos}</td>
                  <td><input type="number" step="0.01" value="${typeof val.tarif === "object" && val.tarif !== null ? val.tarif[annee] || 0 : val.tarif || 0}" onchange="updateGroupPrice('${group}', '${matos}', this.value)" style="width:100%;"></td>
                  <td><input type="number" step="0.01" value="${indiceAafficher !== 0 ? indiceAafficher : "0"}" placeholder="${isUsingCategory}" onchange="updateGroupIndice('${group}', '${matos}', this.value)" style="width:100%;"></td>
                  <td><button class="del-matos-btn" onclick="deleteGlobalMatos('client', '${matos}')" style="color:red; cursor:pointer;">x</button></td>
                </tr>`;
        });
      }
    });
  }
  html += `</tbody></table>`;
  html += `<h3 style="margin-top:30px;">Assignation des Clients</h3><table class="config-table"><thead><tr><th style="width: 40%;">Client</th><th style="width: 40%;">Groupe assigné</th><th style="width: 20%;">Action</th></tr></thead><tbody>`;
  const clients = Object.keys(transportConfigClient).sort();
  if (clients.length === 0) {
    html += `<tr><td colspan="3"><i>Aucun client configuré.</i></td></tr>`;
  } else {
    clients.forEach((client) => {
      const assignedGroup = transportConfigClient[client];
      html += `<tr><td><strong>${client}</strong></td><td><select onchange="setClientGroup('${client}', this.value)" style="width:100%; padding:4px;"><option value="">-- Aucun groupe --</option>${groups.map((g) => `<option value="${g}" ${assignedGroup === g ? "selected" : ""}>${g}</option>`).join("")}</select></td><td><button onclick="deleteClient('${client}')" class="btn-delete">Supprimer</button></td></tr>`;
    });
  }
  container.innerHTML = html + `</tbody></table>`;
}

function renderCategorieManager() {
  const container = document.getElementById("categorie-manager-container");
  if (!container) return;
  const cats = Object.keys(transportConfigCategorie).sort();
  let html = `<h3>Gestion des Catégories</h3><ul style="list-style: none; padding: 0; max-width: 400px;">`;
  cats.forEach((cat) => {
    html += `<li style="display: flex; justify-content: space-between; align-items: center; padding: 8px; border-bottom: 1px solid #eee;"><span>${cat}</span><button onclick="deleteCategorie('${cat}')" style="color: white; background: #dc3545; border: none; padding: 4px 8px; border-radius: 4px; cursor: pointer;">Supprimer</button></li>`;
  });
  html += `</ul>`;
  container.innerHTML = html;
}

function updateGroupCatIndice(associe, categorie, valeur) {
  setCatValue(associe, categorie, "cm1", valeur);
  renderConfigUI();
  renderConfigUI2();
  refreshAllPlanningTarifs();
}

function updateGroupPrice(group, matos, price) {
  if (!transportConfigGroup[group][matos])
    transportConfigGroup[group][matos] = { tarif: {}, indice: 0 };
  if (
    typeof transportConfigGroup[group][matos].tarif !== "object" ||
    transportConfigGroup[group][matos].tarif === null
  ) {
    const ancienTarif = transportConfigGroup[group][matos].tarif || 0;
    transportConfigGroup[group][matos].tarif = {};
    if (ancienTarif)
      transportConfigGroup[group][matos].tarif[annee] = ancienTarif;
  }
  transportConfigGroup[group][matos].tarif[annee] = parseFloat(price) || 0;
  localStorage.setItem(
    "transport_config_group",
    JSON.stringify(transportConfigGroup),
  );
  refreshAllPlanningTarifs();
}

function sauvegarderConfig() {
  localStorage.setItem(
    "transport_config_group",
    JSON.stringify(transportConfigGroup),
  );
  localStorage.setItem(
    "transportConfigGroupInfo",
    JSON.stringify(transportConfigGroupInfo),
  );
  localStorage.setItem(
    "transport_config_associe",
    JSON.stringify(transportConfigAssocie),
  );
}

function updateGroupIndice(group, matos, indice) {
  if (!transportConfigGroup[group][matos])
    transportConfigGroup[group][matos] = { tarif: 0, indice: {} };
  if (typeof transportConfigGroup[group][matos].indice !== "object")
    transportConfigGroup[group][matos].indice = {};
  transportConfigGroup[group][matos].indice[moisAnnee] =
    parseFloat(indice) || 0;
  sauvegarderConfig();
  renderConfigUI();
  refreshAllPlanningTarifs();
}

function updateConfigValue(c, m, f, v) {
  if (typeof transportConfigClient[c][m] !== "object") {
    const old = transportConfigClient[c][m] || 0;
    transportConfigClient[c][m] = { tarif: old, indice: 0 };
  }
  transportConfigClient[c][m][f] = parseFloat(v) || 0;
}

function addNewClient() {
  const name = document
    .getElementById("new-client-name")
    .value.trim()
    .toUpperCase();
  if (!name || transportConfigClient[name]) return;
  transportConfigClient[name] = "";
  localStorage.setItem(
    "transport_config_client",
    JSON.stringify(transportConfigClient),
  );
  document.getElementById("new-client-name").value = "";
  renderConfigUI();
  updateMatosDatalist();
}

function deleteClient(c) {
  if (confirm("Supprimer ce client ?")) {
    delete transportConfigClient[c];
    localStorage.setItem(
      "transport_config_client",
      JSON.stringify(transportConfigClient),
    );
    renderConfigUI();
  }
}

function addNewGroup() {
  const name = document
    .getElementById("new-group-name")
    .value.trim()
    .toUpperCase();
  if (!name) return;
  if (transportConfigGroup[name]) {
    alert("Ce groupe existe déjà");
    return;
  }
  const fullMatosList = new Set(defaultMatosList);
  for (let group in transportConfigGroup)
    Object.keys(transportConfigGroup[group]).forEach((m) =>
      fullMatosList.add(m),
    );
  for (let associe in transportConfigAssocie)
    Object.keys(transportConfigAssocie[associe]).forEach((m) =>
      fullMatosList.add(m),
    );
  transportConfigGroup[name] = {};
  fullMatosList.forEach((m) => {
    transportConfigGroup[name][m] = { tarif: 0, indice: 0 };
  });
  localStorage.setItem(
    "transport_config_group",
    JSON.stringify(transportConfigGroup),
  );
  document.getElementById("new-group-name").value = "";
  renderConfigUI();
  const sn = document.getElementById("save-notif");
  sn.textContent = "Groupe " + name + " créé avec tous les matériels actuels !";
  sn.style.display = "inline";
  setTimeout(() => (sn.style.display = "none"), 2000);
}

function deleteGroup(name) {
  if (confirm(`Supprimer le groupe "${name}" ?`)) {
    delete transportConfigGroup[name];
    for (let c in transportConfigClient) {
      if (transportConfigClient[c] === name) transportConfigClient[c] = "";
    }
    localStorage.setItem(
      "transport_config_group",
      JSON.stringify(transportConfigGroup),
    );
    localStorage.setItem(
      "transport_config_client",
      JSON.stringify(transportConfigClient),
    );
    renderConfigUI();
  }
}

function getClientGroup(client) {
  for (let group in transportConfigGroup) {
    if (transportConfigGroup[group].includes(client)) return group;
  }
  return "";
}

function setClientGroup(client, groupName) {
  transportConfigClient[client] = groupName;
  localStorage.setItem(
    "transport_config_client",
    JSON.stringify(transportConfigClient),
  );
  const sn = document.getElementById("save-notif");
  sn.textContent = "Client assigné au groupe !";
  sn.style.display = "inline";
  setTimeout(() => (sn.style.display = "none"), 2000);
}

function addCategorie(type) {
  const inputId =
    type === "associe" ? "global-categorie-associe" : "global-categorie-client";
  const input = document.getElementById(inputId);
  if (!input) return;
  const name = input.value.trim().toUpperCase();
  if (!name) return;
  if (!transportConfigCategorie[name]) {
    transportConfigCategorie[name] = true;
    localStorage.setItem(
      "transport_config_categorie",
      JSON.stringify(transportConfigCategorie),
    );
    input.value = "";
    renderConfigUI();
    renderConfigUI2();
    const sn = document.getElementById("save-notif");
    sn.textContent = "Catégorie " + name + " créée !";
    sn.style.display = "inline";
    setTimeout(() => (sn.style.display = "none"), 2000);
  }
}

function deleteCategorie(catNom) {
  if (
    confirm(
      `Voulez-vous vraiment supprimer la catégorie "${catNom}" ?\n(Les matériels n'auront plus de catégorie assignée).`,
    )
  ) {
    delete transportConfigCategorie[catNom];
    localStorage.setItem(
      "transport_config_categorie",
      JSON.stringify(transportConfigCategorie),
    );
    for (let group in transportConfigGroupInfo) {
      if (transportConfigGroupInfo[group][catNom] !== undefined)
        delete transportConfigGroupInfo[group][catNom];
    }
    localStorage.setItem(
      "transportConfigGroupInfo",
      JSON.stringify(transportConfigGroupInfo),
    );
    for (let matos in transportConfigMateriel) {
      if (transportConfigMateriel[matos] === catNom)
        transportConfigMateriel[matos] = "";
    }
    localStorage.setItem(
      "transport_config_materiel",
      JSON.stringify(transportConfigMateriel),
    );
    renderConfigUI();
    renderConfigUI2();
    if (typeof renderCategorieManager === "function") renderCategorieManager();
    const sn = document.getElementById("save-notif");
    if (sn) {
      sn.textContent = `Catégorie "${catNom}" supprimée.`;
      sn.style.display = "inline";
      setTimeout(() => (sn.style.display = "none"), 2000);
    }
  }
}

function setMaterielCategorie(matos, categorieName) {
  transportConfigMateriel[matos] = categorieName;
  localStorage.setItem(
    "transport_config_materiel",
    JSON.stringify(transportConfigMateriel),
  );
  renderConfigUI();
  renderConfigUI2();
  refreshAllPlanningTarifs();
  const sn = document.getElementById("save-notif");
  if (sn) {
    sn.textContent = "Catégorie mise à jour !";
    sn.style.display = "inline";
    setTimeout(() => (sn.style.display = "none"), 2000);
  }
}

// ====================================================================
// 🤝  12. CONFIG — ASSOCIÉS
// ====================================================================

function renderConfigUI2() {
  const container = document.getElementById("config-container-2");
  const associeListDatalist = document.getElementById("list-associe");
  if (!container) return;

  ensurePiliersCommunsExists();

  associeListDatalist.innerHTML = "";
  const allAssociesSet = new Set([...Object.keys(transportConfigAssocie)]);
  allAssociesSet.delete(pilierCommunKey);

  Array.from(allAssociesSet)
    .sort()
    .forEach((name) => {
      associeListDatalist.innerHTML += `<option value="${name}">`;
    });

  const groupePiliers = Array.from(allAssociesSet)
    .filter((name) => baseAssocies.includes(name.toUpperCase()))
    .sort((a, b) => a.localeCompare(b));
  const groupeAffretes = Array.from(allAssociesSet)
    .filter((name) => !baseAssocies.includes(name.toUpperCase()))
    .sort((a, b) => a.localeCompare(b));
  const sortedEntries = [...groupePiliers, ...groupeAffretes];
  const categories = Object.keys(transportConfigCategorie).sort();
  const allGlobalMaterials = Object.keys(transportConfigMateriel).sort();

  const communMatos = Array.from(
    new Set([
      ...allGlobalMaterials,
      ...Object.keys(transportConfigAssocie[pilierCommunKey] || {}),
    ]),
  ).sort();

  let htmlCommun = `<div class="piliers-communs-section">
          <div class="piliers-communs-header">
            <div style="flex:1;">
              <strong style="font-size:1.1em;">PILIERS COMMUNS</strong>
              <div style="font-size:0.8em;opacity:0.85;margin-top:2px;">Tarifs et indices partagés entre
                <span style="display:inline-block;padding:1px 6px;border-radius:10px;font-size:11px;background:#fffb00;color:#333;margin:0 2px">ALLARD</span>
                <span style="display:inline-block;padding:1px 6px;border-radius:10px;font-size:11px;background:#00eeff;color:#333;margin:0 2px">LANDAIS</span>
                <span style="display:inline-block;padding:1px 6px;border-radius:10px;font-size:11px;background:#e27aff;color:#333;margin:0 2px">STD</span>
              </div>
            </div>
          </div>
          <div class="piliers-communs-body">
            <table class="config-table"><thead><tr><th>Indices Carburant / Cat.</th><th style="width:20%;">Catégorie</th><th style="width:25%;">Matériel</th><th style="width:13%;">Tarif Commun (€)</th><th style="width:10%;">Indice Commun (%)</th><th>Action</th></tr></thead><tbody>`;

  if (communMatos.length === 0) {
    htmlCommun += `<tr><td colspan="6"><i>Aucun matériel. Ajoutez du matériel ci-dessus.</i></td></tr>`;
  } else {
    const communRowSpan = communMatos.length;
    communMatos.forEach((matos, matosIdx) => {
      if (!transportConfigAssocie[pilierCommunKey][matos])
        transportConfigAssocie[pilierCommunKey][matos] = {
          tarif: {},
          indice: {},
        };
      const val = transportConfigAssocie[pilierCommunKey][matos];
      const currentCat = transportConfigMateriel[matos] || "";
      const tarifVal =
        typeof val.tarif === "object" && val.tarif !== null
          ? val.tarif[annee] || 0
          : val.tarif || 0;
      let indiceAafficher = 0;
      let isUsingCategory = false;
      const indiceData = val.indice;
      if (typeof indiceData === "object")
        indiceAafficher = indiceData[moisAnnee] || 0;
      else indiceAafficher = indiceData || 0;
      if (indiceAafficher === 0 && currentCat) {
        const { k, cm1, c0 } = getCatValues(pilierCommunKey, currentCat);
        indiceAafficher = calculateIndice(k, cm1, c0);
        isUsingCategory = true;
      }
      htmlCommun += `<tr style="background:#f0f4ff;">
              ${
                matosIdx === 0
                  ? `<td rowspan="${communRowSpan}" class="group-header" style="vertical-align:top;background:#e8efff;">
                <div style="padding:5px;">
                  <p style="font-weight:bold;font-size:0.85em;margin-bottom:10px;color:#0f3460;">Indices Carburant Communs :</p>
                  ${categories.map((cat) => buildCatBlock(pilierCommunKey, cat, moisAnnee, annee, true)).join("")}
                  ${categories.length === 0 ? '<i style="font-size:11px;color:#888;">Aucune catégorie créée</i>' : ""}
                </div>
              </td>`
                  : ""
              }
              <td style="text-align:left;padding-left:10px;"><select onchange="setMaterielCategorie('${matos}',this.value)" style="width:100%;padding:4px;"><option value="">-- Sans --</option>${categories.map((c) => `<option value="${c}" ${currentCat === c ? "selected" : ""}>${c}</option>`).join("")}</select></td>
              <td>${matos}</td>
              <td><input type="number" step="0.01" value="${tarifVal || 0}" onchange="updatePilierCommunTarif('${matos}',this.value)" style="width:100%;background:#e8efff;border:1px solid #7a9fd4;border-radius:3px;padding:4px;"></td>
              <td><input type="number" step="0.01" value="${indiceAafficher !== 0 ? indiceAafficher : "0"}" placeholder="${isUsingCategory ? "auto" : ""}" onchange="updatePilierCommunIndice('${matos}',this.value)" style="width:100%;background:#e8efff;border:1px solid #7a9fd4;border-radius:3px;padding:4px;"></td>
              <td><button class="del-matos-btn" onclick="deleteGlobalMatos('associe','${matos}')" style="color:red;cursor:pointer;">x</button></td>
            </tr>`;
    });
  }
  htmlCommun += `</tbody></table></div></div>`;

  let html = `<table class="config-table"><thead><tr><th>Associé</th><th style="width:20%;">Catégorie</th><th style="width:25%;">Matériel</th><th style="width:13%;">Tarif (€)</th><th style="width:10%;">Indice (%)</th><th>Actions</th></tr></thead><tbody>`;

  html += `<tr>
          <td colspan="6" style="text-align:center;color:#000;padding:15px;background-color:#f2f2f2;border:1px solid #ccc;font-size:0.95em">
            TARIFS INDIVIDUELS PILIERS —
            <span style="display:inline-block;padding:2px 8px;border-radius:12px;font-size:11px;background:#fffb00;color:#333;border:1px solid #ccc;margin:0 3px">ALLARD</span>
            <span style="display:inline-block;padding:2px 8px;border-radius:12px;font-size:11px;background:#00eeff;color:#333;border:1px solid #ccc;margin:0 3px">LANDAIS</span>
            <span style="display:inline-block;padding:2px 8px;border-radius:12px;font-size:11px;background:#e27aff;color:#333;border:1px solid #ccc;margin:0 3px">STD</span>
            <div>
              <span style="font-size:11px;color:#666;margin-left:10px;">
                <span class="badge-sync badge-commun"></span> synchronisé avec Piliers Communs &nbsp;|&nbsp;
                <span class="badge-sync badge-custom"></span> valeur individuelle
              </span>
            </div>
          </td>
        </tr>`;

  if (sortedEntries.length === 0) {
    html += `<tr><td colspan="6"><i>Aucun associé configuré.</i></td></tr>`;
  } else {
    sortedEntries.forEach((entry, index) => {
      if (index > 0)
        html += `<tr style="background-color:white;"><td colspan="6" style="height:25px;border:none;"></td></tr>`;
      const estPilier = baseAssocies.includes(entry.toUpperCase());
      const precedentEtaitPilier =
        index > 0 &&
        baseAssocies.includes(sortedEntries[index - 1].toUpperCase());
      if (index > 0 && !estPilier && precedentEtaitPilier)
        html += `<tr style="background-color:#f1f3f4;"><td colspan="6" style="text-align:center;font-weight:bold;color:#5f6368;padding:10px;border:1px solid #dee2e6;">AUTRES AFFRÉTÉS</td></tr>`;
      if (!transportConfigAssocie[entry]) transportConfigAssocie[entry] = {};
      const keys = Array.from(
        new Set([
          ...allGlobalMaterials,
          ...Object.keys(transportConfigAssocie[entry]),
        ]),
      ).sort();
      const rowSpanCount = Math.max(1, keys.length);

      let leftBlockBg = "#f8f9fa";
      if (estPilier) {
        if (entry.toUpperCase() === "ALLARD") leftBlockBg = "#fffde7";
        else if (entry.toUpperCase() === "LANDAIS") leftBlockBg = "#e0feff";
        else if (entry.toUpperCase() === "STD") leftBlockBg = "#f8e8ff";
      }

      const leftBlockHTML = `<td rowspan="${rowSpanCount}" class="group-header" style="vertical-align:top;background-color:${leftBlockBg};">
              <strong>${entry}</strong><br>
              <button class="btn-delete" style="margin-top:5px;" onclick="deleteEntry2('${entry}')">Supprimer</button>
              <div style="margin-top:15px;border-top:1px solid #ccc;padding-top:10px;">
                <p style="font-weight:bold;font-size:0.85em;margin-bottom:10px;color:#555;">Indices Carburant / Cat :</p>
                ${categories.map((cat) => buildCatBlock(entry, cat, moisAnnee, annee)).join("")}
              </div>
            </td>`;

      keys.forEach((k, idx) => {
        if (!transportConfigAssocie[entry][k])
          transportConfigAssocie[entry][k] = { tarif: {}, indice: {} };
        const val = transportConfigAssocie[entry][k];
        const currentCat = transportConfigMateriel[k] || "";
        let indiceAafficher = 0;
        let isUsingCategory = false;
        const indiceDataPilier = val.indice;
        if (typeof indiceDataPilier === "object" && indiceDataPilier !== null) {
          indiceAafficher = parseFloat(indiceDataPilier[moisAnnee]) || 0;
        } else {
          indiceAafficher = parseFloat(indiceDataPilier) || 0;
        }
        if (indiceAafficher === 0 && currentCat) {
          const { k: kVal, cm1, c0 } = getCatValues(entry, currentCat);
          indiceAafficher = calculateIndice(kVal, cm1, c0);
          isUsingCategory = true;
        }
        const tarifVal =
          typeof val.tarif === "object" && val.tarif !== null
            ? val.tarif[annee] || 0
            : val.tarif || 0;

        let syncBadgeHTML = "";
        let rowStyleClass = "";
        if (estPilier) {
          const isOverridden = isPilierOverridden(entry, k);
          if (isOverridden) {
            syncBadgeHTML = `<button class="btn-reset-to-common" onclick="resetPilierToCommon('${entry}','${k}')">↩</button>`;
            rowStyleClass = `class="pilier-row-custom"`;
          } else {
            syncBadgeHTML = "";
            rowStyleClass = `class="pilier-row-synced"`;
          }
        }

        const onChangeTarif = estPilier
          ? `updatePilierIndividuelTarif('${entry}','${k}',this.value);renderConfigUI2();`
          : `updateAssocieTarif('${entry}','${k}',this.value)`;
        const onChangeIndice = estPilier
          ? `updatePilierIndividuelIndice('${entry}','${k}',this.value);renderConfigUI2();`
          : `updateAssocieIndice('${entry}','${k}',this.value)`;

        const inputBg =
          estPilier && !isPilierOverridden(entry, k) ? "#f0f7ff" : "white";
        const inputBorder =
          estPilier && !isPilierOverridden(entry, k) ? "#7a9fd4" : "#ddd";

        html += `<tr ${rowStyleClass}>
                ${idx === 0 ? leftBlockHTML : ""}
                <td style="text-align:left;padding-left:10px;"><select onchange="setMaterielCategorie('${k}',this.value)" style="width:100%;padding:4px;"><option value="">-- Sans --</option>${categories.map((c) => `<option value="${c}" ${currentCat === c ? "selected" : ""}>${c}</option>`).join("")}</select></td>
                <td>${k}</td>
                <td>
                  <div style="display:flex;align-items:center;gap:4px;">
                    <input type="number" step="0.01" value="${tarifVal}" onchange="${onChangeTarif}" style="width:100%;background:${inputBg};border:1px solid ${inputBorder};border-radius:3px;padding:4px;">
                    ${syncBadgeHTML ? syncBadgeHTML : ""}
                  </div>
                </td>
                <td>
                  <input type="number" step="0.01" value="${indiceAafficher !== 0 ? indiceAafficher : "0"}" placeholder="${isUsingCategory ? "auto" : ""}" onchange="${onChangeIndice}" style="width:100%;background:${inputBg};border:1px solid ${inputBorder};border-radius:3px;padding:4px;">
                </td>
                <td><button class="del-matos-btn" onclick="deleteGlobalMatos('associe','${k}')" style="color:red;cursor:pointer;">x</button></td>
              </tr>`;
      });
    });
  }
  container.innerHTML = htmlCommun + html + `</tbody></table>`;
  trierDatalist("list-associe");
}

function updateAssocieTarif(entry, mat, val) {
  if (
    typeof transportConfigAssocie[entry][mat].tarif !== "object" ||
    transportConfigAssocie[entry][mat].tarif === null
  ) {
    const ancienTarif = transportConfigAssocie[entry][mat].tarif || 0;
    transportConfigAssocie[entry][mat].tarif = {};
    if (ancienTarif)
      transportConfigAssocie[entry][mat].tarif[annee] = ancienTarif;
  }
  transportConfigAssocie[entry][mat].tarif[annee] = parseFloat(val) || 0;
  localStorage.setItem(
    "transport_config_associe",
    JSON.stringify(transportConfigAssocie),
  );
  refreshAllPlanningTarifs();
}

function updateAssocieIndice(entry, mat, val) {
  if (!transportConfigAssocie[entry][mat])
    transportConfigAssocie[entry][mat] = { tarif: 0, indice: {} };
  if (typeof transportConfigAssocie[entry][mat].indice !== "object")
    transportConfigAssocie[entry][mat].indice = {};
  transportConfigAssocie[entry][mat].indice[moisAnnee] = parseFloat(val) || 0;
  sauvegarderConfig();
  renderConfigUI2();
  refreshAllPlanningTarifs();
}

function addNewClient2() {
  const input = document.getElementById("new-config2-name");
  const name = input.value.trim().toUpperCase();
  if (!name || transportConfigAssocie[name]) return;
  const fullMatosList = new Set(defaultMatosList);
  for (let group in transportConfigGroup)
    Object.keys(transportConfigGroup[group]).forEach((m) =>
      fullMatosList.add(m),
    );
  for (let associe in transportConfigAssocie)
    Object.keys(transportConfigAssocie[associe]).forEach((m) =>
      fullMatosList.add(m),
    );
  transportConfigAssocie[name] = {};
  fullMatosList.forEach((m) => {
    transportConfigAssocie[name][m] = { tarif: 0, indice: 0 };
  });
  if (baseAssocies.includes(name)) {
    fullMatosList.forEach((m) => syncPiliersFromCommon(m));
  }
  localStorage.setItem(
    "transport_config_associe",
    JSON.stringify(transportConfigAssocie),
  );
  input.value = "";
  renderConfigUI2();
  updateMatosDatalist();
}

function deleteEntry2(e) {
  if (confirm("Supprimer cet associé ?")) {
    delete transportConfigAssocie[e];
    localStorage.setItem(
      "transport_config_associe",
      JSON.stringify(transportConfigAssocie),
    );
    renderConfigUI2();
  }
}

// ====================================================================
// 📞  13. CONFIG — CHAUFFEURS & TÉLÉPHONES
// ====================================================================

function renderConfigUIChauffeur() {
  const container = document.getElementById("config-container-chauffeur");
  const chauffeurListDatalist = document.getElementById("list-chauffeur");
  chauffeurListDatalist.innerHTML = "";
  const allChauffeurs = new Set([
    ...defaultChauffeursList,
    ...Object.keys(transportConfigChauffeur),
  ]);
  Array.from(allChauffeurs)
    .sort()
    .forEach((name) => {
      chauffeurListDatalist.innerHTML += `<option value="${name}">`;
    });
  let html = `<table class="config-table"><thead><tr><th>Chauffeur</th><th>Actions</th></tr></thead><tbody>`;
  const sortedEntries = Object.keys(transportConfigChauffeur).sort();
  if (sortedEntries.length === 0) {
    html += `<tr><td colspan="2"><i>Aucun chauffeur configuré manuellement.</i></td></tr>`;
  } else {
    sortedEntries.forEach((entry) => {
      html += `<tr><td class="client-header">${entry}</td><td><button class="btn btn-reset" onclick="deleteChauffeur('${entry}')">Suppr. Chauffeur</button></td></tr>`;
    });
  }
  container.innerHTML = html + `</tbody></table>`;
  trierDatalist("list-chauffeur");
}

function addNewChauffeur() {
  const name = document
    .getElementById("new-config-chauffeur-name")
    .value.trim()
    .toUpperCase();
  if (!name || transportConfigChauffeur[name]) return;
  transportConfigChauffeur[name] = {};
  localStorage.setItem(
    "transport_config_chauffeur",
    JSON.stringify(transportConfigChauffeur),
  );
  document.getElementById("new-config-chauffeur-name").value = "";
  renderConfigUIChauffeur();
}

function deleteChauffeur(name) {
  if (confirm("Supprimer ce chauffeur ?")) {
    delete transportConfigChauffeur[name];
    localStorage.setItem(
      "transport_config_chauffeur",
      JSON.stringify(transportConfigChauffeur),
    );
    renderConfigUIChauffeur();
  }
}

function updateMatosDatalist() {
  const listVehicule = document.getElementById("list-vehicule");
  if (listVehicule) {
    const uniqueMatos = new Set(defaultMatosList);
    for (let group in transportConfigGroup)
      Object.keys(transportConfigGroup[group]).forEach((m) =>
        uniqueMatos.add(m),
      );
    for (let associe in transportConfigAssocie) {
      if (associe === pilierCommunKey) continue;
      Object.keys(transportConfigAssocie[associe]).forEach((m) =>
        uniqueMatos.add(m),
      );
    }
    listVehicule.innerHTML = "";
    Array.from(uniqueMatos)
      .sort()
      .forEach((m) => {
        const opt = document.createElement("option");
        opt.value = m;
        listVehicule.appendChild(opt);
      });
  }
  const listClient = document.getElementById("list-client");
  if (listClient) {
    listClient.innerHTML = "";
    Object.keys(transportConfigClient)
      .sort()
      .forEach((c) => {
        const opt = document.createElement("option");
        opt.value = c;
        listClient.appendChild(opt);
      });
  }
}

function addNewTelephone() {
  const clientInput = document.getElementById("new-config-telephone-client");
  const nameInput = document.getElementById("new-config-telephone-name");
  const numInput = document.getElementById("new-config-telephone-number");
  const name = nameInput.value.trim().toUpperCase();
  const client = clientInput.value.trim().toUpperCase();
  const numero = formatPhoneNumber(numInput.value.trim());
  if (name) {
    transportConfigTelephone[name] = { numero: numero, client: client };
    localStorage.setItem(
      "transport_config_telephone",
      JSON.stringify(transportConfigTelephone),
    );
    clientInput.value = "";
    nameInput.value = "";
    numInput.value = "";
    updateTelephoneDatalist();
    renderConfigTelephoneUI();
  }
}

function updateTelephoneDatalist() {
  const dl = document.getElementById("list-telephone");
  if (!dl) return;
  dl.innerHTML = "";
  Object.keys(transportConfigTelephone)
    .sort((a, b) => a.localeCompare(b))
    .forEach((name) => {
      const data = transportConfigTelephone[name];
      const opt = document.createElement("option");
      opt.value = name;
      opt.setAttribute("data-numero", data.numero || "");
      dl.appendChild(opt);
    });
}

function renderConfigTelephoneUI() {
  const container = document.getElementById("config-container-telephone");
  if (!container) return;
  let html = `<table class="config-table" style="width:100%; table-layout: fixed; border-collapse: collapse;"><thead><tr><th style="width: 25%;">Client</th><th style="width: 25%;">Contact</th><th style="width: 30%;">Numéro de Téléphone</th><th style="width: 20%;">Action</th></tr></thead><tbody>`;
  const contacts = Object.keys(transportConfigTelephone).sort((a, b) => {
    const clientA = transportConfigTelephone[a].client || "";
    const clientB = transportConfigTelephone[b].client || "";
    return clientA.localeCompare(clientB) || a.localeCompare(b);
  });
  if (contacts.length === 0) {
    html += `<tr><td colspan="4" style="text-align:center;"><i>Aucun contact configuré...</i></td></tr>`;
  } else {
    contacts.forEach((name) => {
      const tel = transportConfigTelephone[name].numero || "";
      const client = transportConfigTelephone[name].client || "";
      html += `<tr><td style="padding: 8px;" contenteditable="true" onblur="updateClientInConfig('${name}', this.innerText)">${client}</td><td style="padding: 8px;" contenteditable="true" onblur="updateNameInConfig('${name}', this.innerText)">${name}</td><td style="text-align: center;" contenteditable="true" onblur="updateTelInConfig('${name}', this.innerText)">${tel}</td><td style="text-align: center;"><button onclick="deleteTelephone('${name}')" class="btn-delete">Supprimer</button></td></tr>`;
    });
  }
  html += `</tbody></table>`;
  container.innerHTML = html;
}

function updateClientInConfig(name, newClientName) {
  transportConfigTelephone[name].client = newClientName.trim().toUpperCase();
  localStorage.setItem(
    "transport_config_telephone",
    JSON.stringify(transportConfigTelephone),
  );
}

function updateNameInConfig(oldName, newName) {
  if (oldName === newName) return;
  transportConfigTelephone[newName] = transportConfigTelephone[oldName];
  delete transportConfigTelephone[oldName];
  localStorage.setItem(
    "transport_config_telephone",
    JSON.stringify(transportConfigTelephone),
  );
  renderConfigTelephoneUI();
  updateTelephoneDatalist();
}

function updateTelInConfig(name, newNum) {
  const formattedNum = formatPhoneNumber(newNum);
  transportConfigTelephone[name].numero = formattedNum;
  localStorage.setItem(
    "transport_config_telephone",
    JSON.stringify(transportConfigTelephone),
  );
  renderConfigTelephoneUI();
  updateTelephoneDatalist();
}

function deleteTelephone(name) {
  if (confirm("Supprimer " + name + " ?")) {
    delete transportConfigTelephone[name];
    localStorage.setItem(
      "transport_config_telephone",
      JSON.stringify(transportConfigTelephone),
    );
    updateTelephoneDatalist();
    renderConfigTelephoneUI();
  }
}

// ====================================================================
// 📤  14. IMPRESSION
// ====================================================================

function imprimerPlanning() {
  const lignes = [];
  document.querySelectorAll("#table-body tr").forEach((tr) => {
    const contactName = tr.querySelector(".select-telephone")?.value || "";
    let phoneNumber = "";
    if (contactName && transportConfigTelephone[contactName])
      phoneNumber = transportConfigTelephone[contactName].numero || "";
    const PH = [
      "Chantier...",
      "Consigne...",
      "N° Chantier...",
      "N° Commande...",
      "N° Bon...",
      "Commentaire...",
      "Véhicule...",
      "Client...",
      "Associé...",
      "Chauffeur...",
      "Contact...",
    ];
    const cp = (v) => (PH.includes((v || "").trim()) ? "" : v || "");
    const associeVal = cp(tr.querySelector(".select-associe")?.value);
    const chantierVal = cp(tr.cells[1]?.innerText);
    const vehiculeVal = cp(tr.querySelector(".select-vehicule")?.value);
    const clientVal = cp(tr.querySelector(".select-client")?.value);
    const consigneVal = cp(tr.cells[6]?.innerText);
    const chauffeurVal = cp(tr.querySelector(".select-chauffeur")?.value);
    const NchantierVal = cp(tr.cells[8]?.innerText);
    const NcommandeVal = cp(tr.cells[9]?.innerText);
    if (
      !associeVal &&
      !chantierVal &&
      !vehiculeVal &&
      !clientVal &&
      !consigneVal &&
      !chauffeurVal &&
      !NchantierVal &&
      !NcommandeVal
    )
      return;
    lignes.push({
      associe: associeVal,
      chantier: chantierVal,
      date: tr.cells[2]?.innerText || "",
      vehicule: vehiculeVal,
      client: clientVal,
      telephoneNom: contactName,
      telephoneNumero: phoneNumber,
      consigne: consigneVal,
      chauffeur: chauffeurVal,
      n_chantier: NchantierVal,
      n_commande: NcommandeVal,
      couleur: tr.style.backgroundColor || "",
    });
  });
  const titre = document.getElementById("display-date").innerText;
  const fenetreImpression = window.open("", "", "height=600,width=800");
  fenetreImpression.document.write(
    `<html><head><title>Impression Planning</title><style>body{font-family:sans-serif;margin:20px;}h2{text-align:center;}table{border-collapse:collapse;width:100%;font-size:11px;}th,td{border:1px solid #ddd;padding:6px;text-align:left;vertical-align:top;}th{background-color:#007bff!important;color:white!important;-webkit-print-color-adjust:exact;print-color-adjust:exact;}.telephone-cell{white-space:normal;line-height:1.3;}.telephone-name{font-size:11px;color:#000;}.telephone-number{font-size:11px;color:#000;font-style:italic;}@media print{th{background-color:#007bff!important;color:white!important;}}</style></head><body><h2>${titre}</h2><table><thead><tr><th>Associé</th><th>Chantier</th><th>Véhicule</th><th>Client</th><th>Contact</th><th>Consigne</th><th>Chauffeur</th><th>N° Chantier</th><th>N° Commande</th></tr></thead><tbody>`,
  );
  lignes.forEach((ligne) => {
    const bgColor = ligne.couleur
      ? ` style="background-color:${ligne.couleur};-webkit-print-color-adjust:exact;print-color-adjust:exact;"`
      : "";
    let telephoneHTML = "";
    if (ligne.telephoneNom) {
      telephoneHTML = `<div class="telephone-name">${ligne.telephoneNom}</div>`;
      if (ligne.telephoneNumero)
        telephoneHTML += `<div class="telephone-number">${ligne.telephoneNumero}</div>`;
    }
    fenetreImpression.document.write(
      `<tr${bgColor}><td>${ligne.associe}</td><td>${ligne.chantier}</td><td>${ligne.vehicule}</td><td>${ligne.client}</td><td class="telephone-cell">${telephoneHTML}</td><td>${ligne.consigne}</td><td>${ligne.chauffeur}</td><td>${ligne.n_chantier}</td><td>${ligne.n_commande}</td></tr>`,
    );
  });
  fenetreImpression.document.write(`</tbody></table></body></html>`);
  fenetreImpression.document.close();
  setTimeout(() => {
    fenetreImpression.print();
    fenetreImpression.close();
  }, 500);
}

// ====================================================================
// 🔒  15. INTÉGRITÉ DES DONNÉES
// ====================================================================

function verifierIntegriteDonnees() {
  let problemesDetectes = [];
  try {
    const groups = JSON.parse(
      localStorage.getItem("transport_config_group") || "{}",
    );
    for (let groupName in groups) {
      for (let matos in groups[groupName]) {
        const val = groups[groupName][matos];
        if (
          typeof val !== "object" ||
          val === null ||
          !("tarif" in val) ||
          !("indice" in val)
        )
          problemesDetectes.push(
            `Groupe ${groupName} / ${matos} : structure invalide`,
          );
      }
    }
  } catch (e) {
    problemesDetectes.push("transport_config_group corrompu");
  }
  try {
    const associes = JSON.parse(
      localStorage.getItem("transport_config_associe") || "{}",
    );
    for (let associeName in associes) {
      if (associeName === pilierCommunKey) continue;
      for (let matos in associes[associeName]) {
        const val = associes[associeName][matos];
        if (
          typeof val !== "object" ||
          val === null ||
          !("tarif" in val) ||
          !("indice" in val)
        )
          problemesDetectes.push(
            `Associé ${associeName} / ${matos} : structure invalide`,
          );
      }
    }
  } catch (e) {
    problemesDetectes.push("transport_config_associe corrompu");
  }
  try {
    for (let i = 0; i < localStorage.length; i++) {
      const key = localStorage.key(i);
      if (key.startsWith("planning_")) {
        const planning = JSON.parse(localStorage.getItem(key));
        planning.forEach((row, index) => {
          if (row.snapshot && typeof row.snapshot === "string")
            problemesDetectes.push(
              `${key} ligne ${index} : snapshot en string`,
            );
        });
      }
    }
  } catch (e) {
    problemesDetectes.push("Plannings corrompus");
  }
  return problemesDetectes;
}

function reparerDonneesAutomatiquement() {
  let reparations = 0;
  try {
    const groups = JSON.parse(
      localStorage.getItem("transport_config_group") || "{}",
    );
    let modifie = false;
    for (let groupName in groups) {
      for (let matos in groups[groupName]) {
        const val = groups[groupName][matos];
        if (
          typeof val !== "object" ||
          val === null ||
          !("tarif" in val) ||
          !("indice" in val)
        ) {
          const oldValue = typeof val === "number" ? val : 0;
          groups[groupName][matos] = { tarif: oldValue, indice: 0 };
          modifie = true;
          reparations++;
        }
      }
    }
    if (modifie)
      localStorage.setItem("transport_config_group", JSON.stringify(groups));
  } catch (e) {}
  try {
    const associes = JSON.parse(
      localStorage.getItem("transport_config_associe") || "{}",
    );
    let modifie = false;
    for (let associeName in associes) {
      if (associeName === pilierCommunKey) continue;
      for (let matos in associes[associeName]) {
        const val = associes[associeName][matos];
        if (
          typeof val !== "object" ||
          val === null ||
          !("tarif" in val) ||
          !("indice" in val)
        ) {
          const oldValue = typeof val === "number" ? val : 0;
          associes[associeName][matos] = { tarif: oldValue, indice: 0 };
          modifie = true;
          reparations++;
        }
      }
    }
    if (modifie)
      localStorage.setItem(
        "transport_config_associe",
        JSON.stringify(associes),
      );
  } catch (e) {}
  try {
    for (let i = 0; i < localStorage.length; i++) {
      const key = localStorage.key(i);
      if (key.startsWith("planning_")) {
        const planning = JSON.parse(localStorage.getItem(key));
        let modifie = false;
        planning.forEach((row) => {
          if ((row.client && row.vehicule) || (row.associe && row.vehicule)) {
            const newSnapshot = captureSnapshot(
              row.client,
              row.vehicule,
              row.associe,
            );
            if (JSON.stringify(row.snapshot) !== JSON.stringify(newSnapshot)) {
              row.snapshot = newSnapshot;
              modifie = true;
              reparations++;
            }
          }
        });
        if (modifie) localStorage.setItem(key, JSON.stringify(planning));
      }
    }
  } catch (e) {}
  return reparations;
}

function verifierEtReparerSiNecessaire() {
  const problemes = verifierIntegriteDonnees();
  if (problemes.length > 0) {
    const derniereReparation = localStorage.getItem("derniere_reparation");
    const maintenant = Date.now();
    if (
      !derniereReparation ||
      maintenant - parseInt(derniereReparation) > 86400000
    ) {
      const reparations = reparerDonneesAutomatiquement();
      if (reparations > 0) {
        localStorage.setItem("derniere_reparation", maintenant.toString());
        setTimeout(() => {
          const notif = document.getElementById("save-notif2");
          if (notif) {
            notif.innerText = `🔧 Données réparées automatiquement\n${reparations} problème(s) corrigé(s)`;
            notif.style.backgroundColor = "#2196F3";
            notif.classList.add("notif-visible");
            setTimeout(() => {
              notif.classList.remove("notif-visible");
            }, 3000);
          }
        }, 1000);
      }
    }
  }
}

function reparerDonneesManuellement() {
  if (
    confirm(
      "Réparer automatiquement les données corrompues ?\n\n(Vos données seront préservées)",
    )
  ) {
    const problemes = verifierIntegriteDonnees();
    if (problemes.length === 0) {
      alert("✅ Aucun problème détecté !");
      return;
    }
    const reparations = reparerDonneesAutomatiquement();
    alert(`✅ Réparation terminée !\n\n${reparations} problème(s) corrigé(s).`);
    location.reload();
  }
}

function resetTout() {
  const premiereConfirmation = confirm(
    "ATTENTION : Vous allez vider TOUTE la mémoire du site. Continuer ?",
  );
  if (premiereConfirmation) {
    const codeSecret = "9051";
    const saisieUtilisateur = prompt("Entrez le mot de passe pour confirmer :");
    const notif = document.getElementById("save-notif2");
    if (saisieUtilisateur === codeSecret) {
      localStorage.clear();
      location.reload();
    } else if (saisieUtilisateur !== null) {
      notif.innerText = `⚠️ CODE INCORRECT ⚠️\nVOUS N'ÊTES PAS AUTORISÉ À EFFECTUER CETTE ACTION !`;
      notif.classList.add("alerte-flash", "notif-visible");
      setTimeout(() => {
        notif.classList.remove("notif-visible");
        setTimeout(() => {
          notif.classList.remove("alerte-flash");
          notif.textContent = "OK";
        }, 500);
      }, 5500);
    }
  }
}

function sauvegarderConfigComplete() {
  const backup = {};
  for (let i = 0; i < localStorage.length; i++) {
    const cle = localStorage.key(i);
    backup[cle] = localStorage.getItem(cle);
  }
  const dataStr =
    "data:text/json;charset=utf-8," +
    encodeURIComponent(JSON.stringify(backup));
  const downloadAnchorNode = document.createElement("a");
  const date = new Date().toLocaleDateString("fr-FR").replace(/\//g, "-");
  downloadAnchorNode.setAttribute("href", dataStr);
  downloadAnchorNode.setAttribute(
    "download",
    "SAUVEGARDE_TOTALE_" + date + ".json",
  );
  document.body.appendChild(downloadAnchorNode);
  downloadAnchorNode.click();
  downloadAnchorNode.remove();
}

function restaurerConfigComplete(input) {
  const file = input.files[0];
  if (!file) return;
  const reader = new FileReader();
  reader.onload = function (e) {
    try {
      const backup = JSON.parse(e.target.result);
      if (
        confirm(
          "Voulez-vous restaurer l'intégralité des données (Planning + Tarifs) ?",
        )
      ) {
        localStorage.clear();
        Object.keys(backup).forEach((cle) => {
          localStorage.setItem(cle, backup[cle]);
        });
        alert("✅ Tout a été restauré avec succès !");
        window.location.reload();
      }
    } catch (err) {
      alert("❌ Erreur de fichier.");
    }
  };
  reader.readAsText(file);
  input.value = "";
}

// ====================================================================
// 🚀  16. INITIALISATION — DOMContentLoaded
// ====================================================================

document.addEventListener("DOMContentLoaded", function () {
  verifierEtReparerSiNecessaire();
  ensurePiliersCommunsExists();

  const datePicker = document.getElementById("current-date-picker");
  const displayDate = document.getElementById("display-date");
  const now = new Date();
  const aujourdhui = `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, "0")}-${String(now.getDate()).padStart(2, "0")}`;
  datePicker.value = aujourdhui;
  displayDate.textContent =
    "Planning du " + new Date(aujourdhui).toLocaleDateString("fr-FR");
  updateMatosDatalist();
  trierDatalist("list-associe");
  trierDatalist("list-chauffeur");
  renderConfigUI();
  renderConfigUI2();
  renderConfigUIChauffeur();
  renderConfigTelephoneUI();
  const saved = localStorage.getItem("planning_" + aujourdhui);
  if (saved && JSON.parse(saved).length > 0) {
    JSON.parse(saved).forEach(ajouterLigne);
    setTimeout(() => {
      forceRecalculTousLesTarifs();
      isInitialLoading = false;
      lastCapturedData = {};
    }, 100);
  } else {
    ajouterLigne();
  }
  setTimeout(updateUndoRedoButtons, 200);

  datePicker.addEventListener("change", function () {
    const nouvelleDate = this.value;
    displayDate.textContent =
      "Planning du " + new Date(nouvelleDate).toLocaleDateString("fr-FR");
    document.getElementById("table-body").innerHTML = "";
    const savedDay = localStorage.getItem("planning_" + nouvelleDate);
    if (savedDay && JSON.parse(savedDay).length > 0)
      JSON.parse(savedDay).forEach(ajouterLigne);
    else ajouterLigne();
    calculerTousLesTotaux();
    updateUndoRedoButtons();
    setTimeout(() => {
      chargerCouleursDate(nouvelleDate);
    }, 100);
  });

  document.getElementById("copy-day").addEventListener("click", copierJournee);
  document.getElementById("paste-day").addEventListener("click", collerJournee);
  document
    .getElementById("toggle-config-client")
    .addEventListener("click", () => {
      document.getElementById("config-section-client").style.display = "block";
      document.body.style.overflow = "hidden";
    });
  document
    .getElementById("toggle-config-associe")
    .addEventListener("click", () => {
      document.getElementById("config-section-associe").style.display = "block";
      document.body.style.overflow = "hidden";
    });
  document
    .getElementById("toggle-config-chauffeur")
    .addEventListener("click", () => {
      document.getElementById("config-section-chauffeur").style.display =
        "block";
      document.body.style.overflow = "hidden";
    });
  document
    .getElementById("toggle-config-telephone")
    .addEventListener("click", () => {
      document.getElementById("config-section-telephone").style.display =
        "block";
      document.body.style.overflow = "hidden";
    });
  document.getElementById("add-row").addEventListener("click", () => {
    ajouterLigne();
    saveCurrentDay();
  });
  document.getElementById("reset-day").addEventListener("click", () => {
    if (confirm("Supprimer toutes les lignes et recommencer ?")) {
      localStorage.removeItem("planning_" + date);
      const allColors = JSON.parse(
        localStorage.getItem("planningColors") || "{}",
      );
      delete allColors[date];
      localStorage.setItem("planningColors", JSON.stringify(allColors));
      document.getElementById("table-body").innerHTML = "";
      ajouterLigne();
      afficherNotification("Planning réinitialisé", "#dc3545");
    }
  });

  document.getElementById("table-body").addEventListener("input", (e) => {
    const row = e.target.closest("tr");
    if (!row) return;
    if (e.target.classList.contains("combo-input")) {
      if (e.target.classList.contains("select-associe"))
        appliquerCouleurAssocie(e.target);
      updateRowTarif(row);
    }
    if (e.target.isContentEditable) {
      if (e.target.classList.contains("cell-bon")) updateRowHighlight(row);
      saveCurrentDay();
    }
  });

  document.getElementById("table-body").addEventListener("change", (e) => {
    if (e.target.classList.contains("select-chauffeur"))
      verifierDoublon(e.target);
  });

  document.getElementById("table-body").addEventListener(
    "blur",
    function (e) {
      const row = e.target.closest("tr");
      if (!row) return;
      if (
        e.target.classList.contains("cell-tarif") ||
        e.target.classList.contains("cell-tarif-affrete")
      ) {
        updateRowTarif(row);
        saveCurrentDay();
        return;
      }
      if (
        !e.target.classList.contains("cell-supplement-client") &&
        !e.target.classList.contains("cell-supplement-affrete")
      )
        return;
      let value = e.target.textContent.replace(/[€\s]/g, "").replace(",", ".");
      const number = parseFloat(value);
      if (!isNaN(number)) e.target.textContent = number.toFixed(2) + " €";
      else e.target.textContent = "0.00 €";
      calculerTotalSupplement_client();
      calculerTotalSupplement_affrete();
      saveCurrentDay();
    },
    true,
  );

  document
    .getElementById("copy-image-client")
    .addEventListener("click", function () {
      const table = document.getElementById("my-table");
      const area = document.getElementById("capture-area");
      const tableContainer = document.querySelector(".table-container");
      const marquerBlock =
        document.querySelector(".marquer-container") ||
        document.getElementById("color-picker-tools");
      const endCol = 12;
      const allRows = table.querySelectorAll("tr");
      const originalTableStyleWidth = table.style.width;
      const originalContainerMaxHeight = tableContainer.style.maxHeight;
      const originalContainerOverflow = tableContainer.style.overflow;
      area.style.width = "1800px";
      area.style.minWidth = "1800px";
      table.style.width = "100%";
      tableContainer.style.maxHeight = "none";
      tableContainer.style.overflow = "visible";
      if (marquerBlock) marquerBlock.style.visibility = "hidden";
      const placeholders = area.querySelectorAll(".placeholder-style");
      placeholders.forEach((el) => {
        el.style.setProperty("color", "transparent", "important");
        el.style.setProperty(
          "-webkit-text-fill-color",
          "transparent",
          "important",
        );
      });
      const selects = area.querySelectorAll("select");
      const hiddenSelects = [];
      selects.forEach((sel) => {
        if (!sel.value || sel.value === "" || sel.value.includes("--")) {
          sel.style.opacity = "0";
          hiddenSelects.push(sel);
        }
      });
      allRows.forEach((row) => {
        Array.from(row.cells).forEach((cell, index) => {
          if (index > endCol) cell.style.display = "none";
        });
      });
      html2canvas(area, {
        backgroundColor: "#ffffff",
        scale: 2,
        useCORS: true,
        width: 1800,
        windowWidth: 1800,
      }).then((canvas) => {
        canvas.toBlob((blob) => {
          if (blob) {
            const item = new ClipboardItem({ "image/png": blob });
            navigator.clipboard.write([item]).then(() => {
              afficherNotification("Planning copié (Vue Client)", "#2196F3");
            });
          }
        });
        area.style.width = "";
        area.style.minWidth = "";
        table.style.width = originalTableStyleWidth;
        tableContainer.style.maxHeight = originalContainerMaxHeight;
        tableContainer.style.overflow = originalContainerOverflow;
        if (marquerBlock) marquerBlock.style.visibility = "visible";
        placeholders.forEach((el) => {
          el.style.color = "";
          el.style.webkitTextFillColor = "";
        });
        hiddenSelects.forEach((sel) => (sel.style.opacity = ""));
        allRows.forEach((row) => {
          Array.from(row.cells).forEach((cell) => {
            cell.style.display = "";
          });
        });
      });
    });

  document.getElementById("copy-image").addEventListener("click", function () {
    const table = document.getElementById("my-table");
    const area = document.getElementById("capture-area");
    const tableContainer = document.querySelector(".table-container");
    const allRows = table.querySelectorAll("tr");
    const endCol = 9;
    const marquerBlock =
      document.querySelector(".marquer-container") ||
      document.getElementById("color-picker-tools");
    const originalTableStyleWidth = table.style.width;
    const forcedWidth = 1800;
    area.style.width = forcedWidth + "px";
    area.style.minWidth = forcedWidth + "px";
    table.style.width = "100%";
    tableContainer.style.maxHeight = "none";
    tableContainer.style.overflow = "visible";
    if (marquerBlock) marquerBlock.style.visibility = "hidden";
    const placeholders = area.querySelectorAll(".placeholder-style");
    placeholders.forEach((el) => {
      el.style.setProperty("color", "transparent", "important");
      el.style.setProperty(
        "-webkit-text-fill-color",
        "transparent",
        "important",
      );
    });
    const selects = area.querySelectorAll("select");
    const hiddenSelects = [];
    selects.forEach((sel) => {
      if (!sel.value || sel.value === "" || sel.value.includes("--")) {
        sel.style.opacity = "0";
        hiddenSelects.push(sel);
      }
    });
    allRows.forEach((row) => {
      Array.from(row.cells).forEach((cell, index) => {
        if (index > endCol) cell.style.display = "none";
      });
    });
    html2canvas(area, {
      backgroundColor: "#ffffff",
      scale: 2,
      width: forcedWidth,
      windowWidth: forcedWidth,
    }).then((canvas) => {
      canvas.toBlob((blob) => {
        const item = new ClipboardItem({ "image/png": blob });
        navigator.clipboard.write([item]).then(() => {
          afficherNotification("Planning copié", "#2196F3");
        });
      });
      area.style.width = "";
      area.style.minWidth = "";
      table.style.width = originalTableStyleWidth;
      tableContainer.style.maxHeight = "";
      tableContainer.style.overflow = "";
      if (marquerBlock) marquerBlock.style.visibility = "visible";
      placeholders.forEach((el) => {
        el.style.color = "";
        el.style.webkitTextFillColor = "";
      });
      hiddenSelects.forEach((sel) => (sel.style.opacity = ""));
      allRows.forEach((row) => {
        Array.from(row.cells).forEach((cell) => {
          cell.style.display = "";
        });
      });
    });
  });

  document
    .getElementById("export-excel")
    .addEventListener("click", function () {
      if (!date) {
        alert(
          "Veuillez sélectionner une date pour définir le mois à exporter.",
        );
        return;
      }

      const clean = (val, def) =>
        !val || val === def || val.trim() === "" ? "" : val;

      const borderThin = { style: "thin", color: { rgb: "CCCCCC" } };
      const allBorders = {
        top: borderThin,
        bottom: borderThin,
        left: borderThin,
        right: borderThin,
      };

      const styleHeader = {
        font: { bold: true, color: { rgb: "FFFFFF" }, sz: 16 },
        fill: { fgColor: { rgb: "1A56A0" } },
        alignment: {
          horizontal: "center",
          vertical: "center",
          wrapText: true,
        },
        border: allBorders,
      };
      const styleDataText = {
        font: { sz: 16 },
        alignment: { vertical: "center", wrapText: false },
        border: allBorders,
      };
      const styleDataMoney = {
        font: { sz: 16 },
        numFmt: '#,##0.00 "€"',
        alignment: { horizontal: "right", vertical: "center" },
        border: allBorders,
        fill: { fgColor: { rgb: "F0F7FF" } },
      };
      const styleDataMoneyAff = {
        ...styleDataMoney,
        fill: { fgColor: { rgb: "FFF0F0" } },
      };
      const styleTotalLabel = {
        font: { bold: true, sz: 16, color: { rgb: "FFFFFF" } },
        fill: { fgColor: { rgb: "2E7D32" } },
        alignment: { horizontal: "right", vertical: "center" },
        border: allBorders,
      };
      const styleTotalValue = {
        font: { bold: true, sz: 16, color: { rgb: "FFFFFF" } },
        numFmt: '#,##0.00 "€"',
        fill: { fgColor: { rgb: "2E7D32" } },
        alignment: { horizontal: "right", vertical: "center" },
        border: allBorders,
      };
      const styleClientHeader = {
        font: { bold: true, sz: 16, color: { rgb: "FFFFFF" } },
        fill: { fgColor: { rgb: "6A1B9A" } },
        alignment: { horizontal: "center", vertical: "center" },
        border: allBorders,
      };
      const styleClientName = {
        font: { bold: true, sz: 16, color: { rgb: "4A148C" } },
        fill: { fgColor: { rgb: "EDE7F6" } },
        alignment: { vertical: "center" },
        border: allBorders,
      };
      const styleClientMoney = {
        font: { sz: 16 },
        numFmt: '#,##0.00 "€"',
        fill: { fgColor: { rgb: "F3E5F5" } },
        alignment: { horizontal: "right", vertical: "center" },
        border: allBorders,
      };
      const styleClientTotal = {
        font: { bold: true, sz: 16, color: { rgb: "FFFFFF" } },
        fill: { fgColor: { rgb: "AB47BC" } },
        alignment: { horizontal: "center", vertical: "center" },
        border: allBorders,
      };
      const styleEmpty = { border: allBorders };

      const monthPrefix = "planning_" + datePicker.substring(0, 7);
      const keys = [];
      for (let i = 0; i < localStorage.length; i++) {
        const key = localStorage.key(i);
        if (key && key.startsWith(monthPrefix)) keys.push(key);
      }
      keys.sort();

      const totauxParClient = {};
      let totalClient = 0,
        totalSupplement_client = 0,
        IC_client = 0;
      let totalAffrete = 0,
        totalSupplement_affrete = 0,
        IC_affrete = 0;

      const ws = {};
      const NCOLS = 18;
      let R = 0;

      function setCell(c, v, s, numFmt) {
        const ref = XLSX.utils.encode_cell({ r: R, c });
        const cell = { v, s };
        if (typeof v === "number") cell.t = "n";
        else cell.t = "s";
        if (numFmt) cell.z = numFmt;
        ws[ref] = cell;
      }
      function setRange(range) {
        ws["!ref"] = range;
      }

      const headers = [
        "Associé",
        "Chantier",
        "Date",
        "Véhicule",
        "Client",
        "Téléphone",
        "Consigne",
        "Chauffeur",
        "N° Chantier",
        "N° Commande",
        "N° Bon",
        "Tarif Client",
        "Suppl. Client",
        "IC Client",
        "Tarif Affrété",
        "Suppl. Affrété",
        "IC Affrété",
        "Commentaire",
      ];
      headers.forEach((h, c) => setCell(c, h, styleHeader));
      R++;

      let dataRowCount = 0;
      keys.forEach((key) => {
        const dayData = JSON.parse(localStorage.getItem(key) || "[]");
        dayData.forEach((row) => {
          const aUnAssocie = row.associe && row.associe !== "Associé...";
          const aUnChantier =
            row.chantier &&
            row.chantier !== "Chantier..." &&
            row.chantier !== "";
          const aUnVehicule = row.vehicule && row.vehicule !== "Véhicule...";
          const aUnClient = row.client && row.client !== "Client...";
          const aUnTelephone = row.telephone && row.telephone !== "Contact...";
          const aUnChauffeur =
            row.chauffeur && row.chauffeur !== "Chauffeur...";
          const aUnChantierNum =
            row.n_chantier &&
            row.n_chantier !== "N° Chantier..." &&
            row.n_chantier !== "";
          const aUnCommande =
            row.n_commande &&
            row.n_commande !== "N° Commande..." &&
            row.n_commande !== "";
          const aUnBon =
            row.n_bon && row.n_bon !== "N° Bon..." && row.n_bon !== "";
          const aUnCommentaire =
            row.commentaire &&
            row.commentaire !== "Commentaire..." &&
            row.commentaire !== "";
          if (
            !aUnAssocie &&
            !aUnChantier &&
            !aUnVehicule &&
            !aUnClient &&
            !aUnTelephone &&
            !aUnChauffeur &&
            !aUnChantierNum &&
            !aUnCommande &&
            !aUnBon &&
            !aUnCommentaire
          )
            return;

          const tarif =
            parseFloat(
              (row.tarif || "0").replace(/[€\s]/g, "").replace(",", "."),
            ) || 0;
          const suppC =
            parseFloat(
              (row.supplement_client || "0")
                .replace(/[€\s]/g, "")
                .replace(",", "."),
            ) || 0;
          const icC =
            parseFloat(
              (row.IC_client || "0").replace(/[€\s]/g, "").replace(",", "."),
            ) || 0;
          const tarifAff =
            parseFloat(
              (row.tarif_affrete || "0")
                .replace(/[€\s]/g, "")
                .replace(",", "."),
            ) || 0;
          const suppAff =
            parseFloat(
              (row.supplement_affrete || "0")
                .replace(/[€\s]/g, "")
                .replace(",", "."),
            ) || 0;
          const icAff =
            parseFloat(
              (row.IC_affrete || "0").replace(/[€\s]/g, "").replace(",", "."),
            ) || 0;

          totalClient += tarif;
          totalSupplement_client += suppC;
          IC_client += icC;
          totalAffrete += tarifAff;
          totalSupplement_affrete += suppAff;
          IC_affrete += icAff;

          const nomClient = clean(row.client, "Client...") || "Sans client";
          if (!totauxParClient[nomClient])
            totauxParClient[nomClient] = {
              tarif: 0,
              supp: 0,
              ic: 0,
              tarifAff: 0,
              suppAff: 0,
              icAff: 0,
            };
          totauxParClient[nomClient].tarif += tarif;
          totauxParClient[nomClient].supp += suppC;
          totauxParClient[nomClient].ic += icC;
          totauxParClient[nomClient].tarifAff += tarifAff;
          totauxParClient[nomClient].suppAff += suppAff;
          totauxParClient[nomClient].icAff += icAff;

          const assVal = (clean(row.associe, "Associé...") || "").toUpperCase();
          let rowFill = "FFFFFF";
          if (assVal === "ALLARD") rowFill = "FFFDE7";
          else if (assVal === "LANDAIS") rowFill = "E0FEFF";
          else if (assVal === "STD") rowFill = "F8E8FF";
          const styleRowText = {
            ...styleDataText,
            fill: { fgColor: { rgb: rowFill } },
          };
          const styleRowMoney = {
            ...styleDataMoney,
            fill: { fgColor: { rgb: "D0E8FF" } },
          };
          const styleRowMoneyAff = {
            ...styleDataMoneyAff,
            fill: { fgColor: { rgb: "FFE0E0" } },
          };

          const textCols = [
            clean(row.associe, "Associé..."),
            clean(row.chantier, "Chantier..."),
            row.date || "",
            clean(row.vehicule, "Véhicule..."),
            clean(row.client, "Client..."),
            clean(row.telephone, "Contact..."),
            clean(row.consigne, "Consigne..."),
            clean(row.chauffeur, "Chauffeur..."),
            clean(row.n_chantier, "N° Chantier..."),
            clean(row.n_commande, "N° Commande..."),
            clean(row.n_bon, "N° Bon..."),
          ];
          textCols.forEach((v, c) => setCell(c, v, styleRowText));
          [tarif, suppC, icC].forEach((v, i) =>
            setCell(11 + i, v, styleRowMoney, '#,##0.00 "€"'),
          );
          [tarifAff, suppAff, icAff].forEach((v, i) =>
            setCell(14 + i, v, styleRowMoneyAff, '#,##0.00 "€"'),
          );
          setCell(17, clean(row.commentaire, "Commentaire..."), styleRowText);
          R++;
          dataRowCount++;
        });
      });

      if (dataRowCount === 0) {
        alert("Aucune donnée enregistrée pour ce mois.");
        return;
      }

      ws["!autofilter"] = {
        ref: XLSX.utils.encode_range({
          s: { r: 0, c: 0 },
          e: { r: 0, c: NCOLS - 1 },
        }),
      };

      const lastDataRowExcel = R;
      R++;

      setCell(10, "TOTAL", styleTotalLabel);
      const colLetters = ["L", "M", "N", "O", "P", "Q"];
      colLetters.forEach((col, i) => {
        const ref = XLSX.utils.encode_cell({ r: R, c: 11 + i });
        ws[ref] = {
          f: `SUBTOTAL(9,${col}2:${col}${lastDataRowExcel})`,
          t: "n",
          z: '#,##0.00 "€"',
          s: styleTotalValue,
        };
      });
      R++;

      ws["!ref"] = XLSX.utils.encode_range({
        s: { r: 0, c: 0 },
        e: { r: R - 1, c: NCOLS - 1 },
      });

      ws["!cols"] = [
        { wch: 18 }, // Associé
        { wch: 26 }, // Chantier
        { wch: 15 }, // Date
        { wch: 15 }, // Véhicule
        { wch: 26 }, // Client
        { wch: 23 }, // Téléphone
        { wch: 38 }, // Consigne
        { wch: 22 }, // Chauffeur
        { wch: 26 }, // N° Chantier
        { wch: 26 }, // N° Commande
        { wch: 18 }, // N° Bon
        { wch: 21 }, // Tarif Client
        { wch: 21 }, // Suppl. Client
        { wch: 19 }, // IC Client
        { wch: 21 }, // Tarif Affrété
        { wch: 21 }, // Suppl. Affrété
        { wch: 19 }, // IC Affrété
        { wch: 28 }, // Commentaire
      ];

      for (i = 0; i < R; i++) {
        ws["!rows"] = [...(ws["!rows"] || []), { hpt: 50, hpx: 50 }];
      }

      ws["!freeze"] = {
        xSplit: 0,
        ySplit: 1,
        topLeftCell: "A2",
        activePane: "bottomLeft",
      };

      const workbook = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(workbook, ws, "Export Mensuel");
      XLSX.writeFile(workbook, `Transport_${datePicker.substring(0, 7)}.xlsx`);
    });
});

window.addEventListener("load", () => {
  const currentDate = document.getElementById("current-date-picker").value;
  chargerCouleursDate(currentDate);
});
