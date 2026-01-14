document.addEventListener("DOMContentLoaded", function () {
  const tableBody = document.getElementById("table-body");
  const inputIndice = document.getElementById("indice-carburant");

  document.addEventListener("copy", function (e) {
    const selection = window.getSelection();
    if (selection.rangeCount > 0) {
      const rows = document.querySelectorAll("#table-body tr");
      let finalOutput = [];

      rows.forEach((row) => {
        let rowData = [];
        const cells = row.querySelectorAll("td");

        cells.forEach((cell) => {
          // ON VÉRIFIE SI CETTE CELLULE PRÉCISE EST SÉLECTIONNÉE
          if (selection.containsNode(cell, true)) {
            const select = cell.querySelector("select");
            if (select) {
              // On prend le texte du menu, mais on ignore "Choisir"
              const text = select.options[select.selectedIndex].text;
              if (text !== "Choisir") {
                rowData.push(text);
              }
            } else {
              // On prend le texte de la cellule, mais on ignore si c'est vide
              const text = cell.innerText.trim();
              if (text !== "") {
                rowData.push(text);
              }
            }
          }
        });

        // On assemble les données de la ligne avec un espace
        if (rowData.length > 0) {
          finalOutput.push(rowData.join(" "));
        }
      });

      // On assemble toutes les lignes avec un retour à la ligne
      const resultString = finalOutput.join("\n");

      if (resultString) {
        e.clipboardData.setData("text/plain", resultString);
        e.preventDefault();
      }
    }
  });

  // --- 2. GRILLE TARIFAIRE ---
  const transportLandais = {
    EUROVIA: {
      "6x2 GRUE": { prix: 500, idx: 0.0085 },
      "8x4": { prix: 200, idx: 0.005 },
      AC3: { prix: 100, idx: 0.002 },
      PB: { prix: 150, idx: 0.004 },
      PC: { prix: 180, idx: 0.006 },
      "6x4": { prix: 220, idx: 0.007 },
      "8x2 AMP GRUE": { prix: 300, idx: 0.009 },
    },
    COCA: {
      "6x2 GRUE": { prix: 450, idx: 0.009 },
      "8x4": { prix: 500, idx: 0.01 },
      AC3: { prix: 90, idx: 0.004 },
      PB: { prix: 140, idx: 0.005 },
      PC: { prix: 170, idx: 0.007 },
      "6x4": { prix: 210, idx: 0.008 },
      "8x2 AMP GRUE": { prix: 290, idx: 0.01 },
    },
    EHTP: {
      "6x2 GRUE": { prix: 480, idx: 0.0075 },
      "8x4": { prix: 190, idx: 0.006 },
      AC3: { prix: 95, idx: 0.003 },
      PB: { prix: 155, idx: 0.0055 },
      PC: { prix: 185, idx: 0.0065 },
      "6x4": { prix: 225, idx: 0.0075 },
      "8x2 AMP GRUE": { prix: 310, idx: 0.0085 },
    },
    "NGE ENERGIES SOLUTION": {
      "6x2 GRUE": { prix: 470, idx: 0.008 },
      "8x4": { prix: 210, idx: 0.0055 },
      AC3: { prix: 85, idx: 0.0045 },
      PB: { prix: 145, idx: 0.0052 },
      PC: { prix: 175, idx: 0.0062 },
      "6x4": { prix: 215, idx: 0.0072 },
      "8x2 AMP GRUE": { prix: 295, idx: 0.0092 },
    },
    KERLEROUX: {
      "6x2 GRUE": { prix: 460, idx: 0.0095 },
      "8x4": { prix: 220, idx: 0.0065 },
      AC3: { prix: 80, idx: 0.0055 },
      PB: { prix: 135, idx: 0.0062 },
      PC: { prix: 165, idx: 0.0072 },
      "6x4": { prix: 205, idx: 0.0082 },
      "8x2 AMP GRUE": { prix: 285, idx: 0.01 },
    },
    "PIGEON TP": {
      "6x2 GRUE": { prix: 490, idx: 0.0082 },
      "8x4": { prix: 205, idx: 0.0058 },
      AC3: { prix: 88, idx: 0.0042 },
      PB: { prix: 148, idx: 0.0053 },
      PC: { prix: 178, idx: 0.0063 },
      "6x4": { prix: 218, idx: 0.0073 },
      "8x2 AMP GRUE": { prix: 308, idx: 0.0093 },
    },
    NGE: {
      "6x2 GRUE": { prix: 475, idx: 0.0078 },
      "8x4": { prix: 215, idx: 0.0062 },
      AC3: { prix: 92, idx: 0.0038 },
      PB: { prix: 142, idx: 0.0051 },
      PC: { prix: 172, idx: 0.0061 },
      "6x4": { prix: 212, idx: 0.0071 },
      "8x2 AMP GRUE": { prix: 292, idx: 0.0091 },
    },
    "CHARIER TP NOZAY": {
      "6x2 GRUE": { prix: 485, idx: 0.0088 },
      "8x4": { prix: 225, idx: 0.0068 },
      AC3: { prix: 87, idx: 0.0048 },
      PB: { prix: 137, idx: 0.0054 },
      PC: { prix: 167, idx: 0.0064 },
      "6x4": { prix: 207, idx: 0.0084 },
      "8x2 AMP GRUE": { prix: 287, idx: 0.01 },
    },
    COLAS: {
      "6x2 GRUE": { prix: 450, idx: 0.009 },
      "8x4": { prix: 500, idx: 0.01 },
      AC3: { prix: 90, idx: 0.004 },
      PB: { prix: 140, idx: 0.005 },
      PC: { prix: 170, idx: 0.007 },
      "6x4": { prix: 210, idx: 0.008 },
      "8x2 AMP GRUE": { prix: 290, idx: 0.01 },
    },
    ATP: {
      "6x2 GRUE": { prix: 460, idx: 0.0092 },
      "8x4": { prix: 510, idx: 0.0102 },
      AC3: { prix: 92, idx: 0.0042 },
      PB: { prix: 142, idx: 0.0052 },
      PC: { prix: 172, idx: 0.0072 },
      "6x4": { prix: 212, idx: 0.0082 },
      "8x2 AMP GRUE": { prix: 292, idx: 0.0102 },
    },
    MODALL: {
      "6x2 GRUE": { prix: 470, idx: 0.0094 },
      "8x4": { prix: 520, idx: 0.0104 },
      AC3: { prix: 94, idx: 0.0044 },
      PB: { prix: 144, idx: 0.0054 },
      PC: { prix: 174, idx: 0.0074 },
      "6x4": { prix: 214, idx: 0.0084 },
      "8x2 AMP GRUE": { prix: 294, idx: 0.0104 },
    },
    "LANDAIS TP": {
      "6x2 GRUE": { prix: 480, idx: 0.0096 },
      "8x4": { prix: 530, idx: 0.0106 },
      AC3: { prix: 96, idx: 0.0046 },
      PB: { prix: 146, idx: 0.0056 },
      PC: { prix: 176, idx: 0.0076 },
      "6x4": { prix: 216, idx: 0.0086 },
      "8x2 AMP GRUE": { prix: 296, idx: 0.0106 },
    },
    EIFFAGE: {
      "6x2 GRUE": { prix: 490, idx: 0.0098 },
      "8x4": { prix: 540, idx: 0.0108 },
      AC3: { prix: 98, idx: 0.0048 },
      PB: { prix: 148, idx: 0.0058 },
      PC: { prix: 178, idx: 0.0078 },
      "6x4": { prix: 218, idx: 0.0088 },
      "8x2 AMP GRUE": { prix: 298, idx: 0.0108 },
    },
    "CHARIER TP SUD": {
      "6x2 GRUE": { prix: 500, idx: 0.01 },
      "8x4": { prix: 550, idx: 0.011 },
      AC3: { prix: 100, idx: 0.005 },
      PB: { prix: 150, idx: 0.006 },
      PC: { prix: 180, idx: 0.008 },
      "6x4": { prix: 220, idx: 0.009 },
      "8x2 AMP GRUE": { prix: 300, idx: 0.011 },
    },
  };

  function mettreAJourTarif(ligne) {
    const clientSelect = ligne.querySelector(".select-client");
    const vehiculeSelect = ligne.querySelector(".select-vehicule");
    const tarifCell = ligne.cells[9];
    if (!clientSelect || !vehiculeSelect || !tarifCell) return;

    const nomClient = clientSelect.value.trim();
    const typeVehicule = vehiculeSelect.value.trim();

    if (
      transportLandais[nomClient] &&
      transportLandais[nomClient][typeVehicule]
    ) {
      const donnees = transportLandais[nomClient][typeVehicule];
      let prixFinal = donnees.prix * (1 + donnees.idx);
      tarifCell.textContent = prixFinal.toFixed(2) + " €";
    } else {
      tarifCell.textContent = "";
    }
  }

  // --- 3. NAVIGATION ET ÉDITION ---
  function activerCellule(cellule) {
    if (
      cellule &&
      cellule.tagName === "TD" &&
      !cellule.classList.contains("date-cell") &&
      !cellule.classList.contains("matos-cell")
    ) {
      cellule.contentEditable = "true";
      cellule.focus();

      cellule.addEventListener(
        "blur",
        function () {
          cellule.contentEditable = "false";
        },
        { once: true }
      );
    }
  }

  tableBody.addEventListener("click", (e) => activerCellule(e.target));

  tableBody.addEventListener("keydown", function (e) {
    if (e.key === "Tab") {
      setTimeout(() => {
        activerCellule(document.activeElement);
      }, 50);
    }
  });

  // --- 4. GESTION DES DATES ET HEURES ---
  function mettreAJourDates() {
    let dateDujour = new Date();
    let demain = new Date(dateDujour);
    demain.setDate(dateDujour.getDate() + 1);

    // Ajout de l'heure dans les options
    const options = {
      day: "2-digit",
      month: "2-digit",
      year: "numeric",
    };
    const dateTexte = demain.toLocaleString("fr-FR", options);

    document.querySelectorAll(".date-cell").forEach((c) => {
      c.textContent = dateTexte;
    });
  }

  // --- 5. INITIALISATION ET ÉVÉNEMENTS ---
  const btnAddRow = document.getElementById("add-row");
  btnAddRow.addEventListener("click", function () {
    const nouvelleLigne = document.createElement("tr");
    nouvelleLigne.innerHTML = `
        <td class="matos-cell">
            <select class="select-associe">
                <option value="">Choisir</option>
                <option value="LANDAIS">LANDAIS</option>
                <option value="ALLARD">ALLARD</option>
                <option value="STD">STD</option>
                <option value="TENDRON">TENDRON</option>
                <option value="GOULARD">GOULARD</option>
                <option value="PINON">PINON</option>
                <option value="MAHE">MAHE</option>
            </select>
        </td>
        <td tabindex="0">0.5</td>
        <td tabindex="0">0.5</td>
        <td tabindex="0"></td>
        <td class="date-cell" style="font-weight: bold; background-color: #eee;"></td>
        <td class="matos-cell"><select class="select-vehicule">
            <option value="">Choisir</option>
            <option value="6x2 GRUE">6x2 GRUE</option>
            <option value="8x4">8x4</option>
            <option value="AC3">AC3</option>
            <option value="PB">PB</option>
            <option value="PC">PC</option>
            <option value="6x4">6x4</option>
            <option value="8x2 AMP GRUE">8x2 AMP GRUE</option>
        </select></td>
        <td class="matos-cell"><select class="select-client">
            <option value="">Choisir</option>
            <option value="EUROVIA">EUROVIA</option>
            <option value="COCA">COCA</option>
            <option value="EHTP">EHTP</option>
            <option value="NGE ENERGIES SOLUTION">NGE ENERGIES SOLUTION</option>
            <option value="KERLEROUX">KERLEROUX</option>
            <option value="PIGEON TP">PIGEON TP</option>
            <option value="NGE">NGE</option>
            <option value="CHARIER TP NOZAY">CHARIER TP NOZAY</option>
            <option value="COLAS">COLAS</option>
            <option value="DALKIA">DALKIA</option>
            <option value="ATP">ATP</option>
            <option value="MODALL">MODALL</option>
            <option value="LANDAIS TP">LANDAIS TP</option>
            <option value="EIFFAGE">EIFFAGE</option>
            <option value="CHARIER TP SUD">CHARIER TP SUD</option>
        </select></td>
        <td tabindex="0"></td>
        <td tabindex="0"></td>
    `;
    tableBody.appendChild(nouvelleLigne);
    mettreAJourDates();
  });

  tableBody.addEventListener("change", function (e) {
    const target = e.target;
    if (
      target.classList.contains("select-vehicule") ||
      target.classList.contains("select-client")
    ) {
      mettreAJourTarif(target.closest("tr"));
    }
    if (target.classList.contains("select-associe")) {
      appliquerCouleurSelect(target);
    }
  });

  function appliquerCouleurSelect(selectElement) {
    const nom = selectElement.value.toUpperCase();
    const cellule = selectElement.parentElement;
    cellule.classList.remove(
      "color-landais",
      "color-allard",
      "color-std",
      "color-tendron",
      "color-goulard",
      "color-pinon",
      "color-mahe"
    );
    if (nom && nom !== "CHOISIR") {
      cellule.classList.add("color-" + nom.toLowerCase());
    }
  }

  // Lancement initial
  mettreAJourDates();
});
