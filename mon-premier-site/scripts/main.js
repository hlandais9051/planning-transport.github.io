const lesElements = document.querySelectorAll("li");

function changerValidation(e) {
  e.currentTarget.classList.toggle("fini");
}

lesElements.forEach((element) => {
  element.addEventListener("click", changerValidation);
});
/*// Stocke une référence vers le <h1> dans une variable
const monTitre = document.querySelector("h1");
// Met à jour le contenu texte du <h1>
monTitre.textContent = "Hello world !";*/
const monImage = document.querySelector("img");

monImage.addEventListener("click", () => {
  const maSrc = monImage.getAttribute("src");
  if (maSrc === "images/krokmou-da.png") {
    monImage.setAttribute("src", "images/krokmou-film.jpg");
  } else {
    monImage.setAttribute("src", "images/krokmou-da.png");
  }
});
let monBouton = document.querySelector("button");
let monTitre = document.querySelector("h1");
function definirNomUtilisateur() {
  const monNom = prompt("Veuillez saisir votre nom.");
  if (!monNom) {
    // Si l'utilisateur n'a rien saisi, on ne change rien
    return;
  }
  localStorage.setItem("nom", monNom);
  monTitre.textContent = `Mozilla est génial, ${monNom}`;
}
if (!localStorage.getItem("nom")) {
  definirNomUtilisateur();
} else {
  const nomEnregistre = localStorage.getItem("nom");
  monTitre.textContent = `Mozilla est génial, ${nomEnregistre}`;
}
monBouton.addEventListener("click", () => {
  definirNomUtilisateur();
});

// 1. Tes données regroupées
const transportLandais = {
  eurovia: {
    grue: { prix: 500, idx: 0.0085 },
    semi: { prix: 200, idx: 0.005 },
    remorque: { prix: 100, idx: 0.002 },
  },
  colas: {
    grue: { prix: 450, idx: 0.009 },
    semi: { prix: 500, idx: 0.01 },
    remorque: { prix: 90, idx: 0.004 },
  },
  fayat: {
    grue: { prix: 480, idx: 0.0075 },
    semi: { prix: 190, idx: 0.006 },
    remorque: { prix: 95, idx: 0.003 },
  },
};

const selectEntreprise = document.getElementById("select-entreprise");
const selectVehicule = document.getElementById("select-vehicule");
const affichagePrix = document.getElementById("resultat-prix");

function mettreAJourPrix() {
  const nomEntreprise = selectEntreprise.value;
  const nomVehicule = selectVehicule.value;

  // On accède à l'objet précis : transportLandais["eurovia"]["grue"]
  const donneesVehicule = transportLandais[nomEntreprise][nomVehicule];

  const prixBase = donneesVehicule.prix;
  const indexCarb = donneesVehicule.idx;

  const prixFinal = prixBase + prixBase * indexCarb;

  affichagePrix.textContent = prixFinal.toFixed(2);
}

// On écoute les changements sur les DEUX listes
selectEntreprise.addEventListener("change", mettreAJourPrix);
selectVehicule.addEventListener("change", mettreAJourPrix);

// Initialisation au chargement
mettreAJourPrix();
