/* ============================================================
   base.js — Planning transport : base de données Supabase
   Remplace le stockage localStorage par une base en ligne,
   sans modifier le fonctionnement de script.js.
   ============================================================ */
(function () {
  "use strict";

  const SUPABASE_URL = "https://kjphjbqjovctdcmbhgtv.supabase.co";
  const SUPABASE_KEY = "sb_publishable_uzXs7ZvAFZ8FTixfDJnqew_YK33CKm2";

  // Clés qui restent locales au navigateur (jamais envoyées en base)
  const LOCAL_ONLY =
    /^(sb-|planning_copy_buffer$|derniere_reparation$|undoStack$|redoStack$|lastCapturedData$|base_)/;
  const CLE_JOUR = /^planning_(\d{4}-\d{2}-\d{2})$/;

  // ---- Accès au vrai localStorage (pour la session de connexion et les clés locales)
  const realLS = window.localStorage;
  const P = Storage.prototype;
  const rawGet = P.getItem,
    rawSet = P.setItem,
    rawRemove = P.removeItem;
  const rawKey = P.key,
    rawClear = P.clear;
  const lengthGetter = Object.getOwnPropertyDescriptor(P, "length").get;
  // Si le localStorage du navigateur est plein (ancien planning), la session de connexion
  // et les petits indicateurs sont gardés dans sessionStorage (valable tant que l'onglet reste ouvert).
  function lireLocal(k) {
    let v = null;
    try {
      v = rawGet.call(realLS, k);
    } catch (e) {}
    if (v === null) {
      try {
        v = window.sessionStorage.getItem(k);
      } catch (e) {}
    }
    return v;
  }
  function ecrireLocal(k, v) {
    try {
      rawSet.call(realLS, k, v);
      return;
    } catch (e) {}
    try {
      window.sessionStorage.setItem(k, v);
    } catch (e) {}
  }
  const adapterSession = {
    getItem: (k) => lireLocal(k),
    setItem: (k, v) => ecrireLocal(k, v),
    removeItem: (k) => {
      try {
        rawRemove.call(realLS, k);
      } catch (e) {}
      try {
        window.sessionStorage.removeItem(k);
      } catch (e) {}
    },
  };

  // ---- Cache en mémoire (remplace le stockage du navigateur)
  const cache = new Map();
  let cleListe = null; // liste des clés, recalculée à la demande
  const dirty = new Set(); // clés à envoyer en base
  let actif = false;
  let sb = null;

  function marquer(k) {
    if (!actif || LOCAL_ONLY.test(k)) return;
    dirty.add(k);
    majStatut("attente");
    clearTimeout(minuteur);
    minuteur = setTimeout(envoyer, 700);
  }

  function installerCache() {
    P.getItem = function (k) {
      if (this !== realLS) return rawGet.call(this, k);
      k = String(k);
      if (LOCAL_ONLY.test(k)) return rawGet.call(this, k);
      return cache.has(k) ? cache.get(k) : null;
    };
    P.setItem = function (k, v) {
      if (this !== realLS) return rawSet.call(this, k, v);
      k = String(k);
      v = String(v);
      if (LOCAL_ONLY.test(k)) return rawSet.call(this, k, v);
      if (!cache.has(k)) cleListe = null;
      cache.set(k, v);
      marquer(k);
    };
    P.removeItem = function (k) {
      if (this !== realLS) return rawRemove.call(this, k);
      k = String(k);
      if (LOCAL_ONLY.test(k)) return rawRemove.call(this, k);
      if (cache.delete(k)) {
        cleListe = null;
        marquer(k);
      }
    };
    P.key = function (i) {
      if (this !== realLS) return rawKey.call(this, i);
      if (!cleListe) cleListe = Array.from(cache.keys());
      return i < cleListe.length ? cleListe[i] : null;
    };
    P.clear = function () {
      if (this !== realLS) return rawClear.call(this);
      Array.from(cache.keys()).forEach((k) => marquer(k));
      cache.clear();
      cleListe = null;
      // vider aussi les clés locales, SAUF la session de connexion
      const locales = [];
      for (let i = 0; i < lengthGetter.call(realLS); i++) {
        const k = rawKey.call(realLS, i);
        if (k && !/^sb-/.test(k)) locales.push(k);
      }
      locales.forEach((k) => rawRemove.call(realLS, k));
    };
    Object.defineProperty(P, "length", {
      configurable: true,
      get: function () {
        if (this !== realLS) return lengthGetter.call(this);
        return cache.size;
      },
    });
  }

  // ---- Conversion texte <-> JSON pour la table config
  function versJson(texte) {
    try {
      const v = JSON.parse(texte);
      if (v !== null) return v;
    } catch (e) {}
    return { __raw: texte };
  }
  function depuisJson(v) {
    if (
      v &&
      typeof v === "object" &&
      !Array.isArray(v) &&
      Object.keys(v).length === 1 &&
      "__raw" in v
    )
      return v.__raw;
    return JSON.stringify(v);
  }

  // ---- Envoi vers la base
  let minuteur = null;
  let enCours = null;
  let erreur = false;

  // Une ligne est "vide" si rien n'a été saisi (textes d'aide, montants à 0, aucune couleur)
  function ligneVide(r) {
    if (!r || typeof r !== "object") return true;
    if (r.couleurs_cellules && Object.keys(r.couleurs_cellules).length)
      return false;
    for (const c of Object.keys(CHAMPS_TEXTE)) {
      const v = r[c];
      if (!(estVide(v) || String(v).trim() === "" || v === CHAMPS_TEXTE[c]))
        return false;
    }
    for (const c of CHAMPS_MONTANT) {
      const v = r[c];
      if (estVide(v)) continue;
      const n = parseFloat(
        String(v)
          .replace(",", ".")
          .replace(/[^0-9.\-]/g, ""),
      );
      if (n) return false;
    }
    return true;
  }

  async function enregistrerJour(date) {
    const texte = cache.get("planning_" + date);
    let lignes = [];
    if (texte) {
      try {
        const p = JSON.parse(texte);
        if (!Array.isArray(p)) return true; // format inattendu : on ne touche pas à la base
        lignes = p;
      } catch (e) {
        return true;
      }
    }
    // Les lignes vides ne sont jamais conservées en base (une journée vide n'y apparaît donc pas)
    lignes = lignes.filter((l) => !ligneVide(l));
    const { error } = await sb.rpc("remplacer_jour", {
      p_date: date,
      p_lignes: lignes,
    });
    if (error) console.error("Erreur enregistrement jour " + date, error);
    return !error;
  }

  function envoyer() {
    clearTimeout(minuteur);
    if (enCours)
      return enCours.then(() =>
        dirty.size && !erreur ? envoyer() : undefined,
      );
    if (!dirty.size) return Promise.resolve();
    enCours = (async () => {
      const cles = Array.from(dirty);
      dirty.clear();
      const echecs = [];
      const ajouts = [],
        retraits = [],
        jours = [];
      cles.forEach((k) => {
        if (CLE_JOUR.test(k)) jours.push(k);
        else if (cache.has(k))
          ajouts.push({
            cle: k,
            valeur: versJson(cache.get(k)),
            maj: new Date().toISOString(),
          });
        else retraits.push(k);
      });
      try {
        if (ajouts.length) {
          const { error } = await sb.from("config").upsert(ajouts);
          if (error) {
            console.error(error);
            ajouts.forEach((r) => echecs.push(r.cle));
          }
        }
        if (retraits.length) {
          const { error } = await sb
            .from("config")
            .delete()
            .in("cle", retraits);
          if (error) {
            console.error(error);
            retraits.forEach((k) => echecs.push(k));
          }
        }
        for (const k of jours) {
          const ok = await enregistrerJour(k.slice(9));
          if (!ok) echecs.push(k);
        }
      } catch (e) {
        console.error(e);
        cles.forEach((k) => {
          if (!echecs.includes(k)) echecs.push(k);
        });
      }
      echecs.forEach((k) => dirty.add(k));
      erreur = echecs.length > 0;
      majStatut(erreur ? "erreur" : dirty.size ? "attente" : "ok");
      if (erreur)
        minuteur = setTimeout(() => {
          erreur = false;
          envoyer();
        }, 5000);
    })().finally(() => {
      enCours = null;
    });
    return enCours.then(() => (dirty.size && !erreur ? envoyer() : undefined));
  }

  // ---- Nettoyage des "undefined" enregistrés par erreur dans d'anciennes données
  const CHAMPS_TEXTE = {
    chantier: "Chantier...",
    consigne: "Consigne...",
    n_chantier: "N° Chantier...",
    n_commande: "N° Commande...",
    n_bon: "N° Bon...",
    heures: "Heures...",
    commentaire: "Commentaire...",
    associe: "",
    vehicule: "",
    client: "",
    chauffeur: "",
    telephone: "",
  };
  const CHAMPS_MONTANT = [
    "tarif",
    "supplement_client",
    "IC_client",
    "tarif_affrete",
    "supplement_affrete",
    "IC_affrete",
  ];
  const estVide = (v) =>
    v === undefined || v === null || String(v).trim() === "undefined";

  // Renvoie le texte corrigé de la journée, ou null si rien à corriger
  function nettoyerJour(date, texte) {
    let lignes;
    try {
      lignes = JSON.parse(texte);
    } catch (e) {
      return null;
    }
    if (!Array.isArray(lignes)) return null;
    let modifie = false;
    lignes.forEach((r) => {
      if (!r || typeof r !== "object") return;
      Object.keys(CHAMPS_TEXTE).forEach((c) => {
        if (estVide(r[c])) {
          r[c] = CHAMPS_TEXTE[c];
          modifie = true;
        }
      });
      CHAMPS_MONTANT.forEach((c) => {
        if (String(r[c]).trim() === "undefined") {
          r[c] = "0.00 €";
          modifie = true;
        }
      });
      if (estVide(r.date)) {
        r.date = new Date(date + "T12:00:00").toLocaleDateString("fr-FR", {
          day: "2-digit",
          month: "2-digit",
          year: "2-digit",
        });
        modifie = true;
      }
    });
    return modifie ? JSON.stringify(lignes) : null;
  }
  const joursCorriges = new Set();

  // ---- Chargement initial depuis la base
  function ligneDepuisColonnes(r) {
    return {
      associe: r.associe,
      chantier: r.chantier,
      date: r.date_ligne,
      vehicule: r.vehicule,
      client: r.client,
      telephone: r.telephone,
      consigne: r.consigne,
      chauffeur: r.chauffeur,
      n_chantier: r.n_chantier,
      n_commande: r.n_commande,
      n_bon: r.n_bon,
      heures: r.heures,
      tarif: r.tarif,
      supplement_client: r.supplement_client,
      IC_client: r.ic_client,
      tarif_affrete: r.tarif_affrete,
      supplement_affrete: r.supplement_affrete,
      IC_affrete: r.ic_affrete,
      commentaire: r.commentaire,
      snapshot: r.snapshot,
      couleurs_cellules: r.couleurs_cellules,
    };
  }

  async function chargerDepuisBase() {
    const cfg = await sb.from("config").select("cle,valeur");
    if (cfg.error) throw cfg.error;
    cfg.data.forEach((r) => cache.set(r.cle, depuisJson(r.valeur)));

    const jours = new Map();
    const pas = 1000;
    for (let debut = 0; ; debut += pas) {
      setOverlayTexte("Chargement du planning… (" + debut + " lignes)");
      const { data, error } = await sb
        .from("lignes_planning")
        .select("*")
        .order("date_jour")
        .order("position")
        .order("id")
        .range(debut, debut + pas - 1);
      if (error) throw error;
      data.forEach((r) => {
        if (!jours.has(r.date_jour)) jours.set(r.date_jour, []);
        jours.get(r.date_jour).push(r.donnees || ligneDepuisColonnes(r));
      });
      if (data.length < pas) break;
    }
    jours.forEach((lignes, date) => {
      const texte = JSON.stringify(lignes);
      const propre = nettoyerJour(date, texte);
      if (propre) joursCorriges.add("planning_" + date);
      cache.set("planning_" + date, propre || texte);
    });
    cleListe = null;
  }

  // ---- Fusion de données (sauvegarde JSON ou anciennes données du navigateur)
  // Règle : rien de ce qui est déjà en base n'est écrasé sans accord explicite.
  function canon(v) {
    if (Array.isArray(v)) return "[" + v.map(canon).join(",") + "]";
    if (v && typeof v === "object")
      return (
        "{" +
        Object.keys(v)
          .sort()
          .map((k) => JSON.stringify(k) + ":" + canon(v[k]))
          .join(",") +
        "}"
      );
    return JSON.stringify(v);
  }
  function lignesNonVides(texte) {
    try {
      const p = JSON.parse(texte);
      if (Array.isArray(p)) return p.filter((r) => !ligneVide(r));
    } catch (e) {}
    return null;
  }
  function canonJour(texte) {
    const l = lignesNonVides(texte);
    return l ? canon(l) : String(texte);
  }
  // Ajoute à "a" ce qui manque depuis "b" ; en cas de désaccord, la valeur de "a" (la base) gagne
  function fusionProfonde(a, b) {
    if (Array.isArray(a) && Array.isArray(b)) {
      const vus = new Set(a.map(canon));
      const res = a.slice();
      b.forEach((x) => {
        const c = canon(x);
        if (!vus.has(c)) {
          vus.add(c);
          res.push(x);
        }
      });
      return res;
    }
    if (
      a &&
      b &&
      typeof a === "object" &&
      typeof b === "object" &&
      !Array.isArray(a) &&
      !Array.isArray(b)
    ) {
      const res = Object.assign({}, a);
      Object.keys(b).forEach((k) => {
        res[k] = k in a ? fusionProfonde(a[k], b[k]) : b[k];
      });
      return res;
    }
    return a;
  }

  // Renvoie "ok", "rien" ou "annule"
  async function fusionner(source, libelle, silencieux) {
    const joursAjoutes = [],
      joursDifferents = [],
      reglagesAjoutes = [];
    const reglagesFusionnes = [];
    Object.keys(source).forEach((k) => {
      if (LOCAL_ONLY.test(k) || typeof source[k] !== "string") return;
      const m = CLE_JOUR.exec(k);
      if (m) {
        const texte = nettoyerJour(m[1], source[k]) || source[k];
        const lignes = lignesNonVides(texte);
        if (!lignes || !lignes.length) return; // journée vide ou illisible : ignorée
        const propre = JSON.stringify(lignes);
        if (!cache.has(k)) joursAjoutes.push([k, propre]);
        else if (canonJour(cache.get(k)) !== canonJour(propre))
          joursDifferents.push([k, propre]);
      } else {
        if (!cache.has(k)) {
          reglagesAjoutes.push([k, source[k]]);
          return;
        }
        try {
          const base = JSON.parse(cache.get(k)),
            ajout = JSON.parse(source[k]);
          const fusion = fusionProfonde(base, ajout);
          if (canon(fusion) !== canon(base))
            reglagesFusionnes.push([k, JSON.stringify(fusion)]);
        } catch (e) {}
      }
    });

    if (
      !joursAjoutes.length &&
      !joursDifferents.length &&
      !reglagesAjoutes.length &&
      !reglagesFusionnes.length
    ) {
      if (!silencieux)
        alert(
          "Rien à fusionner : tout ce que contient " +
            libelle +
            " est déjà dans la base.",
        );
      return "rien";
    }
    const ok = confirm(
      "Fusion avec la base en ligne\nSource : " +
        libelle +
        "\n\n" +
        "• " +
        joursAjoutes.length +
        " jour(s) absents de la base → seront ajoutés\n" +
        "• " +
        (reglagesAjoutes.length + reglagesFusionnes.length) +
        " réglage(s) (clients, tarifs, chauffeurs…) → les éléments manquants seront ajoutés, sans modifier ceux qui existent déjà\n" +
        "• " +
        joursDifferents.length +
        " jour(s) présents des deux côtés avec un contenu différent → je vous demanderai quoi faire\n\n" +
        "Rien de ce qui est déjà en base n'est supprimé ni écrasé à cette étape. Continuer ?",
    );
    if (!ok) return "annule";

    let remplacer = false;
    if (joursDifferents.length) {
      const liste =
        joursDifferents
          .map((x) => x[0].slice(9))
          .slice(0, 12)
          .join(", ") + (joursDifferents.length > 12 ? "…" : "");
      remplacer = confirm(
        joursDifferents.length +
          " jour(s) existent en base ET dans la source (" +
          libelle +
          ") avec un contenu différent :\n" +
          liste +
          "\n\nOK = remplacer ces jours par la version de la source (la version actuelle de la base sera PERDUE)\n" +
          "Annuler = garder la version de la base pour ces jours",
      );
    }
    joursAjoutes.forEach(([k, v]) => {
      cache.set(k, v);
      dirty.add(k);
    });
    reglagesAjoutes.forEach(([k, v]) => {
      cache.set(k, v);
      dirty.add(k);
    });
    reglagesFusionnes.forEach(([k, v]) => {
      cache.set(k, v);
      dirty.add(k);
    });
    if (remplacer)
      joursDifferents.forEach(([k, v]) => {
        cache.set(k, v);
        dirty.add(k);
      });
    cleListe = null;
    setOverlayTexte("Envoi des données vers la base… ne fermez pas la page");
    await envoyer();
    if (erreur || dirty.size) {
      alert(
        "⚠️ L'envoi n'est pas terminé (connexion ?). Rechargez la page pour réessayer.",
      );
      return "annule";
    }
    alert(
      "✅ Fusion terminée :\n• " +
        joursAjoutes.length +
        " jour(s) ajoutés" +
        (remplacer
          ? "\n• " + joursDifferents.length + " jour(s) remplacés"
          : joursDifferents.length
            ? "\n• " +
              joursDifferents.length +
              " jour(s) différents laissés tels quels dans la base"
            : "") +
        "\n• " +
        (reglagesAjoutes.length + reglagesFusionnes.length) +
        " réglage(s) complétés",
    );
    return "ok";
  }

  // ---- Reprise des anciennes données du navigateur (au premier chargement)
  function sourceLocale() {
    const source = {};
    try {
      for (let i = 0; i < lengthGetter.call(realLS); i++) {
        const k = rawKey.call(realLS, i);
        if (k && !LOCAL_ONLY.test(k)) source[k] = rawGet.call(realLS, k);
      }
    } catch (e) {}
    return source;
  }
  async function proposerMigration() {
    if (lireLocal("base_migration_ok")) return;
    const source = sourceLocale();
    if (!Object.keys(source).length) return;
    const res = await fusionner(source, "les données de ce navigateur", true);
    if (res === "ok" || res === "rien")
      ecrireLocal("base_migration_ok", new Date().toISOString());
  }
  // Relance manuelle (lien en bas à gauche) : utile si l'ancien lien a continué d'être utilisé
  async function reprendreDonneesLocales() {
    await fusionner(sourceLocale(), "les données de ce navigateur", false);
  }

  // Fusion d'un fichier de sauvegarde JSON (bouton « Fusionner une sauvegarde »)
  function choisirSauvegarde() {
    const inp = document.createElement("input");
    inp.type = "file";
    inp.accept = ".json";
    inp.onchange = () => {
      const f = inp.files[0];
      if (!f) return;
      const lecteur = new FileReader();
      lecteur.onload = async (e) => {
        let obj;
        try {
          obj = JSON.parse(e.target.result);
        } catch (err) {
          return alert("❌ Fichier illisible.");
        }
        if (!obj || typeof obj !== "object" || Array.isArray(obj))
          return alert("❌ Ce fichier n'est pas une sauvegarde du planning.");
        await fusionner(obj, "le fichier « " + f.name + " »");
      };
      lecteur.readAsText(f);
    };
    inp.click();
  }

  // ---- Interface : écran de connexion, chargement, indicateur
  let overlay = null;
  function creerOverlay() {
    overlay = document.createElement("div");
    overlay.style.cssText =
      "position:fixed;inset:0;background:#f0f2f5;z-index:20000;display:flex;align-items:center;justify-content:center;font-family:sans-serif;";
    document.body.appendChild(overlay);
  }
  function setOverlayTexte(t) {
    if (!overlay) return;
    overlay.innerHTML =
      '<div style="background:#fff;padding:30px 40px;border-radius:10px;box-shadow:0 4px 16px rgba(0,0,0,.15);font-size:15px;color:#333;">⏳ ' +
      t +
      "</div>";
  }
  function setOverlayErreur(msg) {
    overlay.innerHTML =
      '<div style="background:#fff;padding:30px 40px;border-radius:10px;box-shadow:0 4px 16px rgba(0,0,0,.15);max-width:420px;text-align:center;color:#333;">' +
      '<div style="color:#dc3545;font-weight:bold;margin-bottom:10px;">⚠️ Impossible de charger le planning</div>' +
      '<div style="font-size:13px;margin-bottom:16px;">' +
      msg +
      "</div>" +
      '<button id="base-reessayer" style="padding:10px 18px;border:none;border-radius:6px;background:#007bff;color:#fff;font-weight:bold;cursor:pointer;">Réessayer</button></div>';
    document.getElementById("base-reessayer").onclick = () => location.reload();
  }
  function afficherLogin() {
    return new Promise((resolve) => {
      overlay.innerHTML =
        '<form id="base-login" style="background:#fff;padding:30px 36px;border-radius:10px;box-shadow:0 4px 16px rgba(0,0,0,.15);width:320px;">' +
        '<h2 style="margin:0 0 6px;font-size:18px;color:#333;">Planning transport</h2>' +
        '<p style="margin:0 0 18px;font-size:13px;color:#666;">Connectez-vous pour accéder au planning.</p>' +
        '<input id="base-email" type="email" placeholder="Adresse e-mail" required autocomplete="username" style="width:100%;padding:10px;margin-bottom:10px;border:1px solid #ccc;border-radius:6px;font-size:14px;">' +
        '<input id="base-mdp" type="password" placeholder="Mot de passe" required autocomplete="current-password" style="width:100%;padding:10px;margin-bottom:10px;border:1px solid #ccc;border-radius:6px;font-size:14px;">' +
        '<div id="base-msg" style="color:#dc3545;font-size:13px;min-height:18px;margin-bottom:8px;"></div>' +
        '<button type="submit" style="width:100%;padding:11px;border:none;border-radius:6px;background:#007bff;color:#fff;font-weight:bold;font-size:14px;cursor:pointer;">Se connecter</button></form>';
      document.getElementById("base-login").onsubmit = async (e) => {
        e.preventDefault();
        const msg = document.getElementById("base-msg");
        msg.textContent = "";
        const { error } = await sb.auth.signInWithPassword({
          email: document.getElementById("base-email").value.trim(),
          password: document.getElementById("base-mdp").value,
        });
        if (error) {
          msg.textContent =
            "Connexion refusée : e-mail ou mot de passe incorrect.";
          return;
        }
        resolve();
      };
    });
  }

  let badge = null;
  function creerBadge(email) {
    badge = document.createElement("div");
    badge.style.cssText =
      "position:fixed;left:8px;bottom:8px;z-index:9000;background:#fff;border:1px solid #ddd;border-radius:14px;padding:4px 12px;font:12px sans-serif;box-shadow:0 2px 6px rgba(0,0,0,.15);";
    badge.innerHTML =
      '<span id="base-statut"></span> · <span style="color:#666">' +
      email +
      "</span>" +
      (Object.keys(sourceLocale()).length
        ? ' · <a href="#" id="base-reprise" style="color:#007bff;text-decoration:none;">Reprendre les données de ce navigateur</a>'
        : "") +
      ' · <a href="#" id="base-fusion" style="color:#007bff;text-decoration:none;">Fusionner une sauvegarde</a>' +
      ' · <a href="#" id="base-deco" style="color:#007bff;text-decoration:none;">Déconnexion</a>';
    document.body.appendChild(badge);
    const lienReprise = document.getElementById("base-reprise");
    if (lienReprise)
      lienReprise.onclick = (e) => {
        e.preventDefault();
        reprendreDonneesLocales();
      };
    document.getElementById("base-fusion").onclick = (e) => {
      e.preventDefault();
      choisirSauvegarde();
    };
    document.getElementById("base-deco").onclick = async (e) => {
      e.preventDefault();
      await envoyer();
      if (
        dirty.size &&
        !confirm(
          "Des modifications ne sont pas encore enregistrées. Se déconnecter quand même ?",
        )
      )
        return;
      await sb.auth.signOut();
      location.reload();
    };
    majStatut("ok");
  }
  function majStatut(etat) {
    const el = document.getElementById("base-statut");
    if (!el) return;
    if (etat === "ok") {
      el.textContent = "☁️ Enregistré";
      el.style.color = "#28a745";
    } else if (etat === "attente") {
      el.textContent = "⏳ Enregistrement…";
      el.style.color = "#ff9800";
    } else {
      el.textContent = "⚠️ Non enregistré (nouvelle tentative…)";
      el.style.color = "#dc3545";
      el.style.fontWeight = "bold";
    }
  }

  // ---- Démarrage
  function chargerScriptPrincipal() {
    return new Promise((resolve, reject) => {
      const s = document.createElement("script");
      s.src = "script.js";
      s.onload = resolve;
      s.onerror = () => reject(new Error("script.js introuvable"));
      document.body.appendChild(s);
    });
  }

  async function demarrer() {
    creerOverlay();
    setOverlayTexte("Connexion…");
    try {
      if (!window.supabase)
        throw new Error(
          "Bibliothèque Supabase non chargée (connexion Internet ?).",
        );
      sb = window.supabase.createClient(SUPABASE_URL, SUPABASE_KEY, {
        auth: {
          storage: adapterSession,
          persistSession: true,
          autoRefreshToken: true,
        },
      });
      ["undoStack", "redoStack", "lastCapturedData"].forEach((k) =>
        rawRemove.call(realLS, k),
      );

      let { data } = await sb.auth.getSession();
      if (!data.session) {
        await afficherLogin();
        data = (await sb.auth.getSession()).data;
      }
      setOverlayTexte("Chargement du planning…");
      await chargerDepuisBase();
      installerCache();
      actif = true;
      await proposerMigration();
      if (joursCorriges.size) {
        joursCorriges.forEach((k) => dirty.add(k));
        majStatut("attente");
        envoyer();
      }

      sb.auth.onAuthStateChange((evt) => {
        if (evt === "SIGNED_OUT") location.reload();
      });
      await chargerScriptPrincipal();
      overlay.remove();
      creerBadge((data.session && data.session.user.email) || "");
      // script.js a été chargé après les événements de chargement : on les rejoue
      document.dispatchEvent(new Event("DOMContentLoaded"));
      window.dispatchEvent(new Event("load"));
    } catch (e) {
      console.error(e);
      setOverlayErreur(String((e && e.message) || e));
    }
  }

  window.addEventListener("beforeunload", (e) => {
    if (dirty.size || enCours) {
      e.preventDefault();
      e.returnValue = "";
    }
  });
  document.addEventListener("visibilitychange", () => {
    if (document.hidden && actif) envoyer();
  });

  // Fonctions utilisées par script.js
  window.baseAttendreSynchro = () => envoyer();
  window.baseViderTout = async () => {
    localStorage.clear();
    await envoyer();
  };

  if (document.readyState === "loading")
    document.addEventListener("DOMContentLoaded", demarrer, { once: true });
  else demarrer();
})();
