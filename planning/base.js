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
  const adapterSession = {
    getItem: (k) => rawGet.call(realLS, k),
    setItem: (k, v) => rawSet.call(realLS, k, v),
    removeItem: (k) => rawRemove.call(realLS, k),
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
    jours.forEach((lignes, date) =>
      cache.set("planning_" + date, JSON.stringify(lignes)),
    );
    cleListe = null;
  }

  // ---- Reprise des anciennes données du navigateur (une seule fois)
  async function proposerMigration() {
    const anciennes = [];
    for (let i = 0; i < lengthGetter.call(realLS); i++) {
      const k = rawKey.call(realLS, i);
      if (k && !LOCAL_ONLY.test(k)) anciennes.push(k);
    }
    if (!anciennes.length || rawGet.call(realLS, "base_migration_ok")) return;
    const aImporter = anciennes.filter((k) => !cache.has(k));
    const nbJours = aImporter.filter((k) => CLE_JOUR.test(k)).length;
    const nbReglages = aImporter.length - nbJours;
    if (!aImporter.length) {
      rawSet.call(realLS, "base_migration_ok", new Date().toISOString());
      return;
    }
    const ok = confirm(
      "Des données de ce navigateur ne sont pas encore dans la base en ligne :\n" +
        "• " +
        nbJours +
        " jour(s) de planning\n• " +
        nbReglages +
        " réglage(s) (clients, tarifs, chauffeurs…)\n\n" +
        "Les envoyer maintenant dans la base ?\n(Ce qui existe déjà en ligne ne sera jamais écrasé. " +
        "Les données de ce navigateur ne sont pas supprimées.)",
    );
    if (!ok) return;
    setOverlayTexte("Envoi des données vers la base… ne fermez pas la page");
    aImporter.forEach((k) => {
      const v = rawGet.call(realLS, k);
      if (v !== null) {
        cache.set(k, v);
        dirty.add(k);
      }
    });
    cleListe = null;
    await envoyer();
    if (erreur || dirty.size) {
      alert(
        "⚠️ L'envoi n'est pas terminé (connexion ?). Rechargez la page pour réessayer.",
      );
      return;
    }
    rawSet.call(realLS, "base_migration_ok", new Date().toISOString());
    alert(
      "✅ " +
        nbJours +
        " jour(s) et " +
        nbReglages +
        " réglage(s) envoyés dans la base.",
    );
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
      '</span> · <a href="#" id="base-deco" style="color:#007bff;text-decoration:none;">Déconnexion</a>';
    document.body.appendChild(badge);
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
