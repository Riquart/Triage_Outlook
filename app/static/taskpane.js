/**
 * Triage_Outlook — Logique client du volet latéral de tri par lot.
 */
/* global Office */
(() => {
  "use strict";

  const CONFIG = JSON.parse(document.getElementById("app-config").textContent);

  const STORAGE_KEY_RULES = "triage_outlook_folders_rules";
  const STORAGE_KEY_TOKEN = "triage_outlook_auth_token";

  const DEFAULT_FOLDERS = [
    {
      id: "folder-commandes",
      name: "Commandes & Matériel",
      rule: "Mails des clients souhaitant acheter un nouveau lecteur, renouveler leur matériel, commander des accessoires ou obtenir un devis.",
      active: true,
    },
    {
      id: "folder-support",
      name: "Support & Pannes",
      rule: "Blocages techniques, lecteur non reconnu, erreurs de télétransmission, téléphones ou bugs nécessitant une assistance.",
      active: true,
    },
    {
      id: "folder-facturation",
      name: "Facturation & Comptabilité",
      rule: "Demandes de factures, duplicatas, changements de RIB, questions sur les prélèvements et contestations.",
      active: true,
    },
    {
      id: "folder-resiliations",
      name: "Résiliations & Départs",
      rule: "Cessions de cabinet, départs à la retraite, vente d'activité et demandes de résiliation de contrat.",
      active: true,
    },
    {
      id: "folder-divers",
      name: "Information & Divers",
      rule: "Newsletters, informations générales, courriers informatifs sans action urgente requise.",
      active: true,
    },
  ];

  const state = {
    selectedEmails: [],
    folders: [],
    classifications: [],
    busy: false,
    preview: false,
  };

  const el = new Proxy({}, {
    get: (_, id) => document.getElementById(id.replace(/_/g, "-")),
  });

  function getStoredToken() {
    return (localStorage.getItem(STORAGE_KEY_TOKEN) || "").trim();
  }
  function setStoredToken(val) {
    if (val) localStorage.setItem(STORAGE_KEY_TOKEN, val.trim());
    else localStorage.removeItem(STORAGE_KEY_TOKEN);
    updateAuthUI();
  }
  function updateAuthUI() {
    const t = getStoredToken();
    if (el.input_auth_token) el.input_auth_token.value = t;
    if (!el.auth_status_text) return;
    if (t) {
      el.auth_status_text.textContent = "✓ Code secret configuré";
      el.auth_status_text.className = "text-[10px] text-emerald-700 font-medium";
    } else if (CONFIG.authRequired) {
      el.auth_status_text.textContent = "⚠️ Code secret requis par le serveur";
      el.auth_status_text.className = "text-[10px] text-amber-700 font-medium";
    } else {
      el.auth_status_text.textContent = "Optionnel (aucun code exigé)";
      el.auth_status_text.className = "text-[10px] text-slate-400";
    }
  }

  function loadFolders() {
    try {
      const saved = localStorage.getItem(STORAGE_KEY_RULES);
      if (saved) {
        state.folders = JSON.parse(saved);
      } else {
        state.folders = JSON.parse(JSON.stringify(DEFAULT_FOLDERS));
      }
    } catch {
      state.folders = JSON.parse(JSON.stringify(DEFAULT_FOLDERS));
    }
    renderFolders();
  }

  function saveFolders() {
    localStorage.setItem(STORAGE_KEY_RULES, JSON.stringify(state.folders));
    renderFolders();
  }

  function renderFolders() {
    const activeCount = state.folders.filter((f) => f.active).length;
    if (el.active_rules_count) el.active_rules_count.textContent = String(activeCount);

    if (!el.folders_container) return;
    el.folders_container.innerHTML = "";

    state.folders.forEach((f, idx) => {
      const card = document.createElement("div");
      card.className = "rounded-xl border border-slate-200 bg-white p-3 space-y-2 shadow-xs";
      card.innerHTML = `
        <div class="flex items-center justify-between gap-2">
          <label class="flex items-center gap-2 cursor-pointer font-bold text-slate-800 text-xs">
            <input type="checkbox" class="folder-check rounded text-indigo-600 focus:ring-indigo-500" data-idx="${idx}" ${f.active ? "checked" : ""}>
            <span>📁 ${escapeHtml(f.name)}</span>
          </label>
          <button type="button" class="folder-del text-slate-300 hover:text-red-600 text-xs" data-idx="${idx}" title="Supprimer">✕</button>
        </div>
        <div>
          <label class="block text-[10px] text-slate-400 font-semibold uppercase mb-1">Règle d'attribution</label>
          <textarea rows="2" class="folder-rule w-full rounded-lg border border-slate-200 p-2 text-xs text-slate-700 placeholder:text-slate-400 focus:ring-1 focus:ring-indigo-500 focus:outline-none" data-idx="${idx}" placeholder="Décrivez les e-mails à classer ici...">${escapeHtml(f.rule || "")}</textarea>
        </div>
      `;
      el.folders_container.appendChild(card);
    });

    el.folders_container.querySelectorAll(".folder-check").forEach((chk) => {
      chk.addEventListener("change", (e) => {
        const idx = Number(e.target.dataset.idx);
        state.folders[idx].active = e.target.checked;
        saveFolders();
      });
    });

    el.folders_container.querySelectorAll(".folder-rule").forEach((txt) => {
      txt.addEventListener("input", (e) => {
        const idx = Number(e.target.dataset.idx);
        state.folders[idx].rule = e.target.value.trim();
        localStorage.setItem(STORAGE_KEY_RULES, JSON.stringify(state.folders));
      });
    });

    el.folders_container.querySelectorAll(".folder-del").forEach((btn) => {
      btn.addEventListener("click", (e) => {
        const idx = Number(e.target.dataset.idx);
        state.folders.splice(idx, 1);
        saveFolders();
      });
    });
  }

  async function refreshSelection() {
    clearAlerts();
    state.selectedEmails = [];

    if (!window.Office || !Office.context || !Office.context.mailbox) {
      setupMockEmails();
      return;
    }

    const mb = Office.context.mailbox;

    if (typeof mb.getSelectedItemsAsync === "function") {
      mb.getSelectedItemsAsync((asyncResult) => {
        if (asyncResult.status === Office.AsyncResultStatus.Succeeded && asyncResult.value && asyncResult.value.length > 0) {
          state.selectedEmails = asyncResult.value.map((item) => ({
            id: item.itemId,
            subject: item.subject || "(sans objet)",
            sender: "Expéditeur sélectionné",
            body: item.subject || "",
          }));
          renderSelectedEmails();
        } else {
          checkSingleItem();
        }
      });
    } else {
      checkSingleItem();
    }
  }

  function checkSingleItem() {
    if (Office.context && Office.context.mailbox && Office.context.mailbox.item) {
      const item = Office.context.mailbox.item;
      const sender = item.from ? (item.from.displayName ? `${item.from.displayName} <${item.from.emailAddress}>` : item.from.emailAddress) : "";
      item.body.getAsync(Office.CoercionType.Text, (res) => {
        const bodyText = res.status === Office.AsyncResultStatus.Succeeded ? res.value : "";
        state.selectedEmails = [
          {
            id: item.itemId,
            subject: item.subject || "(sans objet)",
            sender: sender,
            body: bodyText,
          },
        ];
        renderSelectedEmails();
      });
    } else {
      renderSelectedEmails();
    }
  }

  function setupMockEmails() {
    state.preview = true;
    state.selectedEmails = [
      {
        id: "mock-1",
        subject: "Commande de 2 lecteurs de carte Vital'Act",
        sender: "Dr Martin <martin@cabinet-kine.fr>",
        body: "Bonjour, nous souhaitons commander 2 nouveaux lecteurs portables pour notre nouveau remplaçant. Pouvez-vous nous envoyer un devis ?",
      },
      {
        id: "mock-2",
        subject: "Erreur télétransmission FSE 429",
        sender: "Cabinet Soins Infirmiers <contact@ide-soins.fr>",
        body: "Bonjour, depuis ce matin nos feuilles de soins sont rejetées avec le code erreur 429. Le lecteur bip en boucle. C'est très urgent merci.",
      },
      {
        id: "mock-3",
        subject: "Demande de duplicata facture n° 2026-08",
        sender: "Comptabilité Santé <compta@centre-medical.fr>",
        body: "Bonjour, pouvez-vous nous renvoyer la facture du mois d'août avec notre nouveau RIB ? Merci d'avance.",
      },
      {
        id: "mock-4",
        subject: "Départ à la retraite fin du mois",
        sender: "M. Lefebvre <lefebvre@orange.fr>",
        body: "Madame, Monsieur, je cesse définitivement mon activité libérale le 30 septembre pour départ à la retraite. Comment résilier mon abonnement ?",
      },
    ];
    renderSelectedEmails();
  }

  function renderSelectedEmails() {
    const count = state.selectedEmails.length;
    if (el.selected_count) el.selected_count.textContent = String(count);

    if (!el.emails_list) return;
    if (count === 0) {
      el.emails_list.innerHTML = `
        <div class="rounded-xl border border-dashed border-slate-300 p-6 text-center">
          <p class="text-xs text-slate-500 font-medium">Aucun e-mail sélectionné</p>
          <p class="text-[10px] text-slate-400 mt-1">Cochez ou sélectionnez des messages dans Outlook puis cliquez sur « Recharger sélection ».</p>
        </div>
      `;
      return;
    }

    el.emails_list.innerHTML = "";
    state.selectedEmails.forEach((email) => {
      const card = document.createElement("div");
      card.className = "rounded-lg border border-slate-200 bg-white p-2.5 shadow-xs space-y-1";
      card.innerHTML = `
        <div class="flex items-start justify-between gap-2">
          <div class="min-w-0 flex-1">
            <p class="font-bold text-slate-900 truncate text-[11px]">${escapeHtml(email.subject)}</p>
            <p class="text-[10px] text-slate-500 truncate">${escapeHtml(email.sender || "")}</p>
          </div>
          <span class="shrink-0 text-[10px] bg-slate-100 text-slate-600 px-1.5 py-0.5 rounded font-mono">En attente</span>
        </div>
      `;
      el.emails_list.appendChild(card);
    });
  }

  async function runTriage() {
    if (state.busy) return;
    clearAlerts();

    const activeFolders = state.folders.filter((f) => f.active);
    if (!activeFolders.length) {
      showAlert("error", "Aucun dossier actif pour le tri. Activez au moins un dossier dans l'onglet Règles.");
      return;
    }
    if (!state.selectedEmails.length) {
      showAlert("error", "Aucun e-mail sélectionné.");
      return;
    }

    const token = getStoredToken();
    if (CONFIG.authRequired && !token) {
      if (el.settings_card) el.settings_card.hidden = false;
      showAlert("error", "Code secret requis pour lancer le tri.");
      return;
    }

    setBusy(true);
    try {
      const payload = {
        emails: state.selectedEmails,
        folders: activeFolders.map((f) => ({ name: f.name, folder_id: f.id, rule: f.rule })),
        auth_token: token,
      };

      const res = await fetch(CONFIG.apiUrl, {
        method: "POST",
        headers: { "Content-Type": "application/json", "X-Auth-Token": token },
        body: JSON.stringify(payload),
      });

      if (!res.ok) {
        const data = await res.json().catch(() => ({}));
        throw new Error(data.detail || `Erreur serveur (${res.status})`);
      }

      const result = await res.json();
      state.classifications = result.classifications || [];
      renderResults();
      showAlert("info", `✓ ${state.classifications.length} e-mails classés. Validez le déplacement ci-dessous.`);
    } catch (err) {
      showAlert("error", err.message || String(err));
    } finally {
      setBusy(false);
    }
  }

  function renderResults() {
    if (!el.results_panel || !el.results_list) return;
    el.results_panel.hidden = false;
    el.results_count.textContent = `${state.classifications.length} classé(s)`;
    el.results_list.innerHTML = "";

    const activeFolders = state.folders.filter((f) => f.active);

    state.classifications.forEach((c, idx) => {
      const email = state.selectedEmails.find((e) => e.id === c.item_id) || { subject: "(message)" };
      const row = document.createElement("div");
      row.className = "rounded-lg border border-slate-200 bg-white p-2.5 shadow-xs space-y-2";

      let optionsHtml = activeFolders
        .map((f) => `<option value="${f.id}" ${f.name === c.target_folder_name ? "selected" : ""}>${escapeHtml(f.name)}</option>`)
        .join("");

      const confBadge = c.confidence === "high"
        ? '<span class="bg-emerald-100 text-emerald-800 text-[9px] px-1 py-0.5 rounded font-bold">Fiable</span>'
        : '<span class="bg-amber-100 text-amber-800 text-[9px] px-1 py-0.5 rounded font-bold">À vérifier</span>';

      row.innerHTML = `
        <div class="flex items-start justify-between gap-1.5">
          <p class="font-bold text-slate-900 text-xs leading-tight flex-1 truncate">${escapeHtml(email.subject)}</p>
          ${confBadge}
        </div>
        <div class="flex items-center gap-2">
          <span class="text-[10px] text-slate-400 font-semibold shrink-0">Vers :</span>
          <select class="res-folder-select flex-1 rounded border border-indigo-200 bg-indigo-50/50 p-1 text-[11px] font-semibold text-indigo-950 focus:outline-none" data-idx="${idx}">
            ${optionsHtml}
          </select>
        </div>
        <p class="text-[10px] text-slate-500 italic">💡 ${escapeHtml(c.justification || "")}</p>
      `;
      el.results_list.appendChild(row);
    });

    el.results_list.querySelectorAll(".res-folder-select").forEach((sel) => {
      sel.addEventListener("change", (e) => {
        const idx = Number(e.target.dataset.idx);
        const selectedOption = e.target.options[e.target.selectedIndex];
        state.classifications[idx].target_folder_id = e.target.value;
        state.classifications[idx].target_folder_name = selectedOption.text;
      });
    });
  }

  async function applyMove() {
    if (!state.classifications.length) return;
    clearAlerts();

    if (state.preview || !window.Office || !Office.context || !Office.context.mailbox) {
      showAlert("info", `✓ Simulation réussie : ${state.classifications.length} e-mails déplacés virtuellement.`);
      return;
    }

    setBusy(true);
    let successCount = 0;
    let failCount = 0;

    for (const c of state.classifications) {
      try {
        await moveSingleEmailEWS(c.item_id, c.target_folder_id);
        successCount++;
      } catch (err) {
        console.warn("MoveItem failed for", c.item_id, err);
        failCount++;
      }
    }

    setBusy(false);
    if (failCount === 0) {
      showAlert("info", `🎉 ${successCount} messages déplacés avec succès !`);
      state.classifications = [];
      el.results_panel.hidden = true;
      refreshSelection();
    } else {
      showAlert("error", `${successCount} déplacé(s), ${failCount} échec(s). Vérifiez vos droits de boîte.`);
    }
  }

  function moveSingleEmailEWS(itemId, folderId) {
    return new Promise((resolve, reject) => {
      const ewsXml = `<?xml version="1.0" encoding="utf-8"?>
<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
               xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages"
               xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types"
               xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
  <soap:Header>
    <t:RequestServerVersion Version="Exchange2013" />
  </soap:Header>
  <soap:Body>
    <m:MoveItem>
      <m:ToFolderId>
        <t:FolderId Id="${folderId}"/>
      </m:ToFolderId>
      <m:ItemIds>
        <t:ItemId Id="${itemId}"/>
      </m:ItemIds>
    </m:MoveItem>
  </soap:Body>
</soap:Envelope>`;

      Office.context.mailbox.makeEwsRequestAsync(ewsXml, (res) => {
        if (res.status === Office.AsyncResultStatus.Succeeded) {
          resolve(res.value);
        } else {
          reject(new Error(res.error ? res.error.message : "Erreur EWS"));
        }
      });
    });
  }

  function detectOutlookFolders() {
    clearAlerts();
    if (!window.Office || !Office.context || !Office.context.mailbox) {
      showAlert("info", "Mode aperçu : 5 dossiers types de démonstration chargés.");
      return;
    }

    const ewsFind = `<?xml version="1.0" encoding="utf-8"?>
<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
               xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages"
               xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types"
               xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
  <soap:Header>
    <t:RequestServerVersion Version="Exchange2013" />
  </soap:Header>
  <soap:Body>
    <m:FindFolder Traversal="Deep">
      <m:FolderShape>
        <t:BaseShape>Default</t:BaseShape>
      </m:FolderShape>
      <m:ParentFolderIds>
        <t:DistinguishedFolderId Id="msgfolderroot"/>
      </m:ParentFolderIds>
    </m:FindFolder>
  </soap:Body>
</soap:Envelope>`;

    Office.context.mailbox.makeEwsRequestAsync(ewsFind, (res) => {
      if (res.status === Office.AsyncResultStatus.Succeeded) {
        const parser = new DOMParser();
        const xml = parser.parseFromString(res.value, "text/xml");
        const folderNodes = xml.querySelectorAll("Folder, CalendarFolder, ContactsFolder");
        const found = [];

        folderNodes.forEach((node) => {
          const idEl = node.querySelector("FolderId");
          const nameEl = node.querySelector("DisplayName");
          const id = idEl ? idEl.getAttribute("Id") : null;
          const name = nameEl ? nameEl.textContent : null;
          if (id && name && !["Boîte de réception", "Éléments envoyés", "Éléments supprimés", "Courrier indésirable", "Boîte d'envoi"].includes(name)) {
            const existing = state.folders.find((f) => f.id === id || f.name.toLowerCase() === name.toLowerCase());
            found.push({
              id: id,
              name: name,
              rule: existing ? existing.rule : "",
              active: existing ? existing.active : true,
            });
          }
        });

        if (found.length) {
          state.folders = found;
          saveFolders();
          showAlert("info", `✓ ${found.length} dossiers Outlook synchronisés !`);
        } else {
          showAlert("info", "Aucun sous-dossier personnalisé trouvé. Vous pouvez en ajouter manuellement.");
        }
      } else {
        showAlert("error", "Lecture EWS refusée par le serveur de messagerie. Vos dossiers par défaut restent actifs.");
      }
    });
  }

  function setBusy(busy) {
    state.busy = busy;
    if (el.btn_run_triage) el.btn_run_triage.disabled = busy;
    if (el.spinner_triage) el.spinner_triage.hidden = !busy;
    if (el.btn_apply_move) el.btn_apply_move.disabled = busy;
  }

  function showAlert(type, msg) {
    if (type === "error") {
      if (el.banner_error_text) el.banner_error_text.textContent = msg;
      if (el.banner_error) el.banner_error.hidden = false;
    } else {
      if (el.banner_info_text) el.banner_info_text.textContent = msg;
      if (el.banner_info) el.banner_info.hidden = false;
    }
  }

  function clearAlerts() {
    if (el.banner_error) el.banner_error.hidden = true;
    if (el.banner_info) el.banner_info.hidden = true;
  }

  function escapeHtml(str) {
    return (str || "").replace(/[&<>"']/g, (m) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" })[m]);
  }

  function bindEvents() {
    const TABS = [
      { btn: "tab_btn_triage", pane: "tab_triage" },
      { btn: "tab_btn_rules", pane: "tab_rules" },
      { btn: "tab_btn_diagnostic", pane: "tab_diagnostic" },
    ];

    function switchTab(activeKey) {
      TABS.forEach((t) => {
        const btn = el[t.btn];
        const pane = el[t.pane];
        if (!btn || !pane) return;
        const isActive = t.pane === activeKey;
        pane.hidden = !isActive;
        btn.className = isActive
          ? "flex-1 py-2.5 text-center border-b-2 border-indigo-600 text-indigo-600 font-semibold"
          : "flex-1 py-2.5 text-center border-b-2 border-transparent text-slate-500 hover:text-slate-800 font-semibold";
      });
    }

    TABS.forEach((t) => {
      const btn = el[t.btn];
      if (btn) btn.addEventListener("click", () => switchTab(t.pane));
    });

    if (el.btn_run_diagnostic) el.btn_run_diagnostic.addEventListener("click", runDiagnostic);
    if (el.btn_copy_diagnostic) el.btn_copy_diagnostic.addEventListener("click", copyDiagnostic);

    if (el.btn_refresh_selection) el.btn_refresh_selection.addEventListener("click", refreshSelection);
    if (el.btn_run_triage) el.btn_run_triage.addEventListener("click", runTriage);
    if (el.btn_apply_move) el.btn_apply_move.addEventListener("click", applyMove);
    if (el.btn_detect_folders) el.btn_detect_folders.addEventListener("click", detectOutlookFolders);

    if (el.btn_add_folder) {
      el.btn_add_folder.addEventListener("click", () => {
        const name = prompt("Nom du nouveau dossier :", "Nouveau dossier");
        if (name && name.trim()) {
          state.folders.push({
            id: `manual-${Date.now()}`,
            name: name.trim(),
            rule: "",
            active: true,
          });
          saveFolders();
        }
      });
    }

    if (el.btn_export_rules) {
      el.btn_export_rules.addEventListener("click", () => {
        const data = JSON.stringify(state.folders, null, 2);
        navigator.clipboard.writeText(data);
        alert("Configuration des règles copiée dans votre presse-papier !");
      });
    }

    if (el.btn_import_rules) {
      el.btn_import_rules.addEventListener("click", () => {
        const val = prompt("Collez le JSON exporté des règles :");
        if (val) {
          try {
            const parsed = JSON.parse(val);
            if (Array.isArray(parsed)) {
              state.folders = parsed;
              saveFolders();
              alert("Règles importées avec succès !");
            }
          } catch {
            alert("Format JSON invalide.");
          }
        }
      });
    }

    if (el.btn_settings_toggle) {
      el.btn_settings_toggle.addEventListener("click", () => {
        el.settings_card.hidden = !el.settings_card.hidden;
      });
    }
    if (el.btn_settings_close) {
      el.btn_settings_close.addEventListener("click", () => {
        el.settings_card.hidden = true;
      });
    }
    if (el.btn_save_auth_token) {
      el.btn_save_auth_token.addEventListener("click", () => {
        setStoredToken(el.input_auth_token ? el.input_auth_token.value : "");
        el.settings_card.hidden = true;
      });
    }
  }

  // ------------------------------------------------------------------ //
  // Diagnostic des capacités de la boîte aux lettres
  // ------------------------------------------------------------------ //

  function _ewsFindFolderXml() {
    return `<?xml version="1.0" encoding="utf-8"?>
<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
               xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages"
               xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types"
               xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
  <soap:Header>
    <t:RequestServerVersion Version="Exchange2013" />
  </soap:Header>
  <soap:Body>
    <m:FindFolder Traversal="Deep">
      <m:FolderShape>
        <t:BaseShape>Default</t:BaseShape>
      </m:FolderShape>
      <m:ParentFolderIds>
        <t:DistinguishedFolderId Id="msgfolderroot"/>
      </m:ParentFolderIds>
    </m:FindFolder>
  </soap:Body>
</soap:Envelope>`;
  }

  async function runDiagnostic() {
    const lines = [];
    const sep = "────────────────────────────────────────";
    const log = (s) => lines.push(s || "");
    const ok = (s) => log("  [OK]    " + s);
    const ko = (s) => log("  [ECHEC] " + s);
    const info = (s) => log("  [info]  " + s);
    const dump = () => { if (el.diagnostic_output) el.diagnostic_output.textContent = lines.join("\n"); };

    if (el.diagnostic_output) el.diagnostic_output.textContent = "Diagnostic en cours…";
    if (el.btn_run_diagnostic) el.btn_run_diagnostic.disabled = true;

    log("=== DIAGNOSTIC Triage_Outlook ===");
    log("Date    : " + new Date().toLocaleString("fr-FR"));
    log("Version : " + CONFIG.version);
    log(sep);

    // --- TEST 1 : hôte et identité ---
    log("TEST 1 — Hôte et identité de session");
    if (!window.Office || !Office.context) {
      ko("Office.js indisponible (page ouverte dans un navigateur)");
      dump();
      if (el.btn_run_diagnostic) el.btn_run_diagnostic.disabled = false;
      return;
    }
    const mb = Office.context.mailbox;
    try { info("Hôte        : " + String(Office.context.host)); } catch (e) { info("Hôte : ?"); }
    try { info("Plateforme  : " + String(Office.context.platform)); } catch (e) { info("Plateforme : ?"); }
    if (mb && mb.userProfile) {
      info("Connecté en : " + (mb.userProfile.emailAddress || "?"));
      info("Compte      : " + (mb.userProfile.displayName || "?"));
    }
    // Détection du contexte "boîte partagée" (accès délégué)
    const sharedCtx = { targetRestUrl: "", targetMailbox: "", owner: "", permissions: null };
    if (mb.item && typeof mb.item.getSharedPropertiesAsync === "function") {
      await new Promise((resolve) => {
        mb.item.getSharedPropertiesAsync((res) => {
          if (res.status === Office.AsyncResultStatus.Succeeded && res.value) {
            const v = res.value;
            sharedCtx.targetMailbox = v.targetMailbox || "";
            sharedCtx.owner = v.owner || "";
            sharedCtx.targetRestUrl = v.targetRestUrl || "";
            sharedCtx.permissions = v.delegatePermissions;
            if (sharedCtx.targetMailbox) {
              info("CONTEXTE PARTAGÉ DÉTECTÉ");
              info("   Boîte visée  : " + sharedCtx.targetMailbox);
              info("   Propriétaire : " + (sharedCtx.owner || "?"));
            } else {
              info("Contexte : boîte personnelle (aucune délégation détectée)");
            }
            info("   Permissions  : " + (typeof v.delegatePermissions === "number" ? v.delegatePermissions : "?"));
            if (sharedCtx.targetRestUrl) info("   REST URL     : " + sharedCtx.targetRestUrl);
          } else {
            info("getSharedPropertiesAsync indisponible/échoué");
          }
          resolve();
        });
      });
    } else {
      info("getSharedPropertiesAsync non disponible sur cet élément");
    }

    const versions = ["1.1","1.3","1.5","1.8","1.10","1.11","1.12","1.13","1.14","1.15"];
    const supported = versions.filter((v) => { try { return Office.context.requirements.isSetSupported("Mailbox", v); } catch (e) { return false; } });
    info("Versions Mailbox supportées : " + (supported.join(", ") || "aucune"));
    log(sep);

    // --- TEST 2 : lecture de la sélection ---
    log("TEST 2 — Lecture de la sélection");
    if (typeof mb.getSelectedItemsAsync === "function") {
      if (supported.indexOf("1.13") !== -1) { ok("getSelectedItemsAsync disponible (multi-sélection)"); }
      else { info("getSelectedItemsAsync exposé, mais Mailbox 1.13 non confirmé"); }
      await new Promise((resolve) => {
        mb.getSelectedItemsAsync((res) => {
          if (res.status === Office.AsyncResultStatus.Succeeded) {
            const n = (res.value || []).length;
            if (n > 0) {
              ok(n + " élément(s) dans la sélection");
              (res.value || []).slice(0, 3).forEach((it) => info("   • " + String(it.subject || "").slice(0, 70)));
            } else { ko("Sélection vide (sélectionnez des messages puis relancez)"); }
          } else {
            ko("Échec : " + (res.error ? res.error.message : "inconnu"));
          }
          resolve();
        });
      });
    } else {
      ko("getSelectedItemsAsync indisponible");
    }
    if (mb.item) { ok("Un message est ouvert (mode lecture simple)"); } else { info("Aucun message ouvert (mode multi-sélection)"); }
    log(sep);

    // --- TEST 3 : EWS ---
    log("TEST 3 — EWS makeEwsRequestAsync (indispensable au déplacement)");
    if (typeof mb.makeEwsRequestAsync !== "function") {
      ko("makeEwsRequestAsync indisponible sur ce client");
      dump();
      if (el.btn_run_diagnostic) el.btn_run_diagnostic.disabled = false;
      return;
    }
    ok("makeEwsRequestAsync exposé par le client");
    const ewsOk = await new Promise((resolve) => {
      mb.makeEwsRequestAsync(_ewsFindFolderXml(), (res) => {
        if (res.status === Office.AsyncResultStatus.Succeeded) {
          try {
            const xml = new DOMParser().parseFromString(res.value, "text/xml");
            let nodes = xml.getElementsByTagName("t:Folder");
            if (!nodes.length) nodes = xml.getElementsByTagName("Folder");
            ok("EWS OPÉRATIONNEL — " + nodes.length + " dossier(s) lus");
            const names = [];
            for (let i = 0; i < nodes.length; i++) {
              const dn = nodes[i].getElementsByTagName("t:DisplayName")[0] || nodes[i].getElementsByTagName("DisplayName")[0];
              if (dn) names.push(dn.textContent);
            }
            if (names.length) info("Dossiers : " + names.slice(0, 15).join(" | "));
            resolve(true);
          } catch (e) {
            ko("Réponse EWS illisible : " + e.message);
            resolve(false);
          }
        } else {
          ko("EWS REFUSÉ : " + (res.error ? res.error.message : "inconnu"));
          info("→ Comportement attendu dans une boîte partagée : le déplacement direct sera impossible.");
          resolve(false);
        }
      });
    });
    log(sep);

    // --- TEST 4 : catégories de couleur ---
    log("TEST 4 — Catégories existantes de la boîte partagée");
    const item = mb.item;
    let existingCats = [];
    if (!item) {
      info("Aucun message ouvert : test ignoré (ouvrez un message de la boîte partagée)");
    } else {
      // 4a. Lire les catégories déjà présentes sur le message
      if (item.categories && typeof item.categories.getAsync === "function") {
        await new Promise((resolve) => {
          item.categories.getAsync((res) => {
            if (res.status === Office.AsyncResultStatus.Succeeded) {
              existingCats = res.value || [];
              ok("Lecture des catégories du message : OK (" + existingCats.length + " sur ce message)");
              if (existingCats.length) info("   Déjà présentes : " + existingCats.slice(0, 8).join(", "));
            } else {
              ko("Lecture des catégories impossible : " + (res.error ? res.error.message : "?"));
            }
            resolve();
          });
        });
      }

      // 4b. Lire la liste maître des catégories
      let masterCats = [];
      if (mb.masterCategories && typeof mb.masterCategories.getAsync === "function") {
        await new Promise((resolve) => {
          mb.masterCategories.getAsync((res) => {
            if (res.status === Office.AsyncResultStatus.Succeeded) {
              masterCats = (res.value || []).map((c) => c.displayName || c);
              ok("Liste maître lisible : " + masterCats.length + " catégories disponibles");
              if (masterCats.length) info("   Exemples : " + masterCats.slice(0, 10).join(", "));
            } else {
              ko("Liste maître illisible : " + (res.error ? res.error.message : "?"));
            }
            resolve();
          });
        });
      }

      // 4c. Tentative d'application d'une catégorie EXISTANTE (et non inventée)
      if (item.categories && typeof item.categories.addAsync === "function" && masterCats.length) {
        const free = masterCats.find((c) => existingCats.indexOf(c) === -1);
        if (!free) {
          info("Le message porte déjà toutes les catégories : test d'écriture ignoré (aucune modification)");
        } else {
          await new Promise((resolve) => {
            item.categories.addAsync([free], (res) => {
              if (res.status === Office.AsyncResultStatus.Succeeded) {
                ok("APPLICATION D'UNE CATÉGORIE EXISTANTE : OK  (« " + free + " »)");
                item.categories.removeAsync([free], (r2) => {
                  if (r2.status === Office.AsyncResultStatus.Succeeded) { ok("Catégorie de test retirée (aucune trace laissée)"); }
                  else { info("Nettoyage non confirmé : vérifiez la catégorie « " + free + " » sur ce message"); }
                  resolve();
                });
              } else {
                ko("Échec application (« " + free + " ») : " + (res.error ? res.error.message : "?"));
                info("   → Confirme que les catégories sont verrouillées pour un délégué.");
                resolve();
              }
            });
          });
        }
      } else if (!masterCats.length) {
        info("Aucune catégorie maître disponible : test d'écriture ignoré");
      }
    }
    log(sep);

    // --- TEST 5 : Outlook REST API (voie officielle Microsoft) ---
    log("TEST 5 — Outlook REST API (voie recommandée par Microsoft hors EWS)");
    if (!sharedCtx.targetMailbox) {
      info("Non applicable : vous n'êtes pas dans un contexte partagé.");
    } else if (typeof mb.getCallbackTokenAsync !== "function") {
      ko("getCallbackTokenAsync indisponible");
    } else {
      await new Promise((resolve) => {
        mb.getCallbackTokenAsync({ isRest: true }, async (res) => {
          if (res.status !== Office.AsyncResultStatus.Succeeded || !res.value) {
            ko("Jeton REST non obtenu : " + (res.error ? res.error.message : "?"));
            resolve();
            return;
          }
          ok("Jeton REST obtenu");
          const base = (sharedCtx.targetRestUrl || "https://outlook.office.com/api").replace(/\/+$/, "");
          const headers = {
            "Authorization": "Bearer " + res.value,
            "Accept": "application/json",
            "X-AnchorMailbox": sharedCtx.targetMailbox,
          };
          // 5a. Lister les dossiers de la boîte partagée
          try {
            const r = await fetch(base + "/v2.0/me/MailFolders?$top=100&$select=Id,DisplayName", { headers: headers });
            if (r.ok) {
              const data = await r.json();
              const folders = (data.value || []);
              ok("LISTE DES DOSSIERS : OK (" + folders.length + " dossiers via REST)");
              const names = folders.map((f) => f.DisplayName).filter(Boolean);
              if (names.length) info("   Dossiers : " + names.slice(0, 15).join(" | "));
              // 5b. Lire un message
              try {
                const r2 = await fetch(base + "/v2.0/me/messages?$top=1&$select=Subject", { headers: headers });
                if (r2.ok) {
                  ok("LECTURE DES MESSAGES : OK");
                  info("   → Le déplacement via REST devrait être possible !");
                } else {
                  ko("Lecture des messages refusée (HTTP " + r2.status + ")");
                }
              } catch (e2) {
                ko("Lecture des messages bloquée : " + e2.message);
              }
            } else {
              ko("Liste des dossiers refusée (HTTP " + r.status + ")");
              if (r.status === 401 || r.status === 403) info("   → Jeton refusé ou REST API désactivée par le tenant.");
            }
          } catch (e) {
            ko("Appel REST bloqué : " + e.message);
            info("   → Cause possible : politique CORS du tenant ou REST API désactivée.");
          }
          resolve();
        });
      });
    }
    log(sep);

    // --- Synthèse ---
    log("SYNTHÈSE");
    if (ewsOk) {
      ok("Déplacement automatique par EWS : POSSIBLE sur cette boîte");
      info("→ L'Add-in peut trier et déplacer les messages directement.");
    } else {
      ko("Déplacement automatique par EWS : IMPOSSIBLE sur cette boîte");
      info("   (Confirmé par la doc Microsoft : EWS n'est pas supporté en contexte délégué)");
      info("→ Voie officielle Microsoft : Outlook REST API ou Microsoft Graph.");
    }
    log("");
    log("Fin du diagnostic — utilisez le bouton « Copier » pour me transmettre ce rapport.");

    dump();
    if (el.btn_run_diagnostic) el.btn_run_diagnostic.disabled = false;
  }

  async function copyDiagnostic() {
    const txt = el.diagnostic_output ? el.diagnostic_output.textContent : "";
    try {
      await navigator.clipboard.writeText(txt);
      showAlert("info", "✓ Rapport de diagnostic copié dans le presse-papier.");
    } catch (e) {
      showAlert("error", "Copie impossible : sélectionnez le texte manuellement.");
    }
  }

  bindEvents();
  loadFolders();
  updateAuthUI();

  if (window.Office) {
    Office.onReady((info) => {
      refreshSelection();
    });
  } else {
    setupMockEmails();
  }
})();
