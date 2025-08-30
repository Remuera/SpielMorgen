/* globals XLSX */

(() => {
  "use strict";

  // DOM
  const form = document.getElementById("config-form");
  const inputGameCount = document.getElementById("input-game-count");
  const inputCapacity = document.getElementById("input-capacity");
  const btnReset = document.getElementById("btn-reset");
  const messages = document.getElementById("messages");
  const resultsSection = document.getElementById("results-section");
  const table = document.getElementById("result-table");
  const tbody = table.querySelector("tbody");

  const filterVorname = document.getElementById("filter-vorname");
  const filterNachname = document.getElementById("filter-nachname");
  const filterBlock = document.getElementById("filter-block");
  const filterSpiel = document.getElementById("filter-spiel");

  // Zustand
  let assignments = []; // {vorname, nachname, block, spiel, repeated:boolean}
  let sortState = { key: "__default__", dir: "asc" };

  // Utils
  const html = (strings, ...vals) =>
    strings.map((s, i) => s + (vals[i] ?? "")).join("");

  function showMessage(type, text) {
    const div = document.createElement("div");
    div.className = `msg ${type}`;
    div.textContent = text;
    messages.appendChild(div);
  }
  function clearMessages() {
    messages.innerHTML = "";
  }
  function clamp(n, min, max) {
    return Math.max(min, Math.min(max, n));
  }
  function shuffle(arr) {
    const a = arr.slice();
    for (let i = a.length - 1; i > 0; i--) {
      const j = Math.floor(Math.random() * (i + 1));
      [a[i], a[j]] = [a[j], a[i]];
    }
    return a;
  }
  function escapeHtml(s) {
    return String(s)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;");
  }

  // Excel laden (erst .xlsx, dann .xls)
  async function loadExcel() {
    const tryFiles = ["Spielpraeferenzen.xlsx", "Spielpraeferenzen.xls"];
    let lastErr = null;
    for (const name of tryFiles) {
      try {
        const res = await fetch(name);
        if (!res.ok) throw new Error(`${name}: HTTP ${res.status}`);
        const buf = await res.arrayBuffer();
        const wb = XLSX.read(buf, { type: "array" });
        const sheet = wb.Sheets[wb.SheetNames[0]];
        const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });
        return rows;
      } catch (e) {
        lastErr = e;
      }
    }
    throw new Error(
      `Excel-Datei nicht gefunden oder nicht lesbar. Erwarte 'Spielpraeferenzen.xlsx' (oder .xls) im selben Verzeichnis. ${lastErr ? "(" + lastErr.message + ")" : ""}`
    );
  }

  // Excel-Daten parsen
  function parseChildren(rows) {
    // rows[0] = Header; relevante Spalten: F (Index 5), G (6), H (7)
    const children = [];
    for (let i = 1; i < rows.length; i++) {
      const r = rows[i] || [];
      const prefsRaw = (r[5] || "").toString();
      const vorname = (r[6] || "").toString().trim();
      const nachname = (r[7] || "").toString().trim();
      if (!vorname || !nachname) continue;
      const prefs = prefsRaw
        .split(";")
        .map((s) => s.trim())
        .filter((s) => s.length > 0);
      children.push({ vorname, nachname, prefs });
    }
    return children;
  }

  // Popularität der Spiele (Gewichtete Punkte je Präferenzrang)
  function computePopularity(children) {
    const score = new Map();
    for (const c of children) {
      c.prefs.forEach((g, idx) => {
        const w = idx === 0 ? 5 : idx === 1 ? 3 : idx === 2 ? 2 : 1;
        score.set(g, (score.get(g) || 0) + w);
      });
    }
    const allGames = Array.from(score.entries())
      .sort((a, b) => b[1] - a[1])
      .map(([g]) => g);
    return allGames;
  }

  // Zuteilungs-Algorithmus
  function allocate(children, gamePool, N, capacity, blocks = 4) {
    const K = children.length;
    const minHalf = Math.ceil(capacity / 2);

    if (N < 1) throw new Error("Anzahl Spiele muss ≥ 1 sein.");
    if (capacity < 1) throw new Error("Kapazität muss ≥ 1 sein.");
    if (K === 0) throw new Error("Keine Kinder gefunden.");

    const distinctPool = Array.from(new Set(gamePool));
    if (distinctPool.length === 0) {
      throw new Error("Keine Spielnamen in den Präferenzen gefunden.");
    }

    const games = distinctPool.slice(0, N);
    if (games.length < 4) {
      showMessage(
        "warn",
        "Es sind insgesamt weniger als 4 unterschiedliche Spiele verfügbar. Wiederholungen sind ggf. unvermeidbar."
      );
    }

    // Anzahl aktive Spiele pro Block (M) nach Kapazität und Mindestbelegung
    // M_min: M*capacity >= K  -> M >= ceil(K/capacity)
    // M_max: K >= M*minHalf  -> M <= floor(K/minHalf)
    const M_min = Math.ceil(K / capacity);
    const M_max = Math.min(games.length, Math.floor(K / minHalf));
    let M;
    if (M_min <= M_max) {
      M = clamp(M_min, 1, M_max);
    } else {
      M = Math.min(games.length, M_min);
      showMessage(
        "warn",
        "Mindestbelegung pro Spiel lässt sich nicht überall einhalten. Spiele werden zusammengelegt oder ausgelassen."
      );
    }

    const strictNoRepeatPossible = games.length >= 4;

    const playedByChild = new Map(); // key -> Set(spiele)
    const keyOf = (c) => `${c.vorname}|||${c.nachname}`;

    // blockweise rotierende Auswahl
    const rotatedBlockGames = (b) => {
      if (games.length <= M) return games.slice();
      const offset = ((b - 1) * M) % games.length;
      const cycle = games.slice(offset).concat(games.slice(0, offset));
      return cycle.slice(0, M);
    };

    const out = [];

    for (let block = 1; block <= blocks; block++) {
      let activeGames = rotatedBlockGames(block);

      // Primäre Zuordnung nach Präferenzen
      const cap = new Map(activeGames.map((g) => [g, capacity]));
      const assigned = new Map(); // childKey -> spiel
      const order = shuffle(children);

      const maxRanks = Math.max(1, ...children.map((c) => c.prefs.length || 0));
      for (let rank = 0; rank < maxRanks; rank++) {
        for (const c of order) {
          const ck = keyOf(c);
          if (assigned.has(ck)) continue;
          const already = playedByChild.get(ck) || new Set();
          const candidate = c.prefs.find(
            (g) => activeGames.includes(g) && !already.has(g)
          );
          if (!candidate) continue;
          if ((cap.get(candidate) || 0) > 0) {
            assigned.set(ck, candidate);
            cap.set(candidate, cap.get(candidate) - 1);
          }
        }
      }

      // Rest auf verfügbare Spiele verteilen (ohne Wiederholung bevorzugt)
      for (const c of order) {
        const ck = keyOf(c);
        if (assigned.has(ck)) continue;
        const already = playedByChild.get(ck) || new Set();
        let placed = false;
        for (const g of activeGames) {
          if (already.has(g)) continue;
          if ((cap.get(g) || 0) > 0) {
            assigned.set(ck, g);
            cap.set(g, cap.get(g) - 1);
            placed = true;
            break;
          }
        }
        if (!placed) {
          // notfalls Wiederholung
          const allowRepeat = !strictNoRepeatPossible || activeGames.length < 4;
          if (allowRepeat) {
            for (const g of activeGames) {
              if ((cap.get(g) || 0) > 0) {
                assigned.set(ck, g);
                cap.set(g, cap.get(g) - 1);
                break;
              }
            }
          }
        }
      }

      // Mindestbelegung prüfen, zu kleine Spiele auslassen und umverteilen
      const assignedByGame = new Map(activeGames.map((g) => [g, []]));
      for (const c of order) {
        const g = assigned.get(keyOf(c));
        if (g && assignedByGame.has(g)) assignedByGame.get(g).push(c);
      }

      const tooSmall = activeGames.filter(
        (g) => assignedByGame.get(g).length > 0 && assignedByGame.get(g).length < minHalf
      );

      if (tooSmall.length > 0) {
        showMessage(
          "info",
          `Block ${block}: Spiele mit zu geringer Nachfrage werden ausgelassen.`
        );
        activeGames = activeGames.filter((g) => !tooSmall.includes(g));

        const cap2 = new Map(activeGames.map((g) => [g, capacity]));
        for (const g of activeGames) {
          for (const c of assignedByGame.get(g) || []) {
            cap2.set(g, cap2.get(g) - 1);
          }
        }

        const toReassign = tooSmall.flatMap((g) => assignedByGame.get(g) || []);
        for (const c of toReassign) assigned.delete(keyOf(c));

        for (const c of toReassign) {
          const ck = keyOf(c);
          const already = playedByChild.get(ck) || new Set();
          let placed = false;

          for (const g of c.prefs) {
            if (!activeGames.includes(g)) continue;
            if (already.has(g)) continue;
            if ((cap2.get(g) || 0) > 0) {
              assigned.set(ck, g);
              cap2.set(g, cap2.get(g) - 1);
              placed = true;
              break;
            }
          }
          if (!placed) {
            let choice = null;
            for (const g of activeGames) {
              if ((cap2.get(g) || 0) > 0 && !already.has(g)) {
                choice = g;
                break;
              }
            }
            if (!choice) {
              for (const g of activeGames) {
                if ((cap2.get(g) || 0) > 0) {
                  choice = g;
                  break;
                }
              }
            }
            if (choice) {
              assigned.set(ck, choice);
              cap2.set(choice, cap2.get(choice) - 1);
            }
          }
        }
      }

      // Ergebnisse sammeln und Historie aktualisieren
      for (const c of children) {
        const ck = keyOf(c);
        const g = assigned.get(ck);
        const already = playedByChild.get(ck) || new Set();
        const repeated = !!(g && already.has(g));
        if (g) {
          out.push({
            vorname: c.vorname,
            nachname: c.nachname,
            block,
            spiel: g,
            repeated,
          });
          already.add(g);
          playedByChild.set(ck, already);
        } else {
          out.push({
            vorname: c.vorname,
            nachname: c.nachname,
            block,
            spiel: "—",
            repeated: false,
          });
        }
      }
    }

    return out;
  }

  // Sortierung
  function defaultSort(a, b) {
    const av = a.vorname.localeCompare(b.vorname, "de", { sensitivity: "base" });
    if (av !== 0) return av;
    const an = a.nachname.localeCompare(b.nachname, "de", { sensitivity: "base" });
    if (an !== 0) return an;
    return a.block - b.block;
  }
  function applySort(data, key, dir) {
    const mult = dir === "asc" ? 1 : -1;
    const collator = new Intl.Collator("de", { numeric: true, sensitivity: "base" });
    return data.slice().sort((a, b) => {
      if (key === "block") return (a.block - b.block) * mult;
      return collator.compare(a[key], b[key]) * mult;
    });
  }

  // Rendern
  function renderTable(data) {
    const fv = filterVorname.value.trim().toLowerCase();
    const fn = filterNachname.value.trim().toLowerCase();
    const fb = filterBlock.value.trim().toLowerCase();
    const fs = filterSpiel.value.trim().toLowerCase();

    let rows = data.filter((r) => {
      const m1 = r.vorname.toLowerCase().includes(fv);
      const m2 = r.nachname.toLowerCase().includes(fn);
      const m3 = String(r.block).toLowerCase().includes(fb);
      const m4 = r.spiel.toLowerCase().includes(fs);
      return m1 && m2 && m3 && m4;
    });

    rows = sortState.key === "__default__"
      ? rows.slice().sort(defaultSort)
      : applySort(rows, sortState.key, sortState.dir);

    tbody.innerHTML = "";
    for (const r of rows) {
      const tr = document.createElement("tr");
      if (r.spiel === "—") tr.classList.add("row-warn");
      if (r.repeated) tr.classList.add("row-repeat");

      tr.innerHTML = html`
        <td>${escapeHtml(r.vorname)}</td>
        <td>${escapeHtml(r.nachname)}</td>
        <td>${r.block}</td>
        <td>${escapeHtml(r.spiel)}</td>
      `;
      tbody.appendChild(tr);
    }
  }

  // Events
  form.addEventListener("submit", async (e) => {
    e.preventDefault();
    clearMessages();
    resultsSection.classList.add("hidden");
    tbody.innerHTML = "";

    const N = parseInt(inputGameCount.value, 10);
    const capacity = parseInt(inputCapacity.value, 10);

    if (!Number.isFinite(N) || N < 1) {
      showMessage("error", "Bitte eine gültige Anzahl Spiele ≥ 1 eingeben.");
      return;
    }
    if (!Number.isFinite(capacity) || capacity < 1) {
      showMessage("error", "Bitte eine gültige Kapazität ≥ 1 eingeben.");
      return;
    }

    try {
      showMessage("info", "Excel wird geladen und Zuteilung berechnet...");
      const rows = await loadExcel();
      const children = parseChildren(rows);

      if (children.length === 0) {
        showMessage("error", "Keine gültigen Kinder-Einträge in der Excel-Datei gefunden.");
        return;
      }

      const gamePool = computePopularity(children);
      if (gamePool.length === 0) {
        showMessage("error", "Keine Spielnamen in den Präferenzen gefunden.");
        return;
      }

      assignments = allocate(children, gamePool, N, capacity, 4);
      sortState = { key: "__default__", dir: "asc" };
      renderTable(assignments);
      resultsSection.classList.remove("hidden");
      showMessage(
        "ok",
        `Zuteilung erstellt: ${assignments.length} Einträge (Kinder × 4 Blöcke).`
      );
    } catch (err) {
      console.error(err);
      showMessage("error", (err && err.message) ? err.message : String(err));
    }
  });

  btnReset.addEventListener("click", () => {
    form.reset();
    clearMessages();
    resultsSection.classList.add("hidden");
    tbody.innerHTML = "";
    assignments = [];
    sortState = { key: "__default__", dir: "asc" };
    [filterVorname, filterNachname, filterBlock, filterSpiel].forEach((i) => (i.value = ""));
  });

  [filterVorname, filterNachname, filterBlock, filterSpiel].forEach((inp) => {
    inp.addEventListener("input", () => renderTable(assignments));
  });

  // Sortierung per Klick
  table.querySelectorAll("th[data-key]").forEach((th) => {
    th.style.cursor = "pointer";
    th.addEventListener("click", () => {
      const key = th.dataset.key;
      if (!key) return;
      if (sortState.key === key) {
        sortState.dir = sortState.dir === "asc" ? "desc" : "asc";
      } else {
        sortState = { key, dir: "asc" };
      }
      renderTable(assignments);
    });
  });
})();
