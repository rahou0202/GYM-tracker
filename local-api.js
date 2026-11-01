// Phone version: replaces the PC server. Everything is stored on the phone (localStorage),
// and window.localApi answers the same /api/... calls the page makes to the PC server.
(function () {
  const SETUP_KEY = "gt.setup", SETS_KEY = "gt.sets", BACKUP_KEY = "gt.lastBackup";

  const DEFAULT_SETUP = {
    unit: "kg",
    days: [
      { name: "Arms", exercises: [
        { name: "Barbell Curl", sets: 3 }, { name: "Hammer Curl", sets: 3 },
        { name: "Tricep Pushdown", sets: 3 }, { name: "Overhead Tricep Extension", sets: 3 } ] },
      { name: "Legs", exercises: [
        { name: "Squat", sets: 4 }, { name: "Romanian Deadlift", sets: 3 }, { name: "Leg Press", sets: 3 },
        { name: "Leg Curl", sets: 3 }, { name: "Calf Raise", sets: 4 } ] },
    ],
  };

  const read = (key, fallback) => {
    try { return JSON.parse(localStorage.getItem(key)) ?? fallback; } catch { return fallback; }
  };
  const write = (key, value) => localStorage.setItem(key, JSON.stringify(value));
  const loadSetup = () => read(SETUP_KEY, null) || structuredClone(DEFAULT_SETUP);
  const loadSets = () => read(SETS_KEY, []);
  // keep sets oldest first; sort is stable, so exercise/set order inside a workout is kept
  const saveSets = (rows) => write(SETS_KEY, rows.sort((a, b) => a.date.localeCompare(b.date)));

  // ask iOS not to clear the app's storage
  navigator.storage?.persist?.();

  function num(v) {
    if (v === null || v === undefined || v === "") return null;
    const f = parseFloat(String(v).replace(",", "."));
    return Number.isFinite(f) ? f : null;
  }

  function weekName(iso) {
    const [y, m, d] = iso.split("-").map(Number);
    const t = new Date(y, m - 1, d);
    t.setDate(t.getDate() + 3 - ((t.getDay() + 6) % 7));
    const jan4 = new Date(t.getFullYear(), 0, 4);
    const week = 1 + Math.round(((t - jan4) / 864e5 - 3 + ((jan4.getDay() + 6) % 7)) / 7);
    return `week ${week}`;
  }

  function saveWorkout({ date, day, exercises }) {
    const fresh = [];
    for (const ex of exercises) {
      let n = 0;
      for (const s of ex.sets) {
        if ((s.weight ?? "") === "" && (s.reps ?? "") === "") continue; // empty set row
        fresh.push({ date, day, exercise: ex.name, set: ++n, weight: num(s.weight), reps: num(s.reps), failure: !!s.failure });
      }
    }
    saveSets(loadSets().filter((r) => !(r.date === date && r.day === day)).concat(fresh));
    return { sheet: weekName(date), sets: fresh.length };
  }

  // sets already saved for (date, day), plus the most recent earlier sets of each exercise
  function history(date, day) {
    const existing = {}, last = {};
    for (const r of loadSets()) {
      const s = { weight: r.weight, reps: r.reps, failure: r.failure };
      if (r.date === date && r.day === day) {
        (existing[r.exercise] ??= []).push(s);
      } else if (r.date <= date) {
        const cur = last[r.exercise];
        if (!cur || r.date > cur.date) last[r.exercise] = { date: r.date, sets: [s] };
        else if (r.date === cur.date) cur.sets.push(s);
      }
    }
    return { existing: Object.keys(existing).length ? existing : null, last };
  }

  window.localApi = async function (path, body) {
    const url = new URL(path, location.href);
    switch (url.pathname.split("/").pop()) {
      case "setup":
        if (body) { write(SETUP_KEY, body); return { ok: true }; }
        return loadSetup();
      case "workout": return saveWorkout(body);
      case "history": return history(url.searchParams.get("date"), url.searchParams.get("day"));
      case "log": return loadSets();
    }
    throw new Error("Unknown request " + path);
  };

  // ---------- backup / restore as workouts.xlsx ----------

  const HEADERS = ["Date", "Day", "Exercise", "Set", "Weight", "Reps", "To Failure"];

  function loadXlsxLib() {
    if (window.XLSX) return Promise.resolve();
    return new Promise((resolve, reject) => {
      const s = document.createElement("script");
      s.src = "xlsx.full.min.js";
      s.onload = resolve;
      s.onerror = () => reject(new Error("Could not load the Excel tool"));
      document.head.appendChild(s);
    });
  }

  const isoOf = (v) => {
    if (v instanceof Date) return `${v.getFullYear()}-${String(v.getMonth() + 1).padStart(2, "0")}-${String(v.getDate()).padStart(2, "0")}`;
    const s = String(v ?? "").trim();
    return /^\d{4}-\d{2}-\d{2}/.test(s) ? s.slice(0, 10) : null;
  };

  async function backup() {
    await loadXlsxLib();
    const setup = loadSetup();
    const headers = HEADERS.map((h) => (h === "Weight" ? `Weight (${setup.unit || "kg"})` : h));
    const sets = loadSets().map((r) => [r.date, r.day, r.exercise, r.set, r.weight, r.reps, r.failure ? "Yes" : ""]);
    const split = setup.days.flatMap((d) => d.exercises.map((e) => [d.name, e.name, e.sets]));

    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.aoa_to_sheet([headers, ...sets]);
    ws["!cols"] = [{ wch: 12 }, { wch: 14 }, { wch: 28 }, { wch: 5 }, { wch: 12 }, { wch: 6 }, { wch: 10 }];
    XLSX.utils.book_append_sheet(wb, ws, "All Sets");
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([["Day", "Exercise", "Sets", "Unit"], ...split.map((r, i) => [...r, i ? "" : setup.unit])]), "Split");
    const data = XLSX.write(wb, { bookType: "xlsx", type: "array" });
    const file = new File([data], "workouts.xlsx", { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });

    if (navigator.canShare?.({ files: [file] })) {
      try {
        await navigator.share({ files: [file], title: "Gym Tracker backup" }); // iPhone: "Save to Files"
      } catch (err) {
        if (err.name === "AbortError") return { cancelled: true };
        throw err;
      }
    } else {
      const a = document.createElement("a");
      a.href = URL.createObjectURL(file);
      a.download = file.name;
      a.click();
      setTimeout(() => URL.revokeObjectURL(a.href), 10000);
    }
    write(BACKUP_KEY, new Date().toISOString());
    return { sets: sets.length };
  }

  // Merge a backup (or a PC workouts.xlsx) into the phone. Workouts on the same date + day are replaced.
  async function restore(file) {
    await loadXlsxLib();
    const wb = XLSX.read(await file.arrayBuffer(), { cellDates: true });
    const sheets = wb.SheetNames.includes("All Sets")
      ? ["All Sets"]
      : wb.SheetNames.filter((n) => String(wb.Sheets[n].A1?.v) === "Date"); // older PC files: one sheet per week
    const rows = [];
    for (const name of sheets) {
      for (const r of XLSX.utils.sheet_to_json(wb.Sheets[name], { header: 1, raw: true }).slice(1)) {
        const date = isoOf(r[0]);
        if (!date || !r[2]) continue;
        rows.push({ date, day: String(r[1] ?? ""), exercise: String(r[2]), set: num(r[3]), weight: num(r[4]), reps: num(r[5]),
                    failure: String(r[6] ?? "").trim().toLowerCase() === "yes" });
      }
    }
    let split = null;
    if (wb.Sheets.Split) {
      const lines = XLSX.utils.sheet_to_json(wb.Sheets.Split, { header: 1, raw: true }).slice(1).filter((r) => r[0] && r[1]);
      if (lines.length) {
        const days = [];
        for (const [day, ex, sets] of lines) {
          let d = days.find((x) => x.name === String(day));
          if (!d) days.push((d = { name: String(day), exercises: [] }));
          d.exercises.push({ name: String(ex), sets: Math.max(1, parseInt(sets) || 3) });
        }
        split = { unit: String(lines[0][3] || "kg"), days };
      }
    }
    const workouts = new Set(rows.map((r) => r.date + "|" + r.day));
    return {
      workouts: workouts.size,
      split,
      apply(withSplit) {
        saveSets(loadSets().filter((r) => !workouts.has(r.date + "|" + r.day)).concat(rows));
        if (withSplit && split) write(SETUP_KEY, split);
      },
    };
  }

  window.localApi.backup = backup;
  window.localApi.restore = restore;
  window.localApi.lastBackup = () => read(BACKUP_KEY, null);
  window.localApi.workoutCount = () => new Set(loadSets().map((r) => r.date + "|" + r.day)).size;
})();
