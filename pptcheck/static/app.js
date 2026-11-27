/* PPTcheck dashboard — plain JavaScript, no build step required. */
(() => {
  "use strict";

  const PPT_RE = /\.(pptx|pptm|ppsx|ppsm)$/i;
  const CONCURRENCY = 2;            // files analysed in parallel
  const results = new Map();        // key (relative path) -> report
  let sortKey = "file_name", sortDir = 1;
  let cancelRequested = false, running = false;
  let currentDetail = null;

  const $ = (sel) => document.querySelector(sel);
  const esc = (v) => String(v ?? "").replace(/[&<>"']/g, (c) =>
    ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
  const fmt = (v, fallback = "—") => (v === null || v === undefined || v === "" ? fallback : v);
  const isIgnored = (name) => name.startsWith("~$") || name.startsWith("._");
  const baseName = (p) => p.split(/[\\/]/).pop();

  // ------------------------------------------------------------------ app bridge
  // All work happens in the PPTcheck app on this computer. The page talks to it
  // through pywebview's built-in bridge (window.pywebview.api) — there is no
  // web server and no network involved.
  let api = null;
  const bridgeReady = new Promise((resolve) => {
    const ready = () => { api = window.pywebview.api; resolve(); };
    if (window.pywebview && window.pywebview.api) ready();
    else window.addEventListener("pywebviewready", ready, { once: true });
  });

  function toast(message, kind = "") {
    const el = $("#toast");
    el.textContent = message;
    el.className = `toast ${kind}`;
    clearTimeout(toast._t);
    toast._t = setTimeout(() => el.classList.add("hidden"), 6000);
  }

  bridgeReady.then(async () => {
    const info = await api.info();
    const el = $("#engineStatus");
    if (info.media_available) { el.textContent = `v${info.version} · video metrics ✓`; el.className = "pill ok"; el.title = `Video engine: ${info.media_engine}`; }
    else { el.textContent = "Video engine missing — duration/FPS unavailable"; el.className = "pill bad"; }
  });

  // ------------------------------------------------------------------ input
  const recursive = () => $("#recursive").checked;

  async function addPaths(paths) {
    if (!paths || !paths.length) return;
    await bridgeReady;
    const res = await api.expand_paths(paths, recursive());
    if (res.errors && res.errors.length) toast(res.errors.join(" "), "bad");
    if (!res.jobs.length) { toast("No PowerPoint files (.pptx) found in what you chose.", "bad"); return; }
    enqueue(res.jobs.map((j) => ({ path: j.path, name: j.name })));
  }

  $("#pickFiles").onclick = async () => { await bridgeReady; addPaths(await api.choose_files()); };
  $("#pickFolder").onclick = async () => { await bridgeReady; addPaths(await api.choose_folder()); };

  // Drag-and-drop: the app receives the dropped files' locations directly
  // (see desktop.py) and calls PPTcheck.dropped(paths).
  const dz = $("#dropZone");
  let dragDepth = 0;
  window.addEventListener("dragenter", (e) => { e.preventDefault(); dragDepth++; dz.classList.add("over"); });
  window.addEventListener("dragleave", () => { if (--dragDepth <= 0) { dragDepth = 0; dz.classList.remove("over"); } });
  window.addEventListener("dragover", (e) => e.preventDefault());
  window.addEventListener("drop", (e) => { e.preventDefault(); dragDepth = 0; dz.classList.remove("over"); });

  // ------------------------------------------------------------------ queue
  const queue = [];
  const active = new Set();
  let total = 0, done = 0, activeWorkers = 0;
  const batchWaiters = [];

  function enqueue(jobs) {
    if (!jobs.length) return;
    if (!running) {
      running = true; cancelRequested = false; total = 0; done = 0;
      $("#progress").classList.remove("hidden");
    }
    queue.push(...jobs);
    total += jobs.length;
    updateProgress();
    pump();
  }

  // Keep up to CONCURRENCY workers busy; new files added mid-scan join the same batch.
  function pump() {
    while (activeWorkers < CONCURRENCY && queue.length && !cancelRequested) {
      activeWorkers++;
      worker();
    }
  }

  async function worker() {
    while (queue.length && !cancelRequested) {
      const job = queue.shift();
      active.add(job.name);
      $("#progressFile").textContent = "Analysing: " + [...active].join(", ");
      try {
        results.set(job.name, await analyse(job));
      } catch (err) {
        results.set(job.name, errorReport(job, (err && err.message) || String(err)));
      }
      active.delete(job.name);
      done++;
      updateProgress();
      renderTable();
    }
    activeWorkers--;
    if (activeWorkers === 0) finishBatch();
  }

  function finishBatch() {
    const cancelled = cancelRequested;
    queue.length = 0;
    $("#progressText").textContent = cancelled ? `Cancelled — ${done} of ${total} analysed.` : `Done — ${done} file${done === 1 ? "" : "s"} analysed.`;
    $("#progressBar").style.width = cancelled ? `${(done / total) * 100}%` : "100%";
    $("#progressFile").textContent = "";
    $("#cancelBtn").classList.add("hidden");
    running = false;
    renderTable();
    batchWaiters.splice(0).forEach((resolve) => resolve());
  }

  function updateProgress() {
    $("#cancelBtn").classList.remove("hidden");
    $("#progressText").textContent = `Analysing ${Math.min(done + 1, total)} of ${total}…`;
    $("#progressBar").style.width = total ? `${(done / total) * 100}%` : "0";
  }
  $("#cancelBtn").onclick = () => { cancelRequested = true; $("#progressText").textContent = "Cancelling after current file(s)…"; };

  async function analyse(job) {
    await bridgeReady;
    const data = await api.analyze(job.path, job.name);
    if (!data || data.bridge_error) throw new Error((data && data.bridge_error) || "No result");
    data._key = job.name;
    return data;
  }

  function errorReport(job, msg) {
    return { _key: job.name, file_name: job.name, error: msg, slide_count: 0, hidden_slide_count: 0,
      file_size_mb: null, summary: {}, videos: [], audio: [], transitions: [], animations: [],
      addins: [], fonts: [], slides: [], warnings: [msg] };
  }

  // ------------------------------------------------------------------ table
  const rowValue = (r, key) => {
    const s = r.summary || {};
    switch (key) {
      case "video_count": return s.video_count ?? 0;
      case "total_video_size_mb": return s.total_video_size_mb ?? 0;
      case "has_transitions": return s.has_transitions ? 1 : 0;
      case "has_animations": return s.has_animations ? 1 : 0;
      case "addins_text": return (s.addins || []).join(", ");
      case "fonts_issue": return (s.fonts_to_install || []).length * 1000 + (s.font_count || 0);
      default: return r[key] ?? "";
    }
  };

  function visibleRows() {
    const text = $("#filterText").value.trim().toLowerCase();
    const fv = $("#filterVideos").value, ft = $("#filterTrans").value, fa = $("#filterAddins").value, ff = $("#filterFonts").value;
    const yn = (want, has) => !want || (want === "yes") === !!has;
    const rows = [...results.values()].filter((r) => {
      const s = r.summary || {};
      if (text && !(`${r.file_name} ${(s.addins || []).join(" ")} ${(s.fonts || []).join(" ")}`.toLowerCase().includes(text))) return false;
      return yn(fv, s.video_count) && yn(ft, s.has_transitions) && yn(fa, (s.addins || []).length)
        && (!ff || yn(ff, (s.fonts_to_install || []).length));
    });
    rows.sort((a, b) => {
      const x = rowValue(a, sortKey), y = rowValue(b, sortKey);
      if (typeof x === "number" && typeof y === "number") return (x - y) * sortDir;
      return String(x).localeCompare(String(y), undefined, { numeric: true, sensitivity: "base" }) * sortDir;
    });
    return rows;
  }

  function renderTable() {
    const rows = visibleRows();
    const tbody = $("#summaryTable tbody");
    tbody.innerHTML = rows.map((r) => {
      const s = r.summary || {};
      const yes = (v) => `<span class="tag ${v ? "yes" : "no"}">${v ? "Yes" : "No"}</span>`;
      const addins = (s.addins || []).map((a) => `<span class="tag addin">${esc(a)}</span>`).join("") || '<span class="muted">—</span>';
      const toInstall = s.fonts_to_install || [];
      const fonts = r.error ? "—" : `${fmt(s.font_count, 0)} font${s.font_count === 1 ? "" : "s"}` +
        (toInstall.length ? "<br>" + toInstall.map((f) => `<span class="tag err" title="Not standard — must be installed">${esc(f)}</span>`).join("")
          : ' <span class="tag yes">all standard</span>');
      const status = r.error ? ` <span class="tag err" title="${esc(r.error)}">Error</span>` : "";
      const trans = s.has_transitions ? `${yes(true)} <span class="muted small-text">${esc((s.transition_types || []).join(", "))}</span>` : yes(false);
      return `<tr data-key="${esc(r._key)}">
        <td><div class="fname">${esc(r.file_name)}${status}</div></td>
        <td class="num">${fmt(r.slide_count)}</td>
        <td class="num">${fmt(r.hidden_slide_count)}</td>
        <td class="num">${fmt(s.video_count, 0)}</td>
        <td class="num">${fmt(s.total_video_size_mb, "0")}</td>
        <td>${r.error ? "—" : trans}</td>
        <td>${r.error ? "—" : yes(s.has_animations)}</td>
        <td>${addins}</td>
        <td>${fonts}</td>
        <td class="num">${fmt(r.file_size_mb)}</td></tr>`;
    }).join("");
    $("#emptyState").classList.toggle("hidden", results.size > 0);
    $("#countLabel").textContent = results.size ? `(${rows.length} of ${results.size} shown)` : "";
    document.querySelectorAll("#summaryTable th[data-sort]").forEach((th) => {
      th.classList.toggle("sorted-asc", th.dataset.sort === sortKey && sortDir === 1);
      th.classList.toggle("sorted-desc", th.dataset.sort === sortKey && sortDir === -1);
    });
  }

  document.querySelectorAll("#summaryTable th[data-sort]").forEach((th) => th.addEventListener("click", () => {
    if (sortKey === th.dataset.sort) sortDir = -sortDir; else { sortKey = th.dataset.sort; sortDir = 1; }
    renderTable();
  }));
  ["#filterText", "#filterVideos", "#filterTrans", "#filterAddins", "#filterFonts"].forEach((s) => $(s).addEventListener("input", renderTable));
  $("#summaryTable tbody").addEventListener("click", (e) => {
    const tr = e.target.closest("tr[data-key]");
    if (tr) openDetail(results.get(tr.dataset.key));
  });
  $("#clearBtn").onclick = () => {
    if (running) { toast("Wait for the current scan to finish (or cancel it) first.", "bad"); return; }
    if (results.size && confirm("Remove all results from the dashboard?")) { results.clear(); renderTable(); $("#progress").classList.add("hidden"); }
  };

  // ------------------------------------------------------------------ detail view
  function table(headers, rows) {
    if (!rows.length) return '<p class="none">None found.</p>';
    return `<table class="data"><thead><tr>${headers.map((h) => `<th class="${h.num ? "num" : ""}">${esc(h.label)}</th>`).join("")}</tr></thead>
      <tbody>${rows.map((r) => `<tr>${r.map((c, i) => `<td class="${headers[i].num ? "num" : ""}">${c}</td>`).join("")}</tr>`).join("")}</tbody></table>`;
  }
  // Only server-generated https:// links are rendered, opened in a new tab.
  // Links open in the computer's normal web browser (see the click handler below).
  const link = (l) => (/^https:\/\//.test(l.url)
    ? `<a href="${esc(l.url)}" class="ext" rel="noopener noreferrer">${esc(l.label)} ↗</a>` : "");
  const slideList = (arr) => (arr && arr.length ? arr.join(", ") : "—");

  function openDetail(r) {
    if (!r) return;
    currentDetail = r;
    const s = r.summary || {};
    $("#detailTitle").textContent = r.file_name;
    const stat = (label, value) => `<div class="stat"><div class="label">${esc(label)}</div><div class="value">${esc(fmt(value))}</div></div>`;
    let html = "";
    if (r.error) html += `<div class="errbox"><strong>Could not analyse:</strong> ${esc(r.error)}</div>`;
    html += `<div class="stats">
      ${stat("File size", r.file_size_mb != null ? r.file_size_mb + " MB" : null)}
      ${stat("Slides", r.slide_count)}${stat("Hidden slides", r.hidden_slide_count)}
      ${stat("Videos", s.video_count)}${stat("Total video size", s.total_video_size_mb != null ? s.total_video_size_mb + " MB" : null)}
      ${stat("Total video length", s.total_video_duration)}
      ${stat("Images", s.image_count)}${stat("Audio clips", s.audio_count)}
      ${stat("Text density", s.text_density ? `${s.text_density} (${s.avg_words_per_slide}/slide)` : null)}
      ${stat("Slides with notes", s.slides_with_notes)}
      ${stat("Transition slides", s.transition_slide_count)}${stat("Animated slides", s.animated_slide_count)}
    </div>`;
    if (s.description) html += `<div class="summary-text"><strong>Summary:</strong> ${esc(s.description)}</div>`;
    if (r.warnings && r.warnings.length && !r.error) html += `<div class="warnbox">${r.warnings.map(esc).join("<br>")}</div>`;

    // Videos
    html += `<div class="section"><h3>Videos (${(r.videos || []).length})</h3>` + table(
      [{ label: "File" }, { label: "Shape name" }, { label: "Source" }, { label: "Slide(s)" }, { label: "Size (MB)", num: 1 },
       { label: "Duration" }, { label: "FPS", num: 1 }, { label: "Resolution" }, { label: "Codec" }, { label: "Format / container" }],
      (r.videos || []).map((v) => [
        `<strong>${esc(v.file)}</strong>${v.probe_error ? `<br><span class="tag warn">${esc(v.probe_error)}</span>` : ""}${v.source !== "Embedded" ? `<div class="links">${esc(v.target)}</div>` : ""}`,
        esc(fmt(v.name)), esc(v.source), esc(slideList(v.slides)), esc(fmt(v.size_mb)), esc(fmt(v.duration)),
        esc(fmt(v.fps)), esc(fmt(v.resolution)), esc(fmt(v.codec)), esc([v.format, v.container].filter(Boolean).join(" · ") || "—"),
      ])) + "</div>";

    if ((r.audio || []).length) {
      html += `<div class="section"><h3>Audio (${r.audio.length})</h3>` + table(
        [{ label: "File" }, { label: "Source" }, { label: "Slide(s)" }, { label: "Size (MB)", num: 1 }, { label: "Duration" }, { label: "Codec" }],
        r.audio.map((a) => [esc(a.file), esc(a.source), esc(slideList(a.slides)), esc(fmt(a.size_mb)), esc(fmt(a.duration)), esc(fmt(a.codec))])) + "</div>";
    }

    // Transitions
    html += `<div class="section"><h3>Transitions — ${s.has_transitions ? `${s.transition_slide_count} slide(s)` : "none"}${s.has_auto_advance ? " · auto-advance timings set" : ""}</h3>` + table(
      [{ label: "Slide", num: 1 }, { label: "Type" }, { label: "Option" }, { label: "Speed / duration" }, { label: "Auto-advance" }, { label: "Advance on click" }, { label: "Sound" }],
      (r.transitions || []).map((t) => [t.slide, `<strong>${esc(t.type)}</strong>`, esc(fmt(t.variant)),
        esc([t.speed, t.duration_ms ? `${t.duration_ms / 1000}s` : null].filter(Boolean).join(" · ") || "—"),
        t.auto_advance ? `<span class="tag yes">After ${esc(t.auto_advance_seconds)}s</span>` : '<span class="tag no">No</span>',
        t.advance_on_click ? "Yes" : '<span class="tag warn">No</span>', esc(fmt(t.sound))])) + "</div>";

    // Animations
    html += `<div class="section"><h3>Animations — ${s.has_animations ? `${s.animation_effect_count} effect(s) on ${s.animated_slide_count} slide(s)` : "none"}</h3>` + table(
      [{ label: "Slide", num: 1 }, { label: "Total effects", num: 1 }, { label: "Entrance", num: 1 }, { label: "Emphasis", num: 1 },
       { label: "Exit", num: 1 }, { label: "Motion path", num: 1 }, { label: "Media play triggers", num: 1 }],
      (r.animations || []).map((a) => [a.slide, a.effects, a.entrance, a.emphasis, a.exit, a.motion_path, a.media_triggers])) + "</div>";

    // Add-ins
    html += `<div class="section"><h3>Add-ins &amp; interactive integrations (${(r.addins || []).length})</h3>
      <p class="muted small-text">Confidence: <span class="tag conf-high">high</span> = found in add-in/tag/OLE parts,
      <span class="tag conf-medium">medium</span> = found in a hyperlink or document metadata,
      <span class="tag conf-low">low</span> = only mentioned in slide text.</p>` + table(
      [{ label: "Add-in" }, { label: "Category" }, { label: "Location" }, { label: "Detected via" }, { label: "Confidence" }, { label: "Details" }],
      (r.addins || []).map((a) => [`<strong>${esc(a.name)}</strong>`, esc(a.category), esc(a.locations.join(", ")),
        esc(a.sources.join(", ")), `<span class="tag conf-${esc(a.confidence)}">${esc(a.confidence)}</span>`,
        `<div class="links">${a.details.map(esc).join("<br>")}</div>`])) + "</div>";

    // Fonts
    const fontsArr = r.fonts || [];
    const installCount = fontsArr.filter((f) => f.needs_install).length;
    const statusTag = (f) => {
      const cls = f.needs_install ? "err" : f.platform_warning ? "warn" : "yes";
      return `<span class="tag ${cls}" title="${esc(f.note)}">${esc(f.status_label)}</span>`;
    };
    const installedCell = (v) => (v === true ? '<span class="tag yes">Yes</span>' : v === false ? '<span class="tag no">No</span>' : '<span class="muted">unknown</span>');
    html += `<div class="section"><h3>Fonts (${fontsArr.length})${installCount ? ` — <span class="tag err">${installCount} need installing</span>` : ""}</h3>
      <p class="muted small-text">"Installed here" means installed on <em>this</em> computer, the one running PPTcheck.
      Fonts that need installing must be installed on every computer that will present this deck, or embedded
      (PowerPoint: File › Options › Save › <em>Embed fonts in the file</em>). Hover over a status for details.</p>` + table(
      [{ label: "Font" }, { label: "Status" }, { label: "Used on" }, { label: "Embedded" }, { label: "Installed here" }, { label: "Download / find" }],
      fontsArr.map((f) => [
        `<strong>${esc(f.name)}</strong>${f.family && f.family !== f.name ? `<div class="links">family: ${esc(f.family)}</div>` : ""}`,
        statusTag(f) + `<div class="links">${esc(f.note)}</div>`,
        esc(f.slides.length ? "Slides " + f.slides.join(", ") : "—") + `<div class="links">${esc(f.uses.join(", "))}</div>`,
        f.embedded ? `<span class="tag yes">Yes</span>${f.embedded_styles.length ? `<div class="links">${esc(f.embedded_styles.join(", "))}</div>` : ""}` : '<span class="tag no">No</span>',
        installedCell(f.installed_here),
        f.links.map(link).join("<br>") || '<span class="muted">—</span>',
      ])) + "</div>";

    // Slide by slide
    html += `<div class="section"><h3>Slide-by-slide findings</h3>` + table(
      [{ label: "#", num: 1 }, { label: "Title" }, { label: "Hidden" }, { label: "Words", num: 1 }, { label: "Images", num: 1 },
       { label: "Videos", num: 1 }, { label: "Audio", num: 1 }, { label: "Transition" }, { label: "Animations", num: 1 },
       { label: "Notes" }, { label: "Add-ins" }, { label: "Fonts" }, { label: "Links / objects" }],
      (r.slides || []).map((sl) => [sl.number, esc(fmt(sl.title)), sl.hidden ? '<span class="tag warn">Hidden</span>' : "",
        sl.word_count, sl.image_count, sl.video_count, sl.audio_count,
        esc(fmt(sl.transition)) + (sl.auto_advance_seconds != null ? ` <span class="muted small-text">(auto ${esc(sl.auto_advance_seconds)}s)</span>` : ""),
        sl.animation_count, sl.has_notes ? `<span title="${esc(sl.notes_preview)}" class="tag yes">Yes</span>` : '<span class="tag no">No</span>',
        sl.addins.map((a) => `<span class="tag addin">${esc(a)}</span>`).join("") || "—",
        `<div class="links">${(sl.fonts || []).map(esc).join("<br>") || "—"}</div>`,
        `<div class="links">${[...sl.hyperlinks, ...sl.ole_objects].map(esc).join("<br>") || "—"}</div>`])) + "</div>";

    $("#detailBody").innerHTML = html;
    $("#detail").classList.remove("hidden");
    $("#detailBody").scrollTop = 0;
  }
  const closeDetail = () => { $("#detail").classList.add("hidden"); currentDetail = null; };
  document.querySelectorAll("[data-close]").forEach((el) => el.addEventListener("click", closeDetail));
  document.addEventListener("keydown", (e) => { if (e.key === "Escape") closeDetail(); });

  // ------------------------------------------------------------------ export
  // The app asks where to save (normal "Save as" window), writes the PDF
  // summary report on this computer, then opens it.
  const clean = (r) => { const { _key, ...rest } = r; return rest; };
  const safeName = (n) => baseName(n).replace(/\.[^.]+$/, "").replace(/[\\/:*?"<>|]+/g, "_").trim() || "presentation";
  const today = () => new Date().toISOString().slice(0, 10);

  async function exportPdf(reports, suggestedName) {
    if (!reports.length) { toast("Nothing to export yet — add some presentations first.", "bad"); return; }
    await bridgeReady;
    toast("Creating PDF report…");
    const res = await api.export_pdf(reports.map(clean), suggestedName);
    if (res && res.saved) toast(`Saved and opened: ${res.saved}`, "ok");
    else if (res && res.error) toast(`Could not create the PDF: ${res.error}`, "bad");
    else $("#toast").classList.add("hidden");   // Save window was cancelled
  }

  $("#exportPdf").onclick = () => {
    const rows = visibleRows();
    const name = rows.length === 1 ? `${safeName(rows[0].file_name)} - PPTcheck report.pdf` : `PPTcheck report ${today()}.pdf`;
    exportPdf(rows, name);
  };
  $("#detailPdf").onclick = () => currentDetail &&
    exportPdf([currentDetail], `${safeName(currentDetail.file_name)} - PPTcheck report.pdf`);

  document.addEventListener("click", async (e) => {
    const a = e.target.closest("a.ext");
    if (!a) return;
    e.preventDefault();
    await bridgeReady;
    api.open_url(a.getAttribute("href"));
  });

  // ------------------------------------------------------------------ called by the app
  window.PPTcheck = {
    // Files/folders dropped on the window (full locations supplied by the app).
    dropped: (paths) => addPaths(paths),

    // Automated check used when building the app: analyse a file through the
    // real interface, open its detail view, and report what was shown.
    async selfTest(path) {
      await bridgeReady;
      const finished = new Promise((resolve) => batchWaiters.push(resolve));
      await addPaths([path]);
      await finished;
      const row = document.querySelector("#summaryTable tbody tr");
      if (row) row.click();
      const text = (sel) => (document.querySelector(sel) || {}).innerText || "";
      await api.selftest_done({
        rows: document.querySelectorAll("#summaryTable tbody tr").length,
        table: text("#summaryTable"),
        detail: text("#detailBody"),
        engine: text("#engineStatus"),
        pdf_written: await api.selftest_export_pdf(visibleRows().map(clean)),
      });
    },
  };

  renderTable();
})();
