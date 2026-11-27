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

  // ------------------------------------------------------------------ health
  fetch("/api/health").then((r) => r.json()).then((h) => {
    const el = $("#ffprobeStatus");
    if (h.ffprobe) { el.textContent = "FFmpeg ✓ video metrics enabled"; el.className = "pill ok"; el.title = h.ffprobe_version || ""; }
    else { el.textContent = "FFmpeg not found — video duration/FPS unavailable"; el.className = "pill bad"; }
  }).catch(() => { $("#ffprobeStatus").textContent = "Server not reachable"; $("#ffprobeStatus").className = "pill bad"; });

  // ------------------------------------------------------------------ input
  const recursive = () => $("#recursive").checked;
  $("#pickFiles").onclick = () => $("#fileInput").click();
  $("#pickFolder").onclick = () => $("#folderInput").click();

  $("#fileInput").onchange = (e) => {
    const jobs = [...e.target.files].filter((f) => PPT_RE.test(f.name) && !isIgnored(f.name))
      .map((f) => ({ kind: "upload", file: f, name: f.name }));
    e.target.value = "";
    enqueue(jobs);
  };
  $("#folderInput").onchange = (e) => {
    const jobs = [];
    for (const f of e.target.files) {
      const rel = f.webkitRelativePath || f.name;
      const parts = rel.split("/");
      if (!PPT_RE.test(f.name) || isIgnored(f.name)) continue;
      if (parts.some((p, i) => i < parts.length - 1 && p.startsWith("."))) continue;
      if (!recursive() && parts.length > 2) continue;   // top level only
      jobs.push({ kind: "upload", file: f, name: rel });
    }
    e.target.value = "";
    if (!jobs.length) alert("No .pptx files found in that folder.");
    enqueue(jobs);
  };

  const dz = $("#dropZone");
  ["dragenter", "dragover"].forEach((ev) => dz.addEventListener(ev, (e) => { e.preventDefault(); dz.classList.add("over"); }));
  ["dragleave", "drop"].forEach((ev) => dz.addEventListener(ev, () => dz.classList.remove("over")));
  // Stop the browser from opening files dropped outside the zone.
  window.addEventListener("dragover", (e) => e.preventDefault());
  window.addEventListener("drop", (e) => e.preventDefault());
  dz.addEventListener("drop", async (e) => {
    e.preventDefault();
    const items = [...(e.dataTransfer.items || [])];
    const entries = items.map((it) => it.webkitGetAsEntry && it.webkitGetAsEntry()).filter(Boolean);
    const jobs = [];
    if (entries.length) {
      for (const entry of entries) await walkEntry(entry, "", 0, jobs);
    } else {
      for (const f of e.dataTransfer.files) if (PPT_RE.test(f.name) && !isIgnored(f.name)) jobs.push({ kind: "upload", file: f, name: f.name });
    }
    if (!jobs.length) alert("No .pptx files found in what you dropped.");
    enqueue(jobs);
  });

  async function walkEntry(entry, prefix, depth, jobs) {
    const path = prefix ? `${prefix}/${entry.name}` : entry.name;
    if (entry.isFile) {
      if (PPT_RE.test(entry.name) && !isIgnored(entry.name)) {
        const file = await new Promise((res, rej) => entry.file(res, rej));
        jobs.push({ kind: "upload", file, name: path });
      }
      return;
    }
    if (!entry.isDirectory || entry.name.startsWith(".")) return;
    if (depth >= 1 && !recursive()) return;
    const reader = entry.createReader();
    // readEntries returns results in chunks; keep reading until empty.
    for (;;) {
      const batch = await new Promise((res, rej) => reader.readEntries(res, rej));
      if (!batch.length) break;
      for (const child of batch) await walkEntry(child, path, depth + 1, jobs);
    }
  }

  $("#scanPath").onclick = scanPath;
  $("#folderPath").addEventListener("keydown", (e) => { if (e.key === "Enter") scanPath(); });
  async function scanPath() {
    const path = $("#folderPath").value.trim();
    if (!path) return;
    try {
      const r = await fetch(`/api/list-folder?path=${encodeURIComponent(path)}&recursive=${recursive()}`);
      const data = await r.json();
      if (!r.ok) throw new Error(data.detail || r.statusText);
      if (!data.files.length) { alert("No .pptx files found in that folder."); return; }
      const folderName = baseName(data.folder.replace(/[\\/]+$/, ""));
      enqueue(data.files.map((f) => ({ kind: "path", path: f.path, name: `${folderName}/${f.relative_path.replace(/\\/g, "/")}` })));
    } catch (err) { alert("Could not scan folder: " + err.message); }
  }

  // ------------------------------------------------------------------ queue
  const queue = [];
  const active = new Set();
  let total = 0, done = 0, activeWorkers = 0;

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

  // Keep up to CONCURRENCY workers busy; new files dropped mid-scan join the same batch.
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
        results.set(job.name, errorReport(job, err.message));
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
  }

  function updateProgress() {
    $("#cancelBtn").classList.remove("hidden");
    $("#progressText").textContent = `Analysing ${Math.min(done + 1, total)} of ${total}…`;
    $("#progressBar").style.width = total ? `${(done / total) * 100}%` : "0";
  }
  $("#cancelBtn").onclick = () => { cancelRequested = true; $("#progressText").textContent = "Cancelling after current file(s)…"; };

  async function analyse(job) {
    let r;
    if (job.kind === "upload") {
      const fd = new FormData();
      fd.append("file", job.file, job.file.name);
      fd.append("relative_path", job.name);
      r = await fetch("/api/analyze", { method: "POST", body: fd });
    } else {
      r = await fetch("/api/analyze-path", { method: "POST", headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ path: job.path, display_name: job.name }) });
    }
    const data = await r.json().catch(() => ({}));
    if (!r.ok) throw new Error(data.detail || `HTTP ${r.status}`);
    data._key = job.name;
    return data;
  }

  function errorReport(job, msg) {
    return { _key: job.name, file_name: job.name, error: msg, slide_count: 0, hidden_slide_count: 0,
      file_size_mb: job.file ? +(job.file.size / 1048576).toFixed(2) : null, summary: {}, videos: [], audio: [],
      transitions: [], animations: [], addins: [], slides: [], warnings: [msg] };
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
      default: return r[key] ?? "";
    }
  };

  function visibleRows() {
    const text = $("#filterText").value.trim().toLowerCase();
    const fv = $("#filterVideos").value, ft = $("#filterTrans").value, fa = $("#filterAddins").value;
    const yn = (want, has) => !want || (want === "yes") === !!has;
    const rows = [...results.values()].filter((r) => {
      const s = r.summary || {};
      if (text && !(`${r.file_name} ${(s.addins || []).join(" ")}`.toLowerCase().includes(text))) return false;
      return yn(fv, s.video_count) && yn(ft, s.has_transitions) && yn(fa, (s.addins || []).length);
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
  ["#filterText", "#filterVideos", "#filterTrans", "#filterAddins"].forEach((s) => $(s).addEventListener("input", renderTable));
  $("#summaryTable tbody").addEventListener("click", (e) => {
    const tr = e.target.closest("tr[data-key]");
    if (tr) openDetail(results.get(tr.dataset.key));
  });
  $("#clearBtn").onclick = () => {
    if (running) { alert("Wait for the current scan to finish (or cancel it) first."); return; }
    if (results.size && confirm("Remove all results from the dashboard?")) { results.clear(); renderTable(); $("#progress").classList.add("hidden"); }
  };

  // ------------------------------------------------------------------ detail view
  function table(headers, rows) {
    if (!rows.length) return '<p class="none">None found.</p>';
    return `<table class="data"><thead><tr>${headers.map((h) => `<th class="${h.num ? "num" : ""}">${esc(h.label)}</th>`).join("")}</tr></thead>
      <tbody>${rows.map((r) => `<tr>${r.map((c, i) => `<td class="${headers[i].num ? "num" : ""}">${c}</td>`).join("")}</tr>`).join("")}</tbody></table>`;
  }
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

    // Slide by slide
    html += `<div class="section"><h3>Slide-by-slide findings</h3>` + table(
      [{ label: "#", num: 1 }, { label: "Title" }, { label: "Hidden" }, { label: "Words", num: 1 }, { label: "Images", num: 1 },
       { label: "Videos", num: 1 }, { label: "Audio", num: 1 }, { label: "Transition" }, { label: "Animations", num: 1 },
       { label: "Notes" }, { label: "Add-ins" }, { label: "Links / objects" }],
      (r.slides || []).map((sl) => [sl.number, esc(fmt(sl.title)), sl.hidden ? '<span class="tag warn">Hidden</span>' : "",
        sl.word_count, sl.image_count, sl.video_count, sl.audio_count,
        esc(fmt(sl.transition)) + (sl.auto_advance_seconds != null ? ` <span class="muted small-text">(auto ${esc(sl.auto_advance_seconds)}s)</span>` : ""),
        sl.animation_count, sl.has_notes ? `<span title="${esc(sl.notes_preview)}" class="tag yes">Yes</span>` : '<span class="tag no">No</span>',
        sl.addins.map((a) => `<span class="tag addin">${esc(a)}</span>`).join("") || "—",
        `<div class="links">${[...sl.hyperlinks, ...sl.ole_objects].map(esc).join("<br>") || "—"}</div>`])) + "</div>";

    $("#detailBody").innerHTML = html;
    $("#detail").classList.remove("hidden");
    $("#detailBody").scrollTop = 0;
  }
  const closeDetail = () => { $("#detail").classList.add("hidden"); currentDetail = null; };
  document.querySelectorAll("[data-close]").forEach((el) => el.addEventListener("click", closeDetail));
  document.addEventListener("keydown", (e) => { if (e.key === "Escape") closeDetail(); });

  // ------------------------------------------------------------------ export
  function csvCell(v) {
    if (v === null || v === undefined) return "";
    let s = Array.isArray(v) ? v.join("; ") : String(v);
    if (/^[=+\-@\t\r]/.test(s)) s = "'" + s;          // stop spreadsheet formula injection
    return /[",\n\r]/.test(s) ? `"${s.replace(/"/g, '""')}"` : s;
  }
  const toCsv = (headers, rows) => "﻿" + [headers, ...rows].map((r) => r.map(csvCell).join(",")).join("\r\n");
  function download(name, content, type) {
    const url = URL.createObjectURL(new Blob([content], { type }));
    const a = Object.assign(document.createElement("a"), { href: url, download: name });
    document.body.appendChild(a); a.click(); a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 1000);
  }
  const stamp = () => new Date().toISOString().slice(0, 19).replace(/[:T]/g, "-");
  const clean = (r) => { const { _key, ...rest } = r; return rest; };

  const builders = {
    summary: (rows) => toCsv(
      ["File name", "File size (MB)", "Slides", "Hidden slides", "Videos", "Embedded videos", "Linked videos", "Total video size (MB)",
       "Total video length", "Images", "Audio clips", "Text density", "Avg words/slide", "Slides with notes", "Has transitions",
       "Transition slides", "Transition types", "Auto-advance", "Has animations", "Animated slides", "Detected add-ins", "Error"],
      rows.map((r) => { const s = r.summary || {}; return [r.file_name, r.file_size_mb, r.slide_count, r.hidden_slide_count, s.video_count,
        s.embedded_video_count, s.linked_video_count, s.total_video_size_mb, s.total_video_duration, s.image_count, s.audio_count,
        s.text_density, s.avg_words_per_slide, s.slides_with_notes, s.has_transitions ? "Yes" : "No", s.transition_slide_count,
        s.transition_types, s.has_auto_advance ? "Yes" : "No", s.has_animations ? "Yes" : "No", s.animated_slide_count, s.addins, r.error]; })),
    videos: (rows) => toCsv(
      ["Presentation", "Video file", "Shape name", "Source", "Slide(s)", "Size (MB)", "Duration (HH:MM:SS)", "Duration (s)", "FPS",
       "Resolution", "Video codec", "Audio codec", "Format", "Container", "Bitrate (kbps)", "Target", "Probe note"],
      rows.flatMap((r) => (r.videos || []).map((v) => [r.file_name, v.file, v.name, v.source, v.slides, v.size_mb, v.duration,
        v.duration_seconds, v.fps, v.resolution, v.video_codec, v.audio_codec, v.format, v.container, v.bitrate_kbps, v.target, v.probe_error]))),
    transitions: (rows) => toCsv(
      ["Presentation", "Slide", "Transition type", "Option", "Speed", "Duration (ms)", "Auto-advance", "Auto-advance after (s)", "Advance on click", "Sound"],
      rows.flatMap((r) => (r.transitions || []).map((t) => [r.file_name, t.slide, t.type, t.variant, t.speed, t.duration_ms,
        t.auto_advance ? "Yes" : "No", t.auto_advance_seconds, t.advance_on_click ? "Yes" : "No", t.sound]))),
    addins: (rows) => toCsv(
      ["Presentation", "Add-in", "Category", "Slides", "Presentation-level", "Detected via", "Confidence", "Details"],
      rows.flatMap((r) => (r.addins || []).map((a) => [r.file_name, a.name, a.category, a.slides, a.presentation_level ? "Yes" : "No",
        a.sources, a.confidence, a.details]))),
    slides: (rows) => toCsv(
      ["Presentation", "Slide", "Title", "Hidden", "Words", "Images", "Videos", "Audio", "Transition", "Auto-advance (s)",
       "Animation effects", "Has notes", "Add-ins", "Hyperlinks", "OLE/ActiveX objects"],
      rows.flatMap((r) => (r.slides || []).map((s) => [r.file_name, s.number, s.title, s.hidden ? "Yes" : "No", s.word_count,
        s.image_count, s.video_count, s.audio_count, s.transition, s.auto_advance_seconds, s.animation_count,
        s.has_notes ? "Yes" : "No", s.addins, s.hyperlinks, s.ole_objects]))),
  };

  document.querySelectorAll("[data-export]").forEach((btn) => btn.addEventListener("click", () => {
    const rows = visibleRows();
    if (!rows.length) { alert("Nothing to export yet."); return; }
    const kind = btn.dataset.export;
    if (kind === "json") download(`pptcheck_${stamp()}.json`, JSON.stringify(rows.map(clean), null, 2), "application/json");
    else download(`pptcheck_${kind}_${stamp()}.csv`, builders[kind](rows), "text/csv;charset=utf-8");
  }));

  const safeName = (n) => baseName(n).replace(/\.[^.]+$/, "").replace(/[^\w.-]+/g, "_");
  $("#detailJson").onclick = () => currentDetail && download(`${safeName(currentDetail.file_name)}_audit.json`,
    JSON.stringify(clean(currentDetail), null, 2), "application/json");
  $("#detailCsv").onclick = () => {
    if (!currentDetail) return;
    const r = [currentDetail];
    // One CSV containing every section, separated by blank lines.
    const parts = ["summary", "videos", "transitions", "addins", "slides"].map((k) => `${k.toUpperCase()}\r\n` + builders[k](r).slice(1));
    download(`${safeName(currentDetail.file_name)}_audit.csv`, "﻿" + parts.join("\r\n\r\n"), "text/csv;charset=utf-8");
  };

  renderTable();
})();
