/* UI wiring for the GRA QA drift-check tool. Depends on core.js (GRADrift) and SheetJS (XLSX), both loaded via <script> before this file. */
(function () {
  "use strict";

  const fileInput = document.getElementById("fileInput");
  const analyzeBtn = document.getElementById("analyzeBtn");
  const statusEl = document.getElementById("status");
  const resultsEl = document.getElementById("results");
  const fileListEl = document.getElementById("fileList");

  let pickedFiles = [];
  let cachedRecords = null;

  fileInput.addEventListener("change", (e) => {
    pickedFiles = Array.from(e.target.files).filter(
      (f) => f.name.endsWith(".xlsx") && !f.name.includes("Percentage_QCed")
    );
    fileListEl.textContent = pickedFiles.length
      ? `${pickedFiles.length} file(s) selected`
      : "No valid Weekly_report_*.xlsx files selected";
    analyzeBtn.disabled = pickedFiles.length === 0;
  });

  analyzeBtn.addEventListener("click", async () => {
    analyzeBtn.disabled = true;
    resultsEl.innerHTML = "";
    const total = pickedFiles.length;
    let allRecords = [];
    for (let i = 0; i < total; i++) {
      const file = pickedFiles[i];
      setStatus(`Reading ${i + 1} / ${total}: ${file.name}`);
      const m = file.name.match(/Weekly_report_(.+)_([A-Za-z]+)\.xlsx$/);
      if (!m) continue;
      const [, period, locale] = m;
      const buf = await file.arrayBuffer();
      const wb = XLSX.read(buf, { cellDates: true, type: "array" });
      const ws = wb.Sheets["Raw_Annotations"];
      if (!ws) continue;
      const rows2d = XLSX.utils.sheet_to_json(ws, { header: 1, raw: false, defval: null });
      const recs = GRADrift.extractRecords(period, locale, rows2d);
      allRecords = allRecords.concat(recs);
      await new Promise((r) => setTimeout(r, 0)); // yield to keep UI responsive
    }
    setStatus(`Parsed ${allRecords.length} answer-rank rows from ${total} file(s). Running analysis...`);
    await new Promise((r) => setTimeout(r, 20));
    cachedRecords = allRecords;
    const population = allRecords.filter(GRADrift.isPopulation);
    const periods = GRADrift.orderedPeriods(population);
    renderBaselinePicker(population, periods, periods[0]);
    render(allRecords, periods[0]);
    setStatus(`Done. ${allRecords.length} answer-rank rows analyzed.`);
    analyzeBtn.disabled = false;
  });

  function renderBaselinePicker(population, periods, selected) {
    const bounds = GRADrift.periodBounds(population);
    const existing = document.getElementById("baselinePicker");
    if (existing) existing.remove();
    const select = el(
      "select",
      { id: "baselineSelect" },
      periods.map((p) => {
        const n = population.filter((r) => r.period === p).length;
        const opt = el("option", { value: p }, [`${p}  (${bounds[p].min} → ${bounds[p].max}, n=${n})`]);
        if (p === selected) opt.setAttribute("selected", "selected");
        return opt;
      })
    );
    select.addEventListener("change", () => {
      render(cachedRecords, select.value);
    });
    const picker = el("div", { id: "baselinePicker", class: "card" }, [
      el("h2", null, ["Baseline period"]),
      el(
        "p",
        { class: "note" },
        [
          "Every Δ and p-value below is computed against this period. A short early period can be noisy " +
            "(small n) and give a misleading baseline — pick a period with reasonable volume, or one you know " +
            "represents normal/expected performance, rather than always defaulting to the very first upload.",
        ]
      ),
      select,
    ]);
    resultsEl.before(picker);
  }

  function setStatus(msg) {
    statusEl.textContent = msg;
  }

  function el(tag, attrs, children) {
    const e = document.createElement(tag);
    if (attrs) {
      for (const k in attrs) {
        if (k === "class") e.className = attrs[k];
        else e.setAttribute(k, attrs[k]);
      }
    }
    (children || []).forEach((c) => e.appendChild(typeof c === "string" ? document.createTextNode(c) : c));
    return e;
  }

  function fmtPct(x) {
    return isNaN(x) ? "–" : x.toFixed(1) + "%";
  }
  function fmtDelta(x) {
    return isNaN(x) ? "–" : (x >= 0 ? "+" : "") + x.toFixed(1);
  }
  function fmtP(x) {
    if (isNaN(x)) return "–";
    return x < 0.0001 ? "<0.0001" : x.toFixed(4);
  }
  function sigBadge(p) {
    if (isNaN(p)) return el("span", { class: "badge muted" }, ["n/a"]);
    if (p < 0.05) return el("span", { class: "badge yes" }, ["yes"]);
    return el("span", { class: "badge no" }, ["no"]);
  }

  function table(headers, rows) {
    const t = el("table", { class: "tbl" });
    const thead = el("thead", null, [el("tr", null, headers.map((h) => el("th", null, [h])))]);
    const tbody = el("tbody", null, rows.map((r) => el("tr", null, r)));
    t.appendChild(thead);
    t.appendChild(tbody);
    return t;
  }

  function section(title, note, content) {
    const s = el("section", { class: "card" }, [el("h2", null, [title])]);
    if (note) s.appendChild(el("p", { class: "note" }, [note]));
    s.appendChild(content);
    return s;
  }

  function lineChart(series, { width = 760, height = 220, yLabel = "%", baseline = null } = {}) {
    // series: [{label, points: [{x: index, y: value|null}], color}]
    const pad = { l: 44, r: 16, t: 16, b: 28 };
    const allY = [];
    series.forEach((s) => s.points.forEach((p) => p.y !== null && !isNaN(p.y) && allY.push(p.y)));
    if (baseline !== null) allY.push(baseline);
    const yMin = Math.min(...allY, 0);
    const yMax = Math.max(...allY, 1);
    const yRange = yMax - yMin || 1;
    const n = Math.max(...series.map((s) => s.points.length), 2);
    const xScale = (i) => pad.l + (i / (n - 1)) * (width - pad.l - pad.r);
    const yScale = (v) => height - pad.b - ((v - yMin) / yRange) * (height - pad.t - pad.b);

    const svgNS = "http://www.w3.org/2000/svg";
    const svg = document.createElementNS(svgNS, "svg");
    svg.setAttribute("viewBox", `0 0 ${width} ${height}`);
    svg.setAttribute("class", "chart");

    // gridlines + y-axis labels
    const ticks = 4;
    for (let i = 0; i <= ticks; i++) {
      const v = yMin + (yRange * i) / ticks;
      const y = yScale(v);
      const line = document.createElementNS(svgNS, "line");
      line.setAttribute("x1", pad.l);
      line.setAttribute("x2", width - pad.r);
      line.setAttribute("y1", y);
      line.setAttribute("y2", y);
      line.setAttribute("class", "gridline");
      svg.appendChild(line);
      const label = document.createElementNS(svgNS, "text");
      label.setAttribute("x", pad.l - 8);
      label.setAttribute("y", y + 4);
      label.setAttribute("class", "axislabel");
      label.setAttribute("text-anchor", "end");
      label.textContent = v.toFixed(0) + (yLabel === "%" ? "%" : "");
      svg.appendChild(label);
    }

    if (baseline !== null) {
      const by = yScale(baseline);
      const bline = document.createElementNS(svgNS, "line");
      bline.setAttribute("x1", pad.l);
      bline.setAttribute("x2", width - pad.r);
      bline.setAttribute("y1", by);
      bline.setAttribute("y2", by);
      bline.setAttribute("class", "baseline");
      svg.appendChild(bline);
    }

    series.forEach((s) => {
      const pts = s.points.filter((p) => p.y !== null && !isNaN(p.y));
      if (pts.length === 0) return;
      const d = pts.map((p, idx) => `${idx === 0 ? "M" : "L"}${xScale(p.x)},${yScale(p.y)}`).join(" ");
      const path = document.createElementNS(svgNS, "path");
      path.setAttribute("d", d);
      path.setAttribute("fill", "none");
      path.setAttribute("stroke", s.color);
      path.setAttribute("stroke-width", "2");
      svg.appendChild(path);
      pts.forEach((p) => {
        const c = document.createElementNS(svgNS, "circle");
        c.setAttribute("cx", xScale(p.x));
        c.setAttribute("cy", yScale(p.y));
        c.setAttribute("r", "3");
        c.setAttribute("fill", s.color);
        const t = document.createElementNS(svgNS, "title");
        t.textContent = `${s.label}: ${p.y.toFixed(1)}${yLabel === "%" ? "%" : ""}`;
        c.appendChild(t);
        svg.appendChild(c);
      });
    });

    const wrap = el("div", { class: "chartwrap" }, [svg]);
    const legend = el(
      "div",
      { class: "legend" },
      series.map((s) =>
        el("span", { class: "legenditem" }, [
          el("span", { class: "swatch", style: `background:${s.color}` }, []),
          s.label,
        ])
      )
    );
    const container = el("div", null, [wrap, legend]);
    return container;
  }

  const PALETTE = {
    Japanese: "#4C72B0",
    Portuguese: "#DD8452",
    Dutch: "#55A868",
    Hindi: "#C44E52",
    Arabic: "#8172B2",
    pooled: "#333333",
  };

  function render(allRecords, baselinePeriod) {
    resultsEl.innerHTML = "";
    const population = allRecords.filter(GRADrift.isPopulation);
    const periods = GRADrift.orderedPeriods(population);
    baselinePeriod = baselinePeriod || periods[0];

    // --- summary ---
    const summaryCard = section(
      "Summary",
      null,
      el("div", { class: "statrow" }, [
        stat("Answer-rank rows parsed", allRecords.length),
        stat("Population rows (non-time-sensitive, NPE passed)", population.length),
        stat("Periods detected", periods.length),
        stat("Baseline period", baselinePeriod || "–"),
      ])
    );
    resultsEl.appendChild(summaryCard);

    // --- pooled + per-locale, for each rating axis (correctness, completeness, utility) ---
    // Rendered as tabs: one axis's tables/charts are visible at a time, switchable without re-running the analysis.
    let perLocale; // keep correctness' per-locale series around for mix-shift/auto-flag below
    const tabButtons = [];
    const tabPanels = [];
    GRADrift.METRIC_KEYS.forEach((metric, metricIdx) => {
      const label = GRADrift.METRICS[metric].label;
      const panel = el("div", { class: "tabpanel" + (metricIdx === 0 ? "" : " hidden") }, []);

      // pooled table
      const pooled = GRADrift.pooledByPeriod(population, baselinePeriod, metric);
      const pooledRows = pooled.map((p) => [
        el("td", null, [p.period]),
        el("td", null, [String(p.n)]),
        el("td", null, [fmtPct(p.rate)]),
        el("td", null, [fmtDelta(p.deltaVsBaseline)]),
        el("td", null, [fmtP(p.p)]),
        el("td", null, [sigBadge(p.p)]),
      ]);
      panel.appendChild(
        section(
          `Pooled ${label} rate by period (baseline: ${baselinePeriod})`,
          "Δ and p are versus the baseline period (earliest period in the uploaded set).",
          table(["Period", "n", "Rate", "Δ vs baseline (pp)", "p", "Significant?"], pooledRows)
        )
      );

      // pooled chart
      const pooledSeries = [
        {
          label: "Pooled",
          color: PALETTE.pooled,
          points: pooled.map((p, i) => ({ x: i, y: p.rate })),
        },
      ];
      panel.appendChild(
        section(
          `Pooled ${label} rate trend`,
          null,
          lineChart(pooledSeries, { baseline: pooled[0] ? pooled[0].rate : null })
        )
      );

      // per-locale chart + tables
      const perLocaleForMetric = GRADrift.perLocaleByPeriod(population, baselinePeriod, metric);
      if (metric === "correctness") perLocale = perLocaleForMetric;
      const locSeries = GRADrift.LOCALES.map((loc) => ({
        label: loc,
        color: PALETTE[loc],
        points: (perLocaleForMetric[loc] || []).map((p, i) => ({ x: i, y: p.rate })),
      }));
      panel.appendChild(section(`Per-locale ${label} rate trend`, null, lineChart(locSeries)));

      GRADrift.LOCALES.forEach((loc) => {
        const rows = (perLocaleForMetric[loc] || []).map((p) => [
          el("td", null, [p.period]),
          el("td", null, [String(p.n)]),
          el("td", null, [fmtPct(p.rate)]),
          el("td", null, [fmtDelta(p.deltaVsBaseline)]),
          el("td", null, [fmtP(p.p)]),
          el("td", null, [sigBadge(p.p)]),
        ]);
        panel.appendChild(
          section(
            `${loc}: ${label} rate by period`,
            null,
            table(["Period", "n", "Rate", "Δ vs baseline (pp)", "p", "Significant?"], rows)
          )
        );
      });

      const btn = el("button", { type: "button", class: "tab" + (metricIdx === 0 ? " active" : "") }, [
        label[0].toUpperCase() + label.slice(1),
      ]);
      btn.addEventListener("click", () => {
        tabButtons.forEach((b) => b.classList.remove("active"));
        tabPanels.forEach((p) => p.classList.add("hidden"));
        btn.classList.add("active");
        panel.classList.remove("hidden");
      });
      tabButtons.push(btn);
      tabPanels.push(panel);
    });

    resultsEl.appendChild(el("div", { class: "tabsnav" }, tabButtons));
    resultsEl.appendChild(el("div", { class: "tabpanels" }, tabPanels));

    resultsEl.appendChild(
      el("h1", { class: "metricheading" }, ["Annotator / QC behavior checks (correctness-based)"])
    );

    // --- mix-shift ---
    const mixRows = GRADrift.mixShiftByPeriod(population);
    const mixRowsEls = mixRows.map((r) => [
      el("td", null, [r.period]),
      el("td", null, [r.locale]),
      el("td", null, [String(r.nTotal)]),
      el("td", null, [String(r.nNew)]),
      el("td", null, [fmtPct(r.newSharePct)]),
      el("td", null, [fmtPct(r.rateNew)]),
      el("td", null, [fmtPct(r.rateContinuing)]),
    ]);
    resultsEl.appendChild(
      section(
        "New-annotator mix-shift check",
        "\"New\" = this annotator's first-seen date (across the whole uploaded corpus) is on/after this period's start. Low new-share + stable rate rules out a hiring/mix-shift explanation for any rate move.",
        table(
          ["Period", "Locale", "n", "n new", "New share", "Rate (new)", "Rate (continuing)"],
          mixRowsEls
        )
      )
    );

    // --- concentration ---
    const concRows = GRADrift.concentrationByPeriod(population);
    const concRowsEls = concRows.map((r) => [
      el("td", null, [r.period]),
      el("td", null, [r.locale]),
      el("td", null, [r.topAnnotator]),
      el("td", null, [fmtPct(r.topShare)]),
      el("td", null, [String(r.topN)]),
      el("td", null, [fmtPct(r.topRate)]),
      el("td", null, [fmtPct(r.restRate)]),
      el("td", null, [String(r.nAnnotators)]),
    ]);
    resultsEl.appendChild(
      section(
        "Annotator concentration (top annotator per locale per period)",
        "Flags any single annotator carrying an outsized share of a locale's weekly volume, and whether their own rate differs from the rest of the locale.",
        table(
          ["Period", "Locale", "Top annotator", "Share", "n", "Their rate", "Rest rate", "# annotators"],
          concRowsEls
        )
      )
    );

    // --- self-contradiction ---
    const contraResults = GRADrift.selfContradiction(allRecords);
    const contraSummary = GRADrift.summarizeContradictions(contraResults);
    const overallCard = section(
      "Intra-annotator self-contradiction check",
      "Same QCer grading the literally identical (query, rank, answer text) item on more than one distinct date. A falling rate alongside a moving pooled rate is evidence of real standard drift (consistent, not noisy); a rising rate alongside a flat pooled rate suggests noise, not drift.",
      el("div", { class: "statrow" }, [
        stat("Matched same-QCer groups", contraSummary.overall.n),
        stat("Self-contradictions", contraSummary.overall.contradictions),
        stat("Overall rate", fmtPct(contraSummary.overall.rate)),
      ])
    );
    resultsEl.appendChild(overallCard);

    const perLocRows = GRADrift.LOCALES.filter((l) => contraSummary.perLocale[l]).map((l) => {
      const s = contraSummary.perLocale[l];
      return [el("td", null, [l]), el("td", null, [String(s.n)]), el("td", null, [String(s.contradictions)]), el("td", null, [fmtPct(s.rate)])];
    });
    resultsEl.appendChild(
      section("Self-contradiction rate, per locale", null, table(["Locale", "n", "Contradictions", "Rate"], perLocRows))
    );

    const contraSeries = GRADrift.LOCALES.map((loc) => {
      const months = Object.keys(contraSummary.perLocaleMonth[loc] || {}).sort();
      return {
        label: loc,
        color: PALETTE[loc],
        points: months.map((m, i) => {
          const d = contraSummary.perLocaleMonth[loc][m];
          return { x: i, y: d.n >= 5 ? (100 * d.c) / d.n : null, _month: m };
        }),
      };
    });
    resultsEl.appendChild(section("Self-contradiction rate over time, per locale", "Months with fewer than 5 matched groups are omitted as too small to trust.", lineChart(contraSeries)));

    // --- auto-flag ---
    const flags = GRADrift.autoFlag(perLocale, mixRows, contraSummary);
    const flagRows = GRADrift.LOCALES.map((loc) => {
      const f = flags[loc];
      const verdict = f.flagged
        ? el("span", { class: "badge yes" }, ["FLAGGED"])
        : el("span", { class: "badge no" }, ["not flagged"]);
      return [
        el("td", null, [loc]),
        el("td", null, [verdict]),
        el("td", null, [f.direction || "–"]),
        el("td", null, [f.reason || `trend:${f.trendOk ? "✓" : "✗"} mix-ruled-out:${f.mixOk ? "✓" : "✗"} noise-not-rising:${f.noiseOk ? "✓" : "✗"}`]),
      ];
    });
    resultsEl.appendChild(
      section(
        "Auto-flagged drift candidates",
        "Heuristic: flagged only if (a) the per-locale rate moved in the same direction vs. baseline for the last 3 periods with p<0.10, (b) new-annotator share in the latest period is <10% (rules out hiring/mix shift), and (c) self-contradiction rate isn't rising month over month (rules out noise). This is a trigger for investigation, not a final verdict — read the underlying tables before acting.",
        table(["Locale", "Verdict", "Direction", "Conditions"], flagRows)
      )
    );
  }

  function stat(label, value) {
    return el("div", { class: "stat" }, [el("div", { class: "statvalue" }, [String(value)]), el("div", { class: "statlabel" }, [label])]);
  }
})();
