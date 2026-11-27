/*
 * GRA QA drift-check core logic.
 * Pure JS, no dependencies, no I/O. Works identically loaded via <script>
 * in a browser (exposes window.GRADrift) or required in Node (module.exports).
 * This is the single source of truth for the math so the in-browser tool and
 * any Node-side test both run the exact same code.
 *
 * Mirrors the Python reference pipeline in GRAv2_verification/scripts/
 * (01_build_long.py, 02_population_and_pooled.py, 10/12_*_persistence.py,
 * 15_intra_annotator_consistency.py).
 */
(function (root) {
  "use strict";

  const LOCALES = ["Japanese", "Portuguese", "Dutch", "Hindi", "Arabic"];

  // ---------- text / date helpers ----------

  function nullIfEmpty(v) {
    // Some exports (observed from the October weekly files onward) store a
    // blank cell as a literal empty string ("") rather than an absent cell.
    // Python's openpyxl reads both as None; SheetJS's raw:false mode reads
    // the former as "", which is truthy in JS. Normalize both to null so
    // "has this rank got any data at all" matches the Python reference.
    if (v === undefined || v === null) return null;
    if (typeof v === "string" && v.trim() === "") return null;
    return v;
  }

  function normText(t) {
    if (t === null || t === undefined) return null;
    const s = String(t).normalize("NFKC");
    const parts = s.split(/\s+/).filter(Boolean);
    return parts.length ? parts.join(" ") : null;
  }

  function toDateOnly(v) {
    // Accepts a JS Date, an Excel serial number, or a date string.
    let d;
    if (v instanceof Date) {
      d = v;
    } else if (typeof v === "number") {
      // Excel serial date (SheetJS can also convert this for us via cellDates,
      // but handle the raw-number case defensively).
      const epoch = new Date(Date.UTC(1899, 11, 30));
      d = new Date(epoch.getTime() + v * 86400000);
    } else {
      d = new Date(v);
    }
    if (isNaN(d.getTime())) return null;
    return d.toISOString().slice(0, 10); // "YYYY-MM-DD"
  }

  // ---------- extraction ----------

  // rows2d: array of arrays as produced by XLSX.utils.sheet_to_json(ws, {header:1})
  // Returns: array of record objects, one per (query-row x rank 1..5) with a
  // non-null correctness-audit or a usable answer/npe field, matching the
  // skip-condition in 01_build_long.py.
  function extractRecords(period, locale, rows2d) {
    const header = rows2d[0];
    const idx = {};
    header.forEach((h, i) => {
      idx[h] = i;
    });
    const records = [];
    for (let r = 1; r < rows2d.length; r++) {
      const row = rows2d[r];
      if (!row) continue;
      const annot = row[idx["Initial_annotator_ID"]];
      const qcer = row[idx["Audit_QCer_ID"]];
      const qid = row[idx["query_id"]];
      const askedAtRaw = row[idx["asked_at"]];
      const date = toDateOnly(askedAtRaw);
      const ts = row[idx["time_sensitive"]];
      if (!date) continue;
      for (let k = 1; k <= 5; k++) {
        const corr = nullIfEmpty(row[idx[`T1_r${k}_correctness`]]);
        const corrA = nullIfEmpty(row[idx[`T1_r${k}_correctness_audit`]]);
        const npeA = nullIfEmpty(row[idx[`T1_r${k}_not_popular_entity_audit`]]);
        if (corr === null && corrA === null && npeA === null) continue;
        const complA = nullIfEmpty(row[idx[`T1_r${k}_completeness_audit`]]);
        const utilA = nullIfEmpty(row[idx[`T1_r${k}_utility_audit`]]);
        const reasonA = nullIfEmpty(row[idx[`T1_r${k}_incorrectness_reason_audit`]]);
        const textCol = idx[`T1_ranked_${k - 1}_result`];
        const text = textCol !== undefined ? nullIfEmpty(row[textCol]) : null;
        records.push({
          period,
          locale,
          annotator: nullIfEmpty(annot) !== null ? String(annot) : null,
          qcer: nullIfEmpty(qcer) !== null ? String(qcer) : null,
          query_id: qid,
          date,
          time_sensitive: ts,
          rank: k,
          npe_audit: npeA,
          correctness_qc: corrA,
          completeness_qc: complA,
          utility_qc: utilA,
          reason_qc: reasonA,
          answer_norm: normText(text),
        });
      }
    }
    return records;
  }

  function isPopulation(rec) {
    return rec.time_sensitive === "non_time_sensitive" && rec.npe_audit === "A_Yes";
  }

  // A rating axis is "field name" -> the record field holding its QC verdict.
  const METRICS = {
    correctness: { field: "correctness_qc", label: "correctness" },
    completeness: { field: "completeness_qc", label: "completeness" },
    utility: { field: "utility_qc", label: "utility" },
  };
  const METRIC_KEYS = Object.keys(METRICS);

  function metricPass(rec, metric) {
    const def = METRICS[metric] || METRICS.correctness;
    return rec[def.field] === "A_Yes";
  }

  function isCorrect(rec) {
    return metricPass(rec, "correctness");
  }

  // ---------- stats ----------

  function erf(x) {
    // Abramowitz-Stegun approximation, good to ~1e-7.
    const sign = x < 0 ? -1 : 1;
    x = Math.abs(x);
    const a1 = 0.254829592,
      a2 = -0.284496736,
      a3 = 1.421413741,
      a4 = -1.453152027,
      a5 = 1.061405429,
      p = 0.3275911;
    const t = 1 / (1 + p * x);
    const y = 1 - ((((a5 * t + a4) * t + a3) * t + a2) * t + a1) * t * Math.exp(-x * x);
    return sign * y;
  }

  function normalCdf(z) {
    return 0.5 * (1 + erf(z / Math.SQRT2));
  }

  function twoPropTest(x1, n1, x2, n2) {
    if (n1 === 0 || n2 === 0) return { z: NaN, p: NaN };
    const p1 = x1 / n1,
      p2 = x2 / n2;
    const pp = (x1 + x2) / (n1 + n2);
    const se = Math.sqrt(pp * (1 - pp) * (1 / n1 + 1 / n2));
    const z = se === 0 ? NaN : (p1 - p2) / se;
    const p = isNaN(z) ? NaN : 2 * (1 - normalCdf(Math.abs(z)));
    return { z, p };
  }

  function chi2Cdf1df(x) {
    // CDF of chi-square with 1 df == CDF of (Z^2) == 2*Phi(sqrt(x)) - 1
    if (x < 0) return 0;
    return 2 * normalCdf(Math.sqrt(x)) - 1;
  }

  function mcnemar(b, c) {
    const n = b + c;
    if (n === 0) return { stat: NaN, p: NaN };
    const stat = Math.pow(Math.abs(b - c) - 1, 2) / n;
    const p = 1 - chi2Cdf1df(stat);
    return { stat, p };
  }

  function rateN(recs, metric) {
    metric = metric || "correctness";
    const n = recs.length;
    if (n === 0) return { rate: NaN, n: 0, x: 0 };
    const x = recs.reduce((s, r) => s + (metricPass(r, metric) ? 1 : 0), 0);
    return { rate: (100 * x) / n, n, x };
  }

  function inWindow(rec, startDate, endDate) {
    return rec.date >= startDate && rec.date <= endDate;
  }

  // ---------- period-level pooled / per-locale ----------

  function periodBounds(records) {
    const byPeriod = {};
    for (const r of records) {
      if (!byPeriod[r.period]) byPeriod[r.period] = { min: r.date, max: r.date };
      else {
        if (r.date < byPeriod[r.period].min) byPeriod[r.period].min = r.date;
        if (r.date > byPeriod[r.period].max) byPeriod[r.period].max = r.date;
      }
    }
    return byPeriod;
  }

  function orderedPeriods(records) {
    const bounds = periodBounds(records);
    return Object.keys(bounds).sort((a, b) => (bounds[a].min < bounds[b].min ? -1 : 1));
  }

  function pooledByPeriod(population, baselinePeriod, metric) {
    metric = metric || "correctness";
    const periods = orderedPeriods(population);
    const byPeriod = {};
    periods.forEach((p) => {
      byPeriod[p] = population.filter((r) => r.period === p);
    });
    const baseline = baselinePeriod ? byPeriod[baselinePeriod] || [] : byPeriod[periods[0]] || [];
    const base = rateN(baseline, metric);
    return periods.map((p) => {
      const sub = byPeriod[p];
      const { rate, n, x } = rateN(sub, metric);
      const { p: pval } = twoPropTest(base.x, base.n, x, n);
      return { period: p, n, rate, deltaVsBaseline: rate - base.rate, p: pval, dateMin: periodBounds(population)[p].min, dateMax: periodBounds(population)[p].max };
    });
  }

  function perLocaleByPeriod(population, baselinePeriod, metric) {
    metric = metric || "correctness";
    const periods = orderedPeriods(population);
    const base = baselinePeriod || periods[0];
    const out = {};
    for (const loc of LOCALES) {
      const baseSub = population.filter((r) => r.locale === loc && r.period === base);
      const baseStat = rateN(baseSub, metric);
      out[loc] = periods.map((p) => {
        const sub = population.filter((r) => r.locale === loc && r.period === p);
        const { rate, n, x } = rateN(sub, metric);
        const { p: pval } = twoPropTest(baseStat.x, baseStat.n, x, n);
        return { period: p, n, rate, deltaVsBaseline: rate - baseStat.rate, p: pval };
      });
    }
    return out;
  }

  // ---------- rolling reconstruction ----------

  function rollingRate(population, windowDays, minPeriods) {
    const daily = {};
    for (const r of population) {
      if (!daily[r.date]) daily[r.date] = { sum: 0, count: 0 };
      daily[r.date].count += 1;
      if (isCorrect(r)) daily[r.date].sum += 1;
    }
    const dates = Object.keys(daily).sort();
    if (dates.length === 0) return [];
    const start = new Date(dates[0]);
    const end = new Date(dates[dates.length - 1]);
    const allDates = [];
    for (let d = new Date(start); d <= end; d.setUTCDate(d.getUTCDate() + 1)) {
      allDates.push(d.toISOString().slice(0, 10));
    }
    const out = [];
    for (let i = 0; i < allDates.length; i++) {
      const windowDates = allDates.slice(Math.max(0, i - windowDays + 1), i + 1).filter((d) => daily[d]);
      if (windowDates.length < minPeriods) {
        out.push({ date: allDates[i], rate: null });
        continue;
      }
      let sum = 0,
        count = 0;
      for (const d of windowDates) {
        sum += daily[d].sum;
        count += daily[d].count;
      }
      out.push({ date: allDates[i], rate: count ? (100 * sum) / count : null });
    }
    return out;
  }

  // ---------- annotator concentration (top annotator per locale per period) ----------

  function concentrationByPeriod(population) {
    const periods = orderedPeriods(population);
    const rows = [];
    for (const period of periods) {
      for (const loc of LOCALES) {
        const sub = population.filter((r) => r.locale === loc && r.period === period);
        if (sub.length === 0) continue;
        const counts = {};
        for (const r of sub) {
          if (!r.annotator) continue;
          counts[r.annotator] = (counts[r.annotator] || 0) + 1;
        }
        const ids = Object.keys(counts).sort((a, b) => counts[b] - counts[a]);
        if (ids.length === 0) continue;
        const topId = ids[0];
        const topRows = sub.filter((r) => r.annotator === topId);
        const restRows = sub.filter((r) => r.annotator !== topId);
        rows.push({
          period,
          locale: loc,
          topAnnotator: topId,
          topShare: (100 * counts[topId]) / sub.length,
          topN: counts[topId],
          topRate: rateN(topRows).rate,
          restRate: rateN(restRows).rate,
          nAnnotators: ids.length,
        });
      }
    }
    return rows;
  }

  // ---------- self-contradiction (intra-annotator) check ----------
  // Uses ALL rows with a non-null correctness_audit + usable answer text + qcer
  // identity (not restricted to the `population` filter) -- matches script 15.

  function selfContradiction(allRecords) {
    const withText = allRecords.filter((r) => r.qcer && r.answer_norm && r.correctness_qc !== null && r.correctness_qc !== undefined);
    const groups = {};
    for (const r of withText) {
      const key = [r.locale, r.qcer, r.query_id, r.rank, r.answer_norm].join("||");
      if (!groups[key]) groups[key] = [];
      groups[key].push(r);
    }
    const results = [];
    for (const key of Object.keys(groups)) {
      const g = groups[key];
      const dates = new Set(g.map((r) => r.date));
      if (g.length < 2 || dates.size < 2) continue;
      g.sort((a, b) => (a.date < b.date ? -1 : 1));
      const first = g[0];
      const last = g[g.length - 1];
      results.push({
        locale: first.locale,
        qcer: first.qcer,
        firstDate: first.date,
        lastDate: last.date,
        firstCorrect: isCorrect(first),
        lastCorrect: isCorrect(last),
        contradicted: isCorrect(first) !== isCorrect(last),
      });
    }
    return results;
  }

  function summarizeContradictions(results) {
    const n = results.length;
    const c = results.filter((r) => r.contradicted).length;
    const perLocale = {};
    for (const loc of LOCALES) {
      const sub = results.filter((r) => r.locale === loc);
      if (sub.length === 0) continue;
      perLocale[loc] = { n: sub.length, contradictions: sub.filter((r) => r.contradicted).length, rate: (100 * sub.filter((r) => r.contradicted).length) / sub.length };
    }
    const perMonth = {};
    for (const r of results) {
      const month = r.lastDate.slice(0, 7);
      if (!perMonth[month]) perMonth[month] = { n: 0, c: 0 };
      perMonth[month].n += 1;
      if (r.contradicted) perMonth[month].c += 1;
    }
    const perLocaleMonth = {};
    for (const loc of LOCALES) {
      perLocaleMonth[loc] = {};
      const sub = results.filter((r) => r.locale === loc);
      for (const r of sub) {
        const month = r.lastDate.slice(0, 7);
        if (!perLocaleMonth[loc][month]) perLocaleMonth[loc][month] = { n: 0, c: 0 };
        perLocaleMonth[loc][month].n += 1;
        if (r.contradicted) perLocaleMonth[loc][month].c += 1;
      }
    }
    return {
      overall: { n, contradictions: c, rate: n ? (100 * c) / n : NaN },
      perLocale,
      perMonth,
      perLocaleMonth,
    };
  }

  // ---------- auto-flag heuristic ----------
  // A locale is flagged as a drift candidate if:
  //  (a) its per-locale rate has moved in the same direction vs baseline for
  //      the last >=3 periods (monotonic run), with the latest period's p<0.10
  //  (b) self-contradiction rate for that locale is flat or falling across
  //      the last two months with data -- rules out "it's just noise"

  function autoFlag(perLocaleSeries, contradictionSummary) {
    const flags = {};
    for (const loc of LOCALES) {
      const series = perLocaleSeries[loc] || [];
      if (series.length < 3) {
        flags[loc] = { flagged: false, reason: "not enough periods yet" };
        continue;
      }
      // A single natural weekly period is noisy (deltas vs. baseline zigzag
      // week to week even under real sustained drift), so instead of
      // requiring an uninterrupted monotonic increase/decrease, require the
      // last 3 periods to all sit on the same side of the baseline and each
      // be individually significant -- a sustained, same-direction deviation.
      const last3 = series.slice(-3);
      const signs = last3.map((s) => Math.sign(s.deltaVsBaseline));
      const sameSign = signs.every((s) => s === signs[0]) && signs[0] !== 0;
      const allSignificant = last3.every((s) => s.p < 0.1);
      const monotonicUp = sameSign && signs[0] > 0;
      const monotonicDown = sameSign && signs[0] < 0;
      const trendOk = sameSign && allSignificant;

      const months = Object.keys(contradictionSummary.perLocaleMonth[loc] || {}).sort();
      let noiseOk = true;
      if (months.length >= 2) {
        const last = months[months.length - 1];
        const prev = months[months.length - 2];
        const lastRate = (100 * contradictionSummary.perLocaleMonth[loc][last].c) / contradictionSummary.perLocaleMonth[loc][last].n;
        const prevRate = (100 * contradictionSummary.perLocaleMonth[loc][prev].c) / contradictionSummary.perLocaleMonth[loc][prev].n;
        noiseOk = lastRate <= prevRate + 2; // allow 2pp slack for small-n noise
      }

      flags[loc] = {
        flagged: trendOk && noiseOk,
        direction: monotonicUp ? "more lenient" : monotonicDown ? "stricter" : "no clear trend",
        trendOk,
        noiseOk,
      };
    }
    return flags;
  }

  const GRADrift = {
    LOCALES,
    METRICS,
    METRIC_KEYS,
    normText,
    toDateOnly,
    extractRecords,
    isPopulation,
    isCorrect,
    metricPass,
    twoPropTest,
    mcnemar,
    rateN,
    inWindow,
    orderedPeriods,
    periodBounds,
    pooledByPeriod,
    perLocaleByPeriod,
    rollingRate,
    concentrationByPeriod,
    selfContradiction,
    summarizeContradictions,
    autoFlag,
  };

  if (typeof module !== "undefined" && module.exports) {
    module.exports = GRADrift;
  }
  if (root) {
    root.GRADrift = GRADrift;
  }
})(typeof window !== "undefined" ? window : typeof global !== "undefined" ? global : null);
