/* Scout Patrol Sorter — fully client-side.
   Forms single-gender patrols (size minSize..maxSize), evenly distributed,
   age-banded so tent pairs respect the max age-gap rule, with <=maxUnit per unit
   where possible, and assigns tent pairs (odd patrol -> one triple). */

'use strict';

const PATROL_NAMES = [
  'Eagle', 'Fox', 'Hawk', 'Bear', 'Wolf', 'Owl', 'Cobra', 'Falcon',
  'Bobcat', 'Stag', 'Raven', 'Otter', 'Lynx', 'Bison', 'Panther', 'Heron'
];

let RAW_ROWS = [];     // array of objects keyed by original headers
let HEADERS = [];      // original header strings
const MS_PER_YEAR = 365.25 * 24 * 3600 * 1000;
const MS_PER_DAY = 24 * 3600 * 1000;

/* ---------- helpers ---------- */
const $ = (id) => document.getElementById(id);
const show = (id) => $(id).classList.remove('hidden');
const esc = (s) => String(s ?? '').replace(/[&<>"']/g, c =>
  ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));

function fuzzyFindHeader(headers, patterns) {
  for (const p of patterns) {
    const hit = headers.find(h => p.test(h));
    if (hit) return hit;
  }
  return '';
}

// Always return a LOCAL-midnight date (no time/timezone drift) so day-gaps and
// adult (18+) calendar math are exact and display dates don't shift.
function parseDateLoose(v) {
  if (v == null || v === '') return null;
  const ymd = (y, m, d) => { const dt = new Date(y, m - 1, d); return isNaN(dt) ? null : dt; };
  // Excel / native Date object: SheetJS serial dates land on UTC midnight — read UTC parts.
  if (v instanceof Date) return isNaN(v) ? null : ymd(v.getUTCFullYear(), v.getUTCMonth() + 1, v.getUTCDate());
  const s = String(v).trim();
  let m;
  // ISO-ish Y-M-D (optionally with time) — take the date part only
  m = s.match(/^(\d{4})[\/\-.](\d{1,2})[\/\-.](\d{1,2})/);
  if (m) return ymd(+m[1], +m[2], +m[3]);
  // M/D/Y or M-D-Y
  m = s.match(/^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{2,4})$/);
  if (m) { let yr = +m[3]; if (yr < 100) yr += yr < 30 ? 2000 : 1900; return ymd(yr, +m[1], +m[2]); }
  // fallback: let the engine parse (e.g. "March 14, 2009"), normalize to local midnight
  const d = new Date(s);
  if (!isNaN(d)) return ymd(d.getFullYear(), d.getMonth() + 1, d.getDate());
  return null;
}

function normGender(v) {
  const s = String(v ?? '').trim().toLowerCase();
  if (!s) return 'Unknown';
  if (s[0] === 'm' || s === 'boy' || s === 'b') return 'Male';
  if (s[0] === 'f' || s[0] === 'w' || s === 'girl' || s === 'g') return 'Female';
  return 'Unknown';
}

function ageYears(birth, ref) {
  if (!birth) return null;
  return (ref - birth) / MS_PER_YEAR;
}

// True if the person is at least `years` old (by calendar) on the reference date.
function isAtLeastAt(birth, ref, years) {
  if (!birth) return false;
  const cutoff = new Date(ref.getFullYear() - years, ref.getMonth(), ref.getDate());
  return birth.getTime() <= cutoff.getTime();
}
// True if the person is 18 or older (by calendar) on the reference date.
function isAdultAt(birth, ref) {
  return isAtLeastAt(birth, ref, 18);
}

/* ---------- input loading ---------- */
function ingestRows(rows) {
  // rows: array of plain objects. Drop fully-empty rows.
  RAW_ROWS = rows.filter(r => Object.values(r).some(v => String(v ?? '').trim() !== ''));
  HEADERS = RAW_ROWS.length ? Object.keys(RAW_ROWS[0]) : [];
  if (!RAW_ROWS.length) { alert('No data rows found.'); return; }
  buildMappingUI();  show('step-map');
  $('step-map').scrollIntoView({ behavior: 'smooth' });
}

function handleFile(file) {
  $('fileName').textContent = file.name;
  const ext = file.name.split('.').pop().toLowerCase();
  if (ext === 'xlsx' || ext === 'xls') {
    const reader = new FileReader();
    reader.onload = (e) => {
      const wb = XLSX.read(e.target.result, { type: 'array', cellDates: true });
      const ws = wb.Sheets[wb.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(ws, { defval: '', raw: true });
      ingestRows(rows);
    };
    reader.readAsArrayBuffer(file);
  } else {
    Papa.parse(file, {
      header: true, skipEmptyLines: 'greedy',
      complete: (res) => ingestRows(res.data)
    });
  }
}

function parsePaste() {
  const text = $('pasteInput').value.trim();
  if (!text) { alert('Paste some data first.'); return; }
  const res = Papa.parse(text, { header: true, skipEmptyLines: 'greedy' });
  ingestRows(res.data);
}

/* ---------- mapping UI ---------- */
const FIELD_DEFS = [
  { key: 'name',   label: 'Name',       pats: [/full ?name/i, /name/i] },
  { key: 'email',  label: 'Email',      pats: [/e-?mail/i] },
  { key: 'unit',   label: 'Unit #',     pats: [/unit/i, /troop/i, /crew/i, /pack|ship|post/i] },
  { key: 'gender', label: 'Gender',     pats: [/gender/i, /sex/i] },
  { key: 'birth',  label: 'Birth date', pats: [/birth/i, /\bdob\b/i, /b-?day/i] },
];

function buildMappingUI() {
  const grid = $('mapGrid');
  grid.innerHTML = '';
  for (const f of FIELD_DEFS) {
    const detected = fuzzyFindHeader(HEADERS, f.pats);
    const opts = ['<option value="">— none —</option>']
      .concat(HEADERS.map(h => `<option value="${esc(h)}"${h === detected ? ' selected' : ''}>${esc(h)}</option>`));
    const wrap = document.createElement('label');
    wrap.innerHTML = `${f.label}<select data-field="${f.key}">${opts.join('')}</select>`;
    grid.appendChild(wrap);
  }
  if (!$('courseDate').value) $('courseDate').value = new Date().toISOString().slice(0, 10);
  $('loadStats').textContent = `${RAW_ROWS.length} rows loaded.`;
}

function getMapping() {
  const map = {};
  document.querySelectorAll('#mapGrid select').forEach(s => map[s.dataset.field] = s.value);
  return map;
}

/* ---------- build participant objects ---------- */
function buildParticipants() {
  const map = getMapping();
  // Course date determines who is 18+ (adult) during the event; also used for display ages.
  const ref = parseDateLoose($('courseDate').value) || new Date();
  const people = [];
  const issues = [];
  RAW_ROWS.forEach((r, i) => {
    const birth = parseDateLoose(map.birth ? r[map.birth] : '');
    const age = ageYears(birth, ref);
    const p = {
      id: i,
      name: (map.name ? r[map.name] : '') || `Row ${i + 1}`,
      email: map.email ? r[map.email] : '',
      unit: String((map.unit ? r[map.unit] : '') ?? '').trim() || '—',
      gender: normGender(map.gender ? r[map.gender] : ''),
      birth, age, adult: isAdultAt(birth, ref), age21: isAtLeastAt(birth, ref, 21)
    };
    if (!birth) issues.push(`Missing/unreadable birth date for "${esc(p.name)}".`);
    if (p.gender === 'Unknown') issues.push(`Unknown gender for "${esc(p.name)}".`);
    people.push(p);
  });
  return { people, issues, ref };
}

/* ---------- patrol sizing (even distribution) ---------- */
function planSizes(n, minS, maxS) {
  if (n <= 0) return { sizes: [], outOfRange: false };
  let best = null;
  for (let k = 1; k <= n; k++) {
    const base = Math.floor(n / k), rem = n % k;
    const sizes = Array.from({ length: k }, (_, i) => base + (i < rem ? 1 : 0));
    const min = Math.min(...sizes), max = Math.max(...sizes);
    const inRange = min >= minS && max <= maxS;
    const spread = max - min;
    const target = (minS + maxS) / 2;
    const score = (inRange ? 0 : 10000) + spread * 100 + Math.abs(target - n / k);
    if (best === null || score < best.score)
      best = { sizes, score, outOfRange: !inRange };
  }
  // largest patrols first for stable naming
  best.sizes.sort((a, b) => b - a);
  return best;
}

/* ---------- tent pairing within a patrol ---------- */
const isAdult = (p) => !!p.adult; // 18+ (calendar) at the course date

// Tent a homogeneous group (all youth OR all adults). Tents are pairs only;
// an odd leftover person gets a solo tent (no tent ever holds three).
function tentsForGroup(list) {
  const m = list.slice().sort((a, b) => (a.birth || 0) - (b.birth || 0));
  const tents = [];
  let i = 0;
  for (; i + 1 < m.length; i += 2) tents.push({ members: [m[i], m[i + 1]], kind: 'pair' });
  if (i < m.length) tents.push({ members: [m[i]], kind: 'solo' });
  return tents;
}

function assignTents(members, maxGapDays) {
  // Adults (18+) and youth (<18) are tented separately — they must never share a tent.
  const adults = members.filter(isAdult);
  const youth = members.filter(p => !isAdult(p));
  let tents = [...tentsForGroup(youth), ...tentsForGroup(adults)];

  // stable display order: youth before adults; pairs, then solos
  const kindRank = { pair: 0, solo: 1 };
  const grp = (t) => t.members.some(isAdult) ? 1 : 0;
  tents.sort((a, b) => grp(a) - grp(b) || kindRank[a.kind] - kindRank[b.kind]);

  const warnings = [];
  tents.forEach((t, idx) => {
    const hasAdult = t.members.some(isAdult);
    const hasYouth = t.members.some(x => !isAdult(x));
    if (hasAdult && hasYouth) {
      const names = t.members.map(x => `${x.name} (${fmtAge(x.age)})`).join(', ');
      warnings.push(`Tent ${idx + 1}: an adult (18+) cannot share a tent with a youth (under 18): ${names}.`);
    }
    const births = t.members.map(x => x.birth);
    if (births.some(b => !b)) {
      warnings.push(`Tent ${idx + 1}: missing birth date — verify ages/adult status manually.`);
      return;
    }
    // 730-day gap is a youth-tenting rule: apply to any tent containing a youth.
    if (hasYouth) {
      const times = births.map(b => +b);
      const spreadDays = (Math.max(...times) - Math.min(...times)) / MS_PER_DAY;
      if (spreadDays > maxGapDays + 1e-6) {
        const names = t.members.map(x => `${x.name} (${fmtDate(x.birth)})`).join(', ');
        warnings.push(`Tent ${idx + 1} exceeds ${maxGapDays}-day gap (${Math.round(spreadDays)} days apart): ${names}.`);
      }
    }
  });
  return { tents, warnings };
}

/* ---------- unit balancing ---------- */
function unitOverage(patrol, maxUnit) {
  const counts = {};
  patrol.forEach(p => counts[p.unit] = (counts[p.unit] || 0) + 1);
  let over = 0;
  for (const u in counts) if (u !== '—' && counts[u] > maxUnit) over += counts[u] - maxUnit;
  return over;
}

function tentGapWarnCount(patrol, maxGap) {
  return assignTents(patrol, maxGap).warnings.filter(w => /exceeds/.test(w)).length;
}

function balanceUnits(patrols, maxUnit, maxGap, minS, maxS) {
  // Reduce >maxUnit-per-unit crowding via 1-for-1 swaps between ANY two patrols.
  // A swap is accepted only if it lowers total unit overage AND does not increase
  // the number of tent age-gap violations. Comparing warning *counts* (rather than
  // demanding zero) means an unavoidable outlier birthdate in a patrol can't block
  // all balancing, while still never trading a unit fix for a new tenting problem.
  for (let pass = 0; pass < 12; pass++) {
    let improved = false;
    for (let i = 0; i < patrols.length; i++) {
      if (unitOverage(patrols[i], maxUnit) === 0) continue;
      const counts = {};
      patrols[i].forEach(p => counts[p.unit] = (counts[p.unit] || 0) + 1);
      const surplusUnits = Object.keys(counts).filter(u => u !== '—' && counts[u] > maxUnit);
      for (const surplusUnit of surplusUnits) {
        const candIdxs = patrols[i].map((p, idx) => p.unit === surplusUnit ? idx : -1).filter(x => x >= 0);
        let done = false;
        for (let j = 0; j < patrols.length && !done; j++) {
          if (j === i) continue;
          for (const ci of candIdxs) {
            const a = patrols[i][ci];
            for (let cj = 0; cj < patrols[j].length; cj++) {
              const b = patrols[j][cj];
              if (b.unit === surplusUnit) continue;
              // tentative swap
              const A = patrols[i].slice(); const B = patrols[j].slice();
              A[ci] = b; B[cj] = a;
              const beforeOver = unitOverage(patrols[i], maxUnit) + unitOverage(patrols[j], maxUnit);
              const afterOver = unitOverage(A, maxUnit) + unitOverage(B, maxUnit);
              const beforeWarn = tentGapWarnCount(patrols[i], maxGap) + tentGapWarnCount(patrols[j], maxGap);
              const afterWarn = tentGapWarnCount(A, maxGap) + tentGapWarnCount(B, maxGap);
              if (afterOver < beforeOver && afterWarn <= beforeWarn) {
                patrols[i] = A; patrols[j] = B; improved = true; done = true; break;
              }
            }
            if (done) break;
          }
        }
      }
    }
    if (!improved) break;
  }
  return patrols;
}

/* ---------- main sort ---------- */
function buildPatrolsForGender(people, opts) {
  const sorted = people.slice().sort((a, b) => (a.birth || 0) - (b.birth || 0));
  const plan = planSizes(sorted.length, opts.minS, opts.maxS);
  // sequential age-banded chunks
  let patrols = [], idx = 0;
  for (const sz of plan.sizes) { patrols.push(sorted.slice(idx, idx + sz)); idx += sz; }
  patrols = balanceUnits(patrols, opts.maxUnit, opts.maxGap, opts.minS, opts.maxS);
  return { patrols, outOfRange: plan.outOfRange };
}

function generate() {
  const { people, issues, ref } = buildParticipants();
  // Tenting gap is capped at 730 days (~2 years) — anything larger is a policy violation.
  let gapDays = +$('ageGap').value;
  if (!Number.isFinite(gapDays)) gapDays = 730;
  gapDays = Math.min(730, Math.max(0, gapDays));
  $('ageGap').value = gapDays; // reflect any clamping back to the UI
  const opts = {
    minS: Math.max(2, +$('minSize').value || 5),
    maxS: Math.max(2, +$('maxSize').value || 8),
    maxUnit: Math.max(1, +$('maxUnit').value || 2),
    maxGap: gapDays,
  };
  if (opts.maxS < opts.minS) opts.maxS = opts.minS;

  const excludeOlder = $('exclude21') && $('exclude21').checked;
  let active = people, excluded = [];
  if (excludeOlder) {
    excluded = people.filter(p => p.age21);
    active = people.filter(p => !p.age21);
  }

  const byGender = { Male: [], Female: [], Unknown: [] };
  active.forEach(p => (byGender[p.gender] || byGender.Unknown).push(p));

  const globalWarn = [];
  if (excluded.length)
    globalWarn.push(`${excluded.length} participant(s) aged 21+ were excluded from patrols (listed at the bottom).`);
  if (byGender.Unknown.length)
    globalWarn.push(`${byGender.Unknown.length} participant(s) have unknown gender and could not be placed in a single-gender patrol. Fix the gender column and regenerate.`);
  if (issues.length) {
    const uniq = [...new Set(issues)];
    globalWarn.push(uniq.slice(0, 8).join(' ') + (uniq.length > 8 ? ` …(+${uniq.length - 8} more)` : ''));
  }

  const result = { ref, genders: [], opts, unplaced: byGender.Unknown, excluded };
  for (const g of ['Male', 'Female']) {
    if (!byGender[g].length) continue;
    const built = buildPatrolsForGender(byGender[g], opts);
    result.genders.push({ gender: g, ...built });
  }

  render(result, globalWarn);
}

/* ---------- rendering ---------- */
let CURRENT = null;

function fmtAge(a) { return a == null ? '?' : a.toFixed(1); }
function fmtDate(d) {
  if (!d) return '?';
  const y = d.getFullYear(), m = String(d.getMonth() + 1).padStart(2, '0'), da = String(d.getDate()).padStart(2, '0');
  return `${y}-${m}-${da}`;
}

function renamePatrol(gi, pi, value) {
  const grp = CURRENT && CURRENT.genders[gi];
  if (!grp || !grp.patrols[pi]) return;
  const v = String(value == null ? '' : value).trim();
  grp.patrols[pi].patrolName = v || PATROL_NAMES[pi % PATROL_NAMES.length];
  render(null, null, false);
}

function moveMember(gi, pid, target) {
  const grp = CURRENT && CURRENT.genders[gi];
  if (!grp) return;
  let person = null;
  for (const patrol of grp.patrols) {
    const idx = patrol.findIndex(p => String(p.id) === String(pid));
    if (idx >= 0) { person = patrol.splice(idx, 1)[0]; break; }
  }
  if (!person) return;
  if (target === 'new') grp.patrols.push([person]);
  else {
    const ti = +target;
    if (grp.patrols[ti]) grp.patrols[ti].push(person);
    else grp.patrols.push([person]);
  }
  grp.patrols = grp.patrols.filter(p => p.length > 0);
  render(null, null, false);
}

function render(result, globalWarn, doScroll = true) {
  if (result) { CURRENT = result; CURRENT.globalWarn = globalWarn; }
  const cur = CURRENT;
  if (!cur) return;

  const gw = $('globalWarnings');
  gw.innerHTML = (cur.globalWarn || []).map(w => `<div class="warn">⚠️ ${w}</div>`).join('');

  const cont = $('resultsContainer');
  cont.innerHTML = '';
  LAST_EXPORT = [];

  cur.genders.forEach((grp, gi) => {
    const section = document.createElement('div');
    section.className = `gender-group ${grp.gender.toLowerCase()}`;
    const youthCount = grp.patrols.reduce((a, p) => a + p.length, 0);
    section.innerHTML = `<h3>${grp.gender} patrols — ${grp.patrols.length} patrol(s), ${youthCount} youth</h3>`;
    const sizeIssue = grp.patrols.some(p => p.length < cur.opts.minS || p.length > cur.opts.maxS);
    if (sizeIssue)
      section.innerHTML += `<div class="warn">⚠️ ${grp.gender}: one or more patrols fall outside ${cur.opts.minS}–${cur.opts.maxS}. Use the Move controls to rebalance.</div>`;

    const grid = document.createElement('div');
    grid.className = 'patrol-grid';
    const npatrols = grp.patrols.length;

    grp.patrols.forEach((patrol, pi) => {
      if (!patrol.patrolName) patrol.patrolName = PATROL_NAMES[pi % PATROL_NAMES.length];
      const name = patrol.patrolName;
      const ages = patrol.map(p => p.age).filter(a => a != null);
      const ageRange = ages.length ? `${fmtAge(Math.min(...ages))}–${fmtAge(Math.max(...ages))} yrs` : 'ages n/a';
      const sizeWarn = (patrol.length < cur.opts.minS || patrol.length > cur.opts.maxS);
      const overage = unitOverage(patrol, cur.opts.maxUnit);

      const unitCounts = {};
      patrol.forEach(p => unitCounts[p.unit] = (unitCounts[p.unit] || 0) + 1);

      const rows = patrol.slice().sort((a, b) => String(a.name).localeCompare(String(b.name), undefined, { sensitivity: 'base' })).map(p => {
        const flag = (p.unit !== '—' && unitCounts[p.unit] > cur.opts.maxUnit) ? ' unit-flag' : '';
        const adultBadge = isAdult(p) ? ' <span class="badge adult">18+</span>' : '';
        return `<tr class="member-row" draggable="true" data-gi="${gi}" data-pi="${pi}" data-pid="${esc(String(p.id))}">
          <td class="drag-cell"><span class="drag-handle" title="Drag to move">⠿</span> ${esc(p.name)}</td>
          <td class="${flag.trim()}">${esc(p.unit)}</td>
          <td>${fmtDate(p.birth)}</td>
          <td>${fmtAge(p.age)}${adultBadge}</td>
          <td class="muted">${esc(p.email)}</td>
        </tr>`;
      }).join('');

      const { tents, warnings } = assignTents(patrol, cur.opts.maxGap);
      const tentHtml = tents.map((t, ti) => {
        const suffix = t.kind === 'solo' ? ' (solo)' : '';
        const cls = t.members.some(isAdult) ? ' adult-tent' : '';
        const who = t.members.map(x => `${esc(x.name)} <span class="muted">(b. ${fmtDate(x.birth)}${isAdult(x) ? ', 18+' : ''})</span>`).join(' &nbsp;·&nbsp; ');
        return `<div class="tent${cls}"><span class="label">Tent ${ti + 1}${suffix}:</span> ${who}</div>`;
      }).join('');

      // export rows
      tents.forEach((t, ti) => t.members.forEach(x => LAST_EXPORT.push({
        patrol: `${grp.gender} - ${name}`, gender: grp.gender, tent: ti + 1,
        name: x.name, unit: x.unit, birth: fmtDate(x.birth), age: x.age == null ? '' : x.age.toFixed(1),
        category: isAdult(x) ? 'Adult (18+)' : 'Youth', email: x.email
      })));

      const card = document.createElement('div');
      card.className = 'patrol';
      card.dataset.gi = String(gi);
      card.dataset.pi = String(pi);
      card.innerHTML = `
        <div class="patrol-head">
          <span class="pname">
            <input class="pname-input no-print" type="text" data-gi="${gi}" data-pi="${pi}" value="${esc(name)}" aria-label="Patrol name" title="Click to rename" />
            <span class="pname-print print-only">${esc(name)} Patrol</span>
          </span>
          <span class="pmeta">${patrol.length} youth · ${ageRange}</span>
        </div>
        <div class="patrol-body">
          ${sizeWarn ? `<div class="err">Size ${patrol.length} is outside ${cur.opts.minS}–${cur.opts.maxS}.</div>` : ''}
          ${overage ? `<div class="warn">More than ${cur.opts.maxUnit} from one unit (couldn't fully separate).</div>` : ''}
          <table class="members">
            <thead><tr><th>Name</th><th>Unit</th><th>Birth date</th><th>Age</th><th>Email</th></tr></thead>
            <tbody>${rows}</tbody>
          </table>
          <div class="tents">
            <h5>Tent assignments</h5>
            ${tentHtml || '<div class="muted small">n/a</div>'}
            ${warnings.map(w => `<div class="warn">⚠️ ${esc(w)}</div>`).join('')}
          </div>
        </div>`;
      grid.appendChild(card);
    });

    const dz = document.createElement('div');
    dz.className = 'patrol new-patrol-zone no-print';
    dz.dataset.gi = String(gi);
    dz.innerHTML = `<div class="dropzone-inner">＋ Drop here to create a new ${esc(grp.gender)} patrol</div>`;
    grid.appendChild(dz);

    section.appendChild(grid);
    cont.appendChild(section);
  });

  if (cur.excluded && cur.excluded.length) {
    const sec = document.createElement('div');
    sec.className = 'gender-group';
    sec.innerHTML = `<h3>Excluded — adults 21 and older (${cur.excluded.length})</h3>` +
      `<div class="patrol-grid"><div class="patrol"><div class="patrol-body"><table class="members"><thead><tr><th>Name</th><th>Unit</th><th>Birth date</th><th>Age</th><th>Email</th></tr></thead><tbody>` +
      cur.excluded.slice().sort((a, b) => (a.birth || 0) - (b.birth || 0)).map(p => `<tr><td>${esc(p.name)}</td><td>${esc(p.unit)}</td><td>${fmtDate(p.birth)}</td><td>${fmtAge(p.age)}</td><td class="muted">${esc(p.email)}</td></tr>`).join('') +
      `</tbody></table></div></div></div>`;
    cont.appendChild(sec);
  }

  if (cur.unplaced.length) {
    const sec = document.createElement('div');
    sec.className = 'gender-group';
    sec.innerHTML = `<h3>Unplaced (unknown gender)</h3>` +
      `<div class="patrol-grid"><div class="patrol"><div class="patrol-body"><table class="members"><thead><tr><th>Name</th><th>Unit</th><th>Email</th></tr></thead><tbody>` +
      cur.unplaced.map(p => `<tr><td>${esc(p.name)}</td><td>${esc(p.unit)}</td><td class="muted">${esc(p.email)}</td></tr>`).join('') +
      `</tbody></table></div></div></div>`;
    cont.appendChild(sec);
  }

  show('step-results');
  if (doScroll) $('step-results').scrollIntoView({ behavior: 'smooth' });
}

/* ---------- export ---------- */
let LAST_EXPORT = [];
function exportCsv() {
  if (!LAST_EXPORT.length) return;
  const csv = Papa.unparse(LAST_EXPORT.map(r => ({
    Patrol: r.patrol, Gender: r.gender, Tent: r.tent, Name: r.name, Unit: r.unit, BirthDate: r.birth, Age: r.age, Category: r.category, Email: r.email
  })));
  const blob = new Blob([csv], { type: 'text/csv' });
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = 'patrols.csv';
  a.click();
  URL.revokeObjectURL(a.href);
}

/* ---------- sample ---------- */
const SAMPLE = `name,email,unit,gender,birthdate
Alex Carter,alex@example.org,101,M,2009-03-14
Ben Diaz,ben@example.org,101,M,2009-07-22
Caleb Fox,caleb@example.org,101,M,2010-01-05
Dylan Grant,dylan@example.org,205,M,2009-11-30
Ethan Hill,ethan@example.org,205,Male,2010-05-18
Finn Jones,finn@example.org,205,M,2008-09-09
Gabe Kim,gabe@example.org,310,M,2009-02-27
Henry Lee,henry@example.org,310,M,2010-08-14
Ian Moore,ian@example.org,412,M,2008-12-01
Jack Nolan,jack@example.org,412,M,2009-06-19
Kyle Owens,kyle@example.org,412,M,2010-03-03
Liam Park,liam@example.org,101,M,2009-10-25
Mason Reed,mason@example.org,205,M,2010-07-07
Noah Stone,noah@example.org,310,M,2008-04-16
Owen Tate,owen@example.org,310,M,2009-09-29
Ava Brooks,ava@example.org,101,F,2009-04-10
Bella Cruz,bella@example.org,101,F,2009-08-21
Chloe Dean,chloe@example.org,205,F,2010-02-12
Dana Ellis,dana@example.org,205,Female,2009-12-05
Ella Frost,ella@example.org,310,F,2008-10-30
Faith Gray,faith@example.org,310,F,2010-06-22
Grace Hunt,grace@example.org,412,F,2009-01-17
Hannah Ives,hannah@example.org,412,F,2010-09-08
Iris James,iris@example.org,101,F,2008-11-26
Jade Klein,jade@example.org,205,F,2009-05-14
Kira Lowe,kira@example.org,310,F,2010-04-01
Lena Maye,lena@example.org,412,F,2009-07-19`;

/* ---------- wiring ---------- */
$('fileInput').addEventListener('change', e => { if (e.target.files[0]) handleFile(e.target.files[0]); });

// Drag-and-drop a CSV/Excel file onto the file tile (or anywhere on the load step).
const ACCEPT_EXT = ['csv', 'tsv', 'txt', 'xlsx', 'xls'];
function acceptableFile(file) {
  if (!file) return false;
  const ext = (file.name.split('.').pop() || '').toLowerCase();
  return ACCEPT_EXT.includes(ext);
}
const fileDrop = $('fileDrop');
['dragenter', 'dragover'].forEach(t => fileDrop.addEventListener(t, e => {
  e.preventDefault(); e.stopPropagation();
  if (e.dataTransfer) e.dataTransfer.dropEffect = 'copy';
  fileDrop.classList.add('drag-over');
}));
['dragleave', 'dragend'].forEach(t => fileDrop.addEventListener(t, e => {
  e.preventDefault(); e.stopPropagation();
  if (t === 'dragleave' && fileDrop.contains(e.relatedTarget)) return;
  fileDrop.classList.remove('drag-over');
}));
fileDrop.addEventListener('drop', e => {
  e.preventDefault(); e.stopPropagation();
  fileDrop.classList.remove('drag-over');
  const file = e.dataTransfer && e.dataTransfer.files && e.dataTransfer.files[0];
  if (!file) return;
  if (!acceptableFile(file)) { alert('Unsupported file type. Please drop a .csv, .tsv, .txt, .xlsx, or .xls file.'); return; }
  handleFile(file);
});
// Prevent the browser from navigating away if a file is dropped outside the zone.
window.addEventListener('dragover', e => { e.preventDefault(); });
window.addEventListener('drop', e => { if (!fileDrop.contains(e.target)) e.preventDefault(); });

$('parsePasteBtn').addEventListener('click', parsePaste);
$('loadSampleBtn').addEventListener('click', () => { $('pasteInput').value = SAMPLE; parsePaste(); });
$('generateBtn').addEventListener('click', generate);
$('regenBtn').addEventListener('click', generate);
$('exportBtn').addEventListener('click', exportCsv);
$('printBtn').addEventListener('click', () => window.print());

// Delegated handler for patrol-name renaming.
$('resultsContainer').addEventListener('change', (e) => {
  const inp = e.target.closest && e.target.closest('.pname-input');
  if (inp) { renamePatrol(+inp.dataset.gi, +inp.dataset.pi, inp.value); }
});
// Commit a patrol rename on Enter (blur triggers the change handler above).
$('resultsContainer').addEventListener('keydown', (e) => {
  if (e.key === 'Enter' && e.target.closest && e.target.closest('.pname-input')) {
    e.preventDefault();
    e.target.blur();
  }
});

// Drag-and-drop to move members between patrols (same gender only).
let DRAG = null;
const rc = $('resultsContainer');
function clearDropTargets() {
  rc.querySelectorAll('.drop-target').forEach(el => el.classList.remove('drop-target'));
}
rc.addEventListener('dragstart', (e) => {
  const tr = e.target.closest && e.target.closest('.member-row');
  if (!tr) return;
  DRAG = { gi: +tr.dataset.gi, pi: +tr.dataset.pi, pid: tr.dataset.pid };
  tr.classList.add('dragging');
  if (e.dataTransfer) {
    e.dataTransfer.effectAllowed = 'move';
    try { e.dataTransfer.setData('text/plain', tr.dataset.pid); } catch (_) {}
  }
});
rc.addEventListener('dragend', (e) => {
  const tr = e.target.closest && e.target.closest('.member-row');
  if (tr) tr.classList.remove('dragging');
  clearDropTargets();
  DRAG = null;
});
rc.addEventListener('dragover', (e) => {
  if (!DRAG) return;
  const tgt = e.target.closest && e.target.closest('.patrol');
  if (!tgt || +tgt.dataset.gi !== DRAG.gi) return;
  const isNew = tgt.classList.contains('new-patrol-zone');
  if (!isNew && +tgt.dataset.pi === DRAG.pi) return; // its own patrol
  e.preventDefault();
  if (e.dataTransfer) e.dataTransfer.dropEffect = 'move';
  if (!tgt.classList.contains('drop-target')) {
    clearDropTargets();
    tgt.classList.add('drop-target');
  }
});
rc.addEventListener('dragleave', (e) => {
  const tgt = e.target.closest && e.target.closest('.patrol');
  if (tgt && !tgt.contains(e.relatedTarget)) tgt.classList.remove('drop-target');
});
rc.addEventListener('drop', (e) => {
  if (!DRAG) return;
  const tgt = e.target.closest && e.target.closest('.patrol');
  if (!tgt || +tgt.dataset.gi !== DRAG.gi) return;
  e.preventDefault();
  const isNew = tgt.classList.contains('new-patrol-zone');
  if (!isNew && +tgt.dataset.pi === DRAG.pi) { DRAG = null; clearDropTargets(); return; }
  const d = DRAG; DRAG = null;
  moveMember(d.gi, d.pid, isNew ? 'new' : String(+tgt.dataset.pi));
});
