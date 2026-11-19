// Supplier stations, shared by the coordinator page and the guests' seat plan so both draw them the same way.
// Where each one sits comes from coordinator.json; the supplier's name and details come from the planner sheet's
// SUPPLIERS DIRECTORY tab (home.exe's Apps Script, ?view=suppliers) whenever it loads.

const STATIONS_SHEET = 'https://script.google.com/macros/s/AKfycbyTuqgUu9wSuGxurX-VFAjPWzx7jsEm-IkYZpl4tld5ngMIGu4EBcFXSL8QnvOOCcZO2A/exec';

// Resolves to [coordinator.json, directory rows]; withSheet false skips the sheet when only the layout is needed.
// Offline, each falls back to the last copy this browser loaded, as venue wifi is unreliable; either can come back null.
function stationsLoad(withSheet = true) {
  const load = key => { try { return JSON.parse(localStorage.getItem(key)); } catch (e) { return null; } };
  const save = (key, value) => { try { localStorage.setItem(key, JSON.stringify(value)); } catch (e) {} };
  const data = fetch('coordinator.json', { cache: 'no-cache' })
    .then(r => r.json()).then(d => { save('coordinator-data', d); return d; })
    .catch(() => load('coordinator-data'));
  const directory = !withSheet ? null : fetch(STATIONS_SHEET + '?view=suppliers')
    .then(r => r.json()).then(d => { if (!d.ok) throw new Error(d.error); save('coordinator-directory', d.suppliers); return d.suppliers; })
    .catch(() => load('coordinator-directory'));
  return Promise.all([data, directory]);
}

// Each station names its row in the directory ("directory"); the sheet's supplier, contact person, arrival time and
// coverage win over the copies in coordinator.json, which only stand in until the sheet loads.
function stationsMerge(list, directory) {
  const key = v => String(v || '').trim().replace(/\s+/g, ' ').toUpperCase();
  const rows = new Map((directory || []).map(r => [key(r.service), r]));
  return (list || []).map((s, n) => {
    const row = s.directory && rows.get(key(s.directory));
    const listed = row && !/^(to follow|tbc|tba|n\/?a|-*)$/i.test(row.name) ? row.name : '';   // "TO FOLLOW" isn't a supplier yet
    const fromSheet = !!(directory && s.directory);
    const supplier = listed || s.supplier || '';
    return { ...s, i: n + 1, supplier, title: supplier || s.type.toLowerCase().replace(/\b\w/g, c => c.toUpperCase()),
      contact: fromSheet ? (listed ? row.contact : '') : s.contact,
      arrive: fromSheet ? (listed ? row.arrive : '') : s.arrive,
      coverage: listed ? row.coverage : '', pending: fromSheet && !listed };
  });
}

// Labelled boxes like the backdrop and LED wall: the supplier's name, with the kind of station underneath.
function stationsDraw(svg, list, interactive) {
  const esc = s => String(s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  const g = document.createElementNS('http://www.w3.org/2000/svg', 'g');
  g.innerHTML = list.map(s => {
    const w = s.w || 110, name = s.supplier ? s.supplier.toUpperCase() : '';
    let names = name ? [name] : [];
    if (name.length * .7 * 7 > w - 12) {   // too long for the box even at the smallest size: break it at the space nearest the middle
      const cut = [...name.matchAll(/ /g)].map(m => m.index).sort((a, b) => Math.abs(a - name.length / 2) - Math.abs(b - name.length / 2))[0];
      if (cut) names = [name.slice(0, cut), name.slice(cut + 1)];
    }
    const lines = names.map(t => ({ t, sub: false })).concat({ t: s.type, sub: !!name });
    const h = Math.max(s.h || 40, lines.length * 13 + 8), top = s.y - (lines.length - 1) * 6.5 + 3.8;   // centre the lines in the box
    const size = l => Math.max(7, Math.min(l.sub ? 8.5 : 10.5, (w - 12) / (l.t.length * .7)));   // a long name from the sheet shrinks to fit
    return '<g class="sup" data-i="' + s.i + '"' + (interactive ? ' tabindex="0" role="button"' : '') + ' aria-label="' + esc(s.title) + '">' +
      '<rect x="' + (s.x - w / 2) + '" y="' + (s.y - h / 2) + '" width="' + w + '" height="' + h + '" rx="4"/>' +
      lines.map((l, n) => '<text' + (l.sub ? ' class="sub"' : '') + ' x="' + s.x + '" y="' + (top + n * 13) + '" style="font-size:' + size(l).toFixed(1) + 'px">' +
        esc(l.t) + '</text>').join('') + '</g>';
  }).join('');
  svg.appendChild(g);
}
