// Tiny backend: one endpoint, talks to your Google Sheet with a service account.
const { JWT } = require('google-auth-library');

const SID = process.env.SHEET_ID, PIN = process.env.APP_PIN;
const cred = JSON.parse(process.env.GOOGLE_SERVICE_ACCOUNT_JSON);
const jwt = new JWT({
  email: cred.client_email,
  key: cred.private_key.replace(/\\n/g, '\n'),
  scopes: ['https://www.googleapis.com/auth/spreadsheets'],
});

const T = { c: 'Customers', d: 'Daily Orders', s: 'Subscribers', e: 'Expenses', r: 'Report', n: 'New Customers' };
const W = { c: 10, d: 10, s: 11, e: 7, n: 10 };
const END = { c: 300, d: 100, s: 100, e: 500, n: 100 };
const today = () => new Date().toLocaleDateString('en-CA', { timeZone: 'Asia/Kolkata' });

async function sheets(path, opt = {}) {
  const { token } = await jwt.getAccessToken();
  const r = await fetch('https://sheets.googleapis.com/v4/spreadsheets/' + SID + '/values' + path, {
    ...opt, headers: { Authorization: 'Bearer ' + token, 'Content-Type': 'application/json' },
  });
  const j = await r.json();
  if (!r.ok) throw new Error((j.error && j.error.message) || 'Google Sheets error');
  return j;
}
const get = async rng => (await sheets('/' + encodeURIComponent(rng) + '?valueRenderOption=FORMATTED_VALUE')).values || [];
const put = data => sheets(':batchUpdate', { method: 'POST', body: JSON.stringify({ valueInputOption: 'RAW', data }) });
const cell = (k, a, b) => ({ range: `'${T[k]}'!${a}${b ? ':' + b : ''}` });

// data rows start at row 3 and stop at a "Total..." row
async function read(k) {
  const v = await get(`'${T[k]}'!A3:${String.fromCharCode(64 + W[k])}${END[k]}`), out = [];
  for (let i = 0; i < v.length; i++) {
    const row = Array.from({ length: W[k] }, (_, j) => (v[i][j] == null ? '' : String(v[i][j])));
    if (/^total/i.test(row[0].trim())) break;
    if (row[0].trim()) out.push({ r: i + 3, v: row });
  }
  return out;
}
// first empty row above the Total row (never inserts or moves anything)
async function freeRow(k) {
  const v = await get(`'${T[k]}'!A3:A${END[k]}`);
  for (let i = 0; i < v.length; i++) {
    const a = String((v[i] && v[i][0]) || '').trim();
    if (/^total/i.test(a)) break;
    if (!a) return i + 3;
    if (i === v.length - 1) return i + 4;
  }
  throw new Error('No empty row left in ' + T[k] + '. Add a blank row above its Total row.');
}
const ll = s => {
  const p = String(s).split(/[,\s]+/).filter(Boolean).map(Number);
  if (p.length < 2 || p.some(isNaN)) throw new Error('Location should look like 12.9719, 77.6412');
  return p;
};

const fns = {
  async tabData(t) {
    if (t === 'today') return { daily: await read('d') };
    if (t === 'cust') { const [c, s] = await Promise.all([read('c'), read('s')]); return { customers: c.map(x => x.v), subs: s.map(x => x.v) }; }
    if (t === 'subs') { const [s, c] = await Promise.all([read('s'), read('c')]); return { subs: s.map(x => x.v), names: c.map(x => x.v[1]) }; }
    if (t === 'trial') return { trials: await read('n') };
    if (t === 'exp') return { exp: (await read('e')).map(x => x.v) };
    if (t === 'rep') {
      const v = await get("'Report'!A1:K60"), pad = r => Array.from({ length: 11 }, (_, i) => (r && r[i]) || '');
      const kpi = [], h = pad(v[3]), val = pad(v[4]);
      h.forEach((x, i) => { if (x) kpi.push([x, val[i]]); });
      return { kpi, partners: v.slice(8, 13).map(r => pad(r).slice(0, 9)), cs: v.slice(17).map(r => pad(r).slice(0, 9)).filter(r => r[0]) };
    }
  },
  // Daily Orders: Ordered Today, Number of Orders, Extra Chapati, Speed Delivery (changed rows only)
  async saveOrders(list) {
    await put(list.map(o => {
      const on = !!o.ordered, n = on ? (+o.n || 1) : (+o.n || 0);
      return { ...cell('d', 'F' + o.r, 'I' + o.r), values: [[on ? 'Yes' : 'No', n, +o.extra || 0, o.speed ? 'Yes' : 'No']] };
    }));
  },
  async addTrial(p) {
    if (!p.name || !p.mobile) throw new Error('Name and mobile are required.');
    const [lat, lng] = ll(p.ll), [r, rows] = await Promise.all([freeRow('n'), read('n')]);
    let mx = 0; rows.forEach(x => { const m = /TRIAL-(\d+)/.exec(x.v[0]); if (m) mx = Math.max(mx, +m[1]); });
    await put([{ ...cell('n', 'A' + r, 'I' + r), values: [['TRIAL-' + String(mx + 1).padStart(3, '0'), p.name, p.mobile, p.loc, lat, lng, today(), 49, 'Pending']] }]);
  },
  async setTrialStatus(r, status) { await put([{ ...cell('n', 'I' + r), values: [[status]] }]); },
  async addSub(p) {
    const r = await freeRow('s');
    await put([
      { ...cell('s', 'A' + r), values: [[p.name]] },
      { ...cell('s', 'F' + r), values: [[today()]] },
      { ...cell('s', 'H' + r, 'I' + r), values: [[+p.amount || 949, +p.orders || 15]] },
    ]);
  },
  async addExpense(p) {
    const who = +p.who, amt = +p.amount;
    if (![1, 2, 3].includes(who)) throw new Error('Choose who you are first.');
    if (!amt) throw new Error('Enter the amount paid.');
    const r = await freeRow('e'), pay = [0, 0, 0]; pay[who - 1] = amt;
    await put([
      { ...cell('e', 'A' + r, 'E' + r), values: [[today(), p.desc || '', ...pay]] },
      { ...cell('e', 'G' + r), values: [[p.mode || '']] },
    ]);
  },
};

module.exports = async (req, res) => {
  try {
    if (req.method !== 'POST') return res.status(405).json({ error: 'POST only' });
    if (!PIN || req.headers['x-pin'] !== PIN) return res.status(401).json({ error: 'Wrong PIN' });
    const { fn, args = [] } = req.body || {};
    if (!fns[fn]) return res.status(400).json({ error: 'Unknown action' });
    res.json({ data: (await fns[fn](...args)) || null });
  } catch (e) {
    res.status(400).json({ error: e.message });
  }
};
