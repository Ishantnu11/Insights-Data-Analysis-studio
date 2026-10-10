import { useState, useMemo, useRef } from 'react'
import { motion, AnimatePresence } from 'framer-motion'
import * as XLSX from 'xlsx'
import { jsPDF } from 'jspdf'
import JSZip from 'jszip'

const M = motion
const AP = AnimatePresence
const COL = ['#7c5cff', '#22d3ee', '#f59e0b', '#10b981', '#f43f5e', '#60a5fa', '#a78bfa', '#fb923c']

const GEMINI_KEY = import.meta.env.VITE_GEMINI_API_KEY || ''

/* ---------- data helpers ---------- */
let D = ','
const splitLine = l => { const r = []; let c = '', q = false; for (const ch of l) { if (ch === '"') q = !q; else if (ch === D && !q) { r.push(c); c = '' } else c += ch } r.push(c); return r }
const toRows = txt => { const L = txt.trim().split(/\r?\n/).filter(Boolean); D = L[0].includes('\t') ? '\t' : ','; const h = splitLine(L[0]).map(s => s.trim()); return L.slice(1).map(l => { const v = splitLine(l), o = {}; h.forEach((k, i) => { const s = (v[i] || '').trim(); o[k] = s !== '' && !isNaN(+s) ? +s : s }); return o }) }
const STOP = new Set('the a an and or of to in is it for on with as at by this that be are was from we you i not have has but they their our'.split(' '))
const fromText = t => { const f = {}; t.toLowerCase().match(/[a-z']{3,}/g)?.forEach(w => { if (!STOP.has(w)) f[w] = (f[w] || 0) + 1 }); return Object.entries(f).sort((a, b) => b[1] - a[1]).slice(0, 40).map(([word, count]) => ({ word, count, length: word.length })) }
const parseAny = (name, txt) => { try { if (/\.json$/i.test(name) || /^\s*[\[{]/.test(txt)) { let j = JSON.parse(txt); if (!Array.isArray(j)) j = Object.values(j).find(Array.isArray) || [j]; return j.map(r => typeof r === 'object' ? r : { value: r }) } } catch (e) { }; if (/\.csv$/i.test(name) || (/[,\t]/.test(txt.split('\n')[0]) && txt.split('\n').length > 2)) return toRows(txt); return fromText(txt) }
const AG = { sum: a => a.reduce((x, y) => x + y, 0), avg: a => a.reduce((x, y) => x + y, 0) / (a.length || 1), count: a => a.length, max: a => Math.max(...a), min: a => Math.min(...a), median: a => { const s = [...a].sort((x, y) => x - y), m = s.length >> 1; return s.length % 2 ? s[m] : (s[m - 1] + s[m]) / 2 }, std: a => { const m = a.reduce((x, y) => x + y, 0) / (a.length || 1); return Math.sqrt(a.reduce((x, y) => x + (y - m) ** 2, 0) / (a.length || 1)) }, p90: a => [...a].sort((x, y) => x - y)[Math.floor(a.length * .9)] ?? 0 }
const profile = rows => Object.keys(rows[0] || {}).map(k => { const v = rows.map(r => r[k]); const nn = v.filter(x => x !== '' && x != null); const num = nn.length && nn.filter(x => typeof x === 'number').length / nn.length > .8; const n = nn.filter(x => typeof x === 'number'); return { name: k, num, missing: v.length - nn.length, uniq: new Set(nn).size, min: num ? Math.min(...n) : null, max: num ? Math.max(...n) : null, mean: num ? n.reduce((a, b) => a + b, 0) / n.length : null, med: num ? AG.median(n) : null, std: num ? AG.std(n) : null, out: num ? (() => { const s = [...n].sort((x, y) => x - y), q1 = s[Math.floor(s.length * .25)], q3 = s[Math.floor(s.length * .75)], i = q3 - q1; return n.filter(x => x < q1 - 1.5 * i || x > q3 + 1.5 * i).length })() : null } })
const group = (rows, x, y, agg, top) => { const g = {}; rows.forEach(r => { const k = r[x]; (g[k] = g[k] || []).push(agg === 'count' ? 1 : +r[y] || 0) }); let o = Object.entries(g).map(([k, v]) => ({ k, v: AG[agg](v) })); const dateLike = o.every(d => /^\d{4}-\d{2}/.test(d.k) || /^(jan|feb|mar)/i.test(d.k)); if (!dateLike) o.sort((a, b) => b.v - a.v); else o.sort((a, b) => a.k > b.k ? 1 : -1); return top ? o.slice(0, top) : o.slice(0, 24) }
const fmt = n => Math.abs(n) >= 1e6 ? (n / 1e6).toFixed(2) + 'M' : Math.abs(n) >= 1e3 ? (n / 1e3).toFixed(1) + 'K' : (+n).toFixed(n % 1 ? 1 : 0)
const trunc = (s, max = 10) => { const str = String(s ?? ''); return str.length > max ? str.slice(0, max - 1) + '…' : str }
const monthNames = ['', 'Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']
const formatTick = (val, n) => {
  const s = String(val ?? '').trim()
  const m = s.match(/^(\d{4})[-/](\d{1,2})$/)
  if (m) {
    const idx = parseInt(m[2], 10)
    if (idx >= 1 && idx <= 12) return monthNames[idx] + (n <= 5 ? " '" + m[1].slice(2) : '')
  }
  const m2 = s.match(/^(\d{4})[-/](\d{1,2})[-/](\d{1,2})$/)
  if (m2) {
    const idx = parseInt(m2[2], 10)
    if (idx >= 1 && idx <= 12) return monthNames[idx] + ' ' + parseInt(m2[3], 10)
  }
  return trunc(s, n <= 4 ? 14 : n <= 8 ? 10 : 7)
}
const bins = (rows, y) => { const v = rows.map(r => +r[y]).filter(n => !isNaN(n)); if (!v.length) return []; const lo = Math.min(...v), hi = Math.max(...v), k = 10, s = (hi - lo) / k || 1, c = Array(k).fill(0); v.forEach(n => c[Math.min(k - 1, Math.floor((n - lo) / s))]++); return c.map((n, i) => ({ k: fmt(lo + i * s), v: n })) }
const xf = (d, t) => { if (!t || t === 'none') return d; const o = d.map(p => ({ ...p })); if (t === 'run') { let s = 0; o.forEach(p => { s += p.v; p.v = s }) } else if (t === 'ma') o.forEach((p, i) => { const q = d.slice(Math.max(0, i - 2), i + 1); p.v = q.reduce((a, b) => a + b.v, 0) / q.length }); else { const lag = t === 'yoy' ? 12 : 1; o.forEach((p, i) => { const b = d[i - lag]; p.v = b && b.v ? (p.v - b.v) / Math.abs(b.v) * 100 : 0 }) }; return o }
const pear = (data, a, b) => { const x = data.map(r => +r[a.name] || 0), y = data.map(r => +r[b.name] || 0); let s = 0, sa = 0, sb = 0; x.forEach((v, k) => { s += (v - a.mean) * (y[k] - b.mean); sa += (v - a.mean) ** 2; sb += (y[k] - b.mean) ** 2 }); return s / Math.sqrt(sa * sb) || 0 }
const corrBg = r => r >= 0 ? `rgba(95,139,62,${r * .55})` : `rgba(244,63,94,${-r * .55})`
const POS = 'good great excellent amazing love loved best fast helpful happy awesome perfect nice easy quick friendly recommend fantastic smooth reliable'.split(' ')
const NEG = 'bad poor terrible awful hate hated worst slow broken useless angry disappointed late expensive buggy difficult rude confusing problem issue refund'.split(' ')
const sentiment = t => t.split(/(?<=[.!?])\s+|\n+/).map(s => s.trim()).filter(Boolean).map(s => { const w = s.toLowerCase().match(/[a-z']+/g) || []; let sc = 0; w.forEach((x, i) => { const ng = /^(not|no|never|isn't|wasn't|don't)$/.test(w[i - 1] || ''); if (POS.includes(x)) sc += ng ? -1 : 1; if (NEG.includes(x)) sc += ng ? 1 : -1 }); return { sentiment: sc > 0 ? 'Positive' : sc < 0 ? 'Negative' : 'Neutral', score: sc, sentence: s } })
const summ = (cols, data) => { const n = cols.filter(c => c.num).slice(0, 4), o = {}; cols.filter(c => !c.num && c.uniq > 1 && c.uniq <= 15).forEach(c => { o[c.name] = Object.fromEntries(n.map(nn => [nn.name, Object.fromEntries(group(data, c.name, nn.name, 'sum', 15).map(p => [p.k, Math.round(p.v)]))])) }); return JSON.stringify(o) }
const fltRows = (rows, fs) => (fs || []).reduce((r, f) => r.filter(x => { const a = x[f.column], b = isNaN(+f.value) ? f.value : +f.value; switch (f.op) { case '!=': return a != b; case '>': return a > b; case '<': return a < b; case '>=': return a >= b; case '<=': return a <= b; case 'contains': return String(a).toLowerCase().includes(String(b).toLowerCase()); default: return a == b } }), rows)

/* ---------- sample data ---------- */
const sample = (() => { const R = ['North', 'South', 'East', 'West'], C = ['Electronics', 'Furniture', 'Office', 'Apparel']; let s = 7; const rnd = () => (s = (s * 9301 + 49297) % 233280) / 233280; const o = []; for (let m = 1; m <= 12; m++)for (const r of R) for (const c of C.slice(0, 3 + (m % 2))) { const u = Math.round(20 + rnd() * 80); const p = +(20 + rnd() * 180).toFixed(0); o.push({ Month: '2025-' + String(m).padStart(2, '0'), Region: r, Category: c, Units: u, Sales: u * p, Profit: Math.round(u * p * (.05 + rnd() * .3)) }) } return o })()

/* ---------- export helpers ---------- */
const css = n => getComputedStyle(document.documentElement).getPropertyValue(n).trim()
const svgCv = svg => new Promise((ok, no) => { const s = new XMLSerializer().serializeToString(svg).replace(/var\((--[a-z0-9]+)\)/g, (_, v) => css(v)); const vb = svg.viewBox.baseVal, sc = 3, cv = document.createElement('canvas'); cv.width = vb.width * sc; cv.height = vb.height * sc; const im = new Image(); im.onload = () => { const c = cv.getContext('2d'); c.fillStyle = css('--panel'); c.fillRect(0, 0, cv.width, cv.height); c.drawImage(im, 0, 0, cv.width, cv.height); ok(cv) }; im.onerror = no; im.src = 'data:image/svg+xml;charset=utf-8,' + encodeURIComponent(s) })
const save2 = async (filename, data) => { const bl = data instanceof Blob ? data : new Blob([data]), u = URL.createObjectURL(bl), el = document.createElement('a'); el.href = u; el.download = filename; document.body.appendChild(el); el.click(); el.remove(); setTimeout(() => URL.revokeObjectURL(u), 2000) }
const toCSV = r => { const k = Object.keys(r[0] || {}); return [k.join(',')].concat(r.map(o => k.map(c => { const v = String(o[c] ?? ''); return /[",\n]/.test(v) ? '"' + v.replace(/"/g, '""') + '"' : v }).join(','))).join('\n') }
const readAny = (f, cb) => { const r = new FileReader(); if (/\.xlsx?$/i.test(f.name)) { r.onload = () => { const wb = XLSX.read(r.result, { type: 'array' }); cb(XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { defval: '' })) }; r.readAsArrayBuffer(f) } else { r.onload = () => cb(parseAny(f.name, r.result)); r.readAsText(f) } }
const report = name => { const esc = s => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;'); const cards = [...document.querySelectorAll('.grid>.card')].map(c => { const t = (c.querySelector('.mut') || {}).textContent || '', sv = c.querySelector('svg'), k = c.querySelector('.kpi'); const body = sv ? new XMLSerializer().serializeToString(sv).replace(/var\((--[a-z0-9]+)\)/g, (_, v) => css(v)) : k ? '<div class="k">' + k.innerHTML + '</div>' : (c.querySelector('table') || { outerHTML: '' }).outerHTML; return '<div class="c"><div class="t">' + esc(t) + '</div>' + body + '</div>' }).join(''); return '<!DOCTYPE html><html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>' + esc(name) + ' – report</title><style>body{font-family:system-ui,sans-serif;background:#f0f0f0;color:#141414;margin:0;padding:20px}h1{margin:0 0 4px}.g{display:grid;grid-template-columns:repeat(auto-fill,minmax(min(100%,300px),1fr));gap:12px;margin-top:16px}.c{background:#fff;border-radius:18px;padding:14px;overflow:auto}.t,.mut{color:#77777f;font-size:12px}.k{font-size:30px;font-weight:800}svg{width:100%}table{border-collapse:collapse;font-size:12px}td,th{padding:4px 8px;border-bottom:1px solid #eee;text-align:left}</style></head><body><h1>' + esc(name) + '</h1><div class="t">Generated ' + new Date().toLocaleString() + ' with InsightForge</div><div class="g">' + cards + '</div></body></html>' }

/* ---------- chart title helper ---------- */
const chartTitle = w => w.title || (w.tf && w.tf !== 'none' ? w.tf + ' of ' + w.y + ' by ' + w.x : w.type === 'stack' || w.type === 'heat' ? w.type + ' ' + w.y + ' by ' + w.x + ' × ' + w.z : w.type === 'box' ? 'Spread of ' + w.y + ' by ' + w.x : w.type === 'hist' ? 'Distribution of ' + w.y : w.type === 'kpi' ? `${w.agg} ${w.y}` : w.type === 'table' ? 'Data table' : `${w.agg} of ${w.y} by ${w.x}`)

/* ---------- Chart component ---------- */
function Chart({ w, rows, onPick }) {
  const { type, x, y, agg, z, tgt } = w
  if (type === 'kpi') {
    const v = AG[agg](rows.map(r => agg === 'count' ? 1 : +r[y] || 0))
    return <M.div initial={{ scale: .8, opacity: 0 }} animate={{ scale: 1, opacity: 1 }} className="kpi">{fmt(v)}{tgt ? <span style={{ fontSize: 13, marginLeft: 8, color: v >= tgt ? '#10b981' : '#f43f5e' }}>{v >= tgt ? '▲' : '▼'} {(v / tgt * 100).toFixed(0)}% of target</span> : null}<div className="mut" style={{ fontSize: 12, fontWeight: 400 }}>{agg} of {agg === 'count' ? 'rows' : y}</div></M.div>
  }
  if (type === 'table') return <div className="scroll" style={{ maxHeight: 220 }}><table><thead><tr>{Object.keys(rows[0] || {}).map(k => <th key={k}>{k}</th>)}</tr></thead><tbody>{rows.slice(0, 30).map((r, i) => <tr key={i}>{Object.values(r).map((v, j) => <td key={j}>{v}</td>)}</tr>)}</tbody></table></div>
  if (type === 'stack' || type === 'heat') {
    const XS = [...new Set(rows.map(r => r[x]))].slice(0, 12), ZS = [...new Set(rows.map(r => r[z]))].slice(0, 8)
    if (!z) return <div className="mut">Pick a series column.</div>
    const m = XS.map(a => ZS.map(b => { const v = rows.filter(r => r[x] === a && r[z] === b).map(r => agg === 'count' ? 1 : +r[y] || 0); return v.length ? AG[agg](v) : 0 })), cw = 240 / XS.length, ch = 140 / ZS.length
    if (type === 'heat') {
      const mh = Math.max(...m.flat(), 1)
      return <svg viewBox="0 0 300 180" width="100%">{XS.map((a, i) => ZS.map((b, j) => <M.rect key={i + '-' + j} x={50 + i * cw} y={10 + j * ch} width={cw - 1} height={ch - 1} rx="2" fill={COL[0]} initial={{ opacity: 0 }} animate={{ opacity: .08 + .92 * m[i][j] / mh }} transition={{ delay: (i + j) * .02 }} />))}{ZS.map((b, j) => <text key={j} x="46" y={10 + j * ch + ch / 2 + 3} fontSize="8" fill="var(--mut)" textAnchor="end">{trunc(b, 8)}</text>)}{XS.map((a, i) => <text key={i} x={50 + i * cw + cw / 2} y="165" fontSize="8" fill="var(--mut)" textAnchor="middle">{trunc(a, 8)}</text>)}</svg>
    }
    const ms = Math.max(...m.map(r => r.reduce((a, b) => a + b, 0)), 1)
    return <svg viewBox="0 0 300 180" width="100%">{XS.map((a, i) => { const bw = Math.min(cw * .7, 40), cx = 50 + i * cw + cw / 2; let y0 = 150; return <g key={i}>{ZS.map((b, j) => { const h = m[i][j] / ms * 130; y0 -= h; return <M.rect key={j} x={cx - bw / 2} width={bw} fill={COL[j % 8]} initial={{ height: 0, y: 150 }} animate={{ height: h, y: y0 }} /> })}<text x={cx} y="165" fontSize="8" fill="var(--mut)" textAnchor="middle">{trunc(a, 8)}</text></g> })}{ZS.map((b, j) => <g key={j}><rect x="4" y={8 + j * 12} width="7" height="7" fill={COL[j % 8]} /><text x="14" y={15 + j * 12} fontSize="7.5" fill="var(--mut)">{trunc(b, 8)}</text></g>)}</svg>
  }
  if (type === 'box') {
    const S = [...new Set(rows.map(r => r[x]))].slice(0, 10).map(g => { const v = rows.filter(r => r[x] === g).map(r => +r[y]).filter(n => !isNaN(n)).sort((a, b) => a - b), q = p => v[Math.floor((v.length - 1) * p)]; return { g, lo: q(0), q1: q(.25), md: q(.5), q3: q(.75), hi: q(1) } }).filter(s => s.hi !== undefined), mb = Math.max(...S.map(s => s.hi), 1), Yb = v => 150 - v / mb * 130
    return <svg viewBox="0 0 300 180" width="100%">{S.map((s, i) => { const cx = 40 + (i + .5) * 250 / S.length, bw = Math.min(250 / S.length * .5, 26); return <g key={i}><line x1={cx} x2={cx} y1={Yb(s.lo)} y2={Yb(s.hi)} stroke="var(--mut)" /><M.rect x={cx - bw / 2} y={Yb(s.q3)} width={bw} height={Math.max(1, Yb(s.q1) - Yb(s.q3))} rx="3" fill={COL[i % 8]} initial={{ opacity: 0 }} animate={{ opacity: .8 }} /><line x1={cx - bw / 2} x2={cx + bw / 2} y1={Yb(s.md)} y2={Yb(s.md)} stroke="#fff" strokeWidth="2" /><text x={cx} y="165" fontSize="8" fill="var(--mut)" textAnchor="middle">{trunc(s.g, 8)}</text></g> })}</svg>
  }
  if (type === 'funnel') {
    const F = group(rows, x, y, agg, 8), mf = Math.max(...F.map(p => p.v), 1), h = 170 / (F.length || 1) - 3
    return <svg viewBox="0 0 300 180" width="100%">{F.map((p, i) => <g key={i}><M.rect x={150 - Math.max(20, p.v / mf * 220) / 2} y={5 + i * (h + 3)} width={Math.max(20, p.v / mf * 220)} height={h} rx="4" fill={COL[i % 8]} initial={{ opacity: 0 }} animate={{ opacity: 1 }} transition={{ delay: i * .06 }} /><text x="150" y={5 + i * (h + 3) + h / 2 + 3} fontSize="9" textAnchor="middle" fill="#fff">{trunc(p.k, 12)} · {fmt(p.v)}</text></g>)}</svg>
  }
  if (type === 'scatter') {
    const P = rows.slice(0, 300).map(r => [+r[x], +r[y]]).filter(p => !isNaN(p[0]) && !isNaN(p[1])); const mx = Math.max(...P.map(p => p[0]), 1), my = Math.max(...P.map(p => p[1]), 1)
    return <svg viewBox="0 0 300 180" width="100%">{P.map((p, i) => <M.circle key={i} initial={{ r: 0 }} animate={{ r: 3 }} transition={{ delay: i * .002 }} cx={20 + p[0] / mx * 270} cy={160 - p[1] / my * 150} fill={COL[0]} opacity=".6" />)}<line x1="20" y1="160" x2="290" y2="160" stroke="var(--line)" /><line x1="20" y1="10" x2="20" y2="160" stroke="var(--line)" /><text x="150" y="176" fontSize="9" fill="var(--mut)" textAnchor="middle">{x} vs {y}</text></svg>
  }
  const d = xf(type === 'hist' ? bins(rows, y) : group(rows, x, y, agg, w.top), w.tf); const mx = Math.max(...d.map(p => p.v), 1); const n = d.length || 1
  if (type === 'pie') {
    const t = d.reduce((a, b) => a + Math.abs(b.v), 0) || 1; let a0 = -Math.PI / 2
    return <svg viewBox="0 0 300 180" width="100%">{d.slice(0, 8).map((p, i) => { const a = Math.abs(p.v) / t * Math.PI * 2, a1 = a0 + a, c = [150 + 70 * Math.cos(a0), 90 + 70 * Math.sin(a0), 150 + 70 * Math.cos(a1), 90 + 70 * Math.sin(a1)]; const path = `M150 90L${c[0]} ${c[1]}A70 70 0 ${a > Math.PI ? 1 : 0} 1 ${c[2]} ${c[3]}Z`; a0 = a1; return <M.path key={i} onClick={() => onPick && onPick(x, p.k)} d={path} fill={COL[i % 8]} initial={{ opacity: 0, scale: .6 }} animate={{ opacity: 1, scale: 1 }} transition={{ delay: i * .06 }} style={{ transformOrigin: '150px 90px' }} /> })}<circle cx="150" cy="90" r="34" fill="var(--panel)" />{d.slice(0, 8).map((p, i) => <g key={i}><rect x="232" y={10 + i * 14} width="8" height="8" fill={COL[i % 8]} /><text x="244" y={18 + i * 14} fontSize="8.5" fill="var(--mut)">{trunc(p.k, 10)} {(Math.abs(p.v) / t * 100).toFixed(0)}%</text></g>)}</svg>
  }
  const X = i => 36 + (i + .5) * (268 / n), Y = v => 145 - v / mx * 125
  const pts = d.map((p, i) => `${X(i)},${Y(p.v)}`).join(' ')
  const isAngled = n > 6
  const step = n > 20 ? Math.ceil(n / 10) : n > 14 ? 2 : 1
  return <svg viewBox="0 0 320 205" width="100%">
    {[0, .5, 1].map(f => <g key={f}><line x1="36" x2="305" y1={145 - f * 125} y2={145 - f * 125} stroke="var(--line)" strokeDasharray="2 3" /><text x="32" y={148 - f * 125} fontSize="8" fill="var(--mut)" textAnchor="end">{fmt(mx * f)}</text></g>)}
    <line x1="36" x2="305" y1="145" y2="145" stroke="var(--line)" />
    {(type === 'bar' || type === 'hist') && d.map((p, i) => { const bw = Math.min(Math.max(8, 268 / n * .7), 36), h = Math.max(0, 145 - Y(p.v)); return <M.rect key={i} onClick={() => onPick && onPick(x, p.k)} style={{ cursor: onPick ? 'pointer' : 'default' }} x={X(i) - bw / 2} width={bw} rx="4" fill={COL[i % 8]} initial={{ height: 0, y: 145 }} animate={{ height: h, y: 145 - h }} transition={{ delay: i * .03, type: 'spring', stiffness: 120 }}><title>{`${String(p.k)}: ${fmt(p.v)}`}</title></M.rect> })}
    {(type === 'line' || type === 'area') && <>{type === 'area' && <M.polygon points={`${X(0)},145 ${pts} ${X(n - 1)},145`} fill={COL[1]} initial={{ opacity: 0 }} animate={{ opacity: .22 }} />}<M.polyline points={pts} fill="none" stroke={COL[type === 'area' ? 1 : 0]} strokeWidth="2.5" initial={{ pathLength: 0 }} animate={{ pathLength: 1 }} transition={{ duration: 1 }} />{d.map((p, i) => <circle key={i} cx={X(i)} cy={Y(p.v)} r="2.5" fill={COL[0]}><title>{`${String(p.k)}: ${fmt(p.v)}`}</title></circle>)}</>}
    {d.map((p, i) => (i % step === 0) && <text key={i} x={X(i)} y={isAngled ? 164 : 162} fontSize={n > 8 ? "7.5" : "8"} fill="var(--mut)" textAnchor={isAngled ? "end" : "middle"} transform={isAngled ? `rotate(-40 ${X(i)} 164)` : undefined}><title>{`${String(p.k)}: ${fmt(p.v)}`}</title>{formatTick(p.k, n)}</text>)}
  </svg>
}

/* ---------- interpret natural language ---------- */
function interpret(msg, cols) {
  const low = msg.toLowerCase()
  const type = /pie|donut|share|proportion/.test(low) ? 'pie' : /scatter|correlat|relationship/.test(low) ? 'scatter' : /area/.test(low) ? 'area' : /line|trend|over time|monthly/.test(low) ? 'line' : /kpi|total|overall|card/.test(low) ? 'kpi' : /table/.test(low) ? 'table' : 'bar'
  const agg = /avg|average|mean/.test(low) ? 'avg' : /count|number of|how many/.test(low) ? 'count' : /max|highest|peak/.test(low) ? 'max' : /min|lowest/.test(low) ? 'min' : 'sum'
  const found = cols.map(c => ({ c, i: low.indexOf(c.name.toLowerCase()) })).filter(o => o.i >= 0).sort((a, b) => a.i - b.i).map(o => o.c)
  const nums = cols.filter(c => c.num), cats = cols.filter(c => !c.num)
  const y = (found.find(c => c.num) || nums[0] || {}).name
  let x = (found.find(c => !c.num) || (type === 'line' || type === 'area' ? cats.find(c => /date|month|year|time/i.test(c.name)) : null) || cats[0] || {}).name
  if (type === 'scatter') return { type, x: (found.filter(c => c.num)[1] || nums[1] || nums[0]).name, y, agg }
  const top = (low.match(/top\s*(\d+)/) || [])[1]
  return { type, x, y, agg, top: top ? +top : undefined }
}

/* ---------- Home screen ---------- */
function Home({ data, cols, go, ask }) {
  const [q, setQ] = useState('')
  const nu = cols.find(c => c.num), ca = cols.find(c => !c.num)
  const g = nu && ca ? group(data, ca.name, nu.name, 'sum', 7) : []; const mx = Math.max(...g.map(p => p.v), 1)
  const tot = nu ? AG.sum(data.map(r => +r[nu.name] || 0)) : 0
  const ap = d => ({ initial: { opacity: 0, y: 12, scale: .988 }, animate: { opacity: 1, y: 0, scale: 1 }, transition: { duration: .7, delay: d, ease: [.16, 1, .3, 1] } })
  return <div className="mosaic">
    <M.section className="card m-notif" {...ap(.1)} aria-label="Status"><M.div className="toast" initial={{ opacity: 0, y: 10 }} animate={{ opacity: 1, y: 0 }} transition={{ delay: .4, duration: .6 }}><span className="spk">✦</span><div><b>Dashboard ready!</b><p>{data.length} rows · {cols.length} columns analyzed</p></div><time>now</time></M.div></M.section>
    <M.section className="card m-conn" {...ap(.2)} aria-label="Data sources"><h2>Bring your<br />Data Now.</h2><p className="ml">CSV, Excel, JSON &amp; free text</p><div className="chips2">{['CSV', 'Excel', 'JSON', 'Free text', 'Joins', 'Calc columns'].map((t, i) => <M.span key={t} className="chip" initial={{ opacity: 0, y: 11 }} animate={{ opacity: 1, y: 0 }} transition={{ delay: .85 + i * .05 }}>{t}</M.span>)}</div></M.section>
    <M.section className="card m-auto" {...ap(.04)} aria-label="Analyze your data"><h2 className="hero"><span className="g">Analyze</span> your data.<br />Focus on what matters.</h2><p className="ml">Dashboards, joins and AI-built charts in one studio.</p><div className="row" style={{ marginTop: 16, position: 'relative', zIndex: 3 }}><button className="b pri" onClick={() => go('dash')}>Open dashboard</button><button className="b" onClick={() => go('data')}>Import data</button></div>{[1, 2].map(k => <M.div key={k} className={'hw hw' + k} initial={{ opacity: 0, y: 16 }} animate={{ opacity: 1, y: 0 }} transition={{ delay: .4 + k * .1, duration: .8 }}><div className="tb" /><div className="wb2">{(g.length ? g : [{ v: 5 }, { v: 8 }, { v: 4 }, { v: 9 }]).slice(0, 6).map((p, i) => <i key={i} style={{ height: (p.v / mx * 80 + 10) + '%' }} />)}</div></M.div>)}<M.div className="hpill" initial={{ opacity: 0 }} animate={{ opacity: 1 }} transition={{ delay: .8 }}>✓ Workflow Automated</M.div></M.section>
    <M.section className="card m-ins" {...ap(.14)} aria-label="Live insights"><span className="tag">Live Insights</span><h2 className="big">{fmt(tot)}</h2><p className="ml">total {nu ? nu.name : 'value'}{ca ? ' by ' + ca.name : ''}</p><div className="mbars">{g.map((p, i) => <div key={i} className="mc"><M.div className={'mb' + (p.v === mx ? ' hot' : '')} initial={{ height: 0 }} animate={{ height: Math.max(26, p.v / mx * 210) }} transition={{ delay: .5 + i * .06, type: 'spring', stiffness: 110 }}>{fmt(p.v)}</M.div><span>{trunc(p.k, 8)}</span></div>)}</div></M.section>
    <M.section className="card m-search" {...ap(.28)} aria-label="Ask"><h2>Ask anything<br />about your data</h2><div className="bs"><span className="mg">⌕</span><input value={q} onChange={e => setQ(e.target.value)} onKeyDown={e => e.key === 'Enter' && q.trim() && ask(q)} placeholder="e.g. average profit by region as a pie chart" /></div></M.section>
  </div>
}

/* ---------- Grid ---------- */
function Grid({ rows }) {
  const [q, setQ] = useState(''), [s, setS] = useState({ k: '', d: 1 }), [p, setP] = useState(0)
  const ks = Object.keys(rows[0] || {})
  const f = useMemo(() => { let r = q ? rows.filter(x => Object.values(x).some(v => String(v).toLowerCase().includes(q.toLowerCase()))) : rows; if (s.k) r = [...r].sort((a, b) => (a[s.k] > b[s.k] ? 1 : a[s.k] < b[s.k] ? -1 : 0) * s.d); return r }, [rows, q, s])
  const n = Math.ceil(f.length / 25) || 1, pg = Math.min(p, n - 1)
  return <div><div className="row"><input placeholder="Search all columns…" value={q} onChange={e => { setQ(e.target.value); setP(0) }} /><span className="mut">{f.length} rows</span><span style={{ flex: 1 }} /><button className="b" disabled={pg < 1} onClick={() => setP(pg - 1)}>‹</button><span>{pg + 1}/{n}</span><button className="b" disabled={pg >= n - 1} onClick={() => setP(pg + 1)}>›</button></div>
    <div className="scroll" style={{ marginTop: 8 }}><table><thead><tr>{ks.map(k => <th key={k} style={{ cursor: 'pointer' }} onClick={() => setS({ k, d: s.k === k ? -s.d : 1 })}>{k}{s.k === k ? (s.d > 0 ? ' ▲' : ' ▼') : ''}</th>)}</tr></thead><tbody>{f.slice(pg * 25, pg * 25 + 25).map((r, i) => <tr key={i}>{ks.map(k => <td key={k} style={{ color: typeof r[k] === 'number' && r[k] < 0 ? '#f43f5e' : 'inherit' }}>{r[k]}</td>)}</tr>)}</tbody></table></div></div>
}

/* ---------- Pivot ---------- */
function Pivot({ rows, cols }) {
  const ca = cols.filter(c => !c.num).map(c => c.name), nm = cols.filter(c => c.num).map(c => c.name)
  const [a, setA] = useState(''), [b, setB] = useState(''), [v, setV] = useState(''), [g, setG] = useState('sum')
  const A = a || ca[0], V = v || nm[0]
  if (!A) return <div className="mut">Needs at least one text column.</div>
  const R = [...new Set(rows.map(r => r[A]))].sort().slice(0, 50), C = b ? [...new Set(rows.map(r => r[b]))].sort().slice(0, 12) : ['']
  const val = f => { const x = rows.filter(f).map(r => g === 'count' ? 1 : +r[V] || 0); return x.length ? AG[g](x) : null }
  const cell = (rv, cv) => val(r => r[A] === rv && (!b || r[b] === cv))
  const all = R.flatMap(r => C.map(c => cell(r, c))).filter(x => x != null), mx = Math.max(...all.map(Math.abs), 1)
  const sel = (l, x, set, o, none) => <label className="mut">{l} <select value={x} onChange={e => set(e.target.value)}>{none && <option value="">{none}</option>}{o.map(k => <option key={k}>{k}</option>)}</select></label>
  return <div><div className="row">{sel('Rows', A, setA, ca)}{sel('Columns', b, setB, ca, '(none)')}{sel('Values', V, setV, nm.length ? nm : cols.map(c => c.name))}{sel('Aggregate', g, setG, Object.keys(AG))}</div>
    <div className="scroll" style={{ marginTop: 10, maxHeight: 420 }}><table><thead><tr><th>{A}</th>{b ? C.map(c => <th key={c}>{c}</th>) : <th>{g} {V}</th>}{b && <th>Total</th>}</tr></thead><tbody>{R.map(r => <tr key={r}><td><b>{r}</b></td>{C.map(c => { const x = cell(r, c); return <td key={c} style={{ background: x == null ? 'none' : `rgba(95,139,62,${.08 + .45 * Math.abs(x) / mx})` }}>{x == null ? '' : fmt(x)}</td> })}{b && <td><b>{fmt(val(z => z[A] === r) || 0)}</b></td>}</tr>)}<tr><td><b>Total</b></td>{C.map(c => <td key={c}><b>{fmt(val(z => !b || z[b] === c) || 0)}</b></td>)}{b && <td><b>{fmt(val(() => true) || 0)}</b></td>}</tr></tbody></table></div></div>
}

/* ---------- App ---------- */
export default function App() {
  const [tab, setTab] = useState('home')
  const [data, setData] = useState(sample)
  const [name, setName] = useState('Sample: Retail Sales')
  const [ws, setWs] = useState(() => [{ id: 1, type: 'kpi', y: 'Sales', agg: 'sum' }, { id: 2, type: 'kpi', y: 'Profit', agg: 'sum' }, { id: 3, type: 'kpi', y: 'Units', agg: 'avg' }, { id: 4, type: 'line', x: 'Month', y: 'Sales', agg: 'sum' }, { id: 5, type: 'bar', x: 'Region', y: 'Profit', agg: 'sum' }, { id: 6, type: 'pie', x: 'Category', y: 'Sales', agg: 'sum' }])
  const [flt, setFlt] = useState({ c: '', v: '' })
  const [paste, setPaste] = useState('')
  const [chat, setChat] = useState([{ r: 'a', t: 'Hi! Tell me what you need, e.g. "average profit by region as a pie chart" or "top 5 categories by sales". I can build it straight onto your dashboard.' }])
  const [inp, setInp] = useState('')
  const [done, setDone] = useState({})
  const idr = useRef(100)
  const cols = useMemo(() => profile(data), [data])
  const [rng, setRng] = useState({ a: '', b: '' })
  const dc = cols.find(c => !c.num && /date|month|year|time|period/i.test(c.name))
  const dvals = useMemo(() => dc ? [...new Set(data.map(r => String(r[dc.name])))].sort() : [], [data, dc])
  const rows = useMemo(() => data.filter(r => (!(flt.c && flt.v) || String(r[flt.c]) === flt.v) && (!dc || ((!rng.a || String(r[dc.name]) >= rng.a) && (!rng.b || String(r[dc.name]) <= rng.b)))), [data, flt, rng, dc])
  const cats = cols.filter(c => !c.num), nums = cols.filter(c => c.num)
  const insights = useMemo(() => { const o = [], n = cols.filter(c => c.num), c = cols.filter(c => !c.num && c.uniq > 1 && c.uniq < 20); if (n[0] && c[0]) { const g = group(data, c[0].name, n[0].name, 'sum'); const t = g.reduce((a, b) => a + b.v, 0) || 1; if (g[0]) o.push(g[0].k + ' drives ' + (g[0].v / t * 100).toFixed(0) + '% of total ' + n[0].name + ' (by ' + c[0].name + ').') } for (let i = 0; i < n.length; i++)for (let j = i + 1; j < n.length; j++) { const r = pear(data, n[i], n[j]); if (Math.abs(r) > .5) o.push(n[i].name + ' and ' + n[j].name + ' are ' + (r > 0 ? 'positively' : 'negatively') + ' correlated (r=' + r.toFixed(2) + ') — try a scatter chart.') } cols.forEach(c => c.missing && o.push(c.name + ' has ' + c.missing + ' missing values — clean before reporting.')); return o }, [data, cols])
  const add = w => setWs(s => [...s, { ...w, id: ++idr.current }])
  const load = (n, t) => { const r = Array.isArray(t) ? t : parseAny(n, t); if (r.length) { setData(r); setName(n); setFlt({ c: '', v: '' }); const p = profile(r), nu = p.find(c => c.num), ca = p.find(c => !c.num); setWs([nu && { id: 1, type: 'kpi', y: nu.name, agg: 'sum' }, nu && ca && { id: 2, type: 'bar', x: ca.name, y: nu.name, agg: 'sum' }, nu && ca && { id: 3, type: 'pie', x: ca.name, y: nu.name, agg: 'sum' }, { id: 4, type: 'table' }].filter(Boolean)); setTab('dash') } }
  const [nw, setNw] = useState({ type: 'bar', x: '', y: '', agg: 'sum', z: '', tf: 'none', tgt: '' })
  const [msg, setMsg] = useState('')
  const [extra, setExtra] = useState([])
  const [hist, setHist] = useState([])
  const [cl, setCl] = useState({ col: '', v: '', op: '=' })
  const mod = f => { setHist(h => [...h.slice(-9), data]); setData(f(data)) }
  const undo = () => { if (!hist.length) return; setData(hist[hist.length - 1]); setHist(h => h.slice(0, -1)) }
  const dedupe = d => { const s = new Set(); return d.filter(r => { const k = JSON.stringify(r); return s.has(k) ? false : (s.add(k), true) }) }
  const numFn = s => { const t = String(s).replace(/[,$%\s]/g, ''); return t === '' || isNaN(+t) ? '' : +t }
  const [ai, setAi] = useState(true)
  const [jn, setJn] = useState({ rows: null, a: '', b: '', t: 'left' })
  const dg = useRef(null)
  const drop = id => { const f = dg.current; dg.current = null; if (f == null || f === id) return; setWs(s => { const a = s.find(z => z.id === f), c = s.filter(z => z.id !== f), i = c.findIndex(z => z.id === id); c.splice(i, 0, a); return c }) }
  const rs = (e, w) => { e.preventDefault(); const c = e.currentTarget.parentElement, cw = c.offsetWidth, ch = c.offsetHeight, sx = e.clientX, sy = e.clientY; const mvf = ev => { const sp = Math.max(1, Math.min(4, Math.round((cw + ev.clientX - sx + 14) / 314))); const h = Math.max(110, ch - 50 + ev.clientY - sy); setWs(s => s.map(z => z.id === w.id ? { ...z, span: sp, h } : z)) }; const up = () => { window.removeEventListener('pointermove', mvf); window.removeEventListener('pointerup', up) }; window.addEventListener('pointermove', mvf); window.addEventListener('pointerup', up) }
  const png = async id => { const s = document.querySelector('#w' + id + ' svg'); if (!s) return alert('PNG export works for charts only.'); const cv = await svgCv(s).catch(() => null); if (!cv) return alert('Could not render this chart.'); save2('chart-' + id + '.png', await new Promise(r => cv.toBlob(r))) }
  const pdf = async ret => { try { const d = new jsPDF({ orientation: 'l', unit: 'pt', format: 'a4' }); const sv = [...document.querySelectorAll('.grid svg')]; d.setFontSize(18); d.text('InsightForge report - ' + name, 30, 36); for (let i = 0; i < sv.length; i++) { const k = i % 4; if (i && k === 0) d.addPage(); const cv = await svgCv(sv[i]); d.addImage(cv.toDataURL('image/png'), 'PNG', 30 + (k % 2) * 400, (i < 4 ? 56 : 30) + Math.floor(k / 2) * 250, 380, 228) } const bl = d.output('blob'); if (ret) return bl; save2('dashboard.pdf', bl) } catch (e) { alert('PDF failed: ' + e.message) } }
  const exportAll = async () => { try { const z = new JSZip(), cards = [...document.querySelectorAll('.grid>.card')]; z.file('data_updated.csv', toCSV(data)); if (rows.length !== data.length) z.file('data_filtered.csv', toCSV(rows)); const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(data), 'Data'); z.file('data_updated.xlsx', XLSX.write(wb, { type: 'array', bookType: 'xlsx' })); const kp = [['Title', 'Value']]; let n = 0; for (const c of cards) { n++; const t = (c.querySelector('.mut') || {}).textContent || 'visual', k = c.querySelector('.kpi'), sv = c.querySelector('svg'); if (k) kp.push([t, k.firstChild.textContent]); else if (sv) { const cv = await svgCv(sv).catch(() => null); if (cv) z.file('charts/' + n + '-' + t.replace(/[^a-z0-9]+/gi, '_').slice(0, 40) + '.png', await new Promise(r => cv.toBlob(r))) } } z.file('kpis.csv', kp.map(r => r.map(v => '"' + String(v).replace(/"/g, '""') + '"').join(',')).join('\n')); z.file('report.html', report(name)); const pb = await pdf(true); if (pb) z.file('dashboard.pdf', pb); save2('insightforge-export.zip', await z.generateAsync({ type: 'blob' })) } catch (e) { alert('Export failed: ' + e.message) } }
  const shareReport = () => save2(name.replace(/\W+/g, '_') + '-report.html', new Blob([report(name)], { type: 'text/html' }))
  const doJoin = () => { const { rows: o, a, b, t } = jn; const idx = {}; o.forEach(r => { (idx[r[b]] = idx[r[b]] || []).push(r) }); const ex = Object.keys(o[0]).filter(k => k !== b); const out = []; data.forEach(r => { const m = idx[r[a]]; if (m) m.forEach(x => { const n = { ...r }; ex.forEach(k => { n[k in r ? k + '_2' : k] = x[k] }); out.push(n) }); else if (t === 'left') { const n = { ...r }; ex.forEach(k => { n[k in r ? k + '_2' : k] = '' }); out.push(n) } }); setData(out); setJn({ rows: null, a: '', b: '', t: 'left' }) }
  const [calc, setCalc] = useState({ n: '', f: '', e: '' })
  const addCalc = () => { try { let e = calc.f;[...cols].sort((a, b) => b.name.length - a.name.length).forEach(c => { e = e.split(c.name).join('r[' + JSON.stringify(c.name) + ']') }); const fn = new Function('r', 'return (' + e + ')'); const nm = calc.n.trim() || 'Calc'; setData(data.map(r => { const v = +fn(r); return { ...r, [nm]: isFinite(v) ? +v.toFixed(2) : '' } })); setCalc({ n: '', f: '', e: '' }) } catch (x) { setCalc({ ...calc, e: 'Invalid formula: ' + x.message }) } }
  const save = () => { try { localStorage.setItem('ifdash', JSON.stringify({ ws, flt, data: data.length < 20000 ? data : null, name })); setMsg('Layout saved') } catch (e) { setMsg('Saving blocked in this browser') } }
  const restore = () => { try { const j = JSON.parse(localStorage.getItem('ifdash')); if (j) { if (j.data) { setData(j.data); setName(j.name || 'Saved project') } setWs(j.ws); setFlt(j.flt); setMsg('Layout loaded') } else setMsg('Nothing saved yet') } catch (e) { setMsg('Load failed') } }
  const mv = (id, d) => setWs(s => { const i = s.findIndex(z => z.id === id), j = i + d; if (j < 0 || j >= s.length) return s; const c = [...s];[c[i], c[j]] = [c[j], c[i]]; return c })
  const wide = id => setWs(s => s.map(z => z.id === id ? { ...z, wide: !z.wide } : z))

  /* ---- Gemini AI ---- */
  const askAI = async t => {
    setInp(''); setChat(c => [...c, { r: 'u', t }, { r: 'a', t: 'Thinking…', tmp: 1 }])
    try {
      if (!GEMINI_KEY.trim()) throw { message: 'VITE_GEMINI_API_KEY env var not set. Add it in Vercel → Settings → Environment Variables.' }
      const prompt = 'You are a senior BI analyst. The user has a dataset with these columns: ' + JSON.stringify(cols.map(c => ({ name: c.name, type: c.num ? 'number' : 'text', uniqueValues: c.uniq, min: c.min, max: c.max, mean: c.mean }))) + '. There are ' + data.length + ' rows. Precomputed grouped sums: ' + summ(cols, data) + '. First 5 rows: ' + JSON.stringify(data.slice(0, 5)) + '.\nUser request: ' + t + '\nReply ONLY with valid JSON (no markdown fences): {"reply":"direct answer with exact numbers, then one short tip","widgets":[{"type":"bar|line|area|pie|scatter|kpi|table","x":"column","y":"column","agg":"sum|avg|count|max|min","top":5,"title":"optional label"}]}. Use only the exact column names listed. widgets may be an empty array.'
      const res = await fetch('https://generativelanguage.googleapis.com/v1beta/models/gemini-3.8-flash:generateContent?key=' + GEMINI_KEY.trim(), {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          contents: [{ parts: [{ text: prompt }] }],
          generationConfig: {
            responseMimeType: 'application/json',
            responseSchema: {
              type: 'OBJECT',
              properties: {
                reply: { type: 'STRING' },
                widgets: {
                  type: 'ARRAY',
                  items: {
                    type: 'OBJECT',
                    properties: {
                      type: { type: 'STRING' },
                      x: { type: 'STRING' },
                      y: { type: 'STRING' },
                      agg: { type: 'STRING' },
                      top: { type: 'INTEGER' },
                      title: { type: 'STRING' }
                    },
                    required: ['type']
                  }
                }
              },
              required: ['reply', 'widgets']
            },
            temperature: 0.2,
            maxOutputTokens: 4096
          }
        })
      })
      if (!res.ok) { const err = await res.json().catch(() => ({})); throw { message: (err.error?.message || 'HTTP ' + res.status) } }
      const rawText = await res.text()
      let raw = {}
      try {
        raw = JSON.parse(rawText)
      } catch {
        throw { message: 'Incomplete response received from Gemini API. Please try again.' }
      }
      let txt = (raw.candidates?.[0]?.content?.parts?.[0]?.text || '').trim()
      let j = { reply: txt, widgets: [] }
      if (txt) {
        try {
          j = JSON.parse(txt)
        } catch {
          const firstBrace = txt.indexOf('{')
          const lastBrace = txt.lastIndexOf('}')
          if (firstBrace !== -1 && lastBrace !== -1) {
            try {
              j = JSON.parse(txt.slice(firstBrace, lastBrace + 1))
            } catch {
              const m = txt.match(/"reply"\s*:\s*"([^"]+)"/)
              j = { reply: m ? m[1] : txt.replace(/[{}"\\]/g, ' ').trim(), widgets: [] }
            }
          }
        }
      }
      const okc = n => cols.some(c => c.name === n)
      const W = (j.widgets || []).filter(w => w.type && (w.type === 'table' || (w.type === 'kpi' ? okc(w.y) : okc(w.y) && okc(w.x)))).map(w => ({ ...w, agg: w.agg || 'sum' }))
      W.forEach(add)
      setChat(c => [...c.filter(m => !m.tmp), { r: 'a', t: (j.reply || 'Done.') + (W.length ? ' (added ' + W.length + ' visual' + (W.length > 1 ? 's' : '') + ' to dashboard)' : ''), w: W[0] }])
    } catch (e) {
      setChat(c => [...c.filter(m => !m.tmp), { r: 'a', t: 'AI error: ' + (e.message || 'Unknown') + '. Uncheck ✦ Gemini AI to use the built-in parser.' }])
    }
  }

  const send = t => { if (!t.trim()) return; if (ai) return askAI(t); const w = interpret(t, cols); const ok = w.type === 'table' || w.type === 'kpi' ? w.y || w.type === 'table' : w.x && w.y; setChat(c => [...c, { r: 'u', t }, ok ? { r: 'a', t: `Done — added a ${w.type} chart (${w.agg}${w.y ? ' of ' + w.y : ''}${w.x && w.type !== 'kpi' ? ' by ' + w.x : ''}) to your dashboard.`, w } : { r: 'a', t: "I couldn't find matching columns. Mention a column name like: " + cols.map(c => c.name).join(', ') }]); if (ok) add(w); setInp('') }
  const sugg = [nums[0] && cats[0] && `${nums[0].name} by ${cats[0].name} bar chart`, nums[0] && cats[0] && `share of ${nums[0].name} by ${cats[0].name} pie`, nums[0] && `total ${nums[0].name} kpi`, nums[1] && `average ${nums[1].name} kpi`, nums[0] && cats[1] && `top 5 ${cats[1].name} by ${nums[0].name}`, nums.length > 1 && `scatter ${nums[0].name} vs ${nums[1].name}`].filter(Boolean)
  const plan = [['Phase 1 · Ingest & Clean', ['Import CSV / JSON / raw text', 'Check missing values & types in Profile', 'Remove duplicates, fix date formats']], ['Phase 2 · Define KPIs', ['Pick 3–5 headline KPIs (total, average, growth %)', 'Agree targets & time grain (monthly/weekly)', 'Add slicers for Region, Category, Date']], ['Phase 3 · Visualize', ['Trend: line/area for time series', 'Comparison: bar for categories (top N)', 'Composition: pie for shares (≤6 slices)', 'Relationship: scatter for correlations']], ['Phase 4 · Explore with Chat', ['Ask for ad-hoc views & drill-downs', 'Cross-filter the dashboard', 'Save the combos that answer real questions']], ['Phase 5 · Share & Iterate', ['Review with stakeholders', 'Refine visuals, remove clutter', 'Schedule data refresh & monitor']]]
  const ask = t => { setTab('chat'); setTimeout(() => send(t), 60) }
  const tabs = [['home', '◆', 'Home'], ['dash', '▦', 'Dashboard'], ['data', '⌗', 'Data & Profile'], ['pivot', '⊞', 'Pivot'], ['chat', '✦', 'AI Chat'], ['plan', '✓', 'Action Plan']]

  return <div className="app">
    <aside className="side">
      <div className="logo">◆ InsightForge</div>
      {tabs.map(([k, i, l]) => <button key={k} className={'nav' + (tab === k ? ' on' : '')} onClick={() => setTab(k)}>{tab === k && <M.span layoutId="pill" style={{ position: 'absolute', left: 0, top: 8, bottom: 8, width: 3, borderRadius: 3, background: 'var(--ac)' }} />}<span>{i}</span>{l}</button>)}
      <div style={{ flex: 1 }} />
      <div className="mut" style={{ fontSize: 11, padding: 8 }}>Dataset: {name}<br />{data.length} rows · {cols.length} cols</div>
    </aside>
    <main className="main"><AP mode="wait"><M.div key={tab} initial={{ opacity: 0, y: 12 }} animate={{ opacity: 1, y: 0 }} exit={{ opacity: 0 }} transition={{ duration: .25 }}>
      {tab === 'home' && <Home data={data} cols={cols} go={setTab} ask={ask} />}
      {tab === 'pivot' && <><h2>Pivot table</h2><div className="sub">Cross-tabulate any fields with totals and a heat map.</div><div className="card"><Pivot rows={rows} cols={cols} /></div></>}
      {tab === 'dash' && <>
        <h2>Dashboard</h2><div className="sub">Build, filter and rearrange your report.</div>
        <div className="card row" style={{ marginBottom: 14 }}>
          <b>Add visual</b>
          <select value={nw.type} onChange={e => setNw({ ...nw, type: e.target.value })}>{['bar', 'line', 'area', 'pie', 'scatter', 'hist', 'stack', 'heat', 'box', 'funnel', 'kpi', 'table'].map(t => <option key={t}>{t}</option>)}</select>
          {nw.type !== 'kpi' && nw.type !== 'table' && <select value={nw.x} onChange={e => setNw({ ...nw, x: e.target.value })}><option value="">X / group…</option>{cols.map(c => <option key={c.name}>{c.name}</option>)}</select>}
          {nw.type !== 'table' && <select value={nw.y} onChange={e => setNw({ ...nw, y: e.target.value })}><option value="">Value…</option>{cols.map(c => <option key={c.name}>{c.name}</option>)}</select>}
          {nw.type !== 'scatter' && nw.type !== 'table' && <select value={nw.agg} onChange={e => setNw({ ...nw, agg: e.target.value })}>{Object.keys(AG).map(a => <option key={a}>{a}</option>)}</select>}
          {(nw.type === 'stack' || nw.type === 'heat') && <select value={nw.z} onChange={e => setNw({ ...nw, z: e.target.value })}><option value="">Series…</option>{cols.map(c => <option key={c.name}>{c.name}</option>)}</select>}
          {['line', 'area', 'bar'].includes(nw.type) && <select value={nw.tf} onChange={e => setNw({ ...nw, tf: e.target.value })}>{[['none', 'As is'], ['run', 'Running total'], ['ma', '3-pt moving avg'], ['mom', 'Period-over-period %'], ['yoy', 'Year-over-year %']].map(([k, l]) => <option key={k} value={k}>{l}</option>)}</select>}
          {nw.type === 'kpi' && <input type="number" placeholder="Target" style={{ width: 90 }} value={nw.tgt} onChange={e => setNw({ ...nw, tgt: +e.target.value || '' })} />}
          <button className="b pri" onClick={() => add(nw)}>+ Add</button>
          <button className="b" onClick={save}>Save layout</button><button className="b" onClick={restore}>Load layout</button>
          <button className="b" onClick={() => save2('data.csv', toCSV(rows))}>Export CSV</button>
          <button className="b" onClick={() => pdf()}>Export PDF</button>
          <button className="b pri" onClick={exportAll} title="CSV + Excel + charts + PDF + HTML in one ZIP">Download all (ZIP)</button>
          <button className="b" onClick={shareReport}>Share report</button>
          {dc && <><span className="mut">Range</span><select value={rng.a} onChange={e => setRng({ ...rng, a: e.target.value })}><option value="">from…</option>{dvals.map(v => <option key={v}>{v}</option>)}</select><select value={rng.b} onChange={e => setRng({ ...rng, b: e.target.value })}><option value="">to…</option>{dvals.map(v => <option key={v}>{v}</option>)}</select></>}
          {(flt.v || rng.a || rng.b) && <button className="b" onClick={() => { setFlt({ c: '', v: '' }); setRng({ a: '', b: '' }) }}>Clear filters</button>}
          {flt.v && <span className="mut">Filtered: {flt.c} = {flt.v}</span>}
          {msg && <span className="mut">{msg}</span>}
          <span style={{ flex: 1 }} /><b>Slicer</b>
          <select value={flt.c} onChange={e => setFlt({ c: e.target.value, v: '' })}><option value="">Filter by…</option>{cats.map(c => <option key={c.name}>{c.name}</option>)}</select>
          {flt.c && <select value={flt.v} onChange={e => setFlt({ ...flt, v: e.target.value })}><option value="">All</option>{[...new Set(data.map(r => r[flt.c]))].map(v => <option key={v}>{v}</option>)}</select>}
        </div>
        <div className="grid"><AP>{ws.map(w => <M.div key={w.id} layout="position" initial={{ opacity: 0, scale: .92 }} animate={{ opacity: 1, scale: 1 }} exit={{ opacity: 0, scale: .9 }} whileHover={{ y: -3 }} id={'w' + w.id} onDragOver={e => e.preventDefault()} onDrop={() => drop(w.id)} style={{ position: 'relative', gridColumn: 'span ' + (w.span || (w.type === 'table' || w.wide ? 2 : 1)) }} className="card">
          <button className="x" title="Remove" onClick={() => setWs(s => s.filter(z => z.id !== w.id))}>×</button>
          <button className="x" title="Wide / narrow" onClick={() => wide(w.id)}>⇔</button>
          <button className="x" title="Move later" onClick={() => mv(w.id, 1)}>›</button>
          <button className="x" title="Move earlier" onClick={() => mv(w.id, -1)}>‹</button>
          <button className="x" title="Export PNG" onClick={() => png(w.id)}>⤓</button>
          <span className="x" draggable onDragStart={() => { dg.current = w.id }} title="Drag to reorder" style={{ cursor: 'grab' }}>⠿</span>
          <div onPointerDown={e => rs(e, w)} title="Drag to resize" style={{ position: 'absolute', right: 5, bottom: 5, width: 16, height: 16, cursor: 'nwse-resize', borderRight: '2px solid var(--mut)', borderBottom: '2px solid var(--mut)', borderRadius: 2, touchAction: 'none' }} />
          <div className="mut" style={{ fontSize: 12, marginBottom: 6 }}>{chartTitle(w)}{w.top ? ` (top ${w.top})` : ''}</div>
          <div className={w.h ? 'cw' : ''} style={w.h ? { height: w.h } : {}}><Chart w={w} rows={rows} onPick={w.type === 'hist' ? null : (c, v) => setFlt({ c, v: String(v) })} /></div>
        </M.div>)}</AP></div>
        {!ws.length && <div className="card mut">Empty canvas — add a visual above or ask the AI Chat.</div>}
      </>}
      {tab === 'data' && <>
        <h2>Data &amp; Profile</h2><div className="sub">Structured (CSV, JSON) or unstructured (any text → keyword table).</div>
        <div className="card" style={{ marginBottom: 14 }}><div className="row">
          <input type="file" accept=".csv,.json,.txt,.md,.xlsx,.xls" onChange={e => { const f = e.target.files[0]; if (!f) return; const r = new FileReader(); if (/\.xlsx?$/i.test(f.name)) { r.onload = () => { try { const wb = XLSX.read(r.result, { type: 'array' }); load(f.name, XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { defval: '' })); setExtra(wb.SheetNames.slice(1).map(n => ({ name: n, rows: XLSX.utils.sheet_to_json(wb.Sheets[n], { defval: '' }) }))) } catch (x) { alert('Could not read workbook') } }; r.readAsArrayBuffer(f) } else { r.onload = () => load(f.name, r.result); r.readAsText(f) } }} />
          <button className="b" onClick={() => { setData(sample); setName('Sample: Retail Sales'); setFlt({ c: '', v: '' }) }}>Load sample</button>
        </div>
          <textarea value={paste} onChange={e => setPaste(e.target.value)} placeholder="…or paste CSV / JSON / free text" rows="4" style={{ width: '100%', marginTop: 10 }} />
          <button className="b pri" style={{ marginTop: 8 }} onClick={() => paste.trim() && load('pasted data', paste)}>Analyze pasted data</button>
          <button className="b" style={{ marginTop: 8, marginLeft: 8 }} onClick={() => paste.trim() && load('sentiment', sentiment(paste))}>Sentiment analysis</button></div>
        <div className="card" style={{ marginBottom: 14 }}><b>Clean data</b> <span className="mut">({data.length} rows{hist.length ? ' · undo available' : ''})</span>
          <div className="row" style={{ marginTop: 8 }}><button className="b" onClick={() => mod(dedupe)}>Remove duplicates</button><button className="b" onClick={() => mod(d => d.filter(r => Object.values(r).every(v => v !== '' && v != null)))}>Drop blank rows</button><button className="b" onClick={() => mod(d => d.map(r => Object.fromEntries(Object.entries(r).map(([k, v]) => [k, typeof v === 'string' ? v.trim() : v]))))}>Trim text</button><button className="b" disabled={!hist.length} onClick={undo}>↶ Undo</button></div>
          <div className="row" style={{ marginTop: 8 }}><select value={cl.col} onChange={e => setCl({ ...cl, col: e.target.value })}><option value="">Column…</option>{cols.map(c => <option key={c.name}>{c.name}</option>)}</select><input placeholder="Value / new name" value={cl.v} onChange={e => setCl({ ...cl, v: e.target.value })} />
            <button className="b" onClick={() => cl.col && mod(d => { const c = cols.find(x => x.name === cl.col), fv = cl.v !== '' ? (c.num && !isNaN(+cl.v) ? +cl.v : cl.v) : (c.num ? +c.mean.toFixed(2) : 'Unknown'); return d.map(r => r[cl.col] === '' || r[cl.col] == null ? { ...r, [cl.col]: fv } : r) })}>Fill blanks</button>
            <button className="b" onClick={() => cl.col && mod(d => d.map(r => ({ ...r, [cl.col]: numFn(r[cl.col]) })))}>To number</button>
            <button className="b" onClick={() => cl.col && cl.v && mod(d => d.map(r => Object.fromEntries(Object.entries(r).map(([k, v]) => [k === cl.col ? cl.v : k, v]))))}>Rename</button>
            <button className="b" onClick={() => cl.col && mod(d => d.map(r => { const o = { ...r }; delete o[cl.col]; return o }))}>Delete column</button>
            <select value={cl.op} onChange={e => setCl({ ...cl, op: e.target.value })}>{['=', '!=', '>', '<', '>=', '<=', 'contains'].map(o => <option key={o}>{o}</option>)}</select>
            <button className="b" onClick={() => cl.col && mod(d => fltRows(d, [{ column: cl.col, op: cl.op, value: cl.v }]))}>Keep rows where</button>
          </div></div>
        {nums.length > 1 && <div className="card" style={{ marginBottom: 14 }}><b>Correlation matrix</b><div className="scroll"><table><thead><tr><th></th>{nums.slice(0, 8).map(c => <th key={c.name}>{c.name}</th>)}</tr></thead><tbody>{nums.slice(0, 8).map(a => <tr key={a.name}><td><b>{a.name}</b></td>{nums.slice(0, 8).map(b => { const r = pear(data, a, b); return <td key={b.name} style={{ background: corrBg(r) }}>{r.toFixed(2)}</td> })}</tr>)}</tbody></table></div></div>}
        <div className="card" style={{ marginBottom: 14 }}><b>Join datasets</b> <span className="mut">(like a Power Query merge)</span><div className="row" style={{ marginTop: 8 }}><input type="file" accept=".csv,.json,.txt,.xlsx,.xls" onChange={e => { const f = e.target.files[0]; f && readAny(f, rows => setJn({ ...jn, rows })) }} /><button className="b" onClick={() => setJn({ ...jn, rows: [{ Region: 'North', Manager: 'Asha', Target: 400000 }, { Region: 'South', Manager: 'Ravi', Target: 350000 }, { Region: 'East', Manager: 'Meera', Target: 300000 }, { Region: 'West', Manager: 'Kabir', Target: 380000 }] })}>Use sample lookup</button>{extra.map(s => <button key={s.name} className="b" onClick={() => setJn({ ...jn, rows: s.rows })}>Sheet: {s.name}</button>)}</div>{jn.rows && <div className="row" style={{ marginTop: 8 }}><span>Current key</span><select value={jn.a} onChange={e => setJn({ ...jn, a: e.target.value })}><option value="">…</option>{cols.map(c => <option key={c.name}>{c.name}</option>)}</select><span>= new key</span><select value={jn.b} onChange={e => setJn({ ...jn, b: e.target.value })}><option value="">…</option>{Object.keys(jn.rows[0] || {}).map(k => <option key={k}>{k}</option>)}</select><select value={jn.t} onChange={e => setJn({ ...jn, t: e.target.value })}><option>left</option><option>inner</option></select><button className="b pri" disabled={!jn.a || !jn.b} onClick={doJoin}>Join</button><span className="mut">{jn.rows.length} rows loaded</span><button className="b" onClick={() => { mod(d => [...d, ...jn.rows]); setJn({ ...jn, rows: null }) }}>Stack (append rows)</button></div>}</div>
        <div className="card" style={{ marginBottom: 14 }}><b>Calculated column</b> <span className="mut">(e.g. Profit / Sales * 100)</span><div className="row" style={{ marginTop: 8 }}><input placeholder="New name" value={calc.n} onChange={e => setCalc({ ...calc, n: e.target.value })} /><input style={{ flex: 1, minWidth: 180 }} placeholder="Formula using column names" value={calc.f} onChange={e => setCalc({ ...calc, f: e.target.value })} /><button className="b pri" onClick={addCalc}>Create</button></div>{calc.e && <div style={{ color: '#f43f5e', marginTop: 6 }}>{calc.e}</div>}</div>
        <div className="card" style={{ marginBottom: 14 }}><b>Auto insights</b>{insights.map(t => <div key={t} style={{ marginTop: 6 }}>• {t}</div>)}{!insights.length && <div className="mut">No notable patterns yet.</div>}</div>
        <div className="card" style={{ marginBottom: 14 }}><b>Column profile</b><div className="scroll"><table><thead><tr><th>Column</th><th>Type</th><th>Unique</th><th>Missing</th><th>Min</th><th>Mean</th><th>Median</th><th>Std</th><th>Max</th><th>Outliers</th></tr></thead>
          <tbody>{cols.map(c => <tr key={c.name}><td>{c.name}</td><td>{c.num ? 'Number' : 'Text'}</td><td>{c.uniq}</td><td style={{ color: c.missing ? '#f43f5e' : 'inherit' }}>{c.missing}</td><td>{c.num ? fmt(c.min) : '–'}</td><td>{c.num ? fmt(c.mean) : '–'}</td><td>{c.num ? fmt(c.med) : '–'}</td><td>{c.num ? fmt(c.std) : '–'}</td><td>{c.num ? fmt(c.max) : '–'}</td><td style={{ color: c.out ? '#f43f5e' : 'inherit' }}>{c.num ? c.out : '–'}</td></tr>)}</tbody></table></div></div>
        <div className="card"><b>Data grid</b> <span className="mut">search, click headers to sort</span><Grid rows={data} /></div>
      </>}
      {tab === 'chat' && <>
        <h2>AI Chat</h2><div className="sub">Describe what you need in plain language — visuals are generated onto your dashboard.</div>
        <div className="card" style={{ display: 'flex', flexDirection: 'column', gap: 10, minHeight: 320 }}>
          {chat.map((m, i) => <M.div key={i} initial={{ opacity: 0, y: 8 }} animate={{ opacity: 1, y: 0 }} className={'msg ' + m.r}>
            <div style={{ lineHeight: 1.5 }}>{m.t}</div>
            {m.w && <div className="chat-card-widget">
              <div className="chat-card-header">
                <div className="chat-card-meta">
                  <span className="chat-badge">{m.w.type}</span>
                  <span className="chat-card-title">{m.w.title || chartTitle(m.w)}</span>
                </div>
                <div style={{ display: 'flex', gap: 6, flexWrap: 'wrap' }}>
                  <button
                    className="chat-pin-btn pri-btn"
                    onClick={() => {
                      add(m.w)
                      setTab('dash')
                    }}
                    title="Add chart and open dashboard"
                  >
                    + Add to Dashboard
                  </button>
                  <button
                    className="chat-pin-btn"
                    onClick={() => setTab('dash')}
                    title="Open dashboard tab"
                  >
                    View Dashboard →
                  </button>
                </div>
              </div>
              <div className="chat-card-canvas">
                <Chart w={m.w} rows={rows} />
              </div>
            </div>}
          </M.div>)}
          <div className="row" style={{ marginTop: 'auto' }}>
            <label className="chip"><input type="checkbox" checked={ai} onChange={e => setAi(e.target.checked)} /> ✦ Gemini AI</label>
            <input style={{ flex: 1, minWidth: 180 }} value={inp} onChange={e => setInp(e.target.value)} onKeyDown={e => e.key === 'Enter' && send(inp)} placeholder="e.g. average profit by region pie chart" />
            <button className="b pri" onClick={() => send(inp)}>Send</button>
          </div>
        </div>
        <h3>Suggested for this dataset</h3><div className="row">{sugg.map(s => <button key={s} className="chip" onClick={() => send(s)}>{s}</button>)}</div>
        <div className="card mut" style={{ marginTop: 14, fontSize: 13 }}><b style={{ color: 'var(--tx)' }}>Recommended methods:</b> trend analysis (line), Pareto / top-N (bar), composition (pie ≤6 slices), correlation (scatter), KPI cards (total, average, max), outlier check via Profile min/max.</div>
      </>}
      {tab === 'plan' && <>
        <h2>Action Plan</h2><div className="sub">Your path from raw data to a decision-ready dashboard.</div>
        <div className="grid">{plan.map(([h, it], pi) => <M.div key={h} className="card" initial={{ opacity: 0, y: 14 }} animate={{ opacity: 1, y: 0 }} transition={{ delay: pi * .07 }}><b>{h}</b>{it.map(t => <label key={t} style={{ display: 'flex', gap: 8, marginTop: 10, cursor: 'pointer', opacity: done[t] ? .5 : 1 }}><input type="checkbox" checked={!!done[t]} onChange={() => setDone({ ...done, [t]: !done[t] })} />{t}</label>)}</M.div>)}</div>
        <div className="card" style={{ marginTop: 14 }}><b>Progress</b><div style={{ height: 8, background: 'var(--panel2)', borderRadius: 8, marginTop: 8, overflow: 'hidden' }}><M.div animate={{ width: Object.values(done).filter(Boolean).length / plan.flat(2).filter(x => !x.startsWith('Phase')).length * 100 + '%' }} style={{ height: '100%', background: 'linear-gradient(90deg,var(--ac),var(--ac2))' }} /></div></div>
      </>}
    </M.div></AP></main>
  </div>
}
