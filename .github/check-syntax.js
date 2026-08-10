#!/usr/bin/env node
/**
 * 자바스크립트 문법 검사 — 화면이 통째로 안 뜨는 사고를 배포 전에 잡는다.
 *
 * 주의: 이 파일들은 HTML 안에 <script>가 여러 개 들어 있고, 문자열 안에
 *       "</script>" 나 "<" 가 섞여 있어 블록을 기계적으로 자르면 일부는
 *       원래부터 파싱에 실패한다(오탐). 그래서 절대 개수가 아니라
 *       '기준선보다 늘었는가'로 판단한다.
 */
const fs = require('fs');
const path = require('path');
const vm = require('vm');

const ROOT = path.join(__dirname, '..');
const BASE = path.join(__dirname, 'syntax-baseline.json');
const baseline = fs.existsSync(BASE) ? JSON.parse(fs.readFileSync(BASE, 'utf8')) : {};

function scan(file) {
  const src = fs.readFileSync(path.join(ROOT, file), 'utf8');
  const lines = src.split('\n');
  const blocks = [];
  let open = null, buf = [];
  lines.forEach((l, i) => {
    if (/<script(?![^>]*src=)[^>]*>/.test(l) && !open) { open = i + 1; buf = []; return; }
    if (/<\/script>/.test(l) && open) { blocks.push([open, i + 1, buf.join('\n')]); open = null; return; }
    if (open) buf.push(l);
  });
  const bad = [];
  blocks.forEach(([a, b, code]) => {
    try { new vm.Script(code); } catch (e) { bad.push({ at: `${a}~${b}`, msg: String(e.message).slice(0, 80) }); }
  });
  return { total: blocks.length, bad };
}

const files = Object.keys(baseline).length ? Object.keys(baseline) : ['index.html', 'total.html'];
let worse = false;

for (const f of files) {
  if (!fs.existsSync(path.join(ROOT, f))) continue;
  const r = scan(f);
  const expected = baseline[f] != null ? baseline[f] : 0;
  const mark = r.bad.length > expected ? '실패' : '통과';
  console.log(`${mark}  ${f} — script ${r.total}개 / 파싱 실패 ${r.bad.length}개 (기준선 ${expected})`);
  if (r.bad.length > expected) {
    worse = true;
    r.bad.forEach(x => console.log(`        ${x.at}  ${x.msg}`));
  }
}

if (worse) {
  console.log('');
  console.log('문법 오류가 기준선보다 늘었습니다. 방금 수정한 부분을 확인하세요.');
  console.log('구조를 바꿔 기준선이 달라진 것이라면 .github/syntax-baseline.json 을 갱신하세요.');
  process.exit(1);
}
console.log('문법 검사 통과');
