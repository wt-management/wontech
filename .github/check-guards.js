#!/usr/bin/env node
/**
 * 회귀 검사 — .github/guards.json 에 적힌 코드가 아직 남아 있는지 확인한다.
 *
 * 배경: 한 번 고친 로직이 다른 사람 작업에 밀려 조용히 사라지는 일이 반복됐다.
 *       (해외 조회기간 5월 고정 버그는 고친 뒤 3개 커밋 만에 되돌아갔다)
 *       화면은 멀쩡히 뜨고 숫자만 틀리기 때문에 눈으로는 잡히지 않는다.
 *
 * 사용: node .github/check-guards.js
 *       실패하면 종료코드 1 → GitHub Actions 가 빨간불로 알려준다.
 */
const fs = require('fs');
const path = require('path');

const ROOT = path.join(__dirname, '..');
const conf = JSON.parse(fs.readFileSync(path.join(__dirname, 'guards.json'), 'utf8'));

let checked = 0;
const failed = [];

for (const [file, rules] of Object.entries(conf.files || {})) {
  const p = path.join(ROOT, file);
  if (!fs.existsSync(p)) {
    failed.push({ file, id: '(파일)', why: '파일을 찾을 수 없습니다' });
    continue;
  }
  const src = fs.readFileSync(p, 'utf8');
  for (const r of rules) {
    checked++;
    if (src.indexOf(r.must) < 0) failed.push({ file, id: r.id, why: r.why, must: r.must });
  }
}

console.log(`회귀 검사 ${checked}건 실행`);

if (!failed.length) {
  console.log('모두 통과 — 고쳐둔 로직이 그대로 남아 있습니다.');
  process.exit(0);
}

console.log('');
console.log(`실패 ${failed.length}건 — 아래 로직이 사라졌습니다.`);
for (const f of failed) {
  console.log('');
  console.log(`  [${f.file}] ${f.id}`);
  console.log(`   사유: ${f.why}`);
  if (f.must) console.log(`   찾던 코드: ${f.must.slice(0, 100)}`);
}
console.log('');
console.log('의도한 변경이라면 .github/guards.json 의 해당 항목을 새 코드에 맞게 고치세요.');
process.exit(1);
