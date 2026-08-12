#!/usr/bin/env node
/* 대량 삭제 경보 — 한 커밋이 파일을 통째로 갈아엎는 사고를 잡는다.
 *
 * 2026-08-11 사고: '수주관리 파이프라인 3단계 간소화' 커밋 하나가 index.html 을
 * 10,014줄 바꿨고(4,992 추가 / 5,022 삭제), 그 안에서 조회기간 자동설정·
 * 반품충당 합계·담당자 대수·딜러명 말줄임이 한꺼번에 사라졌다.
 * 작업 내용에 비해 삭제가 지나치게 많으면 옛 파일 위에 덮어쓴 것이다.
 *
 * 통과시키려면 커밋 메시지에 [대량변경] 을 넣는다(의도한 리팩터링임을 밝히는 표시).
 */
const { execSync } = require('child_process');

const LIMIT = 1500;           // 파일 하나에서 이만큼 넘게 지워지면 멈춘다
const RATIO = 0.35;           // 또는 파일의 35% 넘게 지워지면

const sh = c => { try { return execSync(c, { encoding: 'utf8' }); } catch (e) { return ''; } };

// 비교 기준
//  · 내 PC(push 직전): origin/main = 아직 안 올라간 원격 상태 → 이번에 올릴 변화량이 나온다
//  · GitHub Actions : 이미 올라간 뒤라 origin/main 이 HEAD 와 같다 → 그대로 두면 변화량 0으로
//    무조건 통과해 버린다. 이때는 직전 커밋과 비교한다.
const head = sh('git rev-parse HEAD').trim();
let base = sh('git rev-parse --verify --quiet origin/main').trim();
if (!base || base === head) base = sh('git rev-parse --verify --quiet HEAD~1').trim();
if (!base) process.exit(0);   // 첫 커밋이면 검사할 게 없다

const stat = sh(`git diff --numstat ${base} HEAD`).trim();
if (!stat) process.exit(0);

// 의도한 대량 변경이면 통과 — 본문이 아니라 '제목줄'에만 표시를 인정한다.
// (본문까지 보면 이 검사를 설명하는 커밋 메시지 자체가 스스로를 통과시켜 버린다)
const msgs = sh(`git log --format=%s ${base}..HEAD`);
if (/\[대량변경\]/.test(msgs)) {
  console.log('대량 삭제 검사 — [대량변경] 표시가 있어 건너뜁니다.');
  process.exit(0);
}

const hits = [];
stat.split('\n').forEach(line => {
  const [add, del, file] = line.split('\t');
  if (!file || add === '-' ) return;                    // 바이너리는 제외
  const d = parseInt(del, 10) || 0;
  if (!d) return;
  const total = (sh(`git show ${base}:${file}`).match(/\n/g) || []).length;
  const pct = total ? d / total : 0;                      // 새 파일(원본 없음)은 비율 판정을 하지 않는다
  if (d > LIMIT || pct > RATIO) hits.push({ file, add: parseInt(add, 10) || 0, del: d, total, pct });
});

if (!hits.length) { console.log('대량 삭제 검사 통과'); process.exit(0); }

console.error('\n대량 삭제가 감지됐습니다 — 옛 파일 위에 덮어쓴 것은 아닌지 확인하세요.\n');
hits.forEach(h => {
  console.error(`  ${h.file}`);
  console.error(`    +${h.add.toLocaleString()} / -${h.del.toLocaleString()} 줄` +
    (h.total ? ` (원본 ${h.total.toLocaleString()}줄의 ${(h.pct * 100).toFixed(0)}% 삭제)` : ''));
});
console.error('\n  · 최신 상태에서 작업한 게 맞나요?  git pull --rebase 후 다시 확인');
console.error('  · 지워진 줄에 남이 고쳐둔 로직이 섞이지 않았나요?  git diff ' + base + ' HEAD -- <파일> | grep "^-"');
console.error('  · 의도한 정리라면 커밋 메시지에 [대량변경] 을 넣으세요.\n');
process.exit(1);
