# wontech — 원텍 한국영업 · 전사 종합

정적 웹앱(빌드 없음, HTML 하나에 CSS·JS 전부 내장).

| 파일 | 사이트 |
|---|---|
| `index.html` | 한국영업 |
| `total.html` | 전사 종합 |
| `upload.html` | 회계 매출자료 업로드(파서 진입점) |
| `wtu_engine.js` · `portfolio_engine.js` · `ilbo_engine.js` | 업로드 파서 |

- 라이브: https://wt-management.github.io/wontech/
- 데이터는 Supabase(`cons_cache` 등)에 있다. 저장소를 고쳐도 데이터는 안 지워진다.

## 배포 전 검사 — 반드시 실행

이 저장소를 고쳤으면 **push 하기 전에** 아래를 돌린다. 셋 다 통과해야 한다.

```bash
node .github/check-syntax.js && node .github/check-guards.js && node .github/check-bulk.js
```

| 검사 | 무엇을 잡나 |
|---|---|
| `check-syntax.js` | 화면이 통째로 안 뜨는 문법 오류 (기준선보다 늘면 실패) |
| `check-guards.js` | 한 번 고쳐둔 로직이 사라졌는지 (`.github/guards.json` 목록) |
| `check-bulk.js` | 한 파일에서 1,500줄 이상 또는 35% 넘게 삭제 — 옛 파일 위에 덮어쓴 사고 |

**버그를 고쳤으면 `.github/guards.json` 에 한 줄 추가한다.** 이게 이 저장소의 핵심 규칙이다.
같은 버그가 다른 작업에 밀려 조용히 되돌아가는 일이 실제로 여러 번 있었고(해외 조회기간은
고친 뒤 3개 커밋 만에 사라졌다), 화면은 멀쩡히 뜨고 숫자만 틀리기 때문에 눈으로는 못 잡는다.

```json
{ "id": "짧은이름", "why": "빠지면 무슨 일이 생기는지", "must": "파일에 남아 있어야 할 코드 한 조각" }
```

`check-bulk.js` 가 막아설 때는 대개 **최신 상태에서 작업하지 않은 것**이다. `git pull --rebase`
후 다시 확인하고, 정말 의도한 대규모 정리라면 커밋 **제목**에 `[대량변경]` 을 넣는다.
