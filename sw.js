/* wontech 앱 서비스워커 — 일반 새로고침에도 항상 최신 HTML을 받도록(network-first).
   페이지/문서 요청은 캐시를 무시하고 서버에서 새로 받고(cache:'reload'),
   네트워크 실패 시에만 브라우저 기본 동작으로 폴백. 앱 자원을 별도 캐싱하지 않아
   '옛 버전이 계속 뜨는' 서비스워커 함정을 피함. */
self.addEventListener('install', function(e){ self.skipWaiting(); });
self.addEventListener('activate', function(e){ e.waitUntil(self.clients.claim()); });
self.addEventListener('fetch', function(e){
  var req = e.request;
  if(req.mode === 'navigate'){
    e.respondWith(
      fetch(req, {cache: 'reload'}).catch(function(){ return fetch(req); })
    );
  }
});
