/* WTU 엔진 — 재무팀 회계원본(상품/제품/기타매출, 차변/대변) → cons_cache(main/device/intl)
 * 순수 계산(네트워크·DOM 없음). 브라우저(window.WTU) + Node(module.exports) 공용.
 * rebuild_from_raw.py 충실 포팅 + 국내/해외=거래처명 접두사 OR 사용부서. 2026만 계산, 2025는 기존 cons_cache 머지.
 */
(function(root){
'use strict';
// ───────── 헬퍼 ─────────
function num(v){ if(v===null||v===undefined||v==='') return 0; var n=parseFloat(v); return isNaN(n)?0:n; }
function pint(v){ if(v===null||v===undefined||v==='') return 0; var n=parseInt(parseFloat(v),10); if(!isNaN(n)) return n; var m=String(v).match(/-?\d+/); return m?parseInt(m[0],10):0; }
function monthNum(v){ if(!v&&v!==0) return null; var m=String(v).match(/(\d+)/); return m?parseInt(m[1],10):null; }
function moN(v){ if(!v&&v!==0) return 0; var d=String(v).replace(/[^0-9]/g,''); var n=d?parseInt(d,10):0; return (n>=1&&n<=12)?n:0; }
function toDt(v){ if(v===null||v===undefined||v==='') return null; if(v instanceof Date) return isNaN(v.getTime())?null:v;
  var s=String(v).slice(0,19); var d=new Date(s.replace(/\./g,'-').replace(/\//g,'-')); return isNaN(d.getTime())?null:d; }
function moOf(v){ var d=toDt(v); return d?d.getMonth()+1:null; }
function dsOf(v){ var d=toDt(v); if(!d) return null; var y=d.getFullYear(),m=d.getMonth()+1,da=d.getDate(); return y+'-'+(m<10?'0'+m:m)+'-'+(da<10?'0'+da:da); }
function topN(o,n){ n=n||10; var e=Object.keys(o).map(function(k){return [k,o[k]];}); e.sort(function(a,b){return b[1]-a[1];}); var out={}; e.slice(0,n).forEach(function(p){out[p[0]]=p[1];}); return out; }
function sortTop(o,n){ return topN(o,n); }
function maxKey(o){ var bk=null,bv=-Infinity; for(var k in o){ if(o[k]>bv){bv=o[k];bk=k;} } return bk; }
function addInto(dst,src){ for(var k in src){ dst[k]=(dst[k]||0)+src[k]; } return dst; }

var FOR=/^[가-힣A-Za-z]+\s*\//, SUB=/^\s*\[/, USD=/\$\s*([\d,]+)/;
// ── 데이터 오분류 보정(회계 원본 라벨 오류) — 사이트 딜러뷰/파트 override와 동일 ──
var INTL_COUNTRY_FIX={'더마케이':'러시아'};  // 국내법인 라벨이나 실제 러시아 수출(간접수출) → 러시아(region_map 러시아=유럽=유럽·중아파트)
var INTL_EXCLUDE=['에스준의원'];             // 국내 병원이 해외로 오분류 → 국내로 되돌림(해외집계 제외)
function _intlCtryFix(nm){ for(var k in INTL_COUNTRY_FIX){ if(nm.indexOf(k)>=0) return INTL_COUNTRY_FIX[k]; } return null; }

// ───────── CATMAP (제품→소모품/제품군 역매핑, 기존 cons_cache 기준) ─────────
function buildCatMap(main,device,intl){
  var CM={}; main=main||{}; device=device||{}; intl=intl||{};
  var I=(intl.INTL2026)||{};
  Object.keys(main.products2026||{}).forEach(function(n){ CM[n]='소모품'; });
  Object.keys(device.products2026||{}).forEach(function(n){ CM[n]='제품군'; });
  Object.keys(I.prods_cons||{}).forEach(function(n){ if(!(n in CM)) CM[n]='소모품'; });
  Object.keys(I.prods_dev||{}).forEach(function(n){ if(!(n in CM)) CM[n]='제품군'; });
  return CM;
}
function catOf(prod,CM){
  if(prod in CM) return CM[prod];
  if(prod.indexOf('의료기기')>=0) return '의료기기';
  if(/tip|팁|cartridge|카트리지|소모품|spray|goggle|switch|oil|오일|패드|fiber|bare|air ?gap|side fiber/i.test(prod)) return '소모품';
  return '제품군';
}
// 제품명 별칭 통일 — 회계 원본 표기 흔들림(버전 접미사)을 대표명으로 정규화. 새 흔들림 발견 시 여기 추가
var PROD_ALIAS={'Ultraskin Tightan II':'Ultraskin Tightan','SANDRO Dual-N':'SANDRO Dual'};
function normProd(p){ p=String(p||'').trim(); return PROD_ALIAS[p]||p; }
function prodName(proj){ var s=String(proj||''); var i=s.indexOf('.'); return normProd(i>=0?s.slice(i+1).trim():s.trim()); }

// ───────── 재무팀 raw → 표준행(2026) ─────────
// sheets: [{name, rows:[{컬럼:값}]}]  (rows = XLSX.utils.sheet_to_json 결과)
function transformRaw(sheets, CM){
  var out=[];
  sheets.forEach(function(sh){
    (sh.rows||[]).forEach(function(row){
      var jeok=String(row['적요']==null?'':row['적요']);
      var cust=String(row['거래처명']==null?'':row['거래처명']).trim();
      if(SUB.test(jeok) || !cust) return;
      var net=num(row['대변'])-num(row['차변']);
      if(net===0) return;
      var prod=prodName(row['프로젝트']); var cat=catOf(prod,CM);
      var dept=String(row['사용부서']==null?'':row['사용부서']);
      // 국내/해외: 거래처명 "국가/" 접두사 OR 사용부서(해외법인영업·해외CS·수출·구매확인서) — 더마케이 등 간접수출 포착
      var region=(FOR.test(cust)||dept.indexOf('해외')>=0||dept.indexOf('수출')>=0||dept.indexOf('구매확인서')>=0)?'해외':'국내';
      for(var _ex=0;_ex<INTL_EXCLUDE.length;_ex++){ if(cust.indexOf(INTL_EXCLUDE[_ex])>=0){ region='국내'; break; } }  // 국내 병원 오분류 보정(에스준의원)
      var dt=toDt(row['승인일자']); var mo=dt?dt.getMonth()+1:null;
      var um=USD.exec(jeok); var usd=um?parseFloat(um[1].replace(/,/g,'')):0;
      out.push({'월':mo,'금액':net,'국내외':region,'구분#2':cat,'구분#4':prod,'구분#1':prod,
        '분기':mo?(Math.floor((mo-1)/3)+1):null,'거래처명':cust,'담당자':row['담당자2'],
        '수량':row['수량'],'USD':usd,'날짜':dt,'사용부서':dept});
    });
  });
  return out;
}

// ───────── CONS (소모품 국내) 2026 ─────────
function aggregateCons(rows){
  var monthly={},products={},groups={},qtr={},oligioMo={};
  rows.forEach(function(r){
    var amt=num(r['금액']); if(!amt) return;
    var mo=monthNum(r['월']),prod=r['구분#4'],grp=r['구분#1'],q=r['분기'];
    if(mo) monthly[mo]=(monthly[mo]||0)+amt;
    if(prod) products[prod]=(products[prod]||0)+amt;
    if(grp) groups[grp]=(groups[grp]||0)+amt;
    if(q) qtr[q]=(qtr[q]||0)+amt;
    if(mo&&prod){ var og=null;
      if(/Oligio Tip \[(F|E)-/.test(String(prod))) og='올리지오';
      else if(/Oligio Tip \[(XF|XE)-/.test(String(prod))) og='X/XM';
      if(og){ oligioMo[og]=oligioMo[og]||{}; oligioMo[og][mo]=(oligioMo[og][mo]||0)+amt; }
    }
  });
  return {monthly:monthly,products:products,groups:groups,qtr:qtr,oligioMo:oligioMo};
}
function groups26final(groups){
  var e=Object.keys(groups).map(function(k){return [k,groups[k]];}).sort(function(a,b){return b[1]-a[1];});
  var top=e.slice(0,7), etc=e.slice(7).reduce(function(s,p){return s+p[1];},0);
  var out={}; top.forEach(function(p){out[p[0]]=p[1];}); if(etc>0) out['기타']=etc; return out;
}
// 2026 병원 레코드(소모품)
function consHosp2026(rows){
  var INT=['원텍','wontech','Wontech'], EXCL={'의료기기':1};
  var hosp={};
  rows.forEach(function(r){
    var nm=r['거래처명']; if(!nm) return;
    if(INT.some(function(k){return String(nm).indexOf(k)>=0;})) return;
    var amt2=num(r['금액']); if(!amt2) return;
    var q2=pint(r['수량']); var qty2=q2>0?q2:null;
    var prod=r['구분#4']||''; if(prod in EXCL) return;
    var sales=r['담당자']||'';
    var rd=r['날짜'], mo=moOf(rd), ds=dsOf(rd);
    var h=hosp[nm]; if(!h){ h={y2026:0,orders26:0,prodsAll:{},m2026:{},prods2026:{},txns2026:[],lastDate:null,_sc:{}}; hosp[nm]=h; }
    h.y2026+=amt2; h.orders26++;
    if(prod){ h.prodsAll[prod]=(h.prodsAll[prod]||0)+amt2; h.prods2026[prod]=(h.prods2026[prod]||0)+amt2; }
    if(sales) h._sc[sales]=(h._sc[sales]||0)+1;
    if(mo) h.m2026[mo]=(h.m2026[mo]||0)+amt2;
    if(ds) h.txns2026.push([mo||0,ds,prod,amt2,qty2]);
    if(ds&&(!h.lastDate||ds>h.lastDate)) h.lastDate=ds;
  });
  Object.keys(hosp).forEach(function(nm){ var h=hosp[nm];
    h.sales=maxKey(h._sc)||''; delete h._sc;
    h.txns2026.sort(function(a,b){return a[1]<b[1]?-1:1;});
  });
  return hosp;
}

// ───────── DEVICE (제품군 국내) 2026 ─────────
function aggDevice2026(rows26){
  var monthly={},products={},prodMo={},monthlyQty={},prodQty={},prodMoQty={};
  var devHosp={},devReps={};
  rows26.forEach(function(r){
    var amt=num(r['금액']); if(amt===0) return;     // 음수(차감) 유지=net
    var qtyRaw=pint(r['수량']);
    var qtyForExcl=qtyRaw>0?qtyRaw:1;               // 저단가 제외 판정용(빈칸=1로 가정해 판정만)
    var qty=qtyRaw>0?qtyRaw:0;                      // 실제 대수: 빈 수량은 0(세지 않음, 재무 원본 기준) — 금액은 아래서 그대로 반영
    var mo=monthNum(r['월']); var prod=(r['구분#4']||'').trim();
    if(/hair|헤어/i.test(prod)) prod='Hair Beam';
    else if(amt>0 && amt/qtyForExcl<1000000) return;       // 양수 저단가(팁/소모품) 제외
    var name=(r['거래처명']||'').trim(), sales=(r['담당자']||'').trim();
    var ds=dsOf(r['날짜'])||'';
    if(mo){ monthly[mo]=(monthly[mo]||0)+amt; monthlyQty[mo]=(monthlyQty[mo]||0)+qty; }
    if(prod){ products[prod]=(products[prod]||0)+amt; prodQty[prod]=(prodQty[prod]||0)+qty; }
    if(mo&&prod){ prodMo[prod]=prodMo[prod]||{}; prodMo[prod][mo]=(prodMo[prod][mo]||0)+amt;
      prodMoQty[prod]=prodMoQty[prod]||{}; prodMoQty[prod][mo]=(prodMoQty[prod][mo]||0)+qty; }
    if(name){ var dh=devHosp[name]; if(!dh){ dh={y2026:0,orders26:0,prodsAll:{},prods2026:{},m2026:{},txns2026:[],lastDate:'',sales:''}; devHosp[name]=dh; }
      dh.y2026+=amt; dh.orders26++;
      if(prod){ dh.prodsAll[prod]=(dh.prodsAll[prod]||0)+amt; dh.prods2026[prod]=(dh.prods2026[prod]||0)+amt; }
      if(mo) dh.m2026[mo]=(dh.m2026[mo]||0)+amt;
      if(mo&&prod&&ds) dh.txns2026.push([mo,ds,name,prod,Math.round(amt),qty]);
      if(ds&&ds>dh.lastDate) dh.lastDate=ds;
      if(sales) dh.sales=sales;
    }
    if(sales){ var dr=devReps[sales]; if(!dr){ dr={y2026:0,c26:{},prods2026:{},m2026:{},qty2026:0}; devReps[sales]=dr; }
      dr.y2026+=amt; dr.c26[name||'?']=1;
      if(prod) dr.prods2026[prod]=(dr.prods2026[prod]||0)+amt;
      if(mo) dr.m2026[mo]=(dr.m2026[mo]||0)+amt;
      dr.qty2026+=qty;
    }
  });
  Object.keys(devHosp).forEach(function(n){ var dh=devHosp[n]; dh.txns2026.sort(function(a,b){return a[1]<b[1]?-1:1;}); });
  // 랭킹(products/prodQty)을 월별 상세(prodMo/prodMoQty) 합과 일치시켜 자기모순 제거 + 전 품목 유지(상한 제거)
  var _pAll={},_qAll={};
  Object.keys(prodMo).forEach(function(p){ var s=0,mm=prodMo[p]; Object.keys(mm).forEach(function(m){s+=mm[m];}); _pAll[p]=s; });
  Object.keys(prodMoQty).forEach(function(p){ var s=0,mm=prodMoQty[p]; Object.keys(mm).forEach(function(m){s+=mm[m];}); _qAll[p]=s; });
  return {monthly2026:monthly,monthlyQty2026:monthlyQty,products2026:_pAll,prodQty2026:_qAll,
    prodMo2026:prodMo,prodMoQty2026:prodMoQty,hospitals2026:devHosp,reps2026:devReps};
}

// ───────── INTL (해외) 2026 ─────────
function aggIntl2026(rows26){
  var mo={},mu={},mq={},pr={},pc={},pd={},pm={},pmq={},pq={},ct={},rp={};
  var _muK={},_muU={},_muMiss={};   // USD 미기재 거래 보정용(월별 내재환율)
  rows26.forEach(function(r){
    var amt=num(r['금액']); if(amt===0) return;
    var usd=num(r['USD']); var qty=pint(r['수량']); qty=qty>0?qty:1;
    var m=moN(r['월']); var p=(r['구분#4']||'').trim(); var cat=(r['구분#2']||'').trim();
    var nm=(r['거래처명']||'').trim(); var sl=(r['담당자']||'').trim();
    var co=(!nm)?'기타':(_intlCtryFix(nm)||(FOR.test(nm)?(nm.match(/^([가-힣a-zA-Z]+)\s*\//)[1]):'국내법인'));
    if(m){ mo[m]=(mo[m]||0)+amt; mu[m]=(mu[m]||0)+usd; mq[m]=(mq[m]||0)+qty;
      if(usd){ _muK[m]=(_muK[m]||0)+amt; _muU[m]=(_muU[m]||0)+usd; } else { _muMiss[m]=(_muMiss[m]||0)+amt; } }
    if(p){ pr[p]=(pr[p]||0)+amt; pq[p]=(pq[p]||0)+qty;
      if(cat==='소모품') pc[p]=(pc[p]||0)+amt; else if(cat==='제품군') pd[p]=(pd[p]||0)+amt; }
    if(m&&p){ pm[p]=pm[p]||{}; pm[p][m]=(pm[p][m]||0)+amt; pmq[p]=pmq[p]||{}; pmq[p][m]=(pmq[p][m]||0)+qty; }
    if(co){ var c=ct[co]; if(!c){ c={total:0,usd:0,orders:0,clients:{},prods:{},m:{},clientData:{}}; ct[co]=c; }
      c.total+=amt; c.usd+=usd; c.orders++; c.clients[nm]=1;
      if(p) c.prods[p]=(c.prods[p]||0)+amt; if(m) c.m[m]=(c.m[m]||0)+amt;
      var cd=c.clientData[nm]; if(!cd){ cd={total:0,m:{},prods:{},cnt:0,last:"",txns:[]}; c.clientData[nm]=cd; } cd.total+=amt; if(m) cd.m[m]=(cd.m[m]||0)+amt; if(p)cd.prods[p]=(cd.prods[p]||0)+amt; cd.cnt++;
      var _cds=dsOf(r["날짜"])||""; if(_cds>cd.last)cd.last=_cds;
      cd.txns.push([_cds,p||"",pint(r["수량"]),Math.round(amt),Math.round(usd)]);   // 거래 원장 [날짜,제품,수량,금액,USD]
    }
    if(sl){ var s=rp[sl]; if(!s){ s={total:0,usd:0,orders:0,clients:{},cset:{},prods:{},m:{}}; rp[sl]=s; }
      s.total+=amt; s.usd+=usd; s.orders++; s.clients[nm]=1; s.cset[co]=1;
      if(p) s.prods[p]=(s.prods[p]||0)+amt; if(m) s.m[m]=(s.m[m]||0)+amt;
    }
  });
  function fct(ctm){ var out={}; Object.keys(ctm).map(function(k){return [k,ctm[k]];}).sort(function(a,b){return b[1].total-a[1].total;}).forEach(function(p){
    var v=p[1]; var cdo={}; Object.keys(v.clientData).forEach(function(ck){ var cv=v.clientData[ck]; cdo[ck]={total:Math.round(cv.total),m:cv.m,prods:topN(cv.prods||{},5),cnt:cv.cnt||0,last:cv.last||"",txns:(cv.txns||[]).sort(function(a,b){return a[0]<b[0]?-1:1;})}; });
    out[p[0]]={total:Math.round(v.total),usd:Math.round(v.usd),orders:v.orders,clients:Object.keys(v.clients).length,prods:topN(v.prods,4),m:v.m,clientData:cdo};
  }); return out; }
  function frp(rpm){ var out={}; Object.keys(rpm).map(function(k){return [k,rpm[k]];}).sort(function(a,b){return b[1].total-a[1].total;}).forEach(function(p){
    var v=p[1]; out[p[0]]={total:Math.round(v.total),usd:Math.round(v.usd),orders:v.orders,countries:Object.keys(v.cset).length,clients:Object.keys(v.clients).length,prods:topN(v.prods,4),m:v.m};
  }); return out; }
  // USD 미기재 거래(적요에 $ 없음 — 반품 미차감·신규출하 등) 보정: 해당월 USD 병기 거래의 내재환율로 환산해 monthly_usd에 합산.
  // 내재환율이 비정상(800~2500 밖)이거나 표본이 없으면 1450 고정. 부호 보존(반품 음수는 차감).
  Object.keys(_muMiss).forEach(function(m){
    var rt=(_muU[m]>0&&_muK[m]>0)?(_muK[m]/_muU[m]):0;
    if(!(rt>=800&&rt<=2500)) rt=1450;
    mu[m]=(mu[m]||0)+_muMiss[m]/rt;
  });
  Object.keys(mu).forEach(function(m){ mu[m]=Math.round(mu[m]); });
  function sum(o){ var s=0; for(var k in o) s+=o[k]; return s; }
  return {monthly:mo,monthly_usd:mu,monthly_qty:mq,total:sum(mo),total_usd:sum(mu),
    products:topN(pr,12),prods_cons:topN(pc,10),prods_dev:topN(pd,10),prodQty:topN(pq,12),
    prodMo:pm,prodMoQty:pmq,countries:fct(ct),reps:frp(rp)};
}

// ───────── 2025 머지 ─────────
function mergeConsHospitals(existing,new26){
  existing=existing||{}; var names={}; Object.keys(existing).forEach(function(n){names[n]=1;}); Object.keys(new26).forEach(function(n){names[n]=1;});
  var out={};
  Object.keys(names).forEach(function(n){
    var E=existing[n]||{}, N=new26[n]||{};
    var y2025=E.y2025||0, m2025=E.m2025||{}, prods2025=E.prods2025||{}, txns2025=E.txns2025||[];
    var y2026=N.y2026||0, m2026=N.m2026||{}, txns2026=N.txns2026||[];
    var prods2026=N.prodsAll?topN(N.prodsAll,4):(N.prods2026||{});
    var comb={}; addInto(comb,prods2025); addInto(comb,(N.prodsAll||N.prods2026||{}));
    var prods=topN(comb,4); var topProd=maxKey(comb)||E.topProd||'';
    out[n]={total:y2025+y2026,y2025:y2025,y2026:y2026,orders:(txns2025.length+txns2026.length)||E.orders||0,
      prods:prods,prods2025:prods2025,prods2026:prods2026,m2025:m2025,m2026:m2026,
      txns2025:txns2025,txns2026:txns2026,lastDate:N.lastDate||E.lastDate||null,sales:N.sales||E.sales||'',topProd:topProd};
  });
  return out;
}
function mergeDevHospitals(existing,new26){
  existing=existing||{}; var names={}; Object.keys(existing).forEach(function(n){names[n]=1;}); Object.keys(new26).forEach(function(n){names[n]=1;});
  var out={};
  Object.keys(names).forEach(function(n){
    var E=existing[n]||{}, N=new26[n]||{};
    var y2025=E.y2025||0, m2025=E.m2025||{}, prods2025=E.prods2025||{}, txns2025=E.txns2025||[];
    var y2026=N.y2026||0, m2026=N.m2026||{}, txns2026=N.txns2026||[];
    var prods2026=N.prodsAll?topN(N.prodsAll,4):(N.prods2026||{});
    var comb={}; addInto(comb,prods2025); addInto(comb,(N.prodsAll||N.prods2026||{}));
    out[n]={total:y2025+y2026,y2025:y2025,y2026:y2026,orders:(txns2025.length+txns2026.length)||E.orders||0,
      prods:topN(comb,4),prods2025:prods2025,prods2026:prods2026,m2025:m2025,m2026:m2026,
      txns2025:txns2025,txns2026:txns2026,lastDate:N.lastDate||E.lastDate||'',sales:N.sales||E.sales||'',topProd:maxKey(comb)||E.topProd||''};
  });
  return out;
}
function mergeDevReps(existing,new26){
  existing=existing||{}; var names={}; Object.keys(existing).forEach(function(n){names[n]=1;}); Object.keys(new26).forEach(function(n){names[n]=1;});
  var out={};
  Object.keys(names).forEach(function(n){
    var E=existing[n]||{}, N=new26[n]||{};
    var y2025=E.y2025||0, y2026=N.y2026||0;
    out[n]={total:Math.round(y2025+y2026),y2025:Math.round(y2025),y2026:Math.round(y2026),
      cust2025:E.cust2025||0,cust2026:N.c26?Object.keys(N.c26).length:0,
      prods2025:E.prods2025||{},prods2026:N.prods2026?topN(N.prods2026,6):{},
      m2025:E.m2025||{},m2026:N.m2026||{},qty2025:E.qty2025||0,qty2026:N.qty2026||0};
  });
  return out;
}

// ───────── 메인 빌드 ─────────
// existing = {main, device, intl} (기존 cons_cache data). 반환 {main, device, intl, summary}
function build(sheets, existing){
  existing=existing||{}; var exMain=existing.main||{}, exDev=existing.device||{}, exIntl=existing.intl||{};
  var CM=buildCatMap(exMain,exDev,exIntl);
  var rows26=transformRaw(sheets,CM);
  // 국내영업 제외 부서: 서지컬·B2C·CS(고객만족)·비영업 지원부서
  var EXCL_DEPT=/지컬|B2C|CS|연구|구매팀|인사|총무/;
  var _incl=function(r){return !EXCL_DEPT.test(String(r['사용부서']||''));};
  var consR=rows26.filter(function(r){return r['구분#2']==='소모품'&&r['국내외']==='국내'&&_incl(r);});
  var devR =rows26.filter(function(r){return r['구분#2']==='제품군'&&r['국내외']==='국내'&&_incl(r);});
  var intlR=rows26.filter(function(r){return r['국내외']==='해외';});
  // updatedAt: 소모품 2026 날짜 max (없으면 전체 rows26 max)
  var dates=consR.map(function(r){return r['날짜'];}).filter(Boolean);
  if(!dates.length) dates=rows26.map(function(r){return r['날짜'];}).filter(Boolean);
  var md=dates.length?new Date(Math.max.apply(null,dates.map(function(d){return d.getTime();}))):new Date();
  var pad=function(x){return x<10?'0'+x:''+x;};
  var updatedAt=md.getFullYear()+'.'+pad(md.getMonth()+1)+'.'+pad(md.getDate());

  // CONS
  var a26=aggregateCons(consR);
  var consH26=consHosp2026(consR);
  var CONS={ monthly2025:exMain.monthly2025||{}, monthly2026:a26.monthly,
    products2025:exMain.products2025||{}, products2026:topN(a26.products,10),
    groups2026:groups26final(a26.groups), qtr2025:exMain.qtr2025||{}, qtr2026:a26.qtr,
    oligioMo2025:exMain.oligioMo2025||{}, oligioMo2026:a26.oligioMo,
    hospitals:mergeConsHospitals(exMain.hospitals,consH26), updatedAt:updatedAt };
  // DEVICE
  var d26=aggDevice2026(devR);
  var DEVICE={ monthly2025:exDev.monthly2025||{}, monthly2026:d26.monthly2026,
    monthlyQty2025:exDev.monthlyQty2025||{}, monthlyQty2026:d26.monthlyQty2026,
    products2025:exDev.products2025||{}, products2026:d26.products2026,
    prodQty2025:exDev.prodQty2025||{}, prodQty2026:d26.prodQty2026,
    prodMo2025:exDev.prodMo2025||{}, prodMo2026:d26.prodMo2026,
    prodMoQty2025:exDev.prodMoQty2025||{}, prodMoQty2026:d26.prodMoQty2026,
    hospitals:mergeDevHospitals(exDev.hospitals,d26.hospitals2026),
    reps:mergeDevReps(exDev.reps,d26.reps2026), updatedAt:updatedAt };
  // INTL
  var i26=aggIntl2026(intlR);
  var INTL={ INTL2026:i26, INTL2025:(exIntl.INTL2025||{}), updatedAt:updatedAt };

  // 기타(종합 전용): 국내 전체 net − 국내소모품 − 국내제품 = 서지컬·B2C·고객만족·의료기기·지원부서.
  // 한국영업 분석(소모품/제품)엔 미포함, 종합사이트 전사 합계에만 얹음 → 회계 전체와 정합.
  var _domAll={};
  rows26.forEach(function(r){ if(r['국내외']==='국내'){ var _m=monthNum(r['월']); var _a=num(r['금액']); if(_m&&_a) _domAll[_m]=(_domAll[_m]||0)+_a; } });
  var _etcMo={};
  for(var _em=1;_em<=12;_em++){ var _ev=Math.round((_domAll[_em]||0)-(a26.monthly[_em]||0)-(d26.monthly2026[_em]||0)); if(_ev) _etcMo[_em]=_ev; }
  CONS.etcMonthly2026=_etcMo; CONS.etcMonthly2025=exMain.etcMonthly2025||{};
  // 기타를 사용부서별로 분리(서지컬/B2C/고객만족/기타=의료기기·조정) — 매출현황 서지컬 행 표시용. 버킷합≠잔차분은 '기타'로 흡수해 etcMonthly와 정합 유지.
  var _incSet=new Set(); consR.forEach(function(r){_incSet.add(r);}); devR.forEach(function(r){_incSet.add(r);});
  var _dept={'서지컬':{},'B2C':{},'고객만족':{},'기타':{}};
  rows26.forEach(function(r){ if(r['국내외']!=='국내'||_incSet.has(r)) return;
    var _m2=monthNum(r['월']); var _a2=num(r['금액']); if(!_m2||!_a2) return;
    var _d2=String(r['사용부서']||''); var _b=/지컬/.test(_d2)?'서지컬':(/B2C/i.test(_d2)?'B2C':(/고객만족|CS/.test(_d2)?'고객만족':'기타'));
    _dept[_b][_m2]=(_dept[_b][_m2]||0)+_a2; });
  for(var _em2=1;_em2<=12;_em2++){ var _bs=0; for(var _bk in _dept) _bs+=(_dept[_bk][_em2]||0); var _adj=(_etcMo[_em2]||0)-_bs; if(Math.round(_adj)) _dept['기타'][_em2]=(_dept['기타'][_em2]||0)+_adj; }
  Object.keys(_dept).forEach(function(_k){ var _o={}; Object.keys(_dept[_k]).forEach(function(_m3){ var _v3=Math.round(_dept[_k][_m3]); if(_v3) _o[_m3]=_v3; }); _dept[_k]=_o; });
  CONS.etcDept2026=_dept; CONS.etcDept2025=exMain.etcDept2025||{};

  var e8=function(x){return Math.round(x/1e8*10)/10;};
  var sum=function(o){var s=0;for(var k in o)s+=o[k];return s;};
  var dermak = ('주식회사 더마케이' in CONS.hospitals)||('주식회사 더마케이' in DEVICE.hospitals);
  var summary={ updatedAt:updatedAt,
    cons26:e8(sum(a26.monthly)), consHosp:Object.keys(CONS.hospitals).length,
    dev26:e8(sum(d26.monthly2026)), devHosp:Object.keys(DEVICE.hospitals).length, devReps:Object.keys(DEVICE.reps).length,
    intl26:e8(i26.total), intlCountries:Object.keys(i26.countries).length,
    us:e8((i26.countries['미국']||{}).total||0),
    domestic:e8(sum(a26.monthly)+sum(d26.monthly2026)),
    etc26:e8(sum(_etcMo)),
    total:e8(sum(a26.monthly)+sum(d26.monthly2026)+sum(_etcMo)+i26.total),
    dermakInDomestic:dermak,
    monConsT:a26.monthly, monDevT:d26.monthly2026, monIntlT:i26.monthly,
    rows26:rows26.length };
  // rev=내용해시 → 같은날 재기록해도 클라이언트 캐시 자동갱신(updatedAt 표시용 유지)
  var _hash=function(s){ var h=5381,i=s.length; while(i) h=(h*33)^s.charCodeAt(--i); return (h>>>0).toString(36); };
  [CONS,DEVICE,INTL].forEach(function(o){ delete o.rev; o.rev=updatedAt+'#'+_hash(JSON.stringify(o)); });
  return {main:CONS, device:DEVICE, intl:INTL, summary:summary};
}

var WTU={ build:build, transformRaw:transformRaw, buildCatMap:buildCatMap,
  aggregateCons:aggregateCons, aggDevice2026:aggDevice2026, aggIntl2026:aggIntl2026 };
if(typeof module!=='undefined'&&module.exports) module.exports=WTU;
root.WTU=WTU;
})(typeof window!=='undefined'?window:(typeof global!=='undefined'?global:this));
