/* WTILBO — 일일매출일보(ilbo_acc) 집계 엔진
 * wt-consumables/ilbo.html 의 parseAccounting 을 verbatim 포팅.
 * 재무팀 회계원본(차변/대변/사용부서/품번) → ilbo_acc 구조. upload.html 통합용.
 * XLSX 전역 필요(브라우저). 순수계산.
 */
(function(root){
'use strict';
var DEPT_MAP={'A0402040':'국내영업','A0402020':'국내영업','A0402030':'국내영업','A0402010':'국내영업','A0402050':'국내영업','A0401020':'국내영업','A0401040':'국내영업','A0401070':'국내영업','A0401030':'해외영업','A0401050':'Surgical','B0901030':'Surgical','A0501020':'고객만족','A0501030':'고객만족','A0501010':'고객만족','A0401060':'B2C','A0701020':'기타'};
var CONS_RE=/tip|cartridge|cathet|카트리지|팁|니들|needle|소모|consum|spray|쿨링/i;
var DIVS=['국내제품','국내소모품','해외영업','Surgical','B2C','고객만족','기타'];
function divRegion(div){ return div==='해외영업'?'해외':'국내'; }
function prodName(pn){ pn=String(pn||''); var i=pn.indexOf('.'); return (i>=0?pn.slice(i+1):pn).trim(); }
function classify(deptStr,pnum){ var code=String(deptStr||'').split('.')[0]; var div=DEPT_MAP[code]||'기타'; if(div==='국내영업') div=CONS_RE.test(String(pnum||''))?'국내소모품':'국내제품'; return div; }

function parseAccounting(wb,fname){
  var monthly={}; DIVS.forEach(function(d){monthly[d]={};});
  var consMonthly={};
  var lines=[]; var products={'국내':{},'해외':{}};
  wb.SheetNames.forEach(function(sn){
    var rows=XLSX.utils.sheet_to_json(wb.Sheets[sn],{header:1,raw:true}); var cur=null;
    for(var i=1;i<rows.length;i++){
      var r=rows[i]; if(!r) continue;
      var adq=r[0],memo=r[1],cname=r[3],debit=Number(r[5])||0,credit=Number(r[6])||0,dept=r[8],pnum=r[10],qty=Number(r[11])||0;
      if(typeof adq==='string'&&/^\d{4}-\d{2}-\d{2}/.test(adq)) cur=adq.slice(0,10);
      if(dept==null||dept==='') continue;
      if(typeof memo==='string'&&memo.trim().charAt(0)==='[') continue;
      var d=(typeof adq==='string'&&/^\d{4}/.test(adq))?adq.slice(0,10):cur; if(!d) continue;
      var mo=parseInt(d.slice(5,7),10); var amt=credit-debit; var div=classify(dept,pnum);
      monthly[div][mo]=(monthly[div][mo]||0)+amt;
      if(CONS_RE.test(String(pnum||''))) consMonthly[mo]=(consMonthly[mo]||0)+amt;
      lines.push({date:d,month:mo,region:divRegion(div),div:div,customer:String(cname||''),product:prodName(pnum),qty:qty,amount:amt});
    }
  });
  var rd=null,m=String(fname||'').match(/(\d{4})\D+(\d{1,2})\D+(\d{1,2})/);
  if(m) rd=m[1]+'-'+('0'+m[2]).slice(-2)+'-'+('0'+m[3]).slice(-2);
  if(!rd) lines.forEach(function(l){ if(!rd||l.date>rd) rd=l.date; });
  var curMo=rd?parseInt(rd.slice(5,7),10):new Date().getMonth()+1;
  // 월 롤오버 가드: 당월(파일명 월)에 매출 데이터가 없으면 데이터 있는 마지막 월로 캡 (총합계 무영향, 당월 표시/미반영대조 정상화)
  var _dm=lines.map(function(l){return l.month;}).filter(Boolean); var _maxMo=_dm.length?Math.max.apply(null,_dm):curMo; if(curMo>_maxMo) curMo=_maxMo;
  var cl=lines.filter(function(l){return l.month===curMo;});
  return {reportDate:rd,curMonth:curMo,monthly:monthly,consMonthly:consMonthly,lines:cl,products:products,updatedAt:rd};
}
function snapFig(acc){
  var cur=acc.curMonth; var tot=0; DIVS.forEach(function(d){ tot+=(acc.monthly[d]&&acc.monthly[d][cur])||0; });
  var cons=(acc.consMonthly&&acc.consMonthly[cur])||0;
  return {date:acc.reportDate,tot:tot,cons:cons,prod:tot-cons};
}
// 전일 스냅샷 보존 포함 빌드 (ilbo.html handleUpload 로직과 동일)
function buildIlbo(wb,fname,prevAcc){
  var parsed=parseAccounting(wb,fname);
  if(!parsed.reportDate) throw new Error('일일보고: 날짜를 인식하지 못했습니다.');
  if(prevAcc&&prevAcc.reportDate&&prevAcc.reportDate!==parsed.reportDate) parsed.prev=snapFig(prevAcc);
  else if(prevAcc&&prevAcc.prev) parsed.prev=prevAcc.prev;
  return parsed;
}
var WTILBO={ buildIlbo:buildIlbo, parseAccounting:parseAccounting, snapFig:snapFig, DIVS:DIVS };
if(typeof module!=='undefined'&&module.exports) module.exports=WTILBO;
root.WTILBO=WTILBO;
})(typeof window!=='undefined'?window:(typeof global!=='undefined'?global:this));
