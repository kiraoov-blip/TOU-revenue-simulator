console.info("PRAS - TOU weekend discount control update loaded");
const DATA_URL="./data/tou_data.xlsx",PERIODS=["경부하","중간부하","최대부하"],PERIOD_CLASS={"경부하":"off","중간부하":"mid","최대부하":"peak"},SEASONS=["하계","춘추계","동계"],DAY_TYPES=["평일","토요일","일·공휴일"];
const state={workbook:null,catalog:[],tariffs:[],schedules:[],usage:[],selectedCatalog:null,activeSeason:"하계",scenarioSchedule:null,scenarioDiscount:null,scenarioRates:null,lastResult:null};
const $=id=>document.getElementById(id),clone=o=>JSON.parse(JSON.stringify(o)),text=v=>String(v??"").trim(),number=v=>{const n=Number(String(v??"").replace(/,/g,""));return Number.isFinite(n)?n:0};
const normalizeDate=v=>{if(v instanceof Date)return v;const s=text(v).replace(/[^0-9]/g,"");return s.length===8?new Date(+s.slice(0,4),+s.slice(4,6)-1,+s.slice(6,8)):null},ymd=d=>d?`${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,"0")}-${String(d.getDate()).padStart(2,"0")}`:"-";
const fmtWon=n=>{const a=Math.abs(n),s=n<0?"-":"";if(a>=1e12)return s+(a/1e12).toLocaleString("ko-KR",{maximumFractionDigits:3})+"조원";if(a>=1e8)return s+(a/1e8).toLocaleString("ko-KR",{maximumFractionDigits:1})+"억원";if(a>=1e4)return s+(a/1e4).toLocaleString("ko-KR",{maximumFractionDigits:1})+"만원";return s+Math.round(a).toLocaleString("ko-KR")+"원"};
const fmtEnergy=n=>{const a=Math.abs(n),s=n<0?"-":"";if(a>=1e9)return s+(a/1e9).toLocaleString("ko-KR",{maximumFractionDigits:2})+"TWh";if(a>=1e6)return s+(a/1e6).toLocaleString("ko-KR",{maximumFractionDigits:2})+"GWh";if(a>=1e3)return s+(a/1e3).toLocaleString("ko-KR",{maximumFractionDigits:2})+"MWh";return s+a.toLocaleString("ko-KR",{maximumFractionDigits:1})+"kWh"};
const signedWon=n=>(n>=0?"+":"")+fmtWon(n),cssSign=n=>n>=0?"pos":"neg";
function sheetRows(wb,name){const ws=wb.Sheets[name];return ws?XLSX.utils.sheet_to_json(ws,{defval:"",raw:true}):[]}
async function loadGithubData(){setStatus("loading","GitHub 데이터 불러오는 중",DATA_URL);try{const res=await fetch(`${DATA_URL}?t=${Date.now()}`,{cache:"no-store"});if(!res.ok)throw new Error(`HTTP ${res.status}`);loadWorkbook(await res.arrayBuffer(),"GitHub")}catch(e){setStatus("error","GitHub 데이터 불러오기 실패",`${e.message} · Pages에서 실행 중인지, data/tou_data.xlsx가 있는지 확인`)}}
function loadWorkbook(buffer,source){try{if(typeof XLSX==="undefined")throw new Error("SheetJS 라이브러리를 불러오지 못함");const wb=XLSX.read(buffer,{type:"array",cellDates:true});state.workbook=wb;state.catalog=sheetRows(wb,"종별목록");state.tariffs=sheetRows(wb,"요금표");state.schedules=sheetRows(wb,"시간대기준");["종별목록","요금표","시간대기준"].forEach(n=>{if(!wb.Sheets[n])throw new Error(`필수 시트 '${n}' 없음`)});populateCategories();setStatus("ok",`${source} 데이터 연결 완료`,`${wb.SheetNames.length}개 시트 · 사용량 ${state.usage.length.toLocaleString()}일 · ${[...new Set(state.usage.map(r=>r.year))].sort().join(", ")}년`)}catch(e){setStatus("error","엑셀 구조 오류",e.message)}}
function setStatus(type,title,detail){
  const message=`${title}: ${detail}`;
  if(type==="error"){
    console.error(message);
    const warning=$("warning");
    if(warning) warning.innerHTML=`<b>데이터 오류:</b> ${title} · ${detail}`;
  }else{
    console.info(message);
  }
}
function populateCategories(){const rows=state.catalog.filter(r=>text(r["활성화"]).toUpperCase()==="Y"&&state.workbook.Sheets[text(r["사용량시트명"])]);if(!rows.length)throw new Error("활성화된 종별 또는 사용량 시트가 없음");$("category").innerHTML=rows.map(r=>`<option value="${text(r["종별ID"])}">${text(r["표시명"])}</option>`).join("");selectCategory()}
function selectCategory(){
  const id=$("category").value||text(state.catalog[0]["종별ID"]);
  state.selectedCatalog=state.catalog.find(r=>text(r["종별ID"])===id);
  updateWeekendDiscountControl();
  const sheetName=text(state.selectedCatalog["사용량시트명"]);
  const ws=state.workbook.Sheets[sheetName];
  state.usage=parseUsageSheet(ws,sheetName);
  if(!state.usage.length){
    throw new Error(`사용량 시트 '${sheetName}'에서 유효한 날짜별 부하자료를 찾지 못함`);
  }
  populateYears();
  populateTariffDimensions();
  populateVersions();
  loadScenarioPreset();
  updateWarning();
}
function updateWeekendDiscountControl(){
  const categoryId=text(state.selectedCatalog?.["종별ID"]);
  const eligible=["IND_EUL","EV"].includes(categoryId);
  const checkbox=$("useScenarioDiscount");
  const label=$("weekendDiscountLabel");
  checkbox.disabled=!eligible;
  checkbox.checked=eligible;
  if(label){
    label.classList.toggle("disabled",!eligible);
    label.title=eligible
      ? "봄·가을철 토요일·일요일·공휴일 11~14시 전력량요금 50% 할인 적용 여부"
      : "산업용(을)과 전기차충전전력에만 적용되는 항목";
  }
}
function normalizeSeason(v){
  const s=text(v).replace(/\s/g,"");
  if(["하계","여름","여름철"].includes(s)) return "하계";
  if(["춘추계","봄가을","봄·가을","봄가을철","봄·가을철"].includes(s)) return "춘추계";
  if(["동계","겨울","겨울철"].includes(s)) return "동계";
  return s;
}
function normalizeDayType(v){
  const s=text(v).replace(/\s/g,"");
  if(["평일","주중"].includes(s)) return "평일";
  if(["토요일","토"].includes(s)) return "토요일";
  if(["일·공휴일","일/공휴일","일공휴일","일요일·공휴일","일요일/공휴일","일요일","공휴일"].includes(s)) return "일·공휴일";
  return s;
}
function dateFromRaw(v){
  if(v instanceof Date && !Number.isNaN(v.getTime())) return v;
  if(typeof v==="number"){
    const digits=String(Math.trunc(v));
    if(/^\d{8}$/.test(digits)){
      const y=+digits.slice(0,4),m=+digits.slice(4,6),d=+digits.slice(6,8);
      const dt=new Date(y,m-1,d);
      return Number.isNaN(dt.getTime())?null:dt;
    }
    if(v>20000 && v<80000 && typeof XLSX!=="undefined" && XLSX.SSF){
      const p=XLSX.SSF.parse_date_code(v);
      if(p) return new Date(p.y,p.m-1,p.d);
    }
  }
  const digits=text(v).replace(/[^0-9]/g,"");
  if(digits.length>=8){
    const y=+digits.slice(0,4),m=+digits.slice(4,6),d=+digits.slice(6,8);
    const dt=new Date(y,m-1,d);
    return Number.isNaN(dt.getTime())?null:dt;
  }
  return null;
}
function parseUsageSheet(ws,sheetName){
  if(!ws) throw new Error(`사용량 시트 '${sheetName}' 없음`);
  const matrix=XLSX.utils.sheet_to_json(ws,{header:1,defval:"",raw:true});
  if(matrix.length<2) return [];
  const headers=matrix[0].map(v=>text(v).replace(/\s/g,""));
  const findIndex=(aliases)=>headers.findIndex(h=>aliases.includes(h));
  const dateIdx=findIndex(["날짜","일자","일시"]);
  const yearIdx=findIndex(["연도","연"]);
  const monthIdx=findIndex(["월"]);
  const seasonIdx=findIndex(["계절"]);
  const dayIdx=findIndex(["요일","요일구분"]);
  const hourIdx=Array.from({length:24},(_,i)=>findIndex([`H${String(i+1).padStart(2,"0")}`]));
  const missing=[];
  if(dateIdx<0) missing.push("날짜/일자");
  if(seasonIdx<0) missing.push("계절");
  if(dayIdx<0) missing.push("요일");
  hourIdx.forEach((idx,i)=>{if(idx<0) missing.push(`H${String(i+1).padStart(2,"0")}`)});
  if(missing.length) throw new Error(`'${sheetName}' 시트 필수열 누락: ${missing.slice(0,8).join(", ")}${missing.length>8?" 외":""}`);
  const parsed=[];
  for(let r=1;r<matrix.length;r++){
    const row=matrix[r];
    const date=dateFromRaw(row[dateIdx]);
    const rawDateDigits=text(row[dateIdx]).replace(/[^0-9]/g,"");
    let year=yearIdx>=0?number(row[yearIdx]):0;
    if(!year && date) year=date.getFullYear();
    if(!year && rawDateDigits.length>=4) year=number(rawDateDigits.slice(0,4));
    let month=monthIdx>=0?number(row[monthIdx]):0;
    if(!month && date) month=date.getMonth()+1;
    if(!month && rawDateDigits.length>=6) month=number(rawDateDigits.slice(4,6));
    const season=normalizeSeason(row[seasonIdx]);
    const dayType=normalizeDayType(row[dayIdx]);
    const hours=hourIdx.map(idx=>number(row[idx]));
    if(year>=1900 && year<=2200 && SEASONS.includes(season) && DAY_TYPES.includes(dayType)){
      parsed.push({date:date||new Date(year,Math.max(0,month-1),1),year,month,season,dayType,hours});
    }
  }
  return parsed;
}
function populateYears(){const ys=[...new Set(state.usage.map(r=>r.year))].sort((a,b)=>b-a);$("year").innerHTML=ys.map(y=>`<option value="${y}">${y}년</option>`).join("")+(ys.length>1?`<option value="AVG">전체기간 연평균</option>`:"")}
function categoryTariffs(){return state.tariffs.filter(r=>text(r["요금그룹ID"])===text(state.selectedCatalog["요금그룹ID"]))}function categorySchedules(){return state.schedules.filter(r=>text(r["시간대그룹ID"])===text(state.selectedCatalog["시간대그룹ID"]))}function unique(rows,key){return[...new Set(rows.map(r=>text(r[key])).filter(Boolean))]}
function populateTariffDimensions(){
  updateCustomerScopeControls();
}
function updateCustomerScopeControls(){
  const all=$("customerScope").value==="all";
  ["contract","voltage","choice"].forEach(id=>$(id).disabled=all);
  if(all){
    $("contract").innerHTML='<option value="ALL">전체</option>';
    $("voltage").innerHTML='<option value="ALL">전체</option>';
    $("choice").innerHTML='<option value="ALL">전체</option>';
  }else{
    const contracts=unique(categoryTariffs(),"계약전력구간");
    $("contract").innerHTML=contracts.map(v=>`<option>${v}</option>`).join("");
    populateVoltage();
  }
}
function populateVoltage(){
  if($("customerScope").value==="all") return;
  const rows=categoryTariffs().filter(x=>text(x["계약전력구간"])===$("contract").value);
  $("voltage").innerHTML=unique(rows,"전압구분").map(v=>`<option>${v}</option>`).join("");
  populateChoice();
}
function populateChoice(){
  if($("customerScope").value==="all") return;
  const rows=categoryTariffs().filter(x=>
    text(x["계약전력구간"])===$("contract").value &&
    text(x["전압구분"])===$("voltage").value
  );
  $("choice").innerHTML=unique(rows,"선택요금").map(v=>`<option>${v}</option>`).join("");
}
function populateVersions(){const tr=categoryTariffs(),sr=categorySchedules(),ids=unique(tr,"버전ID").filter(id=>sr.some(r=>text(r["버전ID"])===id)),meta=id=>{const r=tr.find(x=>text(x["버전ID"])===id);return`${text(r["버전명"])} (${text(r["적용시작일"])}~${text(r["적용종료일"])})`},opts=ids.map(id=>`<option value="${id}">${meta(id)}</option>`).join("");$("baseVersion").innerHTML=opts;$("scenarioVersion").innerHTML=opts;if(ids.includes("PRE"))$("baseVersion").value="PRE";if(ids.includes("POST"))$("scenarioVersion").value="POST"}
function buildSchedule(versionId){const rows=categorySchedules().filter(r=>text(r["버전ID"])===versionId),schedule={},discount={};SEASONS.forEach(s=>{schedule[s]={};discount[s]={};DAY_TYPES.forEach(d=>{schedule[s][d]=Array(24).fill("경부하");discount[s][d]=Array(24).fill(1)})});rows.forEach(r=>{const s=text(r["계절"]),d=text(r["요일구분"]),h=number(r["시간"]);if(schedule[s]?.[d]&&h>=0&&h<24){schedule[s][d][h]=text(r["부하시간대"]);discount[s][d][h]=number(r["할인율"])||1}});return{schedule,discount}}
function buildRates(versionId){
  let rows=categoryTariffs().filter(r=>text(r["버전ID"])===versionId);
  const all=$("customerScope").value==="all";
  if(!all){
    rows=rows.filter(r=>
      text(r["계약전력구간"])===$("contract").value &&
      text(r["전압구분"])===$("voltage").value &&
      text(r["선택요금"])===$("choice").value
    );
  }
  const values={};
  SEASONS.forEach(s=>{
    values[s]={};
    PERIODS.forEach(p=>values[s][p]=[]);
  });
  rows.forEach(r=>{
    const s=text(r["계절"]),p=text(r["부하시간대"]);
    if(values[s]&&PERIODS.includes(p)) values[s][p].push(number(r["전력량요금"]));
  });
  const rates={};
  SEASONS.forEach(s=>{
    rates[s]={};
    PERIODS.forEach(p=>{
      const list=values[s][p];
      rates[s][p]=list.length
        ? Math.round((list.reduce((a,b)=>a+b,0)/list.length)*10)/10
        : 0;
    });
  });
  return rates;
}
function loadScenarioPreset(){const b=buildSchedule($("scenarioVersion").value);state.scenarioSchedule=clone(b.schedule);state.scenarioDiscount=clone(b.discount);state.scenarioRates=buildRates($("scenarioVersion").value);renderAll()}function copyBase(){const b=buildSchedule($("baseVersion").value);state.scenarioSchedule=clone(b.schedule);state.scenarioDiscount=clone(b.discount);state.scenarioRates=buildRates($("baseVersion").value);$("scenarioVersion").value=$("baseVersion").value;renderAll()}
function updateWarning(){
  const proxy=text(state.selectedCatalog["대리요금적용"]).toUpperCase()==="Y";
  const dates=state.usage.map(r=>r.date).sort((a,b)=>a-b);
  const period=dates.length?`${ymd(dates[0])}~${ymd(dates.at(-1))}`:"자료 없음";
  const all=$("customerScope").value==="all";
  const messages=[];
  if(proxy){
    messages.push(`<b>환산 분석:</b> ${text(state.selectedCatalog["표시명"])} 부하에 ${text(state.selectedCatalog["요금그룹ID"])} 요금표를 대리 적용함`);
  }else{
    messages.push(`<b>데이터 범위:</b> ${text(state.selectedCatalog["표시명"])}`);
  }
  if(all){
    messages.push(`<b>전체 고객:</b> 요금제별 고객수·판매량 가중치가 없어 계약전력·전압·선택요금별 전력량요금의 단순평균을 적용함`);
  }
  messages.push(`사용량 자료기간 ${period}`);
  $("warning").innerHTML=messages.join(" · ");
}
function renderAll(){renderHours();renderScheduleCompare();renderRates();calculate()}function renderHours(){const d=$("dayType").value,a=state.scenarioSchedule?.[state.activeSeason]?.[d]||[];$("hourGrid").innerHTML=a.map((p,h)=>`<div class="hour ${PERIOD_CLASS[p]}" data-hour="${h}"><b>${String(h).padStart(2,"0")}~${String((h+1)%24).padStart(2,"0")}</b>${p}</div>`).join("");document.querySelectorAll(".hour").forEach(el=>el.onclick=()=>{const h=number(el.dataset.hour),c=state.scenarioSchedule[state.activeSeason][d][h];state.scenarioSchedule[state.activeSeason][d][h]=PERIODS[(PERIODS.indexOf(c)+1)%3];renderAll()})}
function renderScheduleCompare(){const d=$("dayType").value,b=buildSchedule($("baseVersion").value).schedule[state.activeSeason][d],s=state.scenarioSchedule[state.activeSeason][d];let h=`<div></div>${Array.from({length:24},(_,i)=>`<div class="muted" style="text-align:center">${i}</div>`).join("")}`;h+=`<div class="label">기준안</div>${b.map(p=>`<div class="mini ${PERIOD_CLASS[p]}">${p[0]}</div>`).join("")}`;h+=`<div class="label">시나리오</div>${s.map(p=>`<div class="mini ${PERIOD_CLASS[p]}">${p[0]}</div>`).join("")}`;$("scheduleCompare").innerHTML=h}
function rateTable(r,e){let h=`<thead><tr><th>계절</th>${PERIODS.map(p=>`<th>${p}</th>`).join("")}</tr></thead><tbody>`;SEASONS.forEach(s=>{h+=`<tr><td>${s==="춘추계"?"봄·가을철":s}</td>`;PERIODS.forEach(p=>h+=e?`<td><input type="number" step="0.1" data-season="${s}" data-period="${p}" value="${r[s][p].toFixed(1)}"></td>`:`<td>${r[s][p].toFixed(1)}</td>`);h+="</tr>"});return h+"</tbody>"}function renderRates(){$("baseRates").innerHTML=rateTable(buildRates($("baseVersion").value),false);$("scenarioRates").innerHTML=rateTable(state.scenarioRates,true);document.querySelectorAll("#scenarioRates input").forEach(i=>i.oninput=()=>{state.scenarioRates[i.dataset.season][i.dataset.period]=number(i.value);calculate()})}
function filteredUsage(){const y=$("year").value,s=$("scope").value,rows=state.usage.filter(r=>(y==="AVG"||r.year===number(y))&&(s==="연간"||r.season===s)),n=[...new Set(rows.map(r=>r.year))].length;return{rows,factor:y==="AVG"&&n?1/n:1}}
function calc(schedule,rates,discount,useDiscount){const{rows,factor}=filteredUsage(),bySeason={},byPeriod={};SEASONS.forEach(s=>bySeason[s]={usage:0,rev:0});PERIODS.forEach(p=>byPeriod[p]={usage:0,rev:0});let totalUsage=0,totalRev=0,days=0,minDate=null,maxDate=null,hourUsage=Array(24).fill(0);rows.forEach(r=>{days+=factor;if(!minDate||r.date<minDate)minDate=r.date;if(!maxDate||r.date>maxDate)maxDate=r.date;r.hours.forEach((raw,h)=>{const q=raw*factor,p=schedule[r.season][r.dayType][h],d=useDiscount?discount[r.season][r.dayType][h]:1,rev=q*rates[r.season][p]*d;totalUsage+=q;totalRev+=rev;hourUsage[h]+=q;bySeason[r.season].usage+=q;bySeason[r.season].rev+=rev;byPeriod[p].usage+=q;byPeriod[p].rev+=rev})});return{totalUsage,totalRev,days,minDate,maxDate,hourUsage,bySeason,byPeriod}}
function calculate(){if(!state.scenarioSchedule||!state.scenarioRates)return;const bb=buildSchedule($("baseVersion").value),br=buildRates($("baseVersion").value),base=calc(bb.schedule,br,bb.discount,true),sch=calc(state.scenarioSchedule,br,bb.discount,true),rat=calc(state.scenarioSchedule,state.scenarioRates,bb.discount,true),fin=calc(state.scenarioSchedule,state.scenarioRates,state.scenarioDiscount,$("useScenarioDiscount").checked),delta=fin.totalRev-base.totalRev;$("kUsage").textContent=fmtEnergy(fin.totalUsage);$("kPeriod").textContent=`${ymd(fin.minDate)}~${ymd(fin.maxDate)}`;$("kBase").textContent=fmtWon(base.totalRev);$("kScenario").textContent=fmtWon(fin.totalRev);$("kDelta").textContent=signedWon(delta);$("kDelta").className=cssSign(delta);$("kDeltaPct").textContent=base.totalRev?`${delta>=0?"+":""}${(delta/base.totalRev*100).toFixed(3)}%`:"-";$("kBaseAvg").textContent=base.totalUsage?(base.totalRev/base.totalUsage).toFixed(1):"-";$("kScenarioAvg").textContent=fin.totalUsage?(fin.totalRev/fin.totalUsage).toFixed(1):"-";const ds=sch.totalRev-base.totalRev,dr=rat.totalRev-sch.totalRev,dd=fin.totalRev-rat.totalRev;[["dSchedule",ds],["dRate",dr],["dDiscount",dd]].forEach(([id,v])=>{$(id).textContent=signedWon(v);$(id).className=cssSign(v)});renderSeasonTable(base,fin);renderPeriodTable(base,fin);renderChart(fin);state.lastResult={base,fin,ds,dr,dd}}
function renderSeasonTable(b,f){let h="<thead><tr><th>계절</th><th>기준안</th><th>시나리오</th><th>증감</th></tr></thead><tbody>";SEASONS.forEach(s=>{const d=f.bySeason[s].rev-b.bySeason[s].rev;h+=`<tr><td>${s==="춘추계"?"봄·가을철":s}</td><td>${fmtWon(b.bySeason[s].rev)}</td><td>${fmtWon(f.bySeason[s].rev)}</td><td class="${cssSign(d)}">${signedWon(d)}</td></tr>`});$("seasonTable").innerHTML=h+"</tbody>"}
function renderPeriodTable(b,f){let h="<thead><tr><th>시간대</th><th>기준 사용량</th><th>시나리오 사용량</th><th>변화</th></tr></thead><tbody>";PERIODS.forEach(p=>{const d=f.byPeriod[p].usage-b.byPeriod[p].usage;h+=`<tr><td>${p}</td><td>${fmtEnergy(b.byPeriod[p].usage)}</td><td>${fmtEnergy(f.byPeriod[p].usage)}</td><td class="${cssSign(d)}">${d>=0?"+":""}${fmtEnergy(d)}</td></tr>`});$("periodTable").innerHTML=h+"</tbody>"}
function getYAxisBounds(vals){
  const dataMin=Math.min(...vals);
  const dataMax=Math.max(...vals);
  const mode=$("yAxisMode").value;
  let yMin=0,yMax=Math.max(dataMax*1.08,1);
  if(mode==="zoom"){
    const span=Math.max(dataMax-dataMin,dataMax*0.02,1);
    yMin=Math.max(0,dataMin-span*0.18);
    yMax=dataMax+span*0.18;
  }else if(mode==="custom"){
    const customMin=number($("yAxisMin").value)*1e6;
    const customMax=number($("yAxisMax").value)*1e6;
    if(customMax>customMin){
      yMin=customMin;
      yMax=customMax;
    }else{
      const span=Math.max(dataMax-dataMin,dataMax*0.02,1);
      yMin=Math.max(0,dataMin-span*0.18);
      yMax=dataMax+span*0.18;
    }
  }
  if(yMax<=yMin) yMax=yMin+1;
  if(mode!=="custom"){
    $("yAxisMin").value=(yMin/1e6).toFixed(1);
    $("yAxisMax").value=(yMax/1e6).toFixed(1);
  }
  const outside=dataMin<yMin||dataMax>yMax;
  $("yAxisInfo").textContent=
    `현재 세로축 ${(yMin/1e6).toFixed(1)}~${(yMax/1e6).toFixed(1)}백만 kWh/일`
    +(outside?" · 일부 값이 직접 입력 범위를 벗어남":"");
  return {yMin,yMax};
}
function updateYAxisControlState(){
  const custom=$("yAxisMode").value==="custom";
  $("yAxisMin").disabled=!custom;
  $("yAxisMax").disabled=!custom;
}
function renderChart(r){
  const svg=$("loadChart"),w=900,h=285,l=64,rr=18,t=36,b=35;
  const vals=r.hourUsage.map(v=>r.days?v/r.days:0);
  const {yMin,yMax}=getYAxisBounds(vals);
  const x=i=>l+i*(w-l-rr)/23;
  const y=v=>h-b-(v-yMin)/(yMax-yMin)*(h-t-b);
  const day=$("dayType").value;
  let out="";
  for(let i=0;i<24;i++){
    const p=state.scenarioSchedule[state.activeSeason][day][i];
    const xx=l+i*(w-l-rr)/24,ww=(w-l-rr)/24;
    out+=`<rect x="${xx}" y="3" width="${ww-1}" height="19" fill="var(--${PERIOD_CLASS[p]})"/>`
      +`<text x="${xx+ww/2}" y="16" text-anchor="middle" font-size="8">${p[0]}</text>`;
  }
  for(let k=0;k<=4;k++){
    const ratio=k/4;
    const yy=t+(h-t-b)*ratio;
    const tick=yMax-(yMax-yMin)*ratio;
    out+=`<line x1="${l}" y1="${yy}" x2="${w-rr}" y2="${yy}" stroke="#e4e8ed"/>`
      +`<text x="${l-7}" y="${yy+4}" text-anchor="end" font-size="9" fill="#687386">${(tick/1e6).toFixed(1)}</text>`;
  }
  out+=`<text x="8" y="${t-7}" font-size="9" fill="#687386">백만 kWh/일</text>`;
  out+=`<polyline points="${vals.map((v,i)=>`${x(i)},${y(v)}`).join(" ")}" fill="none" stroke="#2468ad" stroke-width="3"/>`;
  vals.forEach((v,i)=>{
    out+=`<circle cx="${x(i)}" cy="${y(v)}" r="2.3" fill="#2468ad"/>`;
    if(i%2===0) out+=`<text x="${x(i)}" y="${h-13}" text-anchor="middle" font-size="9" fill="#687386">${i}</text>`;
  });
  svg.innerHTML=out;
}
function exportCsv(){const r=state.lastResult;if(!r)return;const rows=[["항목","값"],["종별",text(state.selectedCatalog["표시명"])],["연도",$("year").value],["범위",$("scope").value],["계약전력구간",$("contract").value],["전압구분",$("voltage").value],["선택요금",$("choice").value],["기준안 매출(원)",r.base.totalRev],["시나리오 매출(원)",r.fin.totalRev],["증감액(원)",r.fin.totalRev-r.base.totalRev],["시간대 효과(원)",r.ds],["단가 효과(원)",r.dr],["주말 할인 효과(원)",r.dd],[],["계절","기준안(원)","시나리오(원)","증감(원)"]];SEASONS.forEach(s=>rows.push([s,r.base.bySeason[s].rev,r.fin.bySeason[s].rev,r.fin.bySeason[s].rev-r.base.bySeason[s].rev]));const csv="\ufeff"+rows.map(row=>row.map(v=>`"${String(v??"").replace(/"/g,'""')}"`).join(",")).join("\r\n"),blob=new Blob([csv],{type:"text/csv;charset=utf-8"}),a=document.createElement("a");a.href=URL.createObjectURL(blob);a.download="TOU_매출_분석결과.csv";a.click();URL.revokeObjectURL(a.href)}
$("category").onchange=selectCategory;
$("customerScope").onchange=()=>{
  updateCustomerScopeControls();
  populateVersions();
  loadScenarioPreset();
  updateWarning();
};
$("contract").onchange=()=>{
  populateVoltage();
  populateVersions();
  loadScenarioPreset();
};
$("voltage").onchange=()=>{
  populateChoice();
  populateVersions();
  loadScenarioPreset();
};
$("choice").onchange=()=>{
  populateVersions();
  loadScenarioPreset();
};
$("baseVersion").onchange=renderAll;
$("scenarioVersion").onchange=loadScenarioPreset;
$("year").onchange=calculate;
$("scope").onchange=calculate;
$("dayType").onchange=renderAll;
$("useScenarioDiscount").onchange=calculate;
$("loadScenario").onclick=loadScenarioPreset;
$("copyBase").onclick=copyBase;
$("exportCsv").onclick=exportCsv;
$("yAxisMode").onchange=()=>{
  updateYAxisControlState();
  if(state.lastResult) renderChart(state.lastResult.fin);
};
$("applyYAxis").onclick=()=>{
  if(state.lastResult) renderChart(state.lastResult.fin);
};
$("yAxisMin").onchange=()=>{if(state.lastResult) renderChart(state.lastResult.fin)};
$("yAxisMax").onchange=()=>{if(state.lastResult) renderChart(state.lastResult.fin)};
document.querySelectorAll("#seasonTabs .tab").forEach(btn=>btn.onclick=()=>{
  state.activeSeason=btn.dataset.season;
  document.querySelectorAll("#seasonTabs .tab").forEach(x=>x.classList.toggle("active",x===btn));
  renderAll();
});
document.querySelectorAll(".rate-adjust").forEach(btn=>btn.onclick=()=>{
  const m=number(btn.dataset.mult);
  SEASONS.forEach(s=>PERIODS.forEach(p=>
    state.scenarioRates[s][p]=Math.round(state.scenarioRates[s][p]*m*10)/10
  ));
  renderAll();
});
updateYAxisControlState();
loadGithubData();