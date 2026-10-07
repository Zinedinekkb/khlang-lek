import * as XLSX from "xlsx";
import qrcode from "qrcode-generator";
import "./lib/claude-shim.js";
import { assetUrl } from "./lib/claude-shim.js";

window.XLSX = XLSX;
window.qrcode = qrcode;

const CATS = ["เหล็กแผ่น","เหล็กกล่อง","ท่อเหล็ก","เหล็กฉาก","เหล็กราง","เพลา","เมทัลชีท","อื่นๆ"];
const $ = s => document.querySelector(s);
const esc = s => String(s ?? "").replace(/[&<>"']/g, c => ({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}[c]));
const fmt = n => Number(n||0).toLocaleString("th-TH",{maximumFractionDigits:2});
const money = n => Number(n||0).toLocaleString("th-TH",{minimumFractionDigits:2,maximumFractionDigits:2});
const uid = () => Date.now().toString(36) + Math.random().toString(36).slice(2,7);
const dstr = t => new Date(t).toLocaleString("th-TH",{dateStyle:"short",timeStyle:"short"});
const dlong = t => new Date(t).toLocaleDateString("th-TH",{year:"numeric",month:"long",day:"numeric"});
const DEFAULT_SHOP = {name:"", phone:"", tax:"", addr:"", promptpay:"", bankName:"", bankAcc:"", bankHolder:"", billNote:"สินค้าซื้อแล้วไม่รับเปลี่ยนหรือคืน · โปรดตรวจนับสินค้าก่อนรับมอบ", ownerPin:"8888", staffPin:"1111"};

let items=[], moves=[], buys=[], docs=[], cash=[], shop={...DEFAULT_SHOP};
let cat="ทั้งหมด", docFilter="ทั้งหมด", buyFilter="ทั้งหมด", tab="stock", current=null, store, downloads=null, assets=null, sample=null, sampleImg=false;

/* ---------- storage: shared db when published, local demo otherwise ---------- */
function localStore(){
  const KEY="steel-stock-demo";
  let data={items:{},moves:{},buys:{},docs:{},cash:{},settings:{},closes:{},counts:{},prices:{},orders:{},customers:{}};
  try{ const raw=localStorage.getItem(KEY); if(raw){ const d=JSON.parse(raw);
    if(Array.isArray(d.moves)){ const m={}; d.moves.forEach(x=>m[x.id||uid()]=x); d.moves=m; }
    data={...data,...d}; } }catch(e){}
  const subs=[];
  const list=(c,o)=>{ let a=Object.entries(data[c]||{}).map(([id,v])=>({id,...v}));
    if(o.order) a.sort((x,y)=>(y[o.order]||0)-(x[o.order]||0)); if(o.limit) a=a.slice(0,o.limit); return a; };
  const save=()=>{ try{localStorage.setItem(KEY,JSON.stringify(data));}catch(e){} subs.forEach(s=>s.cb(list(s.c,s.o))); };
  const split=p=>p.split("/");
  return { mode:"local",
    sub(c,o,cb){ subs.push({c,o,cb}); cb(list(c,o)); },
    async set(p,v){ const [c,id]=split(p); (data[c]=data[c]||{})[id]=v; save(); },
    async update(p,v){ const [c,id]=split(p); Object.assign(data[c][id],v); save(); },
    async del(p){ const [c,id]=split(p); delete data[c][id]; save(); } };
}
function dbStore(db){
  return { mode:"db",
    sub(c,o,cb){ let q=db.collection(c); if(o.order) q=q.orderBy(o.order,"desc"); if(o.limit) q=q.limit(o.limit);
      q.onSnapshot(s=>cb(s.docs.map(d=>({id:d.id,...d.data()}))), onDbErr); },
    set:(p,v)=>db.doc(p).set(v), update:(p,v)=>db.doc(p).update(v), del:p=>db.doc(p).delete() };
}
function onDbErr(e){ toast("ฐานข้อมูลขัดข้อง: "+(e&&e.code||"ไม่ทราบสาเหตุ")); }
const addMove = m => store.set("moves/"+uid(), m);

async function init(){
  render();
  let db=null;
  try{ if(window.claude&&claude.use) db=await claude.use("db"); }catch(e){}
  store = db ? dbStore(db) : localStore(); wrapAudit();
  checkLocalDemoMigration();
  try{ if(window.claude&&claude.use){ userCap=await claude.use("user"); if(userCap){ try{ me=await userCap.id(); }catch(e){} } } }catch(e){}
  applyRole();
  $("#mode").innerHTML = db ? "<b>●</b> ออนไลน์" : "โหมดทดลอง"; $("#mode").title = db ? "ฐานข้อมูลออนไลน์ ซิงก์ทุกเครื่อง" : "บันทึกเฉพาะเครื่องนี้";
  store.sub("items",{},v=>{items=v;render();renderProfit();});
  store.sub("moves",{order:"at",limit:600},v=>{moves=v;renderLog();});
  store.sub("buys",{order:"at",limit:1000},v=>{buys=v;renderBuys();renderProfit();renderCash();renderAlerts();});
  store.sub("docs",{order:"at",limit:1000},v=>{docs=v;renderDocs();renderProfit();renderCash();renderAlerts();if(moneyView==="report")renderReport();});
  store.sub("closes",{order:"at",limit:120},v=>{closes=v.map(c=>({...c,date:c.date||c.id}));if(moneyView==="close"&&!$("#tab-close").hidden&&document.activeElement?.closest?.("#closeBody")==null)renderClose();});
  store.sub("prices",{order:"at",limit:30},v=>{priceHist=v;});
  store.sub("orders",{order:"at",limit:500},v=>{orders=v;renderDocs();});
  store.sub("customers",{},v=>{customers=v;fillCustList();if(docFilter==="ลูกค้า")renderCustDir();});
  store.sub("cash",{order:"at",limit:2000},v=>{cash=v;renderCash();renderProfit();});
  store.sub("settings",{},v=>{ const s=v.find(x=>x.id==="shop"); shop={...DEFAULT_SHOP,...(s||{})}; const g=v.find(x=>x.id==="scrap"); grades=g&&Array.isArray(g.grades)?g.grades:[]; $("#shopTitle").textContent = shop.name ? `คลังเหล็ก · ${shop.name}` : "คลังเหล็ก"; });
  try{ if(window.claude&&claude.use) downloads=await claude.use("downloads"); }catch(e){}
  $("#xlsBtn").hidden = !downloads; $("#vSave").hidden = !downloads;
  try{ assets=await claude.use("assets"); }catch(e){}
  try{ sample=await claude.use("sample"); const lim=sample&&await sample.limits().catch(()=>null); sampleImg=!!sample; }catch(e){}
}

/* ---------- render ---------- */
const isLow = i => (+i.min||0)>0 && (+i.qty||0)<=(+i.min||0);
function render(){
  $("#chips").innerHTML=["ทั้งหมด",...CATS].map(c=>`<button aria-pressed="${c===cat}" data-cat="${esc(c)}">${esc(c)}</button>`).join("");
  const q=$("#q").value.trim().toLowerCase();
  const shown=items.filter(i=>cat==="ทั้งหมด"||i.cat===cat)
    .filter(i=>!q||[i.name,i.spec,i.loc].join(" ").toLowerCase().includes(q))
    .sort((a,b)=>(isLow(b)-isLow(a))||String(a.name).localeCompare(String(b.name),"th"));
  $("#list").innerHTML = shown.length ? shown.map(i=>`
    <button class="item ${isLow(i)?"low":""}" data-id="${esc(i.id)}">
      <h3>${esc(i.name)}</h3>
      <div class="qty"><span class="num">${fmt(i.qty)}</span><small>${esc(i.unit)}${remOf(i).length?" เต็ม":""}</small></div>
      <div class="meta"><span class="tag ${i.cond==="used"?"used":""}">${i.cond==="used"?"มือสอง":"ใหม่"}</span>${esc(i.spec)}${i.loc?" · "+esc(i.loc):""}${+i.kg?` · ≈${fmt(Math.round(itemKg(i)*100)/100)} กก.`:""}</div>
      ${remOf(i).length?`<div class="rem"><b>✂</b>${esc(remSummary(i))}</div>`:""}
      <div class="money">${+i.price>0
        ? `<span>ขาย <b class="num">${money(i.price)}</b>/${esc(i.unit)}${+i.kg>0&&i.unit!=="กก."?` <small>(${money(i.price/i.kg)}/กก.)</small>`:""}</span><span>มูลค่า <b class="num">${money(itemVal(i))}</b> ฿</span>`
        : `<span class="noprice">ยังไม่ตั้งราคาขาย</span>`}</div>
    </button>`).join("")
    : `<div class="empty">${items.length?"ไม่พบสินค้าที่ตรงกับการค้นหา":"ยังไม่มีสินค้าในคลัง<br><button class='btn primary' id='emptyAdd'>เพิ่มสินค้าชิ้นแรก</button>"}</div>`;
  $("#sCount").textContent=fmt(items.length);
  $("#sValue").textContent=fmt(Math.round(items.reduce((s,i)=>s+itemVal(i),0)));
  $("#sKg").textContent=fmt(Math.round(items.reduce((s,i)=>s+itemKg(i),0)));
  const low=items.filter(isLow).length; $("#sLow").textContent=low; $("#sLowBox").classList.toggle("alert",low>0);
}
function renderLog(){
  $("#log").innerHTML = moves.length ? moves.slice(0,150).map(m=>`
    <div><span class="num ${m.type}">${m.label?esc(m.label):(m.type==="in"?"+":"−")+fmt(m.qty)}</span>
    <span>${esc(m.name)}${m.note||m.by?`<br><small>${esc(m.note||"")}${byTxt(m)}</small>`:""}</span>
    <small>${dstr(m.at)}</small></div>`).join("") : `<p class="empty">ยังไม่มีการรับเข้าหรือจ่ายออก</p>`;
}
function renderBuys(){
  const nOwe=buys.filter(b=>buyOwed(b)>0).length;
  $("#buyChips").innerHTML=["ทั้งหมด","เหล็กเก่า","เหล็กใหม่","ค้างจ่าย"].map(c=>`<button aria-pressed="${c===buyFilter}" data-bf="${c}">${c}${c==="ค้างจ่าย"&&nOwe?` (${nOwe})`:""}</button>`).join("");
  const now=new Date(), mStart=new Date(now.getFullYear(),now.getMonth(),1).getTime();
  const kindOk=b=>buyFilter==="ทั้งหมด"||(buyFilter==="ค้างจ่าย"?buyOwed(b)>0:(buyFilter==="เหล็กเก่า")===((b.kind||"used")==="used"));
  const shown=buys.filter(kindOk), mb=shown.filter(b=>(b.date||b.at)>=mStart);
  $("#buySum").innerHTML=`<span>เดือนนี้ซื้อเข้า<br><b class="num">${fmt(mb.reduce((s,b)=>s+(+b.total||0),0))}</b> บาท</span>
    <span>น้ำหนัก (ที่ระบุเป็น กก./ตัน)<br><b class="num">${fmt(mb.reduce((s,b)=>s+(+b.kg||0),0))}</b> กก.</span>
    <span>จำนวนบิล<br><b class="num">${mb.length}</b></span>`;
  $("#buyList").innerHTML = shown.length ? shown.map(b=>{ const ls=legacyLines(b), used=(b.kind||"used")==="used";
    return `<button class="rec" data-buy="${esc(b.id)}" style="border-left-color:${used?"var(--hazard)":"var(--steel)"}">
      <h3><span class="tag ${used?"used":"new"}">${used?"เก่า":"ใหม่"}</span>${esc(b.seller)}${b.photoId?" 📷":""} ${buyTag(b)}</h3>
      <div class="amt"><span class="num">${fmt(b.total)}</span><small>บาท</small></div>
      <div class="meta">${esc(ls.map(l=>`${l.desc} ${fmt(l.qty)}${l.unit}`).join(", "))}<br>${dlong(b.date||b.at)}${b.billNo?" · บิล "+esc(b.billNo):""}</div>
    </button>`; }).join("") : `<div class="empty">ยังไม่มีรายการซื้อเข้า<br>กด "+ ซื้อเข้า" แล้วถ่ายรูปใบเสร็จ ระบบจะกรอกรายการให้</div>`;
}
const DOCNAME={quote:"ใบเสนอราคา",receipt:"ใบเสร็จรับเงิน",credit:"ใบส่งของ (ขายเชื่อ)"};
function renderDocs(){
  const opts=["ทั้งหมด","ออเดอร์","รอจัดส่ง","ลูกค้า","ลูกหนี้","ใบเสร็จรับเงิน","ขายเชื่อ","ใบเสนอราคา"];
  const cnt={"ลูกหนี้":debtors().length,"ออเดอร์":openOrders().length,"รอจัดส่ง":toShip().length};
  $("#docChips").innerHTML=opts.map(c=>`<button aria-pressed="${c===docFilter}" data-df="${c}">${c}${cnt[c]?` (${cnt[c]})`:""}</button>`).join("");
  if(docFilter==="ลูกค้า"){
    if(!$("#cuSearch")) $("#docList").innerHTML=`<input id="cuSearch" type="search" placeholder="ค้นหาชื่อหรือเบอร์ลูกค้า" style="width:100%;padding:10px 12px;border:1px solid var(--line);border-radius:6px;background:var(--panel);margin-bottom:8px"><div id="cuListBox"></div>`;
    $("#cuSearch").value=custSearch; renderCustDir(); return;
  }
  if(docFilter==="รอจัดส่ง"){
    const s=toShip();
    $("#docList").innerHTML=s.length?s.map(d=>`<button class="rec ${d.type}" data-doc="${esc(d.id)}"><h3>${esc(d.customer)} ${fulTag(d)}</h3>
      <div class="amt"><span class="num">${fmt(d.total)}</span><small>บาท</small></div>
      <div class="meta">📍 ${esc(d.shipAddr||"ไม่ระบุที่อยู่")}${d.shipPhone?" · ☎ "+esc(d.shipPhone):""}<br>${DOCNAME[d.type]} ${esc(d.no)}${d.shipDate&&daysFrom(d.shipDate)>0?` · <b class="neg">เลยวันส่ง ${daysFrom(d.shipDate)} วัน</b>`:""}</div></button>`).join("")
      :`<div class="empty">ไม่มีของรอจัดส่ง</div>`;
    return;
  }
  if(docFilter==="ออเดอร์"){
    const op=openOrders().sort((x,y)=>(x.need||9e15)-(y.need||9e15)), past=orders.filter(o=>["done","cancel"].includes(o.status)).slice(0,10);
    $("#docList").innerHTML=renderOrders(op)+(past.length?`<h4 class="sec">ปิดแล้วล่าสุด</h4>`+renderOrders(past):"");
    return;
  }
  if(docFilter==="ลูกหนี้"){
    const ds=debtors(), tot=ds.reduce((s,x)=>s+x.owed,0);
    $("#docList").innerHTML = ds.length ? `<div class="sumbar"><span>ลูกหนี้ทั้งหมด<br><b class="num">${fmt(tot)}</b> บาท</span><span>จำนวนราย<br><b class="num">${ds.length}</b></span>
      <span>เกินกำหนด<br><b class="num neg">${fmt(ds.filter(x=>daysFrom(x.oldestDue)>0).reduce((s,x)=>s+x.owed,0))}</b></span></div>`+
      ds.map(x=>{ const od=daysFrom(x.oldestDue); return `<button class="rec" data-cust="${esc(x.name)}" style="border-left-color:${od>0?"var(--low)":"var(--hazard)"}">
        <h3>${esc(x.name)}</h3><div class="amt"><span class="num ${od>0?"neg":""}">${fmt(x.owed)}</span><small>บาท</small></div>
        <div class="meta">${x.n} ใบค้าง${x.phone?" · "+esc(x.phone):""} · ${od>0?`<b class="neg">เกินกำหนด ${od} วัน</b>`:`ครบกำหนด ${dlong(x.oldestDue)}`}</div></button>`; }).join("")
      : `<div class="empty">ไม่มีลูกหนี้ค้างชำระ 👍</div>`;
    return;
  }
  const want={"ใบเสร็จรับเงิน":"receipt","ขายเชื่อ":"credit","ใบเสนอราคา":"quote"}[docFilter];
  const shown=docs.filter(d=>!want||d.type===want);
  $("#docList").innerHTML = shown.length ? shown.map(d=>`
    <button class="rec ${d.type}" data-doc="${esc(d.id)}">
      <h3>${esc(d.customer)} ${docStatusTag(d)} ${d.ful==="ship"&&d.type!=="quote"?fulTag(d):""}</h3>
      <div class="amt"><span class="num">${fmt(d.total)}</span><small>บาท</small></div>
      <div class="meta">${DOCNAME[d.type]} ${esc(d.no)} · ${dlong(d.date||d.at)}${d.via?" · "+esc(d.via):""}${d.cut?" · ตัดสต็อกแล้ว":""}${byTxt(d)}</div>
    </button>`).join("") : `<div class="empty">ยังไม่มีเอกสาร<br>กด "⚡ ขายด่วน" หรือ "+ ออกเอกสาร"</div>`;
}
/* ---------- tabs & fab ---------- */
const FAB={stock:"+ เพิ่มสินค้า",buy:"+ ซื้อเข้า",doc:"⚡ ขายด่วน",money:"+ รายรับ/รายจ่าย",more:""};
const TITLE={stock:"สต็อก",buy:"ซื้อเข้า",doc:"ใบเสนอราคา / ใบเสร็จ",money:"การเงิน",more:"เพิ่มเติม"};
let moneyView="cash";
function setFab(){ const t= tab==="money"&&moneyView!=="cash" ? "" : FAB[tab]; $("#fab").hidden=!t; $("#fab").textContent=t; }
document.querySelectorAll(".tabbar button").forEach(b=>b.addEventListener("click",()=>{
  tab=b.dataset.tab;
  document.querySelectorAll(".tabbar button").forEach(x=>x.setAttribute("aria-selected",x===b));
  ["stock","buy","doc","money","more"].forEach(t=>$("#tab-"+t).hidden=t!==tab);
  $("#pageTitle").textContent=TITLE[tab]; $("#main").scrollTop=0; setFab();
}));
document.querySelectorAll(".segbar button").forEach(b=>b.addEventListener("click",()=>{
  moneyView=b.dataset.mv;
  document.querySelectorAll(".segbar button").forEach(x=>x.setAttribute("aria-selected",x===b));
  ["cash","profit","report","close"].forEach(k=>$("#tab-"+k).hidden=moneyView!==k); setFab();
  if(moneyView==="report") renderReport(); if(moneyView==="close") renderClose();
}));
$("#fab").addEventListener("click",()=>{ if(tab==="stock")openEdit(null); else if(tab==="buy")openBuy(null); else if(tab==="doc")openQuick(); else if(tab==="money")openCash(null,"out"); });
$("#q").addEventListener("input",render);
$("#chips").addEventListener("click",e=>{const b=e.target.closest("[data-cat]"); if(b){cat=b.dataset.cat;render();}});
$("#buyChips").addEventListener("click",e=>{const b=e.target.closest("[data-bf]"); if(b){buyFilter=b.dataset.bf;renderBuys();}});
$("#docChips").addEventListener("click",e=>{const b=e.target.closest("[data-df]"); if(b){docFilter=b.dataset.df;renderDocs();}});
$("#list").addEventListener("click",e=>{ if(e.target.id==="emptyAdd") return openEdit(null);
  const b=e.target.closest("[data-id]"); if(b) openMove(b.dataset.id); });
$("#buyList").addEventListener("click",e=>{const b=e.target.closest("[data-buy]"); if(b) openBuy(buys.find(x=>x.id===b.dataset.buy));});
$("#docList").addEventListener("click",e=>{const o=e.target.closest("[data-ord]"); if(o) return openOrder(orders.find(x=>x.id===o.dataset.ord)); const c=e.target.closest("[data-cust]"); if(c) return openCust(c.dataset.cust); const b=e.target.closest("[data-doc]"); if(b) openView(docs.find(x=>x.id===b.dataset.doc));});

/* ---------- stock move ---------- */
function openMove(id){
  current=items.find(i=>i.id===id); if(!current) return;
  $("#mName").textContent=current.name;
  $("#mMeta").textContent=[current.cat,current.cond==="used"?"มือสอง":"ใหม่",current.spec,current.loc].filter(Boolean).join(" · ");
  $("#mQty").textContent=fmt(current.qty); $("#mUnit").textContent=current.unit;
  $("#mKg").textContent= +current.kg ? `(รวมท่อนเหลือ ≈ ${fmt(Math.round(itemKg(current)*100)/100)} กก.)` : "";
  $("#mAmt").value=1; $("#mNote").value=""; drawRem();
  $("#dMove").showModal();
}
document.querySelectorAll("[data-step]").forEach(b=>b.addEventListener("click",()=>{ $("#mAmt").value=Math.max(0,(+$("#mAmt").value||0)+(+b.dataset.step)); }));
$("#mClose").addEventListener("click",()=>$("#dMove").close());
$("#mIn").addEventListener("click",()=>move("in"));
$("#mOut").addEventListener("click",()=>move("out"));
$("#mEdit").addEventListener("click",()=>{ $("#dMove").close(); openEdit(current); });
async function move(type){
  const n=+$("#mAmt").value; if(!(n>0)) return toast("ใส่จำนวนมากกว่า 0");
  const cur=+current.qty||0;
  if(type==="out"&&n>cur) return toast(`ของไม่พอ เหลือ ${fmt(cur)} ${current.unit}`);
  try{
    await store.update("items/"+current.id,{qty:type==="in"?cur+n:cur-n,updatedAt:Date.now()});
    await addMove({type,qty:n,itemId:current.id,name:current.name,note:$("#mNote").value.trim(),at:Date.now()});
    $("#dMove").close(); toast(type==="in"?`รับเข้า ${fmt(n)} ${current.unit}`:`จ่ายออก ${fmt(n)} ${current.unit}`);
  }catch(e){ toast("บันทึกไม่สำเร็จ ลองอีกครั้ง"); }
}

/* ---------- item edit ---------- */
const f=$("#fEdit"); f.cat.innerHTML=CATS.map(c=>`<option>${c}</option>`).join("");
let editing=null;
function openEdit(item){
  editing=item; $("#eTitle").textContent=item?"แก้ไขสินค้า":"เพิ่มสินค้า"; $("#eDel").hidden=!item;
  const d=item||{name:"",cat:CATS[0],cond:"new",spec:"",qty:0,unit:"เส้น",kg:0,min:0,cost:0,price:0,loc:""};
  for(const k of ["name","cat","cond","spec","qty","unit","kg","min","cost","price","loc"]) f[k].value=d[k]??"";
  f.lenFull.value=d.lenFull||""; f.sheetW.value=d.sheetW||""; f.sheetL.value=d.sheetL||""; editCutRows();
  editMargin();
  $("#dEdit").showModal();
}
f.addEventListener("submit",async e=>{
  if(e.submitter&&e.submitter.value!=="save") return;
  e.preventDefault(); if(!f.reportValidity()) return;
  const v={name:f.name.value.trim(),cat:f.cat.value,cond:f.cond.value,spec:f.spec.value.trim(),qty:+f.qty.value||0,unit:f.unit.value,
    kg:+f.kg.value||0,min:+f.min.value||0,cost:+f.cost.value||0,price:+f.price.value||0,loc:f.loc.value.trim(),updatedAt:Date.now(),
    lenFull:+f.lenFull.value||0,sheetW:+f.sheetW.value||0,sheetL:+f.sheetL.value||0,rem:editing?remOf(editing):[]};
  try{ await store.set("items/"+(editing?editing.id:uid()),v); $("#dEdit").close(); toast(editing?"บันทึกการแก้ไขแล้ว":"เพิ่มสินค้าแล้ว"); }
  catch(err){ toast("บันทึกไม่สำเร็จ ลองอีกครั้ง"); }
});
$("#eDel").addEventListener("click",async()=>{
  if(!editing||!confirm(`ลบ "${editing.name}" ออกจากคลัง?`)) return;
  try{ await store.del("items/"+editing.id); $("#dEdit").close(); toast("ลบสินค้าแล้ว"); }catch(e){ toast("ลบไม่สำเร็จ"); }
});
$("#eCalc").addEventListener("click",()=>openCalc(true));
function editCutRows(){ const m=cutMode({cat:f.cat.value,unit:f.unit.value}); $("#eLenRow").hidden=m!=="len"; $("#eSheetRow").hidden=m!=="area"; }
f.cat.addEventListener("change",editCutRows); f.unit.addEventListener("change",editCutRows);

/* ---------- weight calculator (kg) ---------- */
const LEN_UNITS={"มม.":1,"ซม.":10,"ม.":1000,"นิ้ว":25.4,"หุน":3.175,"ฟุต":304.8,"หลา":914.4};
// [key, label, default value, default unit, base unit (mm|m)]
const SHAPES={
  plate:{n:"เหล็กแผ่น",f:[["t","หนา",3,"มม.","mm"],["w","กว้าง",4,"ฟุต","mm"],["l","ยาว",8,"ฟุต","mm"]],kg:v=>v.t*v.w*v.l*7.85e-6},
  box:{n:"เหล็กกล่อง / แป๊บเหลี่ยม",f:[["a","ด้าน A",2,"นิ้ว","mm"],["b","ด้าน B",4,"นิ้ว","mm"],["t","หนา",2.3,"มม.","mm"],["len","ยาว",6,"ม.","m"]],kg:v=>Math.max(0,2*v.t*(v.a+v.b)-4*v.t*v.t)*0.00785*v.len},
  pipe:{n:"ท่อกลม",f:[["d","โตนอก",60.5,"มม.","mm"],["t","หนา",3.2,"มม.","mm"],["len","ยาว",6,"ม.","m"]],kg:v=>0.02466*v.t*Math.max(0,v.d-v.t)*v.len},
  shaft:{n:"เพลากลม",f:[["d","โต",1,"นิ้ว","mm"],["len","ยาว",6,"ม.","m"]],kg:v=>0.006165*v.d*v.d*v.len},
  sq:{n:"เหล็กสี่เหลี่ยมตัน",f:[["a","ด้าน",12,"มม.","mm"],["len","ยาว",6,"ม.","m"]],kg:v=>v.a*v.a*0.00785*v.len},
  angle:{n:"เหล็กฉาก",f:[["a","ขา",2,"นิ้ว","mm"],["t","หนา",5,"มม.","mm"],["len","ยาว",6,"ม.","m"]],kg:v=>Math.max(0,v.t*(2*v.a-v.t))*0.00785*v.len},
  flat:{n:"เหล็กแบน",f:[["w","กว้าง",2,"นิ้ว","mm"],["t","หนา",6,"มม.","mm"],["len","ยาว",6,"ม.","m"]],kg:v=>v.w*v.t*0.00785*v.len},
  chan:{n:"เหล็กราง U / รางน้ำ",f:[["h","สูง (หลัง)",100,"มม.","mm"],["b","ปีก",50,"มม.","mm"],["t","หนา",5,"มม.","mm"],["len","ยาว",6,"ม.","m"]],kg:v=>Math.max(0,v.t*(v.h+2*v.b-2*v.t))*0.00785*v.len},
  clip:{n:"เหล็กตัวซี (ราง C มีขอบ)",f:[["h","สูง (หลัง)",100,"มม.","mm"],["b","ปีก",50,"มม.","mm"],["c","ขอบพับ",20,"มม.","mm"],["t","หนา",2.3,"มม.","mm"],["len","ยาว",6,"ม.","m"]],kg:v=>Math.max(0,v.t*(v.h+2*v.b+2*v.c-4*v.t))*0.00785*v.len},
  metal:{n:"เมทัลชีท / แผ่นลอน",f:[["t","หนา (รวมเคลือบ)",0.35,"มม.","mm"],["w","หน้ากว้างก่อนรีดลอน",914,"มม.","mm"],["c","หน้ากว้างใช้งาน (มุงได้)",760,"มม.","mm"],["len","ยาว",6,"ม.","m"]],kg:v=>v.t*v.w*0.00785*v.len},
  rebar:{n:"เหล็กเส้น (กลม/ข้ออ้อย)",f:[["d","โต",12,"มม.","mm"],["len","ยาว",10,"ม.","m"]],kg:v=>0.006165*v.d*v.d*v.len}
};
const dimState={}; // remembers value+unit per shape/field
$("#cShape").innerHTML=Object.entries(SHAPES).map(([k,s])=>`<option value="${k}">${s.n}</option>`).join("");
function drawDims(){
  const sk=$("#cShape").value, s=SHAPES[sk];
  $("#cDims").innerHTML=s.f.map(([k,l,dv,du])=>{ const st=dimState[sk+k]||{v:dv,u:du};
    return `<label>${l}<span class="dim"><input data-dim="${k}" type="number" inputmode="decimal" min="0" step="any" value="${st.v}">
      <select data-unit="${k}" aria-label="หน่วยของ${l}">${Object.keys(LEN_UNITS).map(u=>`<option ${u===st.u?"selected":""}>${u}</option>`).join("")}</select></span></label>`; }).join("");
  calc();
}
let calcPer=0;
function calc(){
  const sk=$("#cShape").value, s=SHAPES[sk], v={};
  let lenM=0, areaM2=0;
  s.f.forEach(([k,,,,base])=>{
    const val=+($(`#cDims [data-dim="${k}"]`)?.value)||0, u=$(`#cDims [data-unit="${k}"]`)?.value||"มม.";
    dimState[sk+k]={v:val,u}; const mm=val*LEN_UNITS[u]; v[k]= base==="m" ? mm/1000 : mm;
  });
  calcPer=s.kg(v)||0;
  const n=+$("#cN").value||0, all=calcPer*n, p=+$("#cPrice").value||0, out=+$("#cOut").value||1, uname=$("#cOut").selectedOptions[0].text.replace(/ \(.+\)/,"").replace("กิโลกรัม","กก.");
  document.querySelectorAll("#dCalc .cU").forEach(x=>x.textContent=uname);
  const dec = out>=1000 ? 4 : 2, f2=x=>Number(x).toLocaleString("th-TH",{maximumFractionDigits:dec});
  $("#cOne").textContent=f2(calcPer/out); $("#cAll").textContent=f2(all/out);
  if(sk==="plate"){ areaM2=v.w*v.l/1e6*n; $("#cExtra").textContent=`พื้นที่ ${fmt(areaM2)} ตร.ม.`; }
  else if(sk==="metal"){ lenM=(v.len||0)*n; areaM2=(v.c||0)/1000*lenM; const kgm=v.len?calcPer/v.len:0;
    $("#cExtra").textContent=`${fmt(Math.round(kgm*1000)/1000)} กก./ม. · ยาวรวม ${fmt(lenM)} ม. · มุงได้ ${fmt(Math.round(areaM2*100)/100)} ตร.ม.`; }
  else { lenM=(v.len||0)*n; $("#cExtra").textContent=`${fmt(lenM)} ม. (${fmt(lenM/0.3048)} ฟุต)`; }
  $("#cMoneyRow").hidden=!p; $("#cMoney").textContent=fmt(all*p);
}
$("#cShape").addEventListener("change",drawDims);
$("#dCalc").addEventListener("input",e=>{ if(e.target.id!=="cShape") calc(); });
$("#dCalc").addEventListener("change",e=>{ if(e.target.dataset.unit!==undefined||e.target.id==="cOut") calc(); });
function openCalc(fill){
  $("#cUse").hidden=!fill;
  const map={"เหล็กแผ่น":"plate","เมทัลชีท":"metal","เหล็กกล่อง":"box","ท่อเหล็ก":"pipe","เพลา":"shaft","เหล็กฉาก":"angle","เหล็กราง":"chan"};
  if(fill&&map[f.cat.value]&&$("#cShape").value!==map[f.cat.value]){ $("#cShape").value=map[f.cat.value]; drawDims(); }
  calc(); $("#dCalc").showModal();
}
$("#calcBtn").addEventListener("click",()=>openCalc(false));
$("#cClose").addEventListener("click",()=>$("#dCalc").close());
$("#cUse").addEventListener("click",()=>{ f.kg.value=Math.round(calcPer*100)/100; $("#dCalc").close(); });
drawDims();

/* ---------- photo → fields, red warnings for unknown values ---------- */
function setMsg(id,kind,html){ const el=$(id); el.hidden=!html; el.className="msg "+(kind||""); el.innerHTML=html||""; }
function warnInput(inp,msg){
  inp.classList.add("warn"); inp.title=msg;
  const lab=inp.closest("label"); if(!lab) return;
  let s=lab.querySelector(".wmsg"); if(!s){ s=document.createElement("small"); s.className="wmsg"; lab.appendChild(s); } s.textContent=msg;
}
function clearWarn(inp){ inp.classList.remove("warn"); inp.removeAttribute("title"); inp.closest("label")?.querySelector(".wmsg")?.remove(); }
function clearAllWarns(form){ form.querySelectorAll(".warn").forEach(clearWarn); }
const ocrState={fBuy:{msg:"#bMsg",mismatch:"",read:false},fDoc:{msg:"#dMsg",mismatch:"",read:false}};
function refreshBanner(form){
  const st=ocrState[form.id]; if(!st.read) return;
  const n=form.querySelectorAll(".warn").length;
  if(st.total!=null){ const cur= form.id==="fBuy" ? bSum() : Math.max(0,sumLines(lines)-(+fd.discount.value||0));
    st.mismatch = !st.total ? (form.id==="fBuy"?"ไม่พบยอดรวมในบิล ตรวจยอดเองอีกครั้ง":"") : Math.abs(st.total-cur)>1 ? `ยอดรวมในรูป ${money(st.total)} บาท ไม่ตรงกับที่คำนวณได้ ${money(cur)} บาท` : ""; }
  if(n) setMsg(st.msg,"bad",`⚠ มี ${n} ช่องสีแดงที่อ่านไม่ได้หรือไม่แน่ใจ — กรอกหรือตรวจก่อนบันทึก${st.mismatch?"<br>⚠ "+st.mismatch:""}`);
  else if(st.mismatch) setMsg(st.msg,"bad","⚠ "+st.mismatch);
  else setMsg(st.msg,"ok","✓ กรอกจากรูปครบแล้ว ตรวจตัวเลขอีกครั้งก่อนบันทึก");
}
["fBuy","fDoc"].forEach(id=>$("#"+id).addEventListener("input",e=>{ if(e.target.classList.contains("warn")) clearWarn(e.target); refreshBanner($("#"+id)); }));
const blankVal=(k,v)=> v==null || (typeof v==="string"&&!v.trim()) || (["qty","price"].includes(k)&&!(+v>0));
function ocrError(err,btn,msgId){
  const c=err&&err.code;
  if(c==="not_granted"){ sampleImg=false; btn.hidden=true; setMsg(msgId,"info","ไม่ได้อนุญาตให้อ่านรูป รูปจะถูกแนบไว้ กรอกข้อมูลเองได้ตามปกติ"); }
  else setMsg(msgId,"bad", c==="rate_limited" ? "อ่านถี่เกินไป รอสักครู่แล้วกด \"อ่านรูปอีกครั้ง\"" : "อ่านรูปไม่สำเร็จ ลองถ่ายให้ชัด ตรง และมีแสงพอ แล้วกด \"อ่านรูปอีกครั้ง\" หรือกรอกเอง");
}
// apply header fields: list of [input, value, key, kind]
function fillHead(form,pairs,unc){
  pairs.forEach(([k,val,isDate])=>{
    const inp=form[k], ok=isDate ? /^\d{4}-\d{2}-\d{2}$/.test(val||"") : !blankVal(k,val);
    if(ok){ inp.value=val; if(unc.has(k)) warnInput(inp,"อ่านไม่ชัด ตรวจอีกครั้ง"); }
    else warnInput(inp, isDate ? "ไม่พบวันที่ในรูป ใช้วันนี้แทน" : "ไม่พบในรูป กรอกเอง");
  });
}
const lineWarns=(l,i,keys,unc)=>{ const w={}; keys.forEach(([k,src])=>{ if(blankVal(k,l[src]??l[k])) w[k]="ไม่พบในรูป"; else if(unc.has(`lines[${i}].${src||k}`)) w[k]="อ่านไม่ชัด"; }); return w; };
const wc=(l,k)=> l._w&&l._w[k] ? ` warn" title="${esc(l._w[k])}` : "";
const UNC_SPEC=`"uncertain": array of field paths you read but are NOT confident about (blurry, handwriting, guessed), e.g. ["seller","date","lines[0].price","lines[2].qty"]. Use null for any field that is not present or unreadable — never invent values.`;

/* ---------- purchases: used & new steel, receipt photo + auto-read ---------- */
const fb=$("#fBuy"); let editingBuy=null, blines=[];
const UNITS=["กก.","เส้น","แผ่น","ท่อน","ชิ้น","ตัน"];
function legacyLines(b){
  if(Array.isArray(b.lines)) return b.lines.map(l=>({...l}));
  return [{desc:b.desc||"",qty:+b.kg||0,unit:"กก.",price:+b.ppk||0,itemId:b.itemId||"",addQty:+b.addQty||0}];
}
const blankBLine=()=>({desc:"",qty:0,unit:"กก.",price:0,itemId:"",addQty:0});
function openBuy(b){
  editingBuy=b||null; pendingPhotos.forEach(p=>URL.revokeObjectURL(p.url)); pendingPhotos=[]; drawThumbs();
  $("#bTitle").textContent=b?"รายการซื้อเข้า":"บันทึกการซื้อเข้า"; $("#bDel").hidden=!b;
  const kind=b?.kind||"used"; fb.querySelector(`[name=kind][value=${kind}]`).checked=true;
  fb.seller.value=b?.seller||""; fb.phone.value=b?.phone||""; fb.billNo.value=b?.billNo||"";
  fb.date.value=b?.date?isoDate(b.date):today(); fb.note.value=b?.note||"";
  blines=b?legacyLines(b):[blankBLine()];
  $("#bImg").hidden=true;
  const ids=b?(b.photoIds||(b.photoId?[b.photoId]:[])):[];
  ids.forEach((id,i)=>{ const im=document.createElement("img"); im.src=assetUrl(id); im.alt="หน้า "+(i+1); im.onclick=()=>{ $("#iBig").src=im.src; $("#dImg").showModal(); }; $("#bThumbs").appendChild(im); });
  clearAllWarns(fb); ocrState.fBuy={msg:"#bMsg",mismatch:"",read:false}; setMsg("#bMsg");
  $("#bPick").hidden=!!b; $("#bPick").firstChild.textContent="ถ่าย/เลือกรูปใบเสร็จ"; $("#bPick").disabled=false;
  showManual(!!b); $("#bManualBtn").hidden=!!b;
  fb.payStatus.value=b?.payStatus||"paid"; fb.payVia.value=b?.payVia||"เงินสด"; fb.due.value=b?.due?isoDate(b.due):""; payFields(fb); buyPayInfo(b); drawGradeChips();
  $("#bRead").hidden=true; drawBLines(); $("#dBuy").showModal();
}
const isoDate=t=>{ const d=new Date(t); return new Date(d-d.getTimezoneOffset()*6e4).toISOString().slice(0,10); };
function stockOptions(sel){
  return `<option value="">— ไม่เข้าสต็อก —</option>`+[...items].sort((a,c)=>String(a.name).localeCompare(String(c.name),"th"))
    .map(i=>`<option value="${esc(i.id)}" ${i.id===sel?"selected":""}>${esc(i.name)} ${esc(i.spec)} (${esc(i.unit)})</option>`).join("");
}
function drawBLines(){
  const locked=!!editingBuy;
  $("#bLines").innerHTML=blines.map((l,i)=>`
    <div class="pline" data-i="${i}">
      <input class="nm${wc(l,"desc")}" data-k="desc" value="${esc(l.desc)}" placeholder="รายการ เช่น เหล็กกล่อง 2x4 หนา 2.3 / เศษเหล็กหนา" maxlength="100">
      <input class="${wc(l,"qty")}" data-k="qty" type="number" inputmode="decimal" min="0" step="any" value="${l.qty||""}" placeholder="จำนวน" aria-label="จำนวน">
      <input class="${wc(l,"unit")}" data-k="unit" list="unitList" value="${esc(l.unit)}" placeholder="หน่วย" maxlength="10">
      <input class="${wc(l,"price")}" data-k="price" type="number" inputmode="decimal" min="0" step="any" value="${l.price||""}" placeholder="ราคา/หน่วย" aria-label="ราคาต่อหน่วย">
      <button type="button" class="rm" data-rm="${i}" aria-label="ลบรายการ" ${locked?"hidden":""}>×</button>
      <div class="stk"><select data-k="itemId" ${locked?"disabled":""} aria-label="เข้าสต็อกสินค้า">${stockOptions(l.itemId)}</select>
        <input data-k="addQty" type="number" inputmode="decimal" min="0" step="any" value="${l.addQty||""}" placeholder="จำนวนเข้าสต็อก" ${locked?"disabled":""} aria-label="จำนวนเข้าสต็อก"></div>
      <div class="sub">รวม <span class="num">${money((+l.qty||0)*(+l.price||0))}</span> บาท</div>
    </div>`).join("")+`<datalist id="unitList">${UNITS.map(u=>`<option>${u}</option>`).join("")}</datalist>`;
  buyTotal();
}
const bSum=()=>blines.reduce((s,l)=>s+(+l.qty||0)*(+l.price||0),0);
function buyTotal(){ $("#bTotal").textContent=money(bSum()); }
$("#bLines").addEventListener("input",e=>{
  const row=e.target.closest("[data-i]"); if(!row) return; const l=blines[+row.dataset.i], k=e.target.dataset.k; if(!k) return;
  l[k]=["qty","price","addQty"].includes(k)?(+e.target.value||0):e.target.value; if(l._w) delete l._w[k];
  if(k==="itemId"){ const it=items.find(x=>x.id===l.itemId); if(it&&!l.addQty){ l.addQty = it.unit===l.unit ? l.qty : 0; row.querySelector("[data-k=addQty]").value=l.addQty||""; } }
  row.querySelector(".sub .num").textContent=money((+l.qty||0)*(+l.price||0)); buyTotal();
});
$("#bLines").addEventListener("click",e=>{ const r=e.target.dataset.rm; if(r!==undefined){ blines.splice(+r,1); if(!blines.length) blines.push(blankBLine()); drawBLines(); } });
$("#bAdd").addEventListener("click",()=>{ blines.push(blankBLine()); drawBLines(); });
$("#bPick").addEventListener("click",()=>$("#bFile").click());
$("#bImg").addEventListener("click",()=>{ $("#iBig").src=$("#bImg").src; $("#dImg").showModal(); });
$("#iClose").addEventListener("click",()=>$("#dImg").close());

function shrink(file,max=1600){
  return new Promise((res,rej)=>{
    const img=new Image(), url=URL.createObjectURL(file);
    img.onload=()=>{ const r=Math.min(1,max/Math.max(img.width,img.height)); const c=document.createElement("canvas");
      c.width=Math.round(img.width*r); c.height=Math.round(img.height*r); c.getContext("2d").drawImage(img,0,0,c.width,c.height);
      URL.revokeObjectURL(url); c.toBlob(b=>b?res(b):rej(new Error("encode")),"image/jpeg",0.85); };
    img.onerror=()=>{ URL.revokeObjectURL(url); rej(new Error("decode")); }; img.src=url;
  });
}
let pendingPhotos=[];
function drawThumbs(){
  const t=$("#bThumbs"); t.innerHTML="";
  pendingPhotos.forEach((p,i)=>{ const im=document.createElement("img"); im.src=p.url; im.alt="หน้า "+(i+1); im.onclick=()=>{ $("#iBig").src=p.url; $("#dImg").showModal(); }; t.appendChild(im); });
}
function showManual(on){ $("#bManual").hidden=!on; $("#bManualBtn").hidden=on; }
$("#bManualBtn").addEventListener("click",()=>showManual(true));
$("#bClose2").addEventListener("click",()=>$("#dBuy").close());
$("#bFile").addEventListener("change",async e=>{
  const files=[...e.target.files]; e.target.value=""; if(!files.length) return;
  for(const f0 of files.slice(0,6)){ try{ const blob=await shrink(f0); pendingPhotos.push({blob,url:URL.createObjectURL(blob)}); }catch(err){ toast("เปิดรูปบางรูปไม่ได้ ข้ามไป"); } }
  drawThumbs(); $("#bPick").firstChild.textContent = "เพิ่มรูปหน้าถัดไป";
  if(!pendingPhotos.length) return;
  if(sample) readBuyPhoto();
  else { setMsg("#bMsg","bad","หน้านี้เรียกใช้ Claude อ่านรูปไม่ได้ (ต้องเปิดผ่านแอป Claude หรือ claude.ai ที่ล็อกอินอยู่) — กด \"ไม่มีรูป กรอกเอง\" เพื่อกรอกแทน รูปจะแนบไปด้วย"); showManual(true); }
});
$("#bRead").addEventListener("click",()=>readBuyPhoto());

async function readBuyPhoto(){
  if(!sample||!pendingPhotos.length) return;
  const btn=$("#bRead"); btn.hidden=false; btn.disabled=true; $("#bPick").disabled=true;
  setMsg("#bMsg","info busy",`กำลังอ่านใบเสร็จ ${pendingPhotos.length} หน้า และคำนวณน้ำหนัก/ต้นทุน… (ราว 20–90 วินาที)`);
  const stock=items.slice(0,250).map(i=>({id:i.id,name:i.name,spec:i.spec,unit:i.unit}));
  const prompt=`You read ${pendingPhotos.length} photo(s) (pages of ONE document) of a Thai purchase document: receipt, tax invoice, delivery note (ใบส่งสินค้า ใบกำกับภาษี ใบเสร็จ บิลเงินสด ใบรับซื้อ).
Context: OUR shop is the BUYER — ${JSON.stringify(shop.name||"ร้านเรา")} (phone ${JSON.stringify(shop.phone||"")}). Text in the customer box (นามลูกค้า/ที่อยู่/ลูกค้า) is US — never report it as the seller. The seller is the company that issued the document (letterhead, logo, footer, stamp). If the seller's name is not visible, return null.
Read every printed and handwritten text and number. Reply with ONLY a JSON object, no markdown:
{"seller": string|null, "phone": string|null (seller's phone), "billNo": string|null (document number เลขที่), "date": "YYYY-MM-DD"|null (convert พ.ศ. to C.E. by subtracting 543; 2-digit Thai year 69 = 2569),
 "kind": "new"|"used"|null (used = scrap / second-hand steel bought by weight),
 "lines":[{
   "desc": string (full description as printed, without the long product code),
   "name": string (short Thai stock name, e.g. "ท่อเหลี่ยม ZM 32x32 หนา 1.1"),
   "spec": string (size/thickness/length in Thai, e.g. "32x32 มม. หนา 1.10 มม. ยาว 6 ม."),
   "category": one of ${JSON.stringify(CATS)},
   "qty": number|null (count of pieces เส้น/แผ่น/ท่อน; for scrap sold by weight use the kg), "unit": string|null (Thai: เส้น, แผ่น, ท่อน, ชิ้น, กก.),
   "kgTotal": number|null (total weight of this line in kg if printed, e.g. "620.00 KG."; convert ตัน to kg),
   "unitPrice": number|null, "priceBasis": "kg"|"unit"|null (does unitPrice apply per kg or per piece? if qty*unitPrice ≠ amount but kgTotal*unitPrice ≈ amount, it is "kg"),
   "amount": number|null (line amount printed),
   "shape": {"type": one of ["box","pipe","shaft","sq","angle","flat","plate","metal","chan","clip","rebar"]|null, "a": mm|null, "b": mm|null, "d": outer diameter mm|null, "h": channel height mm|null, "t": thickness mm|null, "w": width mm|null (for metal = coil width before roll-forming if printed, else 914), "c": metal sheet effective/cover width mm|null, "l": plate length mm|null, "len": length in metres|null (if not printed assume 6 for pipes/tubes/bars and say so by adding "lines[i].shape.len" to uncertain)},
   "itemId": string ("" unless clearly the same product as one in STOCK)
 }],
 "shipping": number|null (transport/delivery charge ค่าขนส่ง ค่าส่ง written anywhere, including handwritten notes),
 "discount": number|null, "vat": number|null, "total": number|null (grand total if visible on any page),
 ${UNC_SPEC}}
box = square/rectangular tube (ท่อเหลี่ยม, เหล็กกล่อง, แป๊บเหลี่ยม: a x b x t); pipe = round tube (d, t); shaft = solid round bar; rebar = เหล็กเส้น; sq = solid square bar; flat = flat bar; angle = เหล็กฉาก; plate = flat sheet/plate; chan = U channel / ราง U (h,b,t); clip = C channel with lips / ตัวซี (h,b,c,t); metal = เมทัลชีท / roofing or wall sheet (t = thickness incl. coating, w = coil width, len = sheet length in metres as ordered/cut).
If a line's kgTotal is unreadable but amount and per-kg price are readable, compute kgTotal = amount / unitPrice and list "lines[i].kgTotal" in uncertain. Plain numbers, no commas.
STOCK (for itemId matching): ${JSON.stringify(stock)}`;
  try{
    const r=await sample.json(prompt,{images:pendingPhotos.map(p=>p.blob),modelTier:"complex"});
    setMsg("#bMsg","ok","✓ อ่านเสร็จแล้ว ตรวจรายการในหน้าถัดไป");
    openReview(r);
  }catch(err){
    const c=err&&err.code;
    if(c==="images_unavailable") setMsg("#bMsg","bad","แอปที่เปิดอยู่ยังส่งรูปให้ Claude อ่านไม่ได้ ลองเปิดลิงก์นี้ใน Safari/Chrome ที่ล็อกอิน claude.ai หรืออัปเดตแอป Claude — หรือกด \"ไม่มีรูป กรอกเอง\"");
    else if(c==="image_rejected") setMsg("#bMsg","bad","รูปมากเกินไปหรือไฟล์ไม่รองรับ ลองเลือกทีละ 1–3 หน้า");
    else ocrError(err,btn,"#bMsg");
    $("#bManualBtn").hidden=false;
  }finally{ btn.disabled=false; $("#bPick").disabled=false; }
}

/* ---------- review sheet: per-piece weight, cost, selling price ---------- */
const fr=$("#fRev"); let rv=null; // {lines, total, ...}
try{ const m=localStorage.getItem("steel-sellkg"); if(m) fr.sellKg.value=m; }catch(e){}
function theoKg(sh){
  if(!sh||!SHAPES[sh.type]) return 0;
  const v={a:+sh.a||0,b:+sh.b||0,d:+sh.d||0,t:+sh.t||0,w:+sh.w||0,l:+sh.l||0,len:+sh.len||0,h:+sh.h||0,c:+sh.c||0};
  if(sh.type==="metal"&&!v.w) v.w=914;
  if(sh.type==="plate"&&!v.w){ v.w=1219; v.l=v.l||2438; }
  try{ return SHAPES[sh.type].kg(v)||0; }catch(e){ return 0; }
}
function openReview(r){
  const unc=new Set(Array.isArray(r.uncertain)?r.uncertain:[]);
  clearAllWarns(fr); ocrState.fRev={msg:"#rMsg",mismatch:"",read:true,total:null};
  fillHead(fr,[["seller",r.seller],["billNo",r.billNo],["date",r.date,true]],unc);
  fr.kind.value = r.kind==="used" ? "used" : (r.kind==="new" ? "new" : fb.querySelector("[name=kind]:checked").value);
  fr.ship.value=+r.shipping||0; if(unc.has("shipping")) warnInput(fr.ship,"อ่านไม่ชัด ตรวจอีกครั้ง");
  rv={billTotal:+r.total||0, discount:+r.discount||0, vat:+r.vat||0, lines:(Array.isArray(r.lines)?r.lines:[]).map((l,i)=>{
    const it=items.find(x=>x.id===l.itemId), w={};
    const miss=(k,v)=>{ if(v==null||v===""||(typeof v==="number"&&!(v>0))) w[k]="ไม่พบในรูป"; else if(unc.has(`lines[${i}].${k}`)) w[k]="อ่านไม่ชัด"; };
    miss("name",l.name||l.desc); miss("qty",l.qty); miss("unitPrice",l.unitPrice);
    if(l.priceBasis==="kg"&&!(+l.kgTotal>0)&&+l.amount>0&&+l.unitPrice>0){ l.kgTotal=Math.round(l.amount/l.unitPrice*100)/100; w.kgTotal="อ่านไม่ได้ คำนวณจากยอดเงิน ตรวจอีกครั้ง"; }
    else if(unc.has(`lines[${i}].kgTotal`)) w.kgTotal="คำนวณจากยอดเงิน ตรวจอีกครั้ง";
    const theo=theoKg(l.shape);
    return {name:String(l.name||l.desc||""),desc:String(l.desc||""),spec:String(l.spec||""),cat:CATS.includes(l.category)?l.category:"อื่นๆ",
      qty:+l.qty||0,unit:String(l.unit||"เส้น"),kgTotal:+l.kgTotal||0,unitPrice:+l.unitPrice||0,priceBasis:l.priceBasis==="kg"?"kg":"unit",
      amount:+l.amount||0,theo,lenGuess:unc.has(`lines[${i}].shape.len`),target:it?it.id:"new",sell:0,sellKg:0,mode:"global",_w:w};
  })};
  if(!rv.lines.length) rv.lines.push({name:"",desc:"",spec:"",cat:"อื่นๆ",qty:0,unit:"เส้น",kgTotal:0,unitPrice:0,priceBasis:"unit",amount:0,theo:0,target:"new",sell:0,_w:{name:"ไม่พบในรูป",qty:"ไม่พบในรูป",unitPrice:"ไม่พบในรูป"}});
  $("#dBuy").close(); drawReview(); $("#dReview").showModal();
}
function lineAmt(l){ return l.priceBasis==="kg" ? (l.kgTotal||0)*(l.unitPrice||0) : (l.qty||0)*(l.unitPrice||0); }
function recalc(){
  const act=rv.lines.filter(l=>l.target!=="skip");
  const ship=+fr.ship.value||0, gk=+fr.sellKg.value||0;
  const kgAll=act.reduce((s,l)=>s+(l.kgTotal||(l.theo*l.qty)||0),0), amtAll=act.reduce((s,l)=>s+lineAmt(l),0);
  rv.lines.forEach(l=>{
    l.amt=lineAmt(l);
    const kgL=l.kgTotal||(l.theo*l.qty)||0;
    l.shipShare = l.target==="skip" ? 0 : (kgAll>0 ? ship*kgL/kgAll : (amtAll>0 ? ship*l.amt/amtAll : 0));
    l.kgPer = l.unit==="กก." ? 1 : (l.kgTotal&&l.qty ? l.kgTotal/l.qty : l.theo);
    l.costPer = l.qty ? (l.amt+l.shipShare)/l.qty : 0;
    if(l.mode==="global") l.sellKg=gk;
    if(l.mode==="piece") l.sellKg = l.kgPer ? Math.round(l.sell/l.kgPer*100)/100 : 0;
    else l.sell = l.sellKg&&l.kgPer ? Math.ceil(l.sellKg*l.kgPer) : 0;
  });
  return {ship,goods:amtAll,kg:kgAll};
}
function itemOpts(sel){
  return `<option value="new" ${sel==="new"?"selected":""}>+ สร้างสินค้าใหม่ในสต็อก</option><option value="skip" ${sel==="skip"?"selected":""}>ไม่เข้าสต็อก</option>`+
    [...items].sort((a,c)=>String(a.name).localeCompare(String(c.name),"th")).map(i=>`<option value="${esc(i.id)}" ${i.id===sel?"selected":""}>เพิ่มเข้า: ${esc(i.name)} ${esc(i.spec)}</option>`).join("");
}
function drawReview(){
  const t=recalc();
  $("#rLines").innerHTML=rv.lines.map((l,i)=>{ const W=k=>l._w[k]?` class="warn" title="${esc(l._w[k])}"`:"", M=k=>l._w[k]?`<small class="wmsg">${esc(l._w[k])}</small>`:"";
    const pu=l.sell-l.costPer, pk=l.unit==="กก."?"":`/${esc(l.unit||"หน่วย")}`;
    const diff=l.amount&&Math.abs(l.amount-l.amt)>1;
    return `<div class="rl ${l.target==="skip"?"skip":""}" data-i="${i}">
    <div class="hd"><b>${i+1}</b><label>ชื่อสินค้าในสต็อก<input data-k="name" value="${esc(l.name)}"${W("name")}>${M("name")}</label></div>
    ${l.desc&&l.desc!==l.name?`<div class="desc">ในบิล: ${esc(l.desc)}</div>`:""}
    <div class="g3">
      <label>จำนวน<input data-k="qty" type="number" inputmode="decimal" step="any" value="${l.qty||""}"${W("qty")}>${M("qty")}</label>
      <label>หน่วย<input data-k="unit" list="unitList" value="${esc(l.unit)}"></label>
      <label>น้ำหนักรวม กก.<input data-k="kgTotal" type="number" inputmode="decimal" step="any" value="${l.kgTotal||""}" placeholder="${l.theo&&l.qty?fmt(l.theo*l.qty)+" (สูตร)":""}"${W("kgTotal")}>${M("kgTotal")}</label>
      <label>ราคาซื้อ<input data-k="unitPrice" type="number" inputmode="decimal" step="any" value="${l.unitPrice||""}"${W("unitPrice")}>${M("unitPrice")}</label>
      <label>คิดราคาต่อ<select data-k="priceBasis"><option value="kg" ${l.priceBasis==="kg"?"selected":""}>กก.</option><option value="unit" ${l.priceBasis==="unit"?"selected":""}>${esc(l.unit||"หน่วย")}</option></select></label>
    </div>
    <div class="g2" style="margin-top:6px">
      <label>ราคาขายต่อ กก.<input data-k="sellKg" type="number" inputmode="decimal" step="any" value="${l.sellKg||""}" placeholder="ใส่ราคาขาย/กก."></label>
      <label>ราคาขาย${pk}<input data-k="sell" type="number" inputmode="decimal" step="any" value="${l.sell||""}"></label>
    </div>
    <div class="calc">
      <span>น้ำหนัก${pk}</span><span>${l.kgPer?fmt(Math.round(l.kgPer*100)/100)+" กก.":"—"}</span>
      ${l.theo&&l.unit!=="กก."?`<span class="theo">ตามสูตรเหล็ก ${fmt(Math.round(l.theo*100)/100)} กก.${l.lenGuess?" (สมมติยาว 6 ม.)":""}${l.kgTotal&&l.qty?` · ต่างจากบิล ${fmt(Math.round((l.kgTotal/l.qty-l.theo)/l.theo*1000)/10)}%`:""}</span>`:""}
      <span>ค่าสินค้าทั้งรายการ</span><span${diff?' class="neg"':""}>${money(l.amt)}${diff?` (บิล ${money(l.amount)})`:""}</span>
      <span>ต้นทุน${pk}${l.shipShare?" รวมค่าส่ง":""}</span><span>${money(l.costPer)}</span>
      <span>ทุนต่อ กก.</span><span>${l.kgPer?money(l.costPer/l.kgPer):"—"}</span>
      <span>กำไร${pk}</span><span class="${pu<0?"neg":"pos"}">${l.sell?money(pu)+(l.kgPer?` (${money(pu/l.kgPer)}/กก.)`:""):"ยังไม่ใส่ราคาขาย"}</span>
    </div>
    <div class="g2"><label>ลงสต็อก<select data-k="target">${itemOpts(l.target)}</select></label>
      ${l.target==="new"?`<label>หมวด<select data-k="cat">${CATS.map(c=>`<option ${c===l.cat?"selected":""}>${c}</option>`).join("")}</select></label>`:"<span></span>"}</div>
    ${l.target==="new"?`<label style="margin-top:6px">ขนาด/สเปก<input data-k="spec" value="${esc(l.spec)}"></label>`:""}
  </div>`; }).join("")+`<datalist id="unitList">${UNITS.map(u=>`<option>${u}</option>`).join("")}</datalist>`;
  sumReview(t);
}
function sumReview(t){
  t=t||recalc();
  $("#rGoods").textContent=money(t.goods); $("#rShip").textContent=money(t.ship);
  $("#rTotal").textContent=money(t.goods+t.ship); $("#rKg").textContent=fmt(Math.round(t.kg));
  const n=rv.lines.filter(l=>l.target!=="skip").length; $("#rOk").textContent=`ยืนยัน เข้าสต็อก ${n} รายการ`;
  const nW=fr.querySelectorAll(".warn").length, tot=rv.billTotal, cur=t.goods+t.ship-rv.discount+rv.vat;
  const mm = tot&&Math.abs(tot-cur)>1 ? `ยอดในบิล ${money(tot)} บาท ไม่ตรงกับที่คำนวณได้ ${money(cur)} บาท (ใบนี้อาจมีหลายหน้า หรือมี VAT/ส่วนลด)` : (!tot?"ไม่พบยอดรวมสุทธิในรูป ตรวจยอดกับบิลอีกครั้ง":"");
  if(nW) setMsg("#rMsg","bad",`⚠ มี ${nW} ช่องสีแดงที่อ่านไม่ได้หรือไม่แน่ใจ — กรอกหรือตรวจก่อนยืนยัน${mm?"<br>⚠ "+mm:""}`);
  else if(mm) setMsg("#rMsg","bad","⚠ "+mm);
  else setMsg("#rMsg","ok",`✓ อ่านครบ ${rv.lines.length} รายการ ยอดตรงกับบิล`);
}
$("#rLines").addEventListener("input",e=>{
  const card=e.target.closest("[data-i]"), k=e.target.dataset.k; if(!card||!k) return; const l=rv.lines[+card.dataset.i];
  l[k]=["qty","kgTotal","unitPrice","sell","sellKg"].includes(k)?(+e.target.value||0):e.target.value;
  if(k==="sell") l.mode="piece"; if(k==="sellKg") l.mode="kg";
  if(l._w[k]){ delete l._w[k]; clearWarn(e.target); }
  if(["target","priceBasis","unit","cat"].includes(k)) drawReview();
  else { recalc(); drawCalcOnly(card,l); sumReview(); }
});
function drawCalcOnly(card,l){
  const pk=l.unit==="กก."?"":`/${esc(l.unit||"หน่วย")}`, pu=l.sell-l.costPer, sp=card.querySelectorAll(".calc > span:not(.theo)");
  if(sp.length>=10){ sp[1].textContent=l.kgPer?fmt(Math.round(l.kgPer*100)/100)+" กก.":"—"; sp[3].textContent=money(l.amt); sp[5].textContent=money(l.costPer);
    sp[7].textContent=l.kgPer?money(l.costPer/l.kgPer):"—";
    sp[9].textContent=l.sell?money(pu)+(l.kgPer?` (${money(pu/l.kgPer)}/กก.)`:""):"ยังไม่ใส่ราคาขาย"; sp[9].className=pu<0&&l.sell?"neg":"pos"; }
  const si=card.querySelector("[data-k=sell]"), sk=card.querySelector("[data-k=sellKg]");
  if(si&&document.activeElement!==si) si.value=l.sell||""; if(sk&&document.activeElement!==sk) sk.value=l.sellKg||"";
}
fr.ship.addEventListener("input",()=>{ drawReview(); });
fr.sellKg.addEventListener("input",()=>{ try{ localStorage.setItem("steel-sellkg",fr.sellKg.value); }catch(e){} rv.lines.forEach(l=>l.mode="global"); drawReview(); });
fr.addEventListener("input",e=>{ if(e.target.closest("#rLines")) return; if(e.target.classList.contains("warn")){ clearWarn(e.target); } if(rv) sumReview(); });
fr.addEventListener("submit",async e=>{
  if(e.submitter&&e.submitter.value!=="save") return;
  e.preventDefault();
  const act=rv.lines.filter(l=>l.target!=="skip");
  if(act.some(l=>!l.name.trim()||!(l.qty>0))) return toast("ใส่ชื่อและจำนวนให้ครบทุกรายการที่เข้าสต็อก (หรือเลือก ไม่เข้าสต็อก)");
  const nW=fr.querySelectorAll(".warn").length; if(nW&&!confirm(`ยังมี ${nW} ช่องสีแดงที่ยังไม่ได้ตรวจ ยืนยันเข้าสต็อกเลยไหม?`)) return;
  const ok=$("#rOk"); ok.disabled=true;
  const t=recalc(), kind=fr.kind.value, at=Date.now();
  const v={kind,seller:fr.seller.value.trim()||"ไม่ระบุผู้ขาย",phone:"",billNo:fr.billNo.value.trim(),
    date:new Date((fr.date.value||today())+"T12:00").getTime(),at,shipping:t.ship,total:Math.round((t.goods+t.ship)*100)/100,note:"",kg:Math.round(t.kg*100)/100,lines:[],...payData(fr)};
  try{
    if(assets&&pendingPhotos.length){ const ids=[]; for(const p of pendingPhotos){ try{ ids.push((await assets.upload(p.blob,{type:"image/jpeg"})).id); }catch(err){} }
      if(ids.length){ v.photoIds=ids; v.photoId=ids[0]; } if(ids.length<pendingPhotos.length) toast("แนบรูปไม่ครบ บันทึกรายการต่อ"); }
    for(const l of rv.lines){
      let itemId="", addQty=0;
      if(l.target!=="skip"){
        const kgPer=Math.round(l.kgPer*1000)/1000, cost=Math.round(l.costPer*100)/100;
        if(l.target==="new"){
          itemId=uid();
          await store.set("items/"+itemId,{name:l.name.trim(),cat:l.cat,cond:kind==="used"?"used":"new",spec:l.spec.trim(),qty:l.qty,unit:l.unit||"เส้น",
            kg:kgPer,min:0,cost,price:l.sell||0,loc:"",updatedAt:at});
        }else{
          const it=items.find(i=>i.id===l.target); if(it){ itemId=it.id;
            const oq=+it.qty||0, oc=+it.cost||0, nq=oq+l.qty, nc=Math.round((oq>0&&oc>0?(oq*oc+cost*l.qty)/nq:cost)*100)/100;
            await store.update("items/"+it.id,{qty:nq,cost:nc,kg:kgPer||+it.kg||0,price:l.sell||+it.price||0,updatedAt:at}); }
        }
        addQty=l.qty;
        await addMove({type:"in",qty:l.qty,itemId,name:l.name.trim(),note:`${kind==="used"?"รับซื้อเก่า":"ซื้อใหม่"}จาก ${v.seller}${v.billNo?" บิล "+v.billNo:""}`,at});
      }
      v.lines.push({desc:l.desc||l.name,qty:l.qty,unit:l.unit,price:l.qty?Math.round(l.amt/l.qty*100)/100:0,itemId,addQty,
        kgTotal:l.kgTotal||0,kgPer:Math.round((l.kgPer||0)*1000)/1000,unitPrice:l.unitPrice,priceBasis:l.priceBasis,costPer:Math.round(l.costPer*100)/100,sell:l.sell||0,sellKg:l.sellKg||0});
    }
    await store.set("buys/"+uid(),v);
    $("#dReview").close(); toast(`เข้าสต็อก ${act.length} รายการ · จ่าย ${money(v.total)} บาท`);
    pendingPhotos.forEach(p=>URL.revokeObjectURL(p.url)); pendingPhotos=[];
  }catch(err){ toast("บันทึกไม่สำเร็จ ลองอีกครั้ง"); }
  finally{ ok.disabled=false; }
});

fb.addEventListener("submit",async e=>{
  if(e.submitter&&e.submitter.value!=="save") return;
  e.preventDefault(); if(!fb.reportValidity()) return;
  if(!fb.seller.value.trim()){ warnInput(fb.seller,"ใส่ชื่อผู้ขาย"); fb.seller.focus(); return; }
  if(!fb.date.value) fb.date.value=today();
  const nW=fb.querySelectorAll(".warn").length; if(nW&&!confirm(`ยังมี ${nW} ช่องสีแดงที่ยังไม่ได้ตรวจ บันทึกเลยไหม?`)) return;
  const ls=blines.filter(l=>String(l.desc).trim()&&(+l.qty||0)>0).map(l=>({desc:String(l.desc).trim(),qty:+l.qty,unit:l.unit||"",price:+l.price||0,
    itemId:l.itemId||"",addQty:l.itemId?(+l.addQty||0):0,...(l.grade?{grade:l.grade}:{})}));
  if(!ls.length) return toast("ใส่รายการอย่างน้อย 1 รายการ (ชื่อ + จำนวน)");
  const saveBtn=$("#bSave"); saveBtn.disabled=true;
  const kind=fb.querySelector("[name=kind]:checked").value, total=Math.round(ls.reduce((s,l)=>s+l.qty*l.price,0)*100)/100;
  const v={kind,seller:fb.seller.value.trim(),phone:fb.phone.value.trim(),billNo:fb.billNo.value.trim(),
    date:new Date(fb.date.value+"T12:00").getTime(),lines:ls,total,note:fb.note.value.trim(),
    kg:ls.filter(l=>l.unit==="กก.").reduce((s,l)=>s+l.qty,0)+ls.filter(l=>l.unit==="ตัน").reduce((s,l)=>s+l.qty*1000,0),...payData(fb)};
  if(editingBuy&&editingBuy.payStatus==="credit"&&v.payStatus==="credit") v.payments=editingBuy.payments||[];
  try{
    if(pendingPhotos.length&&assets){ const ids=[]; for(const p of pendingPhotos){ try{ ids.push((await assets.upload(p.blob,{type:"image/jpeg"})).id); }catch(err){} }
      if(ids.length){ v.photoIds=ids; v.photoId=ids[0]; } if(ids.length<pendingPhotos.length) toast("แนบรูปไม่ครบ"); }
    if(editingBuy){
      const keep=legacyLines(editingBuy);
      v.lines=ls.map((l,i)=>({...(keep[i]||{}),...l,itemId:keep[i]?.itemId||"",addQty:keep[i]?.addQty||0}));
      const {id:_old,...prev}=editingBuy; await store.set("buys/"+editingBuy.id,{...prev,...v,at:editingBuy.at||Date.now(),photoId:v.photoId||editingBuy.photoId||"",photoIds:v.photoIds||editingBuy.photoIds||[]});
    }else{
      v.at=Date.now(); const local={};
      for(const l of ls){ if(l.grade&&!l.itemId){ const gi=await gradeItem(l.grade); l.itemId=gi.id; l.addQty=l.qty; local[gi.id]=local[gi.id]||{...gi}; } }
      for(const l of ls){
        if(!l.itemId||!(l.addQty>0)) continue;
        const it=local[l.itemId]||{...items.find(i=>i.id===l.itemId)}; if(!it.id) continue;
        const oq=+it.qty||0, oc=+it.cost||0, nq=oq+l.addQty, lineTotal=l.qty*l.price;
        it.cost=Math.round((oq>0&&oc>0 ? (oq*oc+lineTotal)/nq : lineTotal/l.addQty)*100)/100; it.qty=nq; local[it.id]=it;
        await store.update("items/"+it.id,{qty:nq,cost:it.cost,updatedAt:Date.now()});
        await addMove({type:"in",qty:l.addQty,itemId:it.id,name:it.name,note:`${kind==="used"?"รับซื้อเก่า":"ซื้อใหม่"}จาก ${v.seller}${v.billNo?" บิล "+v.billNo:""}`,at:v.at});
      }
      await store.set("buys/"+uid(),v);
    }
    $("#dBuy").close(); toast(editingBuy?"บันทึกการแก้ไขแล้ว":`บันทึกซื้อเข้า ${money(total)} บาท`);
  }catch(err){ toast("บันทึกไม่สำเร็จ ลองอีกครั้ง"); }
  finally{ saveBtn.disabled=false; }
});
$("#bDel").addEventListener("click",async()=>{
  if(!editingBuy||!confirm("ลบรายการซื้อเข้านี้? (สต็อกที่เพิ่มไปแล้วจะไม่ถูกหักคืน)")) return;
  try{ await store.del("buys/"+editingBuy.id); $("#dBuy").close(); toast("ลบรายการแล้ว"); }catch(e){ toast("ลบไม่สำเร็จ"); }
});

/* ---------- cost / price / profit calculator ---------- */
let pLast="kg", pFill=false;
function pCalc(){
  const ppk=+$("#pPpk").value||0, kg=+$("#pKg").value||0;
  if(ppk>0&&kg>0&&document.activeElement&&["pPpk","pKg"].includes(document.activeElement.id)) $("#pCost").value=Math.round(ppk*kg*100)/100;
  const tc=(+$("#pCost").value||0)+(+$("#pExtra").value||0);
  if(kg>0){
    if(pLast==="kg"){ const sk=+$("#pSellKg").value||0; if(sk) $("#pPrice").value=Math.round(sk*kg*100)/100; }
    else { const pr=+$("#pPrice").value||0; $("#pSellKg").value= pr ? Math.round(pr/kg*100)/100 : ""; }
  }
  $("#pHint").innerHTML = kg>0 ? "ใส่ราคาขายต่อ กก. ระบบคูณน้ำหนักเป็นราคาต่อหน่วยให้ หรือใส่ราคาต่อหน่วย ระบบบอกราคาต่อ กก. ให้"
    : '<b class="neg">ใส่ "น้ำหนักต่อหน่วย" ก่อน</b> ราคาต่อ กก. จึงจะคิดเป็นราคาต่อหน่วยได้';
  const price=+$("#pPrice").value||0, pu=price-tc, n=+$("#pN").value||0;
  $("#pTc").textContent=money(tc); $("#pTk").textContent= kg ? money(tc/kg) : "—";
  $("#pPu").textContent=money(pu); $("#pPu").className="num "+(pu<0?"neg":"pos");
  $("#pPk").textContent= kg ? money(pu/kg) : "—"; $("#pPk").className="num "+(pu<0?"neg":"pos");
  $("#pMg").textContent= price ? fmt(Math.round(pu/price*1000)/10) : "0";
  $("#pAll").textContent=money(pu*n); $("#pAll").className="num "+(pu<0?"neg":"pos");
}
$("#dProfit").addEventListener("input",e=>{ if(e.target.id==="pSellKg") pLast="kg"; if(e.target.id==="pPrice") pLast="price"; pCalc(); });
function openProfit(fill){
  pFill=fill; $("#pUse").hidden=!fill;
  if(fill){ $("#pCost").value=+f.cost.value||0; $("#pExtra").value=0; $("#pKg").value=+f.kg.value||""; $("#pPpk").value="";
    $("#pPrice").value=+f.price.value||0; $("#pSellKg").value=""; pLast="price"; }
  pCalc(); $("#dProfit").showModal();
}
$("#profitBtn").addEventListener("click",()=>openProfit(false));
$("#eProfit").addEventListener("click",()=>openProfit(true));
$("#pClose").addEventListener("click",()=>$("#dProfit").close());
$("#pUse").addEventListener("click",()=>{
  f.cost.value=Math.round(((+$("#pCost").value||0)+(+$("#pExtra").value||0))*100)/100; f.price.value=+$("#pPrice").value||0;
  editMargin(); $("#dProfit").close();
});
function editMargin(){
  const c=+f.cost.value||0, p=+f.price.value||0;
  const kg=+f.kg.value||0;
  $("#eMargin").innerHTML = c&&p ? `กำไร <b class="${p-c<0?"neg":"pos"}">${money(p-c)}</b> บาท/หน่วย${kg?` · ขาย ${money(p/kg)}/กก. ทุน ${money(c/kg)}/กก.`:""}` : "ใส่ทุนและราคาขายเพื่อดูกำไรต่อหน่วย";
}
f.cost.addEventListener("input",editMargin); f.price.addEventListener("input",editMargin);

/* ---------- profit report tab ---------- */
let prPeriod="เดือนนี้";
function periodRange(){
  const n=new Date();
  if(prPeriod==="เดือนนี้") return [new Date(n.getFullYear(),n.getMonth(),1).getTime(),Infinity];
  if(prPeriod==="เดือนที่แล้ว") return [new Date(n.getFullYear(),n.getMonth()-1,1).getTime(),new Date(n.getFullYear(),n.getMonth(),1).getTime()];
  return [0,Infinity];
}
function renderProfit(){
  $("#prChips").innerHTML=["เดือนนี้","เดือนที่แล้ว","ทั้งหมด"].map(c=>`<button aria-pressed="${c===prPeriod}" data-pr="${c}">${c}</button>`).join("");
  const [a,b]=periodRange(), inR=t=>t>=a&&t<b;
  const rc=docs.filter(d=>(d.type==="receipt"||d.type==="credit")&&inR(d.date||d.at));
  const sales=rc.reduce((s,d)=>s+(+d.total||0),0);
  let cogs=0, noCost=0;
  rc.forEach(d=>d.lines.forEach(l=>{ if(l.itemId){ if(+l.cost>0) cogs+=l.qty*l.cost; else noCost++; } }));
  const opex=cash.filter(c=>c.type==="out"&&inR(c.date||c.at)).reduce((s,c)=>s+(+c.amount||0),0), oinc=cash.filter(c=>c.type==="in"&&inR(c.date||c.at)).reduce((s,c)=>s+(+c.amount||0)-(+c.cogs||0),0);
  const gp=sales-cogs, bought=buys.filter(x=>inR(x.date||x.at)).reduce((s,x)=>s+(+x.total||0),0);
  const stCost=items.reduce((s,i)=>s+itemVal(i,"cost"),0), stSell=items.filter(i=>+i.cost>0).reduce((s,i)=>s+itemVal(i),0);
  const withM=items.filter(i=>+i.cost>0&&+i.price>0).map(i=>({...i,pu:i.price-i.cost,mg:(i.price-i.cost)/i.price*100})).sort((x,y)=>x.mg-y.mg);
  const missing=items.filter(i=>!(+i.cost>0)).length;
  $("#profitBody").innerHTML=`
  <div class="pgrid">
    <div class="big2"><span>กำไรขั้นต้นจากการขาย (${prPeriod})</span><b class="num ${gp<0?"neg":"pos"}">${money(gp)}</b><span>${sales?fmt(Math.round(gp/sales*1000)/10)+"% ของยอดขาย":"ยังไม่มียอดขาย"}</span></div>
    <div><span>ยอดขาย รวมขายเชื่อ (${rc.length} ใบ)</span><b class="num">${fmt(sales)}</b></div>
    <div><span>ต้นทุนสินค้าที่ขาย</span><b class="num">${fmt(cogs)}</b></div>
    <div><span>ซื้อเข้า (เก่า+ใหม่)</span><b class="num">${fmt(bought)}</b></div>
    <div><span>กำไรที่รอขายในสต็อก</span><b class="num">${fmt(stSell-stCost)}</b></div>
    <div><span>ค่าใช้จ่ายดำเนินงาน</span><b class="num">${fmt(opex)}</b></div>
    <div><span>กำไรสุทธิ (ขั้นต้น + กำไรรายรับอื่น − ค่าใช้จ่าย)</span><b class="num ${gp+oinc-opex<0?"neg":"pos"}">${fmt(gp+oinc-opex)}</b></div>
  </div>
  ${noCost?`<p class="hint">มี ${noCost} รายการขายที่สินค้ายังไม่ได้ใส่ทุน จึงคิดทุนเป็น 0 — กำไรจริงจะต่ำกว่านี้</p>`:""}
  <h4 class="sec">กำไรต่อหน่วยของแต่ละสินค้า (น้อยไปมาก)</h4>
  ${withM.length?withM.map(i=>`<div class="mrow"><span>${esc(i.name)} <small>${esc(i.spec)}</small><br><small>ทุน ${money(i.cost)} → ขาย ${money(i.price)} / ${esc(i.unit)}</small></span>
    <span class="num ${i.pu<0?"neg":i.mg<10?"":"pos"}">${money(i.pu)}<br><small>${fmt(Math.round(i.mg*10)/10)}%</small></span></div>`).join("")
    :`<p class="empty">ยังไม่มีสินค้าที่ใส่ทั้งทุนและราคาขาย</p>`}
  ${missing?`<p class="hint">สินค้า ${missing} รายการยังไม่มีทุน — ใส่ได้ในหน้าแก้ไขสินค้า หรือบันทึกซื้อเข้าแล้วเลือก "เข้าสต็อก"</p>`:""}`;
}
$("#prChips").addEventListener("click",e=>{const b=e.target.closest("[data-pr]"); if(b){prPeriod=b.dataset.pr;renderProfit();}});

/* ---------- income / expense dashboard ---------- */
const CASH_CATS={out:["ค่าแรง","ค่าน้ำมัน/ขนส่ง","ค่าไฟ/ค่าน้ำ","ค่าเช่า","ค่าซ่อม/เครื่องมือ","แก๊ส/ใบตัด/วัสดุสิ้นเปลือง","ภาษี/ค่าธรรมเนียม","อื่นๆ"],
  in:["ค่าตัด/ค่าบริการ","ค่าขนส่งที่เก็บลูกค้า","ขายเศษเหล็ก","รายรับอื่นๆ"]};
const fx=$("#fCash"); let editingCash=null, cashPeriod="เดือนนี้";
function setCashCats(sel){ const t=fx.querySelector("[name=type]:checked").value; fx.cat.innerHTML=CASH_CATS[t].map(c=>`<option ${c===sel?"selected":""}>${c}</option>`).join(""); }
fx.addEventListener("change",e=>{ if(e.target.name==="type") setCashCats(); });
function openCash(c,type){
  editingCash=c||null; $("#xTitle").textContent=c?"แก้ไขรายการ":"บันทึกรายรับ-รายจ่าย"; $("#xDel").hidden=!c;
  fx.querySelector(`[name=type][value=${c?.type||type||"out"}]`).checked=true; setCashCats(c?.cat);
  fx.amount.value=c?.amount??""; fx.date.value=c?.date?isoDate(c.date):today(); fx.via.value=c?.via||"เงินสด"; fx.note.value=c?.note||"";
  $("#dCash").showModal();
}
fx.addEventListener("submit",async e=>{
  if(e.submitter&&e.submitter.value!=="save") return;
  e.preventDefault(); if(!fx.reportValidity()) return;
  const v={type:fx.querySelector("[name=type]:checked").value,cat:fx.cat.value,amount:+fx.amount.value||0,
    date:new Date(fx.date.value+"T12:00").getTime(),via:fx.via.value,note:fx.note.value.trim(),at:editingCash?.at||Date.now()};
  if(!(v.amount>0)) return toast("ใส่จำนวนเงินมากกว่า 0");
  try{ await store.set("cash/"+(editingCash?editingCash.id:uid()),v); $("#dCash").close(); toast(`บันทึก${v.type==="in"?"รายรับ":"รายจ่าย"} ${money(v.amount)} บาท`); }
  catch(err){ toast("บันทึกไม่สำเร็จ ลองอีกครั้ง"); }
});
$("#xDel").addEventListener("click",async()=>{
  if(!editingCash||!confirm("ลบรายการนี้?")) return;
  try{ await store.del("cash/"+editingCash.id); $("#dCash").close(); toast("ลบรายการแล้ว"); }catch(e){ toast("ลบไม่สำเร็จ"); }
});

function cashRange(){
  const n=new Date(), d0=new Date(n.getFullYear(),n.getMonth(),n.getDate()).getTime(), DAY=864e5;
  switch(cashPeriod){
    case "วันนี้": return {a:d0,b:d0+DAY,step:"day"};
    case "7 วัน": return {a:d0-6*DAY,b:d0+DAY,step:"day"};
    case "เดือนนี้": return {a:new Date(n.getFullYear(),n.getMonth(),1).getTime(),b:new Date(n.getFullYear(),n.getMonth()+1,1).getTime(),step:"day"};
    case "เดือนที่แล้ว": return {a:new Date(n.getFullYear(),n.getMonth()-1,1).getTime(),b:new Date(n.getFullYear(),n.getMonth(),1).getTime(),step:"day"};
    default: return {a:new Date(n.getFullYear(),0,1).getTime(),b:new Date(n.getFullYear()+1,0,1).getTime(),step:"month"};
  }
}
// unified ledger: receipts (in), other income (in), purchases (out), other expenses (out)
function ledger(){
  const L=[];
  docs.filter(d=>d.type==="receipt").forEach(d=>L.push({src:"doc",id:d.id,dir:"in",cat:"ขายสินค้า",t:d.date||d.at,amt:+d.total||0,title:`${d.customer}`,sub:`ใบเสร็จ ${d.no}`,via:d.via||"เงินสด"}));
  docs.filter(d=>d.type==="credit").forEach(d=>(d.payments||[]).forEach(p=>L.push({src:"doc",id:d.id,dir:"in",cat:"รับชำระหนี้",t:p.date||p.at,amt:+p.amount||0,title:d.customer,sub:`ชำระ ${d.no}`,via:p.via||"เงินสด"})));
  buys.forEach(b=>{ const cat=(b.kind||"used")==="used"?"รับซื้อเหล็กเก่า":"ซื้อเหล็กใหม่";
    if(b.payStatus==="credit") (b.payments||[]).forEach(p=>L.push({src:"buy",id:b.id,dir:"out",cat,t:p.date||p.at,amt:+p.amount||0,title:b.seller,sub:"จ่ายหนี้"+(b.billNo?" บิล "+b.billNo:""),via:p.via||"เงินสด"}));
    else L.push({src:"buy",id:b.id,dir:"out",cat,t:b.date||b.at,amt:+b.total||0,title:b.seller,sub:b.billNo?"บิล "+b.billNo:"",via:b.payVia||"เงินสด"}); });
  cash.forEach(c=>L.push({src:"cash",id:c.id,dir:c.type,cat:c.cat,t:c.date||c.at,amt:+c.amount||0,title:c.note||c.cat,sub:c.via||"",via:c.via||"เงินสด"}));
  return L.sort((x,y)=>y.t-x.t);
}
function renderCash(){
  if(typeof renderAlerts==="function"&&$("#alerts")) try{ renderAlerts(); }catch(e){}
  $("#cashChips").innerHTML=["วันนี้","7 วัน","เดือนนี้","เดือนที่แล้ว","ปีนี้"].map(c=>`<button aria-pressed="${c===cashPeriod}" data-cp="${c}">${c}</button>`).join("");
  const r=cashRange(), all=ledger(), L=all.filter(x=>x.t>=r.a&&x.t<r.b);
  const inc=L.filter(x=>x.dir==="in").reduce((s,x)=>s+x.amt,0), exp=L.filter(x=>x.dir==="out").reduce((s,x)=>s+x.amt,0), net=inc-exp;
  // buckets
  const B=[]; if(r.step==="day"){ for(let t=r.a;t<r.b;t+=864e5){ const d=new Date(t); B.push({a:t,b:t+864e5,lab:String(d.getDate())}); } }
  else { for(let m=0;m<12;m++){ const y=new Date(r.a).getFullYear(); B.push({a:new Date(y,m,1).getTime(),b:new Date(y,m+1,1).getTime(),lab:new Date(y,m,1).toLocaleDateString("th-TH",{month:"narrow"})}); } }
  B.forEach(k=>{ k.i=0;k.o=0; L.forEach(x=>{ if(x.t>=k.a&&x.t<k.b){ if(x.dir==="in")k.i+=x.amt; else k.o+=x.amt; } }); });
  const W=340,H=150,P=22, max=Math.max(1,...B.map(k=>Math.max(k.i,k.o))), bw=(W-P)/B.length, gw=Math.max(1,bw*0.38);
  const y=v=>H-16-(v/max)*(H-28);
  const every=B.length<=12?1:Math.ceil(B.length/8);
  const svg=`<svg viewBox="0 0 ${W} ${H}" role="img" aria-label="กราฟรายรับรายจ่าย">
    <line x1="${P}" x2="${W}" y1="${H-16}" y2="${H-16}" stroke="var(--line)"/>
    <text x="${P-3}" y="${y(max)+4}" text-anchor="end" font-size="8" fill="var(--muted)">${max>=1000?fmt(Math.round(max/100)/10)+"k":fmt(max)}</text>
    ${B.map((k,i)=>{ const x=P+i*bw+bw/2;
      return `<g><title>${k.lab}: รับ ${fmt(k.i)} / จ่าย ${fmt(k.o)}</title>
      ${k.i?`<rect x="${x-gw}" y="${y(k.i)}" width="${gw}" height="${H-16-y(k.i)}" fill="var(--ok)" rx="1"/>`:""}
      ${k.o?`<rect x="${x}" y="${y(k.o)}" width="${gw}" height="${H-16-y(k.o)}" fill="var(--low)" rx="1"/>`:""}
      ${i%every===0?`<text x="${x}" y="${H-4}" text-anchor="middle" font-size="8" fill="var(--muted)">${k.lab}</text>`:""}</g>`; }).join("")}
  </svg>`;
  const byCat=dir=>{ const m={}; L.filter(x=>x.dir===dir).forEach(x=>m[x.cat]=(m[x.cat]||0)+x.amt); return Object.entries(m).sort((a,b)=>b[1]-a[1]); };
  const catBars=(rows,cls,tot)=>rows.length?`<div class="bars">${rows.map(([c,v])=>`<div class="br"><span>${esc(c)}</span><span class="tr"><b class="${cls}" style="width:${Math.max(2,v/tot*100)}%"></b></span><span class="num">${fmt(v)}</span></div>`).join("")}</div>`:`<p class="hint">ไม่มีรายการ</p>`;
  $("#cashBody").innerHTML=`
  <div class="kpi">
    <div><span>รายรับ</span><b class="num pos">${fmt(inc)}</b></div>
    <div><span>รายจ่าย</span><b class="num neg">${fmt(exp)}</b></div>
    <div><span>คงเหลือสุทธิ</span><b class="num ${net<0?"neg":""}">${fmt(net)}</b></div>
  </div>
  <div class="chart"><div class="legend"><span><i style="background:var(--ok)"></i>รายรับ</span><span><i style="background:var(--low)"></i>รายจ่าย</span><span>${r.step==="day"?"รายวัน":"รายเดือน"}</span></div>${svg}</div>
  <div class="row" style="margin:0 0 6px"><button class="btn grow" data-new="in">+ รายรับอื่น</button><button class="btn grow" data-new="out">+ ค่าใช้จ่าย</button></div>
  <h4 class="sec">รายจ่ายแยกหมวด</h4>${catBars(byCat("out"),"",exp||1)}
  <h4 class="sec">รายรับแยกหมวด</h4>${catBars(byCat("in"),"inc",inc||1)}
  <p class="hint">เงินซื้อเหล็กเข้าสต็อกนับเป็นรายจ่ายเงินสดตรงนี้ ถ้าต้องการดูกำไรจริงจากของที่ขายไป ให้ดูแท็บกำไร</p>
  <h4 class="sec">รายการล่าสุด (${L.length})</h4>
  ${L.length?L.slice(0,60).map(x=>`<button class="tx" data-src="${x.src}" data-id="${esc(x.id)}"><span>${esc(x.title)}<br><small>${esc(x.cat)}${x.sub?" · "+esc(x.sub):""} · ${dlong(x.t)}</small></span>
    <span class="num ${x.dir==="in"?"pos":"neg"}">${x.dir==="in"?"+":"−"}${fmt(x.amt)}</span></button>`).join(""):`<p class="empty">ยังไม่มีรายการในช่วงนี้</p>`}`;
}
$("#cashChips").addEventListener("click",e=>{const b=e.target.closest("[data-cp]"); if(b){cashPeriod=b.dataset.cp;renderCash();}});
$("#cashBody").addEventListener("click",e=>{
  const n=e.target.closest("[data-new]"); if(n) return openCash(null,n.dataset.new);
  const t=e.target.closest("[data-src]"); if(!t) return; const id=t.dataset.id;
  if(t.dataset.src==="doc") openView(docs.find(d=>d.id===id));
  else if(t.dataset.src==="buy") openBuy(buys.find(b=>b.id===id));
  else openCash(cash.find(c=>c.id===id));
});

/* ---------- partial sale / cutting: full pieces + remnants ---------- */
const CUT_LEN_UNITS={"ม.":1,"ซม.":0.01,"ฟุต":0.3048,"นิ้ว":0.0254};
const CUT_MM_UNITS={"มม.":1,"ซม.":10,"นิ้ว":25.4,"ฟุต":304.8};
const cutMode=i=> !i||i.unit==="กก." ? "" : (i.cat==="เหล็กแผ่น" ? "area" : "len");
const fullLen=i=> +i.lenFull>0 ? +i.lenFull : 6;                    // m
const sheetWL=i=> [ +i.sheetW>0?+i.sheetW:1219, +i.sheetL>0?+i.sheetL:2438 ]; // mm
const remOf=i=> Array.isArray(i.rem)?i.rem:[];
function remRatio(i,r){ if(cutMode(i)==="area"){ const [W,L]=sheetWL(i); return (r.w*r.l)/(W*L); } return r.len/fullLen(i); }
const remUnits=i=> remOf(i).reduce((s,r)=>s+remRatio(i,r),0);      // in "full piece" equivalents
const itemUnits=i=> (+i.qty||0)+remUnits(i);
const itemKg=i=> itemUnits(i)*(+i.kg||0);
const itemVal=(i,k="price")=> itemUnits(i)*(+i[k]||0);
const remLabel=(i,r)=> cutMode(i)==="area" ? `${fmt(Math.round(r.w))}×${fmt(Math.round(r.l))} มม.` : `${fmt(Math.round(r.len*100)/100)} ม.`;
function remSummary(i){
  const r=remOf(i); if(!r.length) return "";
  const s=[...r].sort((a,b)=>remRatio(i,b)-remRatio(i,a));
  return `${cutMode(i)==="area"?"แผ่นเหลือ":"ท่อนเหลือ"} ${s.length}: ${s.slice(0,4).map(x=>remLabel(i,x)).join(", ")}${s.length>4?" …":""}`;
}
let cutCfg={minLen:0.3,minMm:100};
try{ const c=JSON.parse(localStorage.getItem("steel-cutcfg")||"null"); if(c) cutCfg={...cutCfg,...c}; }catch(e){}

// plan cutting n pieces from item. spec: {x:m} or {w,l:mm}. returns plan or {error}
function planCut(item,spec,n,useOld=true){
  const mode=cutMode(item); if(!mode) return {error:"สินค้านี้ขายเป็น กก. ใช้จ่ายออกตามน้ำหนักได้เลย"};
  let full=+item.qty||0; const pool=remOf(item).map(r=>({...r,old:true})); const steps=[]; let fullUsed=0;
  if(mode==="len"){
    const L=fullLen(item), x=spec.x; if(!(x>0)) return {error:"ใส่ความยาวที่ตัด"}; if(x>L+1e-9) return {error:`ยาวเกินเส้นเต็ม (${fmt(L)} ม.)`};
    for(let k=0;k<n;k++){
      const cand=pool.filter(r=>r.len>=x-1e-9&&(useOld||!r.old)).sort((a,b)=>a.len-b.len)[0];
      if(cand){ steps.push({from:`ท่อน ${fmt(Math.round(cand.len*100)/100)} ม.`,left:cand.len-x}); cand.len=Math.round((cand.len-x)*1000)/1000; }
      else{ if(full<1) return {error:`เส้นเต็มไม่พอ (เหลือ ${fmt(full)} ${item.unit})`}; full--; fullUsed++;
        const r={id:uid(),len:Math.round((L-x)*1000)/1000,old:false}; pool.push(r); steps.push({from:`เส้นเต็ม ${fmt(L)} ม.`,left:r.len}); }
    }
    const keep=[], scrap=[];
    pool.forEach(r=>{ if(r.len<=1e-6) return; (r.len<cutCfg.minLen?scrap:keep).push(r); });
    const ratio=x/L;
    return {mode,full,fullUsed,keep,scrap,steps,ratio,scrapRatio:scrap.reduce((s,r)=>s+r.len/L,0)};
  }
  const [W,L]=sheetWL(item), w=spec.w, l=spec.l; if(!(w>0&&l>0)) return {error:"ใส่ขนาดที่ตัด"};
  const fits=(r,a,b)=>r.w>=a-1e-6&&r.l>=b-1e-6;
  if(!fits({w:W,l:L},w,l)&&!fits({w:W,l:L},l,w)) return {error:`ใหญ่เกินแผ่นเต็ม (${W}×${L} มม.)`};
  for(let k=0;k<n;k++){
    // pick smallest sheet that fits; orientation leaving the biggest single offcut
    const score=(r,a,b)=>Math.max(r.w*(r.l-b),(r.w-a)*b);
    const choose=list=>{ let best=null; list.forEach(r=>{ [[w,l],[l,w]].forEach(([a,b])=>{ if(!fits(r,a,b)) return; const s=score(r,a,b), ar=r.w*r.l;
      if(!best||ar<best.ar-1e-6||(Math.abs(ar-best.ar)<1e-6&&s>best.s)) best={r,a,b,s,ar}; }); }); return best; };
    let best=choose(pool.filter(r=>useOld||!r.old));
    if(!best){ if(full<1) return {error:`แผ่นเต็มไม่พอ (เหลือ ${fmt(full)} ${item.unit})`}; full--; fullUsed++;
      const r={id:uid(),w:W,l:L,old:false}; pool.push(r); best=choose([r]); }
    const {r,a,b}=best, from=`${r.old?"แผ่นเหลือ":"แผ่นเต็ม"} ${fmt(Math.round(r.w))}×${fmt(Math.round(r.l))}`;
    pool.splice(pool.indexOf(r),1);
    const A={id:uid(),w:r.w,l:r.l-b,old:r.old}, B={id:uid(),w:r.w-a,l:b,old:r.old};
    [A,B].forEach(x=>{ if(x.w>0.5&&x.l>0.5) pool.push(x); });
    steps.push({from,left:[A,B].filter(x=>x.w>0.5&&x.l>0.5).map(x=>`${fmt(Math.round(x.w))}×${fmt(Math.round(x.l))}`).join(" + ")||"ไม่เหลือ"});
  }
  const keep=[], scrap=[]; pool.forEach(r=>(Math.min(r.w,r.l)<cutCfg.minMm?scrap:keep).push(r));
  return {mode,full,fullUsed,keep,scrap,steps,ratio:(w*l)/(W*L),scrapRatio:scrap.reduce((s,r)=>s+r.w*r.l/(W*L),0)};
}
// write plan to stock
async function applyCut(item,plan,n,label,note,at){
  await store.update("items/"+item.id,{qty:plan.full,rem:plan.keep.map(({old,...r})=>r),updatedAt:at});
  await addMove({type:"out",qty:n,label:`−${fmt(n)} ${label}`,itemId:item.id,name:item.name,
    note:`${note}${plan.fullUsed?` · ใช้${plan.mode==="area"?"แผ่น":"เส้น"}เต็ม ${plan.fullUsed}`:""}`,at});
  const sk=plan.scrapRatio*(+item.kg||0);
  if(sk>0.01){
    let s=items.find(i=>i.unit==="กก."&&/เศษ/.test(i.name));
    if(s) await store.update("items/"+s.id,{qty:Math.round(((+s.qty||0)+sk)*100)/100,updatedAt:at});
    else { const sid=uid(); await store.set("items/"+sid,{name:"เศษเหล็ก (จากการตัด)",cat:"อื่นๆ",cond:"used",spec:"",qty:Math.round(sk*100)/100,unit:"กก.",kg:1,min:0,cost:0,price:0,loc:"",updatedAt:at}); }
    await addMove({type:"in",qty:Math.round(sk*100)/100,label:`+${fmt(Math.round(sk*100)/100)} กก.`,itemId:"",name:"เศษเหล็ก (จากการตัด)",note:`เศษจากตัด ${item.name}`,at});
  }
}

/* ---------- cut sheet UI ---------- */
let cutItem=null, cutFor="stock", cutPlan=null;
function openCut(item,forWhat){
  cutItem=item; cutFor=forWhat||"stock"; const mode=cutMode(item);
  $("#kTitle").textContent="แบ่งขาย / ตัด — "+item.name;
  $("#kLenRow").hidden=mode!=="len"; $("#kAreaRow").hidden=mode!=="area";
  $("#kInfo").textContent = mode==="area" ? `แผ่นเต็ม ${sheetWL(item).join("×")} มม. มี ${fmt(item.qty)} ${item.unit}` : `เส้นเต็ม ${fmt(fullLen(item))} ม. มี ${fmt(item.qty)} ${item.unit}`;
  $("#kN").value=1; $("#kX").value=""; $("#kW").value=""; $("#kL").value=""; $("#kFee").value=0; $("#kOld").checked=true; $("#kNote").value="";
  $("#kMin").value = mode==="area" ? cutCfg.minMm : Math.round(cutCfg.minLen*100);
  $("#kMinU").textContent = mode==="area" ? "มม. (ด้านสั้น)" : "ซม.";
  $("#kPkg").value = +item.kg>0&&+item.price>0 ? Math.round(item.price/item.kg*100)/100 : "";
  $("#kPkgRow").hidden=!(+item.kg>0);
  $("#kStock").hidden=cutFor!=="stock"; $("#kToDoc").hidden=cutFor!=="stock"; $("#kAdd").hidden=cutFor==="stock"; $("#kNoteRow").hidden=cutFor!=="stock";
  cutCalc(); $("#dCut").showModal();
}
function cutSpec(){
  if(cutMode(cutItem)==="area") return {w:(+$("#kW").value||0)*CUT_MM_UNITS[$("#kWU").value], l:(+$("#kL").value||0)*CUT_MM_UNITS[$("#kLU").value]};
  return {x:(+$("#kX").value||0)*CUT_LEN_UNITS[$("#kXU").value]};
}
function cutLabel(spec){ return cutMode(cutItem)==="area" ? `${fmt(Math.round(spec.w))}×${fmt(Math.round(spec.l))} มม.` : `${fmt(Math.round(spec.x*1000)/1000)} ม.`; }
function cutCalc(){
  const mode=cutMode(cutItem), m=+$("#kMin").value||0;
  if(mode==="area") cutCfg.minMm=m; else cutCfg.minLen=m/100;
  try{ localStorage.setItem("steel-cutcfg",JSON.stringify(cutCfg)); }catch(e){}
  const n=Math.max(1,Math.round(+$("#kN").value||1)), spec=cutSpec(), it=cutItem;
  const has = mode==="area" ? spec.w>0&&spec.l>0 : spec.x>0;
  if(!has){ cutPlan=null; $("#kPlan").innerHTML=`<p class="hint">ใส่ขนาดที่จะตัด ระบบจะเลือกท่อนให้และบอกราคา</p>`; $("#kRes").hidden=true; return; }
  cutPlan=planCut(it,spec,n,$("#kOld").checked);
  if(cutPlan.error){ $("#kPlan").innerHTML=`<div class="msg bad">⚠ ${esc(cutPlan.error)}</div>`; $("#kRes").hidden=true; return; }
  const p=cutPlan, kgPiece=p.ratio*(+it.kg||0), pkg=+$("#kPkg").value||0, fee=+$("#kFee").value||0;
  const pricePiece = (+it.kg>0&&pkg>0 ? kgPiece*pkg : p.ratio*(+it.price||0)) + fee;
  p.pricePiece=Math.ceil(pricePiece); p.costPiece=Math.round(p.ratio*(+it.cost||0)*100)/100; p.kgPiece=kgPiece; p.n=n; p.spec=spec;
  $("#kPlan").innerHTML=`<ol class="steps">${p.steps.map(s=>`<li>ตัดจาก${esc(s.from)} → เหลือ ${typeof s.left==="number"?fmt(Math.round(s.left*100)/100)+" ม.":esc(s.left)}</li>`).join("")}</ol>
    <p class="hint">หลังตัด: ${it.unit}เต็มเหลือ ${fmt(p.full)} · ${p.mode==="area"?"แผ่น":"ท่อน"}เหลือ ${p.keep.length} ${p.keep.length?"("+p.keep.slice(0,5).map(r=>remLabel(it,r)).join(", ")+")":""}${p.scrap.length?` · <b class="neg">เป็นเศษ ${p.scrap.length} ชิ้น ≈${fmt(Math.round(p.scrapRatio*(+it.kg||0)*100)/100)} กก.</b>`:""}</p>`;
  $("#kRes").hidden=false;
  $("#kKg").textContent= +it.kg>0 ? fmt(Math.round(kgPiece*n*100)/100) : "—";
  $("#kPP").textContent=money(p.pricePiece); $("#kAll").textContent=money(p.pricePiece*n);
  const pr=(p.pricePiece-p.costPiece)*n; $("#kPr").textContent= +it.cost>0 ? money(pr) : "ยังไม่มีทุน"; $("#kPr").className="num "+(pr<0?"neg":"pos");
}
$("#dCut").addEventListener("input",cutCalc); $("#dCut").addEventListener("change",cutCalc);
$("#kClose").addEventListener("click",()=>$("#dCut").close());
function cutLine(){
  const p=cutPlan, it=cutItem, lab=cutLabel(p.spec);
  return {itemId:it.id,name:`${it.name}${it.spec?" "+it.spec:""} ตัด ${lab}`,qty:p.n,unit:cutMode(it)==="area"?"ชิ้น":"ท่อน",price:p.pricePiece,cost:p.costPiece,
    cut:{mode:p.mode,x:p.spec.x||0,w:p.spec.w||0,l:p.spec.l||0,useOld:$("#kOld").checked,label:lab}};
}
$("#kStock").addEventListener("click",async()=>{
  if(!cutPlan||cutPlan.error) return toast("ตรวจขนาดที่ตัดก่อน");
  const b=$("#kStock"); b.disabled=true;
  try{ const at=Date.now(), lab=cutLabel(cutPlan.spec);
    await applyCut(cutItem,cutPlan,cutPlan.n,`${cutPlan.mode==="area"?"ชิ้น":"ท่อน"} ${lab}`,`แบ่งขาย ${fmt(cutPlan.n)} × ${lab} · ${money(cutPlan.pricePiece*cutPlan.n)} บาท${$("#kNote").value.trim()?" · "+$("#kNote").value.trim():""}`,at);
    $("#dCut").close(); $("#dMove").close(); toast(`ตัดสต็อกแล้ว ${fmt(cutPlan.n)} × ${lab}`);
  }catch(e){ toast("บันทึกไม่สำเร็จ ลองอีกครั้ง"); } finally{ b.disabled=false; }
});
$("#kToDoc").addEventListener("click",()=>{
  if(!cutPlan||cutPlan.error) return toast("ตรวจขนาดที่ตัดก่อน");
  const l=cutLine(); $("#dCut").close(); $("#dMove").close(); openDoc(); fd.type.value="receipt"; typeChanged(); lines.push(l); drawLines();
});
$("#kAdd").addEventListener("click",()=>{
  if(!cutPlan||cutPlan.error) return toast("ตรวจขนาดที่ตัดก่อน");
  if(cutFor==="quick"){ qcart.push(cutLine()); drawQS(); } else { lines.push(cutLine()); drawLines(); } $("#dCut").close();
});

/* remnants in item sheet */
function drawRem(){
  const it=current, mode=cutMode(it); $("#mCutBox").hidden=!mode; if(!mode) return;
  const r=remOf(it);
  $("#mRem").innerHTML = r.length ? r.map(x=>`<span class="chip">${remLabel(it,x)}<button type="button" data-rmrem="${esc(x.id)}" aria-label="เอาท่อนนี้ออก">×</button></span>`).join("")
    : `<span class="hint">ยังไม่มี${mode==="area"?"แผ่น":"ท่อน"}เหลือ</span>`;
  $("#mRemLbl").textContent = mode==="area" ? "แผ่นเหลือ (เศษแผ่น)" : "ท่อนเหลือ";
  $("#mRemAddW").hidden=mode!=="area";
  $("#mRemAdd").placeholder = mode==="area" ? "ยาว มม." : "ยาว ม.";
}
$("#mCut").addEventListener("click",()=>openCut(current,"stock"));
$("#mRem").addEventListener("click",async e=>{
  const id=e.target.dataset.rmrem; if(!id) return;
  if(!confirm("เอาท่อน/แผ่นนี้ออกจากสต็อก? (เช่น ทิ้งเป็นเศษหรือนับแล้วไม่มี)")) return;
  const rem=remOf(current).filter(r=>r.id!==id);
  try{ await store.update("items/"+current.id,{rem,updatedAt:Date.now()}); current={...current,rem}; drawRem(); }catch(err){ toast("บันทึกไม่สำเร็จ"); }
});
$("#mRemBtn").addEventListener("click",async()=>{
  const mode=cutMode(current), v=+$("#mRemAdd").value||0, w=+$("#mRemAddW").value||0;
  if(!(v>0)||(mode==="area"&&!(w>0))) return toast("ใส่ขนาดท่อนที่จะเพิ่ม");
  const r= mode==="area" ? {id:uid(),w,l:v} : {id:uid(),len:v};
  const rem=[...remOf(current),r];
  try{ await store.update("items/"+current.id,{rem,updatedAt:Date.now()}); current={...current,rem}; $("#mRemAdd").value=""; $("#mRemAddW").value=""; drawRem(); toast("เพิ่มท่อนเหลือแล้ว"); }catch(err){ toast("บันทึกไม่สำเร็จ"); }
});

/* ---------- quotes & receipts ---------- */
const fd=$("#fDoc"); let lines=[];
const today=()=>{ const d=new Date(); return new Date(d-d.getTimezoneOffset()*6e4).toISOString().slice(0,10); };
let docPhoto=null, docPhotoUrl=null;
function openDoc(from){
  orderLink=null;
  lines=from?from.lines.map(l=>({...l})):[];
  docPhoto=null; if(docPhotoUrl){URL.revokeObjectURL(docPhotoUrl);docPhotoUrl=null;} $("#dImgPrev").hidden=true; $("#dRead").hidden=true;
  $("#dPickImg").hidden=!assets&&!sampleImg; clearAllWarns(fd); ocrState.fDoc={msg:"#dMsg",mismatch:"",read:false}; setMsg("#dMsg");
  fd.type.value=from?"receipt":"quote"; fd.date.value=today(); fd.via.value="เงินสด"; fd.dueDays.value=30;
  shipLoad(fd,from?{...from,shipFee:0}:null);
  fd.customer.value=from?.customer||""; fd.phone.value=from?.phone||""; fd.addr.value=from?.addr||"";
  fd.discount.value=from?.discount||0; fd.note.value=""; fd.cut.checked=true; fd.showPhone.checked=from?.showPhone!==false;
  $("#dPick").innerHTML=`<option value="">— เลือกสินค้า —</option>`+[...items].sort((a,c)=>String(a.name).localeCompare(String(c.name),"th"))
    .map(i=>`<option value="${esc(i.id)}">${esc(i.name)} ${esc(i.spec)} · ${fmt(i.price)}฿</option>`).join("");
  $("#dPickCut").innerHTML=`<option value="">— เลือกสินค้าที่จะตัด —</option>`+[...items].filter(i=>cutMode(i)).sort((a,c)=>String(a.name).localeCompare(String(c.name),"th"))
    .map(i=>`<option value="${esc(i.id)}">${esc(i.name)} ${esc(i.spec)} · เต็ม ${fmt(i.qty)}${remOf(i).length?" + เหลือ "+remOf(i).length:""}</option>`).join("");
  drawLines(); typeChanged(); $("#dDoc").showModal();
}
$("#dPickCut").addEventListener("change",e=>{ const it=items.find(i=>i.id===e.target.value); e.target.value=""; if(it) openCut(it,"doc"); });
fd.type.addEventListener("change",typeChanged);
$("#dPick").addEventListener("change",e=>{
  const it=items.find(i=>i.id===e.target.value); if(!it) return;
  lines.push({itemId:it.id,name:[it.name,it.spec].filter(Boolean).join(" "),qty:1,unit:it.unit,price:+it.price||0,cost:+it.cost||0});
  e.target.value=""; drawLines();
});
$("#dBlank").addEventListener("click",()=>{ lines.push({itemId:"",name:"",qty:1,unit:"",price:0}); drawLines(); });
function drawLines(){
  $("#dLines").innerHTML = lines.length ? lines.map((l,i)=>`
    <div class="line" data-i="${i}">
      <input class="nm${wc(l,"name")}" data-k="name" value="${esc(l.name)}" placeholder="รายการ (เช่น ค่าตัดเหล็ก)" maxlength="100">
      <input class="${wc(l,"qty")}" data-k="qty" type="number" inputmode="decimal" min="0" step="any" value="${l.qty||""}" aria-label="จำนวน" placeholder="จำนวน">
      <input class="${wc(l,"unit")}" data-k="unit" value="${esc(l.unit)}" placeholder="หน่วย" maxlength="10">
      <input class="${wc(l,"price")}" data-k="price" type="number" inputmode="decimal" min="0" step="any" value="${l.price||""}" aria-label="ราคาต่อหน่วย" placeholder="ราคา/หน่วย">
      <button type="button" class="rm" data-rm="${i}" aria-label="ลบบรรทัด">×</button>
      <div class="sub">${l.itemId?"จากสต็อก · ":""}รวม <span class="num">${money(l.qty*l.price)}</span></div>
    </div>`).join("") : `<p class="hint">เลือกสินค้าจากสต็อก หรือเพิ่มบรรทัดเปล่าสำหรับค่าตัด/ค่าขนส่ง</p>`;
  docTotal();
}
$("#dLines").addEventListener("input",e=>{
  const row=e.target.closest("[data-i]"); if(!row) return; const l=lines[+row.dataset.i], k=e.target.dataset.k;
  l[k]= (k==="qty"||k==="price") ? (+e.target.value||0) : e.target.value; if(l._w) delete l._w[k];
  row.querySelector(".sub .num").textContent=money(l.qty*l.price); docTotal();
});
$("#dLines").addEventListener("click",e=>{ const r=e.target.dataset.rm; if(r!==undefined){ lines.splice(+r,1); drawLines(); } });
fd.discount.addEventListener("input",docTotal);
$("#dPickImg").addEventListener("click",()=>$("#dFile").click());
$("#dImgPrev").addEventListener("click",()=>{ $("#iBig").src=$("#dImgPrev").src; $("#dImg").showModal(); });
$("#dFile").addEventListener("change",async e=>{
  const file=e.target.files[0]; e.target.value=""; if(!file) return;
  try{ docPhoto=await shrink(file); }catch(err){ return toast("เปิดรูปนี้ไม่ได้ ลองถ่ายใหม่หรือเลือกไฟล์ JPG/PNG"); }
  if(docPhotoUrl) URL.revokeObjectURL(docPhotoUrl); docPhotoUrl=URL.createObjectURL(docPhoto);
  $("#dImgPrev").src=docPhotoUrl; $("#dImgPrev").hidden=false;
  if(sampleImg) readDocPhoto(); else setMsg("#dMsg","info","รูปจะแนบไปกับเอกสารเมื่อบันทึก (อ่านข้อมูลจากรูปอัตโนมัติไม่ได้ในหน้านี้)");
});
$("#dRead").addEventListener("click",()=>readDocPhoto());
async function readDocPhoto(){
  if(!sample||!docPhoto) return;
  const btn=$("#dRead"); btn.hidden=false; btn.disabled=true;
  setMsg("#dMsg","info busy","กำลังอ่านข้อความและตัวเลขในรูป… (ราว 10–60 วินาที)");
  const stock=items.slice(0,250).map(i=>({id:i.id,name:i.name,spec:i.spec,unit:i.unit,price:+i.price||0}));
  const prompt=`You read a photo for a Thai steel shop that is SELLING to a customer: a customer's order list, handwritten note, LINE chat screenshot, purchase order, or an old quotation/bill.
Read every printed or handwritten text and number. Reply with ONLY a JSON object, no markdown:
{"customer": string|null (the BUYER's name or company), "phone": string|null, "addr": string|null, "date": "YYYY-MM-DD"|null (convert พ.ศ. to C.E. by subtracting 543),
 "lines":[{"desc": string|null (item as written, include size/thickness/length), "qty": number|null, "unit": string|null (Thai: เส้น, แผ่น, ท่อน, กก., ชิ้น, ตัน, ครั้ง), "price": number|null (price per unit ONLY if written in the photo), "itemId": string ("" unless clearly the same product as one in STOCK)}],
 "discount": number|null, "total": number|null,
 ${UNC_SPEC}}
Plain numbers without commas.
STOCK (for itemId matching): ${JSON.stringify(stock)}`;
  try{
    const r=await sample.json(prompt,{images:[docPhoto],modelTier:"default"});
    const unc=new Set(Array.isArray(r.uncertain)?r.uncertain:[]);
    clearAllWarns(fd);
    fillHead(fd,[["customer",r.customer],["date",r.date,true]],unc);
    if(r.phone) fd.phone.value=r.phone; if(r.addr) fd.addr.value=r.addr;
    if(+r.discount>0) fd.discount.value=+r.discount;
    let fromStock=0;
    const src=Array.isArray(r.lines)&&r.lines.length ? r.lines : [{}];
    lines=src.map((l,i)=>{ const it=items.find(x=>x.id===l.itemId);
      const w=lineWarns(l,i,[["name","desc"],["qty"],["unit"]],unc);
      let price=+l.price||0;
      if(!price&&it&&+it.price>0){ price=+it.price; fromStock++; }
      else if(!price) w.price="ไม่พบราคาในรูป";
      else if(unc.has(`lines[${i}].price`)) w.price="อ่านไม่ชัด";
      if(it&&!l.unit){ delete w.unit; }
      return {itemId:it?it.id:"",name:String(l.desc??"")||(it?[it.name,it.spec].filter(Boolean).join(" "):""),qty:+l.qty||0,unit:String(l.unit??"")||(it?it.unit:""),price,cost:it?+it.cost||0:0,_w:w}; });
    drawLines();
    ocrState.fDoc.read=true; ocrState.fDoc.total=+r.total||0;
    refreshBanner(fd);
    if(fromStock) $("#dMsg").insertAdjacentHTML("beforeend",`<br><small>ใส่ราคาขายจากสต็อกให้ ${fromStock} รายการ (ในรูปไม่มีราคา)</small>`);
  }catch(err){ ocrError(err,btn,"#dMsg"); }
  finally{ btn.disabled=false; }
}
const sumLines=ls=>ls.reduce((s,l)=>s+(+l.qty||0)*(+l.price||0),0);
function docTotal(){
  const net=Math.max(0,sumLines(lines)-(+fd.discount.value||0)); $("#dTotal").textContent=money(net);
  const cost=lines.reduce((s,l)=>s+(+l.qty||0)*(+l.cost||0),0), miss=lines.some(l=>l.itemId&&!(+l.cost>0));
  $("#dProfitEst").innerHTML = lines.length ? `กำไรประมาณ <b class="${net-cost<0?"neg":"pos"}">${money(net-cost)}</b> บาท${miss?" (บางรายการยังไม่มีทุน)":""} — ไม่แสดงบนเอกสาร` : "";
}
function nextNo(type,date){
  const p=({quote:"QT",receipt:"RC",credit:"DN"}[type]||"RC")+date.replace(/-/g,"").slice(2,6);
  const max=docs.filter(d=>String(d.no).startsWith(p)).reduce((m,d)=>Math.max(m,+String(d.no).slice(p.length+1)||0),0);
  return `${p}-${String(max+1).padStart(3,"0")}`;
}
fd.addEventListener("submit",async e=>{
  if(e.submitter&&e.submitter.value!=="save") return;
  e.preventDefault(); if(!fd.reportValidity()) return;
  const nW=fd.querySelectorAll(".warn").length; if(nW&&!confirm(`ยังมี ${nW} ช่องสีแดงที่ยังไม่ได้ตรวจ บันทึกเลยไหม?`)) return;
  const ls=lines.filter(l=>String(l.name).trim()&&l.qty>0);
  if(!ls.length) return toast("เพิ่มรายการอย่างน้อย 1 บรรทัด");
  const type=fd.type.value;
  if(type==="credit"&&/ลูกค้าหน้าร้าน|ทั่วไป/.test(fd.customer.value)) return toast("ขายเชื่อต้องใส่ชื่อลูกค้าจริง");
  try{
    const d=await commitDoc({type,dateStr:fd.date.value,customer:fd.customer.value.trim(),phone:fd.phone.value.trim(),addr:fd.addr.value.trim(),
      lines:ls,discount:fd.discount.value,note:fd.note.value.trim(),cut:fd.cut.checked,showPhone:fd.showPhone.checked,via:fd.via.value,dueDays:fd.dueDays.value,photo:docPhoto,ship:shipData(fd)});
    if(orderLink){ try{ await store.update("orders/"+orderLink.id,{status:type==="quote"?"quoted":"done",docId:d.id,docNo:d.no,updatedAt:Date.now()}); }catch(e){} orderLink=null; }
    $("#dDoc").close(); toast(`บันทึก${DOCNAME[type]} ${d.no} แล้ว`); openView(d);
  }catch(err){ toast(err&&err.message&&!/^\[/.test(err.message)?err.message:"บันทึกไม่สำเร็จ ลองอีกครั้ง"); }
});

/* ---------- printable paper ---------- */
function bahtText(n){
  const D=["","หนึ่ง","สอง","สาม","สี่","ห้า","หก","เจ็ด","แปด","เก้า"], P=["","สิบ","ร้อย","พัน","หมื่น","แสน"];
  const rd=(s,big)=>{ s=String(s).replace(/^0+/,""); if(!s) return "";
    if(s.length>6) return rd(s.slice(0,-6))+"ล้าน"+rd(s.slice(-6),true);
    let o=""; const L=s.length; for(let i=0;i<L;i++){ const d=+s[i], p=L-1-i; if(!d) continue;
      if(p===1&&d===1) o+="สิบ"; else if(p===1&&d===2) o+="ยี่สิบ"; else if(p===0&&d===1&&(L>1||big)) o+="เอ็ด"; else o+=D[d]+P[p]; } return o; };
  let b=Math.floor(n), st=Math.round((n-b)*100); if(st===100){b++;st=0;}
  return (b?rd(b)+"บาท":(st?"":"ศูนย์บาท"))+(st?rd(st)+"สตางค์":"ถ้วน");
}
function paperHTML(d, format = "a4"){
  if(format === "slip"){
    return `<div class="paper paper-slip">
  <div class="ph">
    <div class="shop">
      <b>${esc(shop.name || "คลังเหล็ก")}</b>
      ${shop.addr ? esc(shop.addr) + "<br>" : ""}
      ${shop.phone && d.showPhone !== false ? "โทร " + esc(shop.phone) : ""}
      ${shop.tax ? "<br>เลขผู้เสียภาษี " + esc(shop.tax) : ""}
    </div>
    <div class="dt">
      <h2>${d.type === "credit" ? "ใบส่งของ / ใบแจ้งหนี้" : DOCNAME[d.type]}</h2>
      <div>เลขที่: <b>${esc(d.no)}</b></div>
      <div>วันที่: ${dstr(d.date || d.at)}</div>
      ${d.type === "credit" ? `<div>ครบกำหนด: ${dlong(d.due || d.date)}</div>` : ""}
      <div>ออกโดย: ${currentRole === "owner" ? "เจ้าของร้าน" : "พนักงาน (ลูกน้อง)"}</div>
    </div>
  </div>
  <div class="cust">
    <b>ลูกค้า:</b> ${esc(d.customer)}
    ${d.phone && d.showPhone !== false ? "<br>โทร: " + esc(d.phone) : ""}
    ${d.addr ? "<br>" + esc(d.addr) : ""}
    ${d.ful === "ship" ? `<br><b>จัดส่งที่:</b> ${esc(d.shipAddr || "-")}${d.shipPhone && d.showPhone !== false ? " · ผู้รับ: " + esc(d.shipPhone) : ""}` : ""}
  </div>
  <table>
    <thead><tr><th style="text-align:left">รายการ</th><th class="r">จน.</th><th class="r">ราคา</th><th class="r">รวม</th></tr></thead>
    <tbody>
    ${d.lines.map((l, i) => `<tr>
      <td colspan="4" class="slip-item-name">${i + 1}. ${esc(l.name)}</td>
    </tr>
    <tr class="slip-item-calc">
      <td></td>
      <td class="r">${fmt(l.qty)} ${esc(l.unit)}</td>
      <td class="r">${money(l.price)}</td>
      <td class="r"><b>${money(l.qty * l.price)}</b></td>
    </tr>`).join("")}
    </tbody>
  </table>
  <div class="tot">
    <table>
      <tr><td>รวมเป็นเงิน</td><td class="r">${money(d.sub)}</td></tr>
      ${d.discount ? `<tr><td>ส่วนลด</td><td class="r">-${money(d.discount)}</td></tr>` : ""}
      <tr><th style="font-size:13px">ยอดสุทธิ</th><th class="r" style="font-size:14px"><b>${money(d.total)} บาท</b></th></tr>
    </table>
    <div class="words">(${bahtText(d.total)})</div>
  </div>
  ${d.type === "receipt" && d.via ? `<div class="note">ชำระโดย: <b>${esc(d.via)}</b></div>` : ""}
  ${d.type === "credit" ? `<div class="note" style="border:1px dashed #000;padding:4px">ชำระแล้ว ${money(sumPay(d.payments))} บาท · <b>คงค้าง ${money(docOwed(d))} บาท</b></div>` : ""}
  ${shop.bankName && shop.bankAcc ? `
  <div class="bank-box">
    <div><b>โอนเข้าบัญชี:</b> ${esc(shop.bankName)}</div>
    <div>เลขที่: <b>${esc(shop.bankAcc)}</b></div>
    ${shop.bankHolder ? `<div>ชื่อบัญชี: ${esc(shop.bankHolder)}</div>` : ""}
  </div>` : ""}
  ${(d.type === "quote" || (d.type === "credit" && docOwed(d) > 0) || d.type === "receipt") && shop.promptpay ? (() => {
    const amt = d.type === "credit" ? docOwed(d) : (+d.total || 0);
    const u = qrDataURL(ppPayload(shop.promptpay, amt));
    return u ? `<div class="qrbox"><img src="${u}" alt="QR"><div>สแกนจ่ายพร้อมเพย์: ${esc(shop.promptpay)}<br>ยอดชำระ <b>${money(amt)} บาท</b></div></div>` : "";
  })() : ""}
  ${d.note ? `<div class="note">หมายเหตุ: ${esc(d.note)}</div>` : ""}
  <div class="slip-footer">
    <div>${shop.billNote ? esc(shop.billNote) : "สินค้าซื้อแล้วไม่รับเปลี่ยนหรือคืน · โปรดตรวจนับสินค้าก่อนรับมอบ"}</div>
    <div style="margin-top:4px">ขอบคุณที่ใช้บริการ 🙏</div>
  </div>
</div>`;
  }

  // A4 / Full paper format
  return `<div class="paper paper-a4">
  <div class="ph"><div class="shop"><b>${esc(shop.name || "คลังเหล็ก")}</b>${shop.addr ? esc(shop.addr) + "<br>" : ""}${shop.phone && d.showPhone !== false ? "โทร " + esc(shop.phone) : ""}${shop.tax ? "<br>เลขผู้เสียภาษี " + esc(shop.tax) : ""}</div>
  <div class="dt"><h2>${d.type === "credit" ? "ใบส่งของ / ใบแจ้งหนี้" : DOCNAME[d.type]}</h2><div>เลขที่ ${esc(d.no)}</div><div>วันที่ ${dlong(d.date || d.at)}</div>${d.type === "credit" ? `<div>ครบกำหนดชำระ ${dlong(d.due || d.date)}</div>` : ""}<div>ออกโดย: ${currentRole === "owner" ? "เจ้าของร้าน" : "พนักงาน (ลูกน้อง)"}</div></div></div>
  <div class="cust"><b>ลูกค้า:</b> ${esc(d.customer)}${d.phone && d.showPhone !== false ? " · โทร " + esc(d.phone) : ""}${d.addr ? "<br>" + esc(d.addr) : ""}
  ${d.ful === "ship" ? `<br><b>จัดส่งที่:</b> ${esc(d.shipAddr || "-")}${d.shipPhone && d.showPhone !== false ? " · ผู้รับ โทร " + esc(d.shipPhone) : ""}${d.shipDate ? " · วันที่ส่ง " + dlong(d.shipDate) : ""}` : d.ful === "pickup" ? `<br><b>การรับสินค้า:</b> ลูกค้ามารับเอง` : ""}</div>
  <table><thead><tr><th class="r">#</th><th>รายการ</th><th class="r">จำนวน</th><th class="r">ราคา/หน่วย</th><th class="r">จำนวนเงิน</th></tr></thead><tbody>
  ${d.lines.map((l, i) => `<tr><td class="r">${i + 1}</td><td>${esc(l.name)}</td><td class="r">${fmt(l.qty)} ${esc(l.unit)}</td><td class="r">${money(l.price)}</td><td class="r">${money(l.qty * l.price)}</td></tr>`).join("")}
  </tbody></table>
  <div class="tot"><div class="words">(${bahtText(d.total)})</div>
  <table><tr><td>รวม</td><td class="r">${money(d.sub)}</td></tr>${d.discount ? `<tr><td>ส่วนลด</td><td class="r">-${money(d.discount)}</td></tr>` : ""}<tr><th>ยอดสุทธิ</th><th class="r">${money(d.total)}</th></tr></table></div>
  ${shop.bankName && shop.bankAcc ? `
  <div class="bank-box"><b>ข้อมูลการโอนเงินเข้าบัญชี:</b> ธนาคาร ${esc(shop.bankName)} · เลขที่บัญชี: <b>${esc(shop.bankAcc)}</b> ${shop.bankHolder ? `· ชื่อบัญชี: ${esc(shop.bankHolder)}` : ""}</div>` : ""}
  ${d.note ? `<div class="note">หมายเหตุ: ${esc(d.note)}</div>` : ""}
  ${d.type === "quote" ? `<div class="note">ยืนราคา 7 วันนับจากวันที่ในเอกสาร</div>` : ""}
  ${d.type === "receipt" && d.via ? `<div class="note">ชำระโดย ${esc(d.via)}</div>` : ""}
  ${d.type === "credit" ? `<div class="paystat">ชำระแล้ว ${money(sumPay(d.payments))} บาท · <b>คงค้าง ${money(docOwed(d))} บาท</b>${(d.payments || []).length ? "<br>" + d.payments.map(p => `${dlong(p.date)} ${esc(p.via)} ${money(p.amount)}`).join(" · ") : ""}</div>` : ""}
  ${(d.type === "quote" || (d.type === "credit" && docOwed(d) > 0) || d.type === "receipt") && shop.promptpay ? (() => { const amt = d.type === "credit" ? docOwed(d) : (+d.total || 0), u = qrDataURL(ppPayload(shop.promptpay, amt)); return u ? `<div class="qrbox"><img src="${u}" alt="QR"><div>สแกนจ่ายพร้อมเพย์<br>${money(amt)} บาท</div></div>` : ""; })() : ""}
  ${shop.billNote ? `<div class="note" style="margin-top:8px;font-size:12px;color:#555">เงื่อนไข: ${esc(shop.billNote)}</div>` : ""}
  <div class="sign" style="clear:both"><div>${d.type === "quote" ? "ผู้เสนอราคา" : d.type === "credit" ? "ผู้ส่งของ" : "ผู้รับเงิน / พนักงาน"}</div><div>${d.type === "quote" ? "ผู้อนุมัติสั่งซื้อ" : d.type === "credit" ? "ผู้รับของ" : "ผู้จ่ายเงิน / ลูกค้า"}</div></div>
</div>`;
}

let viewing=null;
function openView(d){
  if(!d) return; viewing=d;
  $("#vPaper").innerHTML=paperHTML(d, "a4"); $("#vConvert").hidden=d.type!=="quote"; $("#vPay").hidden=!(docOwed(d)>0); $("#vShip").hidden=!(d.ful==="ship"&&!d.deliveredAt&&d.type!=="quote");
  $("#vPhone").textContent = d.showPhone===false ? "แสดงเบอร์โทร" : "ซ่อนเบอร์โทร"; $("#vPhoto").hidden=!d.photoId;
  if(!$("#dView").open) $("#dView").showModal();
}
$("#vClose").addEventListener("click",()=>$("#dView").close());
$("#vPay").addEventListener("click",()=>{ const d=viewing; openPay({title:"รับชำระ "+d.no,info:`${d.customer} · คงค้าง ${money(docOwed(d))} บาท`,amount:docOwed(d),onSave:async p=>{
  const payments=[...(d.payments||[]),p]; await store.update("docs/"+d.id,{payments}); openView({...d,payments}); }}); });
$("#vPhone").addEventListener("click",async()=>{
  if(!viewing) return; const show=viewing.showPhone===false;
  try{ await store.update("docs/"+viewing.id,{showPhone:show}); viewing={...viewing,showPhone:show}; openView(viewing); toast(show?"แสดงเบอร์โทรบนเอกสารแล้ว":"ซ่อนเบอร์โทรบนเอกสารแล้ว"); }
  catch(e){ toast("บันทึกไม่สำเร็จ"); }
});
$("#vPhoto").addEventListener("click",()=>{ if(viewing?.photoId){ $("#iBig").src=assetUrl(viewing.photoId); $("#dImg").showModal(); } });

function printDoc(d, format = "a4"){
  if(!d) return;
  $("#printArea").innerHTML = paperHTML(d, format);
  setTimeout(()=>{ try{ window.print(); }catch(e){ toast("พิมพ์จากหน้านี้ไม่ได้ ใช้ปุ่มบันทึกไฟล์แทน"); } }, 50);
}

$("#vPrint").addEventListener("click", () => printDoc(viewing, "a4"));
$("#vPrintSlip")?.addEventListener("click", () => printDoc(viewing, "slip"));

$("#vSave").addEventListener("click",async()=>{
  if(!downloads) return;
  const css=document.getElementById("paperCss").textContent;
  const html=`<!doctype html><html lang="th"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>${esc(viewing.no)}</title><link href="https://fonts.googleapis.com/css2?family=IBM+Plex+Sans+Thai:wght@400;600&display=swap" rel="stylesheet"><style>body{margin:0;background:#fff}${css}</style></head><body>${paperHTML(viewing,"a4")}<script>window.onload=()=>setTimeout(()=>print(),400)<\/script></body></html>`;
  try{ await downloads.save({filename:`${viewing.no} ${viewing.customer}.html`,data:html}); toast("บันทึกไฟล์แล้ว เปิดไฟล์เพื่อพิมพ์หรือส่งต่อ"); }
  catch(e){ if(e&&e.code!=="declined") toast("บันทึกไฟล์ไม่สำเร็จ"); }
});
$("#vConvert").addEventListener("click",()=>{ const d=viewing; $("#dView").close(); openDoc(d); });
$("#vDel").addEventListener("click",async()=>{
  if(!viewing||!confirm(`ลบ ${DOCNAME[viewing.type]} ${viewing.no}?${viewing.cut?" (สต็อกที่ตัดไปแล้วจะไม่ถูกคืน)":""}`)) return;
  try{ await store.del("docs/"+viewing.id); $("#dView").close(); toast("ลบเอกสารแล้ว"); }catch(e){ toast("ลบไม่สำเร็จ"); }
});

/* ===================== roles, PIN authentication, audit ===================== */
let currentRole = localStorage.getItem("khlang_lek_role") || "owner";
let userCap = null, me = "", realOwner = (currentRole === "owner"), staffPreview = false;
const isStaff = () => !realOwner || staffPreview;

function applyRole(){
  realOwner = (currentRole === "owner");
  document.body.classList.toggle("staff", isStaff());
  document.body.classList.toggle("realstaff", !realOwner);
  const roleBtn = $("#roleBtn");
  if(roleBtn){
    const icon = $("#roleIcon"); if(icon) icon.textContent = realOwner ? "👑" : "👤";
    const text = $("#roleText"); if(text) text.textContent = realOwner ? "เจ้าของ" : "ลูกน้อง";
    roleBtn.title = realOwner ? "สิทธิ์: เจ้าของร้าน (แตะเพื่อสลับเป็นลูกน้อง/ล็อก)" : "สิทธิ์: ลูกน้อง (แตะเพื่อใส่รหัสปลดล็อกเจ้าของร้าน)";
  }
  const roleInfo = $("#roleInfo");
  if(roleInfo){
    roleInfo.textContent = realOwner ? (staffPreview ? "เจ้าของ (กำลังดูแบบลูกน้อง: ไม่เห็นต้นทุน/กำไร)" : "เจ้าของร้าน — มีสิทธิ์เต็มทุกฟังก์ชัน") : "ลูกน้อง (แคชเชียร์) — ซ่อนต้นทุนและกำไร";
  }
  if(isStaff() && moneyView === "profit") {
    const cashTab = document.querySelector('.segbar [data-mv=cash]');
    if(cashTab) cashTab.click();
  }
}
$("#staffPrev")?.addEventListener("change", e => { staffPreview = e.target.checked; applyRole(); render(); });

/* ---------- PIN Keypad Controller ---------- */
let pinBuffer = "";
let pinActiveOptions = null;

function updatePinDots(len = 4){
  const dotsContainer = $("#pinDots");
  if(!dotsContainer) return;
  let html = "";
  for(let i = 0; i < len; i++){
    html += `<span class="dot ${i < pinBuffer.length ? 'filled' : ''}"></span>`;
  }
  dotsContainer.innerHTML = html;
}

function openPinPad({ title, desc, avatar = "🔐", hint, maxLen = 4, target = "any", onDone, cancellable = true }){
  pinBuffer = "";
  pinActiveOptions = { title, desc, avatar, hint, maxLen, target, onDone, cancellable };
  if($("#pinAvatar")) $("#pinAvatar").textContent = avatar;
  if($("#pinTitle")) $("#pinTitle").textContent = title || "ใส่รหัส PIN";
  if($("#pinDesc")) $("#pinDesc").textContent = desc || "กรุณากรอกรหัส PIN 4 หลัก";
  if($("#pinError")) $("#pinError").textContent = "";
  const cancelBtn = $("#pinCancelBtn");
  if(cancelBtn) cancelBtn.style.display = cancellable ? "inline-block" : "none";
  if(hint !== undefined && $("#pinHint")) $("#pinHint").innerHTML = hint;
  updatePinDots(maxLen);
  if(!$("#dPin").open) $("#dPin").showModal();
}

function onPinDigit(digit){
  if(!pinActiveOptions) return;
  const max = pinActiveOptions.maxLen || 4;
  if(pinBuffer.length >= max) return;
  pinBuffer += digit;
  updatePinDots(max);
  if($("#pinError")) $("#pinError").textContent = "";
  if(pinBuffer.length >= max){
    setTimeout(verifyEnteredPin, 60);
  }
}

function onPinBackspace(){
  if(!pinActiveOptions || !pinBuffer.length) return;
  pinBuffer = pinBuffer.slice(0, -1);
  updatePinDots(pinActiveOptions.maxLen || 4);
  if($("#pinError")) $("#pinError").textContent = "";
}

function onPinClear(){
  if(!pinActiveOptions) return;
  pinBuffer = "";
  updatePinDots(pinActiveOptions.maxLen || 4);
  if($("#pinError")) $("#pinError").textContent = "";
}

function verifyEnteredPin(){
  if(!pinActiveOptions) return;
  const ownerPin = String(shop.ownerPin || "8888").trim();
  const staffPin = String(shop.staffPin || "1111").trim();
  const entered = pinBuffer;
  const target = pinActiveOptions.target;

  let success = false;
  let resolvedRole = "";

  if(target === "owner"){
    if(entered === ownerPin){ success = true; resolvedRole = "owner"; }
  } else if(target === "staff"){
    if(entered === staffPin){ success = true; resolvedRole = "staff"; }
  } else {
    if(entered === ownerPin){ success = true; resolvedRole = "owner"; }
    else if(entered === staffPin){ success = true; resolvedRole = "staff"; }
  }

  if(success){
    $("#dPin").close();
    const cb = pinActiveOptions.onDone;
    pinActiveOptions = null;
    pinBuffer = "";
    if(cb) cb(resolvedRole);
  } else {
    const dots = $("#pinDots");
    if(dots){
      dots.classList.add("shake");
      setTimeout(() => dots.classList.remove("shake"), 400);
    }
    if($("#pinError")) $("#pinError").textContent = "รหัส PIN ไม่ถูกต้อง ลองอีกครั้ง";
    setTimeout(() => {
      pinBuffer = "";
      if(pinActiveOptions) updatePinDots(pinActiveOptions.maxLen || 4);
    }, 450);
  }
}

document.querySelectorAll(".pin-pad .pad-btn[data-key]").forEach(b => {
  b.addEventListener("click", () => onPinDigit(b.dataset.key));
});
$("#pinBack")?.addEventListener("click", onPinBackspace);
$("#pinClear")?.addEventListener("click", onPinClear);
$("#pinCancelBtn")?.addEventListener("click", () => {
  if(pinActiveOptions && pinActiveOptions.cancellable){
    $("#dPin").close();
    pinActiveOptions = null;
    pinBuffer = "";
  }
});
window.addEventListener("keydown", e => {
  const pinModal = $("#dPin");
  if(!pinModal || !pinModal.open) return;
  if(e.key >= "0" && e.key <= "9"){
    e.preventDefault();
    onPinDigit(e.key);
  } else if(e.key === "Backspace"){
    e.preventDefault();
    onPinBackspace();
  } else if(e.key === "Escape"){
    if(pinActiveOptions && pinActiveOptions.cancellable){
      pinModal.close();
      pinActiveOptions = null;
      pinBuffer = "";
    }
  }
});

function handleSwitchUser(){
  if(currentRole === "staff"){
    openPinPad({
      title: "ปลดล็อกสิทธิ์เจ้าของร้าน",
      desc: "กรุณาใส่รหัส PIN เจ้าของร้านเพื่อดูต้นทุนและกำไร",
      avatar: "👑",
      target: "owner",
      hint: `รหัสเริ่มต้นเจ้าของร้านคือ <b>${shop.ownerPin || "8888"}</b>`,
      onDone: () => {
        currentRole = "owner";
        localStorage.setItem("khlang_lek_role", "owner");
        applyRole();
        render();
        toast("ปลดล็อกสิทธิ์เจ้าของร้านเรียบร้อย 👑");
      }
    });
  } else {
    if(confirm("ต้องการสลับเป็นโหมด 'ลูกน้อง' (ซ่อนต้นทุนและกำไร) ใช่หรือไม่?")){
      currentRole = "staff";
      localStorage.setItem("khlang_lek_role", "staff");
      applyRole();
      render();
      toast("สลับเป็นโหมดลูกน้องแล้ว 👤");
    }
  }
}

$("#roleBtn")?.addEventListener("click", handleSwitchUser);
$("#btnSwitchPin")?.addEventListener("click", handleSwitchUser);
$("#switchUserRow")?.addEventListener("click", e => {
  if(e.target.id !== "btnSwitchPin") handleSwitchUser();
});
const names={}; let nameReq=null;
function nameOf(id){
  if(!id||!userCap) return ""; if(id===me) return "คุณ";
  if(names[id]!==undefined) return names[id];
  names[id]=""; clearTimeout(nameReq);
  nameReq=setTimeout(async()=>{ try{ const ids=Object.keys(names).filter(k=>names[k]===""); const ps=await userCap.profiles(ids);
    ids.forEach(k=>names[k]=(ps&&ps[k]&&ps[k].name)||"ผู้ใช้อื่น"); renderLog(); renderDocs(); }catch(e){} },50);
  return "";
}
const byTxt=o=>{ const n=nameOf(o&&o.by); return n?` · โดย ${n}`:""; };
const AUDIT=/^(docs|buys|cash|moves|closes|counts|prices|orders|customers)\//;
function wrapAudit(){ const s=store.set.bind(store); store.set=(p,v)=> s(p, me&&AUDIT.test(p)&&v&&!v.by ? {...v,by:me} : v); }

/* payments */
const sumPay=a=>(Array.isArray(a)?a:[]).reduce((s,p)=>s+(+p.amount||0),0);
const docOwed=d=> d.type==="credit" ? Math.max(0,Math.round(((+d.total||0)-sumPay(d.payments))*100)/100) : 0;
const buyOwed=b=> b.payStatus==="credit" ? Math.max(0,Math.round(((+b.total||0)-sumPay(b.payments))*100)/100) : 0;
const DAY=864e5, todayStart=()=>{ const n=new Date(); return new Date(n.getFullYear(),n.getMonth(),n.getDate()).getTime(); };
const daysFrom=t=>{ const d=new Date(t); return Math.round((todayStart()-new Date(d.getFullYear(),d.getMonth(),d.getDate()).getTime())/DAY); };
const fy=$("#fPay"); let payCb=null;
function openPay({title,info,amount,qr,onSave}){
  $("#yTitle").textContent=title; $("#yInfo").textContent=info||""; fy.amount.value=amount||""; fy.amount.max=amount||""; fy.date.value=today(); fy.via.value="เงินสด";
  payCb=onSave; $("#yQR").innerHTML=""; fy._qr=qr; showPayQR(); $("#dPay").showModal();
}
function showPayQR(){ $("#yQR").innerHTML = fy.via.value==="โอน"&&shop.promptpay ? qrHTML(+fy.amount.value||0) : ""; }
fy.addEventListener("input",showPayQR); fy.addEventListener("change",showPayQR);
fy.addEventListener("submit",async e=>{
  if(e.submitter&&e.submitter.value!=="save") return; e.preventDefault(); if(!fy.reportValidity()) return;
  const amt=+fy.amount.value||0; if(!(amt>0)) return toast("ใส่จำนวนเงิน");
  try{ await payCb({amount:Math.round(amt*100)/100,via:fy.via.value,date:new Date(fy.date.value+"T12:00").getTime(),at:Date.now(),...(me?{by:me}:{})}); $("#dPay").close(); toast(`บันทึก ${money(amt)} บาทแล้ว`); }
  catch(err){ toast("บันทึกไม่สำเร็จ ลองอีกครั้ง"); }
});

/* PromptPay EMVCo QR */
function crc16(s){ let c=0xFFFF; for(let i=0;i<s.length;i++){ c^=s.charCodeAt(i)<<8; for(let j=0;j<8;j++) c=(c&0x8000)?((c<<1)^0x1021)&0xFFFF:(c<<1)&0xFFFF; } return c.toString(16).toUpperCase().padStart(4,"0"); }
function ppPayload(id,amount){
  const f=(t,v)=>t+String(v.length).padStart(2,"0")+v; const d=String(id||"").replace(/\D/g,"");
  let acc; if(d.length>=13) acc=f("02",d.slice(0,13)); else acc=f("01",("0066"+d.replace(/^0/,"")).padStart(13,"0"));
  let p=f("00","01")+f("01",amount>0?"12":"11")+f("29",f("00","A000000677010111")+acc)+f("53","764")+(amount>0?f("54",amount.toFixed(2)):"")+f("58","TH")+"6304";
  return p+crc16(p);
}
function qrDataURL(text){ try{ if(!window.qrcode) return ""; const q=qrcode(0,"M"); q.addData(text); q.make(); return q.createDataURL(6,2); }catch(e){ return ""; } }
function qrHTML(amount){
  if(!shop.promptpay) return ""; const url=qrDataURL(ppPayload(shop.promptpay,amount)); if(!url) return "";
  return `<div class="qr"><img src="${url}" alt="QR พร้อมเพย์"><small>พร้อมเพย์ ${esc(shop.promptpay)}${amount>0?` · ${money(amount)} บาท`:""}</small></div>`;
}

/* ===================== sales docs: commit, credit, receivables ===================== */
async function commitDoc(o){
  // o: {type,dateStr,customer,phone,addr,lines,discount,note,cut,showPhone,via,dueDays,photo}
  const type=o.type, cut=o.cut&&(type==="receipt"||type==="credit"), sh=o.ship||{};
  const ls=sh.shipFee>0?[...o.lines,{itemId:"",name:"ค่าจัดส่ง",qty:1,unit:"เที่ยว",price:+sh.shipFee,cost:0}]:o.lines;
  const sim={}, ops=[];
  if(cut){ for(const l of ls){ const base=items.find(i=>i.id===l.itemId); if(!base) continue;
    const it=sim[base.id]||(sim[base.id]={...base,rem:remOf(base).map(r=>({...r}))});
    if(l.cut){ const spec=l.cut.mode==="area"?{w:l.cut.w,l:l.cut.l}:{x:l.cut.x}; const p=planCut(it,spec,l.qty,l.cut.useOld!==false);
      if(p.error) throw new Error(`${it.name}: ${p.error}`);
      ops.push({kind:"cut",id:it.id,p,l}); it.qty=p.full; it.rem=p.keep.map(({old,...r})=>r); }
    else { if(l.qty>(+it.qty||0)) throw new Error(`${it.name} ไม่พอ เหลือ ${fmt(it.qty)} ${it.unit}`); it.qty=(+it.qty||0)-l.qty; ops.push({kind:"qty",id:it.id,l}); } } }
  const sub=sumLines(ls), discount=+o.discount||0, dateMs=new Date(o.dateStr+"T12:00").getTime(), total=Math.max(0,Math.round((sub-discount)*100)/100);
  const v={type,no:nextNo(type,o.dateStr),date:dateMs,at:Date.now(),customer:o.customer,phone:o.phone||"",addr:o.addr||"",
    lines:ls.map(l=>({itemId:l.itemId||"",name:String(l.name).trim(),qty:+l.qty,unit:l.unit||"",price:+l.price,cost:+l.cost||(l.cut?0:(items.find(i=>i.id===l.itemId)?.cost||0)),...(l.cut?{cut:l.cut}:{})})),
    sub,discount,total,note:o.note||"",cut,showPhone:o.showPhone!==false};
  if(type==="receipt") v.via=o.via||"เงินสด";
  if(sh.ful){ Object.assign(v,sh); v.shipFee=0; }
  try{ const cid=await upsertCust({name:v.customer,phone:v.phone,addr:v.addr,ship:v.shipAddr}); if(cid) v.custId=cid; }catch(e){}
  if(type==="credit"){ v.due=dateMs+(+o.dueDays||0)*DAY; v.payments=[]; }
  if(cut){ const at=Date.now(), live={};
    for(const op of ops){ const base=live[op.id]||items.find(i=>i.id===op.id);
      if(op.kind==="cut"){ await applyCut(base,op.p,op.l.qty,`${op.l.unit||"ท่อน"} ${op.l.cut.label}`,`ขาย ${v.no} ${v.customer}`,at); live[op.id]={...base,qty:op.p.full,rem:op.p.keep}; }
      else { const nq=(+base.qty||0)-op.l.qty; await store.update("items/"+op.id,{qty:nq,updatedAt:at});
        await addMove({type:"out",qty:op.l.qty,itemId:op.id,name:base.name,note:`ขาย ${v.no} ${v.customer}`,at}); live[op.id]={...base,qty:nq}; } } }
  if(o.photo&&assets){ try{ const up=await assets.upload(o.photo,{type:"image/jpeg"}); v.photoId=up.id; }catch(err){ toast("แนบรูปไม่สำเร็จ บันทึกเอกสารโดยไม่มีรูป"); } }
  const id=uid(); await store.set("docs/"+id,v); return {id,...v,...(me?{by:me}:{})};
}
function typeChanged(){ const t=fd.type.value; $("#cutWrap").hidden=t==="quote"; $("#viaWrap").hidden=t!=="receipt"; $("#dueWrap").hidden=t!=="credit"; }

/* receivables */
const custKey=d=>(String(d.customer||"").trim().toLowerCase()+"|"+String(d.phone||"").replace(/\D/g,""));
function debtors(){
  const m={};
  docs.filter(d=>d.type==="credit"&&docOwed(d)>0).forEach(d=>{ const k=String(d.customer||"").trim().toLowerCase(); const x=m[k]||(m[k]={name:d.customer,phone:d.phone,owed:0,n:0,oldestDue:Infinity,docs:[]});
    x.owed+=docOwed(d); x.n++; x.oldestDue=Math.min(x.oldestDue,d.due||d.date); x.docs.push(d); if(d.phone) x.phone=d.phone; });
  return Object.values(m).sort((a,b)=>a.oldestDue-b.oldestDue);
}
function docStatusTag(d){
  if(d.type!=="credit") return "";
  const o=docOwed(d); if(o<=0) return `<span class="tag ok">ชำระครบ</span>`;
  const od=daysFrom(d.due||d.date);
  return `<span class="tag ${od>0?"due":"part"}">ค้าง ${fmt(o)}${od>0?` · เกิน ${od} วัน`:""}</span>`;
}
let custCur=null, custRec=null;
const fcu=$("#fCu");
function openCust(name,isNew){
  const key=ckey(name); custCur=key; custRec=isNew?null:findCust(name);
  const match=d=>(custRec&&d.custId===custRec.id)||(key&&ckey(d.customer)===key);
  const all=isNew?[]:docs.filter(match).sort((a,b)=>(b.date||b.at)-(a.date||a.at));
  const ords=isNew?[]:orders.filter(match);
  const sales=all.filter(d=>d.type!=="quote"), owed=all.reduce((s,d)=>s+docOwed(d),0);
  const src=custRec||{name:all[0]?.customer||name||"",phone:all.find(d=>d.phone)?.phone||"",addr:all.find(d=>d.addr)?.addr||"",ship:all.find(d=>d.shipAddr)?.shipAddr||"",kind:"ทั่วไป"};
  $("#cuName").textContent=isNew?"เพิ่มลูกค้าใหม่":src.name; $("#cuInfo").textContent=!isNew&&!custRec?"ยังไม่ได้บันทึกในฐานข้อมูลลูกค้า — กดบันทึกเพื่อเก็บไว้":"";
  for(const k of ["name","phone","addr","ship","taxId","line","note"]) fcu[k].value=src[k]||""; fcu.kind.value=src.kind||"ทั่วไป"; fcu.creditDays.value=src.creditDays||"";
  $("#cuSum").hidden=!!isNew; $("#cuDel").hidden=!custRec; $("#cuNewOrd").hidden=$("#cuNewDoc").hidden=!!isNew;
  $("#cuSum").innerHTML=`<div><span>ซื้อทั้งหมด (${sales.length} ใบ)</span><b class="num">${fmt(sales.reduce((s,d)=>s+(+d.total||0),0))}</b></div>
    <div><span>ค้างชำระ</span><b class="num ${owed>0?"neg":""}">${fmt(owed)}</b></div>`;
  $("#cuPay").hidden=!(owed>0);
  $("#cuOrders").innerHTML=ords.length?renderOrders(ords):`<p class="hint">ไม่มีออเดอร์</p>`;
  $("#cuDocs").innerHTML=all.map(d=>`<button class="tx" data-doc="${esc(d.id)}"><span>${DOCNAME[d.type]} ${esc(d.no)} ${docStatusTag(d)}<br><small>${dlong(d.date||d.at)}${d.type==="credit"?" · ครบกำหนด "+dlong(d.due||d.date):""}</small></span><span class="num">${fmt(d.total)}</span></button>`).join("")||`<p class="hint">ไม่มีเอกสาร</p>`;
  if(!$("#dCust").open) $("#dCust").showModal();
}
fcu.addEventListener("submit",async e=>{
  e.preventDefault(); if(!fcu.reportValidity()) return;
  const v={name:fcu.name.value.trim(),phone:fcu.phone.value.trim(),addr:fcu.addr.value.trim(),ship:fcu.ship.value.trim(),kind:fcu.kind.value,creditDays:+fcu.creditDays.value||0,
    taxId:fcu.taxId.value.trim(),line:fcu.line.value.trim(),note:fcu.note.value.trim(),updatedAt:Date.now()};
  const dup=findCust(v.name); if(dup&&(!custRec||dup.id!==custRec.id)) return toast("มีชื่อลูกค้านี้อยู่แล้ว");
  try{ if(custRec) await store.update("customers/"+custRec.id,v); else { const id=uid(); await store.set("customers/"+id,{...v,at:Date.now()}); custRec={id,...v}; }
    toast("บันทึกข้อมูลลูกค้าแล้ว"); custRec={...custRec,...v}; $("#cuName").textContent=v.name; $("#cuInfo").textContent=""; $("#cuDel").hidden=false; $("#cuNewOrd").hidden=$("#cuNewDoc").hidden=false;
  }catch(err){ toast("บันทึกไม่สำเร็จ"); }
});
$("#cuDel").addEventListener("click",async()=>{ if(!custRec||!confirm(`ลบ "${custRec.name}" ออกจากรายชื่อลูกค้า? (เอกสารเก่ายังอยู่)`)) return;
  try{ await store.del("customers/"+custRec.id); $("#dCust").close(); toast("ลบรายชื่อแล้ว"); }catch(e){ toast("ลบไม่สำเร็จ"); } });
$("#cuNewOrd").addEventListener("click",()=>{ const n=fcu.name.value; $("#dCust").close(); openOrder(null); fo.customer.value=n; custAutofill(fo); });
$("#cuNewDoc").addEventListener("click",()=>{ const n=fcu.name.value; $("#dCust").close(); openDoc(); fd.customer.value=n; custAutofill(fd); });
$("#cuOrders").addEventListener("click",e=>{ const o=e.target.closest("[data-ord]"); if(o){ $("#dCust").close(); openOrder(orders.find(x=>x.id===o.dataset.ord)); } });
$("#cuClose").addEventListener("click",()=>$("#dCust").close());
$("#cuDocs").addEventListener("click",e=>{ const b=e.target.closest("[data-doc]"); if(b){ openView(docs.find(d=>d.id===b.dataset.doc)); } });
$("#cuPay").addEventListener("click",()=>{
  const open=docs.filter(d=>d.type==="credit"&&docOwed(d)>0&&String(d.customer||"").trim().toLowerCase()===custCur).sort((a,b)=>(a.due||a.date)-(b.due||b.date));
  const owed=open.reduce((s,d)=>s+docOwed(d),0);
  openPay({title:"รับชำระจาก "+(open[0]?.customer||""),info:`ค้าง ${open.length} ใบ รวม ${money(owed)} บาท · ตัดใบที่ครบกำหนดก่อน`,amount:owed,onSave:async p=>{
    let left=p.amount;
    for(const d of open){ if(left<=0) break; const take=Math.min(left,docOwed(d)); left=Math.round((left-take)*100)/100;
      await store.update("docs/"+d.id,{payments:[...(d.payments||[]),{...p,amount:take}]}); }
    setTimeout(()=>openCust(custCur),300);
  }});
});

/* ===================== payables (buy on credit) ===================== */
function payFields(form){
  const credit=form.payStatus.value==="credit";
  form.querySelector(".payVia").hidden=credit; form.querySelector(".payDue").hidden=!credit;
  if(credit&&!form.due.value) form.due.value=isoDate(Date.now()+30*DAY);
}
[fb,fr].forEach(F=>F.payStatus.addEventListener("change",()=>payFields(F)));
function payData(form){
  if(form.payStatus.value==="credit") return {payStatus:"credit",due:new Date((form.due.value||today())+"T12:00").getTime(),payments:[]};
  return {payStatus:"paid",payVia:form.payVia.value};
}
function buyPayInfo(b){
  if(!b||b.payStatus!=="credit"){ $("#bPayInfo").innerHTML=""; $("#bPayBtn").hidden=true; return; }
  const o=buyOwed(b), od=daysFrom(b.due||b.date);
  $("#bPayInfo").innerHTML=`<div class="msg ${o>0&&od>0?"bad":o>0?"info":"ok"}">${o>0?`ค้างจ่าย <b>${money(o)}</b> บาท · ครบกำหนด ${dlong(b.due)}${od>0?` (เกิน ${od} วัน)`:""}`:"จ่ายครบแล้ว"}${(b.payments||[]).length?"<br><small>"+b.payments.map(p=>`${dlong(p.date)} ${esc(p.via)} ${money(p.amount)}`).join(" · ")+"</small>":""}</div>`;
  $("#bPayBtn").hidden=!(o>0);
}
$("#bPayBtn").addEventListener("click",()=>{ const b=editingBuy; openPay({title:"จ่ายเงิน "+b.seller,info:`ค้าง ${money(buyOwed(b))} บาท`,amount:buyOwed(b),onSave:async p=>{
  const payments=[...(b.payments||[]),p]; await store.update("buys/"+b.id,{payments}); editingBuy={...b,payments}; buyPayInfo(editingBuy); }}); });
const buyTag=b=>{ if(b.payStatus!=="credit") return ""; const o=buyOwed(b); if(o<=0) return `<span class="tag ok">จ่ายครบ</span>`; const od=daysFrom(b.due||b.date); return `<span class="tag ${od>0?"due":"part"}">ค้างจ่าย ${fmt(o)}${od>0?` เกิน ${od} วัน`:` · ${dlong(b.due)}`}</span>`; };

/* ===================== alerts ===================== */
function dueAlerts(){
  const soon=todayStart()+7*DAY;
  const pay=buys.filter(b=>buyOwed(b)>0&&(b.due||0)<soon), recv=docs.filter(d=>docOwed(d)>0&&(d.due||0)<todayStart());
  return {pay,recv};
}
function renderAlerts(){
  const {pay,recv}=dueAlerts(); let h="";
  if(pay.length){ const t=pay.reduce((s,b)=>s+buyOwed(b),0); h+=`<div class="msg bad alert">⚠ บิลซื้อครบกำหนดใน 7 วัน/เกินกำหนด ${pay.length} ใบ รวม <b>${money(t)}</b> บาท<br><small>${pay.slice(0,4).map(b=>`${esc(b.seller)} ${money(buyOwed(b))} (${dlong(b.due)})`).join(" · ")}</small></div>`; }
  if(recv.length){ const t=recv.reduce((s,d)=>s+docOwed(d),0); h+=`<div class="msg bad alert">⚠ ลูกหนี้เกินกำหนด ${recv.length} ใบ รวม <b>${money(t)}</b> บาท — ดูที่ ขาย › ลูกหนี้</div>`; }
  $("#alerts").innerHTML=h;
  document.querySelector('.tabbar [data-tab=money]').classList.toggle("dot",!!(pay.length||recv.length));
}

/* ===================== quick sale ===================== */
let qcart=[];
function openQuick(){ qcart=[]; $("#qsSearch").value=""; $("#qsCust").value=""; $("#qsGot").value=""; document.querySelector("[name=qsPay][value=เงินสด]").checked=true; drawQS(); $("#dQuick").showModal(); }
$("#qsOpen").addEventListener("click",openQuick);
$("#docNew").addEventListener("click",()=>openDoc());
function drawQS(){
  const q=$("#qsSearch").value.trim().toLowerCase();
  const res=items.filter(i=>itemUnits(i)>0&&(!q||[i.name,i.spec].join(" ").toLowerCase().includes(q))).sort((a,b)=>String(a.name).localeCompare(String(b.name),"th")).slice(0,30);
  $("#qsResults").innerHTML=res.map(i=>`<div class="qsr"><span>${esc(i.name)}<small>${esc(i.spec)} · มี ${fmt(i.qty)} ${esc(i.unit)}${remOf(i).length?` + เหลือ ${remOf(i).length}`:""} · ${+i.price?money(i.price):"ไม่มีราคา"}</small></span>
    <button class="btn" data-qadd="${esc(i.id)}">+1</button>${cutMode(i)?`<button class="btn" data-qcut="${esc(i.id)}">✂</button>`:"<span></span>"}</div>`).join("")||`<p class="hint">ไม่พบสินค้า</p>`;
  $("#qsCart").innerHTML=qcart.length?qcart.map((l,i)=>`<div class="cart" data-i="${i}"><span>${esc(l.name)}</span>
    <input data-k="qty" type="number" inputmode="decimal" min="0" step="any" value="${l.qty}" aria-label="จำนวน"><input data-k="price" type="number" inputmode="decimal" min="0" step="any" value="${l.price}" aria-label="ราคา">
    <button class="rm" data-rm="${i}" aria-label="ลบ">×</button></div>`).join(""):`<p class="hint">แตะ +1 หรือ ✂ เพื่อเพิ่มสินค้า</p>`;
  qsTotals();
}
function qsTotals(){
  const tot=sumLines(qcart), cost=qcart.reduce((s,l)=>s+(+l.qty||0)*(+l.cost||0),0), pay=document.querySelector("[name=qsPay]:checked").value;
  $("#qsTotal").textContent=money(tot); $("#qsProfit").textContent=money(tot-cost);
  $("#qsCashRow").hidden=pay!=="เงินสด"; const got=+$("#qsGot").value||0;
  $("#qsChange").innerHTML = pay==="เงินสด"&&got ? (got>=tot?`ทอนเงิน <b class="num">${money(got-tot)}</b> บาท`:`<b class="neg">รับเงินยังไม่พอ ขาด ${money(tot-got)}</b>`) : pay==="credit"?"ขายเชื่อ: ต้องใส่ชื่อลูกค้า · ครบกำหนด 30 วัน":"";
  $("#qsQR").innerHTML = pay==="โอน"&&tot>0 ? (qrHTML(tot)||`<p class="hint">ตั้งเลขพร้อมเพย์ที่ เพิ่มเติม › ข้อมูลร้าน เพื่อแสดง QR</p>`) : "";
}
$("#qsSearch").addEventListener("input",drawQS);
$("#dQuick").addEventListener("change",e=>{ if(e.target.name==="qsPay") qsTotals(); });
$("#qsGot").addEventListener("input",qsTotals);
$("#qsResults").addEventListener("click",e=>{
  const a=e.target.dataset.qadd, c=e.target.dataset.qcut;
  if(a){ const it=items.find(i=>i.id===a); const ex=qcart.find(l=>l.itemId===a&&!l.cut);
    if(ex) ex.qty++; else qcart.push({itemId:it.id,name:[it.name,it.spec].filter(Boolean).join(" "),qty:1,unit:it.unit,price:+it.price||0,cost:+it.cost||0}); drawQS(); }
  if(c) openCut(items.find(i=>i.id===c),"quick");
});
$("#qsCart").addEventListener("input",e=>{ const r=e.target.closest("[data-i]"); if(!r) return; qcart[+r.dataset.i][e.target.dataset.k]=+e.target.value||0; qsTotals(); });
$("#qsCart").addEventListener("click",e=>{ const r=e.target.dataset.rm; if(r!==undefined){ qcart.splice(+r,1); drawQS(); } });
$("#qsClose").addEventListener("click",()=>$("#dQuick").close());
$("#qsSave").addEventListener("click",async()=>{
  const ls=qcart.filter(l=>l.qty>0); if(!ls.length) return toast("ยังไม่มีสินค้าในตะกร้า");
  const pay=document.querySelector("[name=qsPay]:checked").value, cust=$("#qsCust").value.trim();
  if(pay==="credit"&&!cust) return toast("ขายเชื่อต้องใส่ชื่อลูกค้า");
  const btn=$("#qsSave"); btn.disabled=true;
  try{
    const d=await commitDoc({type:pay==="credit"?"credit":"receipt",dateStr:today(),customer:cust||"ลูกค้าหน้าร้าน",lines:ls,discount:0,cut:true,showPhone:true,via:pay==="credit"?"":pay,dueDays:30});
    $("#dQuick").close(); toast(`ขายแล้ว ${money(d.total)} บาท (${d.no})`); openView(d);
  }catch(err){ toast(err&&err.message?err.message:"บันทึกไม่สำเร็จ"); }
  finally{ btn.disabled=false; }
});

/* ===================== daily close ===================== */
let closes=[], closeDate=null;
function renderClose(){
  const ds=closeDate||today(), a=new Date(ds+"T00:00").getTime(), b=a+DAY;
  const L=ledger().filter(x=>x.t>=a&&x.t<b);
  const sum=(dir,via)=>L.filter(x=>x.dir===dir&&(via?x.via===via:true)).reduce((s,x)=>s+x.amt,0);
  const creditSales=docs.filter(d=>d.type==="credit"&&(d.date||d.at)>=a&&(d.date||d.at)<b).reduce((s,d)=>s+(+d.total||0),0);
  const prev=closes.filter(c=>c.date<ds).sort((x,y)=>y.date.localeCompare(x.date))[0], saved=closes.find(c=>c.date===ds);
  const float=saved?saved.float:(prev?prev.keep??prev.counted:0);
  const cin=sum("in","เงินสด"), cout=sum("out","เงินสด"), exp=Math.round((float+cin-cout)*100)/100;
  $("#closeBody").innerHTML=`
  <label>วันที่<input type="date" id="clDate" value="${ds}"></label>
  <div class="pgrid" style="margin-top:10px">
    <div><span>เงินสดรับ</span><b class="num pos">${fmt(cin)}</b></div><div><span>เงินโอนรับ</span><b class="num pos">${fmt(sum("in","โอน"))}</b></div>
    <div><span>เงินสดจ่าย</span><b class="num neg">${fmt(cout)}</b></div><div><span>เงินโอนจ่าย</span><b class="num neg">${fmt(sum("out","โอน"))}</b></div>
    <div><span>ขายเชื่อวันนี้ (ยังไม่ได้เงิน)</span><b class="num">${fmt(creditSales)}</b></div><div><span>รายการทั้งหมด</span><b class="num">${L.length}</b></div>
  </div>
  <div class="grid">
    <label>เงินทอนตั้งต้น (เช้า)<input id="clFloat" type="number" inputmode="decimal" step="any" value="${float}"></label>
    <div class="result" style="margin-top:0"><div><span>ควรมีเงินสด</span><b class="num" id="clExp">${money(exp)}</b></div></div>
    <label>นับเงินสดได้จริง<input id="clCount" type="number" inputmode="decimal" step="any" value="${saved?saved.counted:""}"></label>
    <label>เก็บไว้เป็นเงินทอนพรุ่งนี้<input id="clKeep" type="number" inputmode="decimal" step="any" value="${saved?saved.keep??"":""}"></label>
    <label class="full">หมายเหตุ<input id="clNote" maxlength="120" value="${esc(saved?.note||"")}"></label>
  </div>
  <div id="clDiff"></div>
  <button class="btn primary" id="clSave" style="width:100%;margin-top:10px">${saved?"แก้ไขการปิดยอด":"บันทึกปิดยอดวันนี้"}</button>
  <h4 class="sec">ปิดยอดย้อนหลัง</h4>
  ${closes.slice(0,15).map(c=>`<div class="mrow"><span>${dlong(new Date(c.date+"T12:00").getTime())}<br><small>ควรมี ${money(c.expected)} · นับได้ ${money(c.counted)}${byTxt(c)}</small></span><span class="num ${Math.abs(c.diff)>0.5?"neg":"pos"}">${c.diff>0?"+":""}${money(c.diff)}</span></div>`).join("")||`<p class="hint">ยังไม่เคยปิดยอด</p>`}`;
  clDiff();
}
function clDiff(){
  const ds=$("#clDate").value, a=new Date(ds+"T00:00").getTime(), b=a+DAY, L=ledger().filter(x=>x.t>=a&&x.t<b&&x.via==="เงินสด");
  const exp=(+$("#clFloat").value||0)+L.filter(x=>x.dir==="in").reduce((s,x)=>s+x.amt,0)-L.filter(x=>x.dir==="out").reduce((s,x)=>s+x.amt,0);
  $("#clExp").textContent=money(exp); const c=$("#clCount").value;
  $("#clDiff").innerHTML = c==="" ? "" : (()=>{ const d=Math.round(((+c)-exp)*100)/100; return `<div class="msg ${Math.abs(d)<0.5?"ok":"bad"}">${Math.abs(d)<0.5?"✓ เงินสดตรง":d>0?`เงินเกิน ${money(d)} บาท`:`เงินขาด ${money(-d)} บาท`}</div>`; })();
  return exp;
}
$("#closeBody").addEventListener("input",e=>{ if(e.target.id==="clDate"){ closeDate=e.target.value; renderClose(); } else clDiff(); });
$("#closeBody").addEventListener("click",async e=>{
  if(e.target.id!=="clSave") return; const exp=clDiff(), c=$("#clCount").value; if(c==="") return toast("ใส่เงินสดที่นับได้");
  const v={date:$("#clDate").value,float:+$("#clFloat").value||0,expected:Math.round(exp*100)/100,counted:+c,diff:Math.round(((+c)-exp)*100)/100,keep:$("#clKeep").value===""?+c:+$("#clKeep").value,note:$("#clNote").value.trim(),at:Date.now()};
  try{ await store.set("closes/"+v.date,v); toast("บันทึกปิดยอดแล้ว"); }catch(err){ toast("บันทึกไม่สำเร็จ"); }
});

/* ===================== reports ===================== */
let rpPeriod="30 วัน";
function renderReport(){
  $("#rpChips").innerHTML=["30 วัน","90 วัน","ปีนี้"].map(c=>`<button aria-pressed="${c===rpPeriod}" data-rp="${c}">${c}</button>`).join("");
  const now=Date.now(), a= rpPeriod==="ปีนี้" ? new Date(new Date().getFullYear(),0,1).getTime() : now-(rpPeriod==="30 วัน"?30:90)*DAY;
  const sales=docs.filter(d=>d.type==="receipt"||d.type==="credit"), agg={}, last={};
  const eq=l=>{ const it=items.find(i=>i.id===l.itemId); if(!l.cut||!it) return +l.qty||0;
    const r=l.cut.mode==="area"?(l.cut.w*l.cut.l)/(sheetWL(it)[0]*sheetWL(it)[1]):l.cut.x/fullLen(it); return (+l.qty||0)*r; };
  sales.forEach(d=>{ const t=d.date||d.at; d.lines.forEach(l=>{ const k=l.itemId||("n:"+l.name);
    if(l.itemId) last[l.itemId]=Math.max(last[l.itemId]||0,t);
    if(t<a) return; const it=items.find(i=>i.id===l.itemId); const x=agg[k]||(agg[k]={name:(it||{}).name||l.name,qty:0,rev:0,prof:0,unit:it?it.unit:l.unit});
    x.qty+=eq(l); x.rev+=(+l.qty||0)*(+l.price||0); x.prof+=(+l.qty||0)*((+l.price||0)-(+l.cost||0)); }); });
  const best=Object.values(agg).sort((x,y)=>y.rev-x.rev).slice(0,15);
  const slow=items.filter(i=>itemUnits(i)>0&&i.unit!=="กก.").map(i=>({...i,days:last[i.id]?Math.floor((now-last[i.id])/DAY):null,val:itemVal(i,isStaff()?"price":"cost")||itemVal(i)}))
    .filter(i=>i.days===null||i.days>=90).sort((x,y)=>y.val-x.val);
  const sold30={}; sales.forEach(d=>{ if((d.date||d.at)<now-30*DAY) return; d.lines.forEach(l=>{ if(l.itemId&&!l.cut) sold30[l.itemId]=(sold30[l.itemId]||0)+(+l.qty||0); }); });
  const reorder=items.filter(i=>isLow(i)||(sold30[i.id]&&sold30[i.id]>=(+i.qty||0))).map(i=>({...i,s30:sold30[i.id]||0}));
  $("#reportBody").innerHTML=`
  <h4 class="sec">ขายดี (${rpPeriod}) ตามยอดขาย</h4>
  ${best.length?best.map((x,i)=>`<div class="mrow"><span>${i+1}. ${esc(x.name)}<br><small>ขาย ${fmt(Math.round(x.qty*100)/100)} ${esc(x.unit)}</small></span><span class="num">${fmt(x.rev)}<br><small class="owner">กำไร ${fmt(x.prof)}</small></span></div>`).join(""):`<p class="hint">ยังไม่มียอดขายในช่วงนี้</p>`}
  <h4 class="sec">ควรสั่งเพิ่ม</h4>
  ${reorder.length?reorder.map(i=>`<div class="mrow"><span>${esc(i.name)}<br><small>เหลือ ${fmt(i.qty)} ${esc(i.unit)}${i.s30?` · 30 วันขายไป ${fmt(i.s30)}`:""}${isLow(i)?" · ต่ำกว่าจุดเตือน":""}</small></span><span class="num neg">${fmt(i.qty)}</span></div>`).join(""):`<p class="hint">ยังไม่มีสินค้าที่ต้องสั่งเพิ่ม</p>`}
  <h4 class="sec">ค้างนาน ไม่มีขาย 90 วันขึ้นไป (จมทุน)</h4>
  ${slow.length?`<p class="hint">รวม${isStaff()?"มูลค่า":"ทุนจม"} ${money(slow.reduce((s,i)=>s+i.val,0))} บาท</p>`+slow.slice(0,20).map(i=>`<div class="mrow"><span>${esc(i.name)} <small>${esc(i.spec)}</small><br><small>${i.days===null?"ยังไม่เคยขายในระบบ":`ขายล่าสุด ${i.days} วันก่อน`} · ${fmt(itemUnits(i))} ${esc(i.unit)}</small></span><span class="num">${fmt(i.val)}</span></div>`).join(""):`<p class="hint">ไม่มีสินค้าค้างนาน</p>`}`;
}
$("#rpChips").addEventListener("click",e=>{ const b=e.target.closest("[data-rp]"); if(b){ rpPeriod=b.dataset.rp; renderReport(); } });

/* ===================== category price update ===================== */
let priceHist=[];
function pcMatch(){ const c=$("#pcCat").value, q=$("#pcFilter").value.trim().toLowerCase();
  return items.filter(i=>(c==="ทุกหมวด"||i.cat===c)&&i.unit!=="กก."&&(!q||[i.name,i.spec].join(" ").toLowerCase().includes(q))); }
function pcNew(i){ const k=+$("#pcKg").value||0, r=+$("#pcRound").value; if(!(k>0)||!(+i.kg>0)) return null; const v=k*i.kg; return r?Math.ceil(v/r)*r:Math.round(v*100)/100; }
function drawPC(){
  const m=pcMatch();
  $("#pcList").innerHTML=m.length?`<p class="hint">${m.length} รายการ${m.some(i=>!(+i.kg>0))?" · รายการที่ยังไม่มีน้ำหนัก/หน่วยจะข้าม":""}</p>`+m.map(i=>{ const n=pcNew(i);
    return `<div class="mrow"><span>${esc(i.name)} <small>${esc(i.spec)}</small><br><small>${+i.kg?fmt(i.kg)+" กก./"+esc(i.unit):"ไม่มีน้ำหนัก"}${+i.cost&&!isStaff()&&n?` · ทุน ${money(i.cost)}`:""}</small></span>
      <span class="num">${money(i.price)}${n!==null?` → <b class="${n<(+i.cost||0)?"neg":"pos"}">${money(n)}</b>`:""}</span></div>`; }).join(""):`<p class="hint">ไม่พบสินค้าในหมวดนี้</p>`;
  $("#pcHist").innerHTML=priceHist.slice(0,10).map(h=>`<div class="mrow"><span>${esc(h.cat)}${h.filter?` "${esc(h.filter)}"`:""} → ${money(h.perKg)}/กก.<br><small>${dstr(h.at)} · ${h.items.length} รายการ${byTxt(h)}</small></span><span></span></div>`).join("")||`<p class="hint">ยังไม่เคยปรับ</p>`;
}
$("#priceCatBtn").addEventListener("click",()=>{ $("#pcCat").innerHTML=["ทุกหมวด",...CATS].map(c=>`<option>${c}</option>`).join(""); $("#pcFilter").value=""; $("#pcKg").value=""; drawPC(); $("#dPriceCat").showModal(); });
$("#dPriceCat").addEventListener("input",drawPC); $("#dPriceCat").addEventListener("change",drawPC);
$("#pcClose").addEventListener("click",()=>$("#dPriceCat").close());
$("#pcApply").addEventListener("click",async()=>{
  const ch=pcMatch().map(i=>({i,n:pcNew(i)})).filter(x=>x.n!==null&&x.n!==+x.i.price);
  if(!ch.length) return toast("ไม่มีราคาที่เปลี่ยน (ใส่ราคาต่อ กก. และตรวจว่าสินค้ามีน้ำหนัก)");
  if(!confirm(`อัปเดตราคาขาย ${ch.length} รายการ?`)) return;
  const at=Date.now();
  try{ for(const {i,n} of ch) await store.update("items/"+i.id,{price:n,updatedAt:at});
    await store.set("prices/"+uid(),{at,cat:$("#pcCat").value,filter:$("#pcFilter").value.trim(),perKg:+$("#pcKg").value,items:ch.map(({i,n})=>({id:i.id,name:i.name,old:+i.price||0,new:n}))});
    toast(`อัปเดตราคา ${ch.length} รายการแล้ว`); setTimeout(drawPC,300);
  }catch(e){ toast("บันทึกไม่สำเร็จ"); }
});

/* ===================== stock count ===================== */
let ctCat="ทั้งหมด", ctVals={};
function drawCT(){
  $("#ctChips").innerHTML=["ทั้งหมด",...CATS].map(c=>`<button aria-pressed="${c===ctCat}" data-ct="${c}">${c}</button>`).join("");
  const list=items.filter(i=>ctCat==="ทั้งหมด"||i.cat===ctCat).sort((a,b)=>String(a.name).localeCompare(String(b.name),"th"));
  $("#ctList").innerHTML=list.map(i=>{ const v=ctVals[i.id]; const d=v===undefined||v===""?null:(+v)-(+i.qty||0);
    return `<div class="ctr"><span>${esc(i.name)}<small>${esc(i.spec)} · ในระบบ ${fmt(i.qty)} ${esc(i.unit)}${remOf(i).length?` + เหลือ ${remOf(i).length}`:""}${i.loc?" · "+esc(i.loc):""}</small></span>
      <input data-ct="${esc(i.id)}" type="number" inputmode="decimal" step="any" value="${v??""}" placeholder="นับได้"><span class="d ${d?d<0?"neg":"pos":""}">${d===null?"":(d>0?"+":"")+fmt(d)}</span></div>`; }).join("");
  ctSum();
}
function ctSum(){ let n=0,val=0; Object.entries(ctVals).forEach(([id,v])=>{ if(v==="") return; const i=items.find(x=>x.id===id); if(!i) return; const d=(+v)-(+i.qty||0); if(d){ n++; val+=d*(+i.cost||+i.price||0); } });
  $("#ctDiff").textContent=`${n} รายการ${isStaff()?"":" · "+(val>0?"+":"")+money(val)+" บาท"}`; }
$("#countBtn").addEventListener("click",()=>{ ctVals={}; drawCT(); $("#dCount").showModal(); });
$("#ctChips").addEventListener("click",e=>{ const b=e.target.closest("[data-ct]"); if(b){ ctCat=b.dataset.ct; drawCT(); } });
$("#ctList").addEventListener("input",e=>{ const id=e.target.dataset.ct; if(!id) return; ctVals[id]=e.target.value;
  const i=items.find(x=>x.id===id), d=e.target.value===""?null:(+e.target.value)-(+i.qty||0), s=e.target.nextElementSibling;
  s.textContent=d===null?"":(d>0?"+":"")+fmt(d); s.className="d "+(d?d<0?"neg":"pos":""); ctSum(); });
$("#ctClose").addEventListener("click",()=>$("#dCount").close());
$("#ctSave").addEventListener("click",async()=>{
  const ch=Object.entries(ctVals).filter(([,v])=>v!=="").map(([id,v])=>({i:items.find(x=>x.id===id),n:+v})).filter(x=>x.i&&x.n!==(+x.i.qty||0));
  if(!ch.length) return toast("ไม่มีส่วนต่าง — สต็อกตรงกับที่นับ");
  const reason=$("#ctReason").value, at=Date.now();
  try{ for(const {i,n} of ch){ const d=n-(+i.qty||0); await store.update("items/"+i.id,{qty:n,updatedAt:at});
      await addMove({type:d>0?"in":"out",qty:Math.abs(d),label:`ปรับ ${d>0?"+":"−"}${fmt(Math.abs(d))}`,itemId:i.id,name:i.name,note:`ตรวจนับ: ${reason} (ระบบ ${fmt(i.qty)} → นับได้ ${fmt(n)})`,at}); }
    await store.set("counts/"+uid(),{at,reason,items:ch.map(({i,n})=>({id:i.id,name:i.name,before:+i.qty||0,after:n,cost:+i.cost||0}))});
    $("#dCount").close(); toast(`ปรับสต็อก ${ch.length} รายการแล้ว`);
  }catch(e){ toast("บันทึกไม่สำเร็จ"); }
});

/* ===================== import from Excel ===================== */
const IM_MAP={name:["ชื่อสินค้า","ชื่อ","สินค้า","name"],cat:["หมวด","category"],cond:["สภาพ"],spec:["ขนาด/สเปก","ขนาด","สเปก","spec"],qty:["คงเหลือ","จำนวน","qty"],
  unit:["หน่วย","unit"],kg:["กก./หน่วย","น้ำหนัก/หน่วย","น้ำหนักต่อหน่วย"],cost:["ต้นทุน/หน่วย","ทุน","ต้นทุน"],price:["ราคาขาย/หน่วย","ราคาขาย","ราคา","price"],
  min:["แจ้งเตือนต่ำกว่า","จุดสั่งซื้อ"],loc:["ตำแหน่ง","ที่วาง"],lenFull:["ความยาวเต็ม (ม.)","ความยาวเต็ม"]};
let imRows=[];
$("#importBtn").addEventListener("click",()=>{ imRows=[]; $("#imPrev").innerHTML=""; $("#imApply").hidden=true; $("#dImport").showModal(); });
$("#imClose").addEventListener("click",()=>$("#dImport").close());
$("#imPick").addEventListener("click",()=>$("#imFile").click());
$("#imTpl").addEventListener("click",async()=>{
  if(!downloads||!window.XLSX) return toast("ดาวน์โหลดไฟล์ไม่ได้ในหน้านี้");
  const rows=[{"ชื่อสินค้า":"ท่อเหลี่ยม ZM 32x32 หนา 1.1","หมวด":"เหล็กกล่อง","สภาพ":"ใหม่","ขนาด/สเปก":"32x32 มม. หนา 1.1 ยาว 6 ม.","คงเหลือ":100,"หน่วย":"เส้น","กก./หน่วย":6.2,"ต้นทุน/หน่วย":150,"ราคาขาย/หน่วย":174,"แจ้งเตือนต่ำกว่า":20,"ตำแหน่ง":"ชั้น A","ความยาวเต็ม (ม.)":6},
    {"ชื่อสินค้า":"เหล็กฉาก 2x2 หนา 5","หมวด":"เหล็กฉาก","สภาพ":"มือสอง","ขนาด/สเปก":"50x50 มม. หนา 5","คงเหลือ":12,"หน่วย":"เส้น","กก./หน่วย":22.4,"ต้นทุน/หน่วย":380,"ราคาขาย/หน่วย":470,"แจ้งเตือนต่ำกว่า":3,"ตำแหน่ง":"ลานหลัง","ความยาวเต็ม (ม.)":6}];
  const wb=XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(rows),"สต็อก");
  try{ await downloads.save({filename:"ตัวอย่างนำเข้าสต็อก.xlsx",data:XLSX.write(wb,{bookType:"xlsx",type:"array"})}); }catch(e){ if(e&&e.code!=="declined") toast("ดาวน์โหลดไม่สำเร็จ"); }
});
$("#imFile").addEventListener("change",async e=>{
  const file=e.target.files[0]; e.target.value=""; if(!file) return;
  if(!window.XLSX) return toast("อ่านไฟล์ไม่ได้ (โหลดตัวอ่าน Excel ไม่สำเร็จ)");
  try{
    const wb=XLSX.read(await file.arrayBuffer(),{type:"array"}), ws=wb.Sheets[wb.SheetNames.includes("สต็อก")?"สต็อก":wb.SheetNames[0]];
    const raw=XLSX.utils.sheet_to_json(ws,{defval:""}); const hdr=Object.keys(raw[0]||{});
    const col={}; Object.entries(IM_MAP).forEach(([k,al])=>{ const h=hdr.find(x=>al.some(a=>String(x).trim().toLowerCase()===a.toLowerCase())); if(h) col[k]=h; });
    if(!col.name) return $("#imPrev").innerHTML=`<div class="msg bad">ไม่พบคอลัมน์ "ชื่อสินค้า" — ใช้หัวตารางตามไฟล์ตัวอย่าง</div>`;
    imRows=raw.map(r=>{ const g=k=>col[k]?r[col[k]]:""; const n=k=>+String(g(k)).replace(/,/g,"")||0;
      const v={name:String(g("name")).trim(),cat:CATS.includes(String(g("cat")).trim())?String(g("cat")).trim():"อื่นๆ",cond:/มือสอง|used/i.test(g("cond"))?"used":"new",spec:String(g("spec")).trim(),
        qty:n("qty"),unit:String(g("unit")).trim()||"เส้น",kg:n("kg"),cost:n("cost"),price:n("price"),min:n("min"),loc:String(g("loc")).trim(),lenFull:n("lenFull")};
      const ex=items.find(i=>i.name.trim()===v.name&&String(i.spec||"").trim()===v.spec); return {v,ex}; }).filter(x=>x.v.name);
    const nNew=imRows.filter(x=>!x.ex).length;
    $("#imPrev").innerHTML=`<div class="msg ok">อ่านได้ ${imRows.length} รายการ · เพิ่มใหม่ ${nNew} · อัปเดตของเดิม ${imRows.length-nNew}<br><small>คอลัมน์ที่พบ: ${Object.values(col).map(esc).join(", ")}</small></div>`+
      imRows.slice(0,12).map(({v,ex})=>`<div class="mrow"><span>${esc(v.name)} <small>${esc(v.spec)}</small><br><small>${ex?"อัปเดต":"ใหม่"} · ${esc(v.cat)} · ${fmt(v.qty)} ${esc(v.unit)}</small></span><span class="num">${money(v.price)}</span></div>`).join("")+(imRows.length>12?`<p class="hint">…อีก ${imRows.length-12} รายการ</p>`:"");
    $("#imApply").hidden=!imRows.length;
  }catch(err){ $("#imPrev").innerHTML=`<div class="msg bad">อ่านไฟล์ไม่ได้ ลองบันทึกเป็น .xlsx หรือ .csv ใหม่</div>`; }
});
$("#imApply").addEventListener("click",async()=>{
  const b=$("#imApply"); b.disabled=true; const at=Date.now(); let n=0;
  try{ for(const {v,ex} of imRows){ if(ex) await store.update("items/"+ex.id,{...v,updatedAt:at}); else await store.set("items/"+uid(),{...v,rem:[],updatedAt:at}); n++; }
    $("#dImport").close(); toast(`นำเข้า ${n} รายการแล้ว`);
  }catch(e){ toast(`นำเข้าได้ ${n} รายการ แล้วเกิดข้อผิดพลาด`); } finally{ b.disabled=false; }
});

/* ===================== scrap grades & selling scrap ===================== */
let grades=[];
const DEFAULT_GRADES=[{id:"g1",name:"เหล็กหนา",buy:0,sell:0},{id:"g2",name:"เหล็กบาง",buy:0,sell:0},{id:"g3",name:"เศษเหล็กรวม",buy:0,sell:0}];
let sgEdit=[];
function drawSG(){ $("#sgList").innerHTML=sgEdit.map((g,i)=>`<div class="sgr" data-i="${i}"><label>เกรด<input data-k="name" value="${esc(g.name)}"></label>
  <label>รับซื้อ/กก.<input data-k="buy" type="number" inputmode="decimal" step="any" value="${g.buy||""}"></label><label>ขายต่อ/กก.<input data-k="sell" type="number" inputmode="decimal" step="any" value="${g.sell||""}"></label>
  <button class="rm" data-rm="${i}" aria-label="ลบ">×</button></div>`).join(""); }
$("#scrapBtn").addEventListener("click",()=>{ sgEdit=(grades.length?grades:DEFAULT_GRADES).map(g=>({...g})); drawSG(); $("#dScrap").showModal(); });
$("#sgList").addEventListener("input",e=>{ const r=e.target.closest("[data-i]"); if(!r) return; const k=e.target.dataset.k; sgEdit[+r.dataset.i][k]=k==="name"?e.target.value:(+e.target.value||0); });
$("#sgList").addEventListener("click",e=>{ const r=e.target.dataset.rm; if(r!==undefined){ sgEdit.splice(+r,1); drawSG(); } });
$("#sgAdd").addEventListener("click",()=>{ sgEdit.push({id:uid(),name:"",buy:0,sell:0}); drawSG(); });
$("#sgClose").addEventListener("click",()=>$("#dScrap").close());
$("#sgSave").addEventListener("click",async()=>{ const g=sgEdit.filter(x=>x.name.trim()).map(x=>({...x,name:x.name.trim()}));
  try{ await store.set("settings/scrap",{grades:g}); grades=g; $("#dScrap").close(); toast("บันทึกราคารับซื้อแล้ว"); }catch(e){ toast("บันทึกไม่สำเร็จ"); } });
function drawGradeChips(){
  const used=fb.querySelector("[name=kind]:checked").value==="used", gs=grades.length?grades:[];
  $("#bGrades").innerHTML = used&&!editingBuy ? (gs.length? `<span class="hint" style="width:100%;margin:0">แตะเกรดเพื่อเพิ่มรายการ แล้วใส่น้ำหนักจากตาชั่ง</span>`+gs.map(g=>`<span class="chip" style="padding-right:10px;cursor:pointer" data-grade="${esc(g.id)}">${esc(g.name)} · ${money(g.buy)}/กก.</span>`).join("")
    : `<span class="hint" style="margin:0">ตั้งราคารับซื้อตามเกรดได้ที่ เพิ่มเติม › ราคารับซื้อเหล็กเก่า</span>`) : "";
}
fb.addEventListener("change",e=>{ if(e.target.name==="kind") drawGradeChips(); });
$("#bGrades").addEventListener("click",e=>{ const c=e.target.closest("[data-grade]"); if(!c) return; const g=grades.find(x=>x.id===c.dataset.grade);
  if(blines.length===1&&!blines[0].desc&&!blines[0].qty) blines=[];
  blines.push({desc:g.name,qty:0,unit:"กก.",price:+g.buy||0,itemId:"",addQty:0,grade:g.name}); drawBLines();
  setTimeout(()=>{ const q=$("#bLines").querySelectorAll("[data-k=qty]"); q[q.length-1]?.focus(); },50); });
async function gradeItem(name){ // find/create stock item for scrap grade
  const nm="เหล็กเก่า "+name; let it=items.find(i=>i.name===nm&&i.unit==="กก.");
  if(it) return it; const g=grades.find(x=>x.name===name)||{}; const id=uid();
  it={id,name:nm,cat:"อื่นๆ",cond:"used",spec:"",qty:0,unit:"กก.",kg:1,min:0,cost:0,price:+g.sell||0,loc:"",rem:[],updatedAt:Date.now()};
  const {id:_x,...doc}=it; await store.set("items/"+id,doc); items=[...items,it]; return it;
}
const fss=$("#fSS");
$("#scrapSellBtn").addEventListener("click",()=>{
  const opts=items.filter(i=>i.unit==="กก."&&(+i.qty||0)>0);
  fss.item.innerHTML=opts.length?opts.map(i=>`<option value="${esc(i.id)}">${esc(i.name)} (มี ${fmt(i.qty)} กก.)</option>`).join(""):`<option value="">ไม่มีเหล็กเก่า/เศษ (หน่วย กก.) ในสต็อก</option>`;
  fss.kg.value=""; fss.buyer.value=""; ssFill(); $("#dScrapSell").showModal(); });
function ssFill(){ const i=items.find(x=>x.id===fss.item.value); const g=i&&grades.find(x=>"เหล็กเก่า "+x.name===i.name); if(i) fss.ppk.value=(g&&g.sell)||+i.price||""; ssCalc(); }
function ssCalc(){ const i=items.find(x=>x.id===fss.item.value), kg=+fss.kg.value||0, p=+fss.ppk.value||0; $("#ssTotal").textContent=money(kg*p); $("#ssProfit").textContent=money(kg*(p-(+i?.cost||0))); }
fss.item.addEventListener("change",ssFill); fss.addEventListener("input",ssCalc);
fss.addEventListener("submit",async e=>{
  if(e.submitter&&e.submitter.value!=="save") return; e.preventDefault(); if(!fss.reportValidity()) return;
  const i=items.find(x=>x.id===fss.item.value); if(!i) return toast("เลือกสินค้า"); const kg=+fss.kg.value, p=+fss.ppk.value;
  if(kg>(+i.qty||0)) return toast(`มีแค่ ${fmt(i.qty)} กก.`);
  const at=Date.now();
  try{ await store.update("items/"+i.id,{qty:Math.round(((+i.qty||0)-kg)*100)/100,updatedAt:at});
    await addMove({type:"out",qty:kg,itemId:i.id,name:i.name,note:`ขายเศษให้ ${fss.buyer.value.trim()||"โรงหลอม"} ${money(p)}/กก.`,at});
    await store.set("cash/"+uid(),{type:"in",cat:"ขายเศษเหล็ก",amount:Math.round(kg*p*100)/100,date:at,via:fss.via.value,note:`${i.name} ${fmt(kg)} กก. → ${fss.buyer.value.trim()||"โรงหลอม"}`,itemId:i.id,kg,cogs:Math.round(kg*(+i.cost||0)*100)/100,at});
    $("#dScrapSell").close(); toast(`ขายเศษ ${fmt(kg)} กก. ได้ ${money(kg*p)} บาท`);
  }catch(err){ toast("บันทึกไม่สำเร็จ"); }
});

/* ===================== customer orders (to source) ===================== */
const OSTAT={new:"รอจัดหา",ordered:"สั่งของแล้ว",ready:"ของพร้อม",quoted:"ออกใบเสนอราคาแล้ว",done:"ส่ง/ขายแล้ว",cancel:"ยกเลิก"};
let orders=[], editingOrd=null, olines=[], orderLink=null;
const fo=$("#fOrd");
const openOrders=()=>orders.filter(o=>!["done","cancel"].includes(o.status));
function ordTag(o){
  const od=o.need&&!["done","cancel"].includes(o.status)?daysFrom(o.need):null;
  const cls={new:"due",ordered:"part",ready:"ok",quoted:"part",done:"ok",cancel:""}[o.status]||"";
  return `<span class="tag ${od>0?"due":cls}">${OSTAT[o.status]||""}${od>0?` · เลยวันนัด ${od} วัน`:""}</span>`;
}
function renderOrders(list){
  return list.length ? list.map(o=>{ const n=(o.lines||[]).length, d=(o.lines||[]).filter(l=>l.done).length, tot=(o.lines||[]).reduce((s,l)=>s+(+l.qty||0)*(+l.price||0),0);
    return `<button class="rec order" data-ord="${esc(o.id)}"><h3>${esc(o.customer)} ${ordTag(o)}</h3>
      <div class="amt"><span class="num">${tot?fmt(tot):"—"}</span><small>${tot?"บาท (ประมาณ)":""}</small></div>
      <div class="meta">${esc((o.lines||[]).slice(0,3).map(l=>`${l.desc} ${fmt(l.qty)}${l.unit}`).join(", "))}${n>3?" …":""}<br>
      ${o.ful==="ship"?"🚚 ส่ง"+(o.shipAddr?": "+esc(o.shipAddr.slice(0,40)):"")+" · ":"🏪 มารับเอง · "}${esc(o.channel||"")}${o.phone?" · "+esc(o.phone):""}${o.need?` · ต้องการ ${dlong(o.need)}`:""} · จัดหาแล้ว ${d}/${n}${o.docNo?` · ${esc(o.docNo)}`:""}${byTxt(o)}</div></button>`; }).join("")
    : `<div class="empty">ไม่มีออเดอร์ค้าง<br>ลูกค้าโทรมาสั่ง กด "📝 จดออเดอร์" ได้เลย</div>`;
}
const blankOLine=()=>({desc:"",qty:1,unit:"เส้น",price:0,cost:0,itemId:"",done:false});
function openOrder(o){
  editingOrd=o||null; $("#oTitle").textContent=o?"ออเดอร์ "+o.customer:"จดออเดอร์ลูกค้า";
  for(const k of ["customer","phone","note"]) fo[k].value=o?.[k]||""; fo.channel.value=o?.channel||"โทรศัพท์"; fo.status.value=o?.status||"new";
  fo.need.value=o?.need?isoDate(o.need):""; fo.deposit.value=o?.deposit||0; $("#oQuick").value=""; shipLoad(fo,o);
  olines=o?(o.lines||[]).map(l=>({...l})):[]; $("#oActs").hidden=!o; $("#oDel").hidden=!o;
  $("#oLink").textContent=o?.docNo?`เชื่อมกับเอกสาร ${o.docNo}`:""; drawOLines(); $("#dOrder").showModal();
}
function itemOptsO(sel){ return `<option value="">— ยังไม่ผูกสต็อก —</option>`+[...items].sort((a,c)=>String(a.name).localeCompare(String(c.name),"th")).map(i=>`<option value="${esc(i.id)}" ${i.id===sel?"selected":""}>${esc(i.name)} ${esc(i.spec)} (มี ${fmt(i.qty)})</option>`).join(""); }
function drawOLines(){
  $("#oLines").innerHTML=olines.map((l,i)=>`<div class="oline ${l.done?"done":""}" data-i="${i}">
    <input class="nm" data-k="desc" value="${esc(l.desc)}" placeholder="รายการ / ขนาด / ความหนา" maxlength="120">
    <input data-k="qty" type="number" inputmode="decimal" step="any" value="${l.qty||""}" placeholder="จำนวน">
    <input data-k="unit" list="unitList" value="${esc(l.unit)}" placeholder="หน่วย">
    <input data-k="price" type="number" inputmode="decimal" step="any" value="${l.price||""}" placeholder="ราคาขาย/หน่วย">
    <button type="button" class="rm" data-rm="${i}" aria-label="ลบ">×</button>
    <div class="st"><select data-k="itemId">${itemOptsO(l.itemId)}</select>
      <input class="owner" data-k="cost" type="number" inputmode="decimal" step="any" value="${l.cost||""}" placeholder="ทุน/หน่วย">
      <label class="ck"><input type="checkbox" data-k="done" ${l.done?"checked":""}> หาได้แล้ว</label></div>
    <div class="sub">รวม ${money((+l.qty||0)*(+l.price||0))}</div></div>`).join("")+`<datalist id="unitList">${UNITS.map(u=>`<option>${u}</option>`).join("")}</datalist>`;
  oTotals();
}
function oTotals(){ const t=olines.reduce((s,l)=>s+(+l.qty||0)*(+l.price||0),0), wc=olines.filter(l=>+l.cost>0), miss=olines.length-wc.length;
  const pr=wc.reduce((s,l)=>s+(+l.qty||0)*((+l.price||0)-(+l.cost||0)),0);
  $("#oTotal").textContent=money(t); $("#oProfit").textContent=wc.length?money(pr)+(miss?` (ยังไม่มีทุน ${miss} รายการ)`:""):"ใส่ทุนเพื่อดูกำไร"; }
$("#oLines").addEventListener("input",e=>{ const r=e.target.closest("[data-i]"), k=e.target.dataset.k; if(!r||!k) return; const l=olines[+r.dataset.i];
  if(k==="done"){ l.done=e.target.checked; r.classList.toggle("done",l.done); return; }
  l[k]=["qty","price","cost"].includes(k)?(+e.target.value||0):e.target.value;
  if(k==="itemId"){ const it=items.find(x=>x.id===l.itemId); if(it){ if(!l.price) l.price=+it.price||0; if(!l.cost) l.cost=+it.cost||0; if(!l.unit) l.unit=it.unit; if(!l.desc) l.desc=[it.name,it.spec].filter(Boolean).join(" "); if((+it.qty||0)>=(+l.qty||0)) l.done=true; drawOLines(); return; } }
  r.querySelector(".sub").textContent="รวม "+money((+l.qty||0)*(+l.price||0)); oTotals(); });
$("#oLines").addEventListener("change",e=>{ if(e.target.dataset.k==="done") $("#oLines").dispatchEvent(new Event("input")); });
$("#oLines").addEventListener("click",e=>{ const r=e.target.dataset.rm; if(r!==undefined){ olines.splice(+r,1); drawOLines(); } });
$("#oAdd").addEventListener("click",()=>{ olines.push(blankOLine()); drawOLines(); });
$("#oParse").addEventListener("click",()=>{
  const U="เส้น|แผ่น|ท่อน|ชิ้น|ตัว|กก\\.?|กิโล|ตัน|ม\\.|เมตร|ม้วน|อัน|ชุด";
  const re=new RegExp(`^(.*?)[\\s,]+(\\d+(?:\\.\\d+)?)\\s*(${U})\\s*$`);
  const rows=$("#oQuick").value.split(/\n+/).map(s=>s.trim()).filter(Boolean); if(!rows.length) return toast("พิมพ์รายการก่อน");
  rows.forEach(s=>{ const m=s.match(re); let desc=s, qty=1, unit="เส้น";
    if(m){ desc=m[1].trim(); qty=+m[2]; unit=m[3].replace(/^กิโล$/,"กก.").replace(/^กก$/,"กก.").replace(/^เมตร$/,"ม."); }
    const q=desc.toLowerCase(), it=items.find(i=>[i.name,i.spec].join(" ").toLowerCase()===q)||items.find(i=>q.length>3&&String(i.name).toLowerCase().includes(q));
    olines.push({desc,qty,unit,price:it?+it.price||0:0,cost:it?+it.cost||0:0,itemId:it?it.id:"",done:!!(it&&(+it.qty||0)>=qty)}); });
  $("#oQuick").value=""; drawOLines();
});
fo.addEventListener("submit",async e=>{
  if(e.submitter&&e.submitter.value!=="save") return; e.preventDefault(); if(!fo.reportValidity()) return;
  if($("#oQuick").value.trim()) $("#oParse").click();
  const ls=olines.filter(l=>String(l.desc).trim()).map(l=>({desc:String(l.desc).trim(),qty:+l.qty||0,unit:l.unit||"",price:+l.price||0,cost:+l.cost||0,itemId:l.itemId||"",done:!!l.done}));
  if(!ls.length) return toast("ใส่รายการอย่างน้อย 1 รายการ");
  let status=fo.status.value; if(status==="new"&&ls.every(l=>l.done)) status="ready";
  const v={customer:fo.customer.value.trim(),phone:fo.phone.value.trim(),channel:fo.channel.value,need:fo.need.value?new Date(fo.need.value+"T12:00").getTime():0,
    status,deposit:+fo.deposit.value||0,note:fo.note.value.trim(),lines:ls,at:editingOrd?.at||Date.now(),updatedAt:Date.now(),...shipData(fo)};
  try{ const cid=await upsertCust({name:v.customer,phone:v.phone,ship:v.shipAddr}); if(cid) v.custId=cid; }catch(e){}
  try{ if(editingOrd){ const {id:_x,...prev}=editingOrd; await store.set("orders/"+editingOrd.id,{...prev,...v}); } else await store.set("orders/"+uid(),v);
    $("#dOrder").close(); toast(editingOrd?"บันทึกออเดอร์แล้ว":`จดออเดอร์ ${v.customer} แล้ว`); }catch(err){ toast("บันทึกไม่สำเร็จ"); }
});
$("#oDel").addEventListener("click",async()=>{ if(!editingOrd||!confirm("ลบออเดอร์นี้?")) return; try{ await store.del("orders/"+editingOrd.id); $("#dOrder").close(); toast("ลบออเดอร์แล้ว"); }catch(e){ toast("ลบไม่สำเร็จ"); } });
function orderToDoc(type){
  const o=editingOrd; if(!o) return;
  const ls=olines.filter(l=>String(l.desc).trim()&&(+l.qty||0)>0);
  $("#dOrder").close(); openDoc(); fd.type.value=type; typeChanged();
  fd.customer.value=o.customer; fd.phone.value=o.phone||"";
  fd.note.value=[o.deposit?`รับมัดจำแล้ว ${money(o.deposit)} บาท`:"",o.note||""].filter(Boolean).join(" · ");
  lines=ls.map(l=>({itemId:l.itemId||"",name:l.desc,qty:+l.qty,unit:l.unit,price:+l.price||0,cost:+l.cost||0}));
  shipLoad(fd,o); orderLink={id:o.id,type}; drawLines();
}
$("#oQuote").addEventListener("click",()=>orderToDoc("quote"));
$("#oSell").addEventListener("click",()=>orderToDoc("receipt"));
$("#ordNew").addEventListener("click",()=>openOrder(null));

/* ===================== customers DB + delivery ===================== */
let customers=[], custSearch="";
const ckey=s=>String(s||"").trim().toLowerCase();
const GENERIC=/^(ลูกค้าหน้าร้าน|ลูกค้าทั่วไป|ทั่วไป|เงินสด)?$/;
const findCust=name=>{ const k=ckey(name); return k?customers.find(c=>ckey(c.name)===k):null; };
function fillCustList(){ $("#custList").innerHTML=customers.map(c=>`<option value="${esc(c.name)}">${esc([c.phone,c.kind!=="ทั่วไป"?c.kind:""].filter(Boolean).join(" · "))}</option>`).join(""); }
async function upsertCust(o){
  const name=String(o.name||"").trim(); if(GENERIC.test(name)) return "";
  const c=findCust(name);
  if(c){ const up={}; ["phone","addr","ship"].forEach(k=>{ if(o[k]&&!c[k]) up[k]=o[k]; }); if(Object.keys(up).length) await store.update("customers/"+c.id,{...up,updatedAt:Date.now()}); return c.id; }
  const id=uid(); await store.set("customers/"+id,{name,phone:o.phone||"",addr:o.addr||"",ship:o.ship||"",kind:"ทั่วไป",creditDays:0,taxId:"",line:"",note:"",at:Date.now()}); return id;
}
function custAutofill(form){
  const c=findCust(form.customer.value); if(!c) return;
  if(form.phone&&!form.phone.value) form.phone.value=c.phone||"";
  if(form.addr&&!form.addr.value) form.addr.value=c.addr||"";
  if(form.shipAddr&&!form.shipAddr.value&&(c.ship||c.addr)) form.shipAddr.value=c.ship||c.addr;
  if(form.dueDays&&+c.creditDays>0) form.dueDays.value=c.creditDays;
  if(c.note) toast("ℹ "+c.note);
}
fd.customer.addEventListener("change",()=>custAutofill(fd)); fo.customer.addEventListener("change",()=>custAutofill(fo));
function shipToggle(form){
  const s=form.ful.value==="ship"; form.querySelectorAll(".shipF").forEach(x=>x.hidden=!s);
  if(s&&!form.shipAddr.value){ const c=findCust(form.customer.value); form.shipAddr.value=(c&&(c.ship||c.addr))||(form.addr?form.addr.value:"")||""; }
}
[fd,fo].forEach(F=>F.ful.addEventListener("change",()=>shipToggle(F)));
function shipData(form){ if(form.ful.value!=="ship") return {ful:"pickup"};
  return {ful:"ship",shipAddr:form.shipAddr.value.trim(),shipPhone:form.shipPhone.value.trim(),shipDate:form.shipDate.value?new Date(form.shipDate.value+"T12:00").getTime():0,shipFee:+form.shipFee.value||0}; }
function shipLoad(form,o){ form.ful.value=o?.ful||"pickup"; form.shipAddr.value=o?.shipAddr||""; form.shipPhone.value=o?.shipPhone||""; form.shipDate.value=o?.shipDate?isoDate(o.shipDate):""; form.shipFee.value=o?.shipFee||0; shipToggle(form); }
const fulTag=x=> x.ful==="ship" ? (x.deliveredAt?`<span class="tag ok">🚚 ส่งแล้ว</span>`:`<span class="tag part">🚚 รอส่ง${x.shipDate?" "+dlong(x.shipDate):""}</span>`) : (x.ful==="pickup"?`<span class="tag">🏪 มารับเอง</span>`:"");
const toShip=()=>docs.filter(d=>d.type!=="quote"&&d.ful==="ship"&&!d.deliveredAt).sort((a,b)=>(a.shipDate||a.date)-(b.shipDate||b.date));
$("#vShip").addEventListener("click",async()=>{ const d=viewing; if(!d) return;
  try{ await store.update("docs/"+d.id,{deliveredAt:Date.now()}); openView({...d,deliveredAt:Date.now()}); toast("บันทึกส่งของแล้ว"); }catch(e){ toast("บันทึกไม่สำเร็จ"); } });

/* directory */
function custRows(){
  const m={};
  customers.forEach(c=>{ m[ckey(c.name)]={rec:c,name:c.name,phone:c.phone,kind:c.kind,bought:0,owed:0,last:0,n:0}; });
  const touch=(name,phone)=>{ const k=ckey(name); if(!k||GENERIC.test(name.trim())) return null; return m[k]||(m[k]={rec:null,name:name.trim(),phone,kind:"",bought:0,owed:0,last:0,n:0}); };
  docs.forEach(d=>{ const r=touch(d.customer,d.phone); if(!r) return; if(d.type!=="quote"){ r.bought+=+d.total||0; r.n++; } r.owed+=docOwed(d); r.last=Math.max(r.last,d.date||d.at); if(!r.phone&&d.phone) r.phone=d.phone; });
  orders.forEach(o=>{ const r=touch(o.customer,o.phone); if(r) r.last=Math.max(r.last,o.at||0); });
  return Object.values(m);
}
function renderCustDir(){
  const rows=custRows(), missing=rows.filter(r=>!r.rec).length, q=ckey(custSearch);
  const list=rows.filter(r=>!q||ckey(r.name).includes(q)||String(r.phone||"").includes(q)).sort((a,b)=>b.last-a.last||String(a.name).localeCompare(String(b.name),"th"));
  $("#cuListBox").innerHTML=`<div class="row" style="margin:0 0 8px"><button class="btn primary grow" data-cunew>+ เพิ่มลูกค้า</button>${missing?`<button class="btn grow" data-cuimp>เก็บรายชื่อจากเอกสารเก่า (${missing})</button>`:""}</div>`+
    (list.length?list.map(r=>`<button class="rec" data-cust="${esc(r.name)}" style="border-left-color:${r.owed>0?"var(--low)":"var(--steel)"}">
      <h3>${esc(r.name)} ${r.kind&&r.kind!=="ทั่วไป"?`<span class="tag">${esc(r.kind)}</span>`:""}${!r.rec?` <span class="tag">ยังไม่บันทึก</span>`:""}</h3>
      <div class="amt"><span class="num ${r.owed>0?"neg":""}">${fmt(r.owed>0?r.owed:r.bought)}</span><small>${r.owed>0?"ค้างชำระ":"ซื้อรวม"}</small></div>
      <div class="meta">${r.phone?esc(r.phone)+" · ":""}${r.n} บิล${r.last?" · ล่าสุด "+dlong(r.last):""}</div></button>`).join(""):`<div class="empty">ยังไม่มีรายชื่อลูกค้า</div>`);
}
$("#docList").addEventListener("input",e=>{ if(e.target.id==="cuSearch"){ custSearch=e.target.value; renderCustDir(); } });
$("#docList").addEventListener("click",async e=>{
  if(e.target.closest("[data-cunew]")) return openCust("",true);
  if(e.target.closest("[data-cuimp]")){ const miss=custRows().filter(r=>!r.rec); if(!confirm(`บันทึก ${miss.length} รายชื่อจากเอกสาร/ออเดอร์เก่าเข้าฐานข้อมูลลูกค้า?`)) return;
    try{ for(const r of miss){ const d=docs.find(x=>ckey(x.customer)===ckey(r.name)&&(x.addr||x.shipAddr))||{}; await store.set("customers/"+uid(),{name:r.name,phone:r.phone||"",addr:d.addr||"",ship:d.shipAddr||"",kind:"ทั่วไป",creditDays:0,taxId:"",line:"",note:"",at:Date.now()}); }
      toast(`บันทึก ${miss.length} รายชื่อแล้ว`); }catch(err){ toast("บันทึกไม่สำเร็จ"); } }
});

/* ===================== customer search & fill ===================== */
function custLookup(q){
  q=ckey(q); const qd=q.replace(/\D/g,"");
  const rows=custRows().filter(r=>!q||ckey(r.name).includes(q)||(qd.length>=3&&String(r.phone||"").replace(/\D/g,"").includes(qd)));
  return rows.sort((a,b)=>{ const as=ckey(a.name).startsWith(q)?0:1, bs=ckey(b.name).startsWith(q)?0:1; return as-bs||b.last-a.last; }).slice(0,8).map(r=>{
    if(r.rec) return {...r.rec,owed:r.owed};
    const d=docs.filter(x=>ckey(x.customer)===ckey(r.name)).sort((a,b)=>(b.date||b.at)-(a.date||a.at));
    const o=orders.filter(x=>ckey(x.customer)===ckey(r.name));
    return {name:r.name,phone:r.phone||"",addr:(d.find(x=>x.addr)||{}).addr||"",ship:(d.find(x=>x.shipAddr)||o.find(x=>x.shipAddr)||{}).shipAddr||"",owed:r.owed};
  });
}
function applyCust(form,c){
  form.customer.value=c.name;
  if(form.phone) form.phone.value=c.phone||"";
  if(form.addr) form.addr.value=c.addr||"";
  if(form.shipAddr){ form.shipAddr.value=c.ship||c.addr||""; }
  if(form.shipPhone&&!form.shipPhone.value&&c.phone) form.shipPhone.value=c.phone;
  if(form.dueDays&&+c.creditDays>0) form.dueDays.value=c.creditDays;
  form.querySelectorAll(".warn").forEach(x=>{ if(["customer","phone","addr"].includes(x.name)) clearWarn(x); });
  toast(`ดึงข้อมูล ${c.name} แล้ว${c.owed>0?` · ค้างชำระ ${money(c.owed)} บาท`:""}${c.note?" · "+c.note:""}`);
}
function attachCustPicker(input,onPick){
  input.removeAttribute("list");
  const box=document.createElement("div"); box.className="sugg"; box.hidden=true; input.insertAdjacentElement("afterend",box);
  let rows=[], hi=-1;
  const show=()=>{ rows=custLookup(input.value); hi=-1;
    if(!rows.length||(rows.length===1&&ckey(rows[0].name)===ckey(input.value))){ box.hidden=true; return; }
    box.innerHTML=`<div class="sh">ลูกค้าเก่า — แตะเพื่อดึงข้อมูล</div>`+rows.map((c,i)=>`<button type="button" data-i="${i}"><b>${esc(c.name)}</b>${c.kind&&c.kind!=="ทั่วไป"?` <span class="tag">${esc(c.kind)}</span>`:""}${c.owed>0?` <span class="tag due">ค้าง ${fmt(c.owed)}</span>`:""}
      <small>${esc([c.phone,(c.ship||c.addr||"").slice(0,40)].filter(Boolean).join(" · ")||"ไม่มีข้อมูลเพิ่มเติม")}</small></button>`).join("");
    box.hidden=false; };
  input.addEventListener("input",show);
  input.addEventListener("focus",()=>{ if(customers.length||docs.length) show(); });
  input.addEventListener("blur",()=>setTimeout(()=>box.hidden=true,180));
  input.addEventListener("keydown",e=>{ if(box.hidden) return; const bs=box.querySelectorAll("button");
    if(e.key==="ArrowDown"||e.key==="ArrowUp"){ e.preventDefault(); hi=(hi+(e.key==="ArrowDown"?1:-1)+bs.length)%bs.length; bs.forEach((b,i)=>b.classList.toggle("on",i===hi)); }
    else if(e.key==="Enter"&&hi>=0){ e.preventDefault(); onPick(rows[hi]); box.hidden=true; }
    else if(e.key==="Escape") box.hidden=true; });
  box.addEventListener("pointerdown",e=>e.preventDefault());
  box.addEventListener("click",e=>{ const b=e.target.closest("[data-i]"); if(!b) return; onPick(rows[+b.dataset.i]); box.hidden=true; });
}
attachCustPicker(fd.customer,c=>applyCust(fd,c));
attachCustPicker(fo.customer,c=>applyCust(fo,c));
attachCustPicker($("#qsCust"),c=>{ $("#qsCust").value=c.name; toast(`ลูกค้า ${c.name}${c.owed>0?` · ค้างชำระ ${money(c.owed)} บาท`:""}`); });

/* ---------- shop info ---------- */
const fs=$("#fShop");
function openShopModal(){
  for(const k of ["name","phone","tax","addr","promptpay","bankName","bankAcc","bankHolder","billNote","ownerPin","staffPin"]){
    if(fs[k]) fs[k].value = shop[k] || "";
  }
  if(fs.ownerPin) fs.ownerPin.value = shop.ownerPin || "8888";
  if(fs.staffPin) fs.staffPin.value = shop.staffPin || "1111";
  $("#dShop").showModal();
}
$("#shopBtn").addEventListener("click",()=>{
  if(isStaff()){
    openPinPad({
      title: "ตั้งค่าร้านค้า (เจ้าของร้าน)",
      desc: "กรุณาใส่รหัส PIN เจ้าของร้านเพื่อเข้าถึงการตั้งค่า",
      avatar: "🔐",
      target: "owner",
      hint: `รหัสเริ่มต้นเจ้าของร้านคือ <b>${shop.ownerPin || "8888"}</b>`,
      onDone: () => {
        currentRole = "owner";
        localStorage.setItem("khlang_lek_role", "owner");
        applyRole();
        render();
        openShopModal();
      }
    });
    return;
  }
  openShopModal();
});
fs.addEventListener("submit",async e=>{
  if(e.submitter&&e.submitter.value!=="save") return;
  e.preventDefault(); if(!fs.reportValidity()) return;
  const v={};
  for(const k of ["name","phone","tax","addr","promptpay","bankName","bankAcc","bankHolder","billNote","ownerPin","staffPin"]){
    if(fs[k]) v[k]=fs[k].value.trim();
  }
  if(!v.ownerPin) v.ownerPin = "8888";
  if(!v.staffPin) v.staffPin = "1111";
  try{
    await store.set("settings/shop",v);
    shop={...DEFAULT_SHOP,...v};
    $("#shopTitle").textContent = shop.name ? `คลังเหล็ก · ${shop.name}` : "คลังเหล็ก";
    $("#dShop").close();
    toast("บันทึกข้อมูลร้านและการตั้งค่าแล้ว");
  }catch(err){ toast("บันทึกไม่สำเร็จ"); }
});

/* ---------- Excel export ---------- */
$("#xlsBtn").addEventListener("click",async()=>{
  if(!downloads||!window.XLSX) return toast("ส่งออก Excel ไม่ได้ในหน้านี้");
  const wb=XLSX.utils.book_new();
  const add=(name,rows)=>XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(rows.length?rows:[{"":"ไม่มีข้อมูล"}]),name);
  add("สต็อก",items.map(i=>({"ชื่อสินค้า":i.name,"หมวด":i.cat,"สภาพ":i.cond==="used"?"มือสอง":"ใหม่","ขนาด/สเปก":i.spec,"คงเหลือ":+i.qty||0,"หน่วย":i.unit,
    "กก./หน่วย":+i.kg||0,"น้ำหนักรวม (กก.)":Math.round(itemKg(i)*100)/100,"ต้นทุน/หน่วย":+i.cost||0,"ราคาขาย/หน่วย":+i.price||0,
    "กำไร/หน่วย":(+i.cost>0&&+i.price>0)?Math.round((i.price-i.cost)*100)/100:"","ท่อน/แผ่นเหลือ":remOf(i).map(r=>remLabel(i,r)).join(", "),"มูลค่าทุน":Math.round(itemVal(i,"cost")*100)/100,"มูลค่าขาย":Math.round(itemVal(i)*100)/100,
    "แจ้งเตือนต่ำกว่า":+i.min||0,"สถานะ":isLow(i)?"ใกล้หมด":"","ตำแหน่ง":i.loc})));
  add("ซื้อเข้า",buys.flatMap(b=>legacyLines(b).map(l=>({"วันที่":dlong(b.date||b.at),"ประเภท":(b.kind||"used")==="used"?"เหล็กเก่า":"เหล็กใหม่","ผู้ขาย":b.seller,"เบอร์":b.phone||"",
    "เลขที่บิล":b.billNo||"","รายการ":l.desc,"จำนวน":l.qty,"หน่วย":l.unit,"ราคา/หน่วย":l.price,"จำนวนเงิน":Math.round(l.qty*l.price*100)/100,
    "เข้าสต็อก":(items.find(i=>i.id===l.itemId)||{}).name||"","จำนวนเข้าสต็อก":l.addQty||"","ยอดบิล":b.total,"มีรูปใบเสร็จ":b.photoId?"มี":"","การจ่าย":b.payStatus==="credit"?"เครดิต":"จ่ายแล้ว","คงค้างจ่าย":buyOwed(b)||"","ครบกำหนด":b.due?dlong(b.due):"","หมายเหตุ":b.note||""}))));
  add("ขาย-เอกสาร",docs.flatMap(d=>d.lines.map(l=>({"เลขที่":d.no,"ประเภท":DOCNAME[d.type],"วันที่":dlong(d.date||d.at),"ลูกค้า":d.customer,
    "รายการ":l.name,"จำนวน":l.qty,"หน่วย":l.unit,"ราคา/หน่วย":l.price,"จำนวนเงิน":l.qty*l.price,"ทุน/หน่วย":+l.cost||0,"กำไรรายการ":Math.round(l.qty*(l.price-(+l.cost||0))*100)/100,
    "ส่วนลดทั้งใบ":d.discount,"ยอดสุทธิทั้งใบ":d.total,"รับเงินทาง":d.via||"","การรับสินค้า":d.ful==="ship"?"จัดส่ง":d.ful==="pickup"?"มารับเอง":"","ที่อยู่จัดส่ง":d.shipAddr||"","ส่งแล้ว":d.deliveredAt?dlong(d.deliveredAt):"","คงค้าง":docOwed(d)||"","ครบกำหนด":d.due?dlong(d.due):"","ตัดสต็อก":d.cut?"ใช่":""}))));
  add("รายรับ-รายจ่าย",ledger().map(x=>({"วันที่":dlong(x.t),"ประเภท":x.dir==="in"?"รายรับ":"รายจ่าย","หมวด":x.cat,"รายการ":x.title,"อ้างอิง":x.sub,
    "รายรับ":x.dir==="in"?x.amt:"","รายจ่าย":x.dir==="out"?x.amt:""})));
  add("ลูกค้า",customers.map(c=>({"ชื่อ":c.name,"เบอร์":c.phone||"","ประเภท":c.kind||"","ที่อยู่ออกบิล":c.addr||"","ที่อยู่จัดส่ง":c.ship||"","เครดิต (วัน)":c.creditDays||"","เลขผู้เสียภาษี":c.taxId||"","LINE":c.line||"","หมายเหตุ":c.note||""})));
  add("ออเดอร์ลูกค้า",orders.flatMap(o=>(o.lines||[]).map(l=>({"วันที่จด":dlong(o.at),"ลูกค้า":o.customer,"เบอร์":o.phone||"","สั่งทาง":o.channel||"","ต้องการวันที่":o.need?dlong(o.need):"",
    "สถานะ":OSTAT[o.status]||"","รายการ":l.desc,"จำนวน":l.qty,"หน่วย":l.unit,"ราคาขาย/หน่วย":l.price,"หาได้แล้ว":l.done?"ใช่":"","เอกสาร":o.docNo||"","มัดจำ":o.deposit||""}))));
  add("ประวัติเข้า-ออก",moves.map(m=>({"วันเวลา":dstr(m.at),"ประเภท":m.type==="in"?"รับเข้า":"จ่ายออก","สินค้า":m.name,"จำนวน":m.qty,"หมายเหตุ":m.note||""})));
  const buf=XLSX.write(wb,{bookType:"xlsx",type:"array"});
  try{ await downloads.save({filename:`คลังเหล็ก ${today()}.xlsx`,data:buf}); toast("ส่งออก Excel แล้ว"); }
  catch(e){ if(e&&e.code!=="declined") toast("ส่งออกไม่สำเร็จ"); }
});

let tt; function toast(m){ const t=$("#toast"); t.textContent=m; t.classList.add("show"); clearTimeout(tt); tt=setTimeout(()=>t.classList.remove("show"),2400); }

/* ---------- JSON backup & restore ---------- */
$("#backupJsonBtn")?.addEventListener("click", async () => {
  if (!downloads) return toast("ไม่สามารถดาวน์โหลดไฟล์ได้");
  const data = {
    version: 1,
    exportedAt: Date.now(),
    shop,
    grades,
    items,
    moves,
    buys,
    docs,
    cash,
    closes,
    prices: priceHist,
    orders,
    customers
  };
  const jsonStr = JSON.stringify(data, null, 2);
  try {
    await downloads.save({ filename: `คลังเหล็ก-สำรองข้อมูล-${today()}.json`, data: jsonStr });
    toast("สำรองข้อมูลเป็นไฟล์ JSON แล้ว");
  } catch (e) {
    if (e?.code !== "declined") toast("สำรองข้อมูลไม่สำเร็จ");
  }
});

$("#restoreJsonBtn")?.addEventListener("click", () => {
  $("#restoreFile")?.click();
});

$("#restoreFile")?.addEventListener("change", async (e) => {
  const file = e.target.files?.[0];
  e.target.value = "";
  if (!file) return;
  try {
    const text = await file.text();
    const data = JSON.parse(text);
    if (!confirm("ต้องการกู้คืนข้อมูลจากไฟล์นี้หรือไม่? (ข้อมูลที่มีอยู่จะถูกอัปเดต)")) return;
    const collections = ["items", "moves", "buys", "docs", "cash", "closes", "orders", "customers"];
    for (const col of collections) {
      if (Array.isArray(data[col])) {
        for (const item of data[col]) {
          if (item?.id) await store.set(`${col}/${item.id}`, item);
        }
      }
    }
    if (data.shop) await store.set("settings/shop", data.shop);
    if (data.grades) await store.set("settings/scrap", { grades: data.grades });
    toast("กู้คืนข้อมูลสำเร็จ");
  } catch (err) {
    toast("ไฟล์ไม่ถูกต้องหรือไม่สามารถกู้คืนได้");
  }
});

function checkLocalDemoMigration() {
  if (store?.mode === "db") {
    try {
      const raw = localStorage.getItem("steel-stock-demo");
      if (raw) {
        const d = JSON.parse(raw);
        const hasData = Object.values(d).some(v => v && (Array.isArray(v) ? v.length : Object.keys(v).length > 0));
        if (hasData && $("#migrateLocalBtn")) {
          $("#migrateLocalBtn").hidden = false;
        }
      }
    } catch (e) {}
  }
}

$("#migrateLocalBtn")?.addEventListener("click", async () => {
  const raw = localStorage.getItem("steel-stock-demo");
  if (!raw) return toast("ไม่มีข้อมูลในโหมดทดลอง");
  if (!confirm("นำเข้าข้อมูลจากโหมดทดลองในเครื่องนี้ขึ้นฐานข้อมูลออนไลน์?")) return;
  try {
    const d = JSON.parse(raw);
    for (const [col, val] of Object.entries(d)) {
      if (!val) continue;
      const list = Array.isArray(val) ? val : Object.entries(val).map(([id, v]) => ({ id, ...v }));
      for (const item of list) {
        if (item?.id) await store.set(`${col}/${item.id}`, item);
      }
    }
    localStorage.removeItem("steel-stock-demo");
    $("#migrateLocalBtn").hidden = true;
    toast("ย้ายข้อมูลขึ้นฐานข้อมูลออนไลน์เรียบร้อยแล้ว");
  } catch (e) {
    toast("เกิดข้อผิดพลาดในการย้ายข้อมูล");
  }
});

init();
