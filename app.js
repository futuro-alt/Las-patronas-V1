(()=>{
'use strict';

/* ================= Constantes ================= */
const PZ=30;
const TIPOS=['blanco','marron'];
const TIPO_NOMBRE={blanco:'Huevo blanco',marron:'Huevo marrón'};
const TIPO_CORTO={blanco:'blanco',marron:'marrón'};
const CLI_TIPOS={individual:'Individual',tienda:'Tienda',distribuidor:'Distribuidor'};
const METODOS={efectivo:'Efectivo',transferencia:'Transferencia',otro:'Otro'};
const MOTIVOS=['Rotos en traslado','Rotos en bodega','Estrellados o sucios','Devolución de cliente','Viejos, no vendibles','Otro'];
const COLS=['clientes','compras','ventas','mermas','abonos'];
const TABS=['inicio','ventas','bodega','clientes','reportes'];
const TAB_NOMBRE={inicio:'Inicio',ventas:'Ventas',bodega:'Bodega',clientes:'Clientes',reportes:'Reportes'};
const FAB_TXT={venta:'Venta',compra:'Compra',cliente:'Cliente'};
const VERSION_APP='1.0.0';
const IMG={gallinas:'img/gallinas.jpg',blanca:'img/gallina-blanca.jpg',roja:'img/gallina-roja.jpg',negra:'img/gallina-negra.jpg'};
const XLSX_URL='vendor/xlsx.mini.min.js';
const BD_NOMBRE='las-patronas',BD_VERSION=1,ALMACENES=['clientes','compras','ventas','mermas','abonos','config'];
const DIAS_RESPALDO=7;

const ICO={
  inicio:'<path d="M3.5 11 12 4l8.5 7"/><path d="M5.5 9.6V20h13V9.6"/><path d="M10 20v-5.5h4V20"/>',
  ventas:'<path d="M4.5 8.5h15l-1.3 10.7a2 2 0 0 1-2 1.8H7.8a2 2 0 0 1-2-1.8z"/><path d="M9 8.5V7a3 3 0 0 1 6 0v1.5"/>',
  bodega:'<path d="M3 13h18v4.5a2.5 2.5 0 0 1-2.5 2.5h-13A2.5 2.5 0 0 1 3 17.5z"/><path d="M6.6 13c-1.4 0-2.2-1-2.2-2.4 0-1.9 1-4.1 2.2-4.1s2.2 2.2 2.2 4.1c0 1.4-.8 2.4-2.2 2.4zM12 13c-1.4 0-2.2-1-2.2-2.4 0-1.9 1-4.1 2.2-4.1s2.2 2.2 2.2 4.1c0 1.4-.8 2.4-2.2 2.4zM17.4 13c-1.4 0-2.2-1-2.2-2.4 0-1.9 1-4.1 2.2-4.1s2.2 2.2 2.2 4.1c0 1.4-.8 2.4-2.2 2.4z"/>',
  clientes:'<circle cx="9" cy="8" r="3.3"/><path d="M3 20c0-3.4 2.7-6 6-6s6 2.6 6 6"/><path d="M15.5 4.8a3.2 3.2 0 0 1 0 6.3"/><path d="M17.5 14.4c2 .7 3.5 2.8 3.5 5.6"/>',
  reportes:'<path d="M3.5 20h17"/><path d="M6.5 20v-6M11 20V6M15.5 20v-9M20 20v-4"/>',
  ajustes:'<path d="M4 7h9M17 7h3M4 12h3M11 12h9M4 17h11M19 17h1"/><circle cx="15" cy="7" r="2"/><circle cx="9" cy="12" r="2"/><circle cx="17" cy="17" r="2"/>',
  mas:'<path d="M12 5v14M5 12h14"/>',
  chat:'<path d="M20 11.6a8 8 0 0 1-11.7 7.1L4 20l1.3-4.1A8 8 0 1 1 20 11.6z"/><path d="M9 10.5h6M9 13.5h4"/>',
  check:'<path d="M5 12.5l4.5 4.5L19 7.5"/>',
  reloj:'<circle cx="12" cy="12" r="8"/><path d="M12 7.8V12l3 2"/>',
  alerta:'<path d="M12 4.2 21 19.5H3z"/><path d="M12 10v4.2M12 16.9v.2"/>',
  calendario:'<rect x="4" y="5.5" width="16" height="14.5" rx="2"/><path d="M4 10h16M9 3.5v4M15 3.5v4"/>',
  copiar:'<rect x="9" y="9" width="11" height="11" rx="2"/><path d="M5.5 15V6.5a2 2 0 0 1 2-2H15"/>',
  lapiz:'<path d="M4 20h4L19.5 8.5l-4-4L4 16z"/><path d="M13.5 6.5l4 4"/>',
  cerrar:'<path d="M6 6l12 12M18 6 6 18"/>',
  basura:'<path d="M5 7h14M10 7V4.8h4V7M7 7l1 13h8l1-13"/>',
  camion:'<path d="M3 6.5h11v9.5H3z"/><path d="M14 10h4l3 3.2V16h-7"/><circle cx="7" cy="17.6" r="1.9"/><circle cx="17.5" cy="17.6" r="1.9"/>',
  roto:'<path d="M12 3.2c-3.5 0-6.3 5.3-6.3 9.8a6.3 6.3 0 0 0 12.6 0c0-4.5-2.8-9.8-6.3-9.8z"/><path d="M8.3 11.6l2.2 1.6 1.6-2.2 1.7 2.2 2-1.5"/>',
  moneda:'<circle cx="12" cy="12" r="8.2"/><path d="M14.6 9.4c-.5-.8-1.5-1.3-2.6-1.3-1.5 0-2.5.8-2.5 1.9 0 2.6 5.2 1.4 5.2 4 0 1.1-1.1 2-2.6 2-1.2 0-2.2-.5-2.7-1.3M12 6.6v1.5M12 15.9v1.5"/>',
  buscar:'<circle cx="11" cy="11" r="6.2"/><path d="M20 20l-4.6-4.6"/>',
  descargar:'<path d="M12 4v11M7.5 10.5 12 15l4.5-4.5M5 20h14"/>',
  huevo:'<path d="M12 3.2c-3.5 0-6.3 5.3-6.3 9.8a6.3 6.3 0 0 0 12.6 0c0-4.5-2.8-9.8-6.3-9.8z"/>',
  apagado:'<circle cx="12" cy="12" r="8"/><path d="M6.5 17.5l11-11"/>',
  subir:'<path d="M12 20V9M7.5 13.5 12 9l4.5 4.5M5 4.5h14"/>',
  compartir:'<path d="M12 3.5v11M8 7.5l4-4 4 4"/><path d="M8 10.5H6v9.5h12v-9.5h-2"/>',
  respaldo:'<rect x="3.5" y="4.5" width="17" height="4" rx="1"/><path d="M5 8.5v10.5h14V8.5M10 12.5h4"/>',
  celular:'<rect x="7" y="2.8" width="10" height="18.4" rx="2.2"/><path d="M11 18h2"/>'
};
const ico=n=>`<svg class="ico" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true" focusable="false">${ICO[n]||''}</svg>`;

/* ================= Utilidades ================= */
const $=(s,r=document)=>r.querySelector(s);
const $$=(s,r=document)=>[...r.querySelectorAll(s)];
const esc=s=>String(s==null?'':s).replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const n0=v=>{const x=Number(v);return Number.isFinite(x)?x:0;};
const r2=v=>Math.round(n0(v)*100)/100;
const sum=(a,f)=>a.reduce((s,x)=>s+(typeof f==='function'?f(x):n0(x[f])),0);
const normal=s=>String(s||'').normalize('NFD').replace(/[̀-ͯ]/g,'').toLowerCase().trim();
const inicial=s=>{const t=String(s||'').trim();return (t[0]||'?').toUpperCase();};
const ahora=()=>new Date().toISOString();
const uid=()=>'ej-'+Math.random().toString(36).slice(2,10);
const entero=v=>{const n=parseInt(String(v==null?'':v).replace(/,/g,'').trim(),10);return Number.isFinite(n)&&n>0?n:0;};
const decimal=v=>{const s=String(v==null?'':v).trim().replace(/[$\s,]/g,'');if(s==='')return null;const n=Number(s);return Number.isFinite(n)?n:NaN;};
const cap=s=>s?s.charAt(0).toUpperCase()+s.slice(1):s;
function debounce(fn,ms){let t;return(...a)=>{clearTimeout(t);t=setTimeout(()=>fn(...a),ms);};}
function leerLS(k){try{return localStorage.getItem(k);}catch(e){return null;}}
function guardarLS(k,v){try{localStorage.setItem(k,v);}catch(e){}}

/* Fechas: siempre 'AAAA-MM-DD' en hora local */
const DIAS=['domingo','lunes','martes','miércoles','jueves','viernes','sábado'];
const DIAS_C=['dom','lun','mar','mié','jue','vie','sáb'];
const MESES=['enero','febrero','marzo','abril','mayo','junio','julio','agosto','septiembre','octubre','noviembre','diciembre'];
const MESES_C=['ene','feb','mar','abr','may','jun','jul','ago','sep','oct','nov','dic'];
const pad=n=>String(n).padStart(2,'0');
const ymd=(y,m,d)=>y+'-'+pad(m)+'-'+pad(d);
function hoy(){const d=new Date();return ymd(d.getFullYear(),d.getMonth()+1,d.getDate());}
const esFecha=s=>typeof s==='string'&&/^\d{4}-\d{2}-\d{2}$/.test(s);
function dnum(s){const [y,m,d]=String(s).split('-').map(Number);return Math.round(Date.UTC(y,m-1,d)/864e5);}
function dstr(n){const d=new Date(n*864e5);return ymd(d.getUTCFullYear(),d.getUTCMonth()+1,d.getUTCDate());}
const sumarDias=(s,k)=>dstr(dnum(s)+k);
const diffDias=(a,b)=>dnum(a)-dnum(b);
const fdate=s=>new Date(dnum(s)*864e5);
const dow=s=>fdate(s).getUTCDay();
const inicioSemana=s=>sumarDias(s,-((dow(s)+6)%7));
function fFecha(s){const d=fdate(s);return DIAS_C[d.getUTCDay()]+' '+d.getUTCDate()+' '+MESES_C[d.getUTCMonth()];}
function fDiaMes(s){const d=fdate(s);return d.getUTCDate()+' '+MESES_C[d.getUTCMonth()];}
function fFechaLarga(s){const d=fdate(s);return DIAS[d.getUTCDay()]+' '+d.getUTCDate()+' de '+MESES[d.getUTCMonth()];}
function fRel(s){const k=diffDias(s,hoy());if(k===0)return 'hoy';if(k===-1)return 'ayer';if(k===1)return 'mañana';return k<0?'hace '+(-k)+' días':'en '+k+' días';}
function etiquetaDia(s){const k=diffDias(s,hoy());if(k===0)return 'Hoy · '+fFecha(s);if(k===-1)return 'Ayer · '+fFecha(s);return cap(fFechaLarga(s));}
function etiquetaCorta(s){const k=diffDias(s,hoy());if(k===0)return 'Hoy';if(k===-1)return 'Ayer';return cap(fFecha(s));}
function rangoTxt(a,b){if(a===b)return fDiaMes(a);const da=fdate(a),db=fdate(b);if(da.getUTCMonth()===db.getUTCMonth()&&da.getUTCFullYear()===db.getUTCFullYear())return da.getUTCDate()+'–'+db.getUTCDate()+' '+MESES_C[db.getUTCMonth()];return fDiaMes(a)+' – '+fDiaMes(b);}
function saludoTxt(){const h=new Date().getHours();return h<5?'Buenas noches':h<12?'Buenos días':h<19?'Buenas tardes':'Buenas noches';}

/* Números */
const NF0=new Intl.NumberFormat('es-MX',{maximumFractionDigits:0});
const NF1=new Intl.NumberFormat('es-MX',{maximumFractionDigits:1});
const NF2=new Intl.NumberFormat('es-MX',{minimumFractionDigits:2,maximumFractionDigits:2});
const num=v=>NF0.format(Math.round(n0(v)));
function dinero(v){const x=r2(v),a=Math.abs(x);const s=Math.abs(a-Math.round(a))<0.005?NF0.format(Math.round(a)):NF2.format(a);return (x<0?'−':'')+'$'+s;}
function dineroEje(v){const a=Math.abs(v);const s=a>=1000?NF1.format(a/1000)+' mil':NF0.format(a);return (v<0?'−':'')+'$'+s;}
const pct=v=>Number.isFinite(v)?NF0.format(Math.round(v*100))+'%':'—';
const dineroR=v=>dinero(Math.round(n0(v)));
const pzas=n=>num(n)+(Math.round(n)===1?' pza':' pzas');
const dCorto=pz=>{pz=Math.max(0,Math.round(n0(pz)));const ch=Math.floor(pz/PZ),s=pz%PZ;return ch&&s?`${ch} ch. y ${s} pzas`:ch?`${ch} ch.`:`${s} pzas`;};
const chTxt=n=>num(n)+(Math.round(n)===1?' charola':' charolas');
function desglose(pz){pz=Math.max(0,Math.round(n0(pz)));const ch=Math.floor(pz/PZ),s=pz%PZ;const a=[];if(ch)a.push(ch+(ch===1?' charola':' charolas'));if(s||!ch)a.push(s+(s===1?' pieza':' piezas'));return a.join(' y ');}
function telWA(t){const d=String(t||'').replace(/\D/g,'');if(d.length===10)return '52'+d;if(d.length>=11&&d.length<=15)return d;return null;}
function fTel(t){const d=String(t||'').replace(/\D/g,'');return d.length===10?d.slice(0,3)+' '+d.slice(3,6)+' '+d.slice(6):String(t||'');}

/* ================= Estado ================= */
const nuevaBase=()=>({clientes:new Map(),compras:new Map(),ventas:new Map(),mermas:new Map(),abonos:new Map(),cfgDoc:null});
const S={
  modo:'cargando',              // cargando | real | ejemplo | sinalmacen
  real:nuevaBase(),ejemplo:null,meta:{},almacenListo:false,
  escribiendo:0,ver:0,
  sw:null,swEspera:null,instalar:null,
  tab:'inicio',hoja:null,_graf:null,
  f:{vVista:'ventas',vPer:'7d',vBuscar:'',bVista:'compras',cBuscar:'',cTipo:'todos',cOrden:'proximo',rPer:'mes',rDesde:'',rHasta:'',rGraf:'charolas',rTabla:false}
};
const D=()=>S.modo==='ejemplo'&&S.ejemplo?S.ejemplo:S.real;
const puedeEditar=()=>S.modo==='ejemplo'||S.modo==='real';

const CFG0={
  negocio:'Huevos de rancho: Las Patronas',proveedor:'',recordarDias:7,
  plantilla:'¡Hola {nombre}! Le saluda Las Patronas, huevos de rancho. Ya llegó huevo fresco de granja, blanco y marrón. ¿Le aparto sus charolas de esta semana?',
  stockMin:10,frescuraDias:15,
  precios:{compra:{blanco:null,marron:null},individual:{blanco:null,marron:null},tienda:{blanco:null,marron:null},distribuidor:{blanco:null,marron:null}}
};
function cfg(){
  const d=D().cfgDoc||{};const p=d.precios||{};const precios={};
  for(const k of Object.keys(CFG0.precios))precios[k]={...CFG0.precios[k],...(p[k]||{})};
  const o={...CFG0,...d,precios};
  o.recordarDias=Math.max(1,Math.round(n0(o.recordarDias))||7);
  o.stockMin=Math.max(0,Math.round(n0(o.stockMin)));
  o.frescuraDias=Math.max(1,Math.round(n0(o.frescuraDias))||15);
  o.plantilla=String(o.plantilla||CFG0.plantilla);
  o.negocio=String(o.negocio||CFG0.negocio);
  return o;
}

/* ================= Cálculos ================= */
const itemsDe=o=>Array.isArray(o&&o.items)?o.items.filter(it=>it&&TIPOS.includes(it.tipo)):[];
function porFecha(a,b){if(a.fecha!==b.fecha)return a.fecha<b.fecha?-1:1;const ca=a.creado||'',cb=b.creado||'';return ca<cb?-1:ca>cb?1:0;}
let _dc=null,_dk='';
function derivar(){
  const H=hoy(),key=S.modo+'|'+S.ver+'|'+H;
  if(_dc&&_dk===key)return _dc;
  const src=D(),c=cfg();
  const clientesArr=[...src.clientes.values()].filter(x=>x&&x.nombre);
  const valido=x=>x&&esFecha(x.fecha);
  const compras=[...src.compras.values()].filter(valido).sort(porFecha);
  const ventas=[...src.ventas.values()].filter(valido).sort(porFecha);
  const mermas=[...src.mermas.values()].filter(valido).sort(porFecha);
  const abonos=[...src.abonos.values()].filter(valido).sort(porFecha);

  /* Inventario por lotes: primero en entrar, primero en salir */
  const lotes={blanco:[],marron:[]},deuda={blanco:0,marron:0},ultCpz={blanco:null,marron:null};
  const ev=[];
  for(const cp of compras)for(const it of itemsDe(cp))if(n0(it.charolas)>0)ev.push({k:0,f:cp.fecha,cr:cp.creado||'',t:it.tipo,o:cp,it});
  for(const v of ventas)for(const it of itemsDe(v))if(n0(it.charolas)>0)ev.push({k:1,f:v.fecha,cr:v.creado||'',t:it.tipo,o:v,it});
  for(const m of mermas)if(n0(m.piezas)>0&&TIPOS.includes(m.tipo))ev.push({k:2,f:m.fecha,cr:m.creado||'',t:m.tipo,o:m});
  ev.sort((a,b)=>a.f<b.f?-1:a.f>b.f?1:(a.k-b.k)||(a.cr<b.cr?-1:a.cr>b.cr?1:0));
  const costoVT=new Map(),costoM=new Map();
  const respaldo=t=>ultCpz[t]!=null?ultCpz[t]:n0(c.precios.compra[t])/PZ;
  const consumir=(t,pz)=>{let costo=0,r=pz;const q=lotes[t];while(r>0&&q.length){const L=q[0];const x=Math.min(L.restante,r);costo+=x*L.cpz;L.restante-=x;r-=x;if(L.restante<=0)q.shift();}if(r>0){deuda[t]+=r;costo+=r*respaldo(t);}return costo;};
  for(const e of ev){
    if(e.k===0){const cpz=n0(e.it.precio)/PZ;ultCpz[e.t]=cpz;const inicial=Math.max(0,n0(e.it.charolas)*PZ-n0(e.it.danadas));let pz=inicial;if(deuda[e.t]>0){const x=Math.min(deuda[e.t],pz);deuda[e.t]-=x;pz-=x;}if(pz>0)lotes[e.t].push({compraId:e.o.id,fecha:e.f,cpz,inicial,restante:pz});}
    else if(e.k===1){const k=e.o.id+'|'+e.t;costoVT.set(k,(costoVT.get(k)||0)+consumir(e.t,n0(e.it.charolas)*PZ));}
    else costoM.set(e.o.id,consumir(e.t,n0(e.o.piezas)));
  }
  const inv={};
  for(const t of TIPOS){
    const pz=sum(lotes[t],'restante')-deuda[t];const viejo=lotes[t][0];
    inv[t]={piezas:pz,charolas:Math.floor(Math.max(0,pz)/PZ),sueltas:Math.max(0,pz)%PZ,valor:r2(sum(lotes[t],L=>L.restante*L.cpz)),
      lotes:lotes[t].map(L=>({...L,edad:Math.max(0,diffDias(H,L.fecha))})),edad:viejo?Math.max(0,diffDias(H,viejo.fecha)):null,faltan:deuda[t]};
  }

  const ventasD=ventas.map(v=>{
    const its=itemsDe(v).filter(it=>n0(it.charolas)>0).map(it=>({tipo:it.tipo,charolas:n0(it.charolas),precio:n0(it.precio)}));
    const total=r2(sum(its,it=>it.charolas*it.precio));
    const costoT={blanco:r2(costoVT.get(v.id+'|blanco')||0),marron:r2(costoVT.get(v.id+'|marron')||0)};
    const costo=r2(costoT.blanco+costoT.marron);const ch={blanco:0,marron:0};its.forEach(it=>ch[it.tipo]+=it.charolas);
    const pagado=v.pagado==null?total:Math.min(total,Math.max(0,n0(v.pagado)));
    return {...v,items:its,total,costo,costoT,ganancia:r2(total-costo),ch,charolas:ch.blanco+ch.marron,pagadoMomento:r2(pagado)};
  });
  const comprasD=compras.map(cp=>{
    const its=itemsDe(cp).filter(it=>n0(it.charolas)>0).map(it=>({tipo:it.tipo,charolas:n0(it.charolas),precio:n0(it.precio),danadas:n0(it.danadas)}));
    const ch={blanco:0,marron:0},dan={blanco:0,marron:0};let costoDan=0;
    its.forEach(it=>{ch[it.tipo]+=it.charolas;dan[it.tipo]+=it.danadas;costoDan+=it.danadas*it.precio/PZ;});
    return {...cp,items:its,total:r2(sum(its,it=>it.charolas*it.precio)),ch,charolas:ch.blanco+ch.marron,danadas:dan,danadasPz:dan.blanco+dan.marron,costoDanadas:r2(costoDan)};
  });
  const mermasD=mermas.map(m=>({...m,piezas:n0(m.piezas),costo:r2(costoM.get(m.id)||0)}));
  const abonosD=abonos.map(a=>({...a,monto:r2(a.monto)}));

  /* Clientes, saldos y recordatorios */
  const porCli=new Map();const g=id=>{if(!porCli.has(id))porCli.set(id,{ventas:[],abonos:[]});return porCli.get(id);};
  ventasD.forEach(v=>{if(v.clienteId)g(v.clienteId).ventas.push(v);});
  abonosD.forEach(a=>{if(a.clienteId)g(a.clienteId).abonos.push(a);});
  const pendiente=new Map();
  porCli.forEach(pc=>{
    let disp=sum(pc.abonos,'monto');
    pc.ventas.forEach(v=>{let p=r2(v.total-v.pagadoMomento);if(p>0&&disp>0){const x=Math.min(p,disp);p=r2(p-x);disp=r2(disp-x);}pendiente.set(v.id,Math.max(0,p));});
    pc.saldo=r2(sum(pc.ventas,v=>v.total-v.pagadoMomento)-sum(pc.abonos,'monto'));
  });
  const clientes=clientesArr.map(cl=>{
    const pc=porCli.get(cl.id)||{ventas:[],abonos:[],saldo:0};const vs=pc.ventas;const ult=vs.length?vs[vs.length-1]:null;
    const fechas=[...new Set(vs.map(v=>v.fecha))];let frecuencia=null;
    if(fechas.length>=3){const gaps=[];for(let i=1;i<fechas.length;i++)gaps.push(diffDias(fechas[i],fechas[i-1]));gaps.sort((a,b)=>a-b);frecuencia=gaps[Math.floor((gaps.length-1)/2)]||null;}
    const rc=cl.recordarCada;const rec=(rc===0||rc==='0')?0:(n0(rc)>0?Math.round(n0(rc)):c.recordarDias);
    let proximo=null;
    if(rec>0){const base=[ult&&ult.fecha,esFecha(cl.ultimoContacto)?cl.ultimoContacto:null,(!ult&&!cl.ultimoContacto)?(esFecha(cl.alta)?cl.alta:H):null].filter(Boolean).sort().pop();proximo=sumarDias(base,rec);if(esFecha(cl.posponerHasta)&&cl.posponerHasta>proximo)proximo=cl.posponerHasta;}
    const dias=proximo?diffDias(proximo,H):null;
    const estado=proximo==null?'sin':dias<0?'atrasado':dias===0?'hoy':dias<=2?'pronto':'aldia';
    return {...cl,tipo:CLI_TIPOS[cl.tipo]?cl.tipo:'individual',rec,ventas:vs,abonos:pc.abonos,ult,frecuencia,proximo,dias,estado,saldo:pc.saldo||0,
      totalCh:sum(vs,'charolas'),totalGastado:r2(sum(vs,'total')),totalGanancia:r2(sum(vs,'ganancia'))};
  });
  const clientesPorId=new Map(clientes.map(x=>[x.id,x]));
  ventasD.forEach(v=>{const cl=v.clienteId?clientesPorId.get(v.clienteId):null;v.cli=cl||null;v.nombre=cl?cl.nombre:(v.clienteNombre||(v.clienteId?'Cliente eliminado':'Venta de mostrador'));v.pendiente=v.clienteId?(pendiente.get(v.id)||0):r2(v.total-v.pagadoMomento);});
  abonosD.forEach(a=>{const cl=clientesPorId.get(a.clienteId);a.nombre=cl?cl.nombre:(a.clienteNombre||'Cliente eliminado');});

  _dc={H,c,inv,ventas:ventasD,compras:comprasD,mermas:mermasD,abonos:abonosD,clientes,clientesPorId,
    porCobrar:r2(sum(clientes,x=>Math.max(0,x.saldo))),
    vacio:!clientesArr.length&&!compras.length&&!ventas.length&&!mermas.length&&!abonos.length};
  _dk=key;return _dc;
}
function resumen(d,desde,hasta){
  const en=f=>(!desde||f>=desde)&&(!hasta||f<=hasta);
  const vs=d.ventas.filter(v=>en(v.fecha)),cs=d.compras.filter(x=>en(x.fecha)),ms=d.mermas.filter(x=>en(x.fecha));
  const r={vs,cs,ms,n:vs.length,ingreso:0,costo:0,ganancia:0,ch:{blanco:0,marron:0},ingresoT:{blanco:0,marron:0},costoT:{blanco:0,marron:0},compraCh:{blanco:0,marron:0},compraMonto:0,compraPz:0,mermaPz:0,danadasPz:0,mermaCosto:0,mermaPzTotal:0};
  for(const v of vs){r.ingreso+=v.total;r.costo+=v.costo;r.ganancia+=v.ganancia;for(const it of v.items){r.ch[it.tipo]+=it.charolas;r.ingresoT[it.tipo]+=it.charolas*it.precio;}r.costoT.blanco+=v.costoT.blanco;r.costoT.marron+=v.costoT.marron;}
  for(const c of cs){r.compraMonto+=c.total;for(const it of c.items){r.compraCh[it.tipo]+=it.charolas;r.compraPz+=it.charolas*PZ;}r.danadasPz+=c.danadasPz;r.mermaCosto+=c.costoDanadas;}
  for(const m of ms){r.mermaPz+=m.piezas;r.mermaCosto+=m.costo;}
  r.mermaPzTotal=r.mermaPz+r.danadasPz;
  for(const k of ['ingreso','costo','ganancia','compraMonto','mermaCosto'])r[k]=r2(r[k]);
  return r;
}
function precioSugerido(d,tipo,cli){
  if(cli&&cli.precioEspecial&&n0(cli.precioEspecial[tipo])>0)return n0(cli.precioEspecial[tipo]);
  const ct=(cli&&CLI_TIPOS[cli.tipo])?cli.tipo:'individual';
  const lp=d.c.precios[ct]&&d.c.precios[ct][tipo];if(n0(lp)>0)return n0(lp);
  let respaldo=null;
  for(let i=d.ventas.length-1;i>=0;i--){const v=d.ventas[i];const it=v.items.find(x=>x.tipo===tipo);if(!it)continue;if(((v.cli&&v.cli.tipo)||'individual')===ct)return it.precio;if(respaldo==null)respaldo=it.precio;}
  return respaldo;
}
function precioCompraSugerido(d,tipo){
  for(let i=d.compras.length-1;i>=0;i--){const it=d.compras[i].items.find(x=>x.tipo===tipo);if(it)return it.precio;}
  const p=d.c.precios.compra[tipo];return n0(p)>0?n0(p):null;
}
function movimientos(d){
  const a=[];
  d.ventas.forEach(o=>a.push({k:'venta',f:o.fecha,cr:o.creado||'',o}));
  d.compras.forEach(o=>a.push({k:'compra',f:o.fecha,cr:o.creado||'',o}));
  d.mermas.forEach(o=>a.push({k:'merma',f:o.fecha,cr:o.creado||'',o}));
  d.abonos.forEach(o=>a.push({k:'abono',f:o.fecha,cr:o.creado||'',o}));
  return a.sort((x,y)=>x.f<y.f?1:x.f>y.f?-1:(x.cr<y.cr?1:x.cr>y.cr?-1:0));
}
function aplicarPlantilla(tpl,x){
  const pedido=x.ult?itemsLargo(x.ult):'sus charolas de siempre';
  const dias=x.ult?String(Math.max(0,diffDias(hoy(),x.ult.fecha))):'unos';
  return String(tpl||CFG0.plantilla).replace(/\{nombre\}/gi,String(x.saludo||x.nombre||'').trim()).replace(/\{pedido\}/gi,pedido).replace(/\{dias\}/gi,dias).replace(/\{negocio\}/gi,cfg().negocio);
}
const mensajePara=x=>aplicarPlantilla(cfg().plantilla,x);

/* ================= Piezas de interfaz ================= */
const chip=(cls,ic,txt,extra='')=>`<span class="chip ${cls} ${extra}">${ic?ico(ic):''}${esc(txt)}</span>`;
function estadoChip(x){
  switch(x.estado){
    case 'atrasado':return chip('rojo','alerta',`Atrasado ${-x.dias} ${x.dias===-1?'día':'días'}`);
    case 'hoy':return chip('ocre','reloj','Toca hoy');
    case 'pronto':return chip('turquesa','calendario',x.dias===1?'Toca mañana':`En ${x.dias} días`);
    case 'aldia':return chip('olivo','check','Próximo: '+fDiaMes(x.proximo));
    default:return chip('','apagado','Sin recordatorio');
  }
}
const itemsHTML=o=>o.items.map(it=>`<span class="hv"><span class="huevo ${it.tipo}" aria-hidden="true"></span>${num(it.charolas)} ch.<span class="sr"> de ${TIPO_CORTO[it.tipo]}</span></span>`).join(' ');
const itemsTxt=o=>o.items.map(it=>`${num(it.charolas)} ch. ${TIPO_CORTO[it.tipo]}`).join(' · ');
function itemsLargo(o){const p=o.items.map((it,i)=>`${num(it.charolas)} ${it.charolas===1?'charola':'charolas'} de${i===0?' huevo':''} ${TIPO_CORTO[it.tipo]}`);return p.length===2?p[0]+' y '+p[1]:(p[0]||'');}
function filaHTML(o){return `<button type="button" class="fila" data-act="${o.act}" data-id="${esc(o.id)}"><span class="fila-ico">${ico(o.ico)}</span><span class="fila-txt"><span class="fila-tit">${esc(o.tit)}</span><span class="fila-sub">${o.sub||''}</span></span><span class="fila-der">${o.monto!=null?`<span class="fila-monto ${o.montoCls||''}">${o.monto}</span>`:''}${o.extra||''}</span></button>`;}
const vacioMini=(t,x)=>`<div class="parche vacio-mini">${ico('huevo')}<div><b>${esc(t)}</b>${x?`<p>${esc(x)}</p>`:''}</div></div>`;
const segmento=(k,ops,act,cls='')=>`<div class="segmento ${cls}" role="group">${ops.map(([v,l])=>`<button type="button" data-act="filtro" data-k="${k}" data-v="${v}" aria-pressed="${act===v}">${esc(l)}</button>`).join('')}</div>`;
const encabezado=(t,ceja,extra='')=>`<header class="titulo-vista"><div class="tv-txt">${ceja?`<p class="ceja">${esc(ceja)}</p>`:''}<h1>${esc(t)}</h1></div>${extra}</header>`;
const stat=(etq,val,sub)=>`<div class="stat"><span class="stat-etq">${esc(etq)}</span><span class="stat-val">${val}</span><span class="stat-sub">${sub||''}</span></div>`;
function botonWA(x){
  const tel=telWA(x.telefono);
  if(!tel)return `<button type="button" class="btn chico" data-act="copiar-msg" data-id="${esc(x.id)}">${ico('copiar')}Copiar mensaje</button>`;
  if(S.modo==='ejemplo')return `<button type="button" class="btn chico wa" data-act="wa-ejemplo" data-id="${esc(x.id)}">${ico('chat')}WhatsApp</button>`;
  return `<a class="btn chico wa" href="https://wa.me/${tel}?text=${encodeURIComponent(mensajePara(x))}" target="_blank" rel="noopener" data-act="wa" data-id="${esc(x.id)}">${ico('chat')}WhatsApp</a>`;
}

/* ================= Vistas ================= */
function vCargando(){return `<div class="vista"><div class="parche cosido estado-vacio"><img class="gallina" src="${IMG.negra}" alt=""><h1>Abriendo la libreta…</h1><p>Aquí verás las charolas en bodega, las ventas del día y a qué clientes les toca volver a comprar.</p></div></div>`;}
function vSinAlmacen(){return `<div class="vista"><div class="parche cosido estado-vacio"><img class="gallina" src="${IMG.negra}" alt=""><h1>No se pudo abrir la libreta</h1><p>Este navegador no permite guardar datos. Si estás en una ventana privada, abre la app en una ventana normal o desde su ícono en la pantalla de inicio.</p><div class="acc-fila" style="justify-content:center"><button type="button" class="btn primario" data-act="reintentar">Reintentar</button><button type="button" class="btn acento" data-act="ejemplo">Explorar con datos de ejemplo</button></div></div></div>`;}
function vBienvenida(){
  const pe=puedeEditar();
  const paso=(n,t,p,tipo,btn)=>`<li><span class="paso-n">${n}</span><div><b>${t}</b><p>${p}</p>${pe?`<button type="button" class="btn chico" data-act="nueva" data-tipo="${tipo}">${btn}</button>`:''}</div></li>`;
  return `<div class="vista">${encabezado(saludoTxt(),fFechaLarga(hoy()))}
  <article class="parche cosido bienvenida">
    <img class="bv-img" src="${IMG.gallinas}" alt="Gallinas bordadas del banner de Las Patronas">
    <div class="bv-in">
      <h2>Tu libreta está lista</h2>
      <p>Anota lo que compras a la granja y lo que vendes. La app lleva el inventario de charolas, calcula tu ganancia y te recuerda a qué cliente escribirle para que vuelva a comprar. Todo se guarda en este celular y funciona sin señal.</p>
      <ol class="pasos">
        ${paso(1,'Registra tu primera compra','Charolas de huevo blanco o marrón, precio por charola y piezas dañadas al recibir.','compra','Registrar compra')}
        ${paso(2,'Agrega a tus clientes','Nombre, WhatsApp y cada cuántos días quieres escribirles.','cliente','Agregar cliente')}
        ${paso(3,'Anota cada venta','La ganancia, el saldo pendiente y el próximo recordatorio se calculan solos.','venta','Registrar venta')}
      </ol>
      <div class="bv-pie"><span>¿Quieres verla con movimiento?</span><button type="button" class="enlace" data-act="ejemplo">Explorar con datos de ejemplo</button></div>
    </div>
  </article></div>`;
}
function tarjetaInv(d,t,detalle){
  const iv=d.inv[t],c=d.c,chips=[];
  if(iv.piezas<=0)chips.push(chip('rojo','alerta','Sin existencias'));
  else if(iv.charolas<=c.stockMin)chips.push(chip('ocre','alerta',iv.charolas===0?'Menos de una charola':'Quedan pocas'));
  if(iv.edad!=null&&iv.piezas>0){const viejo=iv.edad>=c.frescuraDias,cerca=iv.edad>=Math.ceil(c.frescuraDias*.7);chips.push(chip(viejo?'rojo':cerca?'ocre':'olivo',viejo?'alerta':'reloj',`Lote más antiguo: ${iv.edad===0?'de hoy':iv.edad+(iv.edad===1?' día':' días')}`));}
  const sueltas=iv.sueltas?`+ ${iv.sueltas} ${iv.sueltas===1?'pieza suelta':'piezas sueltas'}`:'Sin piezas sueltas';
  return `<article class="parche cosido inv"><img class="gallina" src="${t==='blanco'?IMG.blanca:IMG.roja}" alt="" width="56" height="56"><p class="inv-tipo"><span class="huevo ${t}"></span>${TIPO_NOMBRE[t]}</p><p class="inv-num">${num(iv.charolas)}<span>${iv.charolas===1?'charola':'charolas'}</span></p><p class="inv-sub">${sueltas}${detalle?` · ${dinero(iv.valor)} a costo`:''}</p><div class="chips">${chips.join('')}</div></article>`;
}
function gridInventario(d,detalle){
  const falt=TIPOS.filter(t=>d.inv[t].faltan>0);
  return `<section class="bloque" aria-label="Inventario">${detalle?'':`<div class="bloque-cab"><h2>En bodega</h2><button type="button" class="enlace" data-act="tab" data-tab="bodega">Ver lotes</button></div>`}
    <div class="inv-grid">${TIPOS.map(t=>tarjetaInv(d,t,detalle)).join('')}</div>
    ${falt.length?`<p class="nota-alerta">${ico('alerta')}<span>Hay más salidas que compras registradas de ${falt.map(t=>`huevo ${TIPO_CORTO[t]} (${desglose(d.inv[t].faltan)})`).join(' y ')}. Revisa que estén todas tus compras.</span></p>`:''}
  </section>`;
}
function filaRecordatorio(x){
  const meta=x.ult?`Última compra ${fRel(x.ult.fecha)} · ${itemsTxt(x.ult)}`:'Todavía no compra';
  return `<div class="rec"><div class="rec-info"><div class="rec-top"><button type="button" class="rec-nombre" data-act="cliente" data-id="${esc(x.id)}">${esc(x.nombre)}</button>${estadoChip(x)}</div><p class="rec-meta">${esc(meta)}${x.saldo>0.009?' · debe '+esc(dinero(x.saldo)):''}</p></div>
  <div class="rec-acc">${botonWA(x)}${puedeEditar()?`<button type="button" class="btn chico" data-act="contactado" data-id="${esc(x.id)}">${ico('check')}Ya le escribí</button><button type="button" class="btn chico solo-icono" data-act="posponer" data-id="${esc(x.id)}" aria-label="Posponer recordatorio de ${esc(x.nombre)}" title="Posponer">${ico('reloj')}</button>`:''}</div></div>`;
}
function bloqueRecordatorios(pend,prox){
  const proxTxt=prox.length?`<div class="prox-wrap">Próximos: ${prox.slice(0,4).map(x=>`<button type="button" class="enlace-suave" data-act="cliente" data-id="${esc(x.id)}">${esc(x.nombre)}</button> (${esc(fRel(x.proximo))})`).join(', ')}</div>`:'';
  return `<section class="bloque"><div class="bloque-cab"><h2 class="h2-grupo">Por contactar ${pend.length?chip('rojo','',String(pend.length)):''}</h2></div>
  <div class="parche">${pend.length?`<div class="lista">${pend.map(filaRecordatorio).join('')}</div>`:`<div class="vacio-mini">${ico('check')}<div><b>Nadie pendiente por hoy.</b><p>Cuando a un cliente le toque volver a comprar, aparecerá aquí con su mensaje de WhatsApp listo.</p></div></div>`}${proxTxt}</div></section>`;
}
const accionesRapidas=()=>`<nav class="acciones" aria-label="Registrar"><button type="button" class="accion" data-act="nueva" data-tipo="venta">${ico('ventas')}Venta</button><button type="button" class="accion" data-act="nueva" data-tipo="compra">${ico('camion')}Compra</button><button type="button" class="accion" data-act="nueva" data-tipo="merma">${ico('roto')}Merma</button><button type="button" class="accion" data-act="nueva" data-tipo="abono">${ico('moneda')}Cobro</button></nav>`;
function bloqueSemana(rs,rp,lun,H){
  const delta=(a,b)=>{if(!(b>0))return '';const r=Math.round((a-b)/b*100);if(!r)return '<span class="delta">Igual que la semana pasada</span>';return `<span class="delta ${r>0?'sube':'baja'}">${r>0?'▲':'▼'} ${Math.abs(r)}% vs. semana pasada</span>`;};
  const chT=rs.ch.blanco+rs.ch.marron;
  return `<section class="bloque"><div class="bloque-cab"><h2>Esta semana</h2><span class="ayuda">${esc(rangoTxt(lun,H))}</span></div><div class="parche stats">
    ${stat('Ventas',dineroR(rs.ingreso),delta(rs.ingreso,rp.ingreso)||`${num(rs.n)} ${rs.n===1?'venta':'ventas'}`)}
    ${stat('Ganancia',dineroR(rs.ganancia),rs.ingreso>0?'Margen '+pct(rs.ganancia/rs.ingreso):'Sin ventas aún')}
    ${stat('Charolas vendidas',num(chT),`<span class="hv"><span class="huevo blanco"></span>${num(rs.ch.blanco)} blanco</span><span class="hv"><span class="huevo marron"></span>${num(rs.ch.marron)} marrón</span>`)}
    ${stat('Merma',`${num(rs.mermaPzTotal)}<small> ${rs.mermaPzTotal===1?'pza':'pzas'}</small>`,rs.mermaPzTotal?dinero(rs.mermaCosto)+' a costo':'Sin piezas dañadas')}
  </div></section>`;
}
function bloqueCobrar(d,deud){
  return `<section class="bloque"><div class="bloque-cab"><h2>Por cobrar</h2><button type="button" class="enlace" data-act="cobranza">Ver cobranza</button></div><div class="parche"><div class="cobrar-total"><span class="stat-etq">Total pendiente</span><b>${dinero(d.porCobrar)}</b><span class="ayuda">${deud.length} ${deud.length===1?'cliente':'clientes'}</span></div>
  <div class="lista">${deud.slice(0,3).map(x=>filaHTML({act:'cliente',id:x.id,ico:'moneda',tit:x.nombre,sub:esc(x.ult?'Última compra '+fRel(x.ult.fecha):'Sin compras'),monto:dinero(x.saldo)})).join('')}</div></div></section>`;
}
function filaMov(m){
  const o=m.o;
  switch(m.k){
    case 'venta':return filaHTML({act:'venta',id:o.id,ico:'ventas',tit:o.nombre,sub:`${esc(etiquetaCorta(o.fecha))} · ${itemsHTML(o)}`,monto:dinero(o.total),extra:o.pendiente>0.009?chip('rojo','','Debe '+dinero(o.pendiente),'mini'):''});
    case 'compra':return filaHTML({act:'compra',id:o.id,ico:'camion',tit:'Compra · '+(o.proveedor||'granja'),sub:`${esc(etiquetaCorta(o.fecha))} · ${itemsHTML(o)}`,monto:dinero(o.total)});
    case 'merma':return filaHTML({act:'merma',id:o.id,ico:'roto',tit:'Merma · '+(o.motivo||'piezas dañadas'),sub:`${esc(etiquetaCorta(o.fecha))} · <span class="hv"><span class="huevo ${o.tipo}"></span>${esc(desglose(o.piezas))}</span>`,monto:pzas(o.piezas)});
    default:return filaHTML({act:'abono',id:o.id,ico:'moneda',tit:'Pago · '+o.nombre,sub:esc(etiquetaCorta(o.fecha)+' · '+(METODOS[o.metodo]||'Efectivo')),monto:'+'+dinero(o.monto),montoCls:'pos'});
  }
}
function vInicio(d){
  if(d.vacio&&S.modo==='real')return vBienvenida();
  const H=d.H,lun=inicioSemana(H),trans=diffDias(H,lun);
  const rs=resumen(d,lun,H),rp=resumen(d,sumarDias(lun,-7),sumarDias(lun,-7+trans));
  const pend=d.clientes.filter(x=>x.estado==='atrasado'||x.estado==='hoy').sort((a,b)=>a.dias-b.dias||a.nombre.localeCompare(b.nombre,'es'));
  const prox=d.clientes.filter(x=>x.estado==='pronto').sort((a,b)=>a.dias-b.dias);
  const deud=d.clientes.filter(x=>x.saldo>0.009).sort((a,b)=>b.saldo-a.saldo);
  const movs=movimientos(d).slice(0,6);
  return `<div class="vista">${encabezado(saludoTxt(),fFechaLarga(H))}
    ${gridInventario(d,false)}
    ${bloqueRecordatorios(pend,prox)}
    ${puedeEditar()?accionesRapidas():''}
    ${bloqueSemana(rs,rp,lun,H)}
    ${deud.length?bloqueCobrar(d,deud):''}
    ${movs.length?`<section class="bloque"><div class="bloque-cab"><h2>Últimos movimientos</h2><button type="button" class="enlace" data-act="tab" data-tab="ventas">Ver ventas</button></div><div class="parche lista">${movs.map(filaMov).join('')}</div></section>`:''}
  </div>`;
}
function filaVenta(v){return filaHTML({act:'venta',id:v.id,ico:'ventas',tit:v.nombre,sub:`${itemsHTML(v)}${v.metodo?' · '+esc(METODOS[v.metodo]||''):''}`,monto:dinero(v.total),extra:v.pendiente>0.009?chip('rojo','','Debe '+dinero(v.pendiente),'mini'):''});}
function agrupar(a,k){const m=new Map();for(const x of a){const key=k(x);if(!m.has(key))m.set(key,[]);m.get(key).push(x);}return [...m.entries()];}
function vVentas(d){
  const f=S.f,H=d.H;
  let cuerpo;
  if(f.vVista==='cobrar')cuerpo=vCobranza(d);
  else{
    const desde=f.vPer==='hoy'?H:f.vPer==='7d'?sumarDias(H,-6):f.vPer==='mes'?H.slice(0,8)+'01':null;
    let vs=d.ventas.filter(v=>!desde||v.fecha>=desde);
    const q=normal(f.vBuscar);if(q)vs=vs.filter(v=>normal(v.nombre).includes(q));
    vs=vs.slice().reverse();
    const grupos=agrupar(vs,v=>v.fecha);
    cuerpo=`<div class="filtros">${segmento('vPer',[['hoy','Hoy'],['7d','7 días'],['mes','Este mes'],['todo','Todo']],f.vPer)}<label class="buscar"><span class="sr">Buscar por cliente</span>${ico('buscar')}<input id="q-ventas" type="search" placeholder="Buscar por cliente" value="${esc(f.vBuscar)}" autocomplete="off"></label></div>
    <p class="resumen-linea">${vs.length?`<b>${num(vs.length)}</b> ${vs.length===1?'venta':'ventas'} · <b>${num(sum(vs,'charolas'))}</b> charolas · <b>${dinero(sum(vs,'total'))}</b> · ganancia <b>${dinero(sum(vs,'ganancia'))}</b>`:''}</p>
    ${vs.length?grupos.map(([fecha,arr])=>`<section class="grupo"><h3 class="grupo-fecha"><span>${esc(etiquetaDia(fecha))}</span><span class="num">${dinero(sum(arr,'total'))}</span></h3><div class="parche lista">${arr.map(filaVenta).join('')}</div></section>`).join(''):(d.ventas.length?vacioMini('No hay ventas con estos filtros.','Prueba con otro periodo o búsqueda.'):vacioMini('Aún no hay ventas registradas.',puedeEditar()?'Toca «Venta» para anotar la primera.':''))}`;
  }
  return `<div class="vista">${encabezado('Ventas',null)}${segmento('vVista',[['ventas','Ventas'],['cobrar','Por cobrar']],f.vVista,'grande')}${cuerpo}</div>`;
}
function vCobranza(d){
  const pe=puedeEditar();
  const deud=d.clientes.filter(x=>x.saldo>0.009).sort((a,b)=>b.saldo-a.saldo);
  const favor=d.clientes.filter(x=>x.saldo<-0.009);
  const pagos=d.abonos.slice().reverse().slice(0,30);
  const filaDeudor=x=>{const n=x.ventas.filter(v=>v.pendiente>0.009).length;return `<div class="fila fila-div"><button type="button" class="fila-boton" data-act="cliente" data-id="${esc(x.id)}"><span class="fila-ico inicial-ico">${esc(inicial(x.nombre))}</span><span class="fila-txt"><span class="fila-tit">${esc(x.nombre)}</span><span class="fila-sub">${n} ${n===1?'venta pendiente':'ventas pendientes'}${x.ult?' · última compra '+esc(fRel(x.ult.fecha)):''}</span></span></button><span class="fila-der"><span class="fila-monto">${dinero(x.saldo)}</span>${pe?`<button type="button" class="btn chico" data-act="nueva" data-tipo="abono" data-cliente="${esc(x.id)}">Cobrar</button>`:''}</span></div>`;};
  return `<div class="parche cobrar-head"><div><span class="stat-etq">Total por cobrar</span><b class="cobrar-num">${dinero(d.porCobrar)}</b></div>${pe?`<button type="button" class="btn acento" data-act="nueva" data-tipo="abono">${ico('moneda')}Registrar pago</button>`:''}</div>
  <section class="bloque"><h2 class="h2">Clientes con saldo</h2>${deud.length?`<div class="parche lista">${deud.map(filaDeudor).join('')}</div>`:vacioMini('Nadie te debe.','Las ventas fiadas aparecerán aquí hasta que registres el pago.')}</section>
  ${favor.length?`<section class="bloque"><h2 class="h2">Saldo a favor del cliente</h2><div class="parche lista">${favor.map(x=>filaHTML({act:'cliente',id:x.id,ico:'moneda',tit:x.nombre,sub:'Pagó de más; se descuenta en su próxima compra fiada',monto:dinero(-x.saldo),montoCls:'pos'})).join('')}</div></section>`:''}
  <section class="bloque"><h2 class="h2">Pagos recibidos</h2>${pagos.length?`<div class="parche lista">${pagos.map(a=>filaHTML({act:'abono',id:a.id,ico:'moneda',tit:a.nombre,sub:esc(etiquetaCorta(a.fecha)+' · '+(METODOS[a.metodo]||'Efectivo')),monto:'+'+dinero(a.monto),montoCls:'pos'})).join('')}</div>`:vacioMini('Aún no hay pagos registrados.','')}</section>`;
}
function filaLote(L,d){
  const c=d.c,viejo=L.edad>=c.frescuraDias,cerca=L.edad>=Math.ceil(c.frescuraDias*.7);
  const porc=L.inicial?Math.max(2,Math.round(L.restante/L.inicial*100)):0;
  const quedan=L.restante>=L.inicial?`Completo · ${dCorto(L.restante)}`:`Quedan ${dCorto(L.restante)} de ${dCorto(L.inicial)}`;
  return `<button type="button" class="fila" data-act="compra" data-id="${esc(L.compraId)}"><span class="fila-ico"><span class="huevo ${L.tipo} grande"></span></span><span class="fila-txt"><span class="fila-tit">${esc(TIPO_NOMBRE[L.tipo])} · ${esc(fFecha(L.fecha))}</span><span class="fila-sub">${esc(quedan)}</span><span class="barrita"><span style="width:${porc}%"></span></span></span><span class="fila-der">${chip(viejo?'rojo':cerca?'ocre':'olivo',viejo?'alerta':'reloj',L.edad===0?'De hoy':L.edad+(L.edad===1?' día':' días'))}</span></button>`;
}
function vBodega(d){
  const f=S.f,lotes=[];
  TIPOS.forEach(t=>d.inv[t].lotes.forEach(L=>lotes.push({...L,tipo:t})));
  lotes.sort((a,b)=>a.fecha<b.fecha?-1:a.fecha>b.fecha?1:0);
  let lista;
  if(f.bVista==='mermas'){const a=d.mermas.slice().reverse();lista=a.length?`<div class="parche lista">${a.map(m=>filaHTML({act:'merma',id:m.id,ico:'roto',tit:m.motivo||'Merma',sub:`${esc(etiquetaCorta(m.fecha))} · <span class="hv"><span class="huevo ${m.tipo}"></span>${esc(desglose(m.piezas))}</span>`,monto:pzas(m.piezas)})).join('')}</div><p class="ayuda">Las piezas dañadas al recibir se anotan en cada compra.</p>`:vacioMini('Sin merma registrada.','Anota aquí los huevos rotos, estrellados o que ya no se pueden vender.');}
  else{const a=d.compras.slice().reverse();lista=a.length?`<div class="parche lista">${a.map(x=>filaHTML({act:'compra',id:x.id,ico:'camion',tit:x.proveedor||'Compra a granja',sub:`${esc(etiquetaCorta(x.fecha))} · ${itemsHTML(x)}${x.danadasPz?' · '+num(x.danadasPz)+(x.danadasPz===1?' dañada':' dañadas'):''}`,monto:dinero(x.total)})).join('')}</div>`:vacioMini('Aún no hay compras registradas.','Registra lo que compras a la granja para llevar el inventario y el costo real de cada charola.');}
  const acc=puedeEditar()?`<div class="tv-acc"><button type="button" class="btn primario" data-act="nueva" data-tipo="compra">${ico('camion')}Registrar compra</button><button type="button" class="btn" data-act="nueva" data-tipo="merma">${ico('roto')}Registrar merma</button></div>`:'';
  return `<div class="vista">${encabezado('Bodega',null,acc)}
    ${gridInventario(d,true)}
    <section class="bloque"><div class="bloque-cab"><h2>Lotes en bodega</h2><span class="ayuda">Salen primero los más antiguos</span></div>
      ${lotes.length?`<div class="parche lista">${lotes.map(L=>filaLote(L,d)).join('')}</div>`:vacioMini('No hay lotes con existencias.','Cada compra a la granja crea un lote; aquí verás cuánto queda de cada uno y cuántos días lleva en bodega.')}
    </section>
    <section class="bloque"><div class="bloque-cab">${segmento('bVista',[['compras','Compras a granja'],['mermas','Merma']],f.bVista)}</div>${lista}</section>
  </div>`;
}
function filaCliente(x){return `<button type="button" class="fila" data-act="cliente" data-id="${esc(x.id)}"><span class="fila-ico inicial-ico">${esc(inicial(x.nombre))}</span><span class="fila-txt"><span class="fila-tit">${esc(x.nombre)}</span><span class="fila-sub">${esc(CLI_TIPOS[x.tipo])}${x.ult?' · compró '+esc(fRel(x.ult.fecha)):' · sin compras'}${x.totalCh?' · '+num(x.totalCh)+' ch. en total':''}</span></span><span class="fila-der">${estadoChip(x)}${x.saldo>0.009?chip('rojo','','Debe '+dinero(x.saldo),'mini'):''}</span></button>`;}
function vClientes(d){
  const f=S.f,total=d.clientes.length;let a=d.clientes.slice();
  if(f.cTipo!=='todos')a=a.filter(x=>x.tipo===f.cTipo);
  const q=normal(f.cBuscar),qd=q.replace(/\D/g,'');
  if(q)a=a.filter(x=>normal(x.nombre).includes(q)||normal(x.saludo).includes(q)||normal(x.zona).includes(q)||(qd.length>=3&&String(x.telefono||'').replace(/\D/g,'').includes(qd)));
  const ord={proximo:(x,y)=>(x.dias==null?1e9:x.dias)-(y.dias==null?1e9:y.dias)||x.nombre.localeCompare(y.nombre,'es'),nombre:(x,y)=>x.nombre.localeCompare(y.nombre,'es'),charolas:(x,y)=>y.totalCh-x.totalCh,saldo:(x,y)=>y.saldo-x.saldo};
  a.sort(ord[f.cOrden]||ord.proximo);
  const cuenta=Object.keys(CLI_TIPOS).map(k=>[k,`${CLI_TIPOS[k]} · ${d.clientes.filter(x=>x.tipo===k).length}`]);
  const acc=puedeEditar()?`<div class="tv-acc"><button type="button" class="btn primario" data-act="nueva" data-tipo="cliente">${ico('mas')}Agregar cliente</button></div>`:'';
  return `<div class="vista">${encabezado('Clientes',total?`${num(total)} ${total===1?'cliente':'clientes'}`:null,acc)}
    ${total?`<div class="filtros"><label class="buscar"><span class="sr">Buscar cliente</span>${ico('buscar')}<input id="q-clientes" type="search" placeholder="Buscar por nombre, zona o teléfono" value="${esc(f.cBuscar)}" autocomplete="off"></label><div class="sel" style="flex:0 1 220px;min-width:170px"><label class="sr" for="c-orden">Ordenar por</label><select id="c-orden">${[['proximo','Próximo recordatorio'],['nombre','Nombre'],['charolas','Más charolas'],['saldo','Saldo pendiente']].map(([k,l])=>`<option value="${k}" ${f.cOrden===k?'selected':''}>${l}</option>`).join('')}</select></div></div>${segmento('cTipo',[['todos','Todos'],...cuenta],f.cTipo)}`:''}
    ${a.length?`<div class="parche lista">${a.map(filaCliente).join('')}</div>`:total?vacioMini('Ningún cliente coincide.','Prueba con otra búsqueda o tipo de cliente.'):vacioMini('Aún no hay clientes.','Agrégalos con su WhatsApp y cada cuántos días quieres escribirles para que vuelvan a comprar.')}
  </div>`;
}
function rangoReporte(d){
  const H=d.H;
  switch(S.f.rPer){
    case 'semana':return [inicioSemana(H),H];
    case 'mes':return [H.slice(0,8)+'01',H];
    case 'mespasado':{const [y,m]=H.split('-').map(Number);const py=m===1?y-1:y,pm=m===1?12:m-1;return [ymd(py,pm,1),sumarDias(ymd(y,m,1),-1)];}
    case '90':return [sumarDias(H,-89),H];
    case 'anio':return [H.slice(0,4)+'-01-01',H];
    case 'fechas':{let a=esFecha(S.f.rDesde)?S.f.rDesde:sumarDias(H,-29),b=esFecha(S.f.rHasta)?S.f.rHasta:H;if(a>b)[a,b]=[b,a];return [a,b];}
    default:return [null,null];
  }
}
function cubetas(d,desde,hasta){
  const fin=hasta||d.H;let ini=desde||(d.ventas[0]&&d.ventas[0].fecha)||d.H;if(ini>fin)ini=fin;
  const porMes=diffDias(fin,ini)+1>126;const b=[];
  if(!porMes){let s=inicioSemana(ini);while(s<=fin){const e=sumarDias(s,6);b.push({ini:s,fin:e,corta:fDiaMes(s),larga:'Semana del '+rangoTxt(s,e),blanco:0,marron:0,ingreso:0,ganancia:0,ventas:0});s=sumarDias(s,7);}}
  else{let [y,m]=ini.split('-').map(Number);const [yf,mf]=fin.split('-').map(Number);while(y<yf||(y===yf&&m<=mf)){b.push({clave:y+'-'+pad(m),corta:MESES_C[m-1],larga:cap(MESES[m-1])+' '+y,blanco:0,marron:0,ingreso:0,ganancia:0,ventas:0});m++;if(m>12){m=1;y++;}}}
  const idx=porMes?new Map(b.map((x,i)=>[x.clave,i])):null;
  for(const v of d.ventas){if((desde&&v.fecha<desde)||v.fecha>fin)continue;const i=porMes?idx.get(v.fecha.slice(0,7)):Math.floor(diffDias(inicioSemana(v.fecha),b[0].ini)/7);const x=b[i];if(!x)continue;x.blanco+=v.ch.blanco;x.marron+=v.ch.marron;x.ingreso+=v.total;x.ganancia+=v.ganancia;x.ventas++;}
  return {b:b.length>24?b.slice(-24):b,porMes};
}
function topClientes(vs){const m=new Map();for(const v of vs){const k=v.clienteId||'__m';if(!m.has(k))m.set(k,{id:v.clienteId,nombre:v.clienteId?v.nombre:'Ventas de mostrador',ch:0,ingreso:0,ganancia:0,n:0});const o=m.get(k);o.ch+=v.charolas;o.ingreso+=v.total;o.ganancia+=v.ganancia;o.n++;}return [...m.values()].sort((a,b)=>b.ch-a.ch||b.ingreso-a.ingreso);}
function mermaMotivos(rs){const m=new Map();const add=(k,pz,c)=>{if(!m.has(k))m.set(k,{motivo:k,pz:0,costo:0});const o=m.get(k);o.pz+=pz;o.costo+=c;};for(const x of rs.ms)add(x.motivo||'Otro',x.piezas,x.costo);let dp=0,dc=0;for(const c of rs.cs){dp+=c.danadasPz;dc+=c.costoDanadas;}if(dp)add('Dañadas al recibir',dp,dc);return [...m.values()].sort((a,b)=>b.pz-a.pz);}
function tablaCubetas(c){return `<div class="tabla-env"><table><thead><tr><th>${c.porMes?'Mes':'Semana'}</th><th class="n">Blanco</th><th class="n">Marrón</th><th class="n">Charolas</th><th class="n">Ventas</th><th class="n">Ganancia</th></tr></thead><tbody>${c.b.map(x=>`<tr><td>${esc(c.porMes?x.larga:rangoTxt(x.ini,x.fin))}</td><td class="n">${num(x.blanco)}</td><td class="n">${num(x.marron)}</td><td class="n">${num(x.blanco+x.marron)}</td><td class="n">${dinero(x.ingreso)}</td><td class="n">${dinero(x.ganancia)}</td></tr>`).join('')}</tbody></table></div>`;}
function vReportes(d){
  const f=S.f,[desde,hasta]=rangoReporte(d);
  const rs=resumen(d,desde,hasta),cub=cubetas(d,desde,hasta);S._graf=rs.n?cub:null;
  const chT=rs.ch.blanco+rs.ch.marron,nCli=new Set(rs.vs.filter(v=>v.clienteId).map(v=>v.clienteId)).size;
  const mermaPct=rs.compraPz?rs.mermaPzTotal/rs.compraPz:null;
  const top=topClientes(rs.vs).slice(0,8),mot=mermaMotivos(rs);
  const rango=desde?(desde===hasta?cap(fFecha(desde)):`${cap(fFecha(desde))} – ${fFecha(hasta)}`):'Todo el historial';
  const acc=`<div class="tv-acc"><button type="button" class="btn" data-act="exportar">${ico('descargar')}Exportar a Excel</button></div>`;
  const gan=f.rGraf==='ganancia';
  return `<div class="vista">${encabezado('Reportes',rango,acc)}
    <div class="filtros">${segmento('rPer',[['semana','Semana'],['mes','Mes'],['mespasado','Mes pasado'],['90','90 días'],['anio','Año'],['todo','Todo'],['fechas','Fechas']],f.rPer)}</div>
    ${f.rPer==='fechas'?`<div class="fila2" style="max-width:440px"><div class="campo"><label for="r-desde">Desde</label><input type="date" id="r-desde" value="${desde}"></div><div class="campo"><label for="r-hasta">Hasta</label><input type="date" id="r-hasta" value="${hasta}"></div></div>`:''}
    <div class="parche stats">
      ${stat('Ventas',dineroR(rs.ingreso),`${num(rs.n)} ${rs.n===1?'venta':'ventas'}`)}
      ${stat('Ganancia',dineroR(rs.ganancia),rs.ingreso>0?'Margen '+pct(rs.ganancia/rs.ingreso):'—')}
      ${stat('Charolas vendidas',num(chT),`<span class="hv"><span class="huevo blanco"></span>${num(rs.ch.blanco)} blanco</span><span class="hv"><span class="huevo marron"></span>${num(rs.ch.marron)} marrón</span>`)}
      ${stat('Merma',`${num(rs.mermaPzTotal)}<small> ${rs.mermaPzTotal===1?'pza':'pzas'}</small>`,(mermaPct!=null?pct(mermaPct)+' de lo comprado · ':'')+dinero(rs.mermaCosto))}
    </div>
    <div class="parche stats">
      ${stat('Compras a granja',dineroR(rs.compraMonto),`${num(rs.compraCh.blanco+rs.compraCh.marron)} charolas`)}
      ${stat('Venta promedio',dineroR(rs.n?rs.ingreso/rs.n:0),rs.n?`${NF1.format(chT/rs.n)} charolas por venta`:'—')}
      ${stat('Clientes que compraron',num(nCli),`de ${num(d.clientes.length)} registrados`)}
      ${stat('Por cobrar hoy',dinero(d.porCobrar),`${num(d.clientes.filter(x=>x.saldo>0.009).length)} con saldo`)}
    </div>
    <section class="parche graf-card">
      <div class="graf-cab"><h2 class="h2">${gan?'Ganancia':'Charolas vendidas'} por ${cub.porMes?'mes':'semana'}</h2>${segmento('rGraf',[['charolas','Charolas'],['ganancia','Ganancia']],f.rGraf,'chico')}</div>
      ${rs.n?`${gan?'':`<div class="leyenda"><span><i style="background:var(--c-blanco)"></i>Huevo blanco</span><span><i style="background:var(--c-marron)"></i>Huevo marrón</span></div>`}<div id="graf" class="graf"></div><button type="button" class="enlace" data-act="filtro" data-k="rTabla" data-v="${f.rTabla?'0':'1'}">${f.rTabla?'Ocultar tabla':'Ver como tabla'}</button>${f.rTabla?tablaCubetas(cub):''}`:'<p class="ayuda">No hay ventas en este periodo.</p>'}
    </section>
    <section class="bloque"><h2 class="h2">Rentabilidad por tipo de huevo</h2><div class="parche tabla-env"><table><thead><tr><th>Tipo</th><th class="n">Vendidas</th><th class="n">Precio prom.</th><th class="n">Costo prom.</th><th class="n">Ganancia por charola</th><th class="n">Ganancia</th></tr></thead><tbody>${TIPOS.map(t=>{const ch=rs.ch[t],ing=rs.ingresoT[t],co=rs.costoT[t];return `<tr><td><span class="hv"><span class="huevo ${t}"></span>${TIPO_NOMBRE[t]}</span></td><td class="n">${num(ch)}</td><td class="n">${ch?dinero(ing/ch):'—'}</td><td class="n">${ch?dinero(co/ch):'—'}</td><td class="n">${ch?dinero((ing-co)/ch):'—'}</td><td class="n">${dinero(ing-co)}</td></tr>`;}).join('')}</tbody></table></div></section>
    <section class="bloque"><h2 class="h2">Mejores clientes</h2>${top.length?`<div class="parche tabla-env"><table><thead><tr><th>Cliente</th><th class="n">Compras</th><th class="n">Charolas</th><th class="n">Ventas</th><th class="n">Ganancia</th></tr></thead><tbody>${top.map(o=>`<tr><td>${o.id?`<button type="button" class="enlace-suave" data-act="cliente" data-id="${esc(o.id)}">${esc(o.nombre)}</button>`:esc(o.nombre)}</td><td class="n">${num(o.n)}</td><td class="n">${num(o.ch)}</td><td class="n">${dinero(o.ingreso)}</td><td class="n">${dinero(o.ganancia)}</td></tr>`).join('')}</tbody></table></div>`:vacioMini('Sin ventas en este periodo.','')}</section>
    <section class="bloque"><h2 class="h2">Merma por motivo</h2>${mot.length?`<div class="parche tabla-env"><table><thead><tr><th>Motivo</th><th class="n">Piezas</th><th class="n">Equivale a</th><th class="n">Costo</th></tr></thead><tbody>${mot.map(o=>`<tr><td>${esc(o.motivo)}</td><td class="n">${num(o.pz)}</td><td class="n">${esc(desglose(o.pz))}</td><td class="n">${dinero(o.costo)}</td></tr>`).join('')}</tbody></table></div>`:vacioMini('Sin merma en este periodo.','No se registraron piezas dañadas.')}</section>
  </div>`;
}

/* ================= Gráfica ================= */
function nice(min,max,enteros){
  if(max===min)max=min+(enteros?4:100);
  let paso=(max-min)/4;const mag=Math.pow(10,Math.floor(Math.log10(paso)));const e=paso/mag;
  paso=(e>=7.5?10:e>=3.5?5:e>=1.5?2:1)*mag;if(enteros)paso=Math.max(1,Math.round(paso));
  const lo=Math.floor(min/paso)*paso,hi=Math.ceil(max/paso)*paso;const t=[];
  for(let v=lo;v<=hi+paso*1e-6;v+=paso)t.push(Math.round(v*1e6)/1e6);
  if(t.length<2)t.push(lo+paso);return t;
}
const f2=v=>(+v).toFixed(2);
function barraArriba(x,y,w,h){if(h<=0)return '';const r=Math.min(4,h,w/2);return `M${f2(x)},${f2(y+h)}V${f2(y+r)}Q${f2(x)},${f2(y)} ${f2(x+r)},${f2(y)}H${f2(x+w-r)}Q${f2(x+w)},${f2(y)} ${f2(x+w)},${f2(y+r)}V${f2(y+h)}Z`;}
function barraAbajo(x,y,w,h){if(h<=0)return '';const r=Math.min(4,h,w/2);return `M${f2(x)},${f2(y)}V${f2(y+h-r)}Q${f2(x)},${f2(y+h)} ${f2(x+r)},${f2(y+h)}H${f2(x+w-r)}Q${f2(x+w)},${f2(y+h)} ${f2(x+w)},${f2(y+h-r)}V${f2(y)}Z`;}
function dibujarGrafica(){
  const cont=document.getElementById('graf');const cub=S._graf;if(!cont||!cub||!cub.b.length)return;
  const datos=cub.b,modo=S.f.rGraf==='ganancia'?'ganancia':'charolas';
  const W=Math.max(260,Math.floor(cont.clientWidth)),H=236;
  const vals=datos.map(b=>modo==='charolas'?b.blanco+b.marron:r2(b.ganancia));
  const maxV=Math.max(0,...vals),minV=Math.min(0,...vals);
  const ticks=nice(minV,maxV,modo==='charolas'),lo=ticks[0],hi=ticks[ticks.length-1];
  const etqTick=t=>modo==='charolas'?num(t):dineroEje(t);
  const m={t:24,r:8,b:28,l:Math.max(30,Math.max(...ticks.map(t=>etqTick(t).length))*7+12)};
  const pw=W-m.l-m.r,ph=H-m.t-m.b;
  const y=v=>m.t+ph-(v-lo)/((hi-lo)||1)*ph;
  const n=datos.length,band=pw/n,bw=Math.max(5,Math.min(24,band*.62)),y0=y(0);
  let s='';
  for(const t of ticks){const yy=Math.round(y(t))+.5;s+=`<line x1="${m.l}" x2="${W-m.r}" y1="${yy}" y2="${yy}" class="g-grid${t===0?' g-cero':''}"/><text x="${m.l-8}" y="${yy+4}" class="g-eje" text-anchor="end">${esc(etqTick(t))}</text>`;}
  const cada=Math.max(1,Math.ceil(46/band));
  datos.forEach((b,i)=>{
    const cx=m.l+band*i+band/2,x=cx-bw/2;let seg='';
    if(modo==='charolas'){
      const yb=y(b.blanco),yt=y(b.blanco+b.marron);
      if(b.blanco>0)seg+=b.marron>0?`<rect x="${f2(x)}" y="${f2(yb)}" width="${f2(bw)}" height="${f2(y0-yb)}" class="g-blanco"/>`:`<path d="${barraArriba(x,yb,bw,y0-yb)}" class="g-blanco"/>`;
      if(b.marron>0){const base=b.blanco>0?yb-2:y0;const h=base-yt;if(h>.5)seg+=`<path d="${barraArriba(x,yt,bw,h)}" class="g-marron"/>`;}
    }else{const v=vals[i];if(v>0)seg=`<path d="${barraArriba(x,y(v),bw,y0-y(v))}" class="g-ganancia"/>`;else if(v<0)seg=`<path d="${barraAbajo(x,y0,bw,y(v)-y0)}" class="g-perdida"/>`;}
    s+=`<g class="g-barra" data-i="${i}">${seg}</g>`;
    if((n-1-i)%cada===0)s+=`<text x="${f2(cx)}" y="${H-m.b+18}" class="g-eje" text-anchor="middle">${esc(b.corta)}</text>`;
  });
  const iMax=vals.indexOf(Math.max(...vals));
  [...new Set([iMax,n-1])].forEach(i=>{const v=vals[i];if(v===0||(modo==='charolas'&&v<0))return;const cx=m.l+band*i+band/2;const yy=v>=0?y(v)-7:y(v)+15;s+=`<text x="${f2(cx)}" y="${f2(yy)}" class="g-val" text-anchor="middle">${esc(modo==='charolas'?num(v):dineroEje(v))}</text>`;});
  datos.forEach((b,i)=>{const et=modo==='charolas'?`${b.larga}: ${num(b.blanco)} charolas de blanco y ${num(b.marron)} de marrón`:`${b.larga}: ganancia de ${dinero(b.ganancia)}`;s+=`<rect class="g-hit" x="${f2(m.l+band*i)}" y="${m.t}" width="${f2(band)}" height="${ph}" data-i="${i}" tabindex="0" role="img" aria-label="${esc(et)}"></rect>`;});
  cont.innerHTML=`<svg width="${W}" height="${H}" viewBox="0 0 ${W} ${H}" role="group" aria-label="${modo==='charolas'?'Charolas vendidas':'Ganancia'} por ${cub.porMes?'mes':'semana'}">${s}</svg><div class="tip" hidden></div>`;
  const svg=cont.querySelector('svg'),tip=cont.querySelector('.tip');
  const fila=(color,valor,txt)=>{const r=document.createElement('div');r.className='tr';const k=document.createElement('i');k.style.background=color;const b=document.createElement('b');b.textContent=valor;const sp=document.createElement('span');sp.textContent=txt;r.append(k,b,sp);return r;};
  const mostrar=i=>{
    const b=datos[i];if(!b)return;tip.textContent='';
    const tt=document.createElement('div');tt.className='tt';tt.textContent=b.larga;tip.appendChild(tt);
    if(modo==='charolas'){tip.append(fila('var(--c-blanco)',num(b.blanco),'Huevo blanco'),fila('var(--c-marron)',num(b.marron),'Huevo marrón'),fila('transparent',dinero(b.ingreso),`Ventas · ${num(b.blanco+b.marron)} charolas`));}
    else tip.append(fila('var(--c-ganancia)',dinero(b.ganancia),'Ganancia'),fila('transparent',dinero(b.ingreso),'Ventas'));
    tip.hidden=false;
    const cx=m.l+band*i+band/2,tw=tip.offsetWidth,th=tip.offsetHeight;
    tip.style.left=Math.max(0,Math.min(W-tw,cx-tw/2))+'px';
    tip.style.top=Math.max(0,y(Math.max(vals[i],0))-th-10)+'px';
    cont.classList.add('activa');svg.querySelectorAll('.g-barra').forEach(g=>g.classList.toggle('on',g.dataset.i===String(i)));
  };
  const ocultar=()=>{tip.hidden=true;cont.classList.remove('activa');svg.querySelectorAll('.g-barra.on').forEach(g=>g.classList.remove('on'));};
  svg.addEventListener('pointermove',e=>{const h=e.target.closest('.g-hit');if(h)mostrar(+h.dataset.i);});
  svg.addEventListener('pointerleave',ocultar);
  svg.addEventListener('focusin',e=>{const h=e.target.closest('.g-hit');if(h)mostrar(+h.dataset.i);});
  svg.addEventListener('focusout',ocultar);
}

/* ================= Hojas (formularios y detalles) ================= */
function abrirHoja(o){
  const capa=$('#capa');const previo=S.hoja?S.hoja.foco:document.activeElement;
  capa.innerHTML=`<div class="hoja" role="dialog" aria-modal="true" aria-labelledby="hoja-tit" tabindex="-1"><div class="hoja-cab"><h2 id="hoja-tit">${esc(o.titulo)}</h2><button type="button" class="btn chico solo-icono" data-act="cerrar" aria-label="Cerrar">${ico('cerrar')}</button></div><div class="hoja-cuerpo">${o.cuerpo}</div>${o.pie?`<div class="hoja-pie">${o.pie}</div>`:''}</div>`;
  capa.hidden=false;capa.onclick=e=>{if(e.target===capa&&!o.formulario)cerrarHoja();};
  S.hoja={refrescar:o.refrescar||null,foco:previo};document.body.style.overflow='hidden';
  const h=capa.querySelector('.hoja');instalarControles(h);if(o.alMontar)o.alMontar(h);
  requestAnimationFrame(()=>{try{h.focus({preventScroll:true});}catch(e){}});
}
function cerrarHoja(){
  const capa=$('#capa');if(capa.hidden)return;
  capa.hidden=true;capa.innerHTML='';capa.onclick=null;document.body.style.overflow='';
  const f=S.hoja&&S.hoja.foco;S.hoja=null;if(f&&document.contains(f)){try{f.focus({preventScroll:true});}catch(e){}}
}
function instalarControles(h){
  h.addEventListener('click',e=>{
    const p=e.target.closest('[data-paso]');
    if(p){const inp=h.querySelector('#'+p.dataset.campo);if(inp){inp.value=String(Math.max(0,entero(inp.value)+Number(p.dataset.paso)));inp.dispatchEvent(new Event('input',{bubbles:true}));}return;}
    const b=e.target.closest('[data-grupo] > [data-valor]');
    if(b&&!b.disabled){const g=b.parentElement;g.dataset.valor=b.dataset.valor;for(const x of g.children)if(x.dataset&&('valor' in x.dataset))x.setAttribute('aria-pressed',String(x===b));g.dispatchEvent(new CustomEvent('grupo',{bubbles:true,detail:{grupo:g.dataset.grupo,valor:b.dataset.valor}}));}
  });
  h.addEventListener('focusin',e=>{const t=e.target;if(t.matches&&t.matches('.stepper input,.pref input'))setTimeout(()=>{try{t.select();}catch(x){}},0);});
  h.addEventListener('focusout',e=>{const t=e.target;if(t.matches&&t.matches('.stepper input')){const v=String(entero(t.value));if(t.value!==v){t.value=v;t.dispatchEvent(new Event('input',{bubbles:true}));}}});
}
function confirmarPie(h,texto,boton,fn){
  const pie=h.querySelector('.hoja-pie');if(!pie)return;const antes=pie.innerHTML;
  pie.innerHTML=`<p class="confirma-txt">${esc(texto)}</p><div class="confirma-btns"><button type="button" class="btn" data-no>Cancelar</button><button type="button" class="btn acento" data-si>${esc(boton)}</button></div>`;
  pie.querySelector('[data-no]').addEventListener('click',()=>{pie.innerHTML=antes;});
  const si=pie.querySelector('[data-si]');si.addEventListener('click',async()=>{si.disabled=true;try{await fn();}catch(e){si.disabled=false;toast(msgError(e));}});si.focus();
}
const stepper=(id,label,val)=>`<div class="campo"><label for="${id}">${label}</label><div class="stepper"><button type="button" data-paso="-1" data-campo="${id}" aria-label="Restar uno">−</button><input id="${id}" inputmode="numeric" pattern="[0-9]*" value="${val}" autocomplete="off"><button type="button" data-paso="1" data-campo="${id}" aria-label="Sumar uno">+</button></div></div>`;
const grupoOpc=(nombre,ops,actual,extra='')=>`<div class="opciones" data-grupo="${nombre}" data-valor="${esc(actual)}">${ops.map(([v,l])=>`<button type="button" class="opcion" data-valor="${esc(v)}" aria-pressed="${String(v)===String(actual)}" ${extra}>${l}</button>`).join('')}</div>`;
const pieDetalle=()=>puedeEditar()?`<button type="button" class="btn peligro" data-borrar>${ico('basura')}Eliminar</button><span class="flex1"></span><button type="button" class="btn primario" data-editar>${ico('lapiz')}Editar</button>`:`<button type="button" class="btn" data-act="cerrar">Cerrar</button>`;
const precioInput=(id,val)=>`<div class="pref"><span>$</span><input id="${id}" inputmode="decimal" value="${val==null?'':esc(val)}" placeholder="0" autocomplete="off"></div>`;
function errorForm(el,m){el.textContent=m;el.hidden=false;try{el.scrollIntoView({block:'nearest',behavior:'smooth'});}catch(e){}}

/* ---- Venta ---- */
function hojaVenta(venta,pre){
  if(!puedeEditar())return avisoSoloLectura();
  const d=derivar();
  const itemV=t=>venta?itemsDe(venta).find(i=>i.tipo===t):null;
  const cant=t=>{const it=itemV(t);return it?n0(it.charolas):0;};
  const totalV=venta?TIPOS.reduce((s,t)=>{const it=itemV(t);return s+(it?n0(it.charolas)*n0(it.precio):0);},0):0;
  const pagadoV=venta?(venta.pagado==null?totalV:n0(venta.pagado)):0;
  const st={clienteId:venta?(venta.clienteId||null):((pre&&pre.clienteId)||null),mostrador:!!(venta&&!venta.clienteId),nuevo:false,q:'',editado:{blanco:!!venta,marron:!!venta}};
  if(st.clienteId&&!d.clientesPorId.has(st.clienteId)){st.clienteId=null;st.mostrador=true;}
  const pago0=!venta?'pagado':(pagadoV>=totalV-0.009?'pagado':(pagadoV<=0.009?'fiado':'parcial'));
  const lineaTipo=t=>{const it=itemV(t);return `<div class="tipo-linea"><div class="tipo-cab"><b><span class="huevo ${t}"></span>${TIPO_NOMBRE[t]}</b><span class="ayuda" id="v-stock-${t}"></span></div><div class="fila2">${stepper('v-ch-'+t,'Charolas',cant(t))}<div class="campo"><label for="v-pr-${t}">Precio por charola</label>${precioInput('v-pr-'+t,it?n0(it.precio):null)}</div></div><p class="ayuda" id="v-sub-${t}"></p></div>`;};
  const cuerpo=`<form id="f-venta" class="form" novalidate>
    <div class="campo"><label for="v-fecha">Fecha de venta</label><input id="v-fecha" type="date" value="${esc(venta?venta.fecha:hoy())}"></div>
    <div class="campo"><span class="etq">Cliente</span><div id="v-cli"></div></div>
    ${TIPOS.map(lineaTipo).join('')}
    <div class="total-linea"><span>Total</span><b id="v-total" class="num">$0</b></div>
    <div class="campo"><span class="etq">Pago</span>${grupoOpc('pago',[['pagado','Pagado'],['fiado','Fiado'],['parcial','Pagó una parte']],pago0)}<p class="ayuda" id="v-pago-ayuda" hidden>Para fiar, elige un cliente.</p>
      <div class="campo" id="v-parcial-campo" ${pago0==='parcial'?'':'hidden'}><label for="v-pagado">¿Cuánto pagó?</label>${precioInput('v-pagado',pago0==='parcial'?r2(pagadoV):null)}</div></div>
    <div class="campo"><span class="etq">Método de pago</span>${grupoOpc('metodo',Object.entries(METODOS),(venta&&venta.metodo)||'efectivo')}</div>
    <div class="campo"><label for="v-notas">Notas</label><textarea id="v-notas" rows="2" placeholder="Opcional">${esc((venta&&venta.notas)||'')}</textarea></div>
    <p class="error" id="v-error" role="alert" hidden></p>
  </form>`;
  abrirHoja({titulo:venta?'Editar venta':'Nueva venta',cuerpo,formulario:true,
    pie:`<button type="button" class="btn" data-act="cerrar">Cancelar</button><button type="submit" form="f-venta" class="btn acento" id="v-guardar">${venta?'Guardar cambios':'Guardar venta'}</button>`,
    alMontar(h){
      const $h=s=>h.querySelector(s);
      const clienteObj=()=>{if(st.nuevo){const s=$h('#v-nc-tipo');return {tipo:s?s.value:'individual'};}return st.clienteId?d.clientesPorId.get(st.clienteId):null;};
      const aplicarPrecios=()=>{for(const t of TIPOS){if(st.editado[t])continue;const p=precioSugerido(d,t,clienteObj());$h('#v-pr-'+t).value=p!=null?String(p):'';}};
      function pintarOpciones(){
        const ops=$h('#v-cli-ops');if(!ops)return;const q=normal(st.q);
        let l=d.clientes.slice();
        if(q)l=l.filter(c=>normal(c.nombre).includes(q)||normal(c.saludo).includes(q));
        l.sort((a,b)=>{if(q)return a.nombre.localeCompare(b.nombre,'es');const fa=a.ult?a.ult.fecha:'',fb=b.ult?b.ult.fecha:'';return fa===fb?a.nombre.localeCompare(b.nombre,'es'):(fa<fb?1:-1);});
        l=l.slice(0,q?8:5);
        ops.innerHTML=l.map(c=>`<button type="button" class="cli-op" data-cli="${esc(c.id)}"><span class="inicial">${esc(inicial(c.nombre))}</span><span><b>${esc(c.nombre)}</b><small>${esc(CLI_TIPOS[c.tipo])}${c.ult?' · compró '+esc(fRel(c.ult.fecha)):''}</small></span></button>`).join('')
          +`<button type="button" class="cli-op" data-cli-accion="nuevo"><span class="inicial">${ico('mas')}</span><span><b>${st.q.trim()?'Nuevo cliente: '+esc(st.q.trim()):'Nuevo cliente'}</b><small>Darlo de alta con esta venta</small></span></button>`
          +`<button type="button" class="cli-op" data-cli-accion="mostrador"><span class="inicial">${ico('ventas')}</span><span><b>Venta de mostrador</b><small>Sin nombre ni recordatorio</small></span></button>`;
      }
      function pintarCliente(){
        const cont=$h('#v-cli');
        if(st.nuevo){cont.innerHTML=`<div class="cli-nuevo"><div class="campo"><label for="v-nc-nombre">Nombre del cliente nuevo</label><input id="v-nc-nombre" value="${esc(st.q.trim())}" autocomplete="off" placeholder="Ej. Tienda Don Beto"></div><div class="fila2 apilar"><div class="campo"><label for="v-nc-tel">WhatsApp</label><input id="v-nc-tel" type="tel" inputmode="tel" placeholder="10 dígitos" autocomplete="off"></div><div class="campo"><label for="v-nc-tipo">Tipo</label><div class="sel"><select id="v-nc-tipo">${Object.entries(CLI_TIPOS).map(([k,l])=>`<option value="${k}">${l}</option>`).join('')}</select></div></div></div><p class="ayuda">Te recordaré escribirle cada ${d.c.recordarDias} días después de su última compra; lo puedes cambiar en su ficha.</p><button type="button" class="enlace" data-cli-accion="cancelar-nuevo">Mejor elegir un cliente existente</button></div>`;return;}
        if(st.clienteId){const c=d.clientesPorId.get(st.clienteId);cont.innerHTML=`<div class="cli-elegido"><span class="inicial">${esc(inicial(c&&c.nombre))}</span><div><b>${esc(c?c.nombre:'Cliente')}</b><small>${esc(c?CLI_TIPOS[c.tipo]:'')}${c&&c.saldo>0.009?' · debe '+esc(dinero(c.saldo)):''}</small></div><button type="button" class="btn chico" data-cli-accion="cambiar">Cambiar</button></div>`;return;}
        if(st.mostrador){cont.innerHTML=`<div class="cli-elegido"><span class="inicial">${ico('ventas')}</span><div><b>Venta de mostrador</b><small>${venta&&venta.clienteNombre?esc(venta.clienteNombre)+' · ':''}Sin cliente ni recordatorio</small></div><button type="button" class="btn chico" data-cli-accion="cambiar">Cambiar</button></div>`;return;}
        cont.innerHTML=`<div class="cli-busca"><label class="buscar"><span class="sr">Buscar cliente</span>${ico('buscar')}<input id="v-cli-q" type="search" placeholder="Buscar o escribir un nombre" value="${esc(st.q)}" autocomplete="off"></label><div class="cli-opciones" id="v-cli-ops"></div></div>`;
        pintarOpciones();
      }
      function actualizar(){
        let total=0;
        for(const t of TIPOS){
          const ch=entero($h('#v-ch-'+t).value),pr=decimal($h('#v-pr-'+t).value),ok=pr!=null&&!isNaN(pr)&&pr>=0,sub=ch*(ok?pr:0);total+=sub;
          const disp=d.inv[t].charolas+cant(t);
          $h('#v-stock-'+t).textContent=chTxt(disp)+' en bodega';
          $h('#v-sub-'+t).innerHTML=ch>0?`${num(ch)} × ${ok?dinero(pr):'$—'} = <b>${dinero(sub)}</b>${ch>disp?` <span class="warn">· Solo hay ${num(disp)} en bodega</span>`:''}`:'';
        }
        $h('#v-total').textContent=dinero(total);
        const sinCli=!st.clienteId&&!st.nuevo,g=$h('[data-grupo="pago"]');
        g.querySelectorAll('[data-valor]').forEach(b=>{b.disabled=sinCli&&b.dataset.valor!=='pagado';});
        if(sinCli&&g.dataset.valor!=='pagado'){g.dataset.valor='pagado';g.querySelectorAll('[data-valor]').forEach(b=>b.setAttribute('aria-pressed',String(b.dataset.valor==='pagado')));}
        $h('#v-pago-ayuda').hidden=!sinCli;
        $h('#v-parcial-campo').hidden=g.dataset.valor!=='parcial';
      }
      h.addEventListener('click',e=>{
        const op=e.target.closest('[data-cli]');
        if(op){st.clienteId=op.dataset.cli;st.mostrador=false;st.nuevo=false;pintarCliente();aplicarPrecios();actualizar();return;}
        const ac=e.target.closest('[data-cli-accion]');
        if(ac){const a=ac.dataset.cliAccion;
          if(a==='cambiar'){st.clienteId=null;st.mostrador=false;st.nuevo=false;}
          else if(a==='nuevo'){st.nuevo=true;st.mostrador=false;st.clienteId=null;}
          else if(a==='mostrador'){st.mostrador=true;st.nuevo=false;st.clienteId=null;}
          else if(a==='cancelar-nuevo'){st.nuevo=false;}
          pintarCliente();aplicarPrecios();actualizar();
          const foco=a==='nuevo'?$h('#v-nc-nombre'):(a==='cambiar'||a==='cancelar-nuevo')?$h('#v-cli-q'):null;if(foco)foco.focus();
        }
      });
      h.addEventListener('input',e=>{const id=e.target.id||'';if(id==='v-cli-q'){st.q=e.target.value;pintarOpciones();return;}if(id.indexOf('v-pr-')===0)st.editado[id.slice(5)]=true;actualizar();});
      h.addEventListener('change',e=>{if(e.target.id==='v-nc-tipo'){aplicarPrecios();actualizar();}});
      h.addEventListener('grupo',actualizar);
      $h('#f-venta').addEventListener('submit',e=>{e.preventDefault();guardar();});
      pintarCliente();if(!venta)aplicarPrecios();actualizar();
      async function guardar(){
        const err=$h('#v-error'),btn=document.getElementById('v-guardar');err.hidden=true;
        const fallo=m=>errorForm(err,m);
        const fecha=$h('#v-fecha').value;if(!esFecha(fecha))return fallo('Elige la fecha de la venta.');
        if(!st.clienteId&&!st.nuevo&&!st.mostrador)return fallo('Elige el cliente, o «Venta de mostrador» si no quieres registrarlo.');
        const items=[];
        for(const t of TIPOS){const ch=entero($h('#v-ch-'+t).value);if(ch<=0)continue;const pr=decimal($h('#v-pr-'+t).value);if(pr==null||isNaN(pr)||pr<0)return fallo(`Escribe el precio por charola de ${TIPO_NOMBRE[t].toLowerCase()}.`);items.push({tipo:t,charolas:ch,precio:r2(pr)});}
        if(!items.length)return fallo('Agrega al menos una charola de huevo blanco o marrón.');
        const total=r2(sum(items,it=>it.charolas*it.precio));
        const pago=$h('[data-grupo="pago"]').dataset.valor;let pagado=total;
        if(pago==='fiado')pagado=0;
        else if(pago==='parcial'){const p=decimal($h('#v-pagado').value);if(p==null||isNaN(p)||p<0||p>total)return fallo('Lo pagado debe estar entre $0 y el total de la venta.');pagado=r2(p);}
        let nc=null;
        if(st.nuevo){const nombre=$h('#v-nc-nombre').value.trim();if(!nombre)return fallo('Escribe el nombre del cliente nuevo.');const tel=$h('#v-nc-tel').value.trim(),dg=tel.replace(/\D/g,'');if(dg&&dg.length<10)return fallo('El WhatsApp debe tener 10 dígitos.');nc={nombre,saludo:'',telefono:tel,tipo:$h('#v-nc-tipo').value,zona:'',recordarCada:d.c.recordarDias,preferencia:'ambos',notas:'',alta:hoy(),creado:ahora()};}
        if(pago!=='pagado'&&!st.clienteId&&!nc)return fallo('Para fiar, elige un cliente.');
        if(btn)btn.disabled=true;
        try{
          let clienteId=st.clienteId,clienteNombre='';
          if(nc){clienteId=await guardarDoc('clientes',null,nc);clienteNombre=nc.nombre;}
          else if(clienteId){const c=d.clientesPorId.get(clienteId);clienteNombre=c?c.nombre:'';}
          else if(venta&&venta.clienteNombre)clienteNombre=venta.clienteNombre;
          const doc={fecha,clienteId:clienteId||null,clienteNombre,items,pagado,metodo:$h('[data-grupo="metodo"]').dataset.valor||'efectivo',notas:$h('#v-notas').value.trim(),creado:(venta&&venta.creado)||ahora()};
          if(venta)doc.editado=ahora();
          await guardarDoc('ventas',venta?venta.id:null,doc);
          if(clienteId&&!nc){const raw=D().clientes.get(clienteId);if(raw&&raw.posponerHasta){try{await actualizarDoc('clientes',clienteId,{posponerHasta:null});}catch(e2){}}}
          cerrarHoja();
          let msg=venta?'Venta actualizada.':'Venta guardada.';
          const c=clienteId?d.clientesPorId.get(clienteId):null,rec=nc?nc.recordarCada:(c?c.rec:0);
          if(clienteId&&rec>0){const base=[fecha,c&&c.ult&&c.ult.fecha,c&&esFecha(c.ultimoContacto)?c.ultimoContacto:null].filter(Boolean).sort().pop();msg+=` Te recordaré escribirle a ${clienteNombre} el ${fFecha(sumarDias(base,rec))}.`;}
          toast(msg);
        }catch(e){if(btn)btn.disabled=false;fallo(msgError(e));}
      }
    }});
}
/* ---- Compra ---- */
function hojaCompra(compra){
  if(!puedeEditar())return avisoSoloLectura();
  const d=derivar();const itemC=t=>compra?itemsDe(compra).find(i=>i.tipo===t):null;
  const ultimoProv=d.compras.length?(d.compras[d.compras.length-1].proveedor||''):'';
  const linea=t=>{const it=itemC(t);const pr=it?n0(it.precio):precioCompraSugerido(d,t);return `<div class="tipo-linea"><div class="tipo-cab"><b><span class="huevo ${t}"></span>${TIPO_NOMBRE[t]}</b></div><div class="fila2">${stepper('c-ch-'+t,'Charolas',it?n0(it.charolas):0)}<div class="campo"><label for="c-pr-${t}">Precio por charola</label>${precioInput('c-pr-'+t,pr)}</div></div><div class="fila2">${stepper('c-dan-'+t,'Piezas dañadas',it?n0(it.danadas):0)}<p class="ayuda" style="align-self:end;padding-bottom:4px">Rotas o estrelladas al recibir; se restan del inventario.</p></div><p class="ayuda" id="c-sub-${t}"></p></div>`;};
  const cuerpo=`<form id="f-compra" class="form" novalidate>
    <div class="fila2 apilar"><div class="campo"><label for="c-fecha">Fecha de compra</label><input id="c-fecha" type="date" value="${esc(compra?compra.fecha:hoy())}"></div><div class="campo"><label for="c-prov">Granja o proveedor</label><input id="c-prov" value="${esc(compra?(compra.proveedor||''):(ultimoProv||d.c.proveedor||''))}" placeholder="Ej. Granja San Isidro" autocomplete="off"></div></div>
    ${TIPOS.map(linea).join('')}
    <div class="total-linea"><span>Total de la compra</span><b id="c-total" class="num">$0</b></div>
    <div class="campo"><label for="c-notas">Notas</label><textarea id="c-notas" rows="2" placeholder="Opcional">${esc((compra&&compra.notas)||'')}</textarea></div>
    <p class="error" id="c-error" role="alert" hidden></p></form>`;
  abrirHoja({titulo:compra?'Editar compra':'Nueva compra a granja',cuerpo,formulario:true,
    pie:`<button type="button" class="btn" data-act="cerrar">Cancelar</button><button type="submit" form="f-compra" class="btn acento" id="c-guardar">${compra?'Guardar cambios':'Guardar compra'}</button>`,
    alMontar(h){
      const $h=s=>h.querySelector(s);
      const actualizar=()=>{let total=0;for(const t of TIPOS){const ch=entero($h('#c-ch-'+t).value),pr=decimal($h('#c-pr-'+t).value),dan=entero($h('#c-dan-'+t).value);const ok=pr!=null&&!isNaN(pr)&&pr>=0,sub=ch*(ok?pr:0);total+=sub;$h('#c-sub-'+t).innerHTML=ch>0?`${num(ch)} × ${ok?dinero(pr):'$—'} = <b>${dinero(sub)}</b>${dan>0?` · ${num(dan)} ${dan===1?'pieza dañada':'piezas dañadas'}${ok?' ('+dinero(dan*pr/PZ)+')':''}`:''}${dan>ch*PZ?' <span class="warn">· Son más piezas que las compradas</span>':''}`:(dan>0?'<span class="warn">Agrega las charolas compradas.</span>':'');}$h('#c-total').textContent=dinero(total);};
      h.addEventListener('input',actualizar);actualizar();
      $h('#f-compra').addEventListener('submit',async e=>{
        e.preventDefault();const err=$h('#c-error'),btn=document.getElementById('c-guardar');err.hidden=true;const fallo=m=>errorForm(err,m);
        const fecha=$h('#c-fecha').value;if(!esFecha(fecha))return fallo('Elige la fecha de la compra.');
        const items=[];
        for(const t of TIPOS){const ch=entero($h('#c-ch-'+t).value),dan=entero($h('#c-dan-'+t).value);if(ch<=0){if(dan>0)return fallo(`Agrega las charolas de ${TIPO_NOMBRE[t].toLowerCase()} o quita las piezas dañadas.`);continue;}const pr=decimal($h('#c-pr-'+t).value);if(pr==null||isNaN(pr)||pr<0)return fallo(`Escribe el precio por charola de ${TIPO_NOMBRE[t].toLowerCase()}.`);if(dan>ch*PZ)return fallo('Las piezas dañadas no pueden ser más que las compradas.');items.push({tipo:t,charolas:ch,precio:r2(pr),danadas:dan});}
        if(!items.length)return fallo('Agrega al menos una charola comprada.');
        if(btn)btn.disabled=true;
        try{await guardarDoc('compras',compra?compra.id:null,{fecha,proveedor:$h('#c-prov').value.trim(),items,notas:$h('#c-notas').value.trim(),creado:(compra&&compra.creado)||ahora(),...(compra?{editado:ahora()}:{})});cerrarHoja();toast(`${compra?'Compra actualizada':'Compra guardada'}: ${chTxt(sum(items,'charolas'))}.`);}
        catch(e2){if(btn)btn.disabled=false;fallo(msgError(e2));}
      });
    }});
}
/* ---- Merma ---- */
function hojaMerma(m){
  if(!puedeEditar())return avisoSoloLectura();
  const d=derivar(),t0=(m&&m.tipo)||'blanco',mot0=(m&&m.motivo)||MOTIVOS[0];
  const cuerpo=`<form id="f-merma" class="form" novalidate>
    <div class="campo"><label for="m-fecha">Fecha</label><input id="m-fecha" type="date" value="${esc(m?m.fecha:hoy())}"></div>
    <div class="campo"><span class="etq">Tipo de huevo</span>${grupoOpc('tipo',TIPOS.map(t=>[t,`<span class="huevo ${t}"></span> ${cap(TIPO_CORTO[t])}`]),t0)}<p class="ayuda" id="m-stock"></p></div>
    ${stepper('m-pz','Piezas dañadas o no vendibles',m?n0(m.piezas):0)}
    <div class="opciones">${[6,12,30].map(n=>`<button type="button" class="opcion" data-sumar="${n}">+${n}${n===30?' (1 charola)':''}</button>`).join('')}</div>
    <p class="ayuda" id="m-desglose"></p>
    <div class="campo"><span class="etq">Motivo</span>${grupoOpc('motivo',MOTIVOS.map(x=>[x,esc(x)]),mot0)}</div>
    <div class="campo"><label for="m-notas">Notas</label><textarea id="m-notas" rows="2" placeholder="Opcional">${esc((m&&m.notas)||'')}</textarea></div>
    <p class="error" id="m-error" role="alert" hidden></p></form>`;
  abrirHoja({titulo:m?'Editar merma':'Registrar merma',cuerpo,formulario:true,
    pie:`<button type="button" class="btn" data-act="cerrar">Cancelar</button><button type="submit" form="f-merma" class="btn acento" id="m-guardar">${m?'Guardar cambios':'Guardar merma'}</button>`,
    alMontar(h){
      const $h=s=>h.querySelector(s);
      const actualizar=()=>{const t=$h('[data-grupo="tipo"]').dataset.valor,pz=entero($h('#m-pz').value);const disp=Math.max(0,d.inv[t].piezas)+(m&&m.tipo===t?n0(m.piezas):0);$h('#m-stock').textContent='En bodega: '+desglose(disp);$h('#m-desglose').innerHTML=pz>0?`Equivale a ${esc(desglose(pz))}${pz>disp?' <span class="warn">· Es más de lo que hay en bodega</span>':''}`:'';};
      h.addEventListener('click',e=>{const b=e.target.closest('[data-sumar]');if(b){const inp=$h('#m-pz');inp.value=String(entero(inp.value)+Number(b.dataset.sumar));actualizar();}});
      h.addEventListener('input',actualizar);h.addEventListener('grupo',actualizar);actualizar();
      $h('#f-merma').addEventListener('submit',async e=>{
        e.preventDefault();const err=$h('#m-error'),btn=document.getElementById('m-guardar');err.hidden=true;const fallo=x=>errorForm(err,x);
        const fecha=$h('#m-fecha').value;if(!esFecha(fecha))return fallo('Elige la fecha.');
        const t=$h('[data-grupo="tipo"]').dataset.valor,pz=entero($h('#m-pz').value);if(pz<=0)return fallo('Escribe cuántas piezas se dañaron.');
        if(btn)btn.disabled=true;
        try{await guardarDoc('mermas',m?m.id:null,{fecha,tipo:t,piezas:pz,motivo:$h('[data-grupo="motivo"]').dataset.valor,notas:$h('#m-notas').value.trim(),creado:(m&&m.creado)||ahora(),...(m?{editado:ahora()}:{})});cerrarHoja();toast(`Merma registrada: ${desglose(pz)} de huevo ${TIPO_CORTO[t]}.`);}
        catch(e2){if(btn)btn.disabled=false;fallo(msgError(e2));}
      });
    }});
}
/* ---- Pago (abono) ---- */
function hojaAbono(a,clienteId){
  if(!puedeEditar())return avisoSoloLectura();
  const d=derivar();if(!d.clientes.length){toast('Primero agrega un cliente.');return;}
  const con=d.clientes.filter(x=>x.saldo>0.009).sort((x,y)=>y.saldo-x.saldo);
  const otros=d.clientes.filter(x=>!(x.saldo>0.009)).sort((x,y)=>x.nombre.localeCompare(y.nombre,'es'));
  const sel=(a&&a.clienteId)||clienteId||(con[0]&&con[0].id)||'';
  const opt=x=>`<option value="${esc(x.id)}" ${x.id===sel?'selected':''}>${esc(x.nombre)}${x.saldo>0.009?' — debe '+esc(dinero(x.saldo)):''}</option>`;
  const cuerpo=`<form id="f-abono" class="form" novalidate>
    <div class="campo"><label for="a-cli">Cliente</label><div class="sel"><select id="a-cli">${sel?'':'<option value="">Elige un cliente</option>'}${con.length?`<optgroup label="Con saldo pendiente">${con.map(opt).join('')}</optgroup>`:''}${otros.length?`<optgroup label="${con.length?'Otros clientes':'Clientes'}">${otros.map(opt).join('')}</optgroup>`:''}</select></div><p class="ayuda" id="a-saldo"></p></div>
    <div class="fila2 apilar"><div class="campo"><label for="a-monto">Monto recibido</label>${precioInput('a-monto',a?r2(a.monto):null)}</div><div class="campo"><label for="a-fecha">Fecha</label><input id="a-fecha" type="date" value="${esc(a?a.fecha:hoy())}"></div></div>
    <div class="campo"><span class="etq">Método</span>${grupoOpc('metodo',Object.entries(METODOS),(a&&a.metodo)||'efectivo')}</div>
    <div class="campo"><label for="a-nota">Nota</label><input id="a-nota" value="${esc((a&&a.nota)||'')}" placeholder="Opcional" autocomplete="off"></div>
    <p class="error" id="a-error" role="alert" hidden></p></form>`;
  abrirHoja({titulo:a?'Editar pago':'Registrar pago',cuerpo,formulario:true,
    pie:`<button type="button" class="btn" data-act="cerrar">Cancelar</button><button type="submit" form="f-abono" class="btn acento" id="a-guardar">${a?'Guardar cambios':'Guardar pago'}</button>`,
    alMontar(h){
      const $h=s=>h.querySelector(s);let editado=!!a;
      const actualizar=prefill=>{const c=d.clientesPorId.get($h('#a-cli').value),s=$h('#a-saldo');if(!c){s.textContent='';return;}s.textContent=c.saldo>0.009?`Debe ${dinero(c.saldo)}`:(c.saldo<-0.009?`Tiene ${dinero(-c.saldo)} a favor`:'No tiene saldo pendiente');if(prefill&&!editado)$h('#a-monto').value=c.saldo>0.009?String(r2(c.saldo)):'';};
      $h('#a-cli').addEventListener('change',()=>actualizar(true));
      $h('#a-monto').addEventListener('input',()=>{editado=true;});
      actualizar(true);
      $h('#f-abono').addEventListener('submit',async e=>{
        e.preventDefault();const err=$h('#a-error'),btn=document.getElementById('a-guardar');err.hidden=true;const fallo=x=>errorForm(err,x);
        const c=d.clientesPorId.get($h('#a-cli').value);if(!c)return fallo('Elige el cliente que pagó.');
        const monto=decimal($h('#a-monto').value);if(monto==null||isNaN(monto)||monto<=0)return fallo('Escribe el monto que recibiste.');
        const fecha=$h('#a-fecha').value;if(!esFecha(fecha))return fallo('Elige la fecha del pago.');
        if(btn)btn.disabled=true;
        try{await guardarDoc('abonos',a?a.id:null,{clienteId:c.id,clienteNombre:c.nombre,monto:r2(monto),fecha,metodo:$h('[data-grupo="metodo"]').dataset.valor||'efectivo',nota:$h('#a-nota').value.trim(),creado:(a&&a.creado)||ahora(),...(a?{editado:ahora()}:{})});cerrarHoja();const resta=r2(c.saldo-monto+(a?r2(a.monto):0));toast(`Pago de ${dinero(monto)} registrado.${resta>0.009?' Aún debe '+dinero(resta)+'.':' Quedó al corriente.'}`);}
        catch(e2){if(btn)btn.disabled=false;fallo(msgError(e2));}
      });
    }});
}
/* ---- Cliente ---- */
function hojaCliente(cl){
  if(!puedeEditar())return avisoSoloLectura();
  const c=cfg(),ops=[3,5,7,10,15,30];
  const rec=cl?((cl.recordarCada===0||cl.recordarCada==='0')?0:(Math.round(n0(cl.recordarCada))||c.recordarDias)):c.recordarDias;
  const esOtro=rec>0&&!ops.includes(rec),tipo0=(cl&&CLI_TIPOS[cl.tipo])?cl.tipo:'individual',pref0=(cl&&cl.preferencia)||'ambos';
  const pe=(cl&&cl.precioEspecial)||{};
  const cuerpo=`<form id="f-cliente" class="form" novalidate>
    <div class="campo"><label for="cl-nombre">Nombre</label><input id="cl-nombre" value="${esc((cl&&cl.nombre)||'')}" autocomplete="off" placeholder="Ej. Abarrotes La Esperanza"></div>
    <div class="campo"><label for="cl-saludo">¿Cómo le saludas?</label><input id="cl-saludo" value="${esc((cl&&cl.saludo)||'')}" autocomplete="off" placeholder="Ej. Don Ramiro"><p class="ayuda">Se usa en el mensaje de WhatsApp. Si lo dejas vacío, se usa el nombre.</p></div>
    <div class="campo"><span class="etq">Tipo de cliente</span>${grupoOpc('tipo',Object.entries(CLI_TIPOS),tipo0)}</div>
    <div class="fila2 apilar"><div class="campo"><label for="cl-tel">WhatsApp</label><input id="cl-tel" type="tel" inputmode="tel" value="${esc((cl&&cl.telefono)||'')}" placeholder="10 dígitos" autocomplete="off"></div><div class="campo"><label for="cl-zona">Zona o dirección</label><input id="cl-zona" value="${esc((cl&&cl.zona)||'')}" placeholder="Opcional" autocomplete="off"></div></div>
    <div class="campo"><span class="etq">Recordarme escribirle</span>${grupoOpc('rec',[...ops.map(n=>[n,`Cada ${n} días`]),['otro','Otro'],[0,'No recordar']],esOtro?'otro':rec)}
      <div class="campo otro-rec" id="cl-rec-otro-c" ${esOtro?'':'hidden'}><label for="cl-rec-otro">Cada cuántos días</label><input id="cl-rec-otro" inputmode="numeric" value="${esOtro?rec:''}" placeholder="Ej. 21" autocomplete="off"></div>
      <p class="ayuda">Se cuenta desde su última compra. Si le escribes y no compra, vuelve a contar desde ese día.</p></div>
    <div class="campo"><span class="etq">Huevo que suele pedir</span>${grupoOpc('pref',[['blanco','Blanco'],['marron','Marrón'],['ambos','Ambos']],pref0)}</div>
    <details class="mas" ${n0(pe.blanco)>0||n0(pe.marron)>0?'open':''}><summary>Precio especial para este cliente</summary><div class="fila2">${TIPOS.map(t=>`<div class="campo"><label for="cl-pe-${t}">${TIPO_NOMBRE[t]}</label><div class="pref"><span>$</span><input id="cl-pe-${t}" inputmode="decimal" value="${n0(pe[t])>0?esc(pe[t]):''}" placeholder="Lista" autocomplete="off"></div></div>`).join('')}</div><p class="ayuda" style="margin-top:8px">Si lo dejas vacío se usa la lista de precios de Ajustes.</p></details>
    <div class="campo"><label for="cl-notas">Notas</label><textarea id="cl-notas" rows="2" placeholder="Ej. Paga los viernes, entregar después de las 5">${esc((cl&&cl.notas)||'')}</textarea></div>
    <p class="error" id="cl-error" role="alert" hidden></p></form>`;
  abrirHoja({titulo:cl?'Editar cliente':'Nuevo cliente',cuerpo,formulario:true,
    pie:`<button type="button" class="btn" data-act="cerrar">Cancelar</button><button type="submit" form="f-cliente" class="btn acento" id="cl-guardar">${cl?'Guardar cambios':'Guardar cliente'}</button>`,
    alMontar(h){
      const $h=s=>h.querySelector(s);
      h.addEventListener('grupo',e=>{if(e.detail.grupo==='rec'){const otro=e.detail.valor==='otro';$h('#cl-rec-otro-c').hidden=!otro;if(otro)$h('#cl-rec-otro').focus();}});
      $h('#f-cliente').addEventListener('submit',async e=>{
        e.preventDefault();const err=$h('#cl-error'),btn=document.getElementById('cl-guardar');err.hidden=true;const fallo=x=>errorForm(err,x);
        const nombre=$h('#cl-nombre').value.trim();if(!nombre)return fallo('Escribe el nombre del cliente.');
        const tel=$h('#cl-tel').value.trim(),dg=tel.replace(/\D/g,'');if(dg&&dg.length<10)return fallo('El WhatsApp debe tener 10 dígitos.');
        const rv=$h('[data-grupo="rec"]').dataset.valor;let recordarCada=rv==='otro'?entero($h('#cl-rec-otro').value):Number(rv);
        if(rv==='otro'&&recordarCada<1)return fallo('Escribe cada cuántos días quieres el recordatorio.');
        const precioEspecial={};for(const t of TIPOS){const v=decimal($h('#cl-pe-'+t).value);if(v!=null&&(isNaN(v)||v<0))return fallo('Revisa el precio especial: solo números.');precioEspecial[t]=v==null?null:r2(v);}
        const base=cl?{...(D().clientes.get(cl.id)||{})}:{};delete base.id;
        const doc={...base,nombre,saludo:$h('#cl-saludo').value.trim(),tipo:$h('[data-grupo="tipo"]').dataset.valor,telefono:tel,zona:$h('#cl-zona').value.trim(),recordarCada,preferencia:$h('[data-grupo="pref"]').dataset.valor,precioEspecial,notas:$h('#cl-notas').value.trim(),alta:base.alta||hoy(),creado:base.creado||ahora()};
        if(cl)doc.editado=ahora();
        if(btn)btn.disabled=true;
        try{await guardarDoc('clientes',cl?cl.id:null,doc);cerrarHoja();toast(cl?'Cliente actualizado.':`Cliente agregado: ${nombre}.`);}
        catch(e2){if(btn)btn.disabled=false;fallo(msgError(e2));}
      });
    }});
}
/* ---- Detalles ---- */
function hojaDetalleVenta(id){
  const d=derivar(),v=d.ventas.find(x=>x.id===id);if(!v){toast('Esa venta ya no existe.');return;}
  const filas=v.items.map(it=>`<tr><td><span class="hv"><span class="huevo ${it.tipo}"></span>${esc(TIPO_NOMBRE[it.tipo])}</span></td><td class="n">${num(it.charolas)}</td><td class="n">${dinero(it.precio)}</td><td class="n">${dinero(it.charolas*it.precio)}</td></tr>`).join('');
  const pago=v.pendiente>0.009?`Debe ${dinero(v.pendiente)}`:(v.pagadoMomento<v.total-0.009?'Pagada con abonos':'Pagada');
  const cuerpo=`<div class="detalle"><div class="det-cab"><div><p class="ceja">${esc(fFechaLarga(v.fecha))}</p><h3 class="det-tit">${v.cli?`<button type="button" class="enlace-tit" data-act="cliente" data-id="${esc(v.cli.id)}">${esc(v.nombre)}</button>`:esc(v.nombre)}</h3></div><b class="det-monto">${dinero(v.total)}</b></div>
    <div class="tabla-env"><table><thead><tr><th>Tipo</th><th class="n">Charolas</th><th class="n">Precio</th><th class="n">Subtotal</th></tr></thead><tbody>${filas}</tbody></table></div>
    <dl class="datos"><div><dt>Costo de las charolas</dt><dd>${dinero(v.costo)}</dd></div><div><dt>Ganancia</dt><dd class="${v.ganancia>=0?'pos':'neg'}">${dinero(v.ganancia)}</dd></div><div><dt>Pago</dt><dd class="${v.pendiente>0.009?'neg':''}">${esc(pago)}</dd></div><div><dt>Método</dt><dd>${esc(METODOS[v.metodo]||'—')}</dd></div>${v.notas?`<div class="ancho"><dt>Notas</dt><dd>${esc(v.notas)}</dd></div>`:''}</dl>
    <p class="ayuda">El costo usa el precio de compra de los lotes más antiguos que había en bodega.</p></div>`;
  abrirHoja({titulo:'Venta',cuerpo,pie:pieDetalle(),alMontar(h){h.addEventListener('click',e=>{if(e.target.closest('[data-editar]')){const raw=D().ventas.get(id);if(raw)hojaVenta(raw);}else if(e.target.closest('[data-borrar]'))confirmarPie(h,'¿Eliminar esta venta? No se puede deshacer.','Sí, eliminar',async()=>{await borrarDoc('ventas',id);cerrarHoja();toast('Venta eliminada.');});});}});
}
function hojaDetalleCompra(id){
  const d=derivar(),c=d.compras.find(x=>x.id===id);if(!c){toast('Esa compra ya no existe.');return;}
  const lotes=TIPOS.map(t=>{const L=d.inv[t].lotes.find(l=>l.compraId===id);return L?{...L,tipo:t}:null;}).filter(Boolean);
  const filas=c.items.map(it=>`<tr><td><span class="hv"><span class="huevo ${it.tipo}"></span>${esc(TIPO_NOMBRE[it.tipo])}</span></td><td class="n">${num(it.charolas)}</td><td class="n">${dinero(it.precio)}</td><td class="n">${num(it.danadas)}</td><td class="n">${dinero(it.charolas*it.precio)}</td></tr>`).join('');
  const cuerpo=`<div class="detalle"><div class="det-cab"><div><p class="ceja">${esc(fFechaLarga(c.fecha))}</p><h3 class="det-tit">${esc(c.proveedor||'Compra a granja')}</h3></div><b class="det-monto">${dinero(c.total)}</b></div>
    <div class="tabla-env"><table><thead><tr><th>Tipo</th><th class="n">Charolas</th><th class="n">Precio</th><th class="n">Dañadas</th><th class="n">Subtotal</th></tr></thead><tbody>${filas}</tbody></table></div>
    <dl class="datos"><div><dt>Piezas dañadas</dt><dd>${num(c.danadasPz)}${c.danadasPz?' · '+dinero(c.costoDanadas):''}</dd></div><div><dt>Quedan en bodega</dt><dd>${lotes.length?lotes.map(L=>esc(desglose(L.restante)+' de '+TIPO_CORTO[L.tipo])).join('<br>'):'Ya salió todo'}</dd></div>${c.notas?`<div class="ancho"><dt>Notas</dt><dd>${esc(c.notas)}</dd></div>`:''}</dl></div>`;
  abrirHoja({titulo:'Compra a granja',cuerpo,pie:pieDetalle(),alMontar(h){h.addEventListener('click',e=>{if(e.target.closest('[data-editar]')){const raw=D().compras.get(id);if(raw)hojaCompra(raw);}else if(e.target.closest('[data-borrar]'))confirmarPie(h,'¿Eliminar esta compra? El inventario se recalcula.','Sí, eliminar',async()=>{await borrarDoc('compras',id);cerrarHoja();toast('Compra eliminada.');});});}});
}
function hojaDetalleMerma(id){
  const d=derivar(),m=d.mermas.find(x=>x.id===id);if(!m){toast('Ese registro ya no existe.');return;}
  const cuerpo=`<div class="detalle"><div class="det-cab"><div><p class="ceja">${esc(fFechaLarga(m.fecha))}</p><h3 class="det-tit">${esc(m.motivo||'Merma')}</h3></div><b class="det-monto">${pzas(m.piezas)}</b></div>
    <dl class="datos"><div><dt>Tipo</dt><dd><span class="hv"><span class="huevo ${m.tipo}"></span>${esc(TIPO_NOMBRE[m.tipo]||'')}</span></dd></div><div><dt>Equivale a</dt><dd>${esc(desglose(m.piezas))}</dd></div><div><dt>Costo perdido</dt><dd class="neg">${dinero(m.costo)}</dd></div>${m.notas?`<div class="ancho"><dt>Notas</dt><dd>${esc(m.notas)}</dd></div>`:''}</dl></div>`;
  abrirHoja({titulo:'Merma',cuerpo,pie:pieDetalle(),alMontar(h){h.addEventListener('click',e=>{if(e.target.closest('[data-editar]')){const raw=D().mermas.get(id);if(raw)hojaMerma(raw);}else if(e.target.closest('[data-borrar]'))confirmarPie(h,'¿Eliminar este registro de merma?','Sí, eliminar',async()=>{await borrarDoc('mermas',id);cerrarHoja();toast('Registro eliminado.');});});}});
}
function hojaDetalleAbono(id){
  const d=derivar(),a=d.abonos.find(x=>x.id===id);if(!a){toast('Ese pago ya no existe.');return;}
  const cl=d.clientesPorId.get(a.clienteId);
  const cuerpo=`<div class="detalle"><div class="det-cab"><div><p class="ceja">${esc(fFechaLarga(a.fecha))}</p><h3 class="det-tit">${cl?`<button type="button" class="enlace-tit" data-act="cliente" data-id="${esc(cl.id)}">${esc(a.nombre)}</button>`:esc(a.nombre)}</h3></div><b class="det-monto pos">+${dinero(a.monto)}</b></div>
    <dl class="datos"><div><dt>Método</dt><dd>${esc(METODOS[a.metodo]||'Efectivo')}</dd></div><div><dt>Saldo actual del cliente</dt><dd>${cl?(cl.saldo>0.009?'Debe '+dinero(cl.saldo):'Al corriente'):'—'}</dd></div>${a.nota?`<div class="ancho"><dt>Nota</dt><dd>${esc(a.nota)}</dd></div>`:''}</dl></div>`;
  abrirHoja({titulo:'Pago recibido',cuerpo,pie:pieDetalle(),alMontar(h){h.addEventListener('click',e=>{if(e.target.closest('[data-editar]')){const raw=D().abonos.get(id);if(raw)hojaAbono(raw);}else if(e.target.closest('[data-borrar]'))confirmarPie(h,'¿Eliminar este pago? El saldo del cliente se recalcula.','Sí, eliminar',async()=>{await borrarDoc('abonos',id);cerrarHoja();toast('Pago eliminado.');});});}});
}
function precioEspTxt(x){const p=x.precioEspecial||{};const a=TIPOS.filter(t=>n0(p[t])>0).map(t=>`${cap(TIPO_CORTO[t])} ${dinero(p[t])}`);return a.length?esc(a.join(' · ')):'Lista de precios';}
function cuerpoCliente(id){
  const d=derivar(),x=d.clientesPorId.get(id);if(!x)return '<p>Este cliente ya no existe.</p>';
  const pe=puedeEditar(),ops=[3,5,7,10,15,30],dis=pe?'':'disabled';
  const pend=x.estado==='atrasado'||x.estado==='hoy'||x.estado==='pronto';
  const sug=x.frecuencia&&x.rec>0&&Math.abs(x.frecuencia-x.rec)>=2&&pe?`<div class="sugerencia">${ico('reloj')}<span>Suele comprar cada <b>${x.frecuencia} días</b>.</span><button type="button" class="btn chico" data-rec="${x.frecuencia}">Usar ${x.frecuencia} días</button></div>`:(x.frecuencia?`<p class="ayuda">Suele comprar cada ${x.frecuencia} días.</p>`:'');
  const hist=x.ventas.slice().reverse().slice(0,12),pagos=x.abonos.slice().reverse().slice(0,8);
  const telDD=x.telefono?`<span class="acc-fila" style="align-items:center">${esc(fTel(x.telefono))}<button type="button" class="btn chico solo-icono" data-copiar-tel aria-label="Copiar teléfono" title="Copiar">${ico('copiar')}</button></span>`:'Sin número';
  return `<div class="detalle">
    <div class="det-cab"><div><p class="ceja">${esc(CLI_TIPOS[x.tipo])}</p><h3 class="det-tit">${esc(x.nombre)}</h3>${x.saludo&&x.saludo!==x.nombre?`<p class="ayuda">Le saludas como «${esc(x.saludo)}»</p>`:''}</div>${estadoChip(x)}</div>
    <dl class="datos"><div><dt>WhatsApp</dt><dd>${telDD}</dd></div><div><dt>Zona</dt><dd>${esc(x.zona||'—')}</dd></div><div><dt>Suele pedir</dt><dd>${esc({blanco:'Huevo blanco',marron:'Huevo marrón'}[x.preferencia]||'Ambos')}</dd></div><div><dt>Precio especial</dt><dd>${precioEspTxt(x)}</dd></div>${x.notas?`<div class="ancho"><dt>Notas</dt><dd>${esc(x.notas)}</dd></div>`:''}</dl>
    <section class="sec-det"><h3>Recordatorio</h3>
      <p><b>Próximo:</b> ${x.proximo?esc(cap(fFechaLarga(x.proximo))+' ('+fRel(x.proximo)+')'):'Sin recordatorio'}</p>
      ${x.rec>0?`<div class="acc-fila">${botonWA(x)}${pe?`<button type="button" class="btn chico" data-act="contactado" data-id="${esc(x.id)}">${ico('check')}Ya le escribí</button>`:''}</div>`:''}
      ${pe&&pend?`<div class="acc-fila"><span class="ayuda" style="align-self:center">Posponer:</span>${[[1,'Mañana'],[3,'3 días'],[7,'1 semana']].map(([n,l])=>`<button type="button" class="opcion" data-posp="${n}">${l}</button>`).join('')}</div>`:''}
      <span class="etq">Escribirle cada</span>
      ${grupoOpc('rec',[...ops.map(n=>[n,n+' días']),[0,'No recordar']],ops.includes(x.rec)||x.rec===0?x.rec:'',dis)}
      ${pe?`<div class="acc-fila" style="align-items:flex-end"><div class="campo otro-rec"><label for="cd-rec-otro">Otro número de días</label><input id="cd-rec-otro" inputmode="numeric" placeholder="${!ops.includes(x.rec)&&x.rec>0?x.rec:'Ej. 21'}" autocomplete="off"></div><button type="button" class="btn chico" data-rec-otro>Usar</button></div>`:''}
      ${sug}
      ${x.rec>0?`<span class="etq">Mensaje de WhatsApp</span><div class="mensaje">${esc(mensajePara(x))}</div><button type="button" class="enlace" data-copiar-msg>Copiar mensaje</button>`:''}
    </section>
    <section class="sec-det"><h3>Resumen</h3>
      <dl class="datos"><div><dt>Compras</dt><dd>${num(x.ventas.length)}</dd></div><div><dt>Charolas</dt><dd>${num(x.totalCh)}</dd></div><div><dt>Total comprado</dt><dd>${dinero(x.totalGastado)}</dd></div><div><dt>Ganancia que deja</dt><dd class="${x.totalGanancia>=0?'pos':'neg'}">${dinero(x.totalGanancia)}</dd></div><div><dt>Saldo</dt><dd class="${x.saldo>0.009?'neg':''}">${x.saldo>0.009?'Debe '+dinero(x.saldo):x.saldo<-0.009?dinero(-x.saldo)+' a favor':'Al corriente'}</dd></div><div><dt>Cliente desde</dt><dd>${x.alta&&esFecha(x.alta)?esc(fDiaMes(x.alta)+' '+x.alta.slice(0,4)):'—'}</dd></div></dl>
      ${pe&&x.saldo>0.009?`<button type="button" class="btn chico" data-cobrar style="align-self:flex-start">${ico('moneda')}Registrar pago</button>`:''}
    </section>
    <section class="sec-det"><h3>Historial de compras</h3>${hist.length?`<div class="lista compacta">${hist.map(v=>filaHTML({act:'venta',id:v.id,ico:'ventas',tit:etiquetaCorta(v.fecha),sub:itemsHTML(v),monto:dinero(v.total),extra:v.pendiente>0.009?chip('rojo','','Debe '+dinero(v.pendiente),'mini'):''})).join('')}</div>`:'<p class="ayuda">Todavía no tiene compras.</p>'}</section>
    ${pagos.length?`<section class="sec-det"><h3>Pagos</h3><div class="lista compacta">${pagos.map(a=>filaHTML({act:'abono',id:a.id,ico:'moneda',tit:etiquetaCorta(a.fecha),sub:esc(METODOS[a.metodo]||'Efectivo'),monto:'+'+dinero(a.monto),montoCls:'pos'})).join('')}</div></section>`:''}
  </div>`;
}
function hojaDetalleCliente(id){
  if(!derivar().clientesPorId.get(id)){toast('Ese cliente ya no existe.');return;}
  const pie=puedeEditar()?`<button type="button" class="btn peligro solo-icono" data-borrar aria-label="Eliminar cliente" title="Eliminar cliente">${ico('basura')}</button><span class="flex1"></span><button type="button" class="btn" data-editar>${ico('lapiz')}Editar</button><button type="button" class="btn acento" data-nueva-venta>${ico('ventas')}Nueva venta</button>`:`<button type="button" class="btn" data-act="cerrar">Cerrar</button>`;
  abrirHoja({titulo:'Cliente',cuerpo:cuerpoCliente(id),pie,
    refrescar(){const b=document.querySelector('#capa .hoja-cuerpo');if(!b)return;const y=b.scrollTop,foco=document.activeElement&&document.activeElement.id,otro=b.querySelector('#cd-rec-otro'),val=otro?otro.value:'';b.innerHTML=cuerpoCliente(id);b.scrollTop=y;const nOtro=b.querySelector('#cd-rec-otro');if(nOtro&&val)nOtro.value=val;if(foco){const el=document.getElementById(foco);if(el)el.focus({preventScroll:true});}},
    alMontar(h){
      h.addEventListener('click',async e=>{const t=e.target;
        if(t.closest('[data-editar]')){const raw=D().clientes.get(id);if(raw)hojaCliente(raw);return;}
        if(t.closest('[data-borrar]')){const xb=derivar().clientesPorId.get(id);confirmarPie(h,`¿Eliminar este cliente?${xb&&xb.saldo>0.009?' Todavía debe '+dinero(xb.saldo)+'.':''} Sus ventas se quedan en el historial.`,'Sí, eliminar',async()=>{await borrarDoc('clientes',id);cerrarHoja();toast('Cliente eliminado.');});return;}
        if(t.closest('[data-nueva-venta]')){hojaVenta(null,{clienteId:id});return;}
        if(t.closest('[data-cobrar]')){hojaAbono(null,id);return;}
        const sg=t.closest('[data-rec]');if(sg){await cambiarRec(id,Number(sg.dataset.rec));return;}
        if(t.closest('[data-rec-otro]')){const v=entero(h.querySelector('#cd-rec-otro').value);if(v<1){toast('Escribe cada cuántos días.');return;}await cambiarRec(id,v);return;}
        const pp=t.closest('[data-posp]');if(pp){try{await posponer(id,sumarDias(hoy(),Number(pp.dataset.posp)));}catch(err){toast(msgError(err));}return;}
        if(t.closest('[data-copiar-msg]')){const x=derivar().clientesPorId.get(id);if(x)copiar(mensajePara(x),'Mensaje copiado. Pégalo en WhatsApp.');return;}
        if(t.closest('[data-copiar-tel]')){const x=derivar().clientesPorId.get(id);if(x)copiar(x.telefono||'','Teléfono copiado.');return;}
      });
      h.addEventListener('grupo',e=>{if(e.detail.grupo==='rec')cambiarRec(id,Number(e.detail.valor));});
    }});
}
function hojaPosponer(id){
  const x=derivar().clientesPorId.get(id);if(!x)return;
  abrirHoja({titulo:'Posponer recordatorio',
    cuerpo:`<p class="det-intro">¿Cuándo quieres que te recuerde escribirle a <b>${esc(x.nombre)}</b>?</p><div class="opciones grandes">${[[1,'Mañana'],[2,'En 2 días'],[3,'En 3 días'],[7,'En una semana']].map(([n,l])=>`<button type="button" class="opcion" data-posp="${n}">${l}<small>${esc(fFecha(sumarDias(hoy(),n)))}</small></button>`).join('')}</div><div class="campo" style="margin-top:16px"><label for="pp-fecha">Otra fecha</label><input type="date" id="pp-fecha" min="${sumarDias(hoy(),1)}"></div>`,
    pie:`<button type="button" class="btn" data-act="cerrar">Cancelar</button><button type="button" class="btn acento" data-posp-fecha>Posponer a esa fecha</button>`,
    alMontar(h){h.addEventListener('click',async e=>{const b=e.target.closest('[data-posp]');let f=null;if(b)f=sumarDias(hoy(),Number(b.dataset.posp));else if(e.target.closest('[data-posp-fecha]')){f=h.querySelector('#pp-fecha').value;if(!esFecha(f)){toast('Elige una fecha.');return;}}if(!f)return;try{await posponer(id,f);cerrarHoja();}catch(err){toast(msgError(err));}});}});
}
function hojaAjustes(){
  const c=cfg(),pe=puedeEditar(),dis=pe?'':'disabled';
  const filaP=(k,l)=>`<tr><th scope="row">${l}</th>${TIPOS.map(t=>`<td><div class="pref"><span>$</span><input id="aj-p-${k}-${t}" inputmode="decimal" value="${c.precios[k][t]!=null?esc(c.precios[k][t]):''}" placeholder="—" aria-label="${esc(l+', '+TIPO_NOMBRE[t])}" autocomplete="off" ${dis}></div></td>`).join('')}</tr>`;
  const cuerpo=`<div class="aj-todo"><form id="f-aj" class="form" novalidate>
    <section class="aj-sec"><h3>Negocio</h3><div class="campo"><label for="aj-negocio">Nombre del negocio</label><input id="aj-negocio" value="${esc(c.negocio)}" autocomplete="off" ${dis}></div><div class="campo"><label for="aj-prov">Granja o proveedor habitual</label><input id="aj-prov" value="${esc(c.proveedor)}" placeholder="Ej. Granja San Isidro" autocomplete="off" ${dis}></div></section>
    <section class="aj-sec"><h3>Lista de precios por charola</h3><p class="ayuda">Llenan el precio al registrar; siempre lo puedes cambiar. Si dejas uno vacío, se usa el último que registraste.</p><div class="tabla-env"><table class="tabla-precios"><thead><tr><th></th><th>Blanco</th><th>Marrón</th></tr></thead><tbody>${filaP('compra','Compra a granja')}${filaP('individual','Cliente individual')}${filaP('tienda','Tienda')}${filaP('distribuidor','Distribuidor')}</tbody></table></div></section>
    <section class="aj-sec"><h3>Recordatorios</h3><div class="campo otro-rec"><label for="aj-rec">Días para clientes nuevos</label><input id="aj-rec" inputmode="numeric" value="${c.recordarDias}" autocomplete="off" ${dis}></div><div class="campo"><label for="aj-plantilla">Mensaje de WhatsApp</label><textarea id="aj-plantilla" rows="4" ${dis}>${esc(c.plantilla)}</textarea><p class="ayuda">Puedes usar {nombre}, {pedido} (su último pedido), {dias} (días desde su última compra) y {negocio}.</p></div><div class="campo"><span class="etq">Así se verá</span><div class="mensaje" id="aj-vista"></div></div></section>
    <section class="aj-sec"><h3>Alertas de bodega</h3><div class="fila2"><div class="campo"><label for="aj-stock">Quedan pocas charolas</label><input id="aj-stock" inputmode="numeric" value="${c.stockMin}" autocomplete="off" ${dis}><p class="ayuda">Avisar con esta cantidad o menos.</p></div><div class="campo"><label for="aj-fresco">Frescura</label><input id="aj-fresco" inputmode="numeric" value="${c.frescuraDias}" autocomplete="off" ${dis}><p class="ayuda">Días en bodega para marcar un lote como viejo.</p></div></div></section>
    <p class="error" id="aj-error" role="alert" hidden></p></form>
    <section class="aj-sec"><h3>Respaldo</h3>
      <p class="ayuda">Tus datos se guardan solo en este celular. Haz un respaldo cada semana y guárdalo en Archivos o en Drive, o mándatelo por WhatsApp. Con ese archivo recuperas todo en otro celular.</p>
      <p><b>Último respaldo:</b> ${esc(textoRespaldo())}</p>
      <div class="acc-fila"><button type="button" class="btn primario" data-act="respaldar" ${S.modo==='real'?'':'disabled'}>${ico('respaldo')}Hacer respaldo</button><button type="button" class="btn" data-act="restaurar" ${S.modo==='real'?'':'disabled'}>${ico('subir')}Restaurar respaldo</button></div>
    </section>
    <section class="aj-sec"><h3>Datos</h3>
      <div class="acc-fila"><button type="button" class="btn" data-act="exportar">${ico('descargar')}Exportar todo a Excel</button></div>
      ${S.modo==='ejemplo'?`<div class="acc-fila"><button type="button" class="btn" data-act="salir-ejemplo">Salir del ejemplo</button></div>`:`<div class="acc-fila"><button type="button" class="btn" data-act="ejemplo">Explorar con datos de ejemplo</button></div><p class="ayuda">El ejemplo no toca tus datos reales.</p>`}
      ${S.modo==='real'&&pe?`<details class="mas"><summary>Borrar movimientos y clientes</summary><div class="form" style="gap:10px"><p class="ayuda">Borra todas las ventas, compras, mermas, pagos y clientes; los ajustes se conservan. No se puede deshacer, así que haz un respaldo antes.</p><div class="campo"><label for="aj-borrar">Escribe BORRAR para confirmar</label><input id="aj-borrar" autocomplete="off"></div><button type="button" class="btn peligro" id="aj-borrar-btn" style="align-self:flex-start">Borrar todo</button><p class="ayuda" id="aj-borrar-estado" aria-live="polite"></p></div></details>`:''}
    </section>
    <section class="aj-sec"><h3>Cómo se calcula</h3><p class="ayuda">La ganancia de cada venta usa el precio de compra de los lotes más antiguos en bodega (primero en entrar, primero en salir). Las piezas dañadas y la merma se restan del inventario y cuentan como pérdida. Los recordatorios cuentan desde la última compra del cliente, o desde el día en que le escribiste.</p><p class="ayuda">Versión ${VERSION_APP}</p></section>
  </div>`;
  abrirHoja({titulo:'Ajustes',cuerpo,formulario:true,
    pie:pe?`<button type="button" class="btn" data-act="cerrar">Cerrar</button><button type="submit" form="f-aj" class="btn acento" id="aj-guardar">Guardar ajustes</button>`:`<button type="button" class="btn" data-act="cerrar">Cerrar</button>`,
    alMontar(h){
      const $h=s=>h.querySelector(s);
      const vista=()=>{const d=derivar();const ej=d.clientes.find(x=>x.ult)||{nombre:'Doña Lupe',saludo:'Doña Lupe',ult:null};$h('#aj-vista').textContent=aplicarPlantilla($h('#aj-plantilla').value,ej);};
      $h('#aj-plantilla').addEventListener('input',vista);vista();
      $h('#f-aj').addEventListener('submit',async e=>{
        e.preventDefault();if(!pe)return;const err=$h('#aj-error'),btn=document.getElementById('aj-guardar');err.hidden=true;const fallo=x=>errorForm(err,x);
        const precios={};for(const k of Object.keys(CFG0.precios)){precios[k]={};for(const t of TIPOS){const v=decimal($h(`#aj-p-${k}-${t}`).value);if(v!=null&&(isNaN(v)||v<0))return fallo('Revisa los precios: solo números positivos.');precios[k][t]=v==null?null:r2(v);}}
        const rec=entero($h('#aj-rec').value);if(rec<1)return fallo('Los días para clientes nuevos deben ser 1 o más.');
        const fresco=entero($h('#aj-fresco').value);if(fresco<1)return fallo('Los días de frescura deben ser 1 o más.');
        const doc={negocio:$h('#aj-negocio').value.trim()||CFG0.negocio,proveedor:$h('#aj-prov').value.trim(),recordarDias:rec,plantilla:$h('#aj-plantilla').value.trim()||CFG0.plantilla,stockMin:entero($h('#aj-stock').value),frescuraDias:fresco,precios,editado:ahora()};
        if(btn)btn.disabled=true;
        try{await guardarConfig(doc);cerrarHoja();toast('Ajustes guardados.');}catch(e2){if(btn)btn.disabled=false;fallo(msgError(e2));}
      });
      const bb=$h('#aj-borrar-btn');
      if(bb)bb.addEventListener('click',async()=>{const est=$h('#aj-borrar-estado');if($h('#aj-borrar').value.trim().toUpperCase()!=='BORRAR'){est.textContent='Escribe BORRAR para confirmar.';return;}bb.disabled=true;try{await borrarTodo(t=>{est.textContent=t;});est.textContent='Listo: se borraron los movimientos y clientes.';toast('Datos borrados.');}catch(err){bb.disabled=false;est.textContent=msgError(err);}});
    }});
}
function avisoSoloLectura(){toast('Con este acceso solo se puede consultar.');}

/* ================= Almacenamiento en el celular (IndexedDB) ================= */
const BD={db:null,abriendo:null};
function abrirBD(){
  if(BD.db)return Promise.resolve(BD.db);
  if(BD.abriendo)return BD.abriendo;
  BD.abriendo=new Promise((res,rej)=>{
    let req;
    try{req=indexedDB.open(BD_NOMBRE,BD_VERSION);}catch(e){BD.abriendo=null;rej(e);return;}
    req.onupgradeneeded=()=>{const db=req.result;for(const s of ALMACENES)if(!db.objectStoreNames.contains(s))db.createObjectStore(s,{keyPath:'id'});};
    req.onsuccess=()=>{const db=req.result;db.onversionchange=()=>{try{db.close();}catch(e){}BD.db=null;};db.onclose=()=>{BD.db=null;};BD.db=db;BD.abriendo=null;res(db);};
    req.onerror=()=>{BD.abriendo=null;rej(req.error||new Error('indexeddb'));};
    req.onblocked=()=>{};
  });
  return BD.abriendo;
}
async function transaccion(almacenes,modo,fn){
  for(let intento=0;;intento++){
    try{
      const db=await abrirBD();
      return await new Promise((res,rej)=>{
        let t;try{t=db.transaction(almacenes,modo);}catch(e){rej(e);return;}
        const os={};for(const s of [].concat(almacenes))os[s]=t.objectStore(s);
        let valor;try{valor=fn(os);}catch(e){try{t.abort();}catch(_){}rej(e);return;}
        t.oncomplete=()=>res(valor);
        t.onerror=()=>rej(t.error||new Error('transaccion'));
        t.onabort=()=>rej(t.error||new Error('abortada'));
      });
    }catch(e){
      const n=e&&e.name;
      if(intento===0&&(n==='InvalidStateError'||n==='TransactionInactiveError'||n==='UnknownError')){try{if(BD.db)BD.db.close();}catch(_){}BD.db=null;continue;}
      throw e;
    }
  }
}
async function leerTodo(){
  const out={};
  await transaccion(ALMACENES,'readonly',os=>{for(const s of ALMACENES){const r=os[s].getAll();r.onsuccess=()=>{out[s]=r.result||[];};}});
  return out;
}
function aplicarLectura(out){
  const b=nuevaBase();
  for(const col of COLS)for(const d of (out[col]||[]))if(d&&d.id!=null)b[col].set(String(d.id),d);
  const cfgs=out.config||[];
  const gen=cfgs.find(x=>x&&x.id==='general');if(gen){const c={...gen};delete c.id;b.cfgDoc=c;}
  const meta=cfgs.find(x=>x&&x.id==='meta');S.meta=meta?{...meta}:{};
  S.real=b;S.ver++;
}
async function recargarDatos(){aplicarLectura(await leerTodo());programar();}
const canal=('BroadcastChannel' in window)?new BroadcastChannel('las-patronas'):null;
if(canal)canal.onmessage=()=>{if(S.almacenListo)recargarDatos().catch(()=>{});};
function avisarOtras(){try{if(canal)canal.postMessage('cambio');}catch(e){}}
const nuevoId=()=>(window.crypto&&crypto.randomUUID)?crypto.randomUUID():Date.now().toString(36)+Math.random().toString(36).slice(2,10);
function limpiar(data){const o=JSON.parse(JSON.stringify(data));delete o.id;return o;}
async function conEscritura(fn){
  S.escribiendo++;renderEstado();
  try{return await fn();}
  finally{S.escribiendo--;renderEstado();programar();}
}
async function guardarDoc(col,id,data){
  const body=limpiar(data);
  if(S.modo==='ejemplo'){const nid=id||uid();S.ejemplo[col].set(nid,{...body,id:nid});S.ver++;programar();return nid;}
  if(S.modo!=='real')throw {code:'sinalmacen'};
  const nid=id||nuevoId();const doc={...body,id:nid};
  await conEscritura(()=>transaccion(col,'readwrite',os=>{os[col].put(doc);}));
  S.real[col].set(nid,doc);S.ver++;programar();avisarOtras();pedirPersistencia();return nid;
}
async function actualizarDoc(col,id,patch){
  if(S.modo==='ejemplo'){const o=S.ejemplo[col].get(id);if(o){S.ejemplo[col].set(id,{...o,...patch});S.ver++;programar();}return;}
  if(S.modo!=='real')throw {code:'sinalmacen'};
  let nuevo=null;
  await conEscritura(()=>transaccion(col,'readwrite',os=>{const r=os[col].get(id);r.onsuccess=()=>{if(r.result){nuevo={...r.result,...patch,id};os[col].put(nuevo);}};}));
  if(nuevo){S.real[col].set(id,nuevo);S.ver++;programar();avisarOtras();}
}
async function borrarDoc(col,id){
  if(S.modo==='ejemplo'){S.ejemplo[col].delete(id);S.ver++;programar();return;}
  if(S.modo!=='real')throw {code:'sinalmacen'};
  await conEscritura(()=>transaccion(col,'readwrite',os=>{os[col].delete(id);}));
  S.real[col].delete(id);S.ver++;programar();avisarOtras();
}
async function guardarConfig(doc){
  if(S.modo==='ejemplo'){S.ejemplo.cfgDoc={...doc};S.ver++;programar();return;}
  if(S.modo!=='real')throw {code:'sinalmacen'};
  const body={...limpiar(doc),id:'general'};
  await conEscritura(()=>transaccion('config','readwrite',os=>{os.config.put(body);}));
  const c={...body};delete c.id;S.real.cfgDoc=c;S.ver++;programar();avisarOtras();
}
async function guardarMeta(patch){
  if(!S.almacenListo)return;
  const meta={...S.meta,...patch,id:'meta'};
  await transaccion('config','readwrite',os=>{os.config.put(meta);});
  S.meta=meta;programar();avisarOtras();
}
async function borrarTodo(prog){
  if(S.modo!=='real')throw {code:'sinalmacen'};
  prog('Borrando…');
  await conEscritura(()=>transaccion(COLS,'readwrite',os=>{for(const c of COLS)os[c].clear();}));
  for(const c of COLS)S.real[c]=new Map();S.ver++;programar();avisarOtras();
}
let _persistencia=false;
function pedirPersistencia(){if(_persistencia)return;_persistencia=true;try{if(navigator.storage&&navigator.storage.persist)navigator.storage.persist().catch(()=>{});}catch(e){}}
function msgError(e){
  const n=e&&(e.name||e.code);
  if(n==='QuotaExceededError')return 'Ya no hay espacio para guardar en este celular. Haz un respaldo y libera espacio.';
  if(n==='sinalmacen')return 'Este navegador no permite guardar datos.';
  return 'No se pudo guardar. Intenta de nuevo; si sigue fallando, cierra y vuelve a abrir la app.';
}
async function marcarContactado(id,btn){
  if(!puedeEditar())return;const x=derivar().clientesPorId.get(id);if(btn)btn.disabled=true;
  try{await actualizarDoc('clientes',id,{ultimoContacto:hoy(),posponerHasta:null});toast(x&&x.rec>0?`Anotado. Te recordaré escribirle a ${x.nombre} el ${fFecha(sumarDias(hoy(),x.rec))}.`:'Anotado.');}
  catch(e){if(btn)btn.disabled=false;toast(msgError(e));}
}
async function posponer(id,f){await actualizarDoc('clientes',id,{posponerHasta:f});const x=derivar().clientesPorId.get(id);toast(`Te recordaré escribirle a ${x?x.nombre:'este cliente'} el ${fFecha(f)}.`);}
async function cambiarRec(id,n){if(!puedeEditar())return;try{await actualizarDoc('clientes',id,{recordarCada:n});toast(n>0?`Listo: recordatorio cada ${n} días.`:'Listo: sin recordatorio para este cliente.');}catch(e){toast(msgError(e));}}
function despuesDeWA(id){if(!puedeEditar())return;const x=derivar().clientesPorId.get(id);if(!x)return;setTimeout(()=>toast(`¿Ya le escribiste a ${x.nombre}?`,{texto:'Sí, anotar',fn:()=>marcarContactado(id)}),700);}
function copiar(texto,ok){
  const respaldo=()=>abrirHoja({titulo:'Copiar',cuerpo:`<p class="det-intro">Selecciona el texto y cópialo:</p><textarea id="copiar-txt" rows="5" readonly>${esc(texto)}</textarea>`,pie:'<button type="button" class="btn" data-act="cerrar">Listo</button>',alMontar(h){const t=h.querySelector('#copiar-txt');setTimeout(()=>{t.focus();t.select();},60);}});
  try{const p=navigator.clipboard&&navigator.clipboard.writeText(texto);if(p&&p.then)p.then(()=>toast(ok||'Copiado.'),respaldo);else respaldo();}catch(e){respaldo();}
}
let toastT=null;
function toast(msg,accion){
  const t=$('#toast');t.innerHTML=`<span>${esc(msg)}</span>${accion?`<button type="button">${esc(accion.texto)}</button>`:''}`;t.hidden=false;
  if(accion)t.querySelector('button').addEventListener('click',()=>{t.hidden=true;accion.fn();});
  clearTimeout(toastT);toastT=setTimeout(()=>{t.hidden=true;},accion?8000:4200);
}

/* ================= Archivos: respaldo, restauración y Excel ================= */
async function entregarArchivo(nombre,blob){
  let file=null;try{file=new File([blob],nombre,{type:blob.type});}catch(e){}
  if(file&&navigator.canShare&&navigator.share){
    let ok=false;try{ok=navigator.canShare({files:[file]});}catch(e){}
    if(ok){try{await navigator.share({files:[file],title:nombre});return 'compartido';}catch(e){if(e&&e.name==='AbortError')return 'cancelado';}}
  }
  const url=URL.createObjectURL(blob);const a=document.createElement('a');a.href=url;a.download=nombre;a.rel='noopener';document.body.appendChild(a);a.click();a.remove();setTimeout(()=>URL.revokeObjectURL(url),60000);
  return 'descargado';
}
function textoRespaldo(){const f=S.meta&&S.meta.ultimoRespaldo;if(!esFecha(f))return 'nunca';const k=diffDias(hoy(),f);return k<=0?'hoy':k===1?'ayer':`hace ${k} días (${fDiaMes(f)})`;}
function necesitaRespaldo(){
  if(S.modo!=='real'||derivar().vacio)return false;
  const H=hoy();if(esFecha(S.meta.posponerRespaldo)&&S.meta.posponerRespaldo>H)return false;
  const base=esFecha(S.meta.ultimoRespaldo)?S.meta.ultimoRespaldo:(esFecha(S.meta.inicio)?S.meta.inicio:H);
  return diffDias(H,base)>=DIAS_RESPALDO;
}
function datosRespaldo(){
  const o={app:'las-patronas',schemaVersion:1,exportedAt:ahora(),appVersion:VERSION_APP,config:S.real.cfgDoc||null};
  for(const c of COLS)o[c]=[...S.real[c].values()];
  return o;
}
async function respaldar(btn){
  if(S.modo!=='real'){toast('Sal del ejemplo para respaldar tus datos.');return;}
  const blob=new Blob([JSON.stringify(datosRespaldo())],{type:'application/json'});
  if(btn)btn.disabled=true;
  try{
    const r=await entregarArchivo(`las-patronas-respaldo-${hoy()}.json`,blob);
    if(r==='cancelado'){toast('Respaldo cancelado.');return;}
    await guardarMeta({ultimoRespaldo:hoy(),posponerRespaldo:null});
    toast(r==='compartido'?'Respaldo listo. Guárdalo en Archivos o mándatelo por WhatsApp.':'Respaldo descargado.');
  }catch(e){toast('No se pudo crear el respaldo. Intenta de nuevo.');}
  finally{if(btn)btn.disabled=false;}
}
function elegirRespaldo(){if(S.modo!=='real'){toast('Sal del ejemplo para restaurar un respaldo.');return;}const inp=$('#archivo-respaldo');inp.value='';inp.click();}
function validarRespaldo(o){
  if(!o||typeof o!=='object'||o.app!=='las-patronas')throw new Error('Ese archivo no es un respaldo de Las Patronas.');
  for(const c of COLS)if(o[c]!=null&&!Array.isArray(o[c]))throw new Error('El respaldo está dañado.');
  return o;
}
async function leerArchivoRespaldo(file){
  let o;
  try{o=JSON.parse(await file.text());}catch(e){toast('No se pudo leer el archivo. Elige el .json del respaldo.');return;}
  try{validarRespaldo(o);}catch(e){toast(e.message);return;}
  const cuenta=c=>Array.isArray(o[c])?o[c].length:0;
  const f=typeof o.exportedAt==='string'?o.exportedAt.slice(0,10):'';
  const fecha=esFecha(f)?fFechaLarga(f)+' de '+f.slice(0,4):'fecha desconocida';
  abrirHoja({titulo:'Restaurar respaldo',formulario:true,
    cuerpo:`<div class="detalle"><p class="det-intro">Respaldo del <b>${esc(fecha)}</b>:</p><dl class="datos"><div><dt>Clientes</dt><dd>${num(cuenta('clientes'))}</dd></div><div><dt>Ventas</dt><dd>${num(cuenta('ventas'))}</dd></div><div><dt>Compras</dt><dd>${num(cuenta('compras'))}</dd></div><div><dt>Mermas</dt><dd>${num(cuenta('mermas'))}</dd></div><div><dt>Pagos</dt><dd>${num(cuenta('abonos'))}</dd></div></dl><p class="nota-alerta">${ico('alerta')}<span>Esto reemplaza todo lo que hay ahora en este celular con el contenido del respaldo.</span></p></div>`,
    pie:`<button type="button" class="btn" data-act="cerrar">Cancelar</button><button type="button" class="btn acento" id="rs-ok">Restaurar</button>`,
    alMontar(h){const b=h.querySelector('#rs-ok');b.addEventListener('click',async()=>{b.disabled=true;try{await restaurar(o);cerrarHoja();toast('Respaldo restaurado.');}catch(e){b.disabled=false;toast(msgError(e));}});}});
}
async function restaurar(o){
  const docs={};for(const c of COLS)docs[c]=(o[c]||[]).filter(d=>d&&typeof d==='object'&&d.id!=null).map(d=>({...d,id:String(d.id)}));
  const cfgGen=o.config&&typeof o.config==='object'?{...o.config,id:'general'}:null;
  const f=typeof o.exportedAt==='string'?o.exportedAt.slice(0,10):'';
  const meta={...S.meta,id:'meta',ultimoRespaldo:esFecha(f)?f:hoy(),posponerRespaldo:null};
  await conEscritura(()=>transaccion(ALMACENES,'readwrite',os=>{
    for(const c of COLS){os[c].clear();for(const d of docs[c])os[c].put(d);}
    os.config.clear();if(cfgGen)os.config.put(cfgGen);os.config.put(meta);
  }));
  await recargarDatos();avisarOtras();
}
let _xlsx=null;
function cargarXLSX(){
  if(window.XLSX)return Promise.resolve(window.XLSX);
  if(_xlsx)return _xlsx;
  _xlsx=new Promise((res,rej)=>{const sc=document.createElement('script');sc.src=XLSX_URL+'?v='+VERSION_APP;sc.async=true;sc.onload=()=>window.XLSX?res(window.XLSX):rej(new Error('xlsx'));sc.onerror=()=>{_xlsx=null;rej(new Error('xlsx'));};document.head.appendChild(sc);});
  return _xlsx;
}
function libroExcel(X,d){
  const wb=X.utils.book_new();
  const F=s=>{const [y,m,dd]=s.split('-').map(Number);return new Date(y,m-1,dd);};
  const hoja=(nombre,enc,filas,anchos)=>{const ws=X.utils.aoa_to_sheet([enc,...filas],{cellDates:true,dateNF:'dd/mm/yyyy'});ws['!cols']=anchos.map(w=>({wch:w}));X.utils.book_append_sheet(wb,ws,nombre);};
  const it=(o,t)=>o.items.find(i=>i.tipo===t);
  hoja('Ventas',['Fecha','Cliente','Tipo de cliente','Charolas blanco','Precio blanco','Charolas marrón','Precio marrón','Total','Pagado al momento','Pendiente','Costo (PEPS)','Ganancia','Método','Notas'],d.ventas.slice().reverse().map(v=>{const b=it(v,'blanco'),m=it(v,'marron');return [F(v.fecha),v.nombre,v.cli?CLI_TIPOS[v.cli.tipo]:'Mostrador',b?b.charolas:0,b?b.precio:null,m?m.charolas:0,m?m.precio:null,v.total,v.pagadoMomento,v.pendiente,v.costo,v.ganancia,METODOS[v.metodo]||'',v.notas||''];}),[12,30,15,10,10,10,10,11,12,11,12,11,14,30]);
  hoja('Compras',['Fecha','Proveedor','Charolas blanco','Precio blanco','Dañadas blanco (pzas)','Charolas marrón','Precio marrón','Dañadas marrón (pzas)','Total','Notas'],d.compras.slice().reverse().map(c=>{const b=it(c,'blanco'),m=it(c,'marron');return [F(c.fecha),c.proveedor||'',b?b.charolas:0,b?b.precio:null,b?b.danadas:0,m?m.charolas:0,m?m.precio:null,m?m.danadas:0,c.total,c.notas||''];}),[12,28,10,10,12,10,10,12,12,30]);
  hoja('Merma',['Fecha','Tipo','Piezas','Charolas equivalentes','Motivo','Costo','Notas'],d.mermas.slice().reverse().map(m=>[F(m.fecha),TIPO_NOMBRE[m.tipo]||m.tipo,m.piezas,r2(m.piezas/PZ),m.motivo||'',m.costo,m.notas||'']),[12,14,9,12,24,10,30]);
  hoja('Pagos',['Fecha','Cliente','Monto','Método','Nota'],d.abonos.slice().reverse().map(a=>[F(a.fecha),a.nombre,a.monto,METODOS[a.metodo]||'',a.nota||'']),[12,30,11,14,30]);
  hoja('Clientes',['Nombre','Saludo','Tipo','WhatsApp','Zona','Recordar cada (días)','Última compra','Próximo contacto','Compras','Charolas','Total comprado','Saldo','Notas'],d.clientes.slice().sort((a,b)=>a.nombre.localeCompare(b.nombre,'es')).map(x=>[x.nombre,x.saludo||'',CLI_TIPOS[x.tipo],x.telefono||'',x.zona||'',x.rec,x.ult?F(x.ult.fecha):null,x.proximo?F(x.proximo):null,x.ventas.length,x.totalCh,x.totalGastado,x.saldo,x.notas||'']),[30,18,13,14,20,12,13,14,9,9,13,11,30]);
  const filasInv=[];TIPOS.forEach(t=>d.inv[t].lotes.forEach(L=>filasInv.push([TIPO_NOMBRE[t],F(L.fecha),L.edad,Math.floor(L.restante/PZ),L.restante%PZ,r2(L.cpz*PZ),r2(L.restante*L.cpz)])));
  hoja('Inventario',['Tipo','Compra del','Días en bodega','Charolas','Piezas sueltas','Costo por charola','Valor a costo'],filasInv,[14,12,13,10,12,14,13]);
  return X.write(wb,{bookType:'xlsx',type:'array'});
}
function hojaExportar(){
  const d=derivar();
  abrirHoja({titulo:'Exportar a Excel',
    cuerpo:`<div class="detalle"><p class="det-intro">Un archivo con seis hojas: ventas, compras, merma, pagos, clientes e inventario por lote.</p><dl class="datos"><div><dt>Ventas</dt><dd>${num(d.ventas.length)}</dd></div><div><dt>Compras</dt><dd>${num(d.compras.length)}</dd></div><div><dt>Clientes</dt><dd>${num(d.clientes.length)}</dd></div><div><dt>Pagos</dt><dd>${num(d.abonos.length)}</dd></div></dl><p class="ayuda" id="ex-estado">Preparando…</p></div>`,
    pie:`<button type="button" class="btn" data-act="cerrar">Cerrar</button><button type="button" class="btn acento" id="ex-ok" disabled>${ico('compartir')}Guardar o enviar</button>`,
    alMontar(h){
      const b=h.querySelector('#ex-ok'),est=h.querySelector('#ex-estado');
      cargarXLSX().then(()=>{b.disabled=false;est.textContent='Listo. En iPhone puedes guardarlo en Archivos o mandarlo por WhatsApp.';}).catch(()=>{est.textContent='No se pudo preparar el Excel. Cierra y vuelve a abrir la app.';});
      b.addEventListener('click',async()=>{
        if(!window.XLSX)return;b.disabled=true;
        try{
          const buf=libroExcel(window.XLSX,derivar());
          const r=await entregarArchivo(`las-patronas-${hoy()}.xlsx`,new Blob([buf],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'}));
          if(r!=='cancelado'){cerrarHoja();toast(r==='compartido'?'Archivo de Excel listo.':'Se descargó el archivo de Excel.');}
        }catch(e){toast('No se pudo generar el archivo. Intenta de nuevo.');}
        finally{b.disabled=false;}
      });
    }});
}

/* ================= Datos de ejemplo (solo en memoria) ================= */
function generarEjemplo(){
  let s=0x2F6E2B1;const rnd=()=>{s=(s+0x6D2B79F5)|0;let t=Math.imul(s^(s>>>15),1|s);t=(t+Math.imul(t^(t>>>7),61|t))^t;return((t^(t>>>14))>>>0)/4294967296;};
  const ri=(a,b)=>a+Math.floor(rnd()*(b-a+1));
  const H=hoy(),INI=-48;
  const db={...nuevaBase(),cfgDoc:{negocio:CFG0.negocio,proveedor:'Granja Avícola San Isidro',recordarDias:7,plantilla:CFG0.plantilla,stockMin:10,frescuraDias:15,precios:{compra:{blanco:null,marron:null},individual:{blanco:95,marron:110},tienda:{blanco:86,marron:99},distribuidor:{blanco:78,marron:90}}}};
  const defs=[
    ['Doña Lupe Hernández','Doña Lupe','individual',7,7,-7,[1,2],.5,'0000000101','Col. Las Flores'],
    ['Abarrotes La Esperanza','Don Ramiro','tienda',5,4,-1,[5,8],.3,'0000000102','Av. Principal'],
    ['Cocina Económica Mary','Mary','tienda',4,3,-2,[4,6],.15,'0000000103','Centro'],
    ['Distribuidora Hnos. Rangel','Sr. Rangel','distribuidor',7,7,-3,[18,24],.35,'0000000104','Ruta norte'],
    ['Carmen Ortiz','Carmen','individual',10,10,-12,[1,2],.8,'0000000105','Fracc. Los Pinos'],
    ['Tienda Don Beto','Don Beto','tienda',6,5,-9,[4,7],.25,'0000000106','Barrio de San Juan'],
    ['Panadería San Judas','','tienda',3,3,-1,[6,9],0,'0000000107','Col. Jardines'],
    ['Laura Méndez','Laura','individual',14,14,-14,[2,3],1,'','Fracc. Los Pinos']
  ];
  const alta=sumarDias(H,INI-10);
  const cids=defs.map((x,i)=>{const id='ej-cli-'+i;db.clientes.set(id,{id,nombre:x[0],saludo:x[1],tipo:x[2],recordarCada:x[3],telefono:x[8],zona:x[9],preferencia:x[7]>=1?'marron':x[7]===0?'blanco':'ambos',notas:i===1?'Paga los viernes.':'',alta,creado:alta+'T09:00:00.000Z'});return id;});
  const plan=new Map();
  defs.forEach((x,i)=>{let k=x[5];while(k>=INI+2){if(!plan.has(k))plan.set(k,[]);plan.get(k).push(i);k-=x[4]+(x[4]>=5?ri(-1,1):0);}});
  const stock={blanco:0,marron:0},fiados=[];let pc={blanco:66,marron:77},nV=0,nC=0,nM=0,nA=0;
  for(let dia=INI;dia<=0;dia++){
    const f=sumarDias(H,dia);
    if((dia-INI)%4===0){if(dia>=-20)pc={blanco:68,marron:79};const chB=ri(31,36),chM=ri(11,14),dB=ri(0,8),dM=ri(0,5);stock.blanco+=chB*PZ-dB;stock.marron+=chM*PZ-dM;const id='ej-com-'+(nC++);db.compras.set(id,{id,fecha:f,proveedor:'Granja Avícola San Isidro',items:[{tipo:'blanco',charolas:chB,precio:pc.blanco,danadas:dB},{tipo:'marron',charolas:chM,precio:pc.marron,danadas:dM}],notas:'',creado:f+'T08:00:00.000Z'});}
    const lista=(plan.get(dia)||[]).slice();if(rnd()<.35)lista.push(-1);
    lista.forEach((ci,j)=>{
      let qB,qM,tipo;
      if(ci===-1){tipo='individual';qB=rnd()<.7?1:0;qM=qB?0:1;}
      else{const x=defs[ci];tipo=x[2];const q=ri(x[6][0],x[6][1]);qM=x[7]>=1?q:Math.min(q,Math.round(q*x[7]*(.6+rnd()*.8)));qB=q-qM;}
      qB=Math.min(qB,Math.floor(stock.blanco/PZ));qM=Math.min(qM,Math.floor(stock.marron/PZ));if(qB+qM<=0)return;
      stock.blanco-=qB*PZ;stock.marron-=qM*PZ;
      const pr=db.cfgDoc.precios[tipo],items=[];if(qB)items.push({tipo:'blanco',charolas:qB,precio:pr.blanco});if(qM)items.push({tipo:'marron',charolas:qM,precio:pr.marron});
      const total=sum(items,i=>i.charolas*i.precio);const pagado=(ci>=0&&tipo==='tienda'&&rnd()<.3)?0:total;
      const id='ej-ven-'+(nV++);const v={id,fecha:f,clienteId:ci>=0?cids[ci]:null,clienteNombre:ci>=0?defs[ci][0]:'',items,pagado,metodo:tipo==='distribuidor'||rnd()<.2?'transferencia':'efectivo',notas:'',creado:f+'T'+pad(10+j)+':20:00.000Z'};
      db.ventas.set(id,v);if(pagado<total)fiados.push({v,total});
    });
    if([-41,-29,-16,-6].includes(dia)){const t=(dia===-29||dia===-6)?'marron':'blanco';const pz=Math.min(ri(6,16),stock[t]);if(pz>0){stock[t]-=pz;const id='ej-mer-'+(nM++);db.mermas.set(id,{id,fecha:f,tipo:t,piezas:pz,motivo:MOTIVOS[[0,2,1,4][nM%4]],notas:'',creado:f+'T18:00:00.000Z'});}}
  }
  fiados.forEach(({v,total})=>{const k=diffDias(v.fecha,H);if(-k<7)return;const dp=Math.min(0,k+ri(3,6));const fp=sumarDias(H,dp);const id='ej-abo-'+(nA++);db.abonos.set(id,{id,clienteId:v.clienteId,clienteNombre:v.clienteNombre,monto:total,fecha:fp,metodo:'efectivo',nota:'',creado:fp+'T17:00:00.000Z'});});
  return db;
}
function entrarEjemplo(){if(!S.ejemplo)S.ejemplo=generarEjemplo();S.modo='ejemplo';S.ver++;cerrarHoja();window.scrollTo(0,0);render();}
function salirEjemplo(){S.modo=S.almacenListo?'real':'sinalmacen';S.ver++;cerrarHoja();window.scrollTo(0,0);render();}

/* ================= Render ================= */
function fabTipo(){return {bodega:'compra',clientes:'cliente'}[S.tab]||'venta';}
function estadoTxt(){
  if(S.modo==='ejemplo')return {cls:'ocre',txt:'Ejemplo'};
  if(S.modo==='cargando')return {cls:'',txt:'Abriendo…'};
  if(S.modo==='sinalmacen')return {cls:'rojo',txt:'Sin guardar'};
  if(S.escribiendo>0)return {cls:'',txt:'Guardando…'};
  return {cls:'olivo',txt:'Guardado'};
}
function renderEstado(){const e=estadoTxt();$$('.estado').forEach(el=>{el.className='estado '+e.cls;el.innerHTML='<i></i>'+esc(e.txt);});}
function esInstalada(){try{return window.matchMedia('(display-mode: standalone)').matches||navigator.standalone===true;}catch(e){return false;}}
function esIOS(){const ua=navigator.userAgent||'';return /iPad|iPhone|iPod/.test(ua)||(navigator.platform==='MacIntel'&&navigator.maxTouchPoints>1);}
function avisos(){
  let h='';
  if(S.swEspera)h+=`<div class="aviso">${ico('descargar')}<span>Hay una versión nueva de la app.</span><button type="button" class="btn chico primario" data-act="actualizar-app">Actualizar</button></div>`;
  if(S.modo==='ejemplo')h+=`<div class="aviso">${ico('alerta')}<span>Estás viendo <b>datos de ejemplo</b>. Puedes probar todo; nada se guarda.</span><button type="button" class="btn chico" data-act="salir-ejemplo">Salir del ejemplo</button></div>`;
  if(necesitaRespaldo()){const f=S.meta.ultimoRespaldo;h+=`<div class="aviso">${ico('respaldo')}<span>Tus datos solo están en este celular. ${esFecha(f)?'Tu último respaldo fue '+esc(textoRespaldo())+'.':'Aún no has hecho un respaldo.'}</span><button type="button" class="btn chico primario" data-act="respaldar">Respaldar</button><button type="button" class="btn chico" data-act="posponer-respaldo">Mañana</button></div>`;}
  if(S.modo==='real'&&!esInstalada()&&!leerLS('lp-instalar-oculto')){
    const txt=S.instalar?'Instálala para abrirla desde su ícono y usarla sin señal.':esIOS()?'Para instalarla, en Safari toca Compartir y luego «Agregar a pantalla de inicio».':'Para instalarla, abre el menú del navegador y elige «Instalar app» o «Agregar a la pantalla principal».';
    h+=`<div class="aviso">${ico('celular')}<span>${txt}</span>${S.instalar?'<button type="button" class="btn chico primario" data-act="instalar">Instalar</button>':''}<button type="button" class="btn chico" data-act="ocultar-instalar">Ocultar</button></div>`;
  }
  return h;
}
let _avisos='';
function renderShell(){
  $$('#nav-riel [data-tab],#nav-tabs [data-tab]').forEach(b=>{if(b.dataset.tab===S.tab)b.setAttribute('aria-current','page');else b.removeAttribute('aria-current');});
  renderEstado();
  const fab=$('#fab');
  let ver=(S.modo==='real'||S.modo==='ejemplo')&&puedeEditar()&&S.tab!=='reportes';
  if(ver&&S.modo==='real'&&S.tab==='inicio'&&derivar().vacio)ver=false;
  fab.hidden=!ver;
  if(ver){const t=fabTipo();if(fab.dataset.tipo!==t){fab.dataset.tipo=t;fab.innerHTML=`${ico('mas')}<span>${FAB_TXT[t]}</span>`;fab.setAttribute('aria-label','Registrar '+FAB_TXT[t].toLowerCase());}}
  const av=avisos();if(av!==_avisos){_avisos=av;$('#avisos').innerHTML=av;}
}
let _pend=false;
function programar(){if(_pend)return;_pend=true;requestAnimationFrame(()=>{_pend=false;render();});}
let ro=null;
function render(){
  const v=$('#vista');const act=document.activeElement;let fid=null,s0=null,s1=null;
  if(act&&act.id&&v.contains(act)){fid=act.id;try{s0=act.selectionStart;s1=act.selectionEnd;}catch(e){}}
  renderShell();
  let html;
  if(S.modo==='cargando')html=vCargando();
  else if(S.modo==='sinalmacen')html=vSinAlmacen();
  else{const d=derivar();html=({inicio:vInicio,ventas:vVentas,bodega:vBodega,clientes:vClientes,reportes:vReportes}[S.tab]||vInicio)(d);}
  v.innerHTML=html;
  if(fid){const el=document.getElementById(fid);if(el){try{el.focus({preventScroll:true});if(s0!=null&&el.setSelectionRange)el.setSelectionRange(s0,s1);}catch(e){}}}
  const g=document.getElementById('graf');
  if(ro)ro.disconnect();
  if(g){if(!ro&&'ResizeObserver' in window)ro=new ResizeObserver(debounce(dibujarGrafica,60));if(ro)ro.observe(g);dibujarGrafica();}
  if(S.hoja&&S.hoja.refrescar){try{S.hoja.refrescar();}catch(e){}}
}
function irA(tab){if(!TABS.includes(tab))return;S.tab=tab;guardarLS('lp-tab',tab);try{history.replaceState(null,'','#'+tab);}catch(e){}cerrarHoja();window.scrollTo(0,0);render();}
function setFiltro(k,v){if(!(k in S.f))return;S.f[k]=k==='rTabla'?v==='1':v;render();}
function nueva(tipo,ds){
  ds=ds||{};if(!puedeEditar()){avisoSoloLectura();return;}
  if(tipo==='venta')hojaVenta(null,{clienteId:ds.cliente||null});
  else if(tipo==='compra')hojaCompra(null);
  else if(tipo==='merma')hojaMerma(null);
  else if(tipo==='abono')hojaAbono(null,ds.cliente||null);
  else if(tipo==='cliente')hojaCliente(null);
}

/* ================= Eventos ================= */
document.addEventListener('click',e=>{
  const el=e.target.closest('[data-act]');if(!el)return;const a=el.dataset.act,id=el.dataset.id;
  switch(a){
    case 'tab':irA(el.dataset.tab);break;
    case 'nueva':nueva(el.dataset.tipo,el.dataset);break;
    case 'fab':nueva(el.dataset.tipo||fabTipo());break;
    case 'venta':hojaDetalleVenta(id);break;
    case 'compra':hojaDetalleCompra(id);break;
    case 'merma':hojaDetalleMerma(id);break;
    case 'abono':hojaDetalleAbono(id);break;
    case 'cliente':hojaDetalleCliente(id);break;
    case 'contactado':marcarContactado(id,el);break;
    case 'posponer':hojaPosponer(id);break;
    case 'wa':despuesDeWA(id);break;
    case 'wa-ejemplo':toast('En el ejemplo no se abre WhatsApp. Con tus clientes reales se abre el chat con este mensaje listo.');break;
    case 'copiar-msg':{const x=derivar().clientesPorId.get(id);if(x)copiar(mensajePara(x),'Mensaje copiado. Pégalo en WhatsApp.');break;}
    case 'ajustes':hojaAjustes();break;
    case 'ejemplo':entrarEjemplo();break;
    case 'salir-ejemplo':salirEjemplo();break;
    case 'cerrar':cerrarHoja();break;
    case 'filtro':setFiltro(el.dataset.k,el.dataset.v);break;
    case 'cobranza':S.f.vVista='cobrar';irA('ventas');break;
    case 'exportar':hojaExportar();break;
    case 'reintentar':reintentar();break;
    case 'respaldar':respaldar(el);break;
    case 'restaurar':elegirRespaldo();break;
    case 'posponer-respaldo':guardarMeta({posponerRespaldo:sumarDias(hoy(),1)}).catch(()=>{});break;
    case 'actualizar-app':actualizarApp();break;
    case 'instalar':instalarApp();break;
    case 'ocultar-instalar':guardarLS('lp-instalar-oculto','1');programar();break;
  }
});
document.addEventListener('keydown',e=>{if(e.key==='Escape'&&S.hoja){e.preventDefault();cerrarHoja();}});
document.addEventListener('input',e=>{const t=e.target;if(t.id==='q-ventas'){S.f.vBuscar=t.value;programar();}else if(t.id==='q-clientes'){S.f.cBuscar=t.value;programar();}});
document.addEventListener('change',e=>{const t=e.target;if(t.id==='c-orden'){S.f.cOrden=t.value;programar();}else if(t.id==='r-desde'){S.f.rDesde=t.value;programar();}else if(t.id==='r-hasta'){S.f.rHasta=t.value;programar();}});
setInterval(()=>{if(_dc&&_dc.H!==hoy())programar();},60000);

/* ================= Arranque ================= */
function montarShell(){
  const nav=TABS.map(t=>`<button type="button" data-act="tab" data-tab="${t}">${ico(t)}<span>${TAB_NOMBRE[t]}</span></button>`).join('');
  $('#nav-riel').innerHTML=nav;$('#nav-tabs').innerHTML=nav;
  $('#btn-ajustes').innerHTML=ico('ajustes');$('#btn-ajustes-riel').innerHTML=ico('ajustes')+'Ajustes';
}
function registrarSW(){
  if(!('serviceWorker' in navigator))return;
  const h=location.hostname;
  if(location.protocol!=='https:'&&h!=='localhost'&&h!=='127.0.0.1')return;
  const habiaControlador=!!navigator.serviceWorker.controller;
  let recargando=false;
  navigator.serviceWorker.addEventListener('controllerchange',()=>{if(!habiaControlador||recargando)return;recargando=true;location.reload();});
  navigator.serviceWorker.register('sw.js',{updateViaCache:'none'}).then(reg=>{
    S.sw=reg;
    const marcar=w=>{if(w&&navigator.serviceWorker.controller){S.swEspera=w;programar();}};
    if(reg.waiting)marcar(reg.waiting);
    reg.addEventListener('updatefound',()=>{const w=reg.installing;if(!w)return;w.addEventListener('statechange',()=>{if(w.state==='installed')marcar(w);});});
    let ultima=Date.now();
    document.addEventListener('visibilitychange',()=>{if(document.visibilityState==='visible'&&Date.now()-ultima>36e5){ultima=Date.now();reg.update().catch(()=>{});}});
  }).catch(()=>{});
}
function actualizarApp(){const w=S.swEspera;if(!w){location.reload();return;}w.postMessage('SKIP_WAITING');}
window.addEventListener('beforeinstallprompt',e=>{e.preventDefault();S.instalar=e;programar();});
window.addEventListener('appinstalled',()=>{S.instalar=null;programar();});
async function instalarApp(){const e=S.instalar;if(!e)return;S.instalar=null;try{await e.prompt();await e.userChoice;}catch(x){}programar();}
async function cargarAlmacen(){
  aplicarLectura(await leerTodo());
  S.almacenListo=true;
  if(S.modo!=='ejemplo')S.modo='real';
  if(!esFecha(S.meta.inicio))guardarMeta({inicio:hoy()}).catch(()=>{});
}
function reintentar(){if(S.modo!=='ejemplo')S.modo='cargando';render();cargarAlmacen().catch(()=>{if(S.modo!=='ejemplo')S.modo='sinalmacen';}).then(render);}
async function iniciar(){
  montarShell();
  const h=(location.hash||'').slice(1);
  if(TABS.includes(h))S.tab=h;else{const t=leerLS('lp-tab');if(TABS.includes(t))S.tab=t;}
  render();
  registrarSW();
  $('#archivo-respaldo').addEventListener('change',e=>{const f=e.target.files&&e.target.files[0];if(f)leerArchivoRespaldo(f);});
  if(!('indexedDB' in window)){S.modo='sinalmacen';render();return;}
  try{await cargarAlmacen();}catch(e){if(S.modo!=='ejemplo')S.modo='sinalmacen';}
  render();
}
iniciar();
})();
