import { useState, useEffect, useRef } from "react";
import * as XLSX from "xlsx";

// ─── SUPABASE CONFIG ───────────────────────────────────────────────────────
const SUPABASE_URL = "https://lrfbupelsjvphbmokfxd.supabase.co";
const SUPABASE_KEY = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImxyZmJ1cGVsc2p2cGhibW9rZnhkIiwicm9sZSI6ImFub24iLCJpYXQiOjE3ODAxMzA2OTgsImV4cCI6MjA5NTcwNjY5OH0.WnkFWFioGGaz33T3pRiuz0eGI62k9Omdo_uKNVZiJWU";

async function sbFetch(path, options = {}) {
  const res = await fetch(`${SUPABASE_URL}/rest/v1/${path}`, {
    headers: { "apikey": SUPABASE_KEY, "Authorization": `Bearer ${SUPABASE_KEY}`, "Content-Type": "application/json", "Prefer": "return=representation", ...options.headers },
    ...options,
  });
  if (!res.ok) { const err = await res.json().catch(()=>({})); throw new Error(err.message||`Error ${res.status}`); }
  const text = await res.text(); return text ? JSON.parse(text) : [];
}

const db = {
  getUsers:        (cid)    => sbFetch(`users?select=*&order=id.asc${cid?`&company_id=eq.${cid}`:""}`),
  getRecords:      (cid)    => sbFetch(`records?select=*&order=entry.desc${cid?`&company_id=eq.${cid}`:""}`),
  getCompanies:    ()       => sbFetch("companies?select=*&order=id.asc"),
  login:           (doc)    => sbFetch(`users?documento=eq.${encodeURIComponent(doc)}&select=*`),
  insertRecord:    (body)   => sbFetch("records", { method:"POST", body:JSON.stringify(body) }),
  updateRecord:    (id,b)   => sbFetch(`records?id=eq.${id}`, { method:"PATCH", body:JSON.stringify(b) }),
  insertUser:      (body)   => sbFetch("users", { method:"POST", body:JSON.stringify(body) }),
  deleteUser:      (id)     => sbFetch(`users?id=eq.${id}`, { method:"DELETE", headers:{"Prefer":""} }),
  updateUser:      (id,b)   => sbFetch(`users?id=eq.${id}`, { method:"PATCH", body:JSON.stringify(b) }),
  insertCompany:   (body)   => sbFetch("companies", { method:"POST", body:JSON.stringify(body) }),
  updateCompany:   (id,b)   => sbFetch(`companies?id=eq.${id}`, { method:"PATCH", body:JSON.stringify(b) }),
  deleteCompany:   (id)     => sbFetch(`companies?id=eq.${id}`, { method:"DELETE", headers:{"Prefer":""} }),
};

// ─── CONSTANTES ────────────────────────────────────────────────────────────
const DEFAULT_SCHEDULE = {
  lunes:    { start:"08:00", end:"17:00", active:true  },
  martes:   { start:"08:00", end:"17:00", active:true  },
  miercoles:{ start:"08:00", end:"17:00", active:true  },
  jueves:   { start:"08:00", end:"17:00", active:true  },
  viernes:  { start:"08:00", end:"17:00", active:true  },
  sabado:   { start:"08:00", end:"13:00", active:false },
  domingo:  { start:"08:00", end:"13:00", active:false },
};
const DAYS_ES    = ["domingo","lunes","martes","miercoles","jueves","viernes","sabado"];
const DAYS_LABEL = { lunes:"Lun", martes:"Mar", miercoles:"Mié", jueves:"Jue", viernes:"Vie", sabado:"Sáb", domingo:"Dom" };
const DAYS_FULL  = ["lunes","martes","miercoles","jueves","viernes","sabado","domingo"];

// ─── HELPERS ───────────────────────────────────────────────────────────────
const fmt      = d => new Date(d).toLocaleTimeString("es-CO",{ hour:"2-digit", minute:"2-digit", second:"2-digit" });
const fmtShort = d => new Date(d).toLocaleTimeString("es-CO",{ hour:"2-digit", minute:"2-digit" });
const fmtDate  = d => new Date(d).toLocaleDateString("es-CO",{ weekday:"short", year:"numeric", month:"short", day:"numeric" });
const diffH    = (a,b) => ((new Date(b)-new Date(a))/3600000).toFixed(1);
const todayKey = () => DAYS_ES[new Date().getDay()];
const getInits = n => (n||"?").split(" ").map(w=>w[0]).join("").slice(0,2).toUpperCase();
const sameDay  = (a,b) => new Date(a).toDateString()===new Date(b).toDateString();

function isLate(record, user) {
  if (!record) return null;
  const day = DAYS_ES[new Date(record.entry).getDay()];
  const sched = user?.schedule?.[day];
  if (!sched?.active) return null;
  const [sh,sm] = sched.start.split(":").map(Number);
  const entry = new Date(record.entry);
  const limit = new Date(entry); limit.setHours(sh,sm,0,0);
  return Math.round((entry - limit)/60000);
}

function exportCSV(records, users) {
  const rows = [["Empleado","Documento","Fecha","Entrada","Salida","Horas","Puntualidad"]];
  records.forEach(r => {
    const u = users.find(x=>x.id===(r.user_id||r.userId));
    const lm = isLate(r,u);
    const punt = lm===null?"—":lm>0?`Tarde ${lm}min`:`Temprano ${Math.abs(lm)}min`;
    rows.push([u?.name||"—",u?.documento||"—",fmtDate(r.entry),fmtShort(r.entry),r.exit?fmtShort(r.exit):"—",r.exit?diffH(r.entry,r.exit)+"h":"—",punt]);
  });
  const csv = rows.map(r=>r.map(c=>`"${c}"`).join(",")).join("\n");
  const blob = new Blob(["\uFEFF"+csv],{type:"text/csv;charset=utf-8;"});
  const a=document.createElement("a"); a.href=URL.createObjectURL(blob); a.download="asistencia.csv"; a.click();
}

function downloadTemplate(users) {
  const emps = users.filter(u=>u.role==="employee");
  const headers = ["nombre","documento","lunes_inicio","lunes_fin","lunes_activo","martes_inicio","martes_fin","martes_activo","miercoles_inicio","miercoles_fin","miercoles_activo","jueves_inicio","jueves_fin","jueves_activo","viernes_inicio","viernes_fin","viernes_activo","sabado_inicio","sabado_fin","sabado_activo","domingo_inicio","domingo_fin","domingo_activo"];
  const rows = emps.map(u => {
    const s = u.schedule||DEFAULT_SCHEDULE;
    return [u.name,u.documento,s.lunes.start,s.lunes.end,s.lunes.active?"SI":"NO",s.martes.start,s.martes.end,s.martes.active?"SI":"NO",s.miercoles.start,s.miercoles.end,s.miercoles.active?"SI":"NO",s.jueves.start,s.jueves.end,s.jueves.active?"SI":"NO",s.viernes.start,s.viernes.end,s.viernes.active?"SI":"NO",s.sabado.start,s.sabado.end,s.sabado.active?"SI":"NO",s.domingo.start,s.domingo.end,s.domingo.active?"SI":"NO"];
  });
  const ws = XLSX.utils.aoa_to_sheet([headers,...rows]);
  ws["!cols"] = headers.map((_,i)=>({wch:i<2?22:14}));
  const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb,"Horarios",ws); XLSX.writeFile(wb,"plantilla_horarios.xlsx");
}

function parseScheduleExcel(file, users) {
  return new Promise(resolve => {
    const reader = new FileReader();
    reader.onload = e => {
      try {
        const wb = XLSX.read(e.target.result,{type:"array"});
        const ws = wb.Sheets[wb.SheetNames[0]];
        const rows = XLSX.utils.sheet_to_json(ws,{header:1,defval:""});
        if(rows.length<2){resolve({ok:false,error:"El archivo está vacío."});return;}
        const headers = rows[0].map(h=>String(h).toLowerCase().trim());
        const docIdx = headers.indexOf("documento");
        if(docIdx===-1){resolve({ok:false,error:"No se encontró la columna 'documento'."});return;}
        const updated=[],errors=[];
        rows.slice(1).forEach((row,i)=>{
          const doc=String(row[docIdx]||"").trim();
          const user=users.find(u=>u.documento===doc);
          if(!user){errors.push(`Fila ${i+2}: documento '${doc}' no encontrado`);return;}
          const sched={};
          DAYS_FULL.forEach(day=>{
            const si=headers.indexOf(`${day}_inicio`),fi=headers.indexOf(`${day}_fin`),ai=headers.indexOf(`${day}_activo`);
            const sv=si!==-1?String(row[si]||"08:00").trim():"08:00";
            const ev=fi!==-1?String(row[fi]||"17:00").trim():"17:00";
            const av=ai!==-1?String(row[ai]||"NO").trim().toUpperCase():"NO";
            const pt=v=>{if(/^\d{1,2}:\d{2}$/.test(v))return v;const n=parseFloat(v);if(!isNaN(n)){const m=Math.round(n*1440);return `${String(Math.floor(m/60)).padStart(2,"0")}:${String(m%60).padStart(2,"0")}`;}return "08:00";};
            sched[day]={start:pt(sv),end:pt(ev),active:av==="SI"||av==="1"||av==="TRUE"};
          });
          updated.push({id:user.id,schedule:sched,name:user.name});
        });
        resolve({ok:true,updated,errors});
      } catch(err){resolve({ok:false,error:"Error al leer: "+err.message});}
    };
    reader.readAsArrayBuffer(file);
  });
}

// ─── ESTILOS ───────────────────────────────────────────────────────────────
const STYLES = `
@import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@300;400;500;600;700&family=Space+Grotesk:wght@500;600;700&display=swap');
*{box-sizing:border-box;margin:0;padding:0;}
body{background:#080d1a;}
input::placeholder{color:#4a5568;}
input:focus,select:focus{outline:none;border-color:#06b6d4!important;box-shadow:0 0 0 3px rgba(6,182,212,.15);}
.btn-grad{background:linear-gradient(135deg,#06b6d4,#3b82f6);transition:all .2s;}
.btn-grad:hover{filter:brightness(1.1);transform:translateY(-1px);box-shadow:0 8px 24px rgba(6,182,212,.3);}
.btn-entry{background:linear-gradient(135deg,#10b981,#059669);transition:all .2s;}
.btn-entry:hover{filter:brightness(1.1);transform:translateY(-2px);box-shadow:0 12px 30px rgba(16,185,129,.35);}
.btn-exit{background:linear-gradient(135deg,#f59e0b,#ef4444);transition:all .2s;}
.btn-exit:hover{filter:brightness(1.1);transform:translateY(-2px);box-shadow:0 12px 30px rgba(239,68,68,.35);}
.btn-danger{background:linear-gradient(135deg,#ef4444,#dc2626);transition:all .2s;}
.btn-danger:hover{filter:brightness(1.1);}
.btn-ghost{background:rgba(51,65,85,.5);border:1px solid rgba(71,85,105,.4);color:#94a3b8;transition:all .2s;}
.btn-ghost:hover{background:rgba(51,65,85,.8);color:#fff;}
.btn-amber{background:linear-gradient(135deg,#f59e0b,#d97706);transition:all .2s;}
.btn-amber:hover{filter:brightness(1.1);}
.card{background:rgba(15,23,42,.7);border:1px solid rgba(51,65,85,.5);border-radius:1rem;}
.modal-bg{position:fixed;inset:0;background:rgba(0,0,0,.75);backdrop-filter:blur(6px);z-index:50;display:flex;align-items:center;justify-content:center;padding:1rem;overflow-y:auto;}
.scrollbar::-webkit-scrollbar{width:4px;height:4px;}
.scrollbar::-webkit-scrollbar-thumb{background:#334155;border-radius:2px;}
.drop-zone{border:2px dashed rgba(6,182,212,.3);border-radius:1rem;transition:all .2s;}
.drop-zone:hover,.drop-zone.drag-over{border-color:rgba(6,182,212,.7);background:rgba(6,182,212,.05);}
@keyframes spin{to{transform:rotate(360deg)}}
.spin{animation:spin .8s linear infinite;}
`;

// ─── UI BASE ───────────────────────────────────────────────────────────────
function Badge({children,color="green"}){
  const c={green:"bg-emerald-500/20 text-emerald-300 border-emerald-500/30",red:"bg-red-500/20 text-red-300 border-red-500/30",amber:"bg-amber-500/20 text-amber-300 border-amber-500/30",blue:"bg-blue-500/20 text-blue-300 border-blue-500/30",purple:"bg-purple-500/20 text-purple-300 border-purple-500/30",slate:"bg-slate-500/20 text-slate-300 border-slate-500/30",cyan:"bg-cyan-500/20 text-cyan-300 border-cyan-500/30",orange:"bg-orange-500/20 text-orange-300 border-orange-500/30"};
  return <span className={`px-2 py-0.5 rounded-full text-xs font-semibold border ${c[color]||c.slate}`}>{children}</span>;
}
function Av({name,size="md"}){
  const s=size==="lg"?"w-14 h-14 text-lg":size==="sm"?"w-8 h-8 text-xs":"w-10 h-10 text-sm";
  return <div className={`${s} rounded-full bg-gradient-to-br from-cyan-500 to-blue-600 flex items-center justify-center font-bold text-white shadow-lg flex-shrink-0`}>{getInits(name)}</div>;
}
function Inp({label,...p}){return <div>{label&&<label className="block text-xs text-slate-400 mb-1">{label}</label>}<input className="w-full bg-slate-800/60 border border-slate-600/60 text-white rounded-xl px-3 py-2.5 text-sm transition-all" {...p}/></div>;}
function Sel({label,children,...p}){return <div>{label&&<label className="block text-xs text-slate-400 mb-1">{label}</label>}<select className="w-full bg-slate-800/60 border border-slate-600/60 text-white rounded-xl px-3 py-2.5 text-sm" {...p}>{children}</select></div>;}
function Spin(){return <div className="w-5 h-5 border-2 border-cyan-400 border-t-transparent rounded-full spin flex-shrink-0"/>;}
function Logo({size="sm"}){
  const s=size==="lg"?"w-14 h-14":"w-7 h-7",i=size==="lg"?"w-8 h-8":"w-4 h-4";
  return <div className={`${s} rounded-xl bg-gradient-to-br from-cyan-500 to-blue-600 flex items-center justify-center shadow-xl flex-shrink-0`}><svg className={`${i} text-white`} fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2}><path strokeLinecap="round" strokeLinejoin="round" d="M12 8v4l3 3m6-3a9 9 0 11-18 0 9 9 0 0118 0z"/></svg></div>;
}

// ─── MODALES ───────────────────────────────────────────────────────────────
function ScheduleModal({user,onSave,onClose}){
  const [sched,setSched]=useState(JSON.parse(JSON.stringify(user.schedule||DEFAULT_SCHEDULE)));
  const [saving,setSaving]=useState(false);
  const toggle=d=>setSched(p=>({...p,[d]:{...p[d],active:!p[d].active}}));
  const setT=(d,f,v)=>setSched(p=>({...p,[d]:{...p[d],[f]:v}}));
  const save=async()=>{setSaving(true);await onSave(sched);setSaving(false);};
  return(
    <div className="modal-bg" onClick={e=>e.target===e.currentTarget&&onClose()}>
      <div className="card p-6 w-full max-w-md" style={{fontFamily:"'DM Sans'",maxHeight:"90vh",overflowY:"auto"}}>
        <div className="flex items-center justify-between mb-5"><div><h3 className="font-bold text-white" style={{fontFamily:"'Space Grotesk'"}}>{user.name}</h3><p className="text-xs text-slate-400 mt-0.5">Horario semanal</p></div><button onClick={onClose} className="text-slate-400 hover:text-white text-xl">✕</button></div>
        <div className="space-y-2.5">{Object.entries(sched).map(([day,s])=>(
          <div key={day} className={`rounded-xl p-3 border transition-all ${s.active?"bg-slate-800/60 border-slate-600/50":"bg-slate-800/20 border-slate-700/30 opacity-60"}`}>
            <div className="flex items-center justify-between">
              <div className="flex items-center gap-3"><button onClick={()=>toggle(day)} className={`w-9 h-5 rounded-full transition-all relative flex-shrink-0 ${s.active?"bg-cyan-500":"bg-slate-600"}`}><span className={`absolute top-0.5 w-4 h-4 bg-white rounded-full shadow transition-all ${s.active?"left-4":"left-0.5"}`}/></button><span className="text-sm font-medium text-slate-200 w-12">{DAYS_LABEL[day]}</span></div>
              {s.active?<div className="flex items-center gap-2"><input type="time" value={s.start} onChange={e=>setT(day,"start",e.target.value)} className="bg-slate-700/60 border border-slate-600/50 text-white rounded-lg px-2 py-1 text-xs"/><span className="text-slate-500 text-xs">→</span><input type="time" value={s.end} onChange={e=>setT(day,"end",e.target.value)} className="bg-slate-700/60 border border-slate-600/50 text-white rounded-lg px-2 py-1 text-xs"/></div>:<span className="text-xs text-slate-500">Descanso</span>}
            </div>
          </div>
        ))}</div>
        <div className="flex gap-3 mt-5"><button onClick={onClose} className="btn-ghost flex-1 rounded-xl py-2.5 text-sm">Cancelar</button><button onClick={save} disabled={saving} className="btn-grad flex-1 text-white font-semibold rounded-xl py-2.5 text-sm flex items-center justify-center gap-2 disabled:opacity-60">{saving&&<Spin/>}{saving?"Guardando...":"Guardar"}</button></div>
      </div>
    </div>
  );
}

function AddUserModal({onSave,onClose,companyId}){
  const [form,setForm]=useState({name:"",documento:"",email:"",password:"",role:"employee"});
  const [err,setErr]=useState(""); const [saving,setSaving]=useState(false);
  const set=(k,v)=>setForm(p=>({...p,[k]:v}));
  const save=async()=>{
    if(!form.name.trim()||!form.documento.trim()||!form.password.trim()){setErr("Nombre, documento y contraseña son obligatorios");return;}
    setSaving(true); setErr("");
    try{await onSave({...form,avatar:getInits(form.name),schedule:JSON.parse(JSON.stringify(DEFAULT_SCHEDULE)),company_id:companyId});}
    catch(e){setErr(e.message);setSaving(false);}
  };
  return(
    <div className="modal-bg" onClick={e=>e.target===e.currentTarget&&onClose()}>
      <div className="card p-6 w-full max-w-sm" style={{fontFamily:"'DM Sans'"}}>
        <div className="flex items-center justify-between mb-5"><h3 className="font-bold text-white" style={{fontFamily:"'Space Grotesk'"}}>Agregar Usuario</h3><button onClick={onClose} className="text-slate-400 hover:text-white text-xl">✕</button></div>
        <div className="space-y-3">
          <Inp label="Nombre completo *" value={form.name} onChange={e=>set("name",e.target.value)} placeholder="Ej: María García"/>
          <Inp label="Número de documento *" type="text" value={form.documento} onChange={e=>set("documento",e.target.value)} placeholder="Ej: 12345678"/>
          <Inp label="Correo (opcional)" type="email" value={form.email} onChange={e=>set("email",e.target.value)} placeholder="maria@empresa.com"/>
          <Inp label="Contraseña *" type="text" value={form.password} onChange={e=>set("password",e.target.value)} placeholder="Contraseña inicial"/>
          <Sel label="Rol" value={form.role} onChange={e=>set("role",e.target.value)}><option value="employee">Empleado</option><option value="admin">Administrador</option></Sel>
          {err&&<p className="text-red-400 text-xs">{err}</p>}
        </div>
        <div className="flex gap-3 mt-5"><button onClick={onClose} className="btn-ghost flex-1 rounded-xl py-2.5 text-sm">Cancelar</button><button onClick={save} disabled={saving} className="btn-grad flex-1 text-white font-semibold rounded-xl py-2.5 text-sm flex items-center justify-center gap-2 disabled:opacity-60">{saving&&<Spin/>}{saving?"Guardando...":"Agregar"}</button></div>
      </div>
    </div>
  );
}

function ConfirmModal({title,message,onConfirm,onClose}){
  const [loading,setLoading]=useState(false);
  const go=async()=>{setLoading(true);await onConfirm();};
  return(
    <div className="modal-bg">
      <div className="card p-6 w-full max-w-sm text-center" style={{fontFamily:"'DM Sans'"}}>
        <div className="w-14 h-14 rounded-full bg-red-500/20 flex items-center justify-center mx-auto mb-4 text-2xl">⚠️</div>
        <h3 className="font-bold text-white mb-2" style={{fontFamily:"'Space Grotesk'"}}>{title||"Confirmar"}</h3>
        <p className="text-slate-400 text-sm mb-5">{message}</p>
        <div className="flex gap-3"><button onClick={onClose} className="btn-ghost flex-1 rounded-xl py-2.5 text-sm">Cancelar</button><button onClick={go} disabled={loading} className="btn-danger flex-1 text-white font-semibold rounded-xl py-2.5 text-sm flex items-center justify-center gap-2 disabled:opacity-60">{loading&&<Spin/>}Confirmar</button></div>
      </div>
    </div>
  );
}

function ImportModal({users,onApply,onClose}){
  const [drag,setDrag]=useState(false); const [file,setFile]=useState(null); const [result,setResult]=useState(null); const [loading,setLoading]=useState(false); const [saving,setSaving]=useState(false);
  const ref=useRef();
  const handle=async f=>{if(!f)return;setFile(f);setResult(null);setLoading(true);const r=await parseScheduleExcel(f,users);setResult(r);setLoading(false);};
  const onDrop=e=>{e.preventDefault();setDrag(false);handle(e.dataTransfer.files[0]);};
  const apply=async()=>{if(!result?.ok)return;setSaving(true);await onApply(result.updated);setSaving(false);};
  return(
    <div className="modal-bg" onClick={e=>e.target===e.currentTarget&&onClose()}>
      <div className="card p-6 w-full max-w-lg" style={{fontFamily:"'DM Sans'",maxHeight:"90vh",overflowY:"auto"}}>
        <div className="flex items-center justify-between mb-5"><div><h3 className="font-bold text-white text-lg" style={{fontFamily:"'Space Grotesk'"}}>Importar Horarios</h3><p className="text-xs text-slate-400 mt-0.5">Sube un Excel con los horarios semanales</p></div><button onClick={onClose} className="text-slate-400 hover:text-white text-xl ml-4">✕</button></div>
        <div className="bg-slate-800/40 border border-slate-700/40 rounded-xl p-4 mb-4">
          <div className="flex items-start gap-3"><div className="w-8 h-8 rounded-lg bg-emerald-500/20 flex items-center justify-center flex-shrink-0">📥</div>
            <div className="flex-1"><p className="text-sm font-semibold text-white">Paso 1 — Descarga la plantilla</p><p className="text-xs text-slate-400 mt-0.5 mb-3">Incluye tus empleados con sus horarios actuales.</p>
              <button onClick={()=>downloadTemplate(users)} className="btn-grad text-white text-xs font-semibold rounded-lg px-3 py-2 flex items-center gap-1.5 w-fit"><svg className="w-3.5 h-3.5" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2}><path strokeLinecap="round" strokeLinejoin="round" d="M4 16v1a3 3 0 003 3h10a3 3 0 003-3v-1m-4-4l-4 4m0 0l-4-4m4 4V4"/></svg>Descargar plantilla</button>
            </div>
          </div>
        </div>
        <div className={`drop-zone p-6 text-center cursor-pointer mb-4 ${drag?"drag-over":""}`} onDragOver={e=>{e.preventDefault();setDrag(true);}} onDragLeave={()=>setDrag(false)} onDrop={onDrop} onClick={()=>ref.current.click()}>
          <input ref={ref} type="file" accept=".xlsx,.xls" className="hidden" onChange={e=>handle(e.target.files[0])}/>
          {!file?<><div className="text-3xl mb-2">📂</div><p className="text-sm text-slate-300 font-medium">Arrastra tu Excel aquí</p><p className="text-xs text-slate-500 mt-1">o haz clic para seleccionar</p></>:<><div className="text-3xl mb-2">📊</div><p className="text-sm text-slate-300 font-medium">{file.name}</p></>}
        </div>
        {loading&&<div className="text-center py-4 flex flex-col items-center gap-2"><Spin/><p className="text-xs text-slate-400">Analizando...</p></div>}
        {result&&!loading&&(!result.ok?<div className="bg-red-500/10 border border-red-500/30 rounded-xl p-4 text-sm text-red-300 mb-4">❌ {result.error}</div>
          :<div className="space-y-3 mb-4"><div className="bg-emerald-500/10 border border-emerald-500/30 rounded-xl p-4"><p className="text-sm font-semibold text-emerald-300 mb-2">✅ {result.updated.length} empleado{result.updated.length!==1?"s":""} encontrado{result.updated.length!==1?"s":""}</p><div className="space-y-1 max-h-36 overflow-y-auto scrollbar">{result.updated.map(u=><div key={u.id} className="flex items-center gap-2 text-xs text-slate-300"><span className="w-1.5 h-1.5 rounded-full bg-emerald-400 flex-shrink-0"/><span className="font-medium">{u.name}</span></div>)}</div></div>{result.errors.length>0&&<div className="bg-amber-500/10 border border-amber-500/30 rounded-xl p-4"><p className="text-xs font-semibold text-amber-300 mb-1">⚠️ {result.errors.length} advertencia(s)</p>{result.errors.map((e,i)=><p key={i} className="text-xs text-amber-200/70">{e}</p>)}</div>}</div>)}
        <div className="flex gap-3"><button onClick={onClose} className="btn-ghost flex-1 rounded-xl py-2.5 text-sm">Cancelar</button><button onClick={apply} disabled={!result?.ok||loading||saving} className="btn-grad flex-1 text-white font-semibold rounded-xl py-2.5 text-sm flex items-center justify-center gap-2 disabled:opacity-40">{saving&&<Spin/>}{saving?"Guardando...":"Aplicar horarios"}</button></div>
      </div>
    </div>
  );
}

// ─── MODAL AGREGAR EMPRESA ─────────────────────────────────────────────────
function AddCompanyModal({onSave,onClose}){
  const [form,setForm]=useState({name:"",nit:"",admin_name:"",admin_doc:"",admin_pass:""});
  const [err,setErr]=useState(""); const [saving,setSaving]=useState(false);
  const set=(k,v)=>setForm(p=>({...p,[k]:v}));
  const save=async()=>{
    if(!form.name.trim()||!form.nit.trim()||!form.admin_name.trim()||!form.admin_doc.trim()||!form.admin_pass.trim()){setErr("Todos los campos son obligatorios");return;}
    setSaving(true); setErr("");
    try{await onSave(form);}catch(e){setErr(e.message);setSaving(false);}
  };
  return(
    <div className="modal-bg" onClick={e=>e.target===e.currentTarget&&onClose()}>
      <div className="card p-6 w-full max-w-md" style={{fontFamily:"'DM Sans'",maxHeight:"90vh",overflowY:"auto"}}>
        <div className="flex items-center justify-between mb-5"><div><h3 className="font-bold text-white" style={{fontFamily:"'Space Grotesk'"}}>Nueva Empresa</h3><p className="text-xs text-slate-400 mt-0.5">Registra una nueva empresa cliente</p></div><button onClick={onClose} className="text-slate-400 hover:text-white text-xl">✕</button></div>
        <div className="space-y-3">
          <p className="text-xs text-cyan-400 font-semibold uppercase tracking-widest">Datos de la empresa</p>
          <Inp label="Nombre de la empresa *" value={form.name} onChange={e=>set("name",e.target.value)} placeholder="Ej: OXI 50 S.A.S"/>
          <Inp label="NIT *" value={form.nit} onChange={e=>set("nit",e.target.value)} placeholder="Ej: 900123456-1"/>
          <p className="text-xs text-cyan-400 font-semibold uppercase tracking-widest pt-2">Administrador de la empresa</p>
          <Inp label="Nombre del admin *" value={form.admin_name} onChange={e=>set("admin_name",e.target.value)} placeholder="Ej: Juan Pérez"/>
          <Inp label="Documento del admin *" value={form.admin_doc} onChange={e=>set("admin_doc",e.target.value)} placeholder="Número de cédula"/>
          <Inp label="Contraseña del admin *" type="text" value={form.admin_pass} onChange={e=>set("admin_pass",e.target.value)} placeholder="Contraseña inicial"/>
          {err&&<p className="text-red-400 text-xs">{err}</p>}
        </div>
        <div className="flex gap-3 mt-5"><button onClick={onClose} className="btn-ghost flex-1 rounded-xl py-2.5 text-sm">Cancelar</button><button onClick={save} disabled={saving} className="btn-grad flex-1 text-white font-semibold rounded-xl py-2.5 text-sm flex items-center justify-center gap-2 disabled:opacity-60">{saving&&<Spin/>}{saving?"Creando...":"Crear empresa"}</button></div>
      </div>
    </div>
  );
}

// ─── SUPER ADMIN PANEL ─────────────────────────────────────────────────────
function SuperAdminPanel({currentUser,onLogout}){
  const [companies,setCompanies]=useState([]);
  const [allUsers,setAllUsers]=useState([]);
  const [allRecords,setAllRecords]=useState([]);
  const [loading,setLoading]=useState(true);
  const [showAdd,setShowAdd]=useState(false);
  const [confirm,setConfirm]=useState(null);
  const [selectedCompany,setSelectedCompany]=useState(null);
  const [tab,setTab]=useState("companies");
  const [msg,setMsg]=useState(null);

  const loadAll=async()=>{
    setLoading(true);
    try{
      const [c,u,r]=await Promise.all([db.getCompanies(),db.getUsers(),db.getRecords()]);
      setCompanies(c||[]); setAllUsers(u||[]); setAllRecords(r||[]);
    }catch(e){console.error(e);}
    setLoading(false);
  };
  useEffect(()=>{loadAll();},[]);

  const showMsg=(text,type="success")=>{setMsg({text,type});setTimeout(()=>setMsg(null),5000);};

  const createCompany=async(form)=>{
    // 1. Crear empresa
    const compRes=await db.insertCompany({name:form.name,nit:form.nit,admin_name:form.admin_name,status:"active"});
    if(!compRes||!compRes[0])throw new Error("Error creando empresa");
    const company=compRes[0];
    // 2. Crear admin de la empresa
    await db.insertUser({name:form.admin_name,documento:form.admin_doc,email:`admin@${form.nit.replace(/[^a-z0-9]/gi,"")}.com`,password:form.admin_pass,role:"admin",avatar:getInits(form.admin_name),schedule:{},company_id:company.id});
    setShowAdd(false);
    await loadAll();
    showMsg(`✅ Empresa "${form.name}" creada exitosamente`);
  };

  const toggleStatus=async(company)=>{
    const newStatus=company.status==="active"?"suspended":"active";
    await db.updateCompany(company.id,{status:newStatus});
    setCompanies(p=>p.map(c=>c.id===company.id?{...c,status:newStatus}:c));
    showMsg(`${newStatus==="active"?"✅ Empresa reactivada":"⏸️ Empresa suspendida"}`);
    setConfirm(null);
  };

  const deleteCompany=async(company)=>{
    // Eliminar usuarios de la empresa primero
    const empUsers=allUsers.filter(u=>u.company_id===company.id);
    for(const u of empUsers) await db.deleteUser(u.id);
    await db.deleteCompany(company.id);
    await loadAll();
    showMsg(`🗑️ Empresa "${company.name}" eliminada`);
    setConfirm(null);
  };

  const companyUsers=selectedCompany?allUsers.filter(u=>u.company_id===selectedCompany.id):[];
  const companyRecords=selectedCompany?allRecords.filter(r=>r.company_id===selectedCompany.id):[];
  const activeCompanies=companies.filter(c=>c.status==="active").length;
  const totalEmployees=allUsers.filter(u=>u.role==="employee").length;

  return(
    <div className="min-h-screen bg-[#080d1a] text-white" style={{fontFamily:"'DM Sans'"}}>
      <style>{STYLES}</style>
      {showAdd&&<AddCompanyModal onSave={createCompany} onClose={()=>setShowAdd(false)}/>}
      {confirm&&<ConfirmModal title={confirm.title} message={confirm.message} onConfirm={confirm.action} onClose={()=>setConfirm(null)}/>}

      {/* Header */}
      <div className="bg-slate-900/80 backdrop-blur border-b border-slate-700/50 px-5 py-3.5 flex items-center justify-between sticky top-0 z-10">
        <div className="flex items-center gap-2.5"><Logo/><span className="font-bold text-sm" style={{fontFamily:"'Space Grotesk'"}}>Control Asistencia Pro</span><Badge color="orange">Super Admin</Badge></div>
        <div className="flex items-center gap-3"><Av name={currentUser.name} size="sm"/><button onClick={onLogout} className="text-slate-400 hover:text-red-400 transition-colors text-sm">Salir</button></div>
      </div>

      {/* Tabs */}
      <div className="bg-slate-900/50 border-b border-slate-700/40 px-5">
        <div className="flex gap-1 max-w-5xl mx-auto">
          {[["companies","🏢 Empresas"],["reports","📊 Reportes"]].map(([key,label])=><button key={key} onClick={()=>{setTab(key);setSelectedCompany(null);}} className={`px-4 py-3 text-sm font-medium transition-all border-b-2 ${tab===key?"border-cyan-400 text-cyan-300":"border-transparent text-slate-400 hover:text-white"}`}>{label}</button>)}
        </div>
      </div>

      <div className="max-w-5xl mx-auto px-4 py-7 space-y-5">
        {msg&&<div className={`rounded-xl px-5 py-3 text-sm ${msg.type==="error"?"bg-red-500/10 border border-red-500/30 text-red-300":"bg-emerald-500/10 border border-emerald-500/30 text-emerald-300"}`}>{msg.text}</div>}

        {loading?<div className="flex items-center justify-center py-20"><Spin/></div>:(<>

        {/* ── EMPRESAS ── */}
        {tab==="companies"&&!selectedCompany&&<>
          <div className="flex items-center justify-between flex-wrap gap-3">
            <div><h1 className="text-xl font-bold" style={{fontFamily:"'Space Grotesk'"}}>Panel Super Admin</h1><p className="text-slate-400 text-sm mt-0.5">Gestiona todas las empresas clientes</p></div>
            <button onClick={()=>setShowAdd(true)} className="btn-grad text-white text-sm font-semibold rounded-xl px-4 py-2 flex items-center gap-2"><svg className="w-4 h-4" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2.5}><path strokeLinecap="round" strokeLinejoin="round" d="M12 4v16m8-8H4"/></svg>Nueva empresa</button>
          </div>
          {/* KPIs */}
          <div className="grid grid-cols-2 md:grid-cols-3 gap-4">
            {[{label:"Empresas activas",value:activeCompanies,icon:"🏢",color:"from-cyan-500 to-blue-600"},{label:"Total empleados",value:totalEmployees,icon:"👥",color:"from-emerald-500 to-teal-600"},{label:"Total empresas",value:companies.length,icon:"📋",color:"from-violet-500 to-purple-600"}].map(k=>(
              <div key={k.label} className="card p-5 relative overflow-hidden"><div className={`absolute -right-3 -top-3 w-16 h-16 rounded-full bg-gradient-to-br ${k.color} opacity-10`}/><p className="text-2xl mb-1">{k.icon}</p><p className="text-3xl font-bold" style={{fontFamily:"'Space Grotesk'"}}>{k.value}</p><p className="text-xs text-slate-400 mt-0.5">{k.label}</p></div>
            ))}
          </div>
          {/* Lista empresas */}
          <div className="card overflow-hidden">
            <div className="px-5 py-4 border-b border-slate-700/40"><p className="text-xs text-slate-500 uppercase tracking-widest">Empresas registradas</p></div>
            {companies.length===0?<p className="text-slate-500 text-sm text-center py-10">No hay empresas registradas aún</p>
              :<div className="divide-y divide-slate-700/20">{companies.map(c=>{
                const empCount=allUsers.filter(u=>u.company_id===c.id&&u.role==="employee").length;
                const suspended=c.status==="suspended";
                return(
                  <div key={c.id} className="px-5 py-4 flex items-center justify-between hover:bg-slate-800/20 transition-colors">
                    <div className="flex items-center gap-3 flex-1 min-w-0">
                      <div className={`w-10 h-10 rounded-xl flex items-center justify-center text-lg flex-shrink-0 ${suspended?"bg-slate-700/50":"bg-cyan-500/20"}`}>🏢</div>
                      <div className="min-w-0">
                        <div className="flex items-center gap-2 flex-wrap"><p className="text-sm font-semibold">{c.name}</p><Badge color={suspended?"red":"green"}>{suspended?"Suspendida":"Activa"}</Badge></div>
                        <p className="text-xs text-slate-500 mt-0.5">NIT: {c.nit} · {empCount} empleado{empCount!==1?"s":""}</p>
                      </div>
                    </div>
                    <div className="flex items-center gap-2 flex-shrink-0 ml-3">
                      <button onClick={()=>setSelectedCompany(c)} className="text-xs btn-ghost rounded-lg px-3 py-1.5">👁️ Ver</button>
                      <button onClick={()=>setConfirm({title:suspended?"Reactivar empresa":"Suspender empresa",message:suspended?`¿Reactivar "${c.name}"?`:`¿Suspender "${c.name}"? Los usuarios no podrán acceder.`,action:()=>toggleStatus(c)})} className={`text-xs rounded-lg px-3 py-1.5 text-white ${suspended?"btn-grad":"btn-amber"}`}>{suspended?"▶️ Activar":"⏸️ Suspender"}</button>
                      <button onClick={()=>setConfirm({title:"Eliminar empresa",message:`¿Eliminar "${c.name}" y todos sus datos? Esta acción no se puede deshacer.`,action:()=>deleteCompany(c)})} className="text-xs btn-danger rounded-lg px-3 py-1.5 text-white">🗑️</button>
                    </div>
                  </div>
                );
              })}</div>}
          </div>
        </>}

        {/* ── DETALLE EMPRESA ── */}
        {tab==="companies"&&selectedCompany&&<>
          <div className="flex items-center gap-3">
            <button onClick={()=>setSelectedCompany(null)} className="btn-ghost rounded-xl px-3 py-2 text-sm flex items-center gap-1.5">← Volver</button>
            <div><h1 className="text-xl font-bold" style={{fontFamily:"'Space Grotesk'"}}>{selectedCompany.name}</h1><p className="text-slate-400 text-sm">NIT: {selectedCompany.nit}</p></div>
          </div>
          <div className="grid grid-cols-2 md:grid-cols-3 gap-4">
            {[{label:"Empleados",value:companyUsers.filter(u=>u.role==="employee").length,icon:"👥"},{label:"Registros totales",value:companyRecords.length,icon:"📋"},{label:"Activos hoy",value:companyRecords.filter(r=>!r.exit&&new Date(r.entry).toDateString()===new Date().toDateString()).length,icon:"✅"}].map(k=>(
              <div key={k.label} className="card p-5"><p className="text-2xl mb-1">{k.icon}</p><p className="text-3xl font-bold" style={{fontFamily:"'Space Grotesk'"}}>{k.value}</p><p className="text-xs text-slate-400 mt-0.5">{k.label}</p></div>
            ))}
          </div>
          <div className="card overflow-hidden">
            <div className="px-5 py-4 border-b border-slate-700/40"><p className="text-xs text-slate-500 uppercase tracking-widest">Usuarios de {selectedCompany.name}</p></div>
            {companyUsers.length===0?<p className="text-slate-500 text-sm text-center py-8">Sin usuarios</p>
              :<div className="divide-y divide-slate-700/20">{companyUsers.map(u=><div key={u.id} className="px-5 py-3.5 flex items-center justify-between"><div className="flex items-center gap-3"><Av name={u.name} size="sm"/><div><p className="text-sm font-medium">{u.name}</p><p className="text-xs text-slate-500">Doc: {u.documento}</p></div></div><Badge color={u.role==="admin"?"purple":"blue"}>{u.role==="admin"?"Admin":"Empleado"}</Badge></div>)}</div>}
          </div>
        </>}

        {/* ── REPORTES ── */}
        {tab==="reports"&&<>
          <h1 className="text-xl font-bold" style={{fontFamily:"'Space Grotesk'"}}>Reportes globales</h1>
          <div className="card overflow-hidden">
            <div className="px-5 py-4 border-b border-slate-700/40"><p className="text-xs text-slate-500 uppercase tracking-widest">Resumen por empresa</p></div>
            {companies.length===0?<p className="text-slate-500 text-sm text-center py-8">Sin datos</p>
              :<div className="overflow-x-auto scrollbar"><table className="w-full"><thead><tr className="border-b border-slate-700/30">{["Empresa","NIT","Estado","Empleados","Registros totales"].map(h=><th key={h} className="text-left text-xs text-slate-500 uppercase tracking-wider px-4 py-3 font-medium whitespace-nowrap">{h}</th>)}</tr></thead>
              <tbody className="divide-y divide-slate-700/20">{companies.map(c=>{
                const emps=allUsers.filter(u=>u.company_id===c.id&&u.role==="employee").length;
                const recs=allRecords.filter(r=>r.company_id===c.id).length;
                return(<tr key={c.id} className="hover:bg-slate-800/30 transition-colors">
                  <td className="px-4 py-3 text-sm font-medium text-white">{c.name}</td>
                  <td className="px-4 py-3 text-sm text-slate-400">{c.nit}</td>
                  <td className="px-4 py-3"><Badge color={c.status==="active"?"green":"red"}>{c.status==="active"?"Activa":"Suspendida"}</Badge></td>
                  <td className="px-4 py-3 text-sm text-slate-300">{emps}</td>
                  <td className="px-4 py-3 text-sm text-slate-300">{recs}</td>
                </tr>);
              })}</tbody></table></div>}
          </div>
        </>}
        </>)}
      </div>
    </div>
  );
}

// ─── PANEL EMPLEADO ────────────────────────────────────────────────────────
function EmployeePanel({user,records,onEntry,onExit,onLogout,actionLoading}){
  const [now,setNow]=useState(new Date());
  useEffect(()=>{const t=setInterval(()=>setNow(new Date()),1000);return()=>clearInterval(t);},[]);
  const mine   = records.filter(r=>(r.user_id||r.userId)===user.id).sort((a,b)=>new Date(b.entry)-new Date(a.entry));
  const today  = mine.find(r=>sameDay(r.entry,now));
  const isIn   = today&&!today.exit;
  const sched  = user.schedule?.[todayKey()];
  const lateMin= today ? isLate(today,user) : null;
  return(
    <div className="min-h-screen bg-[#080d1a] text-white" style={{fontFamily:"'DM Sans'"}}>
      <style>{STYLES}</style>
      <div className="bg-slate-900/80 backdrop-blur border-b border-slate-700/50 px-5 py-3.5 flex items-center justify-between sticky top-0 z-10">
        <div className="flex items-center gap-2.5"><Logo/><span className="font-bold text-sm" style={{fontFamily:"'Space Grotesk'"}}>Control Asistencia Pro</span></div>
        <div className="flex items-center gap-3"><Av name={user.name} size="sm"/><button onClick={onLogout} className="text-slate-400 hover:text-red-400 transition-colors text-sm">Salir</button></div>
      </div>
      <div className="max-w-xl mx-auto px-4 py-7 space-y-5">
        <div><p className="text-slate-400 text-sm">Bienvenido,</p><h1 className="text-2xl font-bold" style={{fontFamily:"'Space Grotesk'"}}>{user.name}</h1></div>
        <div className="card p-6 text-center relative overflow-hidden">
          <div className="absolute inset-0 bg-gradient-to-br from-cyan-500/5 to-blue-500/5"/>
          <p className="text-slate-400 text-sm">{fmtDate(now)}</p>
          <p className="text-5xl font-bold tracking-tight text-white mt-1" style={{fontFamily:"'Space Grotesk'"}}>{fmt(now)}</p>
          <div className="flex items-center justify-center gap-2 mt-3 flex-wrap">
            {isIn?<Badge color="green">● EN OFICINA</Badge>:<Badge color="red">● FUERA</Badge>}
            {sched?.active&&<Badge color="slate">Horario: {sched.start} – {sched.end}</Badge>}
          </div>
        </div>
        {lateMin!==null&&(
          <div className={`card p-4 flex items-center gap-3 ${lateMin>0?"border-red-500/30":"border-emerald-500/30"}`}>
            <div className={`w-10 h-10 rounded-full flex items-center justify-center text-lg flex-shrink-0 ${lateMin>0?"bg-red-500/20":"bg-emerald-500/20"}`}>{lateMin>0?"⏰":"🌟"}</div>
            <div className="flex-1"><p className="font-semibold text-sm">{lateMin>0?`Llegaste ${lateMin} min tarde`:`Llegaste ${Math.abs(lateMin)} min antes`}</p><p className="text-xs text-slate-400">Hora límite: {sched?.start}</p></div>
            {lateMin>0?<Badge color="red">Tarde</Badge>:<Badge color="green">Puntual</Badge>}
          </div>
        )}
        <div className="flex justify-center py-2">
          {actionLoading?<div className="flex items-center gap-3 text-slate-400"><Spin/><span className="text-sm">Procesando...</span></div>
            :!today?<button onClick={onEntry} className="btn-entry text-white font-bold rounded-2xl px-10 py-5 text-lg flex items-center gap-3"><svg className="w-6 h-6" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2.5}><path strokeLinecap="round" strokeLinejoin="round" d="M11 16l-4-4m0 0l4-4m-4 4h14"/></svg>Registrar Entrada</button>
            :isIn?<button onClick={onExit} className="btn-exit text-white font-bold rounded-2xl px-10 py-5 text-lg flex items-center gap-3"><svg className="w-6 h-6" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2.5}><path strokeLinecap="round" strokeLinejoin="round" d="M17 16l4-4m0 0l-4-4m4 4H7"/></svg>Registrar Salida</button>
            :<div className="text-center"><p className="text-slate-400 text-sm">Jornada completada ✓</p><p className="text-emerald-400 font-semibold mt-1">{diffH(today.entry,today.exit)} horas trabajadas</p></div>}
        </div>
        {today&&<div className="card p-5"><p className="text-xs text-slate-500 uppercase tracking-widest mb-3">Registro de hoy</p><div className="grid grid-cols-2 gap-3"><div className="bg-slate-800/50 rounded-xl p-3"><p className="text-xs text-slate-500 mb-1">Entrada</p><p className="font-bold text-emerald-400 text-lg">{fmtShort(today.entry)}</p></div><div className="bg-slate-800/50 rounded-xl p-3"><p className="text-xs text-slate-500 mb-1">Salida</p><p className={`font-bold text-lg ${today.exit?"text-amber-400":"text-slate-500"}`}>{today.exit?fmtShort(today.exit):"—"}</p></div></div></div>}
        <div className="card overflow-hidden">
          <div className="px-5 py-4 border-b border-slate-700/40"><p className="text-xs text-slate-500 uppercase tracking-widest">Historial reciente</p></div>
          {mine.length===0?<p className="text-slate-500 text-sm text-center py-8">Sin registros</p>
            :<div className="divide-y divide-slate-700/30">{mine.slice(0,8).map(r=>{const lm=isLate(r,user);return(<div key={r.id} className="px-5 py-3 flex items-center justify-between"><div><p className="text-sm text-slate-300">{fmtDate(r.entry)}</p><p className="text-xs text-slate-500 mt-0.5">{fmtShort(r.entry)}{r.exit?` → ${fmtShort(r.exit)}`:""}</p></div><div className="flex items-center gap-2">{r.exit&&<Badge color="blue">{diffH(r.entry,r.exit)}h</Badge>}{lm!==null&&(lm>0?<Badge color="red">{lm}m tarde</Badge>:<Badge color="green">Puntual</Badge>)}</div></div>);})}</div>}
        </div>
      </div>
    </div>
  );
}

// ─── PANEL ADMIN ───────────────────────────────────────────────────────────
function AdminPanel({users,setUsers,records,setRecords,currentUser,onLogout}){
  const [tab,setTab]=useState("dashboard");
  const [schedUser,setSchedUser]=useState(null);
  const [delUser,setDelUser]=useState(null);
  const [showAdd,setShowAdd]=useState(false);
  const [showImport,setShowImport]=useState(false);
  const [importMsg,setImportMsg]=useState(null);
  const [search,setSearch]=useState("");
  const [filter,setFilter]=useState("all");

  const employees   = users.filter(u=>u.role==="employee");
  const todayStr    = new Date().toDateString();
  const todayRecs   = records.filter(r=>new Date(r.entry).toDateString()===todayStr);
  const activeNow   = records.filter(r=>!r.exit).length;
  const lateToday   = todayRecs.filter(r=>{const u=users.find(x=>x.id===(r.user_id||r.userId));const lm=isLate(r,u);return lm!==null&&lm>0;}).length;
  const onTimeToday = todayRecs.filter(r=>{const u=users.find(x=>x.id===(r.user_id||r.userId));const lm=isLate(r,u);return lm!==null&&lm<=0;}).length;

  const filtRecs = records.filter(r=>{
    const u=users.find(x=>x.id===(r.user_id||r.userId)); if(!u)return false;
    if(search&&!u.name.toLowerCase().includes(search.toLowerCase()))return false;
    if(filter==="active")return !r.exit;
    if(filter==="complete")return !!r.exit;
    if(filter==="late"){const lm=isLate(r,u);return lm!==null&&lm>0;}
    if(filter==="ontime"){const lm=isLate(r,u);return lm!==null&&lm<=0;}
    return true;
  }).sort((a,b)=>new Date(b.entry)-new Date(a.entry));

  const saveSchedule=async(uid,sched)=>{
    await db.updateUser(uid,{schedule:sched});
    setUsers(p=>p.map(u=>u.id===uid?{...u,schedule:sched}:u)); setSchedUser(null);
  };
  const addUser=async data=>{
    const res=await db.insertUser({...data,company_id:currentUser.company_id});
    if(res&&res[0])setUsers(p=>[...p,res[0]]); setShowAdd(false);
  };
  const deleteUser=async id=>{
    await db.deleteUser(id); setUsers(p=>p.filter(u=>u.id!==id)); setDelUser(null);
  };
  const applyImport=async updated=>{
    for(const u of updated) await db.updateUser(u.id,{schedule:u.schedule});
    setUsers(p=>p.map(u=>{const f=updated.find(x=>x.id===u.id);return f?{...u,schedule:f.schedule}:u;}));
    setShowImport(false);
    setImportMsg(`✅ Horarios actualizados para ${updated.length} empleado${updated.length!==1?"s":""}`);
    setTimeout(()=>setImportMsg(null),5000);
  };

  const TABS=[["dashboard","📊 Dashboard"],["records","📋 Registros"],["users","👥 Usuarios"]];
  return(
    <div className="min-h-screen bg-[#080d1a] text-white" style={{fontFamily:"'DM Sans'"}}>
      <style>{STYLES}</style>
      {schedUser&&<ScheduleModal user={schedUser} onSave={s=>saveSchedule(schedUser.id,s)} onClose={()=>setSchedUser(null)}/>}
      {delUser&&<ConfirmModal title="Eliminar usuario" message={`¿Eliminar a ${delUser.name}?`} onConfirm={()=>deleteUser(delUser.id)} onClose={()=>setDelUser(null)}/>}
      {showAdd&&<AddUserModal onSave={addUser} onClose={()=>setShowAdd(false)} companyId={currentUser.company_id}/>}
      {showImport&&<ImportModal users={users} onApply={applyImport} onClose={()=>setShowImport(false)}/>}

      <div className="bg-slate-900/80 backdrop-blur border-b border-slate-700/50 px-5 py-3.5 flex items-center justify-between sticky top-0 z-10">
        <div className="flex items-center gap-2.5"><Logo/><span className="font-bold text-sm" style={{fontFamily:"'Space Grotesk'"}}>Control Asistencia Pro</span><Badge color="purple">Admin</Badge></div>
        <div className="flex items-center gap-2">
          <button onClick={()=>exportCSV(records,users)} className="btn-grad text-white text-xs font-semibold rounded-xl px-3 py-2 flex items-center gap-1.5"><svg className="w-3.5 h-3.5" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2}><path strokeLinecap="round" strokeLinejoin="round" d="M4 16v1a3 3 0 003 3h10a3 3 0 003-3v-1m-4-4l-4 4m0 0l-4-4m4 4V4"/></svg>Exportar</button>
          <button onClick={onLogout} className="text-slate-400 hover:text-red-400 transition-colors text-sm">Salir</button>
        </div>
      </div>
      <div className="bg-slate-900/50 border-b border-slate-700/40 px-5">
        <div className="flex gap-1 max-w-5xl mx-auto">{TABS.map(([key,label])=><button key={key} onClick={()=>setTab(key)} className={`px-4 py-3 text-sm font-medium transition-all border-b-2 ${tab===key?"border-cyan-400 text-cyan-300":"border-transparent text-slate-400 hover:text-white"}`}>{label}</button>)}</div>
      </div>

      <div className="max-w-5xl mx-auto px-4 py-7 space-y-5">
        {importMsg&&<div className="bg-emerald-500/10 border border-emerald-500/30 rounded-xl px-5 py-3 text-sm text-emerald-300">{importMsg}</div>}

        {tab==="dashboard"&&<>
          <h1 className="text-xl font-bold" style={{fontFamily:"'Space Grotesk'"}}>Dashboard de hoy</h1>
          <div className="grid grid-cols-2 md:grid-cols-4 gap-4">
            {[{label:"Total empleados",value:employees.length,icon:"👥",color:"from-cyan-500 to-blue-600"},{label:"En oficina ahora",value:activeNow,icon:"✅",color:"from-emerald-500 to-teal-600"},{label:"Llegaron tarde hoy",value:lateToday,icon:"⏰",color:"from-red-500 to-rose-600"},{label:"Llegaron a tiempo",value:onTimeToday,icon:"🌟",color:"from-violet-500 to-purple-600"}].map(k=>(
              <div key={k.label} className="card p-5 relative overflow-hidden"><div className={`absolute -right-3 -top-3 w-16 h-16 rounded-full bg-gradient-to-br ${k.color} opacity-10`}/><p className="text-2xl mb-1">{k.icon}</p><p className="text-3xl font-bold" style={{fontFamily:"'Space Grotesk'"}}>{k.value}</p><p className="text-xs text-slate-400 mt-0.5">{k.label}</p></div>
            ))}
          </div>
          <div className="card overflow-hidden">
            <div className="px-5 py-4 border-b border-slate-700/40"><p className="text-xs text-slate-500 uppercase tracking-widest">Estado actual del personal</p></div>
            <div className="divide-y divide-slate-700/20">{employees.map(emp=>{
              const tr=records.filter(r=>(r.user_id||r.userId)===emp.id&&new Date(r.entry).toDateString()===todayStr).sort((a,b)=>new Date(b.entry)-new Date(a.entry))[0];
              const active=tr&&!tr.exit; const lm=tr?isLate(tr,emp):null;
              return(<div key={emp.id} className="px-5 py-3.5 flex items-center justify-between hover:bg-slate-800/20 transition-colors">
                <div className="flex items-center gap-3"><div className="relative"><Av name={emp.name} size="sm"/><div className={`absolute -bottom-0.5 -right-0.5 w-3 h-3 rounded-full border-2 border-slate-900 ${active?"bg-emerald-400":"bg-slate-600"}`}/></div><div><p className="text-sm font-medium">{emp.name}</p><p className="text-xs text-slate-500">Doc: {emp.documento}</p></div></div>
                <div className="flex items-center gap-2 flex-wrap justify-end">{active?<Badge color="green">En oficina</Badge>:tr?<Badge color="amber">Salió</Badge>:<Badge color="red">Sin registro</Badge>}{lm!==null&&(lm>0?<Badge color="red">{lm}m tarde</Badge>:<Badge color="green">Puntual</Badge>)}{tr&&!active&&<Badge color="blue">{diffH(tr.entry,tr.exit)}h</Badge>}</div>
              </div>);
            })}</div>
          </div>
        </>}

        {tab==="records"&&<>
          <h1 className="text-xl font-bold" style={{fontFamily:"'Space Grotesk'"}}>Registros de asistencia</h1>
          <div className="flex flex-col sm:flex-row gap-3">
            <input type="text" placeholder="Buscar empleado..." value={search} onChange={e=>setSearch(e.target.value)} className="flex-1 bg-slate-900/70 border border-slate-700/50 text-white rounded-xl px-4 py-2.5 text-sm transition-all"/>
            <div className="flex gap-2 flex-wrap">{[["all","Todos"],["active","Activos"],["complete","Completos"],["late","Tarde"],["ontime","Puntual"]].map(([v,l])=><button key={v} onClick={()=>setFilter(v)} className={`px-3 py-2 rounded-xl text-xs font-medium transition-all ${filter===v?"bg-cyan-500/20 border border-cyan-500/50 text-cyan-300":"bg-slate-900/70 border border-slate-700/50 text-slate-400 hover:text-white"}`}>{l}</button>)}</div>
          </div>
          <div className="card overflow-hidden">
            <div className="px-5 py-3.5 border-b border-slate-700/40 flex justify-between items-center"><p className="text-xs text-slate-500 uppercase tracking-widest">Registros</p><span className="text-xs text-slate-500">{filtRecs.length} resultados</span></div>
            {filtRecs.length===0?<p className="text-slate-500 text-sm text-center py-10">Sin registros</p>
              :<div className="overflow-x-auto scrollbar"><table className="w-full"><thead><tr className="border-b border-slate-700/30">{["Empleado","Fecha","Entrada","Salida","Horas","Puntualidad","Estado"].map(h=><th key={h} className="text-left text-xs text-slate-500 uppercase tracking-wider px-4 py-3 font-medium whitespace-nowrap">{h}</th>)}</tr></thead>
              <tbody className="divide-y divide-slate-700/20">{filtRecs.map(r=>{const u=users.find(x=>x.id===(r.user_id||r.userId));const lm=isLate(r,u);return(<tr key={r.id} className="hover:bg-slate-800/30 transition-colors">
                <td className="px-4 py-3"><div className="flex items-center gap-2"><Av name={u?.name||"?"} size="sm"/><div><p className="text-sm font-medium text-white">{u?.name}</p><p className="text-xs text-slate-500">Doc: {u?.documento}</p></div></div></td>
                <td className="px-4 py-3 text-sm text-slate-300 whitespace-nowrap">{fmtDate(r.entry)}</td>
                <td className="px-4 py-3 text-sm text-emerald-400 font-mono">{fmtShort(r.entry)}</td>
                <td className="px-4 py-3 text-sm text-amber-400 font-mono">{r.exit?fmtShort(r.exit):"—"}</td>
                <td className="px-4 py-3">{r.exit?<Badge color="blue">{diffH(r.entry,r.exit)}h</Badge>:<span className="text-slate-500 text-sm">—</span>}</td>
                <td className="px-4 py-3 whitespace-nowrap">{lm===null?<span className="text-slate-500 text-xs">—</span>:lm>0?<Badge color="red">{lm}m tarde</Badge>:<Badge color="green">{Math.abs(lm)}m antes</Badge>}</td>
                <td className="px-4 py-3">{r.exit?<Badge color="amber">Completo</Badge>:<Badge color="green">Activo</Badge>}</td>
              </tr>);})}</tbody></table></div>}
          </div>
        </>}

        {tab==="users"&&<>
          <div className="flex items-center justify-between flex-wrap gap-3">
            <h1 className="text-xl font-bold" style={{fontFamily:"'Space Grotesk'"}}>Gestión de usuarios</h1>
            <div className="flex gap-2 flex-wrap">
              <button onClick={()=>setShowImport(true)} className="bg-slate-800/70 border border-slate-600/50 hover:border-emerald-500/50 hover:text-emerald-300 text-slate-300 text-sm font-medium rounded-xl px-4 py-2 flex items-center gap-2 transition-all">📊 Importar horarios</button>
              <button onClick={()=>setShowAdd(true)} className="btn-grad text-white text-sm font-semibold rounded-xl px-4 py-2 flex items-center gap-2"><svg className="w-4 h-4" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2.5}><path strokeLinecap="round" strokeLinejoin="round" d="M12 4v16m8-8H4"/></svg>Agregar usuario</button>
            </div>
          </div>
          <div className="card overflow-hidden">
            <div className="divide-y divide-slate-700/20">{users.map(u=>(
              <div key={u.id} className="px-5 py-4 flex items-center justify-between hover:bg-slate-800/20 transition-colors">
                <div className="flex items-center gap-3"><Av name={u.name}/><div><div className="flex items-center gap-2"><p className="text-sm font-medium">{u.name}</p><Badge color={u.role==="admin"?"purple":"blue"}>{u.role==="admin"?"Admin":"Empleado"}</Badge></div><p className="text-xs text-slate-500 mt-0.5">Doc: {u.documento}</p></div></div>
                <div className="flex items-center gap-2">
                  {u.role==="employee"&&<button onClick={()=>setSchedUser(u)} className="text-xs btn-ghost rounded-lg px-3 py-1.5">🗓️ Horario</button>}
                  {u.id!==currentUser.id&&<button onClick={()=>setDelUser(u)} className="text-xs bg-slate-700/50 hover:bg-red-500/20 hover:text-red-300 border border-slate-600/40 hover:border-red-500/40 text-slate-400 rounded-lg px-3 py-1.5 transition-all">🗑️</button>}
                </div>
              </div>
            ))}</div>
          </div>
          <div className="card overflow-hidden">
            <div className="px-5 py-4 border-b border-slate-700/40 flex items-center justify-between"><p className="text-xs text-slate-500 uppercase tracking-widest">Horarios configurados</p><button onClick={()=>setShowImport(true)} className="text-xs text-cyan-400 hover:text-cyan-300 transition-colors">📊 Actualizar vía Excel →</button></div>
            <div className="divide-y divide-slate-700/20">{employees.map(u=>(
              <div key={u.id} className="px-5 py-3.5"><div className="flex items-center gap-2 mb-2"><Av name={u.name} size="sm"/><p className="text-sm font-medium">{u.name}</p></div>
              <div className="flex flex-wrap gap-1.5">{Object.entries(u.schedule||{}).map(([day,s])=><span key={day} className={`text-xs px-2 py-1 rounded-lg ${s.active?"bg-cyan-500/15 text-cyan-300 border border-cyan-500/25":"bg-slate-700/30 text-slate-500"}`}>{DAYS_LABEL[day]}{s.active?` ${s.start}`:""}</span>)}</div>
              </div>
            ))}</div>
          </div>
        </>}
      </div>
    </div>
  );
}

// ─── LOGIN ─────────────────────────────────────────────────────────────────
function Login({onLogin}){
  const [doc,setDoc]=useState(""); const [pw,setPw]=useState(""); const [err,setErr]=useState(""); const [load,setLoad]=useState(false);
  const go=async()=>{
    if(!doc||!pw){setErr("Completa todos los campos");return;}
    setLoad(true); setErr("");
    try{
      const res=await db.login(doc.trim());
      if(!res||res.length===0){setErr("Documento o contraseña incorrectos");setLoad(false);return;}
      const user=res[0];
      if(user.password!==pw){setErr("Documento o contraseña incorrectos");setLoad(false);return;}
      if(user.role!=="superadmin"){
        // Verificar si empresa está suspendida
        if(user.company_id){
          const comps=await db.getCompanies();
          const comp=comps.find(c=>c.id===user.company_id);
          if(comp&&comp.status==="suspended"){setErr("Esta empresa está suspendida. Contacta al administrador.");setLoad(false);return;}
        }
      }
      onLogin(user);
    }catch(e){setErr("Error: "+e.message);setLoad(false);}
  };
  return(
    <div className="min-h-screen bg-[#080d1a] flex items-center justify-center p-4" style={{fontFamily:"'DM Sans'"}}>
      <style>{STYLES}</style>
      <div className="fixed inset-0 pointer-events-none overflow-hidden"><div className="absolute -top-40 -right-40 w-96 h-96 rounded-full bg-cyan-500/5 blur-3xl"/><div className="absolute -bottom-40 -left-40 w-96 h-96 rounded-full bg-blue-500/5 blur-3xl"/><div className="absolute inset-0" style={{backgroundImage:"radial-gradient(circle at 1px 1px,rgba(255,255,255,.03) 1px,transparent 0)",backgroundSize:"32px 32px"}}/></div>
      <div className="relative w-full max-w-md">
        <div className="text-center mb-8"><Logo size="lg"/><h1 className="text-2xl font-bold text-white mt-4" style={{fontFamily:"'Space Grotesk'"}}>Control Asistencia Pro</h1><p className="text-slate-400 text-sm mt-1">Inicia sesión para continuar</p></div>
        <div className="card p-7" style={{boxShadow:"0 0 40px rgba(6,182,212,.1)"}}>
          <div className="space-y-4">
            <Inp label="Número de documento" type="text" value={doc} onChange={e=>setDoc(e.target.value)} placeholder="Ej: 12345678" onKeyDown={e=>e.key==="Enter"&&go()}/>
            <Inp label="Contraseña" type="password" value={pw} onChange={e=>setPw(e.target.value)} placeholder="••••••••" onKeyDown={e=>e.key==="Enter"&&go()}/>
            {err&&<div className="bg-red-500/10 border border-red-500/30 text-red-300 text-sm rounded-xl px-4 py-3">{err}</div>}
            <button onClick={go} disabled={load} className="btn-grad w-full text-white font-semibold rounded-xl py-3 text-sm disabled:opacity-60 flex items-center justify-center gap-2">{load&&<Spin/>}{load?"Verificando...":"Ingresar"}</button>
          </div>
        </div>
      </div>
    </div>
  );
}

// ─── APP ROOT ──────────────────────────────────────────────────────────────
export default function App(){
  const [users,setUsers]       = useState([]);
  const [records,setRecords]   = useState([]);
  const [currentUser,setCU]    = useState(null);
  const [appLoading,setAL]     = useState(false);
  const [actionLoading,setAcL] = useState(false);

  const loadData=async(user)=>{
    setAL(true);
    try{
      if(user.role==="superadmin"){
        // superadmin no necesita cargar datos aquí
      } else {
        const cid=user.company_id;
        const[u,r]=await Promise.all([db.getUsers(cid),db.getRecords(cid)]);
        setUsers(u||[]); setRecords(r||[]);
      }
    }catch(e){console.error(e);}
    setAL(false);
  };

  const handleLogin=async user=>{
    setCU(user);
    await loadData(user);
  };

  const onEntry=async()=>{
    setAcL(true);
    try{
      const entry=new Date().toISOString();
      const res=await db.insertRecord({user_id:currentUser.id,entry,exit:null,company_id:currentUser.company_id});
      if(res&&res[0])setRecords(p=>[res[0],...p]);
    }catch(e){console.error(e);}
    setAcL(false);
  };

  const onExit=async()=>{
    setAcL(true);
    try{
      const active=records.find(r=>(r.user_id||r.userId)===currentUser.id&&!r.exit);
      if(!active){setAcL(false);return;}
      const res=await db.updateRecord(active.id,{exit:new Date().toISOString()});
      if(res&&res[0])setRecords(p=>p.map(r=>r.id===active.id?res[0]:r));
    }catch(e){console.error(e);}
    setAcL(false);
  };

  const logout=()=>{setCU(null);setUsers([]);setRecords([]);};

  if(!currentUser) return <Login onLogin={handleLogin}/>;
  if(appLoading) return(
    <div className="min-h-screen bg-[#080d1a] flex items-center justify-center" style={{fontFamily:"'DM Sans'"}}>
      <style>{STYLES}</style>
      <div className="text-center"><Logo size="lg"/><div className="flex justify-center mt-6"><Spin/></div><p className="text-slate-400 text-sm mt-3">Cargando datos...</p></div>
    </div>
  );

  if(currentUser.role==="superadmin") return <SuperAdminPanel currentUser={currentUser} onLogout={logout}/>;
  if(currentUser.role==="admin") return <AdminPanel users={users} setUsers={setUsers} records={records} setRecords={setRecords} currentUser={currentUser} onLogout={logout}/>;
  return <EmployeePanel user={currentUser} records={records} onEntry={onEntry} onExit={onExit} actionLoading={actionLoading} onLogout={logout}/>;
}