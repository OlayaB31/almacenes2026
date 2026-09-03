import { useEffect, useMemo, useRef, useState } from 'react'
import * as XLSX from 'xlsx'
import { createStore, deleteStore, importStores, listStores, login, updateStore } from './api'
import type { Store } from './types'

const blank: Store = { codigo_wsiga:'', nit:'', nombre_almacen:'', direccion:'', centro_comercial:'', telefono_fijo:'', celular:'', whatsapp:'', correo_electronico:'', pagina_web:'', facebook:'', instagram:'', tiktok:'' }
const columns: Array<[keyof Store, string]> = [
 ['codigo_wsiga','Codigo WSIGA'],['nit','NIT'],['nombre_almacen','Nombre del almacen'],['direccion','Direccion'],['centro_comercial','Centro comercial'],['telefono_fijo','Telefono fijo'],['celular','Celular'],['whatsapp','WhatsApp'],['correo_electronico','Correo electronico'],['pagina_web','Pagina web'],['facebook','Facebook'],['instagram','Instagram'],['tiktok','TikTok']
]

function App(){
 const [authed,setAuthed]=useState(Boolean(localStorage.getItem('wsiga_token')))
 const [user,setUser]=useState(localStorage.getItem('wsiga_user')||'')
 const [rows,setRows]=useState<Store[]>([])
 const [search,setSearch]=useState('')
 const [editing,setEditing]=useState<Store|null>(null)
 const [loading,setLoading]=useState(false)
 const [notice,setNotice]=useState<{kind:'ok'|'error',text:string}|null>(null)
 const fileRef=useRef<HTMLInputElement>(null)

 async function refresh(){ setLoading(true); try{ const r=await listStores(); setRows(r.data||[])}catch(e){setNotice({kind:'error',text:(e as Error).message})}finally{setLoading(false)} }
 useEffect(()=>{ if(authed) refresh() },[authed])
 useEffect(()=>{ if(!notice) return; const t=setTimeout(()=>setNotice(null),4000); return ()=>clearTimeout(t)},[notice])

 const filtered=useMemo(()=>{
  const q=search.trim().toLowerCase(); if(!q) return rows
  return rows.filter(r=>Object.values(r).some(v=>String(v??'').toLowerCase().includes(q)))
 },[rows,search])

 function signOut(){localStorage.removeItem('wsiga_token');localStorage.removeItem('wsiga_user');setAuthed(false)}
 if(!authed) return <Login onLogged={(u)=>{setUser(u);setAuthed(true)}} />

 async function save(data:Store){
  try{ if(data.id) await updateStore(data.id,data); else await createStore(data); await refresh(); setEditing(null); setNotice({kind:'ok',text:'Registro guardado correctamente.'}) }
  catch(e){setNotice({kind:'error',text:(e as Error).message})}
 }
 async function remove(id:number){
  if(!confirm('¿Eliminar este almacen? Esta accion no se puede deshacer.')) return
  try{await deleteStore(id); await refresh(); setNotice({kind:'ok',text:'Registro eliminado.'})}catch(e){setNotice({kind:'error',text:(e as Error).message})}
 }
 function exportExcel(){
  const data=rows.map(r=>Object.fromEntries(columns.map(([k,h])=>[h,r[k]??''])))
  const ws=XLSX.utils.json_to_sheet(data); const wb=XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb,ws,'Almacenes'); XLSX.writeFile(wb,`almacenes_wsiga_${new Date().toISOString().slice(0,10)}.xlsx`)
 }
 function importExcel(ev:React.ChangeEvent<HTMLInputElement>){
  const f=ev.target.files?.[0]; if(!f) return
  const reader=new FileReader(); reader.onload=async e=>{
   try{
    const wb=XLSX.read(e.target?.result,{type:'array'}); const ws=wb.Sheets[wb.SheetNames[0]]; const json=XLSX.utils.sheet_to_json<Record<string,unknown>>(ws,{defval:''})
    const norm=(x:string)=>x.normalize('NFD').replace(/[\u0300-\u036f]/g,'').toLowerCase().replace(/[^a-z0-9]/g,''); const mapHeader=(h:string)=>columns.find(([,label])=>norm(label)===norm(h))?.[0]
    const stores=json.map(obj=>{const s={...blank}; for(const [h,v] of Object.entries(obj)){const key=mapHeader(h); if(key) s[key]=String(v??'').trim()} return s}).filter(s=>s.codigo_wsiga||s.nombre_almacen||s.nit)
    if(!stores.length) throw new Error('No se encontraron registros validos en el Excel.')
    const result=await importStores(stores); await refresh(); setNotice({kind:'ok',text:`Importacion terminada: ${result.inserted} nuevos y ${result.updated} actualizados.`})
    if(result.errors?.length) alert('Algunos registros tuvieron errores:\n\n'+result.errors.join('\n'))
   }catch(err){setNotice({kind:'error',text:(err as Error).message})}finally{if(fileRef.current) fileRef.current.value=''}
  }; reader.readAsArrayBuffer(f)
 }
 return <div className="app">
  <header className="topbar"><div><div className="brand">WSIGA</div><div className="subtitle">Gestion de almacenes</div></div><div className="top-actions"><span className="user">{user}</span><button className="ghost" onClick={signOut}>Cerrar sesion</button></div></header>
  {notice&&<div className={`notice ${notice.kind}`}>{notice.text}</div>}
  <main className="content">
   <section className="hero"><div><h1>Almacenes registrados</h1><p>Administra la informacion y manten el Excel sincronizado con la base de datos.</p></div><div className="stats"><strong>{rows.length}</strong><span>registros</span></div></section>
   <section className="toolbar"><div className="search-wrap"><span>⌕</span><input value={search} onChange={e=>setSearch(e.target.value)} placeholder="Buscar por codigo, NIT, nombre, direccion..." /></div><div className="tools"><button onClick={()=>setEditing(blank)}>+ Agregar</button><button className="secondary" onClick={()=>fileRef.current?.click()}>Importar Excel</button><button className="secondary" onClick={exportExcel}>Exportar Excel</button><input ref={fileRef} hidden type="file" accept=".xlsx,.xls,.csv" onChange={importExcel}/></div></section>
   <section className="table-card"><div className="table-scroll"><table><thead><tr><th>Codigo WSIGA</th><th>NIT</th><th>Nombre del almacen</th><th>Direccion</th><th>Centro comercial</th><th>Telefono fijo</th><th>Celular</th><th>WhatsApp</th><th>Correo</th><th>Web</th><th>Facebook</th><th>Instagram</th><th>TikTok</th><th></th></tr></thead><tbody>{loading?<tr><td colSpan={14} className="empty">Cargando...</td></tr>:filtered.length===0?<tr><td colSpan={14} className="empty">No hay registros.</td></tr>:filtered.map(r=><tr key={r.id}><td>{r.codigo_wsiga}</td><td>{r.nit}</td><td className="strong">{r.nombre_almacen}</td><td>{r.direccion}</td><td>{r.centro_comercial}</td><td>{r.telefono_fijo}</td><td>{r.celular}</td><td>{r.whatsapp}</td><td>{r.correo_electronico}</td><td>{link(r.pagina_web)}</td><td>{link(r.facebook)}</td><td>{link(r.instagram)}</td><td>{link(r.tiktok)}</td><td><div className="row-actions"><button className="icon" onClick={()=>setEditing(r)}>Editar</button><button className="icon danger" onClick={()=>remove(r.id!)}>Eliminar</button></div></td></tr>)}</tbody></table></div></section>
   <p className="hint">Mostrando {filtered.length} de {rows.length} registros.</p>
  </main>
  {editing&&<Editor initial={editing} onClose={()=>setEditing(null)} onSave={save}/>} 
 </div>
}

function link(value:string){ if(!value) return ''; let href=value; if(!/^https?:\/\//i.test(href)) href='https://'+href; return <a href={href} target="_blank" rel="noreferrer">Abrir</a> }
function Login({onLogged}:{onLogged:(u:string)=>void}){ const [u,setU]=useState(''); const [p,setP]=useState(''); const [err,setErr]=useState(''); const [busy,setBusy]=useState(false); async function go(e:React.FormEvent){e.preventDefault();setBusy(true);setErr('');try{const r=await login(u,p);localStorage.setItem('wsiga_token',r.token);localStorage.setItem('wsiga_user',r.user.username);onLogged(r.user.username)}catch(e){setErr((e as Error).message)}finally{setBusy(false)}} return <div className="login-page"><form className="login-card" onSubmit={go}><div className="brand big">WSIGA</div><p>Panel de gestion de almacenes</p><label>Usuario<input value={u} onChange={e=>setU(e.target.value)} autoComplete="username" required /></label><label>Contrasena<input value={p} onChange={e=>setP(e.target.value)} type="password" autoComplete="current-password" required /></label>{err&&<div className="form-error">{err}</div>}<button disabled={busy}>{busy?'Ingresando...':'Ingresar'}</button></form></div> }
function Editor({initial,onClose,onSave}:{initial:Store,onClose:()=>void,onSave:(x:Store)=>Promise<void>}){ const [form,setForm]=useState<Store>(initial); const [busy,setBusy]=useState(false); const set=(k:keyof Store,v:string)=>setForm(s=>({...s,[k]:v})); async function submit(e:React.FormEvent){e.preventDefault();setBusy(true);try{await onSave(form)}finally{setBusy(false)}} return <div className="modal-back"><form className="modal" onSubmit={submit}><div className="modal-head"><div><h2>{form.id?'Editar almacen':'Nuevo almacen'}</h2><span>Completa la informacion disponible.</span></div><button type="button" className="close" onClick={onClose}>×</button></div><div className="grid">{columns.map(([k,label])=><label key={k}>{label}<input value={String(form[k]??'')} onChange={e=>set(k,e.target.value)} /></label>)}</div><div className="modal-actions"><button type="button" className="secondary" onClick={onClose}>Cancelar</button><button disabled={busy}>{busy?'Guardando...':'Guardar'}</button></div></form></div> }
