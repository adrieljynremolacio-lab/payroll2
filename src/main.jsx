
import React,{useMemo,useState} from 'react';
import {createRoot} from 'react-dom/client';
import * as XLSX from 'xlsx';
import {Users,Calculator,Upload,Download,Plus,Search,Settings,FileSpreadsheet,Trash2,X,Cloud,RefreshCw} from 'lucide-react';
import './styles.css';

const KEY='PH_PAYROLL_V3';
const DEFAULT_POS=['Foreman','Welder','Erector','Helper','Safety Officer','Engineer','Secretary'];
const INITIAL={employees:[],rows:[],positions:DEFAULT_POS,settings:{company:'My Company',defaultRate:100,googleScriptUrl:''}};
const load=()=>{
  try{
    const saved=JSON.parse(localStorage.getItem(KEY)||'null');
    if(!saved) return structuredClone(INITIAL);
    return {
      employees:Array.isArray(saved.employees)?saved.employees:[],
      rows:Array.isArray(saved.rows)?saved.rows:[],
      positions:Array.isArray(saved.positions)&&saved.positions.length?saved.positions:DEFAULT_POS,
      settings:{...INITIAL.settings,...(saved.settings||{})}
    };
  }catch{return structuredClone(INITIAL)}
};
const money=n=>'₱'+Number(n||0).toLocaleString('en-PH',{minimumFractionDigits:2,maximumFractionDigits:2});
const uid=p=>p+Date.now().toString(36).toUpperCase()+Math.random().toString(36).slice(2,5).toUpperCase();
const H=['Employee Name','Position','Pay Rate','P','P','F','S','S','M','T','W','TH','Thours','ADD','Salary','Rose Cntn','Rose Cntn To Thur','Over','PP2 NO.25','April','Uniform','Liza Cntn','Net pay'];
const HOURS=['P2','F','S1','S2','M','T','W','TH'];
const DEDS=['Rose','RoseThur','Over','PP2','April','Uniform','Liza'];

function App(){
 const [db,setDb]=useState(load),[tab,setTab]=useState('payroll'),[period,setPeriod]=useState('August 2026'),[q,setQ]=useState('');
 const [modal,setModal]=useState(null),[msg,setMsg]=useState('');
 const update=fn=>setDb(p=>{const n=structuredClone(p);fn(n);localStorage.setItem(KEY,JSON.stringify(n));return n});
 const notify=s=>{setMsg(s);setTimeout(()=>setMsg(''),2600)};
 const employees=db.employees.filter(e=>e.active!==false);
 const row=e=>{const r=db.rows.find(x=>x.employeeId===e.id&&x.period===period)||{};const hours=HOURS.reduce((a,k)=>a+Number(r[k]||0),0);const salary=hours*Number(e.rate||0);const add=Number(r.ADD||0);const deductions=DEDS.reduce((a,k)=>a+Number(r[k]||0),0);return {...r,Thours:hours,Salary:salary,Net:salary+add-deductions}};
 const ensure=()=>update(d=>{d.rows=d.rows.filter(r=>r.period!==period||d.employees.some(e=>e.id===r.employeeId));d.employees.forEach(e=>{if(!d.rows.find(r=>r.employeeId===e.id&&r.period===period))d.rows.push({employeeId:e.id,period,P1:e.position,P2:0,F:0,S1:0,S2:0,M:0,T:0,W:0,TH:0,ADD:0,Rose:0,RoseThur:0,Over:0,PP2:0,April:0,Uniform:0,Liza:0})});});notify('Payroll rows prepared');
 const cell=(id,k,v)=>update(d=>{let r=d.rows.find(x=>x.employeeId===id&&x.period===period);if(!r){r={employeeId:id,period};d.rows.push(r)}r[k]=k==='P1'?v:Number(v)||0});
 const addEmp=f=>{update(d=>d.employees.push({id:uid('EMP-'),name:f.name,position:f.position,rate:Number(f.rate)||0,active:true}));setModal(null);notify('Employee enrolled')};
 const addPos=p=>{if(!p||db.positions.includes(p))return;update(d=>d.positions.push(p));notify('Position added')};
 const importCSV=text=>{const lines=text.trim().split(/\r?\n/);const parse=s=>{let a=[],x='',q=false;for(const c of s){if(c==='"'){q=!q}else if(c===','&&!q){a.push(x.trim());x=''}else x+=c}a.push(x.trim());return a};const h=parse(lines[0]).map(x=>x.toLowerCase());const idx=n=>h.findIndex(x=>n.includes(x));const ni=idx(['employee name','name','employee']),pi=idx(['position','pos']),ri=idx(['pay rate','rate','hourly rate','pay rate per hour']);if(ni<0||pi<0){notify('CSV must contain Employee Name and Position');return}const a=lines.slice(1).filter(Boolean).map(l=>{const v=parse(l);return{id:uid('EMP-'),name:v[ni]||'',position:v[pi]||'',rate:Number((v[ri]||'0').replace(/[₱,]/g,''))||0,active:true}});update(d=>d.employees.push(...a));setModal(null);notify(`${a.length} employees imported`)};
 const exportExcel=()=>{const wb=XLSX.utils.book_new();const data=[H,...employees.map(e=>{const r=row(e);return[e.name,e.position,e.rate,r.P1||e.position,r.P2||0,r.F||0,r.S1||0,r.S2||0,r.M||0,r.T||0,r.W||0,r.TH||0,null,r.ADD||0,null,r.Rose||0,r.RoseThur||0,r.Over||0,r.PP2||0,r.April||0,r.Uniform||0,r.Liza||0,null]})];const ws=XLSX.utils.aoa_to_sheet(data);employees.forEach((e,i)=>{const n=i+2;ws[`M${n}`]={t:'n',f:`SUM(E${n}:L${n})`};ws[`O${n}`]={t:'n',f:`M${n}*C${n}`};ws[`W${n}`]={t:'n',f:`O${n}+N${n}-SUM(P${n}:V${n})`}});XLSX.utils.book_append_sheet(wb,ws,'Payroll');const guide=XLSX.utils.aoa_to_sheet([['Formula','Excel formula'],['Thours','=SUM(E2:L2)'],['Salary','=M2*C2'],['Net pay','=O2+N2-SUM(P2:V2)'],['Note','P = Position; second P = paid/store-hours field; Liza Canteen replaces Pepito Canteen.']]);XLSX.utils.book_append_sheet(wb,guide,'Formula Guide');XLSX.writeFile(wb,`Payroll-${period.replaceAll(' ','-')}.xlsx`);notify('Excel downloaded')};
 const sync=async()=>{if(!db.settings.googleScriptUrl){notify('Add Apps Script URL in Settings');return}try{await fetch(db.settings.googleScriptUrl,{method:'POST',headers:{'Content-Type':'text/plain;charset=utf-8'},body:JSON.stringify({action:'replaceAll',data:db})});notify('Google Sheets synced')}catch(e){notify('Sync failed: check Apps Script URL and deployment')}};
 return <div className="app"><aside><div className="logo">PAYROLL<span>+</span><small>PH Hourly Payroll</small></div>
  <button className={tab==='payroll'?'nav on':'nav'} onClick={()=>setTab('payroll')}><Calculator/>Payroll</button>
  <button className={tab==='employees'?'nav on':'nav'} onClick={()=>setTab('employees')}><Users/>Employees</button>
  <button className="nav" onClick={()=>setModal('csv')}><Upload/>CSV Enrollment</button>
  <button className="nav" onClick={exportExcel}><FileSpreadsheet/>Excel / Sheets</button>
  <button className="nav" onClick={()=>setModal('settings')}><Settings/>Settings</button>
  <button className="nav" onClick={sync}><Cloud/>Sync Google Sheet</button>
  <div className="side">Standard positions are editable. Tax is OFF. PhilHealth and Pag-IBIG are not included in this payroll layout.</div></aside>
  <main><header><div><h1>{tab==='payroll'?'Payroll':'Employees'}</h1><p>{db.settings.company}</p></div><div className="top"><button className="primary" onClick={()=>setModal('employee')}><Plus/>Enroll Employee</button><button onClick={exportExcel}><Download/>Export</button></div></header>{msg&&<div className="toast">{msg}</div>}
   {tab==='payroll'?<Payroll employees={employees} row={row} cell={cell} period={period} setPeriod={setPeriod} ensure={ensure}/>:<Employees db={db} employees={employees.filter(e=>(e.name+' '+e.position).toLowerCase().includes(q.toLowerCase()))} q={q} setQ={setQ} update={update} addPos={addPos}/>}
  </main>
  {modal==='employee'&&<EmployeeModal positions={db.positions} close={()=>setModal(null)} save={addEmp} addPos={addPos}/>}
  {modal==='csv'&&<CSVModal close={()=>setModal(null)} importCSV={importCSV}/>}
  {modal==='settings'&&<SettingsModal db={db} update={update} close={()=>setModal(null)}/>}
 </div>
}

function Payroll({employees,row,cell,period,setPeriod,ensure}){return <section><div className="bar"><div><label>Payroll Period</label><input value={period} onChange={e=>setPeriod(e.target.value)}/></div><button className="primary" onClick={ensure}><Calculator/>Prepare Payroll</button></div><div className="hint"><b>Formula:</b> Thours = P + F + S + S + M + T + W + TH. Salary = Thours × Pay Rate. Net Pay = Salary + ADD − all deduction columns. The hourly rate is manually set per employee.</div><div className="sheet"><table><thead><tr>{H.map((x,i)=><th className={['Thours','Salary','Net pay'].includes(x)?'calcHead':''} key={i}>{x}</th>)}</tr></thead><tbody>{employees.map(e=>{const r=row(e);return <tr key={e.id}><td className="name">{e.name}</td><td><select value={r.P1||e.position} onChange={x=>cell(e.id,'P1',x.target.value)}><option>{e.position}</option></select></td><td className="rate">{money(e.rate)}/hr</td>{HOURS.map(k=><td key={k}><input type="number" value={r[k]||0} onChange={x=>cell(e.id,k,x.target.value)}/></td>)}<td className="calc">{r.Thours.toFixed(1)}</td><td><input type="number" value={r.ADD||0} onChange={x=>cell(e.id,'ADD',x.target.value)}/></td><td className="calc">{money(r.Salary)}</td>{DEDS.map(k=><td key={k}><input type="number" value={r[k]||0} onChange={x=>cell(e.id,k,x.target.value)}/></td>)}<td className="net">{money(r.Net)}</td></tr>})}</tbody></table></div></section>}

function Employees({db,employees,q,setQ,update,addPos}){return <section><div className="bar"><div className="search"><Search/><input placeholder="Search employee..." value={q} onChange={e=>setQ(e.target.value)}/></div><button className="primary" onClick={()=>{const p=prompt('New position name:');if(p)addPos(p.trim())}}><Plus/>Add Position</button></div><div className="cards">{employees.map(e=><div className="emp" key={e.id}><small>{e.id}</small><h3>{e.name}</h3><p>{e.position}</p><label>Pay Rate / Hour<input type="number" value={e.rate} onChange={x=>update(d=>{d.employees.find(a=>a.id===e.id).rate=Number(x.target.value)||0;return d})}/></label><button className="danger" onClick={()=>update(d=>{d.employees=d.employees.filter(a=>a.id!==e.id);d.rows=d.rows.filter(r=>r.employeeId!==e.id);return d})}><Trash2/> Remove</button></div>)}</div></section>}

function EmployeeModal({positions,close,save,addPos}){const [f,setF]=useState({name:'',position:positions[0]||'',rate:''});const [newP,setNewP]=useState('');return <Modal title="Enroll Employee" close={close}><div className="form"><label>Employee Name<input autoFocus value={f.name} onChange={e=>setF({...f,name:e.target.value})}/></label><label>Position<select value={f.position} onChange={e=>{if(e.target.value==='__NEW__'){const p=prompt('Type the new position:');if(p){addPos(p.trim());setF({...f,position:p.trim()})}}else setF({...f,position:e.target.value})}}>{positions.map(p=><option key={p}>{p}</option>)}<option value="__NEW__">+ Add New Position...</option></select></label><label>Pay Rate per Hour<input type="number" value={f.rate} onChange={e=>setF({...f,rate:e.target.value})}/></label></div><p className="hint">You can change the hourly rate later from Employees.</p><div className="modalbuttons"><button onClick={close}>Cancel</button><button className="primary" disabled={!f.name||!f.position} onClick={()=>save(f)}>Save Employee</button></div></Modal>}
function CSVModal({close,importCSV}){const [text,setText]=useState('');return <Modal title="CSV Employee Enrollment" close={close}><p>Upload a CSV containing <b>Employee Name, Position, Pay Rate</b>.</p><input type="file" accept=".csv,text/csv" onChange={e=>{const f=e.target.files?.[0];if(f){const r=new FileReader();r.onload=()=>setText(r.result);r.readAsText(f)}}}/><textarea value={text} onChange={e=>setText(e.target.value)} placeholder="Employee Name,Position,Pay Rate&#10;Jimmy Valdrez,Foreman,100"/><div className="modalbuttons"><button onClick={close}>Cancel</button><button className="primary" disabled={!text} onClick={()=>importCSV(text)}>Import Employees</button></div></Modal>}
function SettingsModal({db,update,close}){
  const [company,setCompany]=useState(db.settings.company||'');
  const [rate,setRate]=useState(String(db.settings.defaultRate??100));
  const [url,setUrl]=useState(db.settings.googleScriptUrl||'');
  const save=()=>{
    update(d=>{
      d.settings={...d.settings,company,defaultRate:Number(rate)||0,googleScriptUrl:url.trim()};
      return d;
    });
    close();
  };
  return <Modal title="Settings" close={close}>
    <div className="form">
      <label>Company Name<input value={company} onChange={e=>setCompany(e.target.value)}/></label>
      <label>Default Hourly Rate<input type="number" value={rate} onChange={e=>setRate(e.target.value)}/></label>
      <label>Google Apps Script Web App URL<input value={url} onChange={e=>setUrl(e.target.value)} placeholder="https://script.google.com/macros/s/.../exec"/></label>
    </div>
    <p className="hint">Google Sheets sync uses the Apps Script included in this project. Deploy it as a Web App, then paste the URL here.</p>
    <div className="modalbuttons"><button onClick={close}>Cancel</button><button className="primary" onClick={save}>Save Settings</button></div>
  </Modal>
}
function Modal({title,close,children}){return <div className="overlay"><div className="modal"><div className="modalhead"><h2>{title}</h2><button onClick={close}><X/></button></div>{children}</div></div>}
createRoot(document.getElementById('root')).render(<App/>);
