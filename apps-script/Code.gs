/*
GOOGLE SHEETS BACKEND
1. Create a Google Sheet.
2. Extensions > Apps Script.
3. Paste this file.
4. Deploy > New deployment > Web app.
5. Execute as: Me. Access: Anyone with the link (or your organization's choice).
6. Copy the /exec URL into Payroll > Settings > Google Apps Script Web App URL.
*/
const SHEETS=['Employees','Rows','Positions','Settings'];
function setup_(){
 const ss=SpreadsheetApp.getActive();
 const headers={
  Employees:['id','name','position','rate','active'],
  Rows:['employeeId','period','P1','P2','F','S1','S2','M','T','W','TH','ADD','Rose','RoseThur','Over','PP2','April','Uniform','Liza'],
  Positions:['position'],
  Settings:['key','value']
 };
 Object.keys(headers).forEach(n=>{let s=ss.getSheetByName(n)||ss.insertSheet(n);if(s.getLastRow()===0)s.appendRow(headers[n]);});
}
function doGet(){setup_();return out_({ok:true,message:'Payroll Google Sheet API online'});}
function doPost(e){
 try{
  setup_(); const b=JSON.parse(e.postData.contents||'{}');
  if(b.action==='replaceAll'){replace_(b.data);return out_({ok:true});}
  if(b.action==='read'){return out_(read_());}
  return out_({ok:false,error:'Unknown action'});
 }catch(err){return out_({ok:false,error:String(err)});}
}
function replace_(d){
 const ss=SpreadsheetApp.getActive();
 const maps={
  Employees:[['id','name','position','rate','active'],...(d.employees||[]).map(x=>[x.id,x.name,x.position,x.rate,x.active])],
  Rows:[['employeeId','period','P1','P2','F','S1','S2','M','T','W','TH','ADD','Rose','RoseThur','Over','PP2','April','Uniform','Liza'],...(d.rows||[]).map(x=>['employeeId','period','P1','P2','F','S1','S2','M','T','W','TH','ADD','Rose','RoseThur','Over','PP2','April','Uniform','Liza'].map(k=>x[k]??''))],
  Positions:[['position'],...(d.positions||[]).map(x=>[x])],
  Settings:[['key','value'],...Object.entries(d.settings||{}).map(([k,v])=>[k,String(v)])]
 };
 Object.entries(maps).forEach(([n,vals])=>{let s=ss.getSheetByName(n);s.clearContents();s.getRange(1,1,vals.length,vals[0].length).setValues(vals);});
}
function read_(){
 const ss=SpreadsheetApp.getActive();
 const out={};SHEETS.forEach(n=>{let v=ss.getSheetByName(n).getDataRange().getValues();out[n]=v;});return out;
}
function out_(x){return ContentService.createTextOutput(JSON.stringify(x)).setMimeType(ContentService.MimeType.JSON);}
