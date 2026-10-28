(function(){
"use strict";
const $=id=>document.getElementById(id);

/* ---------- mapa de columnas (referencia visible) ---------- */
const MAP=[["Q","Tratamiento"],["R","Primer apellido"],["S","Segundo apellido"],["T","Nombres"],["U","Ciudad de nacimiento"],["V","Fecha de nacimiento"],["W","Estado civil"],["X","Nacionalidad"],["Y","Cédula"],["Z","PRODUBANCO"],["AA","Cuenta Produbanco"],["AB","Tipo de cuenta"],["AC","Dirección"],["AE","Ciudad donde reside"],["AF","Teléfono convencional"],["AG","Provincia (por ciudad)"],["AH","Inicio de estudios"],["AI","ECUADOR"],["AJ","Título obtenido"],["AK","Nivel de estudios"],["AL","Institución"],["AM","Título según nivel y fecha fin"],["AN","Correo"],["AO","Celular"],["AP–AY","Carga 1: tipo, apellidos, nombres, fecha, nacionalidad, ciudad, cédula"],["BC–BL","Carga 2: mismos campos"],["AT / BG","Sexo de la carga (lo eliges aquí)"]];
$("map").innerHTML=MAP.map(([c,t])=>`<div><code>${c}</code><span>${t}</span></div>`).join("");

/* ---------- utilidades ---------- */
const BASURA=new Set(["NA","ND","N","NO","NOTENGO","NOAPLICA","NINGUNO","NINGUNA","0","00","000","0000",""]);
const NAC={ECUADOR:"ECUADOR",COLOMBIA:"COLOMBIA",VENEZUELA:"VENEZUELA",VENUEZUELA:"VENEZUELA",CHILE:"CHILE",OTRA:"OTRA"};
const ECIVIL={"Divorciada/o":"Divorciada/a"};
const PROV=[["GUAYAQUIL","GUAYAS"],["CUENCA","AZUAY"],["AMBATO","TUNGURAHUA"],["LATACUNGA","COTOPAXI"],["IBARRA","IMBABURA"],["RIOBAMBA","CHIMBORAZO"],["LOJA","LOJA"],["ESMERALDAS","ESMERALDAS"],["SANTO DOMINGO","SANTO DOMINGO DE LOS TSÁCHILAS"],["MANTA","MANABI"],["PORTOVIEJO","MANABI"]];
const DATE_COLS=new Set(["V","AH","AU","BH"]);
const ID_COLS=new Set(["Y","AY","BL"]);
const esc=s=>String(s).replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g,"").replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;");
const unesc=s=>s.replace(/&lt;/g,"<").replace(/&gt;/g,">").replace(/&quot;/g,'"').replace(/&apos;/g,"'").replace(/&amp;/g,"&");
function clean(v){if(v==null)return null;if(typeof v==="number")return v;const s=String(v).replace(/\s+/g," ").trim();return s||null}
function util(v){if(v==null)return null;if(typeof v==="number")return v===0?null:v;const s=String(v);const k=s.toUpperCase().replace(/[^A-Z0-9]/g,"");if(BASURA.has(k)||s.startsWith("1/1/0001"))return null;return v}
function serialFromYMD(y,m,d){return Math.round(Date.UTC(y,m-1,d)/864e5)+25569}
function toSerial(v){
  if(v==null)return null;
  if(typeof v==="number")return v>=1?Math.floor(v):null;
  const m=String(v).match(/^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{4})$/);
  if(m&&+m[3]>1900)return serialFromYMD(+m[3],+m[2],+m[1]);
  return null;
}
const HOY=Math.floor((Date.now()-new Date().getTimezoneOffset()*6e4)/864e5)+25569;
function titulo(clase,fin){const s=toSerial(fin);const c=(s&&s<=HOY)?"Completa":"Incompleta";return {"Primaria":"Primaria "+c,"Secundaria":"Secundaria "+c,"Tecnica y Superior":"Técnica "+c,"Universitaria":"Universitaria "+c,"Post Grado":"Maestría "+c}[clase]||null}
function provincia(c){c=String(c||"").toUpperCase();for(const [k,p] of PROV)if(c.includes(k))return p;return "PICHINCHA"}
function splitAp(ap){if(!ap)return [null,null];const s=String(ap);const i=s.indexOf(" ");return i<0?[s,null]:[s.slice(0,i),s.slice(i+1)]}
function idStr(v){if(v==null)return null;let s=typeof v==="number"?String(Math.round(v)):String(v);if(/^\d{9}$/.test(s))s="0"+s;return s}
// celular: Excel borra el 0 inicial (995026435 -> 0995026435); +593 9... -> 09...
function cel(v){if(v==null)return null;const d=String(typeof v==="number"?Math.round(v):v).replace(/\D/g,"");if(/^9\d{8}$/.test(d))return "0"+d;if(/^5939\d{8}$/.test(d))return "0"+d.slice(3);return v}
const colIdx=c=>[...c].reduce((a,ch)=>a*26+ch.charCodeAt(0)-64,0);

// Sexo a partir del nombre (nombres comunes en Ecuador). Devuelve "VARON", "MUJER" o null si no es seguro.
const SEXO_M=new Set("aaron abel abraham adam adan adolfo adrian agustin alan alberto aldo alejandro alessandro alex alexander alexis alfonso alfredo alonso alvaro anderson andre andres andrux angel anibal antonio ariel armando arturo augusto axel bayron benjamin bernardo braulio brayan bruno bryan byron camilo carlos cesar christian christopher cristhian cristian cristobal damian daniel dante dario darlin darwin david dennis derick diego dilan dylan edgar edison eduardo edwin efrain elias emanuel emiliano emilio enrique erick ernesto esteban estefano esteffano fabian fabricio facundo felipe fernando francisco franco frank fredy gabriel gael gerardo german gonzalo guillermo gustavo hector henry henrry hernan hugo ian ignacio iker isaac isaias ismael ithan ivan jack jaime jairo javier jean jefferson jeremy jesus jhon jhonny jhordy jimmy joaquin joel johan john jonathan jordan jorge jose joseph josue joshet jostin juan julian julio junior justin kevin kelvin leandro leonardo liam lucas luciano luis manuel marco marcos mario marlon martin mateo matias mauricio maximiliano mayer miguel milton nelson nicolas noah omar oscar pablo patricio paul pedro rafael ramiro raul ricardo roberto robinson rodrigo rolando ruben salvatore samuel santiago sebastian sergio steven steveen stalin andi teo thiago thomas tiago tomas ulises valentin victor vinicio walter wilmer william wilson wiston xavier yandri zabdiel".split(" "));
const SEXO_F=new Set("abigail adriana aime aitana alejandra alexandra alicia alison alisson allyson amanda amy ambar ana anabel anahi anai rafaella andrea angela angie antonella ariana aylen beatriz belen blanca brenda camila carla carmen carolina catalina chloe cristina dana danna daniela dayana denisse diana domenica dominique elaine elena elizabeth emilia emily erika estefania estefany evelyn fernanda gabriela genesis geraldin gisela guadalupe heydy ines isabel isabella ivanna ivonne jazmin jeanely jennifer jessica johanna julia juliana karen karla katherine keyla laura layla leidy leishy leydi lia liz lizbeth lizeth lorena loren lucia luz madeline maria mariagny mariana marilin martina mayerly meredith micaela melanie melany melisa michelle milena mishell monica nicole noemi paola patricia paula priscila rafaela raquel rebeca rianna romina rosa rubi ruth salome samara sara sarai sarahi scarleth sofia sol susana tatiana valentina valeria vanessa veronica victoria viviana ximena yadira yessenia yomayra yudrei zoe".split(" "));
function sinTilde(s){return s.normalize("NFD").replace(/[̀-ͯ]/g,"").toLowerCase()}
function sexoNombre(nombre){
  if(!nombre)return null;
  for(const w of sinTilde(String(nombre)).split(/[^a-zñ]+/).filter(Boolean).slice(0,2)){
    if(SEXO_M.has(w))return "VARON";
    if(SEXO_F.has(w))return "MUJER";
  }
  const w=sinTilde(String(nombre)).split(/[^a-zñ]+/).filter(Boolean)[0]||"";
  if(/a$/.test(w)&&!/(luca|joshua|elias)$/.test(w))return "MUJER";
  if(/o$/.test(w))return "VARON";
  return null;
}

// cédula ecuatoriana: dígito verificador (módulo 10)
function cedOk(s){s=String(s||"");if(!/^\d{10}$/.test(s))return false;let t=0;for(let i=0;i<9;i++){let p=+s[i]*(i%2?1:2);if(p>9)p-=9;t+=p}return (10-t%10)%10===+s[9]}
// ciudad: variantes mal escritas de Quito / Sangolquí
function ciudad(v){if(v==null)return null;const k=String(v).normalize("NFD").replace(/[\u0300-\u036f]/g,"").toUpperCase();if(/\bQUI[TR]O\b/.test(k))return "Quito";if(/SANGOLQUI/.test(k))return "Sangolquí";return v}
// nombres: dígitos pegados a letras por error de teclado (CASTILL9 -> CASTILLO)
function nombreFix(v){if(v==null)return v;return String(v).replace(/([A-Za-zÁÉÍÓÚÑáéíóúñ])[09](?=[A-Za-zÁÉÍÓÚÑáéíóúñ]|\b)/g,(m,a)=>a+(a===a.toUpperCase()?"O":"o"))}
/* ---------- estado ---------- */
let BASE=null, FMT=null, MODEL=null;
const edits={};

/* ---------- lectura base ---------- */
async function readBase(file){
  const wb=XLSX.read(await file.arrayBuffer(),{type:"array",cellDates:false});
  let rows=null,hdr=-1,fila0=1;
  for(const n of wb.SheetNames){
    const ws=wb.Sheets[n];if(!ws["!ref"])continue;
    const rr=XLSX.utils.sheet_to_json(ws,{header:1,raw:true,defval:null,blankrows:true});
    const h=rr.findIndex(r=>r&&String(r[4]||"").startsWith("Ingresa tus"));
    if(h>=0){rows=rr;hdr=h;fila0=XLSX.utils.decode_range(ws["!ref"]).s.r+1;break}
  }
  if(hdr<0)throw new Error("No encuentro la columna «Ingresa tus dos Nombres». ¿Es la base del Link de ingreso?");
  const H=rows[hdr].map(x=>String(x||"").trim().toLowerCase());
  const ESPERA=[[4,"ingresa tus dos nombres"],[5,"ingresa tu primer apellido"],[11,"número de cédula"],[12,"número de cta"],[13,"tipo de cta"],[33,"estudios"],[39,"correo"],[40,"en el caso de tener cónyuge, ingrese el apellido"]];
  const malas=ESPERA.filter(([i,t])=>!(H[i]||"").startsWith(t));
  if(malas.length)throw new Error("La base cambió de estructura: no encuentro «"+malas[0][1]+"» en la columna "+XLSX.utils.encode_col(malas[0][0])+". Avísale a People Analytics.");
  const TIPOS=["Cuenta de Ahorros","Cuenta Corriente"],NIV=["Primaria","Secundaria","Tecnica y Superior","Universitaria","Post Grado"];
  const data=[];
  rows.forEach((r,i)=>{
    if(i<=hdr||!r||!r[4]||String(r[4]).startsWith("Ingresa tus"))return;
    let a=[];for(let j=0;j<60;j++)a[j]=clean(r[j]);
    // fila corrida: buscar cuántas columnas se movieron desde la M (cuenta) y volver a alinear
    let corr=0;
    if(!TIPOS.includes(a[13])){
      for(let k=1;k<=4;k++){
        if(TIPOS.includes(a[13+k])&&(a[33+k]==null||NIV.includes(a[33+k]))&&(a[39+k]==null||String(a[39+k]).includes("@"))){corr=k;break}
      }
      if(corr)a=a.slice(0,12).concat(a.slice(12+corr));
    }
    a.filaBase=fila0+i;a.corrida=corr;
    data.push(a);
  });
  if(!data.length)throw new Error("La base no tiene colaboradores.");
  const sinPos=data.filter(r=>!r[0]).length;
  if(sinPos===data.length)throw new Error("La columna POSICION (A) está vacía. Abre la base en Excel, deja que calcule las posiciones, guarda y vuelve a subirla.");
  return {data,sinPos,name:file.name};
}

/* ---------- lectura formato ---------- */
function cellText(xml,ref,ss){
  const m=xml.match(new RegExp('<c r="'+ref+'"([^>]*?)(?:/>|>([\\s\\S]*?)</c>)'));
  if(!m||!m[2])return null;
  const attrs=m[1],inner=m[2];
  if(/t="inlineStr"/.test(attrs)){return unesc((inner.match(/<t[^>]*>([\s\S]*?)<\/t>/)||[,""])[1])}
  const v=(inner.match(/<v>([\s\S]*?)<\/v>/)||[,null])[1];
  if(v==null)return null;
  return /t="s"/.test(attrs)?ss[+v]:unesc(v);
}
async function readFmt(file){
  const buf=await file.arrayBuffer();const zip=await JSZip.loadAsync(buf.slice(0));
  const wbx=await zip.file("xl/workbook.xml").async("string");
  const rels=await zip.file("xl/_rels/workbook.xml.rels").async("string");
  const ssf=zip.file("xl/sharedStrings.xml");
  const ss=ssf?[...(await ssf.async("string")).matchAll(/<si>([\s\S]*?)<\/si>/g)].map(m=>unesc([...m[1].replace(/<rPh[\s\S]*?<\/rPh>/g,"").matchAll(/<t[^>]*>([\s\S]*?)<\/t>/g)].map(x=>x[1]).join(""))):[];
  const sheets=[...wbx.matchAll(/<sheet\b([^>]*)\/>/g)].map(m=>{
    const a=m[1];const name=unesc((a.match(/name="([^"]*)"/)||[])[1]||"");const rid=(a.match(/r:id="([^"]*)"/)||[])[1];
    const rel=[...rels.matchAll(/<Relationship\b([^>]*)\/>/g)].map(x=>x[1]).find(x=>x.includes('Id="'+rid+'"'));
    let t=(rel.match(/Target="([^"]*)"/)||[])[1];t=t.startsWith("/")?t.slice(1):"xl/"+t.replace(/^\.\//,"");
    return {name,path:t,hidden:/state="(hidden|veryHidden)"/.test(a)};
  });
  let found=null;
  for(const s of sheets.sort((a,b)=>a.hidden-b.hidden)){
    const f=zip.file(s.path);if(!f)continue;
    const xml=await f.async("string");
    const m4=(cellText(xml,"M4",ss)||"").trim().toUpperCase(), r4=(cellText(xml,"R4",ss)||"").trim().toUpperCase();
    if(m4==="POSICION"&&r4==="APELLIDO"){found={...s,xml};break}
  }
  if(!found)throw new Error("No encuentro la hoja con POSICION en M4 y APELLIDO en R4. ¿Es el formato de Medidas?");
  const filaPos={};
  for(const m of found.xml.matchAll(/<c r="M(\d+)"([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g)){
    const r=+m[1];if(r<5||!m[3])continue;
    const v=cellText(m[0],"M"+r,ss);if(v!=null&&/^\d+$/.test(String(v).trim()))filaPos[String(v).trim()]=r;
  }
  return {buf,sheet:found,filaPos,wbx,name:file.name};
}

/* ---------- construir modelo ---------- */
function build(){
  const rep=[],fill={},pend=[];let ok=0,bad=0;
  const nom=b=>[b[4],b[5],b[6]].filter(Boolean).join(" ");
  const porPos={};
  const ced=b=>String(b[11]??"").replace(/\D/g,"");
  // misma persona que respondió 2 veces: se usa su última respuesta (timestamp)
  const ultima={};
  for(const b of BASE.data){const k=ced(b);if(!k)continue;const t=typeof b[1]==="number"?b[1]:0;if(!ultima[k]||t>=ultima[k].t)ultima[k]={t,b}}
  const repetidas=BASE.data.filter(b=>{const k=ced(b);return k&&ultima[k].b!==b});
  const usar=BASE.data.filter(b=>!repetidas.includes(b));
  for(const b of usar){const p=b[0]!=null?String(b[0]).trim():null;if(p)(porPos[p]=porPos[p]||[]).push(nom(b))}
  for(const b of repetidas)rep.push({pos:b[0]!=null?String(b[0]):"—",nombre:nom(b),fb:b.filaBase,fila:"",estado:"warn",obs:"Respondió el formulario más de una vez: se usó su última respuesta"});
  const NIVELES=new Set(["Primaria","Secundaria","Tecnica y Superior","Universitaria","Post Grado"]);
  for(const b of usar){
    const pos=b[0]!=null?String(b[0]).trim():null;
    const nombre=nom(b);
    if(!pos||!/^\d+$/.test(pos)){bad++;rep.push({pos:pos||"—",nombre,fb:b.filaBase,fila:"",estado:"bad",obs:pos?"La posición en la base no es válida («"+pos+"»): revisa la fórmula de POSICION":"Sin posición en la base"});continue}
    if(!(pos in FMT.filaPos)){bad++;rep.push({pos,nombre,fb:b.filaBase,fila:"",estado:"bad",obs:"Esta posición no existe en el formato"});continue}
    if(porPos[pos].length>1){bad++;rep.push({pos,nombre,fb:b.filaBase,fila:FMT.filaPos[pos],estado:"bad",obs:"Posición repetida en la base ("+porPos[pos].join(" / ")+"). No se cargó ninguno: corrige la posición en la base."});continue}
    const corr=[];
    if(!["Cuenta de Ahorros","Cuenta Corriente"].includes(b[13]))corr.push("tipo de cuenta");
    if(b[39]&&!String(b[39]).includes("@"))corr.push("correo");
    if(b[33]&&!NIVELES.has(b[33]))corr.push("nivel de estudios");
    if(b[12]!=null&&b[11]!=null&&String(b[12])===String(b[11]))corr.push("cuenta igual a la cédula");
    if(toSerial(b[8])==null)corr.push("fecha de nacimiento");
    if(corr.length>=1&&(corr.length>=2||corr.includes("cuenta igual a la cédula"))){bad++;rep.push({pos,nombre,fb:b.filaBase,fila:FMT.filaPos[pos],estado:"bad",obs:"Datos corridos de columna en la base ("+corr.join(", ")+"). No se cargó: revisa esta fila en la base."});continue}
    const r=FMT.filaPos[pos],obs=[];
    for(const i of [4,5,6]){const f=nombreFix(b[i]);if(f!==b[i]){obs.push(`Nombre «${b[i]}» corregido a «${f}»`);b[i]=f}}
    const resi=ciudad(util(b[14]));let resiF=resi;
    if(resi&&!/^(Quito|Sangolquí)$/.test(resi)&&/^(ECUADOR|PICHINCHA)$/i.test(String(resi).trim())){resiF="Quito";obs.push(`Ciudad de residencia «${resi}»: se puso Quito`)}
    const mail=b[39]!=null?String(b[39]).replace(/\s+/g,""):null;
    if(mail&&!/^[^@]+@[^@]+\.[A-Za-z]{2,}$/.test(mail))obs.push(`Correo incompleto («${mail}»)`);
    const v={Q:b[3],R:b[5],S:b[6],T:b[4],U:ciudad(util(b[7])),V:b[8],W:ECIVIL[b[9]]||b[9],X:NAC[String(b[10]||"").toUpperCase()]||b[10],
      Y:idStr(b[11]),Z:"PRODUBANCO",AA:b[12],AB:b[13],AC:util(b[15]),AE:resiF,AF:util(b[16]),AG:provincia(resiF),
      AH:b[36],AI:"ECUADOR",AJ:util(b[34]),AK:b[33],AL:util(b[35]),AM:titulo(b[33],b[37]),AN:mail,AO:cel(b[38])};
    if(b.corrida)obs.push(`Fila corregida: venía corrida ${b.corrida} columna(s) en la base`);
    const sx=sexoNombre(b[4]);const casada=/casad|uni[oó]n/i.test(b[9]||"");
    const tratSug=sx==="VARON"?"Sr.":sx==="MUJER"?(casada?"Sra.":"Srta."):null;
    if(!b[3]){
      if(tratSug){v.Q=tratSug;obs.push(`Tratamiento vacío: se puso ${tratSug} por el nombre`)}
      else pend.push({key:r+"Q",row:r,col:"Q",who:nombre,what:"Tratamiento del colaborador",opts:["Sr.","Sra.","Srta."]});
    }else if(sx&&((sx==="VARON")!==(b[3]==="Sr."))){v.Q=tratSug;obs.push(`Tratamiento «${b[3]}» no coincide con el nombre: se cambió a ${tratSug}`)}
    if(v.Y&&!/^\d{10}$/.test(v.Y))obs.push(`Cédula con ${v.Y.length} dígitos`);
    else if(v.Y&&!cedOk(v.Y)&&v.X==="ECUADOR")obs.push(`Cédula ${v.Y} no pasa el dígito verificador: revisa si está mal escrita`);
    if(v.AA!=null&&!/^\d{10,11}$/.test(String(v.AA)))obs.push(`Cuenta Produbanco con ${String(v.AA).length} caracteres («${v.AA}»)`);
    if(v.AO!=null&&String(v.AO).replace(/\D/g,"").length!==10)obs.push(`Celular raro («${v.AO}»)`);
    const cargas=[];
    for(const [tipo,idx] of [["Cónyuge",[40,18,19,20,21,22]],["Hijo",[41,23,24,25,26,27]],["Hijo",[42,28,29,30,31,32]]]){
      const c=idx.map(i=>util(b[i]));
      if(c[1]){
        const ced=idStr(c[5]);
        if(c[5]&&!/^\d+$/.test(ced)){obs.push(`${tipo} ${c[1]}: en cédula escribió «${c[5]}», quedó vacía`);c[5]=null}
        else if(c[5]&&!cedOk(ced))obs.push(`${tipo} ${c[1]}: cédula ${ced} no pasa el dígito verificador`);
        if(c[2]==null)obs.push(`${tipo} ${c[1]}: sin fecha de nacimiento`);
        c[0]=c[0]&&nombreFix(c[0]);c[4]=ciudad(c[4]);
        cargas.push([tipo,...c]);
      }
    }
    if(cargas.length>2)obs.push(`${cargas.length} cargas y el formato tiene 2 espacios: falta ${cargas[2][2]}`);
    const slots=[["AP","AQ","AR","AS","AT","AU","AV","AW","AX","AY"],["BC","BD","BE","BF","BG","BH","BI","BJ","BK","BL"]];
    cargas.slice(0,2).forEach((c,i)=>{
      const [tipo,ap,nom,fnac,nac,ciu,ced]=c;const [a1,a2]=splitAp(ap);const n=NAC[String(nac||"").toUpperCase()]||nac;
      [tipo,a1,a2,nom,null,fnac,n,ciu,(n&&n!=="OTRA")?n:null,idStr(ced)].forEach((x,j)=>v[slots[i][j]]=x);
      if(!ap)obs.push(`${tipo} ${nom} sin apellidos`);
      let sc=sexoNombre(nom);
      if(!sc&&tipo==="Cónyuge"){const t=v.Q||b[3];sc=t==="Sr."?"MUJER":(t==="Sra."||t==="Srta.")?"VARON":null}
      if(sc)v[slots[i][4]]=sc;
      pend.push({key:r+slots[i][4],row:r,col:slots[i][4],who:nombre,what:`Sexo de ${nom} (${tipo.toLowerCase()})`,opts:["VARON","MUJER"],auto:sc});
    });
    for(const c of ["V","AH","AU","BH"])if(v[c]!=null&&toSerial(v[c])==null){obs.push(`Fecha inválida en ${c} («${v[c]}»), quedó vacía`);v[c]=null}
    fill[r]=v;ok++;
    rep.push({pos,nombre,fb:b.filaBase,fila:r,estado:obs.length?"warn":"ok",obs:obs.join(" · ")||"Listo"});
  }
  return {rep,fill,pend,ok,bad};
}

/* ---------- escribir XML ---------- */
function setCell(row,ref,val,kind){
  const re=new RegExp('<c r="'+ref+'"([^>]*?)(/>|>[\\s\\S]*?</c>)');
  const m=row.match(re);
  const s=m?((m[1].match(/ s="\d+"/)||[""])[0]):"";
  const cell=val==null?`<c r="${ref}"${s}/>`:kind==="num"?`<c r="${ref}"${s}><v>${val}</v></c>`:`<c r="${ref}"${s} t="inlineStr"><is><t xml:space="preserve">${esc(val)}</t></is></c>`;
  if(m)return row.slice(0,m.index)+cell+row.slice(m.index+m[0].length);
  if(val==null)return row;
  if(/\/>$/.test(row))return row.replace(/\/>$/,">"+cell+"</row>");
  const tgt=colIdx(ref.replace(/\d+/,""));
  for(const c of row.matchAll(/<c r="([A-Z]+)\d+"/g)){if(colIdx(c[1])>tgt)return row.slice(0,c.index)+cell+row.slice(c.index)}
  return row.replace(/<\/row>$/,cell+"</row>");
}
function fillRow(row,r,vals){
  for(const [col,raw] of Object.entries(vals)){
    let v=raw,kind="str";
    if(DATE_COLS.has(col)){const s=toSerial(raw);if(s){v=s;kind="num"}}
    if(v!=null&&kind==="str")v=String(v);
    row=setCell(row,col+r,v,kind);
  }
  // borrar valores guardados de fórmulas para que se recalculen al abrir
  row=row.replace(new RegExp('(<c r="[A-Z]+'+r+'"[^>]*?) t="(?:str|e|b)"(>(?:<f[^>]*/>|<f[^>]*>[^<]*</f>))<v>[^<]*</v>',"g"),"$1$2");
  row=row.replace(new RegExp('(<c r="[A-Z]+'+r+'"[^>]*>(?:<f[^>]*/>|<f[^>]*>[^<]*</f>))<v>[^<]*</v>',"g"),"$1");
  row=row.replace(new RegExp('(<c r="[A-Z]+'+r+'"[^>]*>(?:<f[^>]*/>|<f[^>]*>[^<]*</f>))<v/>',"g"),"$1");
  return row;
}
async function generateXlsx(){
  const fill=JSON.parse(JSON.stringify(MODEL.fill));
  for(const p of MODEL.pend){if(edits[p.key])fill[p.row][p.col]=edits[p.key]}
  const xml=FMT.sheet.xml.replace(/<row r="(\d+)"[^>]*?(?:\/>|>[\s\S]*?<\/row>)/g,(m,r)=>fill[r]?fillRow(m,+r,fill[r]):m);
  let wbx=FMT.wbx;
  if(/<calcPr\b/.test(wbx)&&!/fullCalcOnLoad=/.test(wbx))wbx=wbx.replace(/<calcPr\b/,'<calcPr fullCalcOnLoad="1"');
  const zip=await JSZip.loadAsync(FMT.buf.slice(0)); // copia limpia en cada descarga
  zip.file(FMT.sheet.path,xml,{createFolders:false});zip.file("xl/workbook.xml",wbx,{createFolders:false});
  const isM=/\.xlsm$/i.test(FMT.name);
  const xbytes=await zip.generateAsync({type:"uint8array",compression:"DEFLATE",mimeType:isM?"application/vnd.ms-excel.sheet.macroEnabled.12":"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"});
  return {xbytes,isM};
}
function buildCsv(){
  const csvRows=[["POSICION","COLABORADOR","FILA EN LA BASE","FILA EN EL FORMATO","ESTADO","OBSERVACION"],...MODEL.rep.map(x=>[x.pos,x.nombre,x.fb??"",x.fila,{ok:"OK",warn:"REVISAR",bad:"NO CARGADO"}[x.estado],x.obs])];
  const csv="﻿"+csvRows.map(r=>r.map(c=>{let x=String(c??"");if(/^[=+\-@]/.test(x))x="'"+x;return `"${x.replace(/"/g,'""')}"`}).join(",")).join("\r\n");
  return csv;
}
async function generateAll(){
  const {xbytes,isM}=await generateXlsx();
  const csv=buildCsv();
  const stamp=new Date().toISOString().slice(0,10);
  const outName=FMT.name.replace(/\.(xlsm|xlsx)$/i,"")+"_LLENO."+(isM?"xlsm":"xlsx");
  return {xlsx:new Blob([xbytes],{type:isM?"application/vnd.ms-excel.sheet.macroEnabled.12":"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"}),outName,
    csv:new Blob([csv],{type:"text/csv;charset=utf-8"}),csvName:`reporte_validacion_${stamp}.csv`};
}

/* ---------- UI ---------- */
function render(){
  const ready=BASE&&FMT;
  $("btn-dl").disabled=!ready;$("btn-csv").disabled=!ready;
  if(!ready){$("results").hidden=true;return}
  for(const k in edits)delete edits[k];
  MODEL=build();
  for(const p of MODEL.pend)if(p.auto)edits[p.key]=p.auto;
  $("results").hidden=false;
  $("s-base").textContent=BASE.data.length;$("s-ok").textContent=MODEL.ok;$("s-bad").textContent=MODEL.bad;
  updatePendCount();
  $("pend-panel").hidden=!MODEL.pend.length;
  $("pend-list").innerHTML=MODEL.pend.slice().sort((a,b)=>(!!a.auto)-(!!b.auto)).map(p=>`<div class="pend-row"><div><div class="who">${esc(p.who)}</div><div class="what">${esc(p.what)} · fila ${p.row}, columna ${p.col}${p.auto?" · <b>puesto automático, cambia si no es correcto</b>":""}</div></div>
    <select id="sel-${p.key}" data-key="${p.key}" aria-label="${esc(p.what)}"><option value="">Elegir…</option>${p.opts.map(o=>`<option${edits[p.key]===o?" selected":""}>${o}</option>`).join("")}</select></div>`).join("");
  $("pend-list").querySelectorAll("select").forEach(s=>{if(s.value)s.classList.add("filled");s.addEventListener("change",()=>{edits[s.dataset.key]=s.value||undefined;s.classList.toggle("filled",!!s.value);updatePendCount()})});
  renderTable();
  $("dl-note").className="dl-note";
  $("dl-note").textContent=`${MODEL.ok} colaboradores listos para el formato.`;
}
function updatePendCount(){
  const n=MODEL.pend.filter(p=>!edits[p.key]).length;$("s-pend").textContent=n;
}
let onlyAlerts=false;
function renderTable(){
  const rows=MODEL.rep.filter(x=>!onlyAlerts||x.estado!=="ok");
  const w=MODEL.rep.filter(x=>x.estado==="warn").length;
  $("rep-sub").textContent=`${MODEL.rep.length} colaboradores · ${w} con alertas para revisar en el formato.`;
  const lbl={ok:"OK",warn:"Revisar",bad:"No cargado"};
  $("rep-body").innerHTML=rows.map(x=>`<tr><td class="num">${esc(x.pos)}</td><td>${esc(x.nombre)}</td><td class="num">${x.fb??""}</td><td class="num">${x.fila}</td><td><span class="pill ${x.estado}">${lbl[x.estado]}</span></td><td>${esc(x.obs)}</td></tr>`).join("")||`<tr><td colspan="6">Sin alertas.</td></tr>`;
}
$("btn-filter").addEventListener("click",e=>{onlyAlerts=!onlyAlerts;e.currentTarget.setAttribute("aria-pressed",onlyAlerts);e.currentTarget.textContent=onlyAlerts?"Ver todos":"Ver solo alertas";renderTable()});

function wire(kind,reader){
  const drop=$("drop-"+kind),input=$("in-"+kind),meta=$("meta-"+kind),fn=$("fn-"+kind);
  async function handle(file){
    if(!file)return;
    drop.classList.remove("done","err");fn.textContent=file.name;meta.className="meta";meta.textContent="Leyendo…";
    try{
      const res=await reader(file);
      if(kind==="base"){BASE=res;meta.textContent=`${res.data.length} colaboradores encontrados`+(res.sinPos?` · ${res.sinPos} sin posición`:"")}
      else{FMT=res;meta.textContent=`Hoja «${res.sheet.name}» · ${Object.keys(res.filaPos).length} posiciones`}
      drop.classList.add("done");
    }catch(err){
      if(kind==="base")BASE=null;else FMT=null;
      drop.classList.add("err");meta.className="meta bad";meta.textContent=err.message||"No pude leer este archivo.";
    }
    render();
  }
  input.addEventListener("change",()=>{handle(input.files[0]);input.value=""});
  ["dragenter","dragover"].forEach(ev=>drop.addEventListener(ev,e=>{e.preventDefault();drop.classList.add("over")}));
  ["dragleave","drop"].forEach(ev=>drop.addEventListener(ev,e=>{e.preventDefault();drop.classList.remove("over")}));
  drop.addEventListener("drop",e=>handle(e.dataTransfer.files[0]));
}
wire("base",readBase);wire("fmt",readFmt);
["dragover","drop"].forEach(ev=>window.addEventListener(ev,e=>e.preventDefault()));

function bajar(blob,name){
  const url=URL.createObjectURL(blob);const a=document.createElement("a");
  a.href=url;a.download=name;document.body.appendChild(a);a.click();a.remove();
  setTimeout(()=>URL.revokeObjectURL(url),4000);
}
$("btn-dl").addEventListener("click",async()=>{
  const note=$("dl-note"),btn=$("btn-dl");
  const faltan=MODEL.pend.filter(p=>!edits[p.key]).length;
  btn.disabled=true;$("btn-csv").disabled=true;note.className="dl-note";note.textContent="Generando…";
  try{
    const out=await generateAll();bajar(out.xlsx,out.outName);
    note.className="dl-note ok";
    note.textContent="Listo: "+out.outName+(faltan?`. Quedaron ${faltan} datos sin elegir; esas celdas van vacías.`:".");
  }catch(err){note.className="dl-note bad";note.textContent="No se pudo generar el archivo: "+(err&&err.message||err)}
  finally{btn.disabled=false;$("btn-csv").disabled=false}
});
$("btn-csv").addEventListener("click",async()=>{
  try{bajar(new Blob([buildCsv()],{type:"text/csv;charset=utf-8"}),`reporte_validacion_${new Date().toISOString().slice(0,10)}.csv`)}
  catch(err){$("dl-note").className="dl-note bad";$("dl-note").textContent="No se pudo generar el reporte: "+(err&&err.message||err)}
});
})();