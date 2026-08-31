/** Borang Hakim PBL STEM — Google Apps Script (projek baharu sahaja). */
const CONFIG = {
  spreadsheetId: '1if8jWab6csiekUljvsDuu9bu250O4ipzzqtSG2plQK8',
  participantSheet: 'SENARAIPESERTA', scoreSheet: 'DATAPENGISIANHAKIM',
  judges: {
    'SEKOLAH RENDAH': ['Dr. Mohd Arif bin Mohd Sarjidan', 'Dr. Nashrah Hani Jamadon'],
    'SEKOLAH MENENGAH': ['Dr. Mohd Arif bin Mohd Sarjidan', 'Dr. Nashrah Hani Jamadon']
  },
  publishWinners: true
};
const RUBRIC = [
  ['A','keaslian','Keaslian','Idea asli dengan kelainan jelas.',5], ['A','kreativiti','Kreativiti','Reka bentuk atau penyelesaian di luar kebiasaan.',5], ['A','integrasi_stem','Integrasi STEM','Elemen Sains, Teknologi, Kejuruteraan dan Matematik jelas.',5], ['A','mesra_pengguna','Mesra Pengguna','Praktikal, selamat dan sesuai untuk pengguna.',5], ['A','kekemasan_produk','Kekemasan Produk','Kemas, tersusun, berfungsi dan menarik.',5], ['A','sumber_rujukan','Sumber & Rujukan','Sumber relevan, sahih dan dinyatakan dengan baik.',5],
  ['B','isu_objektif','Isu & Objektif','Isu jelas, relevan dan objektif boleh dicapai.',5], ['B','inovasi_optimum','Inovasi Optimum','Inovasi menyelesaikan isu dengan berkesan.',5], ['B','kepuasan_pengguna','Kepuasan Pengguna','Memenuhi keperluan dan mempunyai maklum balas positif.',5], ['B','penjimatan_kos','Penjimatan Kos','Mengambil kira kecekapan kos atau bahan.',5], ['B','signifikan_relevan','Signifikan & Relevan','Manfaat nyata dan relevan dalam kehidupan.',5], ['B','kelestarian_projek','Kelestarian Projek','Tahan lama, mesra alam dan berpotensi dikembangkan.',5],
  ['C','kelancaran_komunikasi','Kelancaran Komunikasi','Lancar, yakin, tersusun dan menjawab soalan.',5], ['C','penyampaian_bermaklumat','Penyampaian Bermaklumat','Kandungan jelas, lengkap dan berfokus.',5], ['C','gaya_bahasa','Gaya Bahasa','Bahasa jelas, tepat, sopan dan sesuai.',5], ['C','kerjasama_pasukan','Kerjasama Pasukan','Ahli terlibat aktif dan bekerjasama.',5], ['C','kekemasan_laporan','Kekemasan Laporan','Laporan lengkap, tersusun dan kemas.',5], ['C','kawalan_masa_visual','Kawalan Masa & Visual','Menepati masa serta menggunakan visual/prototaip dengan baik.',5],
  ['D','keunikan_tarikan','Keunikan & Tarikan','Elemen unik dan mengagumkan.',5], ['D','potensi_masa_depan','Potensi Masa Depan','Potensi pengembangan, paten atau komersialisasi.',5]
].map(x => ({section:x[0], key:x[1], label:x[2], description:x[3], max:x[4]}));

function doGet(e) { const action=(e.parameter.action||'bootstrap').toLowerCase(); try { return json_(action==='winners' ? winners_() : bootstrap_()); } catch(err) { return json_({success:false,message:'Sistem belum dapat memuatkan data. Sila cuba lagi.'}); } }
function doPost(e) { try { const p=JSON.parse((e.postData&&e.postData.contents)||'{}'); if(p.action!=='submitScore') return json_({success:false,message:'Permintaan tidak sah.'}); return json_(save_(p)); } catch(err) { return json_({success:false,message:'Markah tidak dapat disimpan. Sila cuba lagi.'}); } }
function json_(data) { return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON); }
function ss_(){ return SpreadsheetApp.openById(CONFIG.spreadsheetId); }
function normal_(s){ return String(s||'').trim().toUpperCase().replace(/\s+/g,' '); }
function col_(map, names){ for(const n of names) if(map[normal_(n)]!==undefined) return map[normal_(n)]; return -1; }
function participants_(){ const sh=ss_().getSheetByName(CONFIG.participantSheet); if(!sh) throw Error('Tab peserta belum diwujudkan.'); const values=sh.getDataRange().getDisplayValues(); if(values.length<2) return []; const h={}; values[0].forEach((v,i)=>h[normal_(v)]=i); const pick=(r,n)=>{const i=col_(h,n);return i<0?'':String(r[i]).trim();}; return values.slice(1).filter(r=>r.some(Boolean)).map(r=>({bil:pick(r,['BIL','ID','ID PESERTA']), namaPeserta:pick(r,['NAMA PESERTA','NAMA AHLI','AHLI KUMPULAN']), sekolah:pick(r,['SEKOLAH','SEKOLAH/ORGANISASI','ORGANISASI']), kategori:normal_(pick(r,['KATEGORI'])), bidang:pick(r,['BIDANG','MATA PELAJARAN','MATAPELAJARAN']), tajuk:pick(r,['TAJUK PROJEK','TAJUK'])})).filter(x=>x.bil&&x.kategori); }
function scoreHeaders_(){ return ['TIMESTAMP','ID REKOD','BIL','NAMA PESERTA','SEKOLAH/ORGANISASI','KATEGORI','BIDANG','TAJUK PROJEK','NAMA HAKIM'].concat(RUBRIC.map(x=>x.key.toUpperCase()),['JUMLAH A','JUMLAH B','JUMLAH C','JUMLAH D','JUMLAH KESELURUHAN','CATATAN HAKIM']); }
function scoreSheet_(){ const ss=ss_(); let sh=ss.getSheetByName(CONFIG.scoreSheet); if(!sh) sh=ss.insertSheet(CONFIG.scoreSheet); const headers=scoreHeaders_(); if(sh.getLastRow()===0) sh.getRange(1,1,1,headers.length).setValues([headers]).setFontWeight('bold').setFrozenRows(1); return sh; }
function scores_(){ const sh=scoreSheet_(), vals=sh.getDataRange().getDisplayValues(); if(vals.length<2)return []; const m={};vals[0].forEach((v,i)=>m[normal_(v)]=i); const g=(r,k)=>r[m[k]]||'';return vals.slice(1).filter(r=>r.some(Boolean)).map(r=>({id:g(r,'ID REKOD'),bil:g(r,'BIL'),kategori:normal_(g(r,'KATEGORI')),hakim:g(r,'NAMA HAKIM'),total:Number(g(r,'JUMLAH KESELURUHAN'))||0, sections:{A:Number(g(r,'JUMLAH A'))||0,B:Number(g(r,'JUMLAH B'))||0,C:Number(g(r,'JUMLAH C'))||0,D:Number(g(r,'JUMLAH D'))||0}, raw:r})); }
function bootstrap_(){ return {success:true,competition:'Borang Hakim PBL STEM',judges:CONFIG.judges,rubric:RUBRIC,participants:participants_(),scores:scores_().map(x=>({bil:x.bil,kategori:x.kategori,hakim:x.hakim,total:x.total})),categories:Object.keys(CONFIG.judges)}; }
function validate_(p, person){ const category=normal_(p.category), judge=String(p.judge||'').trim(); if(!CONFIG.judges[category]||CONFIG.judges[category].indexOf(judge)<0)throw Error('Hakim atau kategori tidak sah.'); if(!person||person.kategori!==category)throw Error('Peserta tidak ditemui dalam kategori hakim.'); const got=p.scores||{}, clean={}, sections={A:0,B:0,C:0,D:0}; RUBRIC.forEach(x=>{const n=got[x.key];if(!Number.isInteger(n)||n<0||n>x.max)throw Error('Markah tidak sah.');clean[x.key]=n;sections[x.section]+=n;});return {category,judge,clean,sections,total:sections.A+sections.B+sections.C+sections.D}; }
function save_(p){ const lock=LockService.getScriptLock(); if(!lock.tryLock(30000)) return {success:false,message:'Sistem sedang menyimpan rekod lain. Sila cuba semula.'}; try { const requestedCategory=normal_(p.category), people=participants_(), person=people.find(x=>String(x.bil)===String(p.participantBil)&&x.kategori===requestedCategory); const v=validate_(p,person), sh=scoreSheet_(); const existing=scores_().some(x=>String(x.bil)===String(person.bil)&&x.kategori===v.category&&x.hakim===v.judge); if(existing)return {success:false,code:'DUPLICATE',message:'Markah kumpulan ini telah direkodkan oleh hakim ini.'}; const row=[new Date(),Utilities.getUuid(),person.bil,person.namaPeserta,person.sekolah,v.category,person.bidang,person.tajuk,v.judge].concat(RUBRIC.map(x=>v.clean[x.key]),[v.sections.A,v.sections.B,v.sections.C,v.sections.D,v.total,String(p.comment||'').slice(0,2000)]); sh.appendRow(row); return {success:true,message:'Markah berjaya disimpan.',recordId:row[1],total:v.total}; } catch(e){return {success:false,message:e.message||'Markah tidak dapat disimpan.'};} finally {lock.releaseLock();} }
const AWARDS = [
  {key:'PEMBELAJARAN_BESTARI', label:'Anugerah Pembelajaran Bestari', capacity:3, metric:'total'},
  {key:'PEMBENTANGAN_MISALI', label:'Anugerah Pembentangan Misali', capacity:3, metric:'C'},
  {key:'IDEA_INOVATIF', label:'Anugerah Idea Inovatif', capacity:2, metric:'A'},
  {key:'PEMBELAJARAN_BERIMPAK', label:'Anugerah Pembelajaran Berimpak', capacity:2, metric:'B'}
];
function metric_(row, award){ return award.metric==='total'?row.total:(row.sectionAverage[award.metric]/30*100); }
function compareForAward_(award){ return (x,y)=>metric_(y,award)-metric_(x,award)||y.total-x.total||y.sectionAverage.A-x.sectionAverage.A||y.sectionAverage.B-x.sectionAverage.B||y.sectionAverage.C-x.sectionAverage.C||y.sectionAverage.D-x.sectionAverage.D||String(x.bil).localeCompare(String(y.bil),undefined,{numeric:true}); }
function allocateAwards_(rows){
  if(!rows.length)return rows;
  const ordered=rows.slice().sort((x,y)=>y.total-x.total||String(x.bil).localeCompare(String(y.bil),undefined,{numeric:true})), counts=AWARDS.map(()=>0), current=[], best={score:-Infinity,signature:'',assignment:null};
  function walk(i,score){
    if(i===ordered.length){const signature=current.join('');if(score>best.score+1e-9||(Math.abs(score-best.score)<1e-9&&(!best.signature||signature<best.signature))){best.score=score;best.signature=signature;best.assignment=current.slice();}return;}
    AWARDS.forEach((award,a)=>{if(counts[a]>=award.capacity)return;counts[a]++;current[i]=a;walk(i+1,score+metric_(ordered[i],award));counts[a]--;});
  }
  walk(0,0);
  ordered.forEach((row,i)=>{const award=AWARDS[best.assignment[i]];row.award=award.key;row.awardLabel=award.label;row.awardMetric=metric_(row,award);});
  AWARDS.forEach(award=>ordered.filter(x=>x.award===award.key).sort(compareForAward_(award)).forEach((x,i)=>x.awardRank=i+1));
  return ordered;
}
function winners_(){
  if(!CONFIG.publishWinners)return {success:true,published:false,message:'Keputusan belum diterbitkan.',categories:Object.keys(CONFIG.judges),winners:{}};
  const ps=participants_(), ss=scores_(), out={}, progress={};
  Object.keys(CONFIG.judges).forEach(cat=>{
    const categoryPeople=ps.filter(p=>p.kategori===cat), categoryScores=ss.filter(s=>s.kategori===cat), expected=categoryPeople.length*CONFIG.judges[cat].length;
    let rows=categoryPeople.map(p=>{const records=categoryScores.filter(s=>String(s.bil)===String(p.bil)), judgeCount=records.length;if(!judgeCount)return null;const sum=k=>records.reduce((n,x)=>n+(k==='total'?x.total:x.sections[k]),0);const sectionAverage={A:sum('A')/judgeCount,B:sum('B')/judgeCount,C:sum('C')/judgeCount,D:sum('D')/judgeCount};return {bil:p.bil,tajuk:p.tajuk,namaPeserta:p.namaPeserta,sekolah:p.sekolah,total:sum('total')/judgeCount,sectionAverage,judgeCount,complete:judgeCount===CONFIG.judges[cat].length};}).filter(Boolean);
    rows.sort((x,y)=>y.total-x.total||y.sectionAverage.A-x.sectionAverage.A||y.sectionAverage.B-x.sectionAverage.B||y.sectionAverage.C-x.sectionAverage.C||y.sectionAverage.D-x.sectionAverage.D||String(x.bil).localeCompare(String(y.bil),undefined,{numeric:true})).forEach((x,i)=>x.kedudukan=i+1);
    out[cat]=allocateAwards_(rows);
    progress[cat]={received:categoryScores.length,expected,evaluatedProjects:rows.length,totalProjects:categoryPeople.length,complete:categoryScores.length===expected};
  });
  return {success:true,published:true,provisional:Object.values(progress).some(x=>!x.complete),categories:Object.keys(CONFIG.judges),awards:AWARDS.map(x=>({key:x.key,label:x.label,capacity:x.capacity})),progress,winners:out};
}
