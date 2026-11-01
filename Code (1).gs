// API Inventaris Karang Taruna: Google Apps Script, database = Google Sheets
// Cara pasang: lihat PETUNJUK.txt. Tempel seluruh isi file ini ke Apps Script.
const S={users:['id','nama','username','pass','role','aktif'],categories:['id','nama'],items:['id','category_id','kode','nama','kondisi','status'],borrowers:['id','nama','hp','alamat'],loans:['id','bid','aid','pinjam','rencana','kembali','status'],loan_details:['id','loan_id','iid','k','cat','cr','gt','gn','gq','gu','fp','fk']};
const N=['id','category_id','bid','aid','iid','loan_id'];
const H=s=>Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,'ikt|'+s));
const today=()=>Utilities.formatDate(new Date(),'Asia/Jakarta','yyyy-MM-dd');
const FOTO=(d,n)=>{const m=/^data:image\/(\w+);base64,(.+)$/.exec(d||'');if(!m)return '';const fi=DriveApp.getFoldersByName('Foto Inventaris'),fo=fi.hasNext()?fi.next():DriveApp.createFolder('Foto Inventaris'),f=fo.createFile(Utilities.newBlob(Utilities.base64Decode(m[2]),'image/'+m[1],n+'.'+m[1]));f.setSharing(DriveApp.Access.ANYONE_WITH_LINK,DriveApp.Permission.VIEW);return 'https://drive.google.com/thumbnail?id='+f.getId()+'&sz=w1000'};
function izinkan(){DriveApp.getRootFolder()}
function seedDb(){return{users:[{id:1,nama:'Administrator',username:'admin',pass:H('admin1'),role:'administrator',aktif:true}],categories:[],items:[],borrowers:[],loans:[]}}
const strip=d=>({...d,users:d.users.map(({pass,...x})=>x)});
function handle(db,a,p,u){
const nx=t=>Math.max(0,...db[t].map(x=>x.id))+1,it=id=>db.items.find(i=>i.id==id),adm=()=>{if(u.role!='administrator')throw Error('Hanya Administrator yang boleh melakukan ini.')};
if(a=='login'){const x=db.users.find(x=>x.username==String(p.u||'').trim().toLowerCase()&&x.pass==H(p.p||''));if(!x)throw Error('Username atau password salah.');if(!x.aktif)throw Error('Akun nonaktif. Hubungi Administrator.');return x}
if(!u)throw Error('Sesi habis. Silakan masuk lagi.');
const mk=(n,c,q)=>{const ex=db.items.find(i=>i.nama.trim().toLowerCase()==n.toLowerCase()),nm=ex?ex.nama:n,pf=nm.toUpperCase().replace(/[^A-Z0-9]/g,'').slice(0,4)||'BRG';let k=db.items.filter(i=>i.kode.startsWith(pf+'-')).length;for(let j=0;j<q;j++){let z;do{k++;z=pf+'-'+String(k).padStart(2,'0')}while(db.items.some(i=>i.kode==z));db.items.push({id:nx('items'),category_id:c,kode:z,nama:nm,kondisi:'normal',status:'tersedia'})}};
switch(a){
case'getAll':return;
case'addCategory':{adm();const n=String(p.n||'').trim();if(!n||db.categories.some(c=>c.nama.toLowerCase()==n.toLowerCase()))throw Error('Isi nama kategori yang belum ada.');db.categories.push({id:nx('categories'),nama:n});return}
case'addItems':{adm();const n=String(p.n||'').trim().replace(/\s+/g,' '),q=Math.floor(+p.q),c=+p.c;if(!db.categories.some(x=>x.id==c))throw Error('Buat kategori dulu di tab Kelola.');if(!n)throw Error('Isi nama barang.');if(!(q>=1&&q<=100))throw Error('Jumlah unit antara 1 dan 100.');
mk(n,c,q);return}
case'fixItem':{adm();const i=it(p.id);if(!i)throw Error('Barang tidak ditemukan.');i.status='tersedia';i.kondisi='normal';return}
case'addUser':{adm();const n=String(p.n||'').trim(),x=String(p.u||'').trim().toLowerCase();if(!n||!x)throw Error('Isi nama dan username.');if(db.users.some(y=>y.username==x))throw Error('Username sudah dipakai.');if(String(p.p||'').length<6)throw Error('Password minimal 6 karakter.');db.users.push({id:nx('users'),nama:n,username:x,pass:H(p.p),role:'admin',aktif:true});return}
case'toggleUser':{adm();const x=db.users.find(y=>y.id==p.id&&y.role=='admin');if(!x)throw Error('Akun tidak ditemukan.');x.aktif=!x.aktif;return}
case'createLoan':{const ids=[...new Set((p.ids||[]).map(Number))];if(!ids.length)throw Error('Pilih minimal satu barang.');if(ids.some(i=>!it(i)||it(i).status!='tersedia'))throw Error('Ada barang yang sudah tidak tersedia.');if(!p.d1||!p.d2||p.d2<p.d1)throw Error('Tanggal rencana kembali tidak boleh sebelum tanggal pinjam.');
let b=+p.b;if(p.nb){const n=String(p.nb.nama||'').trim();if(!n)throw Error('Isi nama peminjam.');b=nx('borrowers');db.borrowers.push({id:b,nama:n,hp:p.nb.hp||'',alamat:p.nb.alamat||''})}else if(!db.borrowers.some(x=>x.id==b))throw Error('Pilih peminjam.');
ids.forEach(i=>it(i).status='dipinjam');db.loans.push({id:nx('loans'),bid:b,aid:u.id,pinjam:p.d1,rencana:p.d2,kembali:'',status:'dipinjam',d:ids.map(i=>({iid:i,k:'',cat:String((p.cats||{})[i]||'').trim(),fp:(p.fotos||{})[i]?FOTO(p.fotos[i],'pinjam-'+i+'-'+Date.now()):''}))});return}
case'returnLoan':{const l=db.loans.find(x=>x.id==p.id&&x.status=='dipinjam');if(!l)throw Error('Peminjaman tidak aktif.');const rs=p.rs||{};
for(const x of l.d){const r=rs[x.iid]||{};if((r.k||'baik')!='baik'){if(r.t=='uang'&&!(+r.u>0))throw Error('Isi nominal ganti rugi uang.');if(r.t=='barang'&&!(Math.floor(+r.q)>=1&&String(r.n||'').trim()))throw Error('Isi nama dan jumlah barang pengganti.')}}
l.d.forEach(x=>{const r=rs[x.iid]||{},k=r.k||'baik',i=it(x.iid);x.k=k;x.cr=String(r.c||'').trim();if(r.f)x.fk=FOTO(r.f,'kembali-'+x.iid+'-'+Date.now());i.status=k=='baik'?'tersedia':k;i.kondisi=k=='baik'?'normal':k;x.gt='';x.gn='';x.gq='';x.gu='';
if(k!='baik'&&r.t=='uang'){x.gt='uang';x.gu=Math.round(+r.u)}
if(k!='baik'&&r.t=='barang'){x.gt='barang';x.gn=String(r.n).trim();x.gq=Math.floor(+r.q);mk(x.gn,i.category_id,x.gq)}});
l.kembali=today();l.status='selesai';return}
case'editItem':{adm();const i=it(p.id),n=String(p.n||'').trim().replace(/\s+/g,' ');if(!i)throw Error('Barang tidak ditemukan.');if(!n)throw Error('Isi nama barang.');if(!db.categories.some(c=>c.id==p.c))throw Error('Pilih kategori.');i.nama=n;i.category_id=+p.c;return}
case'setKondisi':{const i=it(p.id);if(!i)throw Error('Barang tidak ditemukan.');if(i.status=='dipinjam')throw Error('Barang sedang dipinjam. Ubah kondisinya saat pengembalian.');const m={normal:['tersedia','normal'],rusak:['rusak','rusak'],perbaikan:['perbaikan','perbaikan'],hilang:['hilang','hilang']}[p.k];if(!m)throw Error('Kondisi tidak valid.');i.status=m[0];i.kondisi=m[1];return}
case'deleteItem':{adm();const i=it(p.id);if(!i)throw Error('Barang tidak ditemukan.');if(i.status=='dipinjam'||db.loans.some(l=>l.d.some(x=>x.iid==i.id)))throw Error('Barang pernah dipinjam, tidak bisa dihapus. Ubah statusnya saja.');db.items=db.items.filter(x=>x!==i);return}
case'editCategory':{adm();const c=db.categories.find(x=>x.id==p.id),n=String(p.n||'').trim();if(!c||!n)throw Error('Isi nama kategori.');if(db.categories.some(x=>x!==c&&x.nama.toLowerCase()==n.toLowerCase()))throw Error('Nama kategori sudah ada.');c.nama=n;return}
case'deleteCategory':{adm();if(db.items.some(i=>i.category_id==p.id))throw Error('Kategori masih dipakai barang.');db.categories=db.categories.filter(c=>c.id!=p.id);return}
case'editBorrower':{const b=db.borrowers.find(x=>x.id==p.id),n=String(p.n||'').trim();if(!b||!n)throw Error('Isi nama peminjam.');b.nama=n;b.hp=p.h||'';b.alamat=p.a||'';return}
case'deleteBorrower':{adm();if(db.loans.some(l=>l.bid==p.id))throw Error('Peminjam punya riwayat peminjaman, tidak bisa dihapus.');db.borrowers=db.borrowers.filter(b=>b.id!=p.id);return}
case'editUser':{adm();const x=db.users.find(y=>y.id==p.id),n=String(p.n||'').trim(),nu=String(p.u||'').trim().toLowerCase();if(!x||!n||!nu)throw Error('Isi nama dan username.');if(db.users.some(y=>y!==x&&y.username==nu))throw Error('Username sudah dipakai.');if(p.p){if(String(p.p).length<6)throw Error('Password minimal 6 karakter.');x.pass=H(p.p)}x.nama=n;x.username=nu;return}
case'deleteUser':{adm();const x=db.users.find(y=>y.id==p.id&&y.role=='admin');if(!x)throw Error('Hanya akun admin yang bisa dihapus.');if(db.loans.some(l=>l.aid==x.id))throw Error('Akun pernah mencatat peminjaman. Nonaktifkan saja.');db.users=db.users.filter(y=>y!==x);return}
case'editLoan':{const l=db.loans.find(x=>x.id==p.id);if(!l)throw Error('Peminjaman tidak ditemukan.');if(!p.r||p.r<l.pinjam)throw Error('Tanggal rencana kembali tidak boleh sebelum tanggal pinjam.');l.rencana=p.r;l.d.forEach(x=>{if(('c_'+x.iid) in p)x.cat=String(p['c_'+x.iid]).trim()});return}
case'deleteLoan':{adm();const l=db.loans.find(x=>x.id==p.id);if(!l)throw Error('Peminjaman tidak ditemukan.');if(l.status=='dipinjam')l.d.forEach(x=>{const i=it(x.iid);if(i)i.status='tersedia'});db.loans=db.loans.filter(x=>x!==l);return}
}throw Error('Aksi "'+a+'" tidak dikenal. Code.gs di Apps Script belum versi terbaru.')}
function readAll(){const ss=SpreadsheetApp.getActive(),d={};
Object.keys(S).forEach(n=>{let sh=ss.getSheetByName(n);if(!sh){sh=ss.insertSheet(n);sh.getRange(1,1,1,S[n].length).setValues([S[n]]).setFontWeight('bold');sh.setFrozenRows(1);sh.getRange(2,1,sh.getMaxRows()-1,S[n].length).setNumberFormat('@')}
d[n]=sh.getDataRange().getValues().slice(1).filter(r=>r[0]!=='').map(r=>{const o={};S[n].forEach((h,i)=>o[h]=N.includes(h)?Number(r[i]):r[i]);return o})});
d.users.forEach(u=>u.aktif=String(u.aktif)=='1'||String(u.aktif).toLowerCase()=='true');
d.loans.forEach(l=>l.d=d.loan_details.filter(x=>x.loan_id==l.id).map(({id,loan_id,...r})=>r));delete d.loan_details;
if(!d.users.length)d.users=seedDb().users;return d}
function writeAll(d){const ss=SpreadsheetApp.getActive(),ld=[];d.loans.forEach(l=>l.d.forEach(x=>ld.push({...x,id:ld.length+1,loan_id:l.id})));const T={...d,loan_details:ld};
Object.keys(S).forEach(n=>{const rows=T[n].map(o=>S[n].map(h=>h=='aktif'?(o.aktif?1:0):(o[h]??''))),sh=ss.getSheetByName(n),old=sh.getLastRow()-1;sh.getRange(1,1,1,S[n].length).setValues([S[n]]);
if(rows.length){const g=sh.getRange(2,1,rows.length,S[n].length);g.setNumberFormat('@');g.setValues(rows)}
if(old>rows.length)sh.getRange(2+rows.length,1,old-rows.length,S[n].length).clearContent()})}
const out=o=>ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
function doGet(){return ContentService.createTextOutput('API Inventaris Karang Taruna aktif, versi 4 (edit, hapus, kondisi, foto)')}
function doPost(e){const L=LockService.getScriptLock();L.waitLock(20000);
try{const q=JSON.parse(e.postData.contents),c=CacheService.getScriptCache(),db=readAll();let t=q.t,u=null;
if(t){const id=c.get('t_'+t);u=id&&db.users.find(x=>x.id==id&&x.aktif)||null}
const r=handle(db,q.a,q.p||{},u),o={ok:true};
if(q.a=='login'){t=Utilities.getUuid();c.put('t_'+t,String(r.id),21600);o.t=t;o.user=strip({users:[r]}).users[0]}
else if(q.a!='getAll')writeAll(db);
o.db=strip(db);return out(o)}catch(x){return out({ok:false,e:x.message})}finally{L.releaseLock()}}
