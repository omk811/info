/** صندوق التواصل — الأستاذ خليل البلوشي | الإصدار 2.0
 * جميع عمليات الإدارة تتطلب جلسة؛ الدوال المساعدة خاصة وتنتهي بشرطة سفلية.
 * ملف خادم كامل مستقل لواجهة GitHub: لا تحتاج إلى ملفات HTML هنا.
 * الإعداد: ADMIN_PASSWORD في الخصائص، ثم تشغيل setupSystem_، ثم نشر إصدار جديد.
 * لا تجعل جدول البيانات أو مجلد الصور عامًا.
 */
const CONFIG = Object.freeze({
  SHEET_ID: '19s4770QnU-16ApEqj5M0eIWx_nKvKJTPFXy-Cn-43QE',
  SHEET_NAME: 'الملاحظات', MAX_TEXT: 2000, MAX_NAME: 100,
  MAX_IMAGE: 400 * 1024, PAGE_SIZE: 12, SESSION_SECONDS: 3600,
  TIMEZONE: 'Asia/Muscat'
});
const HEADERS = ['التاريخ','الاسم','الملاحظة','صورة الملاحظة','الرد','صورة الرد','الحالة','رقم الصف القديم',
  'معرف ثابت','بصمة المتابعة','موافقة نشر السؤال','منشور','آخر تعديل','نسخة السجل','مؤرشف','التصنيف','معرف الإرسال','نشر الصور'];
const C = {DATE:0,NAME:1,NOTE:2,IMAGE:3,REPLY:4,REPLY_IMAGE:5,STATUS:6,ID:8,RECEIPT:9,
  CONSENT:10,PUBLISHED:11,UPDATED:12,VERSION:13,ARCHIVED:14,CATEGORY:15,REQUEST:16,PUBLISH_IMAGES:17};
const CATEGORIES = ['عام','القراءة','النصوص الأدبية','النحو والصرف','الإملاء','الأنشطة والاختبارات'];

/** فحص الاتصال: الخادم مستقل، والصفحات منشورة في GitHub. */
function doGet() {
  const configured = !!PropertiesService.getScriptProperties().getProperty('ADMIN_HASH');
  return json_({
    success: true,
    version: '2.0',
    configured: configured,
    message: configured ? 'نظام المراسلات يعمل — الأستاذ خليل البلوشي' : 'الكود مكتمل؛ شغّل setupSystem_ بعد إعداد ADMIN_PASSWORD.',
    website: 'https://omk811.github.io/info/',
    admin: 'https://omk811.github.io/info/admin.html'
  });
}
function json_(value) {
  return ContentService.createTextOutput(JSON.stringify(value)).setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  let result;
  try {
    if (!e || !e.postData || e.postData.contents.length > 650000) fail_('BAD_REQUEST','طلب غير صالح أو كبير جدًا.');
    result = api(JSON.parse(e.postData.contents));
  } catch (err) { result = error_(err); }
  return json_(result);
}
/** مدخل موحد لجميع العمليات، مع التحقق من الصلاحيات في الخادم. */
function api(data) {
  try {
    if (!data || typeof data !== 'object' || Array.isArray(data)) fail_('BAD_REQUEST','طلب غير صالح.');
    if (JSON.stringify(data).length > 650000) fail_('BAD_REQUEST','حجم الطلب أكبر من المسموح.');
    const type = String(data.type || '');
    if (['adminList','reply','archive','restore','logout'].indexOf(type) >= 0) auth_(data.token);
    switch (type) {
      case 'login': return login_(data);
      case 'logout': CacheService.getScriptCache().remove('session:' + sha_(data.token)); return ok_({});
      case 'challenge': return challenge_(data);
      case 'add': return add_(data);
      case 'publicList': return list_(data, false);
      case 'adminList': return list_(data, true);
      case 'track': return track_(data);
      case 'image': return image_(data);
      case 'reply': return reply_(data);
      case 'archive': return archive_(data, true);
      case 'restore': return archive_(data, false);
      default: fail_('UNSUPPORTED','حدّث الصفحة إلى الإصدار الجديد.');
    }
  } catch (err) { return error_(err); }
}

/** تشغيل يدوي في المحرر فقط. أدخل ADMIN_PASSWORD في خصائص النص البرمجي أولًا.
 * ينقل كلمة المرور إلى بصمة HMAC، ويضيف أعمدة الإصدار الجديد دون حذف بيانات قديمة.
 * لتغيير كلمة المرور: ضع ADMIN_PASSWORD جديدًا ثم أعد تشغيل هذه الدالة.
 */
function setupSystem_() {
  return locked_(function() {
    const p = PropertiesService.getScriptProperties();
    const password = p.getProperty('ADMIN_PASSWORD');
    if (password && (password.length < 12 || password.length > 128)) fail_('SETUP','استخدم كلمة مرور جديدة من 12 إلى 128 حرفًا.');
    if (!password && !p.getProperty('ADMIN_HASH')) fail_('SETUP','أضف ADMIN_PASSWORD في خصائص النص البرمجي.');
    if (password) {
      p.setProperty('ADMIN_PEPPER', random_());
      p.setProperty('ADMIN_HASH', hmac_(password, p.getProperty('ADMIN_PEPPER')));
      p.deleteProperty('ADMIN_PASSWORD'); // لا تبقى كلمة المرور بصورتها الأصلية
    }
    const sheet = sheet_();
    const last = sheet.getLastRow();
    if (last) sheet.getRange(1,1,1,HEADERS.length).setValues([HEADERS]);
    else sheet.appendRow(HEADERS);
    sheet.setFrozenRows(1);
    if (last > 1) {
      const rows = sheet.getRange(2,1,last-1,HEADERS.length).getValues();
      const extension = rows.map(function(r) {
        // الحقول القديمة A:H تبقى كما هي، والسجلات السابقة لا تُنشر تلقائيًا.
        if (!r[C.ID]) {
          r[C.ID] = Utilities.getUuid(); r[C.RECEIPT] = ''; r[C.CONSENT] = false;
          r[C.PUBLISHED] = false; r[C.UPDATED] = new Date(); r[C.VERSION] = 1;
          r[C.ARCHIVED] = false; r[C.CATEGORY] = 'عام'; r[C.REQUEST] = ''; r[C.PUBLISH_IMAGES] = false;
        }
        return r.slice(8,18);
      });
      sheet.getRange(2,9,extension.length,10).setValues(extension);
    }
    if (!p.getProperty('IMAGE_FOLDER_ID')) {
      const folder = DriveApp.createFolder('صور صندوق التواصل — الأستاذ خليل البلوشي');
      p.setProperty('IMAGE_FOLDER_ID', folder.getId());
    }
    p.setProperty('DATA_REV', Utilities.getUuid());
    CacheService.getScriptCache().remove('loginFailures');
    Logger.log('اكتملت التهيئة. السجلات القديمة محفوظة؛ انشر إصدارًا جديدًا من التطبيق.');
    return {success:true};
  });
}
function sheet_() {
  const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAME) || ss.insertSheet(CONFIG.SHEET_NAME);
  if (sheet.getMaxColumns() < HEADERS.length) sheet.insertColumnsAfter(sheet.getMaxColumns(), HEADERS.length - sheet.getMaxColumns());
  return sheet;
}
function ready_() {
  if (!PropertiesService.getScriptProperties().getProperty('ADMIN_HASH')) fail_('SETUP','النظام يحتاج إلى تشغيل setupSystem_ من المحرر.');
}
function rows_() {
  ready_(); const s = sheet_(); const n = s.getLastRow();
  return n < 2 ? [] : s.getRange(2,1,n-1,HEADERS.length).getValues();
}
function ok_(value) { return Object.assign({success:true,version:'2.0'},value); }
function fail_(code,message) { const err = new Error(message); err.code = code; throw err; }
function error_(err) {
  if (!err.code) console.error('خطأ داخلي: ' + String(err.message).slice(0,160));
  return {success:false,code:err.code || 'SERVER',error:err.code ? err.message : 'تعذر إتمام العملية. حاول مجددًا.'};
}
function locked_(fn) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) fail_('BUSY','النظام مشغول لحظيًا؛ حاول مجددًا.');
  try { return fn(); } finally { lock.releaseLock(); }
}
function sha_(s) { return Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,String(s),Utilities.Charset.UTF_8)); }
function hmac_(s,key) { return Utilities.base64EncodeWebSafe(Utilities.computeHmacSha256Signature(String(s),String(key),Utilities.Charset.UTF_8)); }
function random_() { return Utilities.getUuid().replace(/-/g,'') + Utilities.getUuid().replace(/-/g,''); }
function equal_(a,b) { a=String(a); b=String(b); if(a.length!==b.length)return false; let diff=0; for(let i=0;i<a.length;i++)diff|=a.charCodeAt(i)^b.charCodeAt(i); return diff===0; }
function login_(d) {
  ready_();
  return locked_(function() {
    const p=PropertiesService.getScriptProperties(), cache=CacheService.getScriptCache();
    const key='loginFailures', failures=Number(cache.get(key)||0);
    if(failures>=8) fail_('RATE_LIMIT','توقفت محاولات الدخول مؤقتًا. حاول بعد عشر دقائق.');
    if(typeof d.password!=='string' || d.password.length>128 || !equal_(hmac_(d.password,p.getProperty('ADMIN_PEPPER')),p.getProperty('ADMIN_HASH'))) {
      cache.put(key,String(failures+1),600); fail_('LOGIN','بيانات الدخول غير صحيحة.');
    }
    cache.remove(key);
    const token=random_(), expires=Date.now()+CONFIG.SESSION_SECONDS*1000;
    cache.put('session:'+sha_(token),JSON.stringify({expires:expires,epoch:sha_(p.getProperty('ADMIN_HASH'))}),CONFIG.SESSION_SECONDS);
    return ok_({token:token,expires:expires});
  });
}
function auth_(token) {
  if(typeof token!=='string' || !/^[a-f0-9]{64}$/.test(token)) fail_('AUTH','انتهت جلسة الإدارة. سجّل الدخول مجددًا.');
  const raw=CacheService.getScriptCache().get('session:'+sha_(token));
  const session=raw && JSON.parse(raw);
  if(!session || session.expires<Date.now() || session.epoch!==sha_(PropertiesService.getScriptProperties().getProperty('ADMIN_HASH')||'')) fail_('AUTH','انتهت جلسة الإدارة. سجّل الدخول مجددًا.');
}
function device_(d) { if(typeof d.device!=='string' || !/^[a-f0-9-]{36}$/.test(d.device)) fail_('BAD_REQUEST','أعد فتح الصفحة ثم حاول مجددًا.'); return d.device; }
function rate_(key,limit,seconds) {
  const cache=CacheService.getScriptCache(), slot=key+':'+Math.floor(Date.now()/(seconds*1000));
  const count=Number(cache.get(slot)||0); if(count>=limit) fail_('RATE_LIMIT','طلبات كثيرة؛ انتظر قليلًا ثم حاول مجددًا.');
  cache.put(slot,String(count+1),seconds);
}
function challenge_(d) {
  ready_();
  return locked_(function() {
    const device=device_(d); rate_('challengeGlobal',120,60); rate_('challenge:'+device,15,600);
    const ticket=random_(); CacheService.getScriptCache().put('ticket:'+sha_(ticket),JSON.stringify({device:device,time:Date.now()}),600);
    return ok_({ticket:ticket});
  });
}
function text_(value,max,required) {
  if(typeof value!=='string') fail_('VALIDATION','تحقق من النص المدخل.');
  const s=value.trim(); if(s.length>max || (required&&!s)) fail_('VALIDATION','تحقق من النص وطوله.');
  return s.replace(/\u0000/g,'');
}
function cell_(s) { return /^[=+\-@]/.test(String(s)) ? "'"+s : s; }
function uuid_(s) { if(typeof s!=='string'||!/^[a-f0-9]{8}(-[a-f0-9]{4}){3}-[a-f0-9]{12}$/.test(s)) fail_('VALIDATION','معرف غير صالح.'); return s; }
function secret_(s) { if(typeof s!=='string'||!/^[a-f0-9]{64}$/.test(s)) fail_('VALIDATION','رمز المتابعة غير صالح.'); return s; }
function bool_(v) { return v === true || v === 'TRUE'; }
function category_(v) { if(CATEGORIES.indexOf(v)<0)fail_('VALIDATION','اختر تصنيفًا صحيحًا.'); return v; }
function imageBlob_(value) {
  if(!value)return null;
  if(typeof value!=='string'||value.length>560000)fail_('IMAGE','الصورة أكبر من المسموح.');
  const m=value.match(/^data:image\/(jpeg|png|webp);base64,([A-Za-z0-9+/]+={0,2})$/);
  if(!m)fail_('IMAGE','استخدم صورة JPEG أو PNG أو WebP.');
  const bytes=Utilities.base64Decode(m[2]);
  if(!bytes.length||bytes.length>CONFIG.MAX_IMAGE)fail_('IMAGE','الصورة أكبر من 400 كيلوبايت بعد الضغط.');
  const b=bytes.map(function(n){return n&255;});
  const valid=m[1]==='jpeg' ? b[0]===255&&b[1]===216&&b[2]===255 : m[1]==='png' ? [137,80,78,71,13,10,26,10].every(function(n,i){return b[i]===n;}) : b.slice(0,4).join(',')==='82,73,70,70'&&b.slice(8,12).join(',')==='87,69,66,80';
  if(!valid)fail_('IMAGE','ملف الصورة غير صالح.');
  return Utilities.newBlob(bytes,'image/'+m[1],'attachment.'+(m[1]==='jpeg'?'jpg':m[1]));
}
function storeImage_(blob) {
  if(!blob)return '';
  const folder=PropertiesService.getScriptProperties().getProperty('IMAGE_FOLDER_ID');
  if(!folder)fail_('SETUP','أعد تشغيل التهيئة لإعداد الصور.');
  return 'drive:'+DriveApp.getFolderById(folder).createFile(blob).getId();
}
function rev_() { PropertiesService.getScriptProperties().setProperty('DATA_REV',Utilities.getUuid()); }
function add_(d) {
  ready_(); const name=text_(d.name||'متابع',CONFIG.MAX_NAME,true), note=text_(d.note,CONFIG.MAX_TEXT,true);
  const id=uuid_(d.requestId), secret=secret_(d.secret), device=device_(d), category=category_(d.category||'عام');
  if(d.website)fail_('VALIDATION','طلب غير صالح.');
  const blob=imageBlob_(d.image);
  return locked_(function() {
    const s=sheet_(), rows=rows_();
    const duplicate=rows.find(function(r){return r[C.REQUEST]===id;});
    if(duplicate) {
      if(!equal_(duplicate[C.RECEIPT],sha_(secret)))fail_('VALIDATION','معرف إرسال مستخدم.');
      return ok_({id:duplicate[C.ID],receipt:duplicate[C.ID]+'.'+secret,duplicate:true});
    }
    const cache=CacheService.getScriptCache();
    if(typeof d.ticket!=='string'||!/^[a-f0-9]{64}$/.test(d.ticket))fail_('TICKET','انتهت صلاحية الإرسال؛ حاول مجددًا.');
    const raw=cache.get('ticket:'+sha_(d.ticket)), ticket=raw&&JSON.parse(raw);
    if(!ticket||ticket.device!==device||Date.now()-ticket.time<1200)fail_('TICKET','انتظر لحظة ثم حاول الإرسال مجددًا.');
    rate_('addGlobal',30,60); rate_('add:'+device,5,600);
    const reference=storeImage_(blob), date=new Date();
    const r=[date,cell_(name),cell_(note),reference,'','','قيد المراجعة','',id,sha_(secret),d.consent===true,false,date,1,false,category,id,false];
    try { s.appendRow(r); SpreadsheetApp.flush(); } catch(err) { if(reference.indexOf('drive:')===0)try{DriveApp.getFileById(reference.slice(6)).setTrashed(true);}catch(_){} throw err; }
    cache.remove('ticket:'+sha_(d.ticket)); rev_();
    return ok_({id:id,receipt:id+'.'+secret});
  });
}
function plain_(r) {
  return {id:r[C.ID],date:iso_(r[C.DATE]),name:String(r[C.NAME]||'متابع'),note:String(r[C.NOTE]||''),reply:String(r[C.REPLY]||''),
    hasImage:!!r[C.IMAGE],hasReplyImage:!!r[C.REPLY_IMAGE],consent:bool_(r[C.CONSENT]),published:bool_(r[C.PUBLISHED]),
    publishImages:bool_(r[C.PUBLISH_IMAGES]),archived:bool_(r[C.ARCHIVED]),category:r[C.CATEGORY]||'عام',version:Number(r[C.VERSION])||1,updated:iso_(r[C.UPDATED])};
}
function public_(r) {
  const n=plain_(r); n.name='سؤال من متابع'; n.hasImage=n.hasImage&&n.publishImages; n.hasReplyImage=n.hasReplyImage&&n.publishImages;
  delete n.consent; delete n.archived; delete n.version; delete n.published; delete n.publishImages;
  return n;
}
function iso_(value) { const d=new Date(value); return isNaN(d.getTime())?'':d.toISOString(); }
function eligible_(r) { return bool_(r[C.CONSENT])&&bool_(r[C.PUBLISHED])&&!bool_(r[C.ARCHIVED])&&!!String(r[C.REPLY]||'').trim(); }
function list_(d,admin) {
  const q=text_(d.q||'',120,false).toLowerCase(), filter=String(d.filter||'all'), category=d.category||'';
  const page=Math.max(1,Math.min(100000,Math.floor(Number(d.page)||1)));
  const revision=PropertiesService.getScriptProperties().getProperty('DATA_REV')||'0';
  const key='list:'+sha_(JSON.stringify([revision,admin,q,filter,category,page]));
  const cache=CacheService.getScriptCache(), saved=cache.get(key);
  if(saved)return JSON.parse(saved);
  const rows=rows_();
  let list=admin ? rows : rows.filter(eligible_);
  const stats=admin ? {total:rows.filter(r=>!bool_(r[C.ARCHIVED])).length,pending:rows.filter(r=>!bool_(r[C.ARCHIVED])&&!String(r[C.REPLY]||'').trim()).length,
    replied:rows.filter(r=>!bool_(r[C.ARCHIVED])&&!!String(r[C.REPLY]||'').trim()).length,archived:rows.filter(r=>bool_(r[C.ARCHIVED])).length} : null;
  if(admin)list=list.filter(function(r){return filter==='archived'?bool_(r[C.ARCHIVED]):!bool_(r[C.ARCHIVED]);});
  if(admin&&filter==='pending')list=list.filter(r=>!String(r[C.REPLY]||'').trim());
  if(admin&&filter==='replied')list=list.filter(r=>!!String(r[C.REPLY]||'').trim());
  if(category)list=list.filter(r=>r[C.CATEGORY]===category);
  if(q)list=list.filter(function(r){return [admin?r[C.NAME]:'',r[C.NOTE],r[C.REPLY]].join(' ').toLowerCase().indexOf(q)>=0;});
  list.sort((a,b)=>new Date(b[C.DATE])-new Date(a[C.DATE]));
  const total=list.length, pages=Math.max(1,Math.ceil(total/CONFIG.PAGE_SIZE)), current=Math.min(page,pages);
  const result=ok_({data:list.slice((current-1)*CONFIG.PAGE_SIZE,current*CONFIG.PAGE_SIZE).map(admin?plain_:public_),total:total,page:current,pages:pages,stats:stats});
  const json=JSON.stringify(result); if(Utilities.newBlob(json).getBytes().length<90000)cache.put(key,json,45);
  return result;
}
function record_(id) {
  uuid_(id); const rows=rows_(); const index=rows.findIndex(r=>r[C.ID]===id);
  if(index<0)fail_('NOT_FOUND','لم يتم العثور على الرسالة.');
  return {row:index+2,value:rows[index]};
}
function receipt_(d,r) { secret_(d.secret); if(!r[C.RECEIPT]||!equal_(r[C.RECEIPT],sha_(d.secret)))fail_('NOT_FOUND','رمز المتابعة غير صحيح.'); }
function track_(d) {
  const r=record_(d.id).value; receipt_(d,r);
  if(bool_(r[C.ARCHIVED]))return ok_({data:{id:r[C.ID],archived:true,note:'تم إغلاق هذه الرسالة.',reply:'',date:iso_(r[C.DATE])}});
  const n=plain_(r); delete n.name; delete n.consent; delete n.published; delete n.publishImages; delete n.version;
  return ok_({data:n});
}
function image_(d) {
  const r=record_(d.id).value;
  if(d.token)auth_(d.token);
  else if(d.secret) { receipt_(d,r); if(bool_(r[C.ARCHIVED]))fail_('NOT_FOUND','الرسالة مغلقة.'); }
  else if(!eligible_(r)||!bool_(r[C.PUBLISH_IMAGES]))fail_('AUTH','الصورة غير متاحة للعامة.');
  if(['question','reply'].indexOf(d.which)<0)fail_('VALIDATION','طلب صورة غير صالح.');
  const ref=String(r[d.which==='reply'?C.REPLY_IMAGE:C.IMAGE]||'');
  if(!ref)fail_('NOT_FOUND','لا توجد صورة.');
  if(ref.indexOf('data:image/')===0) { imageBlob_(ref); return ok_({image:ref}); }
  if(ref.indexOf('drive:')!==0)fail_('IMAGE','صيغة صورة قديمة غير مدعومة؛ أعد رفعها من الإدارة.');
  const blob=DriveApp.getFileById(ref.slice(6)).getBlob();
  if(blob.getBytes().length>CONFIG.MAX_IMAGE)fail_('IMAGE','الصورة أكبر من المسموح.');
  return ok_({image:'data:'+blob.getContentType()+';base64,'+Utilities.base64Encode(blob.getBytes())});
}
function reply_(d) {
  const reply=text_(d.reply,CONFIG.MAX_TEXT,true), blob=imageBlob_(d.image);
  return locked_(function() {
    auth_(d.token); const record=record_(d.id), r=record.value;
    if(bool_(r[C.ARCHIVED]))fail_('ARCHIVED','استعد الرسالة من الأرشيف أولًا.');
    if(Number(d.version)!==Number(r[C.VERSION]))fail_('CONFLICT','تغيّرت الرسالة. حدّث القائمة قبل الحفظ؛ مسودتك محفوظة.');
    if(d.published===true&&!bool_(r[C.CONSENT]))fail_('CONSENT','صاحب السؤال اختار رسالة خاصة؛ لا يمكن نشرها.');
    const newRef=blob?storeImage_(blob):d.removeImage===true?'':r[C.REPLY_IMAGE];
    r[C.REPLY]=cell_(reply); r[C.REPLY_IMAGE]=newRef; r[C.STATUS]='تم الرد'; r[C.PUBLISHED]=d.published===true;
    r[C.PUBLISH_IMAGES]=d.published===true&&d.publishImages===true; r[C.UPDATED]=new Date(); r[C.VERSION]=Number(r[C.VERSION])+1;
    try { sheet_().getRange(record.row,1,1,HEADERS.length).setValues([r]); SpreadsheetApp.flush(); }
    catch(err) { if(blob&&String(newRef).indexOf('drive:')===0)try{DriveApp.getFileById(newRef.slice(6)).setTrashed(true);}catch(_){} throw err; }
    rev_(); return ok_({data:plain_(r)});
  });
}
function archive_(d,value) {
  return locked_(function() {
    auth_(d.token); const record=record_(d.id), r=record.value;
    if(Number(d.version)!==Number(r[C.VERSION]))fail_('CONFLICT','تغيّرت الرسالة؛ حدّث القائمة قبل تنفيذ العملية.');
    r[C.ARCHIVED]=value; r[C.PUBLISHED]=false; r[C.PUBLISH_IMAGES]=false;
    r[C.UPDATED]=new Date(); r[C.VERSION]=Number(r[C.VERSION])+1;
    sheet_().getRange(record.row,1,1,HEADERS.length).setValues([r]); SpreadsheetApp.flush(); rev_();
    return ok_({});
  });
}
