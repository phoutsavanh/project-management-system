// =====================================================
// CONFIGURATION — ແກ້ໄຂສະເພາະ 4 ແຖວນີ້
// =====================================================
const SHEET_ID = "1BCjLMiVasdydKwICC_CezRZXbO6Y_NLGy7zVTyuD2fY";
const WEB_APP_URL = "";
const FOLDER_ID = "1Ut9byWvQgTy__YMQyVKTMMR7mXgjk4lh";
const ADMIN_PASSWORD = "1234";

// =====================================================
// CONSTANTS / SHEET SCHEMA
// =====================================================
const APP_TZ = 'Asia/Bangkok';
const SESSION_TTL = 21600; // 6 ຊົ່ວໂມງ
const SHEETS = {
  Settings: ['key','value','updatedAt','updatedBy'],
  Users: ['userId','username','password','fullName','position','departmentId','role','status','createdAt','updatedAt'],
  Departments: ['departmentId','name','description','status','createdAt','updatedAt'],
  Projects: ['projectId','projectCode','projectName','fiscalYear','planName','departmentId','ownerUserId','coOwners','rationale','objectives','quantitativeTarget','qualitativeTarget','targetGroup','location','startDate','endDate','budget','budgetSource','successIndicators','expectedResults','status','progressPercent','note','createdAt','updatedAt','createdBy','updatedBy'],
  Activities: ['activityId','projectId','activityName','description','owner','startDate','endDate','budget','result','successPercent','status','note','createdAt','updatedAt'],
  Expenses: ['expenseId','projectId','activityId','itemNo','expenseDate','description','expenseType','amount','documentNo','requester','status','note','attachmentFileId','createdAt','createdBy'],
  Progress: ['progressId','projectId','reportDate','detail','output','problem','solution','progressPercent','createdBy','imageFileIds','attachmentFileIds','createdAt'],
  Reports: ['reportId','projectId','resultSummary','indicatorResult','problem','suggestion','activityImages','attachments','createdAt','updatedAt','createdBy','updatedBy'],
  Files: ['fileId','projectId','activityId','fileName','mimeType','driveFileId','driveUrl','folderPath','uploadedAt','uploadedBy'],
  Logs: ['logId','timestamp','userId','username','action','detail','recordId']
};
const DEFAULT_SETTINGS = {
  systemName: 'ລະບົບບໍລິຫານຈັດການໂຄງການ',
  schoolName: 'ຊື່ອົງກອນ / ສະຖານສຶກສາ',
  address: '',
  affiliation: '',
  logoUrl: '',
  bannerUrl: '',
  directorName: 'ຊື່ຜູ້ບໍລິຫານ',
  defaultFiscalYear: String(new Date().getFullYear()),
  footerText: 'ລະບົບບໍລິຫານຈັດການໂຄງການ'
};

// =====================================================
// WEB APP & HTTP API (FOR VERCEL / GITHUB / EMBED)
// =====================================================
function doGet(e) {
  if (e && e.parameter && e.parameter.action === 'ping') {
    return ContentService.createTextOutput(JSON.stringify(ok_('API is active'))).setMimeType(ContentService.MimeType.JSON);
  }
  return HtmlService.createTemplateFromFile('index').evaluate()
    .setTitle('ລະບົບບໍລິຫານຈັດການໂຄງການ')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function doPost(e) {
  try {
    let payload = {};
    if (e && e.postData && e.postData.contents) {
      payload = JSON.parse(e.postData.contents);
    }
    const action = payload.action;
    const args = Array.isArray(payload.args) ? payload.args : [];

    const actions = {
      login, logout, getAppBootstrap, getDashboard,
      listProjects, getProject, getProjectDetail, saveProject, deleteProject, copyProject,
      listActivities, saveActivity, deleteActivity,
      listExpenses, saveExpense, deleteExpense,
      listProgress, saveProgress, deleteProgress,
      getReport, saveReport,
      listFiles, uploadProjectFile, deleteProjectFile,
      listUsers, saveUser, resetUserPassword,
      listDepartments, saveDepartment, deleteDepartment,
      getSettings, saveSettings, createBackup, listLogs,
      initializeSystem
    };

    if (typeof actions[action] !== 'function') {
      return ContentService.createTextOutput(JSON.stringify(fail_('Unknown action: ' + action)))
        .setMimeType(ContentService.MimeType.JSON);
    }

    const result = actions[action].apply(null, args);
    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify(fail_(safeMessage_(err))))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// =====================================================
// INITIAL SETUP / SHEET INITIALIZATION
// =====================================================
function initializeSystem() {
  try {
    initializeSystem_();
    return ok_('ກຽມລະບົບຮຽບຮ້ອຍ', getPublicSettings_());
  } catch (e) { return fail_(safeMessage_(e)); }
}

function initializeSystem_() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  Object.keys(SHEETS).forEach(name => ensureSheet_(ss, name, SHEETS[name]));
  ensureDefaultSettings_();
  ensureDefaultDepartments_();
}

function ensureSheet_(ss, name, headers) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    sh.setFrozenRows(1);
    sh.getRange(1,1,1,headers.length).setFontWeight('bold').setBackground('#f4f1ff');
    sh.autoResizeColumns(1, headers.length);
  } else {
    const current = sh.getRange(1,1,1,Math.max(sh.getLastColumn(), headers.length)).getValues()[0];
    headers.forEach((h,i)=>{ if (!current[i]) sh.getRange(1,i+1).setValue(h); });
  }
  return sh;
}

function ensureDefaultSettings_() {
  const existing = getSettingsMap_();
  const rows = [];
  Object.keys(DEFAULT_SETTINGS).forEach(k => {
    if (existing[k] === undefined || existing[k] === '') rows.push([k, DEFAULT_SETTINGS[k], now_(), 'system']);
  });
  if (rows.length) appendRows_('Settings', rows);
}

function ensureDefaultDepartments_() {
  const rows = getRows_('Departments');
  if (rows.length) return;
  const now = now_();
  [
    ['ບໍລິຫານວິຊາການ','ວຽກງານຫຼັກສູດ ແລະ ການຮຽນຮູ້'],
    ['ບໍລິຫານງົບປະມານ','ວຽກງານແຜນ, ງົບປະມານ, ການເງິນ'],
    ['ບໍລິຫານບຸກຄະລາກອນ','ວຽກງານບຸກຄະລາກອນ'],
    ['ບໍລິຫານທົ່ວໄປ','ວຽກງານທົ່ວໄປ ແລະ ອາຄານສະຖານທີ່']
  ].forEach(x => appendRows_('Departments', [[uid_('D'), x[0], x[1], 'ນຳໃຊ້ງານ', now, now]]));
}

// =====================================================
// AUTHENTICATION / SESSION
// =====================================================
function login(username, password) {
  try {
    username = clean_(username); password = String(password || '').trim();
    if (!username || !password) return fail_('ກະລຸນາປ້ອນຊື່ຜູ້ໃຊ້ ແລະ ລະຫັດຜ່ານ');
    let user = null;
    const adminPass = String(getAdminPassword_()).trim();
    const defaultPass = String(ADMIN_PASSWORD).trim();
    if (username.toLowerCase() === 'admin' && (password === adminPass || password === defaultPass)) {
      user = { userId:'ADMIN', username:'admin', fullName:'ຜູ້ດູແລລະບົບ', position:'ຜູ້ດູແລລະບົບ', departmentId:'', role:'admin', status:'ນຳໃຊ້ງານ' };
    } else {
      const found = getRows_('Users').find(r => String(r.username).toLowerCase() === username.toLowerCase());
      if (!found || String(found.password).trim() !== password || String(found.status) !== 'ນຳໃຊ້ງານ') return fail_('ຊື່ຜູ້ໃຊ້ ຫຼື ລະຫັດຜ່ານບໍ່ຖືກຕ້ອງ');
      user = sanitizeUser_(found);
    }
    const token = Utilities.getUuid();
    CacheService.getScriptCache().put('SESSION_' + token, JSON.stringify(user), SESSION_TTL);
    writeLog_(user,'Login','ເຂົ້າສູ່ລະບົບ','-');
    return ok_('ເຂົ້າສູ່ລະບົບສຳເລັດ', { token, user, settings:getPublicSettings_() });
  } catch(e) { return fail_(safeMessage_(e)); }
}

function logout(token) {
  try {
    const user = getSession_(token, false);
    if (user) writeLog_(user,'Logout','ອອກຈາກລະບົບ','-');
    CacheService.getScriptCache().remove('SESSION_' + token);
    return ok_('ອອກຈາກລະບົບແລ້ວ');
  } catch(e) { return fail_('ອອກຈາກລະບົບບໍ່ສຳເລັດ'); }
}

function getSession_(token, throwIfMissing) {
  const raw = token ? CacheService.getScriptCache().get('SESSION_' + token) : null;
  if (!raw) {
    if (throwIfMissing !== false) throw new Error('SESSION_EXPIRED');
    return null;
  }
  CacheService.getScriptCache().put('SESSION_' + token, raw, SESSION_TTL);
  return JSON.parse(raw);
}

function requireRole_(token, roles) {
  const user = getSession_(token, true);
  if (roles && roles.length && roles.indexOf(user.role) < 0) throw new Error('ທ່ານບໍ່ມີສິດນຳໃຊ້ສ່ວນນີ້');
  return user;
}

// =====================================================
// OPTIMIZATION: BULK DATA PROCESSING
// =====================================================
function bulkCalculateBudgets_(projectIds) {
  const acts = getRows_('Activities');
  const exps = getRows_('Expenses');
  const budgetMap = {};
  
  projectIds.forEach(id => { budgetMap[id] = { budget: 0, activityBudget: 0, spent: 0, remaining: 0, percent: 0 }; });
  
  const projects = getRows_('Projects').filter(p => projectIds.includes(p.projectId));
  projects.forEach(p => budgetMap[p.projectId].budget = num_(p.budget));
  
  acts.forEach(a => { if (budgetMap[a.projectId]) budgetMap[a.projectId].activityBudget += num_(a.budget); });
  exps.forEach(e => { if (budgetMap[e.projectId] && String(e.status) !== 'ຍົກເລີກ') budgetMap[e.projectId].spent += num_(e.amount); });
  
  projectIds.forEach(id => {
    const b = budgetMap[id];
    b.remaining = Math.max(0, b.budget - b.spent);
    b.percent = b.budget ? Math.round((b.spent * 10000) / b.budget) / 100 : 0;
  });
  
  return budgetMap;
}

function enrichProjectsBulk_(projects) {
  if (!projects.length) return [];
  const projectIds = projects.map(p => p.projectId);
  const budgetMap = bulkCalculateBudgets_(projectIds);
  const deps = getRows_('Departments');
  const users = getRows_('Users');
  
  return projects.map(p => {
    const x = Object.assign({}, p);
    const dep = deps.find(d => d.departmentId === p.departmentId);
    x.departmentName = dep ? dep.name : '';
    const u = users.find(u => u.userId === p.ownerUserId);
    x.ownerName = u ? u.fullName : (p.ownerUserId === 'ADMIN' ? 'ຜູ້ດູແລລະບົບ' : p.ownerUserId);
    x.budgetSummary = budgetMap[p.projectId] || {budget:0, activityBudget:0, spent:0, remaining:0, percent:0};
    return x;
  });
}

function calculateProjectBudget_(projectId) {
  const map = bulkCalculateBudgets_([projectId]);
  return map[projectId] || {budget:0, activityBudget:0, spent:0, remaining:0, percent:0};
}

function enrichProject_(p) {
  if (!p) return p;
  const enriched = enrichProjectsBulk_([p]);
  return enriched.length ? enriched[0] : p;
}

// =====================================================
// BOOTSTRAP / DASHBOARD
// =====================================================
function getAppBootstrap(token) {
  try {
    const user = requireRole_(token, ['admin','head','teacher']);
    return ok_('ໂຫຼດຂໍ້ມູນສຳເລັດ', {
      user,
      settings: getPublicSettings_(),
      departments: scopedDepartments_(user),
      dashboard: dashboardData_(user)
    });
  } catch(e) { return fail_(safeMessage_(e)); }
}

function getDashboard(token, filters) {
  try { return ok_('ສຳເລັດ', dashboardData_(requireRole_(token,['admin','head','teacher']), filters || {})); }
  catch(e) { return fail_(safeMessage_(e)); }
}

function dashboardData_(user, filters) {
  filters = filters || {};
  let projects = scopeProjects_(user, getRows_('Projects'));
  if (filters.fiscalYear) projects = projects.filter(p=>String(p.fiscalYear)===String(filters.fiscalYear));
  if (filters.departmentId) projects = projects.filter(p=>String(p.departmentId)===String(filters.departmentId));
  if (filters.status) projects = projects.filter(p=>String(p.status)===String(filters.status));
  
  const ids = projects.map(p=>p.projectId);
  const expenses = getRows_('Expenses').filter(e=>ids.includes(e.projectId));
  const totalBudget = projects.reduce((s,p)=>s+num_(p.budget),0);
  const spent = expenses.filter(e=>String(e.status)!=='ຍົກເລີກ').reduce((s,e)=>s+num_(e.amount),0);
  const summary = {
    totalProjects:projects.length,
    ongoing:projects.filter(p=>p.status==='ກຳລັງດຳເນີນງານ').length,
    completed:projects.filter(p=>p.status==='ສຳເລັດ').length,
    delayed:projects.filter(p=>p.status==='ຫຼ້າຊ້າ').length,
    totalBudget, spent, remaining:Math.max(0,totalBudget-spent), spendingPercent: totalBudget ? Math.round(spent*10000/totalBudget)/100 : 0
  };
  
  const byDepartment = groupSum_(projects, 'departmentId', ()=>1);
  const byStatus = groupSum_(projects, 'status', ()=>1);
  const budgetByDepartment = groupSum_(projects, 'departmentId', p=>num_(p.budget));
  
  const enrichedProjects = enrichProjectsBulk_(projects);
  return {
    summary,
    byDepartment: labelDepartmentGroups_(byDepartment),
    byStatus,
    budgetByDepartment: labelDepartmentGroups_(budgetByDepartment),
    latestProjects:enrichedProjects.slice().sort((a,b)=>String(b.createdAt).localeCompare(String(a.createdAt))).slice(0,6),
    delayedProjects:enrichedProjects.filter(p=>p.status==='ຫຼ້າຊ້າ').slice(0,6),
    recentExpenses:expenses.slice().sort((a,b)=>String(b.createdAt).localeCompare(String(a.createdAt))).slice(0,6).map(enrichExpense_)
  };
}

// =====================================================
// USER MANAGEMENT
// =====================================================
function listUsers(token) {
  try { requireRole_(token,['admin']); return ok_('ສຳເລັດ', getRows_('Users').map(sanitizeUser_)); }
  catch(e){ return fail_(safeMessage_(e)); }
}

function saveUser(token, data) {
  try {
    const actor = requireRole_(token,['admin']); data = data || {};
    const username = clean_(data.username).toLowerCase();
    if (!username || !clean_(data.fullName)) throw new Error('ກະລຸນາປ້ອນຊື່ຜູ້ໃຊ້ ແລະ ຊື່-ນາມສະກຸນ');
    const rows = getRows_('Users');
    const duplicate = rows.find(r=>String(r.username).toLowerCase()===username && r.userId!==data.userId);
    if (duplicate || username==='admin') throw new Error('ຊື່ຜູ້ໃຊ້ນີ້ຖືກນຳໃຊ້ແລ້ວ');
    const now = now_();
    if (data.userId) {
      const old = rows.find(r=>r.userId===data.userId); if (!old) throw new Error('ບໍ່ພົບຜູ້ໃຊ້ງານ');
      const record = Object.assign({}, old, {
        username, fullName:clean_(data.fullName), position:clean_(data.position), departmentId:clean_(data.departmentId),
        role:['teacher','head','admin'].indexOf(data.role)>=0?data.role:'teacher', status:data.status||'ນຳໃຊ້ງານ', updatedAt:now
      });
      if (data.password) record.password = String(data.password);
      updateById_('Users','userId',record.userId,record);
      writeLog_(actor,'ແກ້ໄຂຜູ້ໃຊ້ງານ',record.fullName,record.userId);
      return ok_('ບັນທຶກຜູ້ໃຊ້ງານແລ້ວ', sanitizeUser_(record));
    }
    if (!String(data.password||'')) throw new Error('ກະລຸນາກຳນົດລະຫັດຜ່ານ');
    const record = {userId:uid_('U'),username,password:String(data.password),fullName:clean_(data.fullName),position:clean_(data.position),departmentId:clean_(data.departmentId),role:data.role||'teacher',status:data.status||'ນຳໃຊ້ງານ',createdAt:now,updatedAt:now};
    appendObject_('Users', record); writeLog_(actor,'ເພີ່ມຜູ້ໃຊ້ງານ',record.fullName,record.userId);
    return ok_('ເພີ່ມຜູ້ໃຊ້ງານແລ້ວ',sanitizeUser_(record));
  } catch(e){ return fail_(safeMessage_(e)); }
}

function setUserStatus(token, userId, status) {
  try { const a=requireRole_(token,['admin']); const u=findById_('Users','userId',userId); if(!u)throw new Error('ບໍ່ພົບຜູ້ໃຊ້ງານ'); u.status=status;u.updatedAt=now_();updateById_('Users','userId',userId,u);writeLog_(a,'ປ່ຽນສະຖານະຜູ້ໃຊ້',status,userId);return ok_('ສຳເລັດ'); }
  catch(e){return fail_(safeMessage_(e));}
}

function resetUserPassword(token, userId, newPassword) {
  try { const a=requireRole_(token,['admin']); if(!newPassword)throw new Error('ກະລຸນາປ້ອນລະຫັດຜ່ານໃໝ່');const u=findById_('Users','userId',userId);if(!u)throw new Error('ບໍ່ພົບຜູ້ໃຊ້ງານ');u.password=String(newPassword);u.updatedAt=now_();updateById_('Users','userId',userId,u);writeLog_(a,'Reset Password',u.username,userId);return ok_('ຣີເຊັດລະຫັດຜ່ານແລ້ວ'); }
  catch(e){return fail_(safeMessage_(e));}
}

// =====================================================
// DEPARTMENT MANAGEMENT
// =====================================================
function listDepartments(token) {
  try { const user=requireRole_(token,['admin','head','teacher']); return ok_('ສຳເລັດ', scopedDepartments_(user)); }
  catch(e){return fail_(safeMessage_(e));}
}

function saveDepartment(token, data) {
  try {
    const a=requireRole_(token,['admin']); if(!clean_(data.name))throw new Error('ກະລຸນາປ້ອນຊື່ພະແນກ/ຝ່າຍ'); const now=now_();
    if(data.departmentId){const r=findById_('Departments','departmentId',data.departmentId);if(!r)throw new Error('ບໍ່ພົບພະແນກ/ຝ່າຍ');Object.assign(r,{name:clean_(data.name),description:clean_(data.description),status:data.status||'ນຳໃຊ້ງານ',updatedAt:now});updateById_('Departments','departmentId',r.departmentId,r);writeLog_(a,'ແກ້ໄຂພະແນກ/ຝ່າຍ',r.name,r.departmentId);return ok_('ບັນທຶກແລ້ວ',r);}
    const r={departmentId:uid_('D'),name:clean_(data.name),description:clean_(data.description),status:data.status||'ນຳໃຊ້ງານ',createdAt:now,updatedAt:now};appendObject_('Departments',r);writeLog_(a,'ເພີ່ມພະແນກ/ຝ່າຍ',r.name,r.departmentId);return ok_('ເພີ່ມພະແນກ/ຝ່າຍແລ້ວ',r);
  }catch(e){return fail_(safeMessage_(e));}
}

function deleteDepartment(token, departmentId) {
  try { const a=requireRole_(token,['admin']); if(getRows_('Projects').some(p=>p.departmentId===departmentId))throw new Error('ບໍ່ສາມາດລຶບພະແນກທີ່ມີໂຄງການອ້າງອີງຢູ່');const r=findById_('Departments','departmentId',departmentId);deleteById_('Departments','departmentId',departmentId);writeLog_(a,'ລຶບພະແນກ/ຝ່າຍ',r?r.name:'',departmentId);return ok_('ລຶບພະແນກ/ຝ່າຍແລ້ວ'); }
  catch(e){return fail_(safeMessage_(e));}
}

// =====================================================
// PROJECT CRUD
// =====================================================
function listProjects(token, filters) {
  try {
    const user=requireRole_(token,['admin','head','teacher']); let rows=scopeProjects_(user,getRows_('Projects')); filters=filters||{};
    if(filters.search){const q=String(filters.search).toLowerCase();rows=rows.filter(p=>[p.projectCode,p.projectName,p.ownerUserId,p.coOwners].join(' ').toLowerCase().indexOf(q)>=0);}
    ['fiscalYear','departmentId','status'].forEach(k=>{if(filters[k])rows=rows.filter(p=>String(p[k])===String(filters[k]));});
    if(filters.ownerUserId)rows=rows.filter(p=>String(p.ownerUserId)===String(filters.ownerUserId));
    if(filters.startDate)rows=rows.filter(p=>String(p.startDate)>=String(filters.startDate));
    if(filters.endDate)rows=rows.filter(p=>String(p.endDate)<=String(filters.endDate));
    rows=rows.sort((a,b)=>String(b.updatedAt).localeCompare(String(a.updatedAt)));
    return ok_('ສຳເລັດ', enrichProjectsBulk_(rows));
  }catch(e){return fail_(safeMessage_(e));}
}

function getProjectDetail(token, projectId) {
  try {
    const user=requireRole_(token,['admin','head','teacher']); const p=findById_('Projects','projectId',projectId); assertProjectAccess_(user,p,false);
    const acts=getRows_('Activities').filter(x=>x.projectId===projectId);
    const exps=getRows_('Expenses').filter(x=>x.projectId===projectId).map(enrichExpense_);
    const prog=getRows_('Progress').filter(x=>x.projectId===projectId).sort((a,b)=>String(b.reportDate).localeCompare(String(a.reportDate)));
    const files=getRows_('Files').filter(x=>x.projectId===projectId);
    const report=getRows_('Reports').find(x=>x.projectId===projectId)||null;
    return ok_('ສຳເລັດ',{project:enrichProject_(p),activities:acts,expenses:exps,progress:prog,files,report,budget:calculateProjectBudget_(projectId)});
  }catch(e){return fail_(safeMessage_(e));}
}

function saveProject(token, data) {
  try {
    const user=requireRole_(token,['admin','head','teacher']); data=data||{}; validateProject_(data); const now=now_();
    if(data.projectId){
      const old=findById_('Projects','projectId',data.projectId);assertProjectAccess_(user,old,true);
      const record=buildProjectRecord_(data,old,user,now); updateById_('Projects','projectId',record.projectId,record);writeLog_(user,'ແກ້ໄຂໂຄງການ',record.projectName,record.projectId);return ok_('ບັນທຶກໂຄງການແລ້ວ',enrichProject_(record));
    }
    const record=buildProjectRecord_(data,null,user,now);record.projectId=uid_('P');record.createdAt=now;record.createdBy=user.userId;appendObject_('Projects',record);writeLog_(user,'ເພີ່ມໂຄງການ',record.projectName,record.projectId);return ok_('ເພີ່ມໂຄງການແລ້ວ',enrichProject_(record));
  }catch(e){return fail_(safeMessage_(e));}
}

function deleteProject(token, projectId) {
  try {
    const user=requireRole_(token,['admin','head','teacher']); const p=findById_('Projects','projectId',projectId);assertProjectAccess_(user,p,true);
    ['Activities','Expenses','Progress','Reports','Files'].forEach(s=>deleteWhere_(s,'projectId',projectId)); deleteById_('Projects','projectId',projectId);writeLog_(user,'ລຶບໂຄງການ',p.projectName,projectId);return ok_('ລຶບໂຄງການແລ້ວ');
  }catch(e){return fail_(safeMessage_(e));}
}

function copyProject(token, projectId) {
  try { const user=requireRole_(token,['admin','head','teacher']);const p=findById_('Projects','projectId',projectId);assertProjectAccess_(user,p,false);const c=Object.assign({},p,{projectId:uid_('P'),projectCode:p.projectCode+'-COPY',projectName:p.projectName+' (ສຳເນົາ)',status:'ກຳລັງດຳເນີນງານ',progressPercent:0,createdAt:now_(),updatedAt:now_(),createdBy:user.userId,updatedBy:user.userId});appendObject_('Projects',c);writeLog_(user,'ສຳເນົາໂຄງການ',c.projectName,c.projectId);return ok_('ສຳເນົາໂຄງການແລ້ວ',enrichProject_(c)); }
  catch(e){return fail_(safeMessage_(e));}
}

function buildProjectRecord_(d,old,user,now){
  const r=old?Object.assign({},old):{};
  Object.assign(r,{projectCode:clean_(d.projectCode),projectName:clean_(d.projectName),fiscalYear:clean_(d.fiscalYear),planName:clean_(d.planName),departmentId:clean_(d.departmentId||user.departmentId),ownerUserId:clean_(d.ownerUserId||user.userId),coOwners:clean_(d.coOwners),rationale:clean_(d.rationale),objectives:clean_(d.objectives),quantitativeTarget:clean_(d.quantitativeTarget),qualitativeTarget:clean_(d.qualitativeTarget),targetGroup:clean_(d.targetGroup),location:clean_(d.location),startDate:clean_(d.startDate),endDate:clean_(d.endDate),budget:num_(d.budget),budgetSource:clean_(d.budgetSource),successIndicators:clean_(d.successIndicators),expectedResults:clean_(d.expectedResults),status:d.status||'ກຳລັງດຳເນີນງານ',progressPercent:clamp_(num_(d.progressPercent),0,100),note:clean_(d.note),updatedAt:now,updatedBy:user.userId});
  if(user.role==='teacher'){r.ownerUserId=user.userId;r.departmentId=user.departmentId||r.departmentId;}
  if(user.role==='head'){r.departmentId=user.departmentId||r.departmentId;}
  return r;
}

function validateProject_(d){if(!clean_(d.projectCode)||!clean_(d.projectName))throw new Error('ກະລຸນາປ້ອນລະຫັດ ແລະ ຊື່ໂຄງການ');if(num_(d.budget)<0)throw new Error('ງົບປະມານຕ້ອງບໍ່ຕິດລົບ');if(d.startDate&&d.endDate&&String(d.endDate)<String(d.startDate))throw new Error('ວັນທີສິ້ນສຸດຕ້ອງບໍ່ນ້ອຍກວ່າວັນທີເລີ່ມຕົ້ນ');}

// =====================================================
// ACTIVITY CRUD
// =====================================================
function listActivities(token, projectId) {
  try {
    const u=requireRole_(token,['admin','head','teacher']);
    if (projectId) { const p=findById_('Projects','projectId',projectId);assertProjectAccess_(u,p,false);return ok_('ສຳເລັດ',getRows_('Activities').filter(x=>x.projectId===projectId).map(a=>Object.assign({},a,{projectName:p.projectName}))); }
    const projects=scopeProjects_(u,getRows_('Projects')), names={}; projects.forEach(p=>names[p.projectId]=p.projectName);
    return ok_('ສຳເລັດ',getRows_('Activities').filter(a=>names[a.projectId]).map(a=>Object.assign({},a,{projectName:names[a.projectId]})));
  } catch(e){return fail_(safeMessage_(e));}
}
function saveActivity(token, d) {
  try { const u=requireRole_(token,['admin','head','teacher']);const p=findById_('Projects','projectId',d.projectId);assertProjectAccess_(u,p,true);if(!clean_(d.activityName))throw new Error('ກະລຸນາປ້ອນຊື່ກິດຈະກຳ');if(d.startDate&&d.endDate&&String(d.endDate)<String(d.startDate))throw new Error('ວັນທີສິ້ນສຸດຕ້ອງບໍ່ນ້ອຍກວ່າວັນທີເລີ່ມຕົ້ນ');const other=getRows_('Activities').filter(x=>x.projectId===d.projectId&&x.activityId!==d.activityId).reduce((s,x)=>s+num_(x.budget),0);if(other+num_(d.budget)>num_(p.budget))throw new Error('ງົບກິດຈະກຳລວມເກີນງົບປະມານໂຄງການ');const now=now_();let r=d.activityId?findById_('Activities','activityId',d.activityId):null;r=r||{activityId:uid_('A'),projectId:d.projectId,createdAt:now};Object.assign(r,{activityName:clean_(d.activityName),description:clean_(d.description),owner:clean_(d.owner),startDate:clean_(d.startDate),endDate:clean_(d.endDate),budget:num_(d.budget),result:clean_(d.result),successPercent:clamp_(num_(d.successPercent),0,100),status:d.status||'ກຳລັງດຳເນີນງານ',note:clean_(d.note),updatedAt:now});d.activityId?updateById_('Activities','activityId',r.activityId,r):appendObject_('Activities',r);writeLog_(u,d.activityId?'ແກ້ໄຂກິດຈະກຳ':'ເພີ່ມກິດຈະກຳ',r.activityName,r.activityId);return ok_('ບັນທຶກກິດຈະກຳແລ້ວ',r); }
  catch(e){return fail_(safeMessage_(e));}
}
function deleteActivity(token, activityId) {try{const u=requireRole_(token,['admin','head','teacher']);const a=findById_('Activities','activityId',activityId);const p=a&&findById_('Projects','projectId',a.projectId);assertProjectAccess_(u,p,true);deleteById_('Activities','activityId',activityId);writeLog_(u,'ລຶບກິດຈະກຳ',a?a.activityName:'',activityId);return ok_('ລຶບກິດຈະກຳແລ້ວ');}catch(e){return fail_(safeMessage_(e));}}

// =====================================================
// EXPENSE CRUD / BUDGET
// =====================================================
function listExpenses(token, filters){
  try{const u=requireRole_(token,['admin','head','teacher']);let ex=getRows_('Expenses');const allowed=scopeProjects_(u,getRows_('Projects')).map(p=>p.projectId);ex=ex.filter(e=>allowed.indexOf(e.projectId)>=0);filters=filters||{};if(filters.projectId)ex=ex.filter(e=>e.projectId===filters.projectId);if(filters.expenseType)ex=ex.filter(e=>e.expenseType===filters.expenseType);if(filters.startDate)ex=ex.filter(e=>String(e.expenseDate)>=String(filters.startDate));if(filters.endDate)ex=ex.filter(e=>String(e.expenseDate)<=String(filters.endDate));return ok_('ສຳເລັດ',ex.map(enrichExpense_));}catch(e){return fail_(safeMessage_(e));}
}
function saveExpense(token,d){
  try{const u=requireRole_(token,['admin','head','teacher']);const p=findById_('Projects','projectId',d.projectId);assertProjectAccess_(u,p,true);const amount=num_(d.amount);if(amount<=0)throw new Error('ຈຳນວນເງິນຕ້ອງຫຼາຍກວ່າ 0');const previous=d.expenseId?num_((findById_('Expenses','expenseId',d.expenseId)||{}).amount):0;const budget=calculateProjectBudget_(d.projectId);if(amount>budget.remaining+previous)throw new Error('ຍອດເບີກຈ່າຍເກີນງົບປະມານຄົງເຫຼືອ');if(d.activityId){const a=findById_('Activities','activityId',d.activityId);if(!a||a.projectId!==d.projectId)throw new Error('ກິດຈະກຳບໍ່ກົງກັບໂຄງການ');const activitySpent=getRows_('Expenses').filter(e=>e.projectId===d.projectId&&e.activityId===d.activityId&&e.expenseId!==d.expenseId&&e.status!=='ຍົກເລີກ').reduce((s,e)=>s+num_(e.amount),0);if(activitySpent+amount>num_(a.budget))throw new Error('ຍອດເບີກຈ່າຍເກີນງົບປະມານກິດຈະກຳ');}
    const now=now_();let r=d.expenseId?findById_('Expenses','expenseId',d.expenseId):null;r=r||{expenseId:uid_('E'),projectId:d.projectId,createdAt:now,createdBy:u.userId};Object.assign(r,{projectId:d.projectId,activityId:clean_(d.activityId),itemNo:clean_(d.itemNo),expenseDate:clean_(d.expenseDate),description:clean_(d.description),expenseType:clean_(d.expenseType),amount,documentNo:clean_(d.documentNo),requester:clean_(d.requester),status:d.status||'ລໍຖ້າກວດສອບ',note:clean_(d.note),attachmentFileId:clean_(d.attachmentFileId)});d.expenseId?updateById_('Expenses','expenseId',r.expenseId,r):appendObject_('Expenses',r);writeLog_(u,d.expenseId?'ແກ້ໄຂລາຍການເບີກຈ່າຍ':'ເພີ່ມລາຍການເບີກຈ່າຍ',r.description,r.expenseId);return ok_('ບັນທຶກລາຍການເບີກຈ່າຍແລ້ວ',{expense:enrichExpense_(r),budget:calculateProjectBudget_(d.projectId)});}
  catch(e){return fail_(safeMessage_(e));}
}
function deleteExpense(token,expenseId){try{const u=requireRole_(token,['admin','head','teacher']);const e=findById_('Expenses','expenseId',expenseId);const p=e&&findById_('Projects','projectId',e.projectId);assertProjectAccess_(u,p,true);deleteById_('Expenses','expenseId',expenseId);writeLog_(u,'ລຶບລາຍການເບີກຈ່າຍ',e?e.description:'',expenseId);return ok_('ລຶບລາຍການແລ້ວ',calculateProjectBudget_(e.projectId));}catch(err){return fail_(safeMessage_(err));}}
function getProjectBudget(token,projectId){try{const u=requireRole_(token,['admin','head','teacher']);const p=findById_('Projects','projectId',projectId);assertProjectAccess_(u,p,false);return ok_('ສຳເລັດ',calculateProjectBudget_(projectId));}catch(e){return fail_(safeMessage_(e));}}

// =====================================================
// PROGRESS CRUD
// =====================================================
function listProgress(token,projectId){try{const u=requireRole_(token,['admin','head','teacher']);if(projectId){const p=findById_('Projects','projectId',projectId);assertProjectAccess_(u,p,false);return ok_('ສຳເລັດ',getRows_('Progress').filter(x=>x.projectId===projectId).map(g=>Object.assign({},g,{projectName:p.projectName})).sort((a,b)=>String(b.reportDate).localeCompare(String(a.reportDate))));}const projects=scopeProjects_(u,getRows_('Projects')),names={};projects.forEach(p=>names[p.projectId]=p.projectName);return ok_('ສຳເລັດ',getRows_('Progress').filter(g=>names[g.projectId]).map(g=>Object.assign({},g,{projectName:names[g.projectId]})).sort((a,b)=>String(b.reportDate).localeCompare(String(a.reportDate))));}catch(e){return fail_(safeMessage_(e));}}
function saveProgress(token,d){try{const u=requireRole_(token,['admin','head','teacher']);const p=findById_('Projects','projectId',d.projectId);assertProjectAccess_(u,p,true);const now=now_();let r=d.progressId?findById_('Progress','progressId',d.progressId):null;r=r||{progressId:uid_('G'),projectId:d.projectId,createdAt:now,createdBy:u.userId};Object.assign(r,{reportDate:clean_(d.reportDate),detail:clean_(d.detail),output:clean_(d.output),problem:clean_(d.problem),solution:clean_(d.solution),progressPercent:clamp_(num_(d.progressPercent),0,100),imageFileIds:clean_(d.imageFileIds),attachmentFileIds:clean_(d.attachmentFileIds)});d.progressId?updateById_('Progress','progressId',r.progressId,r):appendObject_('Progress',r);p.progressPercent=r.progressPercent;p.updatedAt=now;p.updatedBy=u.userId;updateById_('Projects','projectId',p.projectId,p);writeLog_(u,'ບັນທຶກຄວາມຄືບໜ້າ',r.progressPercent+'%',r.progressId);return ok_('ບັນທຶກຄວາມຄືບໜ້າແລ້ວ',r);}catch(e){return fail_(safeMessage_(e));}}
function deleteProgress(token,progressId){try{const u=requireRole_(token,['admin','head','teacher']);const r=findById_('Progress','progressId',progressId);const p=r&&findById_('Projects','projectId',r.projectId);assertProjectAccess_(u,p,true);deleteById_('Progress','progressId',progressId);writeLog_(u,'ລຶບຄວາມຄືບໜ້າ','',progressId);return ok_('ລຶບແລ້ວ');}catch(e){return fail_(safeMessage_(e));}}

// =====================================================
// REPORT CRUD
// =====================================================
function getReport(token,projectId){try{const u=requireRole_(token,['admin','head','teacher']);const p=findById_('Projects','projectId',projectId);assertProjectAccess_(u,p,false);return ok_('ສຳເລັດ',{project:enrichProject_(p),report:getRows_('Reports').find(x=>x.projectId===projectId)||null,activities:getRows_('Activities').filter(x=>x.projectId===projectId),budget:calculateProjectBudget_(projectId)});}catch(e){return fail_(safeMessage_(e));}}
function saveReport(token,d){try{const u=requireRole_(token,['admin','head','teacher']);const p=findById_('Projects','projectId',d.projectId);assertProjectAccess_(u,p,true);const now=now_();let r=getRows_('Reports').find(x=>x.projectId===d.projectId);const exists=!!r;r=r||{reportId:uid_('R'),projectId:d.projectId,createdAt:now,createdBy:u.userId};Object.assign(r,{resultSummary:clean_(d.resultSummary),indicatorResult:clean_(d.indicatorResult),problem:clean_(d.problem),suggestion:clean_(d.suggestion),activityImages:clean_(d.activityImages),attachments:clean_(d.attachments),updatedAt:now,updatedBy:u.userId});exists?updateById_('Reports','reportId',r.reportId,r):appendObject_('Reports',r);writeLog_(u,'ບັນທຶກລາຍງານຜົນ',p.projectName,r.reportId);return ok_('ບັນທຶກລາຍງານແລ້ວ',r);}catch(e){return fail_(safeMessage_(e));}}

// =====================================================
// FILE MANAGEMENT / GOOGLE DRIVE
// =====================================================
function listFiles(token,projectId){try{const u=requireRole_(token,['admin','head','teacher']);if(projectId){const p=findById_('Projects','projectId',projectId);assertProjectAccess_(u,p,false);return ok_('ສຳເລັດ',getRows_('Files').filter(x=>x.projectId===projectId).map(f=>Object.assign({},f,{projectName:p.projectName})));}const projects=scopeProjects_(u,getRows_('Projects')),names={};projects.forEach(p=>names[p.projectId]=p.projectName);return ok_('ສຳເລັດ',getRows_('Files').filter(f=>names[f.projectId]).map(f=>Object.assign({},f,{projectName:names[f.projectId]})));}catch(e){return fail_(safeMessage_(e));}}
function uploadProjectFile(token,payload){
  try{const u=requireRole_(token,['admin','head','teacher']);const p=findById_('Projects','projectId',payload.projectId);assertProjectAccess_(u,p,true);if(!payload.base64||!payload.fileName)throw new Error('ບໍ່ພົບເອກະສານ');const root=DriveApp.getFolderById(FOLDER_ID);const dep=getDepartmentName_(p.departmentId)||'ບໍ່ລະບຸພະແນກ';const fy=p.fiscalYear||'ບໍ່ລະບຸປີ';const folder=ensureDriveFolderPath_(root,[safeFolderName_(dep),safeFolderName_(fy),safeFolderName_(p.projectName)]);const bytes=Utilities.base64Decode(payload.base64);const blob=Utilities.newBlob(bytes,payload.mimeType||MimeType.PLAIN_TEXT,payload.fileName);const file=folder.createFile(blob);const r={fileId:uid_('F'),projectId:p.projectId,activityId:clean_(payload.activityId),fileName:file.getName(),mimeType:file.getMimeType(),driveFileId:file.getId(),driveUrl:file.getUrl(),folderPath:[dep,fy,p.projectName].join(' > '),uploadedAt:now_(),uploadedBy:u.userId};appendObject_('Files',r);writeLog_(u,'ອັບໂຫຼດເອກະສານ',r.fileName,r.fileId);return ok_('ອັບໂຫຼດເອກະສານແລ້ວ',r);}catch(e){return fail_(safeMessage_(e));}
}
function deleteProjectFile(token,fileId){try{const u=requireRole_(token,['admin','head','teacher']);const r=findById_('Files','fileId',fileId);if(!r)throw new Error('ບໍ່ພົບເອກະສານ');const p=findById_('Projects','projectId',r.projectId);assertProjectAccess_(u,p,true);try{DriveApp.getFileById(r.driveFileId).setTrashed(true);}catch(ignore){}deleteById_('Files','fileId',fileId);writeLog_(u,'ລຶບເອກະສານ',r.fileName,fileId);return ok_('ລຶບເອກະສານແລ້ວ');}catch(e){return fail_(safeMessage_(e));}}
function ensureDriveFolderPath_(root,names){let f=root;names.forEach(n=>{const it=f.getFoldersByName(n);f=it.hasNext()?it.next():f.createFolder(n);});return f;}

// =====================================================
// SETTINGS
// =====================================================
function getSettings(token){try{requireRole_(token,['admin']);const s=getSettingsMap_();delete s.adminPassword;return ok_('ສຳເລັດ',s);}catch(e){return fail_(safeMessage_(e));}}
function saveSettings(token,data){try{const u=requireRole_(token,['admin']);const allowed=['systemName','schoolName','address','affiliation','logoUrl','bannerUrl','directorName','defaultFiscalYear','footerText'];allowed.forEach(k=>{if(data[k]!==undefined)setSetting_(k,String(data[k]),u.username);});if(data.adminPassword)setSetting_('adminPassword',String(data.adminPassword),u.username);writeLog_(u,'ປ່ຽນ Settings','ແກ້ໄຂການຕັ້ງຄ່າລະບົບ','Settings');return ok_('ບັນທຶກການຕັ້ງຄ່າແລ້ວ',getPublicSettings_());}catch(e){return fail_(safeMessage_(e));}}
function getAdminPassword_(){try{const s=getSettingsMap_();return s.adminPassword?String(s.adminPassword).trim():String(ADMIN_PASSWORD).trim();}catch(e){return String(ADMIN_PASSWORD).trim();}}
function getPublicSettings_(){const s=Object.assign({},DEFAULT_SETTINGS,getSettingsMap_());delete s.adminPassword;return s;}
function getSettingsMap_(){const out={};getRows_('Settings').forEach(r=>{out[String(r.key)]=r.value;});return out;}
function setSetting_(key,value,by){const sh=getSheet_('Settings');const vals=sh.getDataRange().getValues();for(let i=1;i<vals.length;i++){if(String(vals[i][0])===key){sh.getRange(i+1,1,1,4).setValues([[key,value,now_(),by]]);return;}}sh.appendRow([key,value,now_(),by]);}

// =====================================================
// BACKUP
// =====================================================
function createBackup(token){
  try{const u=requireRole_(token,['admin']);const blobs=[];Object.keys(SHEETS).forEach(name=>{const sh=getSheet_(name);const csv=sh.getDataRange().getDisplayValues().map(r=>r.map(csvEscape_).join(',')).join('\r\n');blobs.push(Utilities.newBlob('\ufeff'+csv,'text/csv',name+'.csv'));});const zip=Utilities.zip(blobs,'ProjectSystem_Backup_'+Utilities.formatDate(new Date(),APP_TZ,'yyyyMMdd_HHmmss')+'.zip');const root=DriveApp.getFolderById(FOLDER_ID);const it=root.getFoldersByName('Backups');const folder=it.hasNext()?it.next():root.createFolder('Backups');const f=folder.createFile(zip);writeLog_(u,'Backup','ສຳຮອງຂໍ້ມູນລະບົບ',f.getId());return ok_('ສ້າງສຳຮອງຂໍ້ມູນແລ້ວ',{name:f.getName(),url:f.getUrl(),fileId:f.getId()});}catch(e){return fail_(safeMessage_(e));}
}

// =====================================================
// AUDIT LOG
// =====================================================
function listLogs(token,limit){try{requireRole_(token,['admin']);let rows=getRows_('Logs').sort((a,b)=>String(b.timestamp).localeCompare(String(a.timestamp)));if(limit)rows=rows.slice(0,Number(limit));return ok_('ສຳເລັດ',rows);}catch(e){return fail_(safeMessage_(e));}}
function writeLog_(user,action,detail,recordId){try{appendObject_('Logs',{logId:uid_('L'),timestamp:now_(),userId:user.userId||'',username:user.username||'',action,detail,recordId:recordId||''});}catch(ignore){}}

// =====================================================
// PERMISSION / SCOPING
// =====================================================
function scopeProjects_(user,rows){if(user.role==='admin')return rows;if(user.role==='head')return rows.filter(p=>String(p.departmentId)===String(user.departmentId));return rows.filter(p=>String(p.ownerUserId)===String(user.userId)||String(p.coOwners||'').split(',').map(x=>x.trim()).indexOf(user.userId)>=0);}
function assertProjectAccess_(user,p,write){if(!p)throw new Error('ບໍ່ພົບໂຄງການ');if(user.role==='admin')return true;if(user.role==='head'&&String(p.departmentId)===String(user.departmentId))return true;if(user.role==='teacher'&&String(p.ownerUserId)===String(user.userId))return true;if(!write&&user.role==='teacher'&&String(p.coOwners||'').split(',').map(x=>x.trim()).indexOf(user.userId)>=0)return true;throw new Error('ທ່ານບໍ່ມີສິດເຂົ້າເຖິງໂຄງການນີ້');}
function scopedDepartments_(user){const all=getRows_('Departments').filter(d=>d.status==='ນຳໃຊ້ງານ');if(user.role==='admin')return all;if(user.departmentId)return all.filter(d=>d.departmentId===user.departmentId);return all;}

// =====================================================
// DATA HELPERS WITH TEXT FINDER OPTIMIZATION
// =====================================================
function getSheet_(name){
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sh = ss.getSheetByName(name);
  if (!sh && SHEETS[name]) {
    sh = ensureSheet_(ss, name, SHEETS[name]);
  }
  return sh;
}

function getRows_(name){
  const sh=getSheet_(name); 
  if(!sh||sh.getLastRow()<2)return [];
  const values=sh.getDataRange().getValues();
  const headers=values.shift().map(String);
  return values.filter(r=>r.some(v=>v!==''&&v!==null)).map(r=>{
    const o={}; 
    headers.forEach((h,i)=>o[h]=normalizeCell_(r[i])); 
    return o;
  });
}

function normalizeCell_(v){
  if(Object.prototype.toString.call(v)==='[object Date]')return Utilities.formatDate(v,APP_TZ,'yyyy-MM-dd');
  return v;
}

function appendRows_(name,rows){
  if(!rows.length)return;
  const sh=getSheet_(name);
  sh.getRange(sh.getLastRow()+1,1,rows.length,rows[0].length).setValues(rows);
}

function appendObject_(name,obj){
  const headers=SHEETS[name];
  appendRows_(name,[headers.map(h=>obj[h]===undefined?'':obj[h])]);
}

function updateById_(name, idField, id, obj) {
  const sh = getSheet_(name);
  if (sh.getLastRow() < 2) throw new Error('ບໍ່ພົບຂໍ້ມູນທີ່ຕ້ອງການແກ້ໄຂ');
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);
  const colIndex = headers.indexOf(idField) + 1;
  if (colIndex < 1) throw new Error('ບໍ່ພົບ Primary Key');
  const tf = sh.getRange(2, colIndex, sh.getLastRow() - 1, 1).createTextFinder(String(id)).matchEntireCell(true).findNext();
  if (!tf) throw new Error('ບໍ່ພົບຂໍ້ມູນທີ່ຕ້ອງການແກ້ໄຂ');
  sh.getRange(tf.getRow(), 1, 1, headers.length).setValues([headers.map(h => obj[h] === undefined ? '' : obj[h])]);
}

function deleteById_(name, idField, id) {
  const sh = getSheet_(name);
  if (sh.getLastRow() < 2) return false;
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);
  const colIndex = headers.indexOf(idField) + 1;
  const tf = sh.getRange(2, colIndex, sh.getLastRow() - 1, 1).createTextFinder(String(id)).matchEntireCell(true).findNext();
  if (tf) { sh.deleteRow(tf.getRow()); return true; }
  return false;
}

function deleteWhere_(name, field, value) {
  const sh = getSheet_(name);
  if (sh.getLastRow() < 2) return;
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);
  const colIndex = headers.indexOf(field) + 1;
  const tfAll = sh.getRange(2, colIndex, sh.getLastRow() - 1, 1).createTextFinder(String(value)).matchEntireCell(true).findAll();
  tfAll.reverse().forEach(cell => sh.deleteRow(cell.getRow()));
}

function findById_(name,idField,id){return getRows_(name).find(x=>String(x[idField])===String(id))||null;}
function getDepartmentName_(id){return (findById_('Departments','departmentId',id)||{}).name||'';}
function getUserName_(id){if(id==='ADMIN')return 'ຜູ້ດູແລລະບົບ';return (findById_('Users','userId',id)||{}).fullName||id||'';}
function enrichExpense_(e){const x=Object.assign({},e);const p=findById_('Projects','projectId',e.projectId);x.projectName=p?p.projectName:'';const a=e.activityId&&findById_('Activities','activityId',e.activityId);x.activityName=a?a.activityName:'';return x;}
function sanitizeUser_(u){const x=Object.assign({},u);delete x.password;x.departmentName=getDepartmentName_(x.departmentId);return x;}
function groupSum_(rows,key,fn){const out={};rows.forEach(r=>{const k=String(r[key]||'ບໍ່ລະບຸ');out[k]=(out[k]||0)+fn(r);});return Object.keys(out).map(k=>({key:k,value:out[k]}));}
function labelDepartmentGroups_(groups){return groups.map(g=>({key:g.key,label:getDepartmentName_(g.key)||g.key,value:g.value}));}

// =====================================================
// UTILITIES
// =====================================================
function ok_(message,data){return {success:true,message:message||'ສຳເລັດ',data:data===undefined?null:data};}
function fail_(message){return {success:false,message:message||'ເກີດຂໍ້ຜິດພາດ'};}
function safeMessage_(e){const m=String(e&&e.message?e.message:e||'ເກີດຂໍ້ຜິດພາດ');if(m==='SESSION_EXPIRED')return 'ເຊດຊັນໝົດອາຍຸ ກະລຸນາເຂົ້າສູ່ລະບົບໃໝ່';return m.replace(/Exception: /g,'').slice(0,500);}
function clean_(v){return String(v===undefined||v===null?'':v).trim();}
function num_(v){const n=Number(v);return isFinite(n)?n:0;}
function clamp_(v,min,max){return Math.max(min,Math.min(max,v));}
function uid_(prefix){return prefix+'-'+Utilities.formatDate(new Date(),APP_TZ,'yyyyMMddHHmmss')+'-'+Utilities.getUuid().slice(0,8);}
function now_(){return Utilities.formatDate(new Date(),APP_TZ,'yyyy-MM-dd HH:mm:ss');}
function safeFolderName_(s){return String(s||'ບໍ່ລະບຸ').replace(/[\\/:*?"<>|]/g,'-').slice(0,120);}
function csvEscape_(v){v=String(v===undefined||v===null?'':v);return /[",\n\r]/.test(v)?'"'+v.replace(/"/g,'""')+'"':v;}