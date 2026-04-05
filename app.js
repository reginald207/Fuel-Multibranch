/* ═══════════════════════════════════════════════════════════
   FUELTRACK MULTI-BRANCH — app.js  v4
   ✅ Bulk delete — no 150-record cap; selects ALL records
   ✅ Branch delete also deletes its sales + deactivates its users
      with option to keep / reassign staff & admins
   ✅ Corporate HQ can reassign user to another branch (any time)
   ✅ Click branch card → branch detail modal (staff, admin, stats)
   ✅ Reports: live preview/summary before download
   ✅ Sales Entry page: Download Excel template link
   ✅ Only Corporate HQ can change roles (branch admin cannot)
   ✅ Branch admin / staff dashboard/analytics fixed (branch-scoped sub)
   ✅ Current user always first in Users table
   ✅ Page/section preserved on auth-state refresh (no logout flicker)
   ✅ Search in Users table and Pending Approvals
   ✅ Role change does NOT update createdAt (joined date preserved)
   ✅ Branch admins see only their own branch in Reports
   ✅ Product filter drives KPI cards + charts (3-card or 4-card)
═══════════════════════════════════════════════════════════ */

let state = {
  user: null, userDoc: null, role: 'staff',
  branchId: null, activeBranch: 'all',
  demoMode: false, salesData: [], branches: [],
  currentPeriod: 30, pendingUpload: [],
  selectedRows: new Set(),        // ALL IDs (no display cap)
  unsubSales: null, unsubPending: null,
  currentPage: 'dashboard',       // remember page across refresh
  _initialized: false,            // prevent double-init on auth re-fire
  // modal state
  _pendingToApprove: null, _pendingToReject: null,
  _userToChangeRole: null, _userToDelete: null,
  _branchToDelete: null, _branchDetailId: null,
  _userToReassign: null,
};

let charts = {};
let db, auth, authSecondary, ftLib;

const COLS = {
  salesRecon: { date:0, pmsQty:1, pmsPrice:2, pmsAmt:3, agoQty:4, agoPrice:5, agoAmt:6 },
  stockRecon: { date:0, pmsVariance:7, agoVariance:16 }
};

// ─────────────────────────────────────────────────────────────── INIT ──────
document.addEventListener('ftReady', () => {
  const ft = window._ft;
  auth = ft.auth; authSecondary = ft.authSecondary; db = ft.db; ftLib = ft;

  ft.onAuthStateChanged(auth, async (user) => {
    if (user) {
      showLoading(true);
      try {
        // Always re-read from Firestore so reassignment/role changes take effect
        const snap = await ft.getDoc(ft.doc(db, 'users', user.uid));
        if (!snap.exists()) {
          const pQ = await ft.getDocs(ft.query(ft.collection(db,'pendingUsers'), ft.where('email','==',user.email)));
          if (!pQ.empty) {
            const pd = pQ.docs[0].data();
            pd.status === 'rejected' ? showRejectedState(pd.rejectionReason||'Not approved.') : showPendingState();
          } else {
            showAuthError('Account not found. Please request access.');
            await ftLib.signOut(auth);
          }
          showLoading(false); return;
        }
        const data = snap.data();
        if (data.status === 'inactive') {
          showAuthError('Your account has been deactivated. Contact Corporate HQ.');
          await ftLib.signOut(auth); showLoading(false); return;
        }

        // Skip full re-init only if everything that drives the UI is unchanged.
        // This prevents logout flicker on token refresh while still reacting to
        // reassignments and role changes made by Corporate HQ.
        const sameUser   = state._initialized && state.user?.uid === user.uid;
        const sameBranch = (data.branchId || null) === state.branchId;
        const sameRole   = (data.role || 'staff') === state.role;
        if (sameUser && sameBranch && sameRole) {
          showLoading(false);
          return; // nothing changed — no re-init needed
        }

        // Something changed (new login, reassignment, role change) — full re-init
        if (state.unsubSales)   { state.unsubSales();   state.unsubSales   = null; }
        if (state.unsubPending) { state.unsubPending(); state.unsubPending = null; }
        destroyCharts();
        await initApp(user, data);
      } catch(e) { showLoading(false); showAuthError('Error: '+e.message); }
    } else {
      state._initialized = false;
      showAuthScreen();
    }
  });
});

// ─────────────────────────────────────────────────────────────── AUTH ──────
function switchAuthTab(tab) {
  document.querySelectorAll('.auth-tab').forEach((t,i)=>
    t.classList.toggle('active',(i===0&&tab==='login')||(i===1&&tab==='signup')));
  ['loginForm','signupForm','pendingApprovalForm','rejectedForm'].forEach(id=>
    document.getElementById(id).classList.add('hidden'));
  document.getElementById({login:'loginForm',signup:'signupForm',pending:'pendingApprovalForm',rejected:'rejectedForm'}[tab]||'loginForm').classList.remove('hidden');
  document.getElementById('authError').classList.add('hidden');
  if (tab==='signup') loadBranchesForSignup();
}

async function loadBranchesForSignup() {
  const sel = document.getElementById('signupBranch');
  sel.innerHTML = '<option value="">— Loading branches… —</option>';
  if (!window._ftReady) { sel.innerHTML='<option value="">Firebase not connected</option>'; return; }
  try {
    const snap = await ftLib.getDocs(ftLib.collection(db,'branches'));
    const branches = snap.docs.map(d=>({id:d.id,...d.data()}))
      .filter(b=>b.active!==false).sort((a,b)=>(a.name||'').localeCompare(b.name||''));
    if (!branches.length) { sel.innerHTML='<option value="">No branches — contact HQ</option>'; return; }
    sel.innerHTML = '<option value="">— Select your branch —</option>' +
      branches.map(b=>`<option value="${b.id}">${b.name} (${b.code})</option>`).join('');
  } catch(e) { sel.innerHTML='<option value="">Unable to load — check Firestore rules</option>'; }
}

async function handleLogin() {
  const email=document.getElementById('loginEmail').value.trim();
  const pw=document.getElementById('loginPassword').value;
  if (!email||!pw) return showAuthError('Please enter email and password');
  showLoading(true);
  try { await ftLib.signInWithEmailAndPassword(auth,email,pw); }
  catch(e) { showLoading(false); showAuthError(friendlyAuthError(e.code)); }
}

async function handleSignupRequest() {
  const name=document.getElementById('signupName').value.trim();
  const email=document.getElementById('signupEmail').value.trim();
  const branchId=document.getElementById('signupBranch').value;
  const role=document.getElementById('signupRole').value;
  const phone=document.getElementById('signupPhone').value.trim();
  if (!name||!email||!branchId) return showAuthError('Please fill in Name, Email and Branch');
  showLoading(true);
  try {
    const existQ = await ftLib.getDocs(ftLib.query(ftLib.collection(db,'pendingUsers'),ftLib.where('email','==',email)));
    if (existQ.docs.find(d=>d.data().status==='pending')) {
      showLoading(false); return showAuthError('A pending request already exists for this email.');
    }
    const branchSnap = await ftLib.getDoc(ftLib.doc(db,'branches',branchId));
    const branchName = branchSnap.exists()?branchSnap.data().name:'Unknown';
    await ftLib.addDoc(ftLib.collection(db,'pendingUsers'),{
      name,email,branchId,branchName,role,phone,status:'pending',createdAt:new Date().toISOString()
    });
    showLoading(false); showPendingState();
  } catch(e) { showLoading(false); showAuthError('Error: '+e.message); }
}

function showPendingState() {
  document.getElementById('authScreen').classList.remove('hidden');
  document.getElementById('appShell').classList.add('hidden');
  ['loginForm','signupForm','rejectedForm'].forEach(id=>document.getElementById(id).classList.add('hidden'));
  document.getElementById('pendingApprovalForm').classList.remove('hidden');
  document.querySelectorAll('.auth-tab').forEach(t=>t.classList.remove('active'));
}
function showRejectedState(reason) {
  document.getElementById('authScreen').classList.remove('hidden');
  document.getElementById('appShell').classList.add('hidden');
  ['loginForm','signupForm','pendingApprovalForm'].forEach(id=>document.getElementById(id).classList.add('hidden'));
  document.getElementById('rejectedReason').textContent = reason;
  document.getElementById('rejectedForm').classList.remove('hidden');
  document.querySelectorAll('.auth-tab').forEach(t=>t.classList.remove('active'));
}
async function handleLogout() {
  if (state.demoMode) { state._initialized=false; resetState(); showAuthScreen(); return; }
  if (state.unsubSales) state.unsubSales();
  if (state.unsubPending) state.unsubPending();
  state._initialized = false;
  await ftLib.signOut(auth);
}
function resetState() {
  state.salesData=[]; state.branches=[]; state.user=null; state.userDoc=null;
  state.demoMode=false; state.unsubSales=null; state.unsubPending=null;
  state.selectedRows=new Set(); state.currentPage='dashboard';
  destroyCharts();
}
function showAuthScreen() {
  resetState();
  document.getElementById('authScreen').classList.remove('hidden');
  document.getElementById('appShell').classList.add('hidden');
  ['signupForm','pendingApprovalForm','rejectedForm'].forEach(id=>document.getElementById(id).classList.add('hidden'));
  document.getElementById('loginForm').classList.remove('hidden');
  document.querySelectorAll('.auth-tab').forEach((t,i)=>t.classList.toggle('active',i===0));
}
function friendlyAuthError(code) {
  return {'auth/user-not-found':'No account with this email.','auth/wrong-password':'Incorrect password.',
    'auth/email-already-in-use':'Email already registered.','auth/invalid-email':'Invalid email address.',
    'auth/weak-password':'Password too weak.','auth/invalid-credential':'Invalid email or password.',
    'auth/too-many-requests':'Too many attempts. Please wait.'}[code]||'Auth error: '+code;
}
function showAuthError(msg) { const el=document.getElementById('authError'); el.textContent=msg; el.classList.remove('hidden'); }
function enableDemoMode() {
  state.demoMode=true; state.branches=generateDemoBranches(); state.salesData=generateDemoData();
  initApp({uid:'demo',email:'corporate@fueltrack.com'},{name:'Demo Corporate',email:'corporate@fueltrack.com',role:'corporate',branchId:null,status:'active'});
}

// ──────────────────────────────────────────────────────────── APP INIT ──────
async function initApp(user, userDoc) {
  showLoading(false);
  state._initialized = true;
  state.user = user; state.userDoc = userDoc;
  state.role = userDoc.role||'staff';
  state.branchId = userDoc.branchId||null;
  state.activeBranch = state.role==='corporate' ? 'all' : state.branchId;

  document.getElementById('authScreen').classList.add('hidden');
  document.getElementById('appShell').classList.remove('hidden');

  document.getElementById('sidebarName').textContent = userDoc.name||user.email;
  document.getElementById('sidebarRole').textContent = roleLabel(state.role);
  document.getElementById('sidebarAvatar').textContent = (userDoc.name||'U').charAt(0).toUpperCase();
  document.getElementById('sidebarBranch').textContent = '';
  document.getElementById('currentDate').textContent = new Date().toLocaleDateString('en-GH',
    {weekday:'short',year:'numeric',month:'short',day:'numeric'});
  document.getElementById('entryDate').value = new Date().toISOString().split('T')[0];
  setAnalyticsDefaultDates();

  await loadBranches();
  applyRoleUI();

  document.querySelectorAll('.period-btn').forEach(btn=>{
    btn.addEventListener('click',()=>{
      document.querySelectorAll('.period-btn').forEach(b=>b.classList.remove('active'));
      btn.classList.add('active');
      const p=btn.dataset.period;
      if (p==='custom') {
        document.getElementById('customDateRange').classList.remove('hidden');
      } else {
        document.getElementById('customDateRange').classList.add('hidden');
        if (p !== 'alltime') state.currentPeriod=parseInt(p);
      }
      refreshDashboard();
    });
    // Tooltip for "All Time" button: show first entry date on hover
    if (btn.dataset.period === 'alltime') {
      btn.addEventListener('mouseenter', () => {
        const branchFilter = state.role === 'corporate'
          ? (state.activeBranch === 'all' ? null : state.activeBranch)
          : state.branchId;
        const firstDate = getFirstEntryDate(branchFilter);
        btn.title = firstDate
          ? `Data from ${formatDateFull(firstDate)} to today`
          : 'No data uploaded yet';
      });
    }
  });
  document.getElementById('reportPeriod').addEventListener('change',function(){
    document.getElementById('reportCustomRange').style.display=this.value==='custom'?'block':'none';
    previewReport(); // auto-update preview on option change
  });
  // Auto-preview when branch/product filter changes in reports
  ['reportBranch','reportProduct','reportFrom','reportTo'].forEach(id=>{
    const el=document.getElementById(id);
    if (el) el.addEventListener('change', previewReport);
  });

  if (state.demoMode) {
    populateBranchSelects(); refreshDashboard(); renderSalesTable(); loadUsers();
  } else {
    subscribeToSales();
    if (state.role!=='staff') subscribeToApprovals();
    loadUsers();
  }

  // Show change-password banner if user is still on temp password
  checkPasswordChangeBanner(userDoc);

  // Restore last page instead of always going to dashboard
  showPage(state.currentPage||'dashboard');
}

function roleLabel(r){return{corporate:'Corporate HQ',admin:'Branch Admin',staff:'Staff'}[r]||r;}

function applyRoleUI() {
  const role=state.role;
  document.querySelectorAll('.corporate-only').forEach(el=>el.classList.toggle('hidden',role!=='corporate'));
  document.querySelectorAll('.admin-plus').forEach(el=>el.classList.toggle('hidden',role==='staff'));
  document.getElementById('branchSwitcher').style.display=role!=='corporate'?'none':'';
  const rbg=document.getElementById('reportBranchGroup');
  if (rbg) rbg.style.display=role==='staff'?'none':'';
  const canDel=role==='corporate'||role==='admin';
  document.querySelectorAll('.admin-plus-col').forEach(el=>el.classList.toggle('hidden',!canDel));
  const bdb=document.getElementById('bulkDeleteBtn'); if(bdb) bdb.classList.toggle('hidden',!canDel);
  // Branch admins can only report on their own branch
  if (role==='admin') {
    const rb=document.getElementById('reportBranch');
    if(rb&&state.branchId){
      rb.innerHTML=`<option value="${state.branchId}">${state.branches.find(b=>b.id===state.branchId)?.name||'My Branch'}</option>`;
    }
  }
}

// ─────────────────────────────────────────────────────── BRANCH MGMT ──────
async function loadBranches() {
  if (state.demoMode) { populateBranchSelects(); renderBranchesList(); return; }
  try {
    const snap=await ftLib.getDocs(ftLib.query(ftLib.collection(db,'branches'),ftLib.orderBy('name')));
    state.branches=snap.docs.map(d=>({id:d.id,...d.data()}));
    populateBranchSelects(); renderBranchesList(); updateNavBadges();
  } catch(e){ console.warn('loadBranches:',e.message); }
}

function populateBranchSelects() {
  const active=state.branches.filter(b=>b.active!==false);
  const myBranch=state.branches.find(b=>b.id===state.branchId);

  // Sidebar branch switcher (corporate only)
  const bs=document.getElementById('activeBranchSelect');
  bs.innerHTML='<option value="all">🏢 All Branches (HQ)</option>'+active.map(b=>`<option value="${b.id}">${b.name}</option>`).join('');
  bs.value=state.activeBranch;

  // Entry + Upload branch selects
  ['entryBranch','uploadBranch'].forEach(id=>{
    const el=document.getElementById(id);
    if (state.role==='corporate'||state.role==='admin') {
      el.innerHTML=active.map(b=>`<option value="${b.id}">${b.name}</option>`).join('');
      if (state.branchId) el.value=state.branchId;
    } else {
      el.innerHTML=myBranch?`<option value="${myBranch.id}">${myBranch.name}</option>`:'<option>No branch assigned</option>';
    }
  });

  // Report branch (branch admins locked to own branch in applyRoleUI)
  if (state.role!=='admin') {
    document.getElementById('reportBranch').innerHTML='<option value="all">All Branches</option>'+
      active.map(b=>`<option value="${b.id}">${b.name}</option>`).join('');
  }

  // User branch filter
  document.getElementById('userBranchFilter').innerHTML='<option value="all">All Branches</option>'+
    active.map(b=>`<option value="${b.id}">${b.name}</option>`).join('');

  document.getElementById('sidebarBranch').textContent=myBranch?myBranch.name:'';
}

function renderBranchesList() {
  const container=document.getElementById('branchesList');
  const badge=document.getElementById('branchCountBadge');
  const navBadge=document.getElementById('navBranchCount');
  if(badge) badge.textContent=state.branches.length+' branch'+(state.branches.length!==1?'es':'');
  if(navBadge) navBadge.textContent=state.branches.length;
  if(!state.branches.length){ container.innerHTML='<p class="empty-row">No branches yet.</p>'; return; }
  container.innerHTML=state.branches.map(b=>`
    <div class="branch-item ${b.active===false?'branch-inactive':''}">
      <div class="branch-item-icon" style="cursor:pointer" onclick="openBranchDetail('${b.id}')">${b.active===false?'⏸':'🟢'}</div>
      <div class="branch-item-info" style="cursor:pointer;flex:1" onclick="openBranchDetail('${b.id}')">
        <strong>${b.name}</strong>
        <span>${b.code} · ${b.region||'—'}</span>
        <span style="font-size:.75rem;color:var(--text-muted)">${b.address||''}</span>
      </div>
      <div class="branch-item-actions">
        <button class="btn-outline sm" onclick="openBranchDetail('${b.id}')">View</button>
        <button class="btn-outline sm" onclick="openToggleBranchModal('${b.id}',${b.active!==false})">${b.active===false?'Activate':'Deactivate'}</button>
        <button class="btn-danger" onclick="openDeleteBranchModal('${b.id}','${b.name.replace(/'/g,"\\'")}')">Delete</button>
      </div>
    </div>`).join('');
}

async function createBranch() {
  const name=document.getElementById('branchName').value.trim();
  const code=document.getElementById('branchCode').value.trim().toUpperCase();
  const region=document.getElementById('branchRegion').value;
  const address=document.getElementById('branchAddress').value.trim();
  const phone=document.getElementById('branchPhone').value.trim();
  if(!name||!code) return toast('Branch name and code are required','error');
  if(state.branches.find(b=>b.code===code)) return toast(`Code "${code}" already exists`,'error');
  showLoading(true);
  try {
    if(state.demoMode){ state.branches.push({id:'demo-'+Date.now(),name,code,region,address,phone,active:true,createdAt:new Date().toISOString()}); }
    else {
      const ref=await ftLib.addDoc(ftLib.collection(db,'branches'),{name,code,region,address,phone,active:true,createdBy:state.user.uid,createdAt:new Date().toISOString()});
      state.branches.push({id:ref.id,name,code,region,address,phone,active:true});
    }
    toast(`✅ Branch "${name}" created`,'success');
    ['branchName','branchCode','branchAddress','branchPhone'].forEach(id=>document.getElementById(id).value='');
    populateBranchSelects(); renderBranchesList(); updateNavBadges(); updateHQBanner();
  } catch(e){ toast('Error: '+e.message,'error'); }
  showLoading(false);
}

function openToggleBranchModal(branchId, active) {
  const b=state.branches.find(x=>x.id===branchId);
  const action=active?'Deactivate':'Activate';
  showConfirmModal(`${action} Branch`,`${action} "${b?.name}"?`, async()=>{
    try {
      if(!state.demoMode) await ftLib.updateDoc(ftLib.doc(db,'branches',branchId),{active:!active});
      const bx=state.branches.find(x=>x.id===branchId); if(bx) bx.active=!active;
      renderBranchesList(); populateBranchSelects(); toast(`Branch ${!active?'activated':'deactivated'}`,'success');
    } catch(e){ toast('Error: '+e.message,'error'); }
  });
}

// ── Branch detail modal ────────────────────────────────────────────────────
async function openBranchDetail(branchId) {
  state._branchDetailId = branchId;
  const b=state.branches.find(x=>x.id===branchId);
  if (!b) return;

  // Basic info
  document.getElementById('branchDetailName').textContent=b.name;
  document.getElementById('branchDetailCode').textContent=b.code||'—';
  document.getElementById('branchDetailRegion').textContent=b.region||'—';
  document.getElementById('branchDetailAddress').textContent=b.address||'—';
  document.getElementById('branchDetailPhone').textContent=b.phone||'—';
  document.getElementById('branchDetailStatus').textContent=b.active===false?'Inactive':'Active';

  // Sales stats
  const bData=state.salesData.filter(r=>r.branchId===branchId);
  const rev=bData.reduce((s,r)=>s+(r.pmsRevenue||0)+(r.agoRevenue||0),0);
  const pms=bData.reduce((s,r)=>s+(r.pmsActualSales||0),0);
  const ago=bData.reduce((s,r)=>s+(r.agoActualSales||0),0);
  document.getElementById('branchDetailRev').textContent=`GH\u20B5 ${fmt(rev)}`;
  document.getElementById('branchDetailPMS').textContent=`${fmt(pms)} L`;
  document.getElementById('branchDetailAGO').textContent=`${fmt(ago)} L`;
  document.getElementById('branchDetailRecords').textContent=bData.length;

  // Staff list
  const staffEl=document.getElementById('branchDetailStaff');
  staffEl.innerHTML='<p style="color:var(--text-muted);font-size:.82rem">Loading…</p>';

  if (state.demoMode) {
    staffEl.innerHTML=`
      <div class="branch-detail-user"><span class="badge badge-role-admin">Admin</span> Kwame Asante</div>
      <div class="branch-detail-user"><span class="badge badge-role-staff">Staff</span> Kofi Mensah</div>`;
  } else {
    try {
      const snap=await ftLib.getDocs(ftLib.query(ftLib.collection(db,'users'),ftLib.where('branchId','==',branchId)));
      const users=snap.docs.map(d=>({id:d.id,...d.data()})).filter(u=>u.status!=='inactive');
      if(!users.length){ staffEl.innerHTML='<p style="color:var(--text-muted);font-size:.82rem">No users assigned.</p>'; }
      else {
        staffEl.innerHTML=users.map(u=>`
          <div class="branch-detail-user">
            <span class="badge badge-role-${u.role}">${roleLabel(u.role)}</span>
            <span>${u.name||u.email}</span>
            <button class="btn-outline sm" style="margin-left:auto" onclick="openReassignModal('${u.id}','${(u.name||u.email).replace(/'/g,"\\'")}','${branchId}')">Reassign</button>
          </div>`).join('');
      }
    } catch(e){ staffEl.innerHTML=`<p style="color:var(--accent-red);font-size:.82rem">${e.message}</p>`; }
  }

  document.getElementById('branchDetailModal').classList.remove('hidden');
}

// ── Reassign user to another branch ───────────────────────────────────────
function openReassignModal(uid, name, currentBranchId) {
  state._userToReassign = { uid, name, currentBranchId };
  document.getElementById('reassignUserName').textContent = name;
  const sel=document.getElementById('reassignBranchSelect');
  const others=state.branches.filter(b=>b.id!==currentBranchId&&b.active!==false);
  sel.innerHTML=others.map(b=>`<option value="${b.id}">${b.name}</option>`).join('');
  // Close branch detail first, open reassign
  document.getElementById('branchDetailModal').classList.add('hidden');
  document.getElementById('reassignModal').classList.remove('hidden');
}

async function confirmReassign() {
  if (!state._userToReassign) return;
  const {uid, name} = state._userToReassign;
  const newBranchId = document.getElementById('reassignBranchSelect').value;
  const newBranch = state.branches.find(b=>b.id===newBranchId);
  closeModal();
  showLoading(true);
  try {
    if (!state.demoMode) {
      await ftLib.updateDoc(ftLib.doc(db,'users',uid),{
        branchId: newBranchId,
        branchName: newBranch?.name||''
      });
    }
    toast(`✅ ${name} reassigned to ${newBranch?.name||newBranchId}`,'success');
    loadUsers();
  } catch(e){ toast('Error: '+e.message,'error'); }
  showLoading(false);
  state._userToReassign = null;
}

// ── Delete branch — with option to keep/reassign staff ────────────────────
function openDeleteBranchModal(branchId, branchName) {
  state._branchToDelete = branchId;
  document.getElementById('deleteBranchName').textContent = branchName;
  // Populate reassign-to select
  const sel=document.getElementById('deleteBranchReassignTo');
  const others=state.branches.filter(b=>b.id!==branchId&&b.active!==false);
  sel.innerHTML='<option value="">— Deactivate (no reassign) —</option>'+others.map(b=>`<option value="${b.id}">${b.name}</option>`).join('');
  document.getElementById('deleteBranchModal').classList.remove('hidden');
}

async function confirmDeleteBranch() {
  const branchId=state._branchToDelete; if(!branchId) return;
  const keepStaff=document.getElementById('deleteBranchKeepStaff').checked;
  const reassignTo=document.getElementById('deleteBranchReassignTo').value;
  const reassignBranch=state.branches.find(b=>b.id===reassignTo);
  closeModal(); showLoading(true);
  try {
    if (!state.demoMode) {
      // 1. Delete all sales for this branch
      const salesSnap=await ftLib.getDocs(ftLib.query(ftLib.collection(db,'sales'),ftLib.where('branchId','==',branchId)));
      await Promise.all(salesSnap.docs.map(d=>ftLib.deleteDoc(d.ref)));

      // 2. Handle users in this branch
      const usersSnap=await ftLib.getDocs(ftLib.query(ftLib.collection(db,'users'),ftLib.where('branchId','==',branchId)));
      for (const ud of usersSnap.docs) {
        const u=ud.data();
        if (u.role==='corporate') continue; // never touch corporate accounts
        if (keepStaff && reassignTo) {
          await ftLib.updateDoc(ud.ref,{branchId:reassignTo,branchName:reassignBranch?.name||''});
        } else {
          await ftLib.updateDoc(ud.ref,{status:'inactive'});
        }
      }

      // 3. Delete the branch document
      await ftLib.deleteDoc(ftLib.doc(db,'branches',branchId));
    }
    state.branches=state.branches.filter(b=>b.id!==branchId);
    state.salesData=state.salesData.filter(r=>r.branchId!==branchId);
    populateBranchSelects(); renderBranchesList(); updateNavBadges(); updateHQBanner();
    renderSalesTable(); refreshDashboard();
    toast(`Branch deleted. Sales removed.${keepStaff&&reassignTo?` Staff reassigned to ${reassignBranch?.name}.`:' Staff deactivated.'}`,'success');
  } catch(e){ toast('Error: '+e.message,'error'); }
  showLoading(false); state._branchToDelete=null;
}

function handleBranchSwitch(branchId) {
  if (state.activeBranch === branchId) return; // no-op if same branch
  state.activeBranch = branchId;
  const b = state.branches.find(x=>x.id===branchId);
  document.getElementById('topbarBranch').textContent = branchId==='all' ? '' : ' — '+(b?.name||branchId);
  document.getElementById('anlBranchLabel').textContent = branchId==='all' ? 'All Branches' : (b?.name||branchId);
  if (!state.demoMode) {
    subscribeToSales(); // this will trigger scheduleRender() when data arrives
  } else {
    refreshDashboard();
    if (document.getElementById('page-analytics').classList.contains('active')) refreshAnalytics();
  }
}

// ──────────────────────────────────────────────── FIREBASE SUBSCRIPTIONS ──
// ── Render debounce ─────────────────────────────────────────────────────────
// Prevents rapid-fire re-renders during Firestore snapshot bursts (e.g. on
// first load when multiple docs arrive in quick succession). 80 ms is
// imperceptible to a human but eliminates the flash/glitch effect.
let _renderTimer = null;
function scheduleRender() {
  if (_renderTimer) clearTimeout(_renderTimer);
  _renderTimer = setTimeout(() => {
    refreshDashboard();
    renderSalesTable();
    if (document.getElementById('page-analytics').classList.contains('active')) refreshAnalytics();
    _renderTimer = null;
  }, 80);
}

function subscribeToSales() {
  if (state.unsubSales) { state.unsubSales(); state.unsubSales = null; }

  // Determine which branch to query:
  //  corporate + 'all'          → fetch ALL sales (no branch filter)
  //  corporate + specific branch → only that branch
  //  admin or staff             → ALWAYS their own branchId (never switches)
  let branchFilter = null;
  if (state.role === 'corporate') {
    branchFilter = (state.activeBranch === 'all') ? null : state.activeBranch;
  } else {
    branchFilter = state.branchId || null;
  }

  if (state.role !== 'corporate' && !branchFilter) {
    console.warn('subscribeToSales: no branchId for non-corporate user');
    state.salesData = [];
    refreshDashboard();
    renderSalesTable();
    return;
  }

  // No orderBy — avoids the composite index Firestore requires for
  // where('branchId') + orderBy('date'). We sort in JS after arrival instead.
  const q = branchFilter
    ? ftLib.query(ftLib.collection(db,'sales'), ftLib.where('branchId','==', branchFilter))
    : ftLib.query(ftLib.collection(db,'sales'));

  state.unsubSales = ftLib.onSnapshot(q, snap => {
    state.salesData = snap.docs
      .map(d => ({ id: d.id, ...d.data() }))
      .sort((a, b) => (b.date || '').localeCompare(a.date || ''));
    scheduleRender();
  }, err => {
    console.error('Sales sub error:', err.code, err.message);
    toast('Could not load sales data: ' + err.message, 'error');
  });
}

function subscribeToApprovals() {
  if (state.unsubPending) state.unsubPending();
  let q=ftLib.query(ftLib.collection(db,'pendingUsers'),ftLib.where('status','==','pending'));
  if (state.role==='admin'&&state.branchId)
    q=ftLib.query(ftLib.collection(db,'pendingUsers'),ftLib.where('status','==','pending'),ftLib.where('branchId','==',state.branchId));
  state.unsubPending=ftLib.onSnapshot(q, snap=>{
    updatePendingBadge(snap.size);
    renderPendingUsers(snap.docs.map(d=>({id:d.id,...d.data()})));
  });
}

// ──────────────────────────────────────────────────────────── NAVIGATION ──
function showPage(page) {
  document.querySelectorAll('.page').forEach(p=>{p.classList.remove('active');p.classList.add('hidden');});
  document.querySelectorAll('.nav-item').forEach(n=>n.classList.remove('active'));
  const pageEl=document.getElementById(`page-${page}`);
  if (pageEl){ pageEl.classList.remove('hidden'); pageEl.classList.add('active'); }
  const navEl=document.querySelector(`[data-page="${page}"]`);
  if (navEl) navEl.classList.add('active');
  document.getElementById('pageTitle').textContent={dashboard:'Dashboard',analytics:'Analytics',sales:'Upload Sales',reports:'Reports',branches:'Branch Management',users:'Users & Approvals'}[page]||page;
  state.currentPage=page; // remember for refresh
  if (page==='dashboard') refreshDashboard();
  if (page==='analytics') refreshAnalytics();
  if (page==='users') loadUsers();
  if (page==='branches') renderBranchesList();
  if (page==='reports') previewReport();
  if (window.innerWidth<=900) document.getElementById('sidebar').classList.remove('open');
}
function toggleSidebar(){ document.getElementById('sidebar').classList.toggle('open'); }

// ────────────────────────────────────────────────────────── DATA FILTER ──
function getActiveProduct(){ return document.getElementById('productFilter').value; }

function getActivePeriod() {
  return document.querySelector('.period-btn.active')?.dataset.period || '30';
}

function getFilteredData(){
  const period = getActivePeriod();
  if (period === 'alltime') {
    return [...state.salesData]; // return all records, no date cap
  }
  const now=new Date(); let from, to=new Date(now); to.setHours(23,59,59,999);
  if (period === 'custom') {
    const fv=document.getElementById('dateFrom').value, tv=document.getElementById('dateTo').value;
    from=fv?new Date(fv):new Date(now-7*86400000); to=tv?new Date(tv+'T23:59:59'):to;
  } else {
    from=new Date(now-parseInt(period)*86400000);
  }
  return state.salesData.filter(s=>{ const d=new Date(s.date); return d>=from&&d<=to; });
}

function getPreviousPeriodData(){
  const period = getActivePeriod();
  if (period === 'alltime' || period === 'custom') return []; // no comparison for these
  const p=parseInt(period)||30, now=new Date();
  const from=new Date(now-p*2*86400000), to=new Date(now-p*86400000);
  return state.salesData.filter(s=>{ const d=new Date(s.date); return d>=from&&d<=to; });
}

// Returns the earliest date string in the current salesData (or subset for a branch)
function getFirstEntryDate(branchId) {
  const data = branchId
    ? state.salesData.filter(r => r.branchId === branchId)
    : state.salesData;
  if (!data.length) return null;
  const sorted = [...data].map(r=>r.date).filter(Boolean).sort();
  return sorted[0] || null;
}

// ─────────────────────────────────────────────────────────── DASHBOARD ──
function refreshDashboard(){
  const data=getFilteredData(), product=getActiveProduct(), GH='GH\u20B5';

  // Early-exit KPI rendering if no data — show zeroed state cleanly
  const tPQ=data.reduce((s,r)=>s+(r.pmsActualSales||0),0);
  const tAQ=data.reduce((s,r)=>s+(r.agoActualSales||0),0);
  const tPR=data.reduce((s,r)=>s+(r.pmsRevenue||0),0);
  const tAR=data.reduce((s,r)=>s+(r.agoRevenue||0),0);
  const tPV=data.reduce((s,r)=>s+(r.pmsVariance||0),0);
  const tAV=data.reduce((s,r)=>s+(r.agoVariance||0),0);
  const tRev=tPR+tAR;

  const kpiGrid=document.getElementById('kpiGrid');
  if(product==='pms'){
    kpiGrid.innerHTML=`
      <div class="kpi-card" data-color="orange"><div class="kpi-icon">🔴</div><div class="kpi-data"><div class="kpi-label">PMS Actual Sales</div><div class="kpi-value">${fmt(tPQ)} L</div><div class="kpi-delta" style="color:var(--text-muted)">Petrol sold</div></div></div>
      <div class="kpi-card" data-color="blue"><div class="kpi-icon">💰</div><div class="kpi-data"><div class="kpi-label">PMS Revenue</div><div class="kpi-value">${GH} ${fmt(tPR)}</div><div class="kpi-delta" id="kpiRevenueDelta" style="color:var(--text-muted)">—</div></div></div>
      <div class="kpi-card" data-color="variance"><div class="kpi-icon">⚖️</div><div class="kpi-data"><div class="kpi-label">PMS Variance (G/L)</div><div class="kpi-value" style="color:${tPV>=0?'var(--accent-green)':'var(--accent-red)'}">${tPV>=0?'+':''}${fmt(tPV)} L</div><div class="kpi-delta" style="color:var(--text-muted)">Stock gain/loss</div></div></div>`;
    setRevDelta(tPR,'pms');
  } else if(product==='ago'){
    kpiGrid.innerHTML=`
      <div class="kpi-card" data-color="dark"><div class="kpi-icon">⚫</div><div class="kpi-data"><div class="kpi-label">AGO Actual Sales</div><div class="kpi-value">${fmt(tAQ)} L</div><div class="kpi-delta" style="color:var(--text-muted)">Diesel sold</div></div></div>
      <div class="kpi-card" data-color="blue"><div class="kpi-icon">💰</div><div class="kpi-data"><div class="kpi-label">AGO Revenue</div><div class="kpi-value">${GH} ${fmt(tAR)}</div><div class="kpi-delta" id="kpiRevenueDelta" style="color:var(--text-muted)">—</div></div></div>
      <div class="kpi-card" data-color="variance"><div class="kpi-icon">⚖️</div><div class="kpi-data"><div class="kpi-label">AGO Variance (G/L)</div><div class="kpi-value" style="color:${tAV>=0?'var(--accent-green)':'var(--accent-red)'}">${tAV>=0?'+':''}${fmt(tAV)} L</div><div class="kpi-delta" style="color:var(--text-muted)">Stock gain/loss</div></div></div>`;
    setRevDelta(tAR,'ago');
  } else {
    const nV=tPV+tAV;
    kpiGrid.innerHTML=`
      <div class="kpi-card" data-color="blue"><div class="kpi-icon">💰</div><div class="kpi-data"><div class="kpi-label">Total Revenue</div><div class="kpi-value">${GH} ${fmt(tRev)}</div><div class="kpi-delta" id="kpiRevenueDelta" style="color:var(--text-muted)">—</div></div></div>
      <div class="kpi-card" data-color="orange"><div class="kpi-icon">🔴</div><div class="kpi-data"><div class="kpi-label">PMS Actual Sales</div><div class="kpi-value">${fmt(tPQ)} L</div><div class="kpi-delta" style="color:var(--text-muted)">${GH} ${fmt(tPR)}</div></div></div>
      <div class="kpi-card" data-color="dark"><div class="kpi-icon">⚫</div><div class="kpi-data"><div class="kpi-label">AGO Actual Sales</div><div class="kpi-value">${fmt(tAQ)} L</div><div class="kpi-delta" style="color:var(--text-muted)">${GH} ${fmt(tAR)}</div></div></div>
      <div class="kpi-card" data-color="variance"><div class="kpi-icon">⚖️</div><div class="kpi-data"><div class="kpi-label">Stock Variance (G/L)</div><div class="kpi-value" style="color:${nV>=0?'var(--accent-green)':'var(--accent-red)'}">${nV>=0?'+':''}${fmt(nV)} L</div><div class="kpi-delta" style="color:var(--text-muted)">PMS: ${tPV>=0?'+':''}${fmt(tPV)}L | AGO: ${tAV>=0?'+':''}${fmt(tAV)}L</div></div></div>`;
    setRevDelta(tRev,'all');
  }

  // Only rebuild charts when there is data — avoids flicker on empty state
  if (data.length > 0) {
    renderCharts(data, product);
  } else {
    destroyCharts(); // clean slate — no phantom old charts
  }
  renderInsights(data, product);
  updateHQBanner();
}

function setRevDelta(cur, product){
  const prev=getPreviousPeriodData();
  let prevR=0;
  if(product==='pms') prevR=prev.reduce((s,r)=>s+(r.pmsRevenue||0),0);
  else if(product==='ago') prevR=prev.reduce((s,r)=>s+(r.agoRevenue||0),0);
  else prevR=prev.reduce((s,r)=>s+(r.pmsRevenue||0)+(r.agoRevenue||0),0);
  const delta=prevR?((cur-prevR)/prevR*100).toFixed(1):null;
  const el=document.getElementById('kpiRevenueDelta'); if(!el) return;
  el.textContent=delta?`${delta>0?'▲':'▼'} ${Math.abs(delta)}% vs prev period`:'First period';
  el.style.color=delta&&delta>0?'var(--accent-green)':(delta<0?'var(--accent-red)':'var(--text-muted)');
}

// ────────────────────────────────────────────────────────────── CHARTS ──
const CC={pms:{border:'#EF4444',bg:'rgba(239,68,68,.12)',solid:'#EF4444'},ago:{border:'#1F2937',bg:'rgba(31,41,55,.12)',solid:'#374151'},gain:'#059669',loss:'#DC2626'};
const CB={responsive:true,maintainAspectRatio:false,plugins:{legend:{labels:{font:{family:'Poppins',size:11},boxWidth:10}}}};
function destroyCharts(){Object.values(charts).forEach(c=>c&&c.destroy());charts={};}
function dc(k){if(charts[k]){charts[k].destroy();delete charts[k];}}

function getDailyGroups(data){
  const g={};
  data.forEach(r=>{
    const d=(r.date||'').split('T')[0];
    if(!g[d]) g[d]={pmsQty:0,agoQty:0,pmsRev:0,agoRev:0,pmsVar:0,agoVar:0};
    g[d].pmsQty+=r.pmsActualSales||0; g[d].agoQty+=r.agoActualSales||0;
    g[d].pmsRev+=r.pmsRevenue||0; g[d].agoRev+=r.agoRevenue||0;
    g[d].pmsVar+=r.pmsVariance||0; g[d].agoVar+=r.agoVariance||0;
  });
  return Object.entries(g).sort((a,b)=>a[0].localeCompare(b[0]));
}

function renderCharts(data, product='all'){
  renderSalesTrend(data,product); renderVarianceChart(data,product);
  renderProductPie(data); renderRevenueBar(data,product);
  if(state.role==='corporate'&&state.activeBranch==='all'){
    document.getElementById('branchCompareSection').classList.remove('hidden'); renderBranchCompare();
  } else document.getElementById('branchCompareSection').classList.add('hidden');
}

function renderSalesTrend(data,product){
  const c=document.getElementById('salesTrendChart');if(!c)return;
  const g=getDailyGroups(data); dc('salesTrend');
  const ds=[];
  if(product!=='ago') ds.push({label:'PMS (L)',data:g.map(([,v])=>v.pmsQty),borderColor:CC.pms.border,backgroundColor:CC.pms.bg,fill:true,tension:.35,borderWidth:2.5,pointRadius:3});
  if(product!=='pms') ds.push({label:'AGO (L)',data:g.map(([,v])=>v.agoQty),borderColor:CC.ago.border,backgroundColor:CC.ago.bg,fill:true,tension:.35,borderWidth:2.5,pointRadius:3});
  charts.salesTrend=new Chart(c,{type:'line',data:{labels:g.map(([d])=>formatDateLabel(d)),datasets:ds},options:{...CB,scales:{y:{grid:{color:'#F3F4F6'},ticks:{font:{family:'Poppins',size:10},callback:v=>fmtK(v)+' L'}},x:{grid:{display:false},ticks:{font:{family:'Poppins',size:10}}}}}});
}
function renderVarianceChart(data,product){
  const c=document.getElementById('varianceChart');if(!c)return;
  const g=getDailyGroups(data); dc('variance'); const ds=[];
  if(product!=='ago'){const v=g.map(([,x])=>x.pmsVar);ds.push({label:'PMS Var',data:v,backgroundColor:v.map(x=>x>=0?'rgba(5,150,105,.7)':'rgba(220,38,38,.7)'),borderRadius:4});}
  if(product!=='pms'){const v=g.map(([,x])=>x.agoVar);ds.push({label:'AGO Var',data:v,backgroundColor:v.map(x=>x>=0?'rgba(5,150,105,.4)':'rgba(220,38,38,.4)'),borderRadius:4});}
  charts.variance=new Chart(c,{type:'bar',data:{labels:g.map(([d])=>formatDateLabel(d)),datasets:ds},options:{...CB,scales:{y:{grid:{color:'#F3F4F6'},ticks:{font:{family:'Poppins',size:10},callback:v=>(v>=0?'+':'')+v.toFixed(0)}},x:{grid:{display:false},ticks:{font:{family:'Poppins',size:10}}}}}});
}
function renderProductPie(data){
  const c=document.getElementById('productPieChart');if(!c)return;
  const pr=data.reduce((s,r)=>s+(r.pmsRevenue||0),0),ar=data.reduce((s,r)=>s+(r.agoRevenue||0),0);
  dc('pie');
  charts.pie=new Chart(c,{type:'doughnut',data:{labels:['PMS (Petrol)','AGO (Diesel)'],datasets:[{data:[pr,ar],backgroundColor:[CC.pms.solid,CC.ago.solid],borderWidth:0,hoverOffset:8}]},options:{...CB,cutout:'60%',plugins:{legend:{position:'bottom',labels:{font:{family:'Poppins',size:11},padding:16,boxWidth:10}},tooltip:{callbacks:{label:ctx=>` GH\u20B5 ${fmt(ctx.parsed)}`}}}}});
}
function renderRevenueBar(data,product){
  const c=document.getElementById('revenueBarChart');if(!c)return;
  const g=getDailyGroups(data); dc('revBar'); const ds=[];
  if(product!=='ago') ds.push({label:'PMS Revenue',data:g.map(([,v])=>v.pmsRev),backgroundColor:CC.pms.bg,borderColor:CC.pms.border,borderWidth:1.5,borderRadius:4});
  if(product!=='pms') ds.push({label:'AGO Revenue',data:g.map(([,v])=>v.agoRev),backgroundColor:CC.ago.bg,borderColor:CC.ago.border,borderWidth:1.5,borderRadius:4});
  charts.revBar=new Chart(c,{type:'bar',data:{labels:g.map(([d])=>formatDateLabel(d)),datasets:ds},options:{...CB,scales:{y:{grid:{color:'#F3F4F6'},ticks:{font:{family:'Poppins',size:10},callback:v=>'GH\u20B5 '+fmtK(v)}},x:{grid:{display:false},ticks:{font:{family:'Poppins',size:10}}}}}});
}
function renderBranchCompare(){
  const c=document.getElementById('branchCompareChart');if(!c)return;
  const metric=document.getElementById('branchCompareMetric').value;
  const mL={revenue:'Total Revenue (GH\u20B5)',pms:'PMS Volume (L)',ago:'AGO Volume (L)',variance:'Net Variance (L)'};
  const totals=state.branches.filter(b=>b.active!==false).map(b=>{
    const bd=state.salesData.filter(r=>r.branchId===b.id);
    let v=0;
    if(metric==='revenue') v=bd.reduce((s,r)=>s+(r.pmsRevenue||0)+(r.agoRevenue||0),0);
    if(metric==='pms') v=bd.reduce((s,r)=>s+(r.pmsActualSales||0),0);
    if(metric==='ago') v=bd.reduce((s,r)=>s+(r.agoActualSales||0),0);
    if(metric==='variance') v=bd.reduce((s,r)=>s+(r.pmsVariance||0)+(r.agoVariance||0),0);
    return {name:b.name,value:v};
  }).sort((a,b)=>b.value-a.value);
  const colors=['#1E3A8A','#EF4444','#059669','#D97706','#7C3AED','#06B6D4'];
  dc('branchCompare');
  charts.branchCompare=new Chart(c,{type:'bar',data:{labels:totals.map(b=>b.name),datasets:[{label:mL[metric],data:totals.map(b=>b.value),backgroundColor:totals.map((_,i)=>colors[i%colors.length]+'CC'),borderColor:totals.map((_,i)=>colors[i%colors.length]),borderWidth:1.5,borderRadius:5}]},options:{...CB,indexAxis:'y',plugins:{...CB.plugins,legend:{display:false}},scales:{y:{grid:{display:false},ticks:{font:{family:'Poppins',size:11}}},x:{grid:{color:'#F3F4F6'},ticks:{font:{family:'Poppins',size:10},callback:v=>metric==='revenue'?'GH\u20B5 '+fmtK(v):fmtK(v)+' L'}}}}});
}

function updateHQBanner(){
  const banner=document.getElementById('hqOverviewBanner');
  if(state.role!=='corporate'||state.activeBranch!=='all'){banner.classList.add('hidden');return;}
  banner.classList.remove('hidden');
  const active=state.branches.filter(b=>b.active!==false);
  document.getElementById('hqBranchCountLabel').textContent=`${active.length} active branch${active.length!==1?'es':''}`;
  const cards=document.getElementById('hqBranchCards');
  if(!active.length){cards.innerHTML='<div style="color:rgba(255,255,255,.5);font-size:.85rem;padding:10px">No branches yet.</div>';return;}
  const product=getActiveProduct();
  cards.innerHTML=active.map(b=>{
    const bd=state.salesData.filter(r=>r.branchId===b.id);
    const pms=bd.reduce((s,r)=>s+(r.pmsActualSales||0),0);
    const ago=bd.reduce((s,r)=>s+(r.agoActualSales||0),0);
    const pmsRev=bd.reduce((s,r)=>s+(r.pmsRevenue||0),0);
    const agoRev=bd.reduce((s,r)=>s+(r.agoRevenue||0),0);
    const pmsVar=bd.reduce((s,r)=>s+(r.pmsVariance||0),0);
    const agoVar=bd.reduce((s,r)=>s+(r.agoVariance||0),0);
    const rev=product==='pms'?pmsRev:product==='ago'?agoRev:pmsRev+agoRev;
    const vt=product==='pms'?pmsVar:product==='ago'?agoVar:pmsVar+agoVar;
    const detail=product==='pms'?`PMS: ${fmtK(pms)}L`:product==='ago'?`AGO: ${fmtK(ago)}L`:`PMS: ${fmtK(pms)}L | AGO: ${fmtK(ago)}L`;
    return `<div class="hq-branch-card" onclick="switchToBranch('${b.id}')">
      <div class="hq-branch-name">${b.name}</div><div class="hq-branch-code">${b.code}</div>
      <div class="hq-branch-rev">GH\u20B5 ${fmtK(rev)}</div>
      <div class="hq-branch-detail">${detail}</div>
      <div class="hq-branch-var ${vt>=0?'var-gain':'var-loss'}">${vt>=0?'▲':'▼'} ${Math.abs(vt).toFixed(0)}L variance</div>
    </div>`;
  }).join('');
}

function switchToBranch(id){state.activeBranch=id;document.getElementById('activeBranchSelect').value=id;handleBranchSwitch(id);}
function updateNavBadges(){document.getElementById('navBranchCount').textContent=state.branches.length;}

// ─────────────────────────────────────────────────────────────── INSIGHTS ──
function renderInsights(data,product='all'){
  if(!data.length){['insightBestPMS','insightBestAGO','insightVariance','insightTrend'].forEach(id=>{document.getElementById(id).textContent='No data';});return;}
  const g=getDailyGroups(data);
  if(product!=='ago'){const b=g.reduce((a,x)=>x[1].pmsQty>a[1].pmsQty?x:a,g[0]);document.getElementById('insightBestPMS').textContent=b?`${formatDateLabel(b[0])} — ${fmt(b[1].pmsQty)} L`:'—';}
  else document.getElementById('insightBestPMS').textContent='—';
  if(product!=='pms'){const b=g.reduce((a,x)=>x[1].agoQty>a[1].agoQty?x:a,g[0]);document.getElementById('insightBestAGO').textContent=b?`${formatDateLabel(b[0])} — ${fmt(b[1].agoQty)} L`:'—';}
  else document.getElementById('insightBestAGO').textContent='—';
  const pv=data.reduce((s,r)=>s+(r.pmsVariance||0),0),av=data.reduce((s,r)=>s+(r.agoVariance||0),0);
  const nv=product==='pms'?pv:product==='ago'?av:pv+av;
  const dl=g.filter(([,v])=>{const l=product==='pms'?v.pmsVar:product==='ago'?v.agoVar:v.pmsVar+v.agoVar;return l<-5;}).length;
  document.getElementById('insightVariance').textContent=nv>=0?`📈 Net gain +${fmt(nv)}L`:`⚠️ Net loss ${fmt(Math.abs(nv))}L (${dl} day${dl!==1?'s':''})`;
  if(g.length>=4){
    const half=Math.floor(g.length/2);
    const vol=([,v])=>product==='pms'?v.pmsQty:product==='ago'?v.agoQty:v.pmsQty+v.agoQty;
    const f=g.slice(0,half).reduce((s,x)=>s+vol(x),0),l=g.slice(half).reduce((s,x)=>s+vol(x),0);
    const p=f?((l-f)/f*100).toFixed(1):0;
    document.getElementById('insightTrend').textContent=p>0?`📈 Volume up ${p}%`:p<0?`📉 Volume down ${Math.abs(p)}%`:`➡️ Stable`;
  } else document.getElementById('insightTrend').textContent='Need more data';
}

// ──────────────────────────────────────────────────────────── ANALYTICS ──
function setAnalyticsDefaultDates(){
  const now=new Date(),from=new Date(now-30*86400000);
  document.getElementById('anlFrom').value=from.toISOString().split('T')[0];
  document.getElementById('anlTo').value=now.toISOString().split('T')[0];
}
function refreshAnalytics(){
  const from=document.getElementById('anlFrom').value,to=document.getElementById('anlTo').value;
  const groupBy=document.getElementById('anlGroupBy').value;
  // state.salesData is already branch-scoped by the subscription for admin/staff.
  // For corporate, it reflects whatever branch (or all) they are currently viewing.
  const data=state.salesData.filter(s=>{
    const d=s.date.split('T')[0];
    return(!from||d>=from)&&(!to||d<=to);
  });

  // Update the branch label shown in analytics header
  const lbl = document.getElementById('anlBranchLabel');
  if (lbl) {
    if (state.role === 'corporate') {
      lbl.textContent = state.activeBranch === 'all'
        ? 'All Branches'
        : (state.branches.find(b=>b.id===state.activeBranch)?.name || state.activeBranch);
    } else {
      const myB = state.branches.find(b=>b.id===state.branchId);
      lbl.textContent = myB ? myB.name : 'My Branch';
    }
  }

  renderAnlSalesChart(data,groupBy);
  renderAnlRevenueChart(data,groupBy);
  renderAnlVarianceCharts(data,groupBy);
  renderAnlTable(data,groupBy);
}
function groupDataBy(data,groupBy){
  const g={};
  data.forEach(r=>{
    const d=new Date(r.date);let k;
    if(groupBy==='month') k=`${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}`;
    else if(groupBy==='week'){const j=new Date(d.getFullYear(),0,1);const w=Math.ceil(((d-j)/86400000+j.getDay()+1)/7);k=`${d.getFullYear()}-W${String(w).padStart(2,'0')}`;}
    else k=r.date.split('T')[0];
    if(!g[k]) g[k]={pmsQty:0,agoQty:0,pmsRev:0,agoRev:0,pmsVar:0,agoVar:0};
    g[k].pmsQty+=r.pmsActualSales||0;g[k].agoQty+=r.agoActualSales||0;
    g[k].pmsRev+=r.pmsRevenue||0;g[k].agoRev+=r.agoRevenue||0;
    g[k].pmsVar+=r.pmsVariance||0;g[k].agoVar+=r.agoVariance||0;
  });
  return Object.entries(g).sort((a,b)=>a[0].localeCompare(b[0]));
}
function renderAnlSalesChart(data,gb){
  const c=document.getElementById('anlSalesChart');if(!c)return;
  const g=groupDataBy(data,gb);dc('anlSales');
  charts.anlSales=new Chart(c,{type:'bar',data:{labels:g.map(([k])=>k),datasets:[{label:'PMS (L)',data:g.map(([,v])=>v.pmsQty),backgroundColor:CC.pms.solid,borderRadius:3},{label:'AGO (L)',data:g.map(([,v])=>v.agoQty),backgroundColor:CC.ago.solid,borderRadius:3}]},options:{...CB,scales:{x:{grid:{display:false},ticks:{font:{family:'Poppins',size:10}}},y:{grid:{color:'#F3F4F6'},ticks:{font:{family:'Poppins',size:10},callback:v=>fmtK(v)+' L'}}}}});
}
function renderAnlRevenueChart(data,gb){
  const c=document.getElementById('anlRevenueChart');if(!c)return;
  const g=groupDataBy(data,gb);dc('anlRev');
  charts.anlRev=new Chart(c,{type:'line',data:{labels:g.map(([k])=>k),datasets:[{label:'PMS Revenue',data:g.map(([,v])=>v.pmsRev),borderColor:CC.pms.border,backgroundColor:CC.pms.bg,fill:true,tension:.35,borderWidth:2},{label:'AGO Revenue',data:g.map(([,v])=>v.agoRev),borderColor:CC.ago.border,backgroundColor:CC.ago.bg,fill:true,tension:.35,borderWidth:2}]},options:{...CB,scales:{y:{grid:{color:'#F3F4F6'},ticks:{font:{family:'Poppins',size:10},callback:v=>'GH\u20B5 '+fmtK(v)}},x:{grid:{display:false},ticks:{font:{family:'Poppins',size:10}}}}}});
}
function renderAnlVarianceCharts(data,gb){
  const g=groupDataBy(data,gb);
  ['anlPMSVarianceChart','anlAGOVarianceChart'].forEach((cid,idx)=>{
    const c=document.getElementById(cid);if(!c)return;
    const k=idx===0?'anlPMSVar':'anlAGOVar';
    const vd=g.map(([,v])=>idx===0?v.pmsVar:v.agoVar);
    dc(k);
    charts[k]=new Chart(c,{type:'bar',data:{labels:g.map(([x])=>x),datasets:[{label:idx===0?'PMS Var':'AGO Var',data:vd,backgroundColor:vd.map(v=>v>=0?'rgba(5,150,105,.65)':'rgba(220,38,38,.65)'),borderColor:vd.map(v=>v>=0?'#059669':'#DC2626'),borderWidth:1,borderRadius:3}]},options:{...CB,plugins:{...CB.plugins,legend:{display:false}},scales:{y:{grid:{color:'#F3F4F6'},ticks:{font:{family:'Poppins',size:10},callback:v=>(v>=0?'+':'')+v.toFixed(1)+'L'}},x:{grid:{display:false},ticks:{font:{family:'Poppins',size:10}}}}}});
  });
}
function renderAnlTable(data,gb){
  const g=groupDataBy(data,gb),tbody=document.getElementById('anlTableBody');
  if(!g.length){tbody.innerHTML='<tr><td colspan="7" class="empty-row">No data</td></tr>';return;}
  let prev=0;
  tbody.innerHTML=g.map(([period,v])=>{
    const tQ=v.pmsQty+v.agoQty,tR=v.pmsRev+v.agoRev;
    const gr=prev?((tQ-prev)/prev*100).toFixed(1):'—';
    const gH=gr!=='—'?`<span style="color:${parseFloat(gr)>=0?'var(--accent-green)':'var(--accent-red)'}">${parseFloat(gr)>=0?'▲':'▼'} ${Math.abs(gr)}%</span>`:'—';
    const pH=`<span style="color:${v.pmsVar>=0?'var(--accent-green)':'var(--accent-red)'}">${v.pmsVar>=0?'+':''}${fmt(v.pmsVar)}</span>`;
    const aH=`<span style="color:${v.agoVar>=0?'var(--accent-green)':'var(--accent-red)'}">${v.agoVar>=0?'+':''}${fmt(v.agoVar)}</span>`;
    prev=tQ;
    return `<tr><td><strong>${period}</strong></td><td>${fmt(v.pmsQty)} L</td><td>${fmt(v.agoQty)} L</td><td><strong>GH\u20B5 ${fmt(tR)}</strong></td><td>${pH} L</td><td>${aH} L</td><td>${gH}</td></tr>`;
  }).join('');
}

// ──────────────────────────────────────────────────── SALES TABLE ──────
function renderSalesTable(searchTerm=''){
  const tbody=document.getElementById('salesTableBody');
  let rows=[...state.salesData];
  if(searchTerm){
    const t=searchTerm.toLowerCase();
    rows=rows.filter(r=>{
      // Search raw ISO date ("2024-01-15"), full formatted date ("15 Jan 2024"),
      // branch name, and notes — so any of those words/numbers match
      const rawDate  = (r.date||'').toLowerCase();
      const fullDate = formatDateFull(r.date).toLowerCase();
      const branch   = (r.branchName||'').toLowerCase();
      const notes    = (r.notes||'').toLowerCase();
      return rawDate.includes(t) || fullDate.includes(t) || branch.includes(t) || notes.includes(t);
    });
  }
  if(!rows.length){tbody.innerHTML=`<tr><td colspan="10" class="empty-row">No records found</td></tr>`;return;}
  const canDel=state.role==='corporate'||state.role==='admin';
  tbody.innerHTML=rows.map(r=>{
    const pV=r.pmsVariance||0,aV=r.agoVariance||0,isSel=state.selectedRows.has(r.id);
    return `<tr class="${isSel?'row-selected':''}">
      <td class="admin-plus-col ${!canDel?'hidden':''}"><input type="checkbox" ${isSel?'checked':''} onchange="toggleRowSelect('${r.id}',this.checked)"></td>
      <td style="white-space:nowrap">${formatDateFull(r.date)}</td>
      <td><span class="badge badge-branch">${r.branchName||'—'}</span></td>
      <td>${fmt(r.pmsActualSales)} L</td><td>GH\u20B5 ${fmt(r.pmsRevenue)}</td>
      <td>${fmt(r.agoActualSales)} L</td><td>GH\u20B5 ${fmt(r.agoRevenue)}</td>
      <td class="${pV>=0?'var-pos':'var-neg'}">${pV>=0?'+':''}${fmt(pV)}</td>
      <td class="${aV>=0?'var-pos':'var-neg'}">${aV>=0?'+':''}${fmt(aV)}</td>
      <td class="admin-plus-col ${!canDel?'hidden':''}">${canDel?`<button class="btn-danger" onclick="openDeleteRecordModal('${r.id}')">Delete</button>`:'—'}</td>
    </tr>`;
  }).join('');
}
function toggleRowSelect(id,checked){if(checked)state.selectedRows.add(id);else state.selectedRows.delete(id);}

// Toggle ALL records (not just displayed 150) ──────────────────────────────
function toggleSelectAll(checked){
  // Operate on full salesData (no display cap)
  state.salesData.forEach(r=>{if(checked)state.selectedRows.add(r.id);else state.selectedRows.delete(r.id);});
  renderSalesTable(document.getElementById('tableSearch').value);
}
function filterTable(){renderSalesTable(document.getElementById('tableSearch').value);}

function openDeleteRecordModal(id){
  showConfirmModal('Delete Record','Delete this record? This cannot be undone.', async()=>{
    if(state.demoMode){state.salesData=state.salesData.filter(r=>r.id!==id);state.selectedRows.delete(id);renderSalesTable();refreshDashboard();toast('Deleted','success');return;}
    try{await ftLib.deleteDoc(ftLib.doc(db,'sales',id));state.selectedRows.delete(id);toast('Deleted','success');}
    catch(e){toast('Error: '+e.message,'error');}
  },true);
}

function openBulkDeleteModal(){
  const count=state.selectedRows.size;
  if(!count) return toast('Select records to delete first','warning');
  document.getElementById('bulkDeleteDesc').textContent=`Delete ${count} selected record${count!==1?'s':''}? This cannot be undone.`;
  document.getElementById('bulkDeleteModal').classList.remove('hidden');
}

async function confirmBulkDelete(){
  const ids=[...state.selectedRows]; if(!ids.length) return;
  closeModal(); showLoading(true);
  try{
    if(state.demoMode){ state.salesData=state.salesData.filter(r=>!ids.includes(r.id)); }
    else{
      // Batch in chunks of 499
      const chunks=[];
      for(let i=0;i<ids.length;i+=499) chunks.push(ids.slice(i,i+499));
      for(const chunk of chunks) await Promise.all(chunk.map(id=>ftLib.deleteDoc(ftLib.doc(db,'sales',id))));
    }
    state.selectedRows.clear();
    const sa=document.getElementById('selectAllRows');if(sa)sa.checked=false;
    renderSalesTable();refreshDashboard();
    toast(`✅ ${ids.length} record${ids.length!==1?'s':''} deleted`,'success');
  }catch(e){toast('Bulk delete failed: '+e.message,'error');}
  showLoading(false);
}

// Download Excel template for sales entry ─────────────────────────────────
function downloadSalesTemplate(){
  const data=[
    {Date:'2024-01-15',Branch:'Branch Name','PMS Actual Sales (L)':500,'PMS Price (GHC/L)':14.34,'PMS Revenue (GHC)':7170,'AGO Actual Sales (L)':800,'AGO Price (GHC/L)':13.89,'AGO Revenue (GHC)':11112,'PMS Variance (L)':2.5,'AGO Variance (L)':-1.2,Notes:'Morning shift'},
    {Date:'2024-01-16',Branch:'Branch Name','PMS Actual Sales (L)':620,'PMS Price (GHC/L)':14.34,'PMS Revenue (GHC)':8891,'AGO Actual Sales (L)':910,'AGO Price (GHC/L)':13.89,'AGO Revenue (GHC)':12640,'PMS Variance (L)':-0.8,'AGO Variance (L)':3.1,Notes:''},
  ];
  const ws=XLSX.utils.json_to_sheet(data);
  const wb=XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb,ws,'Sales Data');
  XLSX.writeFile(wb,'FuelTrack_Sales_Template.xlsx');
  toast('Template downloaded','success');
}

// ──────────────────────────────────────────────────── MANUAL ENTRY ──────
function autoCalcEntry(){
  const pQ=parseFloat(document.getElementById('entryPMSQty').value)||0,pP=parseFloat(document.getElementById('entryPMSPrice').value)||0;
  const aQ=parseFloat(document.getElementById('entryAGOQty').value)||0,aP=parseFloat(document.getElementById('entryAGOPrice').value)||0;
  if(pQ&&pP) document.getElementById('entryPMSRevenue').value=(pQ*pP).toFixed(2);
  if(aQ&&aP) document.getElementById('entryAGORevenue').value=(aQ*aP).toFixed(2);
}

async function submitManualEntry(){
  const date=document.getElementById('entryDate').value,branchId=document.getElementById('entryBranch').value;
  const pQ=parseFloat(document.getElementById('entryPMSQty').value)||0,pR=parseFloat(document.getElementById('entryPMSRevenue').value)||0;
  const aQ=parseFloat(document.getElementById('entryAGOQty').value)||0,aR=parseFloat(document.getElementById('entryAGORevenue').value)||0;
  const pV=parseFloat(document.getElementById('entryPMSVariance').value)||0,aV=parseFloat(document.getElementById('entryAGOVariance').value)||0;
  const notes=document.getElementById('entryNotes').value.trim();
  if(!date||!branchId) return toast('Date and Branch required','error');
  if(!pQ&&!aQ) return toast('At least one quantity required','error');
  const b=state.branches.find(x=>x.id===branchId);
  const rec={date,branchId,branchName:b?b.name:'',pmsActualSales:pQ,pmsRevenue:pR,agoActualSales:aQ,agoRevenue:aR,pmsVariance:pV,agoVariance:aV,notes,source:'manual',uploadedAt:new Date().toISOString(),uploadedBy:state.user?.uid||'demo'};
  showLoading(true);
  try{
    if(state.demoMode){rec.id='demo-'+Date.now();state.salesData.unshift(rec);}
    else await ftLib.addDoc(ftLib.collection(db,'sales'),rec);
    toast('Entry saved!','success');
    ['entryPMSQty','entryPMSPrice','entryPMSRevenue','entryPMSVariance','entryAGOQty','entryAGOPrice','entryAGORevenue','entryAGOVariance','entryNotes'].forEach(id=>document.getElementById(id).value='');
    if(state.demoMode){renderSalesTable();refreshDashboard();}
  }catch(e){toast('Error: '+e.message,'error');}
  showLoading(false);
}

// ──────────────────────────────────────────────────── EXCEL UPLOAD ──────
function handleDrop(e){e.preventDefault();document.getElementById('uploadZone').classList.remove('drag-over');const f=e.dataTransfer.files[0];if(f)processExcelFile(f);}
function handleFileUpload(e){const f=e.target.files[0];if(f)processExcelFile(f);}
function cancelUpload(){
  state.pendingUpload=[];
  document.getElementById('uploadPreview').classList.add('hidden');
  document.getElementById('uploadActions').classList.add('hidden');
  document.getElementById('fileInput').value='';
  toast('Upload cancelled','info');
}

function processExcelFile(file){
  const reader=new FileReader();
  reader.onload=(e)=>{
    try{
      const wb=XLSX.read(e.target.result,{type:'binary',cellDates:true});
      const sn=wb.SheetNames.map(s=>s.toLowerCase());
      const hasSR=sn.some(s=>s.includes('sales reconciliation')||s.includes('sales recon'));
      const hasSt=sn.some(s=>s.includes('stock recon'));
      if(!hasSR||!hasSt){toast('File must contain "Sales reconciliation" and "Stock Recon." sheets','error');return;}
      const srSheet=wb.SheetNames.find(s=>s.toLowerCase().includes('sales reconciliation')||s.toLowerCase().includes('sales recon'));
      const stSheet=wb.SheetNames.find(s=>s.toLowerCase().includes('stock recon'));
      const srData=parseSheet(wb,srSheet),stData=parseSheet(wb,stSheet);
      const vm={};
      stData.forEach(row=>{const ds=parseDateVal(row[COLS.stockRecon.date]);if(!ds)return;vm[ds]={pmsVariance:parseFloat(row[COLS.stockRecon.pmsVariance])||0,agoVariance:parseFloat(row[COLS.stockRecon.agoVariance])||0};});
      const branchId=document.getElementById('uploadBranch').value;
      const branch=state.branches.find(b=>b.id===branchId);
      const records=[];
      srData.forEach(row=>{
        const ds=parseDateVal(row[COLS.salesRecon.date]);if(!ds)return;
        const pQ=parseFloat(row[COLS.salesRecon.pmsQty])||0,pA=parseFloat(row[COLS.salesRecon.pmsAmt])||0;
        const aQ=parseFloat(row[COLS.salesRecon.agoQty])||0,aA=parseFloat(row[COLS.salesRecon.agoAmt])||0;
        if(pQ<=0&&aQ<=0)return;
        const vd=vm[ds]||{pmsVariance:0,agoVariance:0};
        records.push({date:ds,branchId,branchName:branch?branch.name:'',pmsActualSales:pQ,pmsRevenue:pA,agoActualSales:aQ,agoRevenue:aA,pmsVariance:vd.pmsVariance,agoVariance:vd.agoVariance,source:'excel',uploadedAt:new Date().toISOString(),uploadedBy:state.user?.uid||'demo'});
      });
      if(!records.length){toast('No valid records found','error');return;}
      state.pendingUpload=records;showUploadPreview(records);
    }catch(err){toast('Error reading file: '+err.message,'error');}
  };
  reader.readAsBinaryString(file);
}
function parseSheet(wb,sheetName){return XLSX.utils.sheet_to_json(wb.Sheets[sheetName],{header:1,defval:null,raw:true}).slice(4);}
function parseDateVal(val){
  if(!val)return null;
  if(val instanceof Date)return val.toISOString().split('T')[0];
  if(typeof val==='number')return new Date((val-25569)*86400000).toISOString().split('T')[0];
  if(typeof val==='string'){const d=new Date(val);if(!isNaN(d))return d.toISOString().split('T')[0];}
  return null;
}
function showUploadPreview(records){
  const preview=document.getElementById('uploadPreview'); preview.classList.remove('hidden');
  const pT=records.reduce((s,r)=>s+r.pmsActualSales,0),aT=records.reduce((s,r)=>s+r.agoActualSales,0);
  const pV=records.reduce((s,r)=>s+r.pmsVariance,0),aV=records.reduce((s,r)=>s+r.agoVariance,0);
  const rT=records.reduce((s,r)=>s+r.pmsRevenue+r.agoRevenue,0);
  preview.innerHTML=`<strong>✅ ${records.length} records parsed</strong>
    <div class="preview-stats">
      <div><span>PMS Sales</span><strong>${fmt(pT)} L</strong></div><div><span>AGO Sales</span><strong>${fmt(aT)} L</strong></div>
      <div><span>Total Revenue</span><strong>GH\u20B5 ${fmt(rT)}</strong></div>
      <div><span>PMS Variance</span><strong class="${pV>=0?'var-pos':'var-neg'}">${pV>=0?'+':''}${fmt(pV)} L</strong></div>
      <div><span>AGO Variance</span><strong class="${aV>=0?'var-pos':'var-neg'}">${aV>=0?'+':''}${fmt(aV)} L</strong></div>
      <div><span>Date range</span><strong>${records[0].date} → ${records[records.length-1].date}</strong></div>
    </div>
    <div style="margin-top:8px;font-size:.78rem;color:var(--text-muted)">First 3: ${records.slice(0,3).map(r=>`${r.date}: PMS ${fmt(r.pmsActualSales)}L / AGO ${fmt(r.agoActualSales)}L`).join(' | ')}</div>`;
  document.getElementById('uploadActions').classList.remove('hidden');
}
async function confirmUpload(){
  if(!state.pendingUpload.length)return; showLoading(true);
  try{
    if(state.demoMode){state.pendingUpload.forEach(r=>{r.id='demo-'+Date.now()+Math.random();state.salesData.unshift(r);});renderSalesTable();refreshDashboard();}
    else await Promise.all(state.pendingUpload.map(rec=>ftLib.addDoc(ftLib.collection(db,'sales'),rec)));
    toast(`✅ ${state.pendingUpload.length} records imported!`,'success');
    state.pendingUpload=[];
    document.getElementById('uploadPreview').classList.add('hidden');
    document.getElementById('uploadActions').classList.add('hidden');
    document.getElementById('fileInput').value='';
  }catch(e){toast('Upload failed: '+e.message,'error');}
  showLoading(false);
}

// ──────────────────────────────────────────── USERS & APPROVALS ──────
async function loadUsers(){
  if(state.role!=='corporate'&&state.role!=='admin')return;
  const tbody=document.getElementById('usersTableBody');
  const branchFilter=document.getElementById('userBranchFilter').value;
  const searchTerm=(document.getElementById('userSearch')?.value||'').toLowerCase();

  if(state.demoMode){
    tbody.innerHTML=`
      <tr><td><strong>Demo Corporate</strong></td><td>corporate@fueltrack.com</td><td>—</td><td><span class="badge badge-role-corporate">Corporate HQ</span></td><td>Today</td><td><span class="status-dot status-active">●</span></td><td>—</td></tr>
      <tr><td><strong>Kwame Asante</strong></td><td>kwame@fueltrack.com</td><td><span class="badge badge-branch">Dormaa Ahenkro</span></td><td><span class="badge badge-role-admin">Branch Admin</span></td><td>Jan 2024</td><td><span class="status-dot status-active">●</span></td>
        <td><div class="user-actions"><button class="btn-role" onclick="toast('Demo','warning')">Change Role</button><button class="btn-outline sm" onclick="toast('Demo','warning')">Reassign</button></div></td></tr>
      <tr><td><strong>Kofi Mensah</strong></td><td>kofi@fueltrack.com</td><td><span class="badge badge-branch">Dormaa Ahenkro</span></td><td><span class="badge badge-role-staff">Staff</span></td><td>Mar 2024</td><td><span class="status-dot status-active">●</span></td>
        <td><div class="user-actions"><button class="btn-role" onclick="toast('Demo','warning')">Change Role</button><button class="btn-outline sm" onclick="toast('Demo','warning')">Reassign</button></div></td></tr>`;
    document.getElementById('pendingApprovalsSection').classList.remove('hidden');
    document.getElementById('pendingUsersBody').innerHTML='<tr><td colspan="6" class="empty-row">No pending in demo</td></tr>';
    return;
  }

  try{
    let q=ftLib.query(ftLib.collection(db,'users'));
    if(branchFilter!=='all') q=ftLib.query(ftLib.collection(db,'users'),ftLib.where('branchId','==',branchFilter));
    else if(state.role==='admin'&&state.branchId) q=ftLib.query(ftLib.collection(db,'users'),ftLib.where('branchId','==',state.branchId));
    const snap=await ftLib.getDocs(q);
    let docs=snap.docs.filter(d=>d.data().status!=='inactive');

    // Apply search
    if(searchTerm){
      docs=docs.filter(d=>{
        const u=d.data();
        return (u.name||'').toLowerCase().includes(searchTerm)||(u.email||'').toLowerCase().includes(searchTerm)||(u.branchName||'').toLowerCase().includes(searchTerm);
      });
    }

    if(!docs.length){tbody.innerHTML='<tr><td colspan="7" class="empty-row">No users found</td></tr>';return;}

    // Current user first ────────────────────────────────────────────────────
    docs.sort((a,b)=>{
      if(a.id===state.user.uid) return -1;
      if(b.id===state.user.uid) return 1;
      return 0;
    });

    tbody.innerHTML=docs.map(d=>{
      const u=d.data();
      const isSelf=d.id===state.user?.uid, isCorp=u.role==='corporate';
      // Only corporate HQ can change roles or reassign users
      const canChangeRole = state.role==='corporate';
      const canReassign   = state.role==='corporate';
      const actions=(isCorp||isSelf)?'<span style="color:var(--text-muted);font-size:.78rem">—</span>':`
        <div class="user-actions">
          ${canChangeRole?`<button class="btn-role" onclick="openChangeRoleModal('${d.id}','${(u.name||'').replace(/'/g,"\\'")}','${u.role}')">Change Role</button>`:''}
          ${canReassign?`<button class="btn-outline sm" onclick="openReassignModal('${d.id}','${(u.name||'').replace(/'/g,"\\'")}','${u.branchId||''}')">Reassign</button>`:''}
          <button class="btn-danger" onclick="openDeleteUserModal('${d.id}','${(u.name||'').replace(/'/g,"\\'")}')">Delete</button>
        </div>`;
      return `<tr ${isSelf?'style="background:var(--blue-pale)"':''}>
        <td><strong>${u.name||'—'}</strong>${isSelf?' <span style="font-size:.7rem;color:var(--blue);font-weight:600">(You)</span>':''}</td>
        <td style="font-size:.8rem">${u.email}</td>
        <td><span class="badge badge-branch">${u.branchName||'—'}</span></td>
        <td><span class="badge badge-role-${u.role}">${roleLabel(u.role)}</span></td>
        <td style="font-size:.78rem;color:var(--text-muted)">${u.createdAt?u.createdAt.split('T')[0]:'—'}</td>
        <td><span class="status-dot status-active">●</span></td>
        <td>${actions}</td>
      </tr>`;
    }).join('');
  }catch(e){tbody.innerHTML=`<tr><td colspan="7" class="empty-row">Error: ${e.message}</td></tr>`;}
}

function filterUsers(){loadUsers();}

function renderPendingUsers(users){
  const section=document.getElementById('pendingApprovalsSection');
  const badge=document.getElementById('pendingCountBadge');
  const tbody=document.getElementById('pendingUsersBody');
  section.classList.toggle('hidden',users.length===0);
  if(badge) badge.textContent=users.length+' pending';
  if(!users.length){tbody.innerHTML='<tr><td colspan="6" class="empty-row">No pending requests</td></tr>';return;}

  const searchTerm=(document.getElementById('pendingSearch')?.value||'').toLowerCase();
  let filtered=users;
  if(searchTerm) filtered=users.filter(u=>(u.name||'').toLowerCase().includes(searchTerm)||(u.email||'').toLowerCase().includes(searchTerm)||(u.branchName||'').toLowerCase().includes(searchTerm));

  tbody.innerHTML=filtered.map(u=>`
    <tr>
      <td><strong>${u.name}</strong></td>
      <td style="font-size:.8rem">${u.email}</td>
      <td><span class="badge badge-branch">${u.branchName||'—'}</span></td>
      <td><span class="badge badge-role-${u.role}">${roleLabel(u.role)}</span></td>
      <td style="font-size:.78rem;color:var(--text-muted)">${u.createdAt?u.createdAt.split('T')[0]:'—'}</td>
      <td>
        <button class="btn-approve" onclick="openApproveModal('${u.id}','${(u.name||'').replace(/'/g,"\\'")}','${u.email}','${u.role}','${u.branchId}','${(u.branchName||'').replace(/'/g,"\\'")}')">✅ Approve</button>
        <button class="btn-reject" onclick="openRejectModal('${u.id}','${(u.name||'').replace(/'/g,"\\'")}')">❌ Reject</button>
      </td>
    </tr>`).join('');
}
function filterPending(){
  // Re-render with search from existing data in the DOM isn't reliable — retrigger subscription render
  subscribeToApprovals();
}
function updatePendingBadge(count){const el=document.getElementById('pendingBadge');if(!el)return;el.textContent=count;el.classList.toggle('hidden',count===0);}

// ── Approve / Reject ───────────────────────────────────────────────────────
function openApproveModal(pendingId,name,email,role,branchId,branchName){
  state._pendingToApprove={pendingId,name,email,role,branchId,branchName};
  document.getElementById('approveModalDesc').textContent=`Approving: ${name} (${email}) — ${branchName} as ${roleLabel(role)}`;
  document.getElementById('approveRole').value=role==='admin'?'admin':'staff';
  document.getElementById('approvePassword').value=generateTempPassword();
  document.getElementById('approveModal').classList.remove('hidden');
}
function openRejectModal(pendingId,name){
  state._pendingToReject={pendingId,name};
  document.getElementById('rejectModalDesc').textContent=`Rejecting access request from: ${name}`;
  document.getElementById('rejectReason').value='';
  document.getElementById('rejectModal').classList.remove('hidden');
}
function closeModal(){
  ['approveModal','rejectModal','changeRoleModal','deleteUserModal','bulkDeleteModal',
   'deleteBranchModal','branchDetailModal','reassignModal'].forEach(id=>{
    const el=document.getElementById(id); if(el) el.classList.add('hidden');
  });
  state._pendingToApprove=null;state._pendingToReject=null;state._userToChangeRole=null;
  state._userToDelete=null;state._branchToDelete=null;state._branchDetailId=null;state._userToReassign=null;
}

async function confirmApproval(){
  if(!state._pendingToApprove)return;
  const{pendingId,name,email,branchId,branchName}=state._pendingToApprove;
  const password=document.getElementById('approvePassword').value;
  const role=document.getElementById('approveRole').value;
  if(!password||password.length<6)return toast('Password must be at least 6 characters','error');
  const adminUser=state.user;
  showLoading(true); closeModal();
  try{
    const cred=await ftLib.createUserWithEmailAndPassword(authSecondary,email,password);
    const newUid=cred.user.uid;
    await ftLib.signOut(authSecondary);
    await ftLib.setDoc(ftLib.doc(db,'users',newUid),{
      name,email,role,branchId,branchName,
      status:'active',
      passwordChanged:false,   // ← user must set a new password on first login
      createdAt:new Date().toISOString(),
      approvedBy:adminUser.uid
    });
    await ftLib.updateDoc(ftLib.doc(db,'pendingUsers',pendingId),{status:'approved',approvedAt:new Date().toISOString(),tempPassword:password});
    toast(`✅ ${name} approved! Temp password: ${password}`,'success');
    loadUsers();
  }catch(e){
    if(e.code==='auth/email-already-in-use') toast(`Account for ${email} already exists`,'error');
    else toast('Approval failed: '+e.message,'error');
  }
  showLoading(false);
}

async function confirmRejection(){
  if(!state._pendingToReject)return;
  const{pendingId,name}=state._pendingToReject;
  const reason=document.getElementById('rejectReason').value.trim()||'Request not approved by Corporate HQ.';
  showLoading(true); closeModal();
  try{
    await ftLib.updateDoc(ftLib.doc(db,'pendingUsers',pendingId),{status:'rejected',rejectionReason:reason,rejectedAt:new Date().toISOString()});
    toast(`Request from ${name} rejected`,'success');
  }catch(e){toast('Error: '+e.message,'error');}
  showLoading(false);
}

// ── Change role — only corporate can do this; createdAt is NOT updated ──────
function openChangeRoleModal(uid,name,currentRole){
  if(state.role!=='corporate') return toast('Only Corporate HQ can change roles','error');
  state._userToChangeRole={uid,name};
  document.getElementById('changeRoleDesc').textContent=`Change role for ${name} (currently: ${roleLabel(currentRole)})`;
  document.getElementById('changeRoleSelect').value=currentRole==='admin'?'admin':'staff';
  document.getElementById('changeRoleModal').classList.remove('hidden');
}
async function confirmRoleChange(){
  if(!state._userToChangeRole)return;
  const{uid,name}=state._userToChangeRole;
  const newRole=document.getElementById('changeRoleSelect').value;
  showLoading(true); closeModal();
  try{
    if(!state.demoMode){
      // ONLY update role — do NOT touch createdAt so joined date is preserved
      await ftLib.updateDoc(ftLib.doc(db,'users',uid),{role:newRole});
    }
    toast(`✅ ${name}'s role updated to ${roleLabel(newRole)}`,'success');
    loadUsers();
  }catch(e){toast('Error: '+e.message,'error');}
  showLoading(false);
}

// ── Delete / deactivate user ───────────────────────────────────────────────
function openDeleteUserModal(uid,name){
  state._userToDelete={uid,name};
  document.getElementById('deleteUserDesc').textContent=`Deactivate "${name}"? They will no longer be able to sign in.`;
  document.getElementById('deleteUserModal').classList.remove('hidden');
}
async function confirmDeleteUser(){
  if(!state._userToDelete)return;
  const{uid,name}=state._userToDelete;
  closeModal(); showLoading(true);
  try{
    if(!state.demoMode) await ftLib.updateDoc(ftLib.doc(db,'users',uid),{status:'inactive'});
    toast(`${name} deactivated`,'success'); loadUsers();
  }catch(e){toast('Error: '+e.message,'error');}
  showLoading(false);
}

function generateTempPassword(){
  const c='ABCDEFGHJKMNPQRSTUVWXYZabcdefghjkmnpqrstuvwxyz23456789';
  return Array.from({length:10},()=>c[Math.floor(Math.random()*c.length)]).join('');
}

// Generic confirm modal ────────────────────────────────────────────────────
function showConfirmModal(title,desc,onConfirm,danger=false){
  document.getElementById('confirmModalTitle').textContent=title;
  document.getElementById('confirmModalDesc').textContent=desc;
  const btn=document.getElementById('confirmModalOk');
  btn.className=danger?'btn-danger-full':'btn-primary';
  btn.onclick=()=>{closeConfirmModal();onConfirm();};
  document.getElementById('confirmModal').classList.remove('hidden');
}
function closeConfirmModal(){document.getElementById('confirmModal').classList.add('hidden');}

// ──────────────────────────────────────────────────────────── REPORTS ──────
function getReportData(){
  const period=document.getElementById('reportPeriod').value;
  const branchFilter=document.getElementById('reportBranch').value;
  const now=new Date(); let from,to=now;
  if(period==='custom'){
    from=new Date(document.getElementById('reportFrom').value||now);
    to=new Date(document.getElementById('reportTo').value||now);
  } else from=new Date(now-parseInt(period)*86400000);
  return state.salesData.filter(s=>{
    const d=new Date(s.date);
    return d>=from&&d<=to&&(branchFilter==='all'||s.branchId===branchFilter);
  });
}

// Live preview — called automatically when report options change ─────────
function previewReport(){
  const data=getReportData();
  if(!data.length){
    document.getElementById('reportPreview').innerHTML='<p class="preview-placeholder">No data for selected options.</p>';
    return;
  }
  const tPR=data.reduce((s,r)=>s+(r.pmsRevenue||0),0);
  const tAR=data.reduce((s,r)=>s+(r.agoRevenue||0),0);
  const tPQ=data.reduce((s,r)=>s+(r.pmsActualSales||0),0);
  const tAQ=data.reduce((s,r)=>s+(r.agoActualSales||0),0);
  const tPV=data.reduce((s,r)=>s+(r.pmsVariance||0),0);
  const tAV=data.reduce((s,r)=>s+(r.agoVariance||0),0);
  const tRev=tPR+tAR;
  document.getElementById('reportPreview').innerHTML=`
    <h4 style="margin-bottom:12px;color:var(--blue);font-size:.9rem">📊 Live Preview — ${data.length} records</h4>
    <div class="rpt-metric-grid">
      <div class="rpt-metric-item"><div class="rpt-val">GH\u20B5 ${fmt(tRev)}</div><div class="rpt-lbl">Total Revenue</div></div>
      <div class="rpt-metric-item"><div class="rpt-val">${fmt(tPQ+tAQ)} L</div><div class="rpt-lbl">Total Volume</div></div>
      <div class="rpt-metric-item"><div class="rpt-val">GH\u20B5 ${fmt(tPR)}</div><div class="rpt-lbl">PMS Revenue</div></div>
      <div class="rpt-metric-item"><div class="rpt-val">GH\u20B5 ${fmt(tAR)}</div><div class="rpt-lbl">AGO Revenue</div></div>
      <div class="rpt-metric-item"><div class="rpt-val">${fmt(tPQ)} L</div><div class="rpt-lbl">PMS Sales</div></div>
      <div class="rpt-metric-item"><div class="rpt-val">${fmt(tAQ)} L</div><div class="rpt-lbl">AGO Sales</div></div>
      <div class="rpt-metric-item"><div class="rpt-val" style="color:${tPV>=0?'var(--accent-green)':'var(--accent-red)'}">${tPV>=0?'+':''}${fmt(tPV)} L</div><div class="rpt-lbl">PMS Variance</div></div>
      <div class="rpt-metric-item"><div class="rpt-val" style="color:${tAV>=0?'var(--accent-green)':'var(--accent-red)'}">${tAV>=0?'+':''}${fmt(tAV)} L</div><div class="rpt-lbl">AGO Variance</div></div>
    </div>
    <p style="margin:12px 0 6px;font-size:.8rem;font-weight:600;color:var(--text-secondary)">Latest 5 records:</p>
    <table>
      <thead><tr><th>Date</th><th>Branch</th><th>PMS (L)</th><th>AGO (L)</th><th>Revenue</th></tr></thead>
      <tbody>${data.slice(0,5).map(r=>`<tr><td>${r.date}</td><td>${r.branchName||'—'}</td><td>${fmt(r.pmsActualSales)}</td><td>${fmt(r.agoActualSales)}</td><td>GH\u20B5 ${fmt((r.pmsRevenue||0)+(r.agoRevenue||0))}</td></tr>`).join('')}</tbody>
    </table>`;
}

function generatePDF(){
  const data=getReportData();
  if(!data.length)return toast('No data for selected period','warning');
  const{jsPDF}=window.jspdf; const doc=new jsPDF({unit:'mm',format:'a4'});
  const now=new Date(), GHC='GHC';
  const branchLabel=document.getElementById('reportBranch').options[document.getElementById('reportBranch').selectedIndex]?.text||'All Branches';
  const periodLabel=document.getElementById('reportPeriod').options[document.getElementById('reportPeriod').selectedIndex]?.text||'';
  doc.setFillColor(15,34,68);doc.rect(0,0,210,46,'F');
  doc.setFillColor(30,58,138);doc.circle(20,23,10,'F');
  doc.setTextColor(255,255,255);doc.setFontSize(14);doc.setFont('helvetica','bold');doc.text('F',17.5,26);
  doc.setFontSize(20);doc.text('FuelTrack',34,20);
  doc.setFontSize(10);doc.setFont('helvetica','normal');doc.text('Multi-Branch Sales Analytics Report',34,28);
  doc.setFontSize(9);doc.text(`Generated: ${now.toLocaleDateString('en-GH')}`,135,18);
  doc.text(`Period: ${periodLabel}`,135,25);doc.text(`Branch: ${branchLabel}`,135,32);
  doc.setFillColor(59,130,246);doc.rect(0,46,210,2,'F');
  const tP=data.reduce((s,r)=>s+(r.pmsActualSales||0),0),tA=data.reduce((s,r)=>s+(r.agoActualSales||0),0);
  const tPR=data.reduce((s,r)=>s+(r.pmsRevenue||0),0),tAR=data.reduce((s,r)=>s+(r.agoRevenue||0),0);
  const tRev=tPR+tAR;
  const tPV=data.reduce((s,r)=>s+(r.pmsVariance||0),0),tAV=data.reduce((s,r)=>s+(r.agoVariance||0),0);
  doc.setTextColor(30,30,30);doc.setFontSize(11);doc.setFont('helvetica','bold');doc.text('SUMMARY METRICS',15,58);
  doc.setDrawColor(200,200,200);doc.line(15,60,195,60);
  const kpis=[[`${GHC} ${fmt(tRev)}`,'Total Revenue'],[`${fmt(tP)} L`,'PMS Sales'],[`${fmt(tA)} L`,'AGO Sales'],[`${GHC} ${fmt(tPR)}`,'PMS Revenue'],[`${GHC} ${fmt(tAR)}`,'AGO Revenue'],[`${tPV>=0?'+':''}${fmt(tPV)} L`,'PMS Variance'],[`${tAV>=0?'+':''}${fmt(tAV)} L`,'AGO Variance'],[data.length.toString(),'Records']];
  const bW=44,bH=20,sX=15,sY=64,gX=45;
  kpis.forEach(([val,lbl],i)=>{const col=i%4,row=Math.floor(i/4),x=sX+col*gX,y=sY+row*22;doc.setFillColor(248,250,252);doc.setDrawColor(220,220,220);doc.rect(x,y,bW,bH,'FD');doc.setFontSize(10);doc.setFont('helvetica','bold');doc.setTextColor(15,34,68);doc.text(val,x+3,y+8);doc.setFontSize(7.5);doc.setFont('helvetica','normal');doc.setTextColor(110,110,110);doc.text(lbl,x+3,y+15);});
  const tY=sY+2*22+10;doc.setFontSize(11);doc.setFont('helvetica','bold');doc.setTextColor(30,30,30);doc.text('TRANSACTION DETAILS',15,tY);
  doc.autoTable({startY:tY+4,head:[['Date','Branch','PMS (L)',`PMS Rev (${GHC})`,'AGO (L)',`AGO Rev (${GHC})`,'PMS Var','AGO Var']],
    body:data.slice(0,100).map(r=>[r.date,r.branchName||'—',fmt(r.pmsActualSales),fmt(r.pmsRevenue),fmt(r.agoActualSales),fmt(r.agoRevenue),(r.pmsVariance>=0?'+':'')+fmt(r.pmsVariance),(r.agoVariance>=0?'+':'')+fmt(r.agoVariance)]),
    headStyles:{fillColor:[15,34,68],fontStyle:'bold',fontSize:8,textColor:255},bodyStyles:{fontSize:8,textColor:[30,30,30]},alternateRowStyles:{fillColor:[248,250,252]},
    columnStyles:{2:{halign:'right'},3:{halign:'right'},4:{halign:'right'},5:{halign:'right'},6:{halign:'right'},7:{halign:'right'}},
    margin:{left:15,right:15},
    didParseCell:(h)=>{if(h.column.index===6||h.column.index===7){const v=parseFloat(h.cell.text[0]);if(!isNaN(v))h.cell.styles.textColor=v>=0?[5,150,105]:[220,38,38];}}
  });
  const pc=doc.internal.getNumberOfPages();
  for(let i=1;i<=pc;i++){doc.setPage(i);doc.setFillColor(248,250,252);doc.rect(0,doc.internal.pageSize.height-12,210,12,'F');doc.setFontSize(8);doc.setTextColor(130,130,130);doc.setFont('helvetica','normal');doc.text(`FuelTrack Analytics — Confidential — Page ${i} of ${pc}`,15,doc.internal.pageSize.height-5);doc.text(`Generated ${now.toLocaleString('en-GH')}`,140,doc.internal.pageSize.height-5);}
  doc.save(`FuelTrack_Report_${now.toISOString().split('T')[0]}.pdf`);
  toast('PDF downloaded!','success');
}

function generateExcel(){
  const data=getReportData();if(!data.length)return toast('No data','warning');
  const tPR=data.reduce((s,r)=>s+(r.pmsRevenue||0),0),tAR=data.reduce((s,r)=>s+(r.agoRevenue||0),0);
  const sum=[{Metric:'Total Revenue',Value:`GHC ${fmt(tPR+tAR)}`},{Metric:'PMS Sales',Value:`${fmt(data.reduce((s,r)=>s+(r.pmsActualSales||0),0))} L`},{Metric:'AGO Sales',Value:`${fmt(data.reduce((s,r)=>s+(r.agoActualSales||0),0))} L`},{Metric:'PMS Revenue',Value:`GHC ${fmt(tPR)}`},{Metric:'AGO Revenue',Value:`GHC ${fmt(tAR)}`},{Metric:'PMS Variance',Value:`${fmt(data.reduce((s,r)=>s+(r.pmsVariance||0),0))} L`},{Metric:'AGO Variance',Value:`${fmt(data.reduce((s,r)=>s+(r.agoVariance||0),0))} L`},{Metric:'Records',Value:data.length},{Metric:'Generated',Value:new Date().toLocaleDateString('en-GH')}];
  const trans=data.map(r=>({Date:r.date,Branch:r.branchName||'','PMS Sales (L)':r.pmsActualSales||0,'PMS Revenue (GHC)':r.pmsRevenue||0,'AGO Sales (L)':r.agoActualSales||0,'AGO Revenue (GHC)':r.agoRevenue||0,'PMS Variance (L)':r.pmsVariance||0,'AGO Variance (L)':r.agoVariance||0,Notes:r.notes||''}));
  const wb=XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(sum),'Summary');
  XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(trans),'Transactions');
  XLSX.writeFile(wb,`FuelTrack_Report_${new Date().toISOString().split('T')[0]}.xlsx`);
  toast('Excel downloaded!','success');
}

// ────────────────────────────────────────────────────────── DEMO DATA ──────
function generateDemoBranches(){
  return[{id:'b1',name:'Dormaa Ahenkro Station',code:'DOR-01',region:'Brong-Ahafo',address:'Dormaa Ahenkro',phone:'0244-000001',active:true},
    {id:'b2',name:'Kumasi Central Station',code:'KSI-01',region:'Ashanti',address:'Adum, Kumasi',phone:'0244-000002',active:true},
    {id:'b3',name:'Accra Ring Road Station',code:'ACC-01',region:'Greater Accra',address:'Ring Road West, Accra',phone:'0244-000003',active:true}];
}
function generateDemoData(){
  const branches=generateDemoBranches(),data=[],today=new Date();
  branches.forEach((b,bi)=>{
    const pB=600+bi*200,aB=900+bi*150;
    for(let i=90;i>=0;i--){
      const date=new Date(today-i*86400000),ds=date.toISOString().split('T')[0];
      const pQ=+(pB*(0.75+Math.random()*0.5)).toFixed(1),aQ=+(aB*(0.75+Math.random()*0.5)).toFixed(1);
      const pP=+(14.2+Math.random()*0.4).toFixed(2),aP=+(12.8+Math.random()*0.3).toFixed(2);
      data.push({id:`demo-${b.id}-${i}`,date:ds,branchId:b.id,branchName:b.name,pmsActualSales:pQ,pmsRevenue:+(pQ*pP).toFixed(2),agoActualSales:aQ,agoRevenue:+(aQ*aP).toFixed(2),pmsVariance:+((Math.random()-0.45)*30).toFixed(2),agoVariance:+((Math.random()-0.45)*20).toFixed(2),source:'demo'});
    }
  });
  return data.reverse();
}

// ─────────────────────────────────────────────────────────────── HELPERS ──
function fmt(n){return(parseFloat(n)||0).toLocaleString('en-GH',{minimumFractionDigits:2,maximumFractionDigits:2});}
function fmtK(v){if(Math.abs(v)>=1000000)return(v/1000000).toFixed(1)+'M';if(Math.abs(v)>=1000)return(v/1000).toFixed(1)+'K';return parseFloat(v).toFixed(0);}

// Short label for charts: "15 Jan"
function formatDateLabel(ds){
  if(!ds)return'';
  return new Date((ds+'').split('T')[0]+'T00:00:00').toLocaleDateString('en-GH',{day:'numeric',month:'short'});
}
// Full label with year for tables: "15 Jan 2024"
function formatDateFull(ds){
  if(!ds)return'';
  return new Date((ds+'').split('T')[0]+'T00:00:00').toLocaleDateString('en-GH',{day:'numeric',month:'short',year:'numeric'});
}

function toast(msg,type='info'){const el=document.getElementById('toast');el.textContent=msg;el.className=`toast ${type}`;el.classList.remove('hidden');clearTimeout(el._t);el._t=setTimeout(()=>el.classList.add('hidden'),4800);}
let _loadingTimer = null;
function showLoading(show) {
  document.getElementById('loadingOverlay').classList.toggle('hidden', !show);
  if (show) {
    // Auto-hide after 8 s to prevent permanently stuck spinner
    if (_loadingTimer) clearTimeout(_loadingTimer);
    _loadingTimer = setTimeout(() => {
      document.getElementById('loadingOverlay').classList.add('hidden');
      _loadingTimer = null;
    }, 8000);
  } else {
    if (_loadingTimer) { clearTimeout(_loadingTimer); _loadingTimer = null; }
  }
}

// ───────────────────────────────────────── CHANGE PASSWORD (new accounts) ──
// Show a persistent banner prompting the user to set a personal password
function checkPasswordChangeBanner(userDoc) {
  // Only show for non-corporate accounts that were explicitly flagged as needing a password change.
  // Corporate HQ never uses temp passwords. Existing accounts without the field are ignored.
  if (!userDoc) return;
  if (userDoc.role === 'corporate') return;              // corporate: never
  if (userDoc.passwordChanged !== false) return;        // only show when explicitly set to false
  const existing = document.getElementById('cpBanner');
  if (existing) return;
  const banner = document.createElement('div');
  banner.id = 'cpBanner';
  banner.className = 'cp-banner';
  banner.innerHTML = `
    <span class="cp-banner-icon">🔐</span>
    <div class="cp-banner-text">
      <strong>You are using a temporary password.</strong>
      Please set a secure personal password now.
    </div>
    <button class="btn-primary" style="font-size:.8rem;padding:7px 16px;flex-shrink:0" onclick="openChangePasswordModal()">
      Change Password
    </button>
    <button class="cp-banner-dismiss" onclick="document.getElementById('cpBanner').remove()" title="Dismiss">✕</button>`;
  const topbar = document.querySelector('.topbar');
  if (topbar && topbar.parentNode) topbar.parentNode.insertBefore(banner, topbar.nextSibling);
}

function openChangePasswordModal() {
  ['cpNew','cpConfirm'].forEach(id => document.getElementById(id).value = '');
  const err = document.getElementById('cpError');
  err.textContent = ''; err.classList.add('hidden');
  document.getElementById('changePasswordModal').classList.remove('hidden');
}

async function confirmChangePassword() {
  const newPw     = document.getElementById('cpNew').value;
  const confirmPw = document.getElementById('cpConfirm').value;
  const err       = document.getElementById('cpError');

  const showErr = (msg) => { err.textContent = msg; err.classList.remove('hidden'); };

  if (!newPw)               return showErr('Please enter a new password.');
  if (newPw.length < 8)     return showErr('Password must be at least 8 characters.');
  if (!/[A-Z]/.test(newPw)) return showErr('Password must contain at least one uppercase letter.');
  if (!/[0-9]/.test(newPw)) return showErr('Password must contain at least one number.');
  if (newPw !== confirmPw)  return showErr('Passwords do not match.');

  showLoading(true);
  try {
    // updatePassword is a Firebase Auth function on the current user
    await ftLib.updatePassword(auth.currentUser, newPw);

    // Mark in Firestore so the banner doesn't reappear
    if (!state.demoMode) {
      await ftLib.updateDoc(ftLib.doc(db,'users',state.user.uid), { passwordChanged: true });
    }

    document.getElementById('changePasswordModal').classList.add('hidden');
    const banner = document.getElementById('cpBanner');
    if (banner) banner.remove();
    toast('✅ Password updated! Keep it safe.', 'success');
  } catch(e) {
    if (e.code === 'auth/requires-recent-login') {
      showErr('For security, please sign out then sign back in before changing your password.');
    } else {
      showErr('Error: ' + e.message);
    }
  }
  showLoading(false);
}
