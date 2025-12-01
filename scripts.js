/* ========================
           CONFIG (SET YOUR VALUES)
           ======================== */
const tenantId = '1b2bcaa0-2052-4943-bd9b-2511e7d57a21';
const clientId = 'c34ceb14-3218-4ec5-8bd4-39a66ede3603';
const redirectUri = "https://flogerber.github.io/EntraIdUserMangement/";

const msalConfig = {
    auth: { clientId, authority: `https://login.microsoftonline.com/${tenantId}`, redirectUri },
    cache: { cacheLocation: "localStorage", storeAuthStateInCookie: false }
};

const GRAPH_SCOPES = ["User.Read.All", "Directory.Read.All", "Group.Read.All"];
const msalInstance = new msal.PublicClientApplication(msalConfig);

/* ========================
   FIELDS / STATE
   ======================== */
const BASE_FIELDS = ["displayName", "userPrincipalName", "mail", "mobile", "accountEnabled"];
const EXTRA_FIELDS = ["jobTitle", "department", "companyName", "usageLocation", "employeeId", "employeeType", "country", "officeLocation", "streetAddress", "id"];
const ALL_FIELDS = [...BASE_FIELDS, ...EXTRA_FIELDS, "id"];

const FRIENDLY = {
    displayName: "Display Name", userPrincipalName: "User Principal Name", mail: "Email", mobile: "Mobile",
    jobTitle: "Job Title", department: "Department", companyName: "Company", accountEnabled: "Enabled",
    usageLocation: "Usage Location", employeeId: "Employee ID", employeeType: "Employee Type",
    country: "Country", officeLocation: "Office", streetAddress: "Address", id: "Object ID"
};

let allUsers = [];
let visibleColumns = [...BASE_FIELDS];
let activeFilters = [];
let sortField = null, sortDir = 1;

/* Token cache in-memory */
let cachedToken = null; // {accessToken, expiresAt (epoch ms)}
const TOKEN_BUFFER = 60; // seconds

/* Cache for subscribedSkus and roles */
let cachedSubscribedSkus = null;
let cachedDirectoryRoles = null;

/* Populate small icons */
function populateStaticIcons() {
    const set = (id, name) => { const el = document.getElementById(id); if (el) el.innerHTML = svg(name); };
    set("iconSearch", "search"); set("iconFilter", "filter"); set("iconColumns", "columns");
    set("detailUserIcon", "user");
    document.querySelectorAll(".icon-close").forEach(s => s.innerHTML = svg("close"));
}

/* ========================
   MSAL AUTH HELPERS
   ======================== */
async function initializeAuth() {
    try {
        const redirectResult = await msalInstance.handleRedirectPromise();
        if (redirectResult && redirectResult.account) {
            msalInstance.setActiveAccount(redirectResult.account);
        } else {
            const accounts = msalInstance.getAllAccounts();
            if (accounts.length === 1) msalInstance.setActiveAccount(accounts[0]);
            else if (accounts.length > 1) msalInstance.setActiveAccount(accounts[0]); // choose first by default
        }
        if (!msalInstance.getActiveAccount()) {
            // start interactive redirect login
            await msalInstance.loginRedirect({ scopes: GRAPH_SCOPES });
            return false;
        }
        return true;
    } catch (err) {
        console.error("initializeAuth error", err);
        try { await msalInstance.loginRedirect({ scopes: GRAPH_SCOPES }); return false; } catch (e) { console.error("loginRedirect failed", e); return false; }
    }
}

/* getToken: uses cachedToken if not expired; otherwise acquireTokenSilent -> popup -> redirect fallback */
async function getToken() {
    const now = Date.now();
    if (cachedToken && cachedToken.accessToken && cachedToken.expiresAt && (cachedToken.expiresAt - (TOKEN_BUFFER * 1000) > now)) {
        return cachedToken.accessToken;
    }

    const account = msalInstance.getActiveAccount();
    if (!account) {
        await msalInstance.loginRedirect({ scopes: GRAPH_SCOPES });
        return null;
    }

    try {
        const res = await msalInstance.acquireTokenSilent({ scopes: GRAPH_SCOPES, account });
        storeToken(res);
        return res.accessToken;
    } catch (silentErr) {
        console.warn("acquireTokenSilent failed:", silentErr);
        try {
            const pop = await msalInstance.acquireTokenPopup({ scopes: GRAPH_SCOPES, account });
            storeToken(pop);
            return pop.accessToken;
        } catch (popupErr) {
            console.warn("acquireTokenPopup failed:", popupErr, "falling back to redirect");
            try {
                await msalInstance.acquireTokenRedirect({ scopes: GRAPH_SCOPES, account });
                return null;
            } catch (redErr) {
                console.error("acquireTokenRedirect failed", redErr);
                return null;
            }
        }
    }
}

function storeToken(result) {
    if (!result) return;
    const expiresAt = (result.expiresOn && result.expiresOn.getTime) ? result.expiresOn.getTime() : (Date.now() + (60 * 60 * 1000));
    cachedToken = { accessToken: result.accessToken, expiresAt };
}

/* ========================
   Graph helpers
   ======================== */
async function fetchAllPages(url, token) {
    const items = [];
    let next = url;
    while (next) {
        const res = await fetch(next, { headers: { Authorization: `Bearer ${token}` } });
        if (!res.ok) {
            const txt = await res.text();
            throw new Error(`Graph fetch failed: ${res.status} ${txt}`);
        }
        const data = await res.json();
        items.push(...(data.value || []));
        next = data["@odata.nextLink"] || null;
    }
    return items;
}

/* subscribedSkus mapping (cached) */
async function loadSubscribedSkus(token) {
    if (cachedSubscribedSkus) return cachedSubscribedSkus;
    const url = `https://graph.microsoft.com/v1.0/subscribedSkus`;
    const res = await fetch(url, { headers: { Authorization: `Bearer ${token}` } });
    if (!res.ok) throw new Error('failed to load subscribedSkus');
    const data = await res.json();
    // Map skuId -> friendly (skuPartNumber or displayName)
    const map = {};
    (data.value || []).forEach(s => {
        map[s.skuId] = s.skuPartNumber || (s.prepaidUnits && s.skuId) || s.skuId;
    });
    cachedSubscribedSkus = map;
    return map;
}

/* load directory roles (cached) */
async function loadDirectoryRoles(token) {
    if (cachedDirectoryRoles) return cachedDirectoryRoles;
    const url = `https://graph.microsoft.com/v1.0/directoryRoles?$select=id,displayName`;
    const res = await fetch(url, { headers: { Authorization: `Bearer ${token}` } });
    if (!res.ok) { cachedDirectoryRoles = {}; return cachedDirectoryRoles; }
    const data = await res.json();
    const map = {};
    (data.value || []).forEach(r => map[r.id] = r.displayName || r.id);
    cachedDirectoryRoles = map;
    return map;
}

/* load user groups (memberOf) - returns array of objects with displayName, id, groupTypes, mailEnabled, securityEnabled */
async function loadUserGroups(userId, token) {
    const url = `https://graph.microsoft.com/v1.0/users/${encodeURIComponent(userId)}/memberOf?$select=id,displayName,groupTypes,mailEnabled,securityEnabled`;
    return await fetchAllPages(url, token);
}

/* load user licenses (licenseDetails) */
async function loadUserLicenses(userId, token) {
    const url = `https://graph.microsoft.com/v1.0/users/${encodeURIComponent(userId)}/licenseDetails`;
    return await fetchAllPages(url, token);
}

/* ========================
   Load users (U1 - all up front)
   ======================== */
async function loadUsers() {
    const token = await getToken();
    if (!token) return;
    // fetch all users (U1)
    const url = `https://graph.microsoft.com/v1.0/users?$select=${ALL_FIELDS.join(",")}&$top=999`;
    try {
        allUsers = await fetchAllPages(url, token);
    } catch (e) {
        console.error("Failed to load users", e);
        allUsers = [];
    }
    document.getElementById("userCount").textContent = `${allUsers.length} users`;
    renderTable(allUsers);
    populateFilterFields();
    populateColumnDrawer();
}

/* ========================
   Table rendering
   ======================== */
function renderTable(data) {
    const head = document.getElementById("tableHead"), body = document.getElementById("tableBody");
    if (!head || !body) return;
    head.innerHTML = `<tr>${visibleColumns.map(c => `<th onclick="sortBy('${c}')">${escapeHtml(FRIENDLY[c] ?? c)} ${sortField === c ? (sortDir === 1 ? '▲' : '▼') : ''}</th>`).join("")}</tr>`;
    const rows = (data || []).map(u => `<tr onclick="openUserDetails('${encodeURIComponent(u.id || '')}')">${visibleColumns.map(f => formatCell(f, u)).join("")}</tr>`).join("");
    body.innerHTML = rows;
}

function formatCell(field, user) {
    const safe = (v) => escapeHtml(String(v === undefined || v === null ? "" : v));
    if (field === 'accountEnabled') {
        const v = user?.[field];
        return `<td><span class="badge ${v ? 'active' : 'disabled'}">${v ? 'Active' : 'Disabled'}</span></td>`;
    }
    if (field === 'displayName') {
        const initials = (user?.displayName || '').split(/\s+/).map(s => s[0]).join('').slice(0, 2).toUpperCase();
        return `<td class="col-name"><span>${safe(user?.displayName)}</span></td>`;
    }
    return `<td>${safe(user?.[field])}</td>`;
}

function escapeHtml(str) { return (str || '').replace(/[&<>"'`=\/]/g, s => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;', '`': '&#96;', '=': '&#61;', '/': '&#47;' })[s]); }

/* ========================
   Sorting
   ======================== */
function sortBy(field) {
    if (sortField === field) sortDir *= -1; else { sortField = field; sortDir = 1; }
    const sorted = [...(allUsers || [])].sort((a, b) => {
        const x = ((a[field] || '') + "").toLowerCase();
        const y = ((b[field] || '') + "").toLowerCase();
        if (x > y) return sortDir; if (x < y) return -sortDir; return 0;
    });
    renderTable(sorted);
}

/* ========================
   Search & Filters (multi-field search)
   ======================== */
function applySearchAndFilters() {
    const q = (document.getElementById("searchInput")?.value || '').toLowerCase().trim();
    const data = (allUsers || []).filter(u => {
        // Multi-field search: name, mail, upn, department, companyName, officeLocation
        const text = `${u?.displayName || ''} ${u?.mail || ''} ${u?.userPrincipalName || ''} ${u?.department || ''} ${u?.companyName || ''} ${u?.officeLocation || ''}`.toLowerCase();
        if (q && !text.includes(q)) return false;
        return activeFilters.every(f => applyFilter(u, f));
    });
    renderTable(data);
}

function applyFilter(user, f) {
    const val = ((user?.[f.field] || '') + '').toLowerCase();
    const needle = (f.value || '').toLowerCase();
    switch (f.op) {
        case 'contains': return val.includes(needle);
        case 'equals': return val === needle;
        case 'starts': return val.startsWith(needle);
        case 'ends': return val.endsWith(needle);
        case 'empty': return val === '';
        case 'notempty': return val !== '';
        default: return true;
    }
}

/* ========================
   Filter drawer logic (add/edit/remove)
   ======================== */
function populateFilterFields() {
    const sel = document.getElementById("filterFieldSelector");
    if (!sel) return;
    sel.innerHTML = ALL_FIELDS.map(f => `<option value="${f}">${escapeHtml(FRIENDLY[f] || f)}</option>`).join('');
}

function addFilterCondition() {
    const field = document.getElementById("filterFieldSelector")?.value;
    const op = document.getElementById("filterOperator")?.value;
    const value = (document.getElementById("filterValue")?.value || '').trim();
    if (!field) return;
    if (op !== 'empty' && op !== 'notempty' && value === '') { alert("Please provide a value (or choose is empty/is not empty)"); return; }
    activeFilters.push({ field, op, value });
    redrawActiveFilters();
    applySearchAndFilters();
}

function redrawActiveFilters() {
    const wrap = document.getElementById("activeFilters");
    if (!wrap) return;
    wrap.innerHTML = activeFilters.map((f, idx) => `
    <div class="chip" id="chip-${idx}">
      <span style="font-weight:600">${escapeHtml(FRIENDLY[f.field] || f.field)}</span>
      <span>:</span>
      <span>${escapeHtml(f.op)}</span>
      <span>"${escapeHtml(f.value || '')}"</span>
      <button class="muted-btn" onclick="editFilter(${idx})">Edit</button>
      <button class="muted-btn" onclick="removeFilter(${idx})">Remove</button>
    </div>
  `).join('');
}

function removeFilter(i) { if (i < 0 || i >= activeFilters.length) return; activeFilters.splice(i, 1); redrawActiveFilters(); applySearchAndFilters(); }
function editFilter(idx) {
    const f = activeFilters[idx]; if (!f) return;
    const container = document.getElementById(`chip-${idx}`); if (!container) return;
    container.innerHTML = `
    <div style="display:flex;gap:8px;align-items:center;">
      <select id="editField-${idx}" style="min-width:160px">${ALL_FIELDS.map(ff => `<option value="${ff}" ${ff === f.field ? 'selected' : ''}>${escapeHtml(FRIENDLY[ff] || ff)}</option>`).join('')}</select>
      <select id="editOp-${idx}" style="min-width:140px">
        <option value="contains" ${f.op === 'contains' ? 'selected' : ''}>contains</option>
        <option value="equals" ${f.op === 'equals' ? 'selected' : ''}>equals</option>
        <option value="starts" ${f.op === 'starts' ? 'selected' : ''}>starts with</option>
        <option value="ends" ${f.op === 'ends' ? 'selected' : ''}>ends with</option>
        <option value="empty" ${f.op === 'empty' ? 'selected' : ''}>is empty</option>
        <option value="notempty" ${f.op === 'notempty' ? 'selected' : ''}>is not empty</option>
      </select>
      <input id="editVal-${idx}" value="${escapeHtml(f.value || '')}" placeholder="value"/>
      <button class="btn" onclick="saveEditedFilter(${idx})">Save</button>
      <button class="btn ghost" onclick="cancelEdit(${idx})">Cancel</button>
    </div>
  `;
}
function saveEditedFilter(idx) {
    const field = document.getElementById(`editField-${idx}`)?.value;
    const op = document.getElementById(`editOp-${idx}`)?.value;
    const value = (document.getElementById(`editVal-${idx}`)?.value || '').trim();
    if (!field) return;
    if (op !== 'empty' && op !== 'notempty' && value === '') { alert('Please enter a value'); return; }
    activeFilters[idx] = { field, op, value };
    redrawActiveFilters();
    applySearchAndFilters();
}
function cancelEdit(idx) { redrawActiveFilters(); }
function applyFilters() { applySearchAndFilters(); closeFilterDrawer(); }

/* ========================
   Column drawer
   ======================== */
function populateColumnDrawer() {
    const wrap = document.getElementById("columnList");
    if (!wrap) return;
    wrap.innerHTML = ALL_FIELDS.map(f => `
    <label style="display:block;padding:6px 0;">
      <input type="checkbox" ${visibleColumns.includes(f) ? 'checked' : ''} onchange="toggleColumn('${f}', this.checked)"/>
      ${escapeHtml(FRIENDLY[f] || f)}
    </label>
  `).join('');
}
function toggleColumn(field, show) { if (show) { if (!visibleColumns.includes(field)) visibleColumns.push(field); } else visibleColumns = visibleColumns.filter(x => x !== field); }
function saveColumnSelection() { closeColumnDrawer(); renderTable(allUsers); }

/* ========================
   Detail drawer: skeleton + enhanced data
   ======================== */
function openDetailSkeleton(title, subtitle) {
    const wrapper = document.getElementById("detailGridWrapper");
    if (!wrapper) return;
    document.getElementById("detailTitle").textContent = title || "User Details";
    document.getElementById("detailSubtitle").textContent = subtitle || "";
    wrapper.innerHTML = `
    <div class="detail-grid">
      <div class="detail-card">
        <div class="skeleton h-24 block"></div><div style="height:10px"></div>
        <div class="skeleton h-12 block" style="width:60%"></div>
        <div style="height:12px"></div>
        <div class="skeleton h-12 block"></div>
        <div style="height:8px"></div>
        <div class="skeleton h-12 block"></div>
      </div>
      <div class="detail-card">
        <h3>Group Memberships</h3>
        <div style="display:flex;gap:8px;flex-wrap:wrap;">
          <div class="skeleton h-12" style="width:30%"></div>
          <div class="skeleton h-12" style="width:30%"></div>
          <div class="skeleton h-12" style="width:30%"></div>
        </div>
      </div>
      <div class="detail-card">
        <h3>Assigned Licenses</h3>
        <div style="display:flex;gap:8px;flex-wrap:wrap;">
          <div class="skeleton h-12" style="width:40%"></div>
          <div class="skeleton h-12" style="width:40%"></div>
        </div>
      </div>
    </div>
  `;
}

async function openUserDetails(encodedId) {

    console.log("Attempting to open detail drawer...");
    const drawer = document.getElementById("detailDrawer");

    if (!drawer) {
        console.error("Drawer element not found!");
        return;
    }

    // Log current class list
    console.log("Drawer classes after update:", drawer.className);

    // Check computed style
    const style = window.getComputedStyle(drawer);
    console.log("Drawer transform:", style.transform);

    const id = decodeURIComponent(encodedId || '');



    console.log("Opening detail drawer, id:", id);
    console.log("Drawer before classList:", drawer.className);
    drawer.classList.add("open");

    console.log("Drawer after classList:", drawer.className);
    console.log("Computed transform:", window.getComputedStyle(drawer).transform);






    console.log("Adding open class to drawer...");

    if (!id) return;
    // show skeleton immediately and open drawer
    document.getElementById("detailDrawer")?.classList.add("open");
    openDetailSkeleton("Loading...", "");

    const token = await getToken();
    if (!token) return;

    // Try to find user in allUsers
    const user = allUsers.find(u => u.id === id) || null;
    console.log("USER OBJECT:", user);

    // Fetch groups, licenses, roles
    try {
        const [groups, licenses] = await Promise.all([
            loadUserGroups(id, token),
            loadUserLicenses(id, token)
        ]);
        const rolesMap = await loadDirectoryRoles(token);

        // Categorize groups
        const categorized = { security: [], m365: [], distribution: [], other: [] };
        (groups || []).forEach(g => {
            const t = g['@odata.type'] || '';
            // groupTypes is present for M365 groups (groupTypes includes 'Unified')
            const isUnified = g.groupTypes && Array.isArray(g.groupTypes) && g.groupTypes.includes('Unified');
            if (t.includes('group')) {
                if (g.securityEnabled) categorized.security.push(g);
                else if (isUnified || g.mailEnabled) categorized.m365.push(g);
                else if (!g.securityEnabled && !g.mailEnabled) categorized.distribution.push(g);
                else categorized.other.push(g);
            } else if (t.includes('directoryRole')) {
                // directory role — treat as role
                categorized.other.push({ ...g, isRole: true });
            } else categorized.other.push(g);
        });

        // Build license friendly names by mapping skuId -> skuPartNumber via subscribedSkus
        const subSkus = await loadSubscribedSkus(token);
        const licenseDisplay = (licenses || []).map(l => {
            const skuId = l.skuId || l.skuId;
            const friendly = (subSkus && subSkus[skuId]) ? subSkus[skuId] : (l.skuPartNumber || skuId);
            return { sku: friendly, skuId };
        });

        // Directory roles can also be retrieved from memberOf or mapped via cachedDirectoryRoles
        const dirRoles = (groups || []).filter(g => (g['@odata.type'] || '').includes('directoryRole')).map(r => rolesMap[r.id] || r.displayName || r.id);

        // render details
        renderUserDetailsPanel(user, categorized, licenseDisplay, dirRoles);
    } catch (err) {
        console.error("Failed to load expanded details", err);
        // fallback: show basic user data
        renderUserDetailsPanel(user, {}, [], []);
    }
}

function renderUserDetailsPanel(user, categorized, licenseDisplay, dirRoles) {
    console.log("USER:", user);
    const wrapper = document.getElementById("detailGridWrapper");
    if (!wrapper) return;
    document.getElementById("detailTitle").textContent = user?.displayName || "User Details";
    document.getElementById("detailSubtitle").textContent = user?.userPrincipalName || user?.mail || "";

    const basicHtml = `
    <div class="detail-card">
      <h3>Basic Information</h3>
      ${ALL_FIELDS.map(f => `<div class="detail-row"><div class="k">${escapeHtml(FRIENDLY[f] || f)}</div><div class="v">${escapeHtml((user?.[f] || '') + '')}</div></div>`).join('')}
    </div>
  `;

    const groupsHtml = `
    <div class="detail-card">
      <h3>Group Memberships</h3>
      ${['security', 'm365', 'distribution', 'other'].map(k => {
        const list = (categorized && categorized[k]) || [];
        if (!list.length) return '';
        const title = k === 'm365' ? 'Microsoft 365 Groups' : (k === 'security' ? 'Security Groups' : (k === 'distribution' ? 'Distribution' : 'Other'));
        return `<div style="margin-bottom:8px;"><strong>${escapeHtml(title)}</strong><div class="chips">${list.map(g => `<div class="chip">${escapeHtml(g.displayName || g.id || 'Unnamed')}</div>`).join('')}</div></div>`;
    }).join('')}
      ${((categorized && Object.values(categorized).flat().length) ? '' : '<div class="muted small">No groups</div>')}
    </div>
  `;

    const licensesHtml = `
    <div class="detail-card">
      <h3>Assigned Licenses</h3>
      <div class="chips">${(licenseDisplay || []).map(l => `<div class="chip">${escapeHtml(l.sku)}${l.skuId ? ` <span class="small muted">(${escapeHtml(l.skuId)})</span>` : ''}</div>`).join('') || '<div class="muted small">No licenses</div>'}</div>
    </div>
  `;

    const rolesHtml = `
    <div class="detail-card">
      <h3>Directory Roles</h3>
      <div class="chips">${(dirRoles || []).map(r => `<div class="chip">${escapeHtml(r)}</div>`).join('') || '<div class="muted small">No directory roles</div>'}</div>
    </div>
  `;

    wrapper.innerHTML = `<div class="detail-grid">${basicHtml}${groupsHtml}${licensesHtml}${rolesHtml}</div>`;
}

/* ========================
   Drawer open/close helpers
   ======================== */
function closeDetailDrawer() { document.getElementById("detailDrawer")?.classList.remove("open"); }
function openFilterDrawer() { document.getElementById("filterDrawer")?.classList.add("open"); }
function closeFilterDrawer() { document.getElementById("filterDrawer")?.classList.remove("open"); }
function openColumnDrawer() { document.getElementById("columnDrawer")?.classList.add("open"); populateColumnDrawer(); }
function closeColumnDrawer() { document.getElementById("columnDrawer")?.classList.remove("open"); }

/* ========================
   Column + event bindings + init
   ======================== */
document.getElementById("searchInput")?.addEventListener("input", applySearchAndFilters);
document.getElementById("openFilter").onclick = openFilterDrawer;
document.getElementById("openColumns").onclick = openColumnDrawer;
document.getElementById("resetAll").onclick = () => {
    const si = document.getElementById("searchInput"); if (si) si.value = '';
    activeFilters = []; redrawActiveFilters(); renderTable(allUsers);
};

/* DOM ready init */
document.addEventListener("DOMContentLoaded", async () => {
    populateStaticIcons();
    populateFilterFields();
    // Initialize auth then load all users
    const ok = await initializeAuth();
    if (ok) {
        await loadUsers();
    }
});