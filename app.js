// --- Database Service (Antigravity Cloud Backend Integration) ---
window.dbStore = {
    users: [], projects: [], stages: [], notifications: [], delayRecords: [], lessonsLearned: [], projectReports: []
};

// --- Antigravity Cloud Backend URL ---
// Change this to your cloud deployment URL when deploying remotely
const API_BASE = "https://pme-nexus.onrender.com/api/data";

// Track pending sync operations for retry
let pendingSyncs = [];
let isSyncing = false;

const DB = {
    get: (key) => window.dbStore[key] || [],
    set: (key, data) => {
        window.dbStore[key] = data;
        syncToBackend(key, data);
    },
    getCurrentUser: () => JSON.parse(localStorage.getItem('currentUser')),
    setCurrentUser: (user) => {
        if (user) localStorage.setItem('currentUser', JSON.stringify(user));
        else localStorage.removeItem('currentUser');
    }
};

// --- Backend Sync with Retry Queue ---
function syncToBackend(key, data) {
    fetch(`${API_BASE}/api/sync`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ key, data })
    })
        .then(res => {
            if (!res.ok) throw new Error(`Sync failed: ${res.status}`);
            // If sync succeeds, flush any pending syncs
            if (pendingSyncs.length > 0) flushPendingSyncs();
        })
        .catch(err => {
            console.error("Antigravity Cloud Sync Failed:", err);
            // Queue for retry — only keep the latest data for each key
            const existing = pendingSyncs.findIndex(s => s.key === key);
            if (existing !== -1) {
                pendingSyncs[existing].data = data;
            } else {
                pendingSyncs.push({ key, data });
            }
        });
}

function flushPendingSyncs() {
    if (isSyncing || pendingSyncs.length === 0) return;
    isSyncing = true;

    const batch = [...pendingSyncs];
    pendingSyncs = [];

    Promise.all(batch.map(({ key, data }) =>
        fetch(`${API_BASE}/api/sync`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ key, data })
        }).catch(() => {
            // Re-queue failed items
            pendingSyncs.push({ key, data });
        })
    )).finally(() => {
        isSyncing = false;
    });
}

// --- Centralized Data Fetch from Backend ---
function fetchAllData(retries = 3) {
    return fetch(`${API_BASE}/api/data`)
        .then(res => {
            if (!res.ok) throw new Error(`Backend responded with ${res.status}`);
            return res.json();
        })
        .then(data => {
            if (data && typeof data === 'object' && Object.keys(data).length > 0) {
                // Merge backend data into dbStore, preserving structure
                const collections = ['users', 'projects', 'stages', 'notifications', 'delayRecords', 'lessonsLearned', 'projectReports'];
                collections.forEach(key => {
                    if (Array.isArray(data[key])) {
                        window.dbStore[key] = data[key];
                    }
                });
            }
            // Flush any pending syncs now that backend is reachable
            if (pendingSyncs.length > 0) flushPendingSyncs();
            return { success: true };
        })
        .catch(err => {
            if (retries > 1) {
                // Exponential backoff: 500ms, 1000ms, 2000ms
                const delay = 500 * Math.pow(2, 3 - retries);
                return new Promise(resolve =>
                    setTimeout(() => resolve(fetchAllData(retries - 1)), delay)
                );
            }
            console.error("Antigravity Cloud unreachable after retries:", err);
            return { success: false, error: err };
        });
}

// Lightweight data refresh (no loading screen, just re-fetches and re-renders active view)
function refreshDataAndView() {
    return fetchAllData(1).then(result => {
        refreshActiveView();
        return result;
    });
}

// Re-renders whichever view is currently active
function refreshActiveView() {
    if (views.dashboard?.classList.contains('active')) {
        renderDashboard();
    } else if (views.projects?.classList.contains('active')) {
        renderProjectsTab();
    } else if (views.users?.classList.contains('active')) {
        renderUsersTable(document.getElementById('search-users-input')?.value || '');
    }
    // projectDetail doesn't need refresh — it's rendered on-demand
}

// --- Project Completion Evaluation (Client-Side) ---
// Mirrors server-side logic: auto-sets completionDate when all stages reach 100%
function evaluateProjectCompletion() {
    const projects = DB.get('projects');
    const stages = DB.get('stages');
    let updated = false;

    projects.forEach(project => {
        const projectStages = stages.filter(s => s.projectId === project.id);
        const isCompleted = projectStages.length > 0 && projectStages.every(s => parseInt(s.progress) === 100);

        if (isCompleted && !project.completionDate) {
            project.completionDate = new Date().toISOString();
            updated = true;
        } else if (!isCompleted && project.completionDate) {
            project.completionDate = null;
            updated = true;
        }
    });

    if (updated) {
        DB.set('projects', projects);
    }
}

// --- Global State ---
let currentUser = DB.getCurrentUser();
let currentProjectId = null;
let currentStageId = null;
let actionPending = null;

// --- DOM Elements ---
const views = {
    login: document.getElementById('login-view'),
    signup: document.getElementById('signup-view'),
    dashboard: document.getElementById('dashboard-view'),
    projects: document.getElementById('projects-view'),
    projectDetail: document.getElementById('project-detail-view'),
    users: document.getElementById('users-view')
};

const navbar = document.getElementById('navbar');
const toastContainer = document.getElementById('toast-container');

// --- Initialization ---
function enforceAdminConstraints() {
    let users = DB.get('users');
    let dbUpdated = false;
    users.forEach(u => {
        if (u.email === 'fredadeefe224@gmail.com') {
            if (u.role !== 'Admin') {
                u.role = 'Admin';
                dbUpdated = true;
            }
        } else {
            if (u.role === 'Admin') {
                u.role = 'Project Manager';
                dbUpdated = true;
            }
        }
    });

    if (dbUpdated) {
        DB.set('users', users);
        if (currentUser) {
            const updatedProfile = users.find(u => u.id === currentUser.id);
            if (updatedProfile) {
                currentUser = updatedProfile;
                DB.setCurrentUser(currentUser);
            }
        }
    }
}

let domReady = false;

// Apply theme early from localStorage to prevent light/dark flash on the loader
(function earlyThemeSync() {
    try {
        const saved = JSON.parse(localStorage.getItem('currentUser'));
        if (saved && saved.theme_preference === 'light') {
            document.body.className = 'light-theme';
        }
    } catch (e) { /* ignore */ }
})();

function startupApp() {
    // --- Global State Runtime Updates ---
    currentUser = DB.getCurrentUser();

    lucide.createIcons();
    setupEventListeners();
    enforceAdminConstraints();

    // Auth Check — decide which view to show BEFORE dismissing loader
    if (currentUser) {
        applyTheme(currentUser.theme_preference || 'dark');
        showView('dashboard');
        updateUserNav();
        renderDashboard();
        renderNotifications(); // Global badge update on startup
    } else {
        applyTheme('dark');
        showView('login');
    }

    // Dismiss loader with fade-out transition
    const loader = document.getElementById('app-loader');
    if (loader) {
        loader.classList.add('loaded');
        // Remove from DOM after transition completes
        loader.addEventListener('transitionend', () => loader.remove(), { once: true });
    }
}

document.addEventListener('DOMContentLoaded', () => {
    domReady = true;
    initializeApp();
});

// --- App Initialization: Fetch data from backend THEN start ---
function initializeApp() {
    if (!domReady) return;

    fetchAllData(3).then(result => {
        // Ensure all collections exist locally even if backend had partial data
        const collections = ['users', 'projects', 'stages', 'notifications', 'delayRecords', 'lessonsLearned', 'projectReports'];
        collections.forEach(key => {
            if (!Array.isArray(window.dbStore[key])) {
                window.dbStore[key] = [];
            }
        });

        startupApp();

        if (!result.success) {
            // Backend unreachable — warn user
            setTimeout(() => {
                showToast('Backend server unreachable. Data may not be up to date.', 'error');
            }, 500);
        }
    });
}

// --- UI Utilities ---
function showView(viewId) {
    if (viewId === 'users' && (!currentUser || currentUser.email !== 'fredadeefe224@gmail.com')) {
        showToast('Unauthorized access', 'error');
        viewId = 'dashboard';
    }

    Object.values(views).forEach(v => {
        if (v) v.classList.remove('active');
    });

    // Close notifications panel when switching views
    const notifPanel = document.getElementById('notifications-panel');
    if (notifPanel) notifPanel.style.display = 'none';

    if (views[viewId]) {
        views[viewId].classList.add('active');
    }

    // Tab Highlight State Management - Strictly Route Based
    const navDashboardBtn = document.getElementById('nav-dashboard-btn');
    const navProjectsBtn = document.getElementById('nav-projects-btn');
    const navUsersBtn = document.getElementById('nav-users-btn');

    if (navDashboardBtn) navDashboardBtn.classList.remove('active');
    if (navProjectsBtn) navProjectsBtn.classList.remove('active');
    if (navUsersBtn) navUsersBtn.classList.remove('active');

    if (viewId === 'dashboard' && navDashboardBtn) {
        navDashboardBtn.classList.add('active');
    } else if (viewId === 'projects' && navProjectsBtn) {
        navProjectsBtn.classList.add('active');
    } else if (viewId === 'users' && navUsersBtn) {
        navUsersBtn.classList.add('active');
    }

    if (viewId === 'login' || viewId === 'signup') {
        navbar.classList.add('hidden');
    } else {
        navbar.classList.remove('hidden');
    }
    lucide.createIcons();
}

function applyTheme(theme) {
    const isLight = theme === 'light';
    document.body.className = isLight ? 'light-theme' : 'dark-theme';

    const iconLight = document.getElementById('theme-icon-light');
    const iconDark = document.getElementById('theme-icon-dark');

    if (iconLight && iconDark) {
        iconLight.style.display = isLight ? 'none' : 'inline-block';
        iconDark.style.display = isLight ? 'inline-block' : 'none';

        // Ensure Lucide icons re-initialize if needed but display toggle is sufficient
    }
}

function handleThemeToggle() {
    if (!currentUser) return;
    const newTheme = document.body.classList.contains('light-theme') ? 'dark' : 'light';

    applyTheme(newTheme);

    // Save to user
    currentUser.theme_preference = newTheme;
    DB.setCurrentUser(currentUser);

    const users = DB.get('users');
    const uIndex = users.findIndex(u => u.id === currentUser.id);
    if (uIndex !== -1) {
        users[uIndex].theme_preference = newTheme;
        DB.set('users', users);
    }
}

function showToast(message, type = 'info') {
    const toast = document.createElement('div');
    toast.className = `toast ${type}`;
    toast.textContent = message;
    toastContainer.appendChild(toast);
    setTimeout(() => {
        if (toast.parentNode) toast.remove();
    }, 3000);
}

function toggleModal(modalId) {
    const modal = document.getElementById(modalId);
    if (modal) {
        modal.classList.toggle('hidden');
    }
}

// --- Event Listeners Setup ---
function setupEventListeners() {
    // Auth navigation
    document.getElementById('go-to-signup')?.addEventListener('click', () => showView('signup'));
    document.getElementById('go-to-login')?.addEventListener('click', () => showView('login'));
    document.getElementById('logout-btn')?.addEventListener('click', logout);

    // Theme Toggle
    document.getElementById('theme-toggle-btn')?.addEventListener('click', handleThemeToggle);

    // Auth Forms
    document.getElementById('login-form')?.addEventListener('submit', handleLogin);
    document.getElementById('signup-form')?.addEventListener('submit', handleSignup);

    // Password visibility toggle logic
    setupPasswordToggles();

    // Modals
    document.querySelectorAll('.close-modal').forEach(btn => {
        btn.addEventListener('click', (e) => {
            if (e.target.closest('#modal-confirm-action') || e.target.id === 'cancel-action-btn') {
                actionPending = null;
                // re-render table to rollback select DOM UI optionally
                if (document.getElementById('users-view')?.classList.contains('active')) {
                    renderUsersTable(document.getElementById('search-users-input')?.value || '');
                }
            }
            const modal = e.target.closest('.modal-backdrop');
            if (modal) modal.classList.add('hidden');
        });
    });

    document.getElementById('confirm-action-btn')?.addEventListener('click', () => {
        if (actionPending) {
            actionPending();
            actionPending = null;
        }
        toggleModal('modal-confirm-action');
        lucide.createIcons();
    });

    // Projects Tab
    document.getElementById('nav-projects-btn')?.addEventListener('click', () => {
        showView('projects');
        renderProjectsTab();
        // Re-fetch backend data in background and refresh
        fetchAllData(1).then(() => renderProjectsTab());
    });

    // Filter listeners for completed projects
    document.getElementById('filter-month')?.addEventListener('change', () => renderProjectsTab());
    document.getElementById('filter-year')?.addEventListener('change', () => renderProjectsTab());

    // Dashboard Actions
    document.getElementById('nav-users-btn')?.addEventListener('click', () => {
        if (!currentUser || currentUser.email !== 'fredadeefe224@gmail.com') return;
        showView('users');
        renderUsersTable();
        // Re-fetch backend data in background and refresh
        fetchAllData(1).then(() => renderUsersTable(document.getElementById('search-users-input')?.value || ''));
    });

    document.getElementById('search-users-input')?.addEventListener('input', (e) => {
        renderUsersTable(e.target.value);
    });

    document.getElementById('create-project-btn')?.addEventListener('click', () => {
        if (currentUser.role === 'Viewer') {
            showToast('Viewers cannot create projects.', 'error');
            return;
        }
        document.getElementById('form-create-project').reset();
        toggleModal('modal-create-project');
    });

    // Project Actions
    document.getElementById('nav-dashboard-btn')?.addEventListener('click', () => {
        showView('dashboard');
        renderDashboard();
        // Re-fetch backend data in background and refresh
        fetchAllData(1).then(() => renderDashboard());
    });

    document.getElementById('back-to-dashboard')?.addEventListener('click', () => {
        showView('dashboard');
        renderDashboard();
        // Re-fetch backend data in background and refresh
        fetchAllData(1).then(() => renderDashboard());
    });

    document.getElementById('form-create-project')?.addEventListener('submit', handleCreateProject);

    // Stage Actions
    document.getElementById('create-stage-btn')?.addEventListener('click', () => {
        if (currentUser.role === 'Viewer') {
            showToast('Viewers cannot modify stages', 'error');
            return;
        }
        document.getElementById('form-stage').reset();
        document.getElementById('stage-id').value = '';
        document.getElementById('modal-stage-title').textContent = 'Add Stage';
        toggleModal('modal-stage');
    });

    document.getElementById('form-stage')?.addEventListener('submit', handleSaveStage);

    document.getElementById('nav-notifications-btn')?.addEventListener('click', () => {
        const panel = document.getElementById('notifications-panel');
        if (panel) {
            panel.style.display = panel.style.display === 'none' ? 'flex' : 'none';
            if (panel.style.display === 'flex') {
                renderNotifications();
            }
        }
    });

    document.getElementById('close-notifications-btn')?.addEventListener('click', () => {
        const panel = document.getElementById('notifications-panel');
        if (panel) panel.style.display = 'none';
    });

    document.getElementById('form-delay')?.addEventListener('submit', handleSaveDelay);

    document.getElementById('create-lesson-btn')?.addEventListener('click', () => {
        if (currentUser.role === 'Viewer') {
            showToast('Viewers cannot create lessons', 'error');
            return;
        }
        document.getElementById('form-lesson').reset();

        // Populate stage options
        const pStages = DB.get('stages').filter(s => s.projectId === currentProjectId);
        const stageSelect = document.getElementById('lesson-stage');
        stageSelect.innerHTML = '<option value="">General Project Lesson</option>';
        pStages.forEach(s => {
            stageSelect.innerHTML += `<option value="${s.id}">${s.name}</option>`;
        });

        toggleModal('modal-lesson');
    });

    document.getElementById('form-lesson')?.addEventListener('submit', handleSaveLesson);

    document.getElementById('generate-report-btn')?.addEventListener('click', handleGenerateReport);
}

function setupPasswordToggles() {
    const toggleBtns = document.querySelectorAll('.password-toggle-btn');

    toggleBtns.forEach(btn => {
        btn.addEventListener('click', function () {
            const input = this.previousElementSibling;
            if (!input || input.tagName !== 'INPUT') return;

            const icon = this.querySelector('i');
            if (input.type === 'password') {
                input.type = 'text';
                input.dataset.toggled = 'true';
                if (icon) {
                    icon.setAttribute('data-lucide', 'eye-off');
                    lucide.createIcons();
                }
                this.setAttribute('aria-label', 'Hide password');
            } else {
                input.type = 'password';
                input.dataset.toggled = 'false';
                if (icon) {
                    icon.setAttribute('data-lucide', 'eye');
                    lucide.createIcons();
                }
                this.setAttribute('aria-label', 'Show password');
            }
        });
    });

    const passwordInputs = document.querySelectorAll('input[type="password"]');
    passwordInputs.forEach(input => {
        let maskTimeout;
        let originalValue = '';

        input.addEventListener('input', function (e) {
            if (this.dataset.toggled === 'true') {
                return; // Normal typing if eye is toggled on
            }

            clearTimeout(maskTimeout);

            const currentVal = this.value;
            // Only trigger behavior on new trailing character typing
            if (currentVal.length > originalValue.length) {
                // Temporarily show the text
                this.type = 'text';

                maskTimeout = setTimeout(() => {
                    // Only revert if we haven't manually toggled it ON in the meanwhile
                    if (this.dataset.toggled !== 'true') {
                        this.type = 'password';
                    }
                }, 1000);
            }
            originalValue = currentVal;
        });
    });
}


// --- Controllers ---

function handleSignup(e) {
    e.preventDefault();
    const username = document.getElementById('signup-username').value.trim();
    let role = document.getElementById('signup-role').value;
    const password = document.getElementById('signup-password').value;
    const email = document.getElementById('signup-email')?.value.trim() || '';

    if (email === 'fredadeefe224@gmail.com') {
        role = 'Admin';
    } else {
        role = 'Viewer';
    }

    const users = DB.get('users');
    if (users.find(u => u.username === username)) {
        showToast('Username already exists.', 'error');
        return;
    }

    const newUser = {
        id: Date.now().toString(),
        username,
        email,
        role,
        password,
        theme_preference: 'dark',
        createdAt: new Date().toISOString(),
        isDisabled: false,
        lastLogin: null
    };
    users.push(newUser);
    DB.set('users', users);

    // Auto login
    currentUser = newUser;
    DB.setCurrentUser(currentUser);

    document.getElementById('signup-form').reset();
    showToast('Account created successfully!', 'success');
    applyTheme(currentUser.theme_preference);
    updateUserNav();
    showView('dashboard');
    renderDashboard();
}

function handleLogin(e) {
    e.preventDefault();
    const username = document.getElementById('login-username').value.trim();
    const password = document.getElementById('login-password').value;

    const users = DB.get('users');
    const userIndex = users.findIndex(u => u.username === username && u.password === password);
    const user = userIndex !== -1 ? users[userIndex] : null;

    if (user) {
        if (user.isDisabled) {
            showToast('Account is disabled. Contact an Admin.', 'error');
            return;
        }

        user.lastLogin = new Date().toISOString();
        users[userIndex] = user;
        DB.set('users', users);

        currentUser = user;
        DB.setCurrentUser(user);
        enforceAdminConstraints(); // Ensure state overrides appropriately globally

        document.getElementById('login-form').reset();
        showToast('Welcome back!', 'success');
        applyTheme(currentUser.theme_preference || 'dark');
        updateUserNav();
        showView('dashboard');
        renderDashboard();
    } else {
        showToast('Invalid credentials.', 'error');
    }
}

function logout() {
    currentUser = null;
    DB.setCurrentUser(null);
    applyTheme('dark');
    showView('login');
}

function updateUserNav() {
    document.getElementById('current-user-name').textContent = currentUser.username;
    const roleBadge = document.getElementById('current-user-role');
    roleBadge.textContent = currentUser.role;

    // Clear old role classes
    roleBadge.classList.remove('role-admin', 'role-pm', 'role-viewer');

    if (currentUser.role === 'Admin') roleBadge.classList.add('role-admin');
    else if (currentUser.role === 'Project Manager') roleBadge.classList.add('role-pm');
    else roleBadge.classList.add('role-viewer');

    // Hide create buttons for Viewer
    const createBtn = document.getElementById('create-project-btn');
    const createStageBtn = document.getElementById('create-stage-btn');
    const createLessonBtn = document.getElementById('create-lesson-btn');
    if (createBtn) createBtn.style.display = currentUser.role === 'Viewer' ? 'none' : 'flex';
    if (createStageBtn) createStageBtn.style.display = currentUser.role === 'Viewer' ? 'none' : 'flex';
    if (createLessonBtn) createLessonBtn.style.display = currentUser.role === 'Viewer' ? 'none' : 'flex';

    const navUsersBtn = document.getElementById('nav-users-btn');
    if (navUsersBtn) navUsersBtn.style.display = (currentUser.email === 'fredadeefe224@gmail.com') ? 'flex' : 'none';
}

function handleCreateProject(e) {
    e.preventDefault();
    const name = document.getElementById('new-project-name').value.trim();
    const desc = document.getElementById('new-project-desc').value.trim();

    const projects = DB.get('projects');
    const newProject = {
        id: Date.now().toString(),
        name,
        description: desc,
        createdBy: currentUser.id,
        createdAt: new Date().toISOString()
    };

    projects.push(newProject);
    DB.set('projects', projects);

    // Auto-create 5 stages
    const stages = DB.get('stages');
    const stageNames = [
        "Project Identification",
        "Detailed Planning",
        "Procurement",
        "Implementation",
        "Closure"
    ];

    const todayStr = new Date().toISOString().split('T')[0];
    const nextWeek = new Date();
    nextWeek.setDate(nextWeek.getDate() + 7);
    const nextWeekStr = nextWeek.toISOString().split('T')[0];

    stageNames.forEach((stageName, index) => {
        stages.push({
            id: Date.now().toString() + '-' + index,
            projectId: newProject.id,
            name: stageName,
            plannedStart: todayStr,
            plannedEnd: nextWeekStr,
            actualStart: '',
            actualEnd: '',
            progress: 0,
            status: 'On Track'
        });
    });
    DB.set('stages', stages);

    toggleModal('modal-create-project');
    showToast('Project created successfully', 'success');
    renderDashboard();
}

function updateAllStageStatuses() {
    const allStages = DB.get('stages');
    let dbUpdated = false;
    const todayStr = new Date().toISOString().split('T')[0];

    allStages.forEach(stage => {
        let newStatus = 'On Track';
        if (parseInt(stage.progress) === 100) {
            newStatus = 'Completed';
        } else if (todayStr > stage.plannedEnd) {
            newStatus = 'Behind Schedule';
        }
        if (stage.status !== newStatus) {
            stage.status = newStatus;
            dbUpdated = true;
            if (newStatus === 'Behind Schedule') {
                createBehindScheduleNotification(stage);
            }
        }
    });

    if (dbUpdated) {
        DB.set('stages', allStages);
        renderNotifications();
    }

    // Evaluate project completion status after stage status updates
    evaluateProjectCompletion();
}

function renderDashboard() {
    updateAllStageStatuses();
    renderNotifications(); // Update global badge count
    const projects = DB.get('projects');
    const container = document.getElementById('projects-container');
    container.innerHTML = '';

    if (projects.length === 0) {
        container.innerHTML = '<div style="grid-column: 1/-1; text-align: center; color: var(--text-muted); padding: 3rem;">No projects found. Create one to get started.</div>';
        return;
    }

    // Sort newest first
    projects.sort((a, b) => new Date(b.createdAt) - new Date(a.createdAt));

    // Calculate progress for each project based on stages
    const allStages = DB.get('stages');

    projects.forEach(project => {
        const pStages = allStages.filter(s => s.projectId === project.id);
        const avgProgress = pStages.length ?
            Math.round(pStages.reduce((sum, s) => sum + parseInt(s.progress), 0) / pStages.length) : 0;

        const card = document.createElement('div');
        card.className = 'project-card glass-panel';
        card.innerHTML = `
            <div class="project-header">
                <h3>${project.name}</h3>
                <p class="project-desc">${project.description || 'No description provided.'}</p>
            </div>
            <div class="project-meta">
                <span class="badge ${avgProgress === 100 ? 'status-completed' : 'status-ontrack'}">
                    ${avgProgress}% Complete
                </span>
                <span style="font-size: 0.8rem; color: var(--text-muted);">
                    ${pStages.length} Stages
                </span>
            </div>
        `;

        card.addEventListener('click', () => openProjectDetail(project.id));
        container.appendChild(card);
    });
}

// --- Projects Tab ---
function renderProjectsTab() {
    evaluateProjectCompletion();
    updateAllStageStatuses();

    const filterMonth = document.getElementById('filter-month')?.value || '';
    const filterYear = document.getElementById('filter-year')?.value || '';

    // Populate year filter dynamically from existing project data
    populateYearFilter();

    // Try to fetch from backend, fall back to client-side data
    const completedPromise = fetchCompletedProjects(filterMonth, filterYear);
    const inProgressPromise = fetchInProgressProjects();

    Promise.all([completedPromise, inProgressPromise]).then(([completed, inProgress]) => {
        renderCompletedProjectsTable(completed);
        renderInProgressProjectsTable(inProgress);
        lucide.createIcons();
    });
}

function populateYearFilter() {
    const yearSelect = document.getElementById('filter-year');
    if (!yearSelect) return;

    const currentVal = yearSelect.value;
    const projects = DB.get('projects');
    const yearsSet = new Set();
    const currentYear = new Date().getFullYear();

    // Add current year and surrounding years
    yearsSet.add(currentYear);
    yearsSet.add(currentYear - 1);
    yearsSet.add(currentYear + 1);

    // Add years from project completionDates
    projects.forEach(p => {
        if (p.completionDate) {
            try {
                const yr = new Date(p.completionDate).getFullYear();
                if (!isNaN(yr)) yearsSet.add(yr);
            } catch (e) { /* skip */ }
        }
        if (p.createdAt) {
            try {
                const yr = new Date(p.createdAt).getFullYear();
                if (!isNaN(yr)) yearsSet.add(yr);
            } catch (e) { /* skip */ }
        }
    });

    const years = Array.from(yearsSet).sort((a, b) => b - a);

    // Only rebuild if options changed
    const existingYears = Array.from(yearSelect.options).slice(1).map(o => o.value);
    const newYears = years.map(y => String(y));
    if (JSON.stringify(existingYears) !== JSON.stringify(newYears)) {
        yearSelect.innerHTML = '<option value="">All Years</option>';
        years.forEach(y => {
            yearSelect.innerHTML += `<option value="${y}">${y}</option>`;
        });
        yearSelect.value = currentVal;
    }
}

function fetchCompletedProjects(month, year) {
    // Build query params
    let url = `${API_BASE}/api/projects/completed`;
    const params = [];
    if (month) params.push(`month=${month}`);
    if (year) params.push(`year=${year}`);
    if (params.length) url += '?' + params.join('&');

    return fetch(url)
        .then(res => res.json())
        .then(data => data.projects || [])
        .catch(() => {
            // Fallback: compute from client-side data
            return getCompletedProjectsLocal(month, year);
        });
}

function fetchInProgressProjects() {
    return fetch(`${API_BASE}/api/projects/in-progress`)
        .then(res => res.json())
        .then(data => data.projects || [])
        .catch(() => {
            // Fallback: compute from client-side data
            return getInProgressProjectsLocal();
        });
}

function getCompletedProjectsLocal(month, year) {
    const projects = DB.get('projects');
    const stages = DB.get('stages');

    let completed = projects.filter(p => {
        const pStages = stages.filter(s => s.projectId === p.id);
        return pStages.length > 0 && pStages.every(s => parseInt(s.progress) === 100) && p.completionDate;
    });

    if (month || year) {
        completed = completed.filter(p => {
            try {
                const cd = new Date(p.completionDate);
                const matchMonth = month ? cd.getMonth() + 1 === parseInt(month) : true;
                const matchYear = year ? cd.getFullYear() === parseInt(year) : true;
                return matchMonth && matchYear;
            } catch (e) {
                return false;
            }
        });
    }

    return completed.map(p => {
        const pStages = stages.filter(s => s.projectId === p.id);
        const avgProgress = pStages.length
            ? Math.round(pStages.reduce((sum, s) => sum + parseInt(s.progress), 0) / pStages.length)
            : 0;
        return { ...p, status: 'Completed', totalProgress: avgProgress, stageCount: pStages.length };
    });
}

function getInProgressProjectsLocal() {
    const projects = DB.get('projects');
    const stages = DB.get('stages');

    const inProgress = projects.filter(p => {
        const pStages = stages.filter(s => s.projectId === p.id);
        return !(pStages.length > 0 && pStages.every(s => parseInt(s.progress) === 100));
    });

    return inProgress.map(p => {
        const pStages = stages.filter(s => s.projectId === p.id);
        const avgProgress = pStages.length
            ? Math.round(pStages.reduce((sum, s) => sum + parseInt(s.progress), 0) / pStages.length)
            : 0;
        return { ...p, status: 'In Progress', totalProgress: avgProgress, stageCount: pStages.length };
    });
}

function renderCompletedProjectsTable(projects) {
    const tbody = document.getElementById('completed-projects-body');
    if (!tbody) return;
    tbody.innerHTML = '';

    if (projects.length === 0) {
        tbody.innerHTML = `
            <tr>
                <td colspan="3">
                    <div class="projects-empty-state">
                        <i data-lucide="check-circle-2"></i>
                        <p>No completed projects found</p>
                        <p class="empty-sub">Complete all stages of a project to see it here</p>
                    </div>
                </td>
            </tr>
        `;
        return;
    }

    // Sort by completionDate descending
    projects.sort((a, b) => new Date(b.completionDate) - new Date(a.completionDate));

    const reports = DB.get('projectReports');
    const isViewer = currentUser && currentUser.role === 'Viewer';

    projects.forEach(p => {
        const completionDate = new Date(p.completionDate);
        const dateFormatted = completionDate.toLocaleDateString('en-US', {
            year: 'numeric', month: 'short', day: 'numeric'
        });

        // Check if a report already exists for this project
        const report = reports.find(r => r.projectId === p.id);
        let reportBtns = '';

        if (isViewer) {
            // Viewers cannot generate or download reports
            reportBtns = `<span style="font-size: 0.8rem; color: var(--text-muted);">No access</span>`;
        } else if (report && report.content) {
            // Report exists — show Download + Regenerate buttons
            reportBtns = `
                <div style="display: flex; gap: 0.5rem; justify-content: flex-end; align-items: center;">
                    <button class="btn-download-report" onclick="downloadProjectReport('${p.id}')">
                        <i data-lucide="download"></i> Download
                    </button>
                    <button class="btn-download-report btn-generate" onclick="generateAndDownloadReport('${p.id}')" title="Regenerate report with latest data">
                        <i data-lucide="refresh-cw"></i>
                    </button>
                </div>`;
        } else {
            // No report yet — show Generate & Download button
            reportBtns = `<button class="btn-download-report btn-generate" onclick="generateAndDownloadReport('${p.id}')">
                <i data-lucide="file-text"></i> Generate &amp; Download
            </button>`;
        }

        const tr = document.createElement('tr');
        tr.innerHTML = `
            <td class="project-name-cell">
                ${p.name}
                ${p.description ? `<div class="project-desc-sub">${p.description}</div>` : ''}
            </td>
            <td class="date-cell">
                ${dateFormatted}
            </td>
            <td style="text-align: right;">
                ${reportBtns}
            </td>
        `;
        tbody.appendChild(tr);
    });
}

function renderInProgressProjectsTable(projects) {
    const tbody = document.getElementById('in-progress-projects-body');
    if (!tbody) return;
    tbody.innerHTML = '';

    if (projects.length === 0) {
        tbody.innerHTML = `
            <tr>
                <td colspan="4">
                    <div class="projects-empty-state">
                        <i data-lucide="loader"></i>
                        <p>No projects in progress</p>
                        <p class="empty-sub">All projects have been completed — great work!</p>
                    </div>
                </td>
            </tr>
        `;
        return;
    }

    const allStages = DB.get('stages');

    // Sort by created date descending
    projects.sort((a, b) => new Date(b.createdAt) - new Date(a.createdAt));

    projects.forEach(p => {
        const pStages = allStages.filter(s => s.projectId === p.id);
        const avgProgress = p.totalProgress !== undefined ? p.totalProgress :
            (pStages.length ? Math.round(pStages.reduce((sum, s) => sum + parseInt(s.progress), 0) / pStages.length) : 0);

        // Find current stage (first non-completed stage, or last stage)
        pStages.sort((a, b) => new Date(a.plannedStart) - new Date(b.plannedStart));
        const currentStage = pStages.find(s => parseInt(s.progress) < 100) || pStages[pStages.length - 1];

        let stageHtml = '<span style="color: var(--text-muted); font-size: 0.85rem;">No stages</span>';
        if (currentStage) {
            let chipClass = '';
            let chipIcon = 'target';
            if (currentStage.status === 'Behind Schedule') {
                chipClass = 'stage-behind';
                chipIcon = 'alert-triangle';
            } else if (currentStage.status === 'Completed') {
                chipClass = 'stage-completed';
                chipIcon = 'check';
            }
            stageHtml = `<span class="stage-chip ${chipClass}"><i data-lucide="${chipIcon}"></i> ${currentStage.name}</span>`;
        }

        // Calculate deadline (latest plannedEnd among all stages)
        let deadlineHtml = '<span style="color: var(--text-muted);">—</span>';
        if (pStages.length > 0) {
            const latestEnd = pStages.reduce((latest, s) => {
                const d = new Date(s.plannedEnd);
                return d > latest ? d : latest;
            }, new Date(0));

            const today = new Date();
            today.setHours(0, 0, 0, 0);
            const isOverdue = latestEnd < today && avgProgress < 100;

            const dateStr = latestEnd.toLocaleDateString('en-US', {
                year: 'numeric', month: 'short', day: 'numeric'
            });

            deadlineHtml = `<span class="deadline-cell ${isOverdue ? 'deadline-overdue' : ''}">${dateStr}${isOverdue ? ' (Overdue)' : ''}</span>`;
        }

        const progressClass = avgProgress === 100 ? 'progress-100' : '';

        const tr = document.createElement('tr');
        tr.innerHTML = `
            <td class="project-name-cell">
                ${p.name}
                ${p.description ? `<div class="project-desc-sub">${p.description}</div>` : ''}
            </td>
            <td>${stageHtml}</td>
            <td>
                <div class="table-progress-wrapper">
                    <div class="table-progress-bar">
                        <div class="table-progress-fill ${progressClass}" style="width: ${avgProgress}%"></div>
                    </div>
                    <span class="table-progress-text">${avgProgress}%</span>
                </div>
            </td>
            <td>${deadlineHtml}</td>
        `;
        tbody.appendChild(tr);
    });
}

// --- Projects Tab: Report Generation & Download ---

// Download an already-generated report for a project
function downloadProjectReport(projectId) {
    if (!currentUser) {
        showToast('Please log in to download reports.', 'error');
        return;
    }
    if (currentUser.role === 'Viewer') {
        showToast('Viewers cannot download reports.', 'error');
        return;
    }

    const reports = DB.get('projectReports') || [];
    const report = reports.find(r => r.projectId === projectId);
    if (!report || !report.content) {
        showToast('No report found. Please generate one first.', 'error');
        return;
    }

    const project = DB.get('projects').find(p => p.id === projectId);
    const fileName = `Project_Report_${(project?.name || 'Report').replace(/\s+/g, '_')}.doc`;

    triggerWordDocDownload(report.content, fileName);
    showToast('Report downloaded successfully.', 'success');
}

// Generate a fresh report for a project and immediately download it
function generateAndDownloadReport(projectId) {
    if (!currentUser) {
        showToast('Please log in to generate reports.', 'error');
        return;
    }
    if (currentUser.role === 'Viewer') {
        showToast('Viewers cannot generate reports.', 'error');
        return;
    }

    const project = DB.get('projects').find(p => p.id === projectId);
    if (!project) {
        showToast('Project not found.', 'error');
        return;
    }

    const stages = DB.get('stages').filter(s => s.projectId === projectId);
    const delays = DB.get('delayRecords').filter(d => d.projectId === projectId);
    const lessons = DB.get('lessonsLearned').filter(l => l.projectId === projectId);

    let totalProgress = stages.length ? Math.round(stages.reduce((sum, s) => sum + parseInt(s.progress), 0) / stages.length) : 0;

    let projectStatus = 'On Track';
    const todayStr = new Date().toISOString().split('T')[0];

    // Determine overall project status
    if (totalProgress === 100) {
        projectStatus = 'Completed';
    } else {
        stages.forEach(s => {
            if (parseInt(s.progress) < 100 && todayStr > s.plannedEnd) {
                projectStatus = 'Behind Schedule';
            }
        });
    }

    let numDelays = delays.length;
    let executiveSummary = `Project "${project.name}" is currently ${projectStatus}. Total average progress is ${totalProgress}%. Delays logged to date: ${numDelays}. Lessons recorded to date: ${lessons.length}.`;

    stages.sort((a, b) => new Date(a.plannedStart) - new Date(b.plannedStart));

    // Build Word-compatible HTML — same template as handleGenerateReport
    let wordHtml = `
    <html xmlns:o='urn:schemas-microsoft-com:office:office' xmlns:w='urn:schemas-microsoft-com:office:word' xmlns='http://www.w3.org/TR/REC-html40'>
    <head>
        <meta charset='utf-8'>
        <title>Project Report - ${project.name}</title>
        <style>
            body { font-family: 'Arial', sans-serif; line-height: 1.6; color: #333; }
            h1 { color: #2c3e50; text-align: center; border-bottom: 2px solid #3498db; padding-bottom: 10px; }
            h2 { color: #34495e; margin-top: 20px; }
            table { width: 100%; border-collapse: collapse; margin-top: 15px; }
            th, td { border: 1px solid #bdc3c7; padding: 8px; text-align: left; }
            th { background-color: #ecf0f1; font-weight: bold; }
            .behind-schedule { color: #e74c3c; font-weight: bold; }
            .on-track { color: #27ae60; }
            .completed { color: #2980b9; }
            .meta-info { margin-bottom: 20px; font-size: 14px; color: #7f8c8d; text-align: right; }
        </style>
    </head>
    <body>
        <h1>PROJECT REPORT: ${project.name}</h1>
        <div class="meta-info">Generated At: ${new Date().toLocaleString()}</div>

        <h2>1. EXECUTIVE SUMMARY</h2>
        <p>${executiveSummary}</p>

        <h2>2. PROJECT OVERVIEW</h2>
        <ul>
            <li><strong>Name:</strong> ${project.name}</li>
            <li><strong>Project ID:</strong> ${project.id}</li>
            <li><strong>Description:</strong> ${project.description || 'N/A'}</li>
            <li><strong>Creation Date:</strong> ${new Date(project.createdAt).toLocaleDateString()}</li>
            <li><strong>Total Stages:</strong> ${stages.length}</li>
            <li><strong>Overall Progress:</strong> ${totalProgress}%</li>
            <li><strong>Current Status:</strong> ${projectStatus}</li>
            ${project.completionDate ? `<li><strong>Completion Date:</strong> ${new Date(project.completionDate).toLocaleDateString()}</li>` : ''}
        </ul>

        <h2>3. STAGE-BY-STAGE PERFORMANCE</h2>
        <table>
            <tr>
                <th>Stage Name</th>
                <th>Planned Start / End</th>
                <th>Actual Start / End</th>
                <th>Progress</th>
                <th>Status</th>
            </tr>`;

    stages.forEach(s => {
        let statusClass = '';
        if (s.status === 'Behind Schedule') statusClass = 'behind-schedule';
        else if (s.status === 'Completed') statusClass = 'completed';
        else statusClass = 'on-track';

        wordHtml += `
            <tr>
                <td>${s.name}</td>
                <td>${s.plannedStart} to ${s.plannedEnd}</td>
                <td>${s.actualStart || 'TBD'} to ${s.actualEnd || 'TBD'}</td>
                <td>${s.progress}%</td>
                <td class="${statusClass}">${s.status}</td>
            </tr>`;
    });

    wordHtml += `
        </table>

        <h2>4. DELAY REASONS</h2>`;

    if (delays.length === 0) {
        wordHtml += "<p>No delays logged.</p>";
    } else {
        wordHtml += "<ul>";
        delays.forEach(d => {
            const stageName = stages.find(s => s.id === d.stageId)?.name || 'Unknown Stage';
            wordHtml += `
            <li>
                <strong>Stage:</strong> ${stageName}<br>
                <strong>Reason:</strong> ${d.reason}<br>
                <strong>Impact:</strong> ${d.impact}
            </li>`;
        });
        wordHtml += "</ul>";
    }

    wordHtml += `
        <h2>5. LESSONS LEARNED</h2>`;

    if (lessons.length === 0) {
        wordHtml += "<p>No lessons logged.</p>";
    } else {
        wordHtml += "<ul>";
        lessons.forEach(l => {
            const stageName = l.stageId ? (stages.find(s => s.id === l.stageId)?.name || 'Unknown Stage') : 'General Project';
            wordHtml += `
            <li>
                <strong>${stageName}</strong><br>
                <strong>Description:</strong> ${l.lessonDesc}<br>
                <strong>Recommendation:</strong> ${l.recommendation}
            </li>`;
        });
        wordHtml += "</ul>";
    }

    wordHtml += `
    </body>
    </html>`;

    // Save report to DB (update or create)
    const newReport = {
        id: Date.now().toString(),
        projectId: projectId,
        currentStageStatus: projectStatus,
        overallProgress: totalProgress,
        executiveSummary: executiveSummary,
        keyDelaysSummary: numDelays > 0 ? `${numDelays} delay(s) recorded` : 'No delays recorded',
        lessonsLearnedSummary: lessons.length > 0 ? `${lessons.length} lesson(s) recorded` : 'No lessons recorded',
        content: wordHtml,
        createdAt: new Date().toISOString()
    };

    let reports = DB.get('projectReports') || [];
    const existingIndex = reports.findIndex(r => r.projectId === projectId);
    if (existingIndex !== -1) {
        reports[existingIndex] = { ...reports[existingIndex], ...newReport };
    } else {
        reports.push(newReport);
    }
    DB.set('projectReports', reports);

    // Trigger download
    const fileName = `Project_Report_${project.name.replace(/\s+/g, '_')}.doc`;
    triggerWordDocDownload(wordHtml, fileName);

    showToast('Report generated and downloaded.', 'success');

    // Refresh the completed projects table to show updated button states
    renderProjectsTab();
}

// Trigger browser download of a Word-compatible HTML document
function triggerWordDocDownload(htmlContent, fileName) {
    const blob = new Blob([htmlContent], { type: 'application/msword' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = fileName;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
}

function openProjectDetail(projectId) {
    currentProjectId = projectId;
    const project = DB.get('projects').find(p => p.id === projectId);
    if (!project) return;

    document.getElementById('detail-project-name').textContent = project.name;
    document.getElementById('detail-project-desc').textContent = project.description || 'No description.';

    renderStages();
    renderLessons();
    updateReportUI();
    showView('projectDetail');

    // Re-fetch backend data in background and refresh
    fetchAllData(1).then(() => {
        // Double check we are still on the same project view
        if (currentProjectId === projectId && document.getElementById('project-detail-view')?.classList.contains('active')) {
            renderStages();
            renderLessons();
            updateReportUI();
        }
    });
}

function handleSaveStage(e) {
    e.preventDefault();
    const id = document.getElementById('stage-id').value;
    const name = document.getElementById('stage-name').value.trim();
    const plannedStart = document.getElementById('stage-planned-start').value;
    const plannedEnd = document.getElementById('stage-planned-end').value;
    const actualStart = document.getElementById('stage-actual-start').value;
    const actualEnd = document.getElementById('stage-actual-end').value;
    const progress = parseInt(document.getElementById('stage-progress').value) || 0;

    // Auto-calculate status
    const todayStr = new Date().toISOString().split('T')[0];
    let status = 'On Track';
    if (progress === 100) {
        status = 'Completed';
    } else if (todayStr > plannedEnd) {
        status = 'Behind Schedule';
    }

    const stages = DB.get('stages');

    if (id) {
        // Edit existing
        const index = stages.findIndex(s => s.id === id);
        if (index !== -1) {
            stages[index] = {
                ...stages[index],
                name, plannedStart, plannedEnd, actualStart, actualEnd, progress, status
            };
        }
    } else {
        // Create new
        const newStage = {
            id: Date.now().toString(),
            projectId: currentProjectId,
            name, plannedStart, plannedEnd, actualStart, actualEnd, progress, status
        };
        stages.push(newStage);
    }

    DB.set('stages', stages);

    // Evaluate project completion after stage progress changes
    evaluateProjectCompletion();

    toggleModal('modal-stage');
    showToast('Stage saved successfully', 'success');
    renderStages();
}

function renderStages() {
    updateAllStageStatuses();
    const allStages = DB.get('stages');
    const pStages = allStages.filter(s => s.projectId === currentProjectId);

    // Sort by planned start date
    pStages.sort((a, b) => new Date(a.plannedStart) - new Date(b.plannedStart));

    const container = document.getElementById('stages-container');
    container.innerHTML = '';

    if (pStages.length === 0) {
        container.innerHTML = '<div style="text-align: center; color: var(--text-muted); padding: 2rem;">No stages added yet.</div>';
    } else {
        pStages.forEach(stage => {
            const el = document.createElement('div');
            el.className = 'stage-card glass-panel';

            let statusClass = 'status-ontrack';
            if (stage.status === 'Behind Schedule') statusClass = 'status-behind';
            if (stage.status === 'Completed') statusClass = 'status-completed';

            const aStartStr = stage.actualStart ? new Date(stage.actualStart).toLocaleDateString() : 'TBD';
            const aEndStr = stage.actualEnd ? new Date(stage.actualEnd).toLocaleDateString() : 'TBD';
            const pStartStr = new Date(stage.plannedStart).toLocaleDateString();
            const pEndStr = new Date(stage.plannedEnd).toLocaleDateString();

            let actionHtml = '';
            let extraInfo = '';
            if (currentUser.role !== 'Viewer') {
                actionHtml = `
                    <div class="stage-actions">
                        <button class="icon-btn edit-stage-btn" data-id="${stage.id}"><i data-lucide="edit-3"></i></button>
                    </div>
                `;
            }

            if (stage.status === 'Behind Schedule') {
                const delays = DB.get('delayRecords').filter(d => d.stageId === stage.id);
                if (delays.length > 0) {
                    extraInfo = `<div style="margin-top: 1.5rem; padding: 1rem; background: rgba(245, 158, 11, 0.1); border-left: 3px solid var(--warning); border-radius: 4px; width: 100%;">
                        <strong style="color: var(--warning); font-size: 0.9rem;"><i data-lucide="alert-triangle" style="width: 14px; height: 14px; margin-right: 4px; vertical-align: middle;"></i>Delay Logged:</strong>
                        <p style="font-size: 0.85rem; margin-top: 0.5rem; color: var(--text-muted)"><strong>Reason:</strong> ${delays[0].reason}</p>
                        <p style="font-size: 0.85rem; margin-top: 0.25rem; color: var(--text-muted)"><strong>Impact:</strong> ${delays[0].impact}</p>
                    </div>`;
                } else if (currentUser.role !== 'Viewer') {
                    extraInfo = `<div style="margin-top: 1.5rem; width: 100%;"><button class="btn-secondary log-delay-btn" data-id="${stage.id}" style="color: var(--warning); border-color: rgba(245, 158, 11, 0.3); font-size: 0.8rem; padding: 0.5rem 1rem;"><i data-lucide="alert-circle" style="width: 16px; height: 16px;"></i> Log Delay Info</button></div>`;
                }
            }

            el.style.flexDirection = 'column';
            el.style.alignItems = 'flex-start';

            el.innerHTML = `
                <div style="display: flex; justify-content: space-between; align-items: center; width: 100%;">
                    <div class="stage-info">
                        <h4>
                            <i data-lucide="target" style="color: var(--primary); width: 18px;"></i>
                            ${stage.name}
                        </h4>
                        <div class="stage-dates">
                            <div class="stage-dates-col">
                                <span class="date-row" title="Planned Start">
                                    <i data-lucide="calendar"></i> ${pStartStr} 
                                </span>
                                <span class="date-row" title="Actual Start">
                                    <i data-lucide="clock"></i> ${aStartStr}
                                </span>
                            </div>
                            <div class="stage-dates-col">
                                <span class="date-row" title="Planned End">
                                    <i data-lucide="calendar"></i> ${pEndStr}
                                </span>
                                <span class="date-row" title="Actual End">
                                    <i data-lucide="clock"></i> ${aEndStr}
                                </span>
                            </div>
                        </div>
                    </div>
                    
                    <div style="display: flex; align-items: center; gap: 1rem;">
                        <div class="stage-stats">
                            <span class="status-badge ${statusClass}">${stage.status}</span>
                            <div class="stage-progress">
                                <span style="font-weight: 600; font-size: 0.9rem;">${stage.progress}%</span>
                                <div class="progress-bar-bg">
                                    <div class="progress-bar-fill" style="width: ${stage.progress}%"></div>
                                </div>
                            </div>
                        </div>
                        ${actionHtml}
                    </div>
                </div>
                ${extraInfo}
            `;
            container.appendChild(el);
        });
    }

    // Attach edit listeners
    document.querySelectorAll('.edit-stage-btn').forEach(btn => {
        btn.addEventListener('click', (e) => {
            const id = e.currentTarget.getAttribute('data-id');
            openEditStageModal(id);
        });
    });

    document.querySelectorAll('.log-delay-btn').forEach(btn => {
        btn.addEventListener('click', (e) => {
            const id = e.currentTarget.getAttribute('data-id');
            document.getElementById('form-delay').reset();
            document.getElementById('delay-stage-id').value = id;
            toggleModal('modal-delay');
        });
    });

    lucide.createIcons();
    updateProjectTotalProgress(pStages);
}

function openEditStageModal(stageId) {
    const stage = DB.get('stages').find(s => s.id === stageId);
    if (!stage) return;

    document.getElementById('modal-stage-title').textContent = 'Edit Stage';
    document.getElementById('stage-id').value = stage.id;
    document.getElementById('stage-name').value = stage.name;
    document.getElementById('stage-planned-start').value = stage.plannedStart;
    document.getElementById('stage-planned-end').value = stage.plannedEnd;
    document.getElementById('stage-actual-start').value = stage.actualStart || '';
    document.getElementById('stage-actual-end').value = stage.actualEnd || '';
    document.getElementById('stage-progress').value = stage.progress;
    document.getElementById('stage-status').value = stage.status;

    toggleModal('modal-stage');
}

function updateProjectTotalProgress(stages) {
    const ringCircle = document.getElementById('project-total-progress');
    const ringText = document.getElementById('project-total-progress-text');

    if (!ringCircle) return;

    let totalProgress = 0;
    if (stages.length > 0) {
        totalProgress = Math.round(stages.reduce((sum, s) => sum + parseInt(s.progress), 0) / stages.length);
    }

    ringText.textContent = `${totalProgress}%`;

    const radius = ringCircle.r.baseVal.value;
    const circumference = radius * 2 * Math.PI;

    ringCircle.style.strokeDasharray = `${circumference} ${circumference}`;
    const offset = circumference - (totalProgress / 100) * circumference;
    ringCircle.style.strokeDashoffset = offset;

    if (totalProgress === 100) {
        ringCircle.style.stroke = 'var(--success)';
    } else {
        ringCircle.style.stroke = 'var(--primary)';
    }
}

function createBehindScheduleNotification(stage) {
    const notifications = DB.get('notifications');
    const existing = notifications.find(n => n.stageId === stage.id && n.message.includes('behind schedule'));
    if (existing) return;

    const project = DB.get('projects').find(p => p.id === stage.projectId);
    const users = DB.get('users');
    const targetUsers = users.filter(u => u.role === 'Admin' || u.role === 'Project Manager');

    targetUsers.forEach(u => {
        notifications.push({
            id: Date.now().toString() + '-' + u.id,
            userId: u.id,
            projectId: project ? project.id : '',
            stageId: stage.id,
            message: `Stage "${stage.name}" in project "${project ? project.name : ''}" is behind schedule.`,
            read: false,
            createdAt: new Date().toISOString()
        });
    });
    DB.set('notifications', notifications);
}

function renderNotifications() {
    if (!currentUser) return;
    const notifications = DB.get('notifications').filter(n => n.userId === currentUser.id);
    notifications.sort((a, b) => new Date(b.createdAt) - new Date(a.createdAt));

    const unreadCount = notifications.filter(n => !n.read).length;
    const badge = document.getElementById('unread-notifications-badge');
    if (badge) {
        if (unreadCount > 0) {
            badge.style.display = 'inline-block';
            badge.textContent = unreadCount;
        } else {
            badge.style.display = 'none';
        }
    }

    const container = document.getElementById('notifications-container');
    if (!container) return;
    container.innerHTML = '';

    if (notifications.length === 0) {
        container.innerHTML = '<div style="color: var(--text-muted); font-size: 0.9rem;">No notifications.</div>';
    } else {
        const unreadList = notifications.filter(n => !n.read);
        const readList = notifications.filter(n => n.read);
        const sortedList = [...unreadList, ...readList];

        sortedList.forEach(n => {
            const el = document.createElement('div');
            el.style.padding = '1rem';
            el.style.background = n.read ? 'rgba(0,0,0,0.2)' : 'rgba(99, 102, 241, 0.1)';
            el.style.borderLeft = n.read ? '3px solid transparent' : '3px solid var(--primary)';
            el.style.borderRadius = '6px';
            el.style.cursor = 'pointer';
            el.style.fontSize = '0.9rem';

            el.innerHTML = `
                <div style="font-weight: ${n.read ? '400' : '600'}; margin-bottom: 0.5rem; color: ${n.read ? 'var(--text-muted)' : 'var(--text-main)'}">${n.message}</div>
                <div style="color: var(--text-muted); font-size: 0.8rem;">${new Date(n.createdAt).toLocaleString()}</div>
            `;

            el.addEventListener('click', () => {
                if (!n.read) {
                    const allNotifs = DB.get('notifications');
                    const target = allNotifs.find(x => x.id === n.id);
                    if (target) {
                        target.read = true;
                        DB.set('notifications', allNotifs);
                        renderNotifications();
                    }
                }
                const panel = document.getElementById('notifications-panel');
                if (panel) panel.style.display = 'none';

                openProjectDetail(n.projectId);
            });
            container.appendChild(el);
        });
    }
}

function handleSaveDelay(e) {
    e.preventDefault();
    const stageId = document.getElementById('delay-stage-id').value;
    const reason = document.getElementById('delay-reason').value.trim();
    const impact = document.getElementById('delay-impact').value.trim();

    const delays = DB.get('delayRecords');
    delays.push({
        id: Date.now().toString(),
        projectId: currentProjectId,
        stageId,
        reason,
        impact,
        createdAt: new Date().toISOString()
    });
    DB.set('delayRecords', delays);
    toggleModal('modal-delay');
    showToast('Delay information logged successfully.', 'success');
    renderStages();
}

function handleSaveLesson(e) {
    e.preventDefault();
    const stageId = document.getElementById('lesson-stage').value;
    const lessonDesc = document.getElementById('lesson-desc').value.trim();
    const recommendation = document.getElementById('lesson-rec').value.trim();

    const lessons = DB.get('lessonsLearned');
    lessons.push({
        id: Date.now().toString(),
        projectId: currentProjectId,
        stageId: stageId || null,
        lessonDesc,
        recommendation,
        recordedAt: new Date().toISOString()
    });
    DB.set('lessonsLearned', lessons);
    toggleModal('modal-lesson');
    showToast('Lesson saved.', 'success');
    renderLessons();
}

function renderLessons() {
    const lessons = DB.get('lessonsLearned').filter(l => l.projectId === currentProjectId);
    const container = document.getElementById('lessons-container');
    const stages = DB.get('stages');

    container.innerHTML = '';
    if (lessons.length === 0) {
        container.innerHTML = '<div style="color: var(--text-muted); font-size: 0.9rem;">No lessons recorded yet.</div>';
        return;
    }

    lessons.sort((a, b) => new Date(b.recordedAt) - new Date(a.recordedAt));

    lessons.forEach(l => {
        const el = document.createElement('div');
        el.className = 'glass-panel';
        el.style.padding = '1.5rem';

        let stageName = 'General Lesson';
        if (l.stageId) {
            const stage = stages.find(s => s.id === l.stageId);
            if (stage) stageName = 'Stage: ' + stage.name;
        }

        el.innerHTML = `
            <div style="color: var(--primary); font-size: 0.8rem; font-weight: 600; text-transform: uppercase; margin-bottom: 0.5rem;"><i data-lucide="book-open" style="width: 14px; height: 14px; margin-right: 4px; vertical-align: middle;"></i>${stageName}</div>
            <div style="margin-bottom: 1rem;">
                <h4 style="font-size: 0.95rem; margin-bottom: 0.25rem;">Description</h4>
                <p style="color: var(--text-muted); font-size: 0.9rem; line-height: 1.6;">${l.lessonDesc}</p>
            </div>
            <div>
                <h4 style="font-size: 0.95rem; margin-bottom: 0.25rem;">Recommendation</h4>
                <p style="color: var(--text-muted); font-size: 0.9rem; line-height: 1.6;">${l.recommendation}</p>
            </div>
        `;
        container.appendChild(el);
    });
    lucide.createIcons();
}

function updateReportUI() {
    const reportStatus = document.getElementById('project-report-status');
    const genBtn = document.getElementById('generate-report-btn');
    const dlBtn = document.getElementById('download-report-btn');

    if (!reportStatus || !genBtn || !dlBtn) return;

    reportStatus.style.display = 'none';
    genBtn.style.display = 'none';
    dlBtn.style.display = 'none';

    const project = DB.get('projects').find(p => p.id === currentProjectId);
    if (!project || currentUser.role === 'Viewer') return;

    reportStatus.style.display = 'inline-block';
    genBtn.style.display = 'flex'; // Report can be generated any number of times during the project

    const reports = DB.get('projectReports') || [];
    const report = reports.find(r => r.projectId === currentProjectId);

    if (report) {
        reportStatus.textContent = `Report: Generated (${new Date(report.createdAt).toLocaleDateString()})`;
        reportStatus.style.color = 'var(--success)';
        dlBtn.style.display = 'flex';
        // Use Blob-based download for reliability with large documents
        dlBtn.href = '#';
        dlBtn.download = `Project_Report_${project.name.replace(/\s+/g, '_')}.doc`;
        dlBtn.onclick = (e) => {
            e.preventDefault();
            triggerWordDocDownload(report.content, `Project_Report_${project.name.replace(/\s+/g, '_')}.doc`);
        };
    } else {
        reportStatus.textContent = 'Report: Pending';
        reportStatus.style.color = 'var(--text-muted)';
    }
}

function handleGenerateReport() {
    const project = DB.get('projects').find(p => p.id === currentProjectId);
    if (!project) return;

    const stages = DB.get('stages').filter(s => s.projectId === currentProjectId);
    const delays = DB.get('delayRecords').filter(d => d.projectId === currentProjectId);
    const lessons = DB.get('lessonsLearned').filter(l => l.projectId === currentProjectId);

    let totalProgress = stages.length ? Math.round(stages.reduce((sum, s) => sum + parseInt(s.progress), 0) / stages.length) : 0;

    let projectStatus = 'On Track';
    const todayStr = new Date().toISOString().split('T')[0];

    // Determine overall project status
    if (totalProgress === 100) {
        projectStatus = 'Completed';
    } else {
        stages.forEach(s => {
            if (parseInt(s.progress) < 100 && todayStr > s.plannedEnd) {
                projectStatus = 'Behind Schedule';
            }
        });
    }

    let numDelays = delays.length;
    let executiveSummary = `Project "${project.name}" is currently ${projectStatus}. Total average progress is ${totalProgress}%. Delays logged to date: ${numDelays}. Lessons recorded to date: ${lessons.length}.`;

    stages.sort((a, b) => new Date(a.plannedStart) - new Date(b.plannedStart));

    // Constructing an HTML payload that MS Word can interpret
    let wordHtml = `
    <html xmlns:o='urn:schemas-microsoft-com:office:office' xmlns:w='urn:schemas-microsoft-com:office:word' xmlns='http://www.w3.org/TR/REC-html40'>
    <head>
        <meta charset='utf-8'>
        <title>Project Report - ${project.name}</title>
        <style>
            body { font-family: 'Arial', sans-serif; line-height: 1.6; color: #333; }
            h1 { color: #2c3e50; text-align: center; border-bottom: 2px solid #3498db; padding-bottom: 10px; }
            h2 { color: #34495e; margin-top: 20px; }
            table { width: 100%; border-collapse: collapse; margin-top: 15px; }
            th, td { border: 1px solid #bdc3c7; padding: 8px; text-align: left; }
            th { background-color: #ecf0f1; font-weight: bold; }
            .behind-schedule { color: #e74c3c; font-weight: bold; }
            .on-track { color: #27ae60; }
            .completed { color: #2980b9; }
            .meta-info { margin-bottom: 20px; font-size: 14px; color: #7f8c8d; text-align: right; }
        </style>
    </head>
    <body>
        <h1>PROJECT REPORT: ${project.name}</h1>
        <div class="meta-info">Generated At: ${new Date().toLocaleString()}</div>

        <h2>1. EXECUTIVE SUMMARY</h2>
        <p>${executiveSummary}</p>

        <h2>2. PROJECT OVERVIEW</h2>
        <ul>
            <li><strong>Name:</strong> ${project.name}</li>
            <li><strong>Description:</strong> ${project.description || 'N/A'}</li>
            <li><strong>Creation Date:</strong> ${new Date(project.createdAt).toLocaleDateString()}</li>
            <li><strong>Total Stages:</strong> ${stages.length}</li>
            <li><strong>Overall Progress:</strong> ${totalProgress}%</li>
            <li><strong>Current Status:</strong> ${projectStatus}</li>
        </ul>

        <h2>3. STAGE-BY-STAGE PERFORMANCE</h2>
        <table>
            <tr>
                <th>Stage Name</th>
                <th>Planned Start / End</th>
                <th>Actual Start / End</th>
                <th>Progress</th>
                <th>Status</th>
            </tr>`;

    stages.forEach(s => {
        let statusClass = '';
        if (s.status === 'Behind Schedule') statusClass = 'behind-schedule';
        else if (s.status === 'Completed') statusClass = 'completed';
        else statusClass = 'on-track';

        wordHtml += `
            <tr>
                <td>${s.name}</td>
                <td>${s.plannedStart} to ${s.plannedEnd}</td>
                <td>${s.actualStart || 'TBD'} to ${s.actualEnd || 'TBD'}</td>
                <td>${s.progress}%</td>
                <td class="${statusClass}">${s.status}</td>
            </tr>`;
    });

    wordHtml += `
        </table>

        <h2>4. DELAY REASONS</h2>`;

    if (delays.length === 0) {
        wordHtml += "<p>No delays logged.</p>";
    } else {
        wordHtml += "<ul>";
        delays.forEach(d => {
            const stageName = stages.find(s => s.id === d.stageId)?.name || 'Unknown Stage';
            wordHtml += `
            <li>
                <strong>Stage:</strong> ${stageName}<br>
                <strong>Reason:</strong> ${d.reason}<br>
                <strong>Impact:</strong> ${d.impact}
            </li>`;
        });
        wordHtml += "</ul>";
    }

    wordHtml += `
        <h2>5. LESSONS LEARNED</h2>`;

    if (lessons.length === 0) {
        wordHtml += "<p>No lessons logged.</p>";
    } else {
        wordHtml += "<ul>";
        lessons.forEach(l => {
            const stageName = l.stageId ? (stages.find(s => s.id === l.stageId)?.name || 'Unknown Stage') : 'General Project';
            wordHtml += `
            <li>
                <strong>${stageName}</strong><br>
                <strong>Description:</strong> ${l.lessonDesc}<br>
                <strong>Recommendation:</strong> ${l.recommendation}
            </li>`;
        });
        wordHtml += "</ul>";
    }

    wordHtml += `
    </body>
    </html>`;

    const newReport = {
        id: Date.now().toString(),
        projectId: currentProjectId,
        currentStageStatus: projectStatus,
        overallProgress: totalProgress,
        executiveSummary: executiveSummary,
        keyDelaysSummary: numDelays > 0 ? `${numDelays} delay(s) recorded` : 'No delays recorded',
        lessonsLearnedSummary: lessons.length > 0 ? `${lessons.length} lesson(s) recorded` : 'No lessons recorded',
        content: wordHtml,
        createdAt: new Date().toISOString()
    };

    let reports = DB.get('projectReports') || [];

    // Update or Overwrite existing record
    const existingIndex = reports.findIndex(r => r.projectId === currentProjectId);
    if (existingIndex !== -1) {
        reports[existingIndex] = { ...reports[existingIndex], ...newReport };
    } else {
        reports.push(newReport);
    }

    DB.set('projectReports', reports);

    showToast('Word Report generated successfully.', 'success');
    updateReportUI();
}

// --- Admin Controls ---
function renderUsersTable(query = '') {
    const tbody = document.getElementById('users-table-body');
    if (!tbody) return;

    const users = DB.get('users');
    const q = query.toLowerCase();

    // Filter by name or email (fallback gracefully)
    const filtered = users.filter(u =>
        (u.username && u.username.toLowerCase().includes(q)) ||
        (u.email && u.email.toLowerCase().includes(q))
    ).sort((a, b) => new Date(b.createdAt || 0) - new Date(a.createdAt || 0));

    tbody.innerHTML = '';

    if (filtered.length === 0) {
        tbody.innerHTML = '<tr><td colspan="7" style="text-align: center; padding: 2rem;">No users found.</td></tr>';
        return;
    }

    filtered.forEach(u => {
        const isSelf = u.id === currentUser.id;

        // Inline styles inherit easily into dynamic themes where required since they use CSS vars,
        // but inputs/select native items require direct text/bg coloring
        let roleOptions = '';
        if (u.email === 'fredadeefe224@gmail.com') {
            roleOptions = `<option value="Admin" selected>Admin</option>`;
        } else {
            roleOptions = `
                <option value="Project Manager" ${u.role === 'Project Manager' ? 'selected' : ''}>Project Manager</option>
                <option value="Viewer" ${u.role === 'Viewer' ? 'selected' : ''}>Viewer</option>
            `;
        }

        const isAdminEmail = (u.email === 'fredadeefe224@gmail.com');
        const roleSelect = (isSelf || isAdminEmail) ?
            `<select disabled style="background:var(--input-bg); color:var(--text-main); border:1px solid var(--panel-border); padding:0.25rem; border-radius:4px;">${roleOptions}</select>` :
            `<select class="role-select" data-id="${u.id}" style="background:var(--input-bg); color:var(--text-main); border:1px solid var(--panel-border); padding:0.25rem; border-radius:4px;">
                ${roleOptions}
            </select>`;

        const disableBtn = (isSelf || isAdminEmail) ?
            `<button class="btn-secondary" disabled style="padding:0.25rem 0.5rem; opacity:0.5; border:none; cursor:not-allowed;">Disabled</button>` :
            `<button class="btn-secondary toggle-status-btn" data-id="${u.id}" style="padding:0.25rem 0.6rem; font-size:0.8rem; background:${u.isDisabled ? 'var(--success)' : 'var(--danger)'}; border:none; color:#fff;">
                ${u.isDisabled ? 'Enable' : 'Disable'}
             </button>`;

        const joinedDate = u.createdAt ? new Date(u.createdAt).toLocaleDateString() : 'N/A';
        const lastLoginDate = u.lastLogin ? new Date(u.lastLogin).toLocaleDateString() : 'N/A';

        const tr = document.createElement('tr');
        tr.style.borderBottom = '1px solid var(--panel-border)';
        if (u.isDisabled) tr.style.opacity = '0.6';

        tr.innerHTML = `
            <td style="padding: 1rem; font-weight: 500;">${u.username}</td>
            <td style="padding: 1rem; color: var(--text-muted);">${u.email || 'N/A'}</td>
            <td style="padding: 1rem;">${roleSelect}</td>
            <td style="padding: 1rem;">
                <span class="badge" style="background:${u.isDisabled ? 'rgba(239, 68, 68, 0.2)' : 'rgba(16, 185, 129, 0.2)'}; color:${u.isDisabled ? 'var(--danger)' : 'var(--success)'};">${u.isDisabled ? 'Disabled' : 'Active'}</span>
            </td>
            <td style="padding: 1rem; color: var(--text-muted);">${joinedDate}</td>
            <td style="padding: 1rem; color: var(--text-muted);">${lastLoginDate}</td>
            <td style="padding: 1rem; text-align: right;">${disableBtn}</td>
        `;

        tbody.appendChild(tr);
    });

    // Add event listeners within the newly generated table DOM hooks
    document.querySelectorAll('.role-select').forEach(sel => {
        sel.addEventListener('change', (e) => {
            const userId = e.target.getAttribute('data-id');
            const newRole = e.target.value;
            confirmUserAction('Change Role', `Are you sure you want to change this user's role to ${newRole}?`, () => {
                const usersDb = DB.get('users');
                const idx = usersDb.findIndex(x => x.id === userId);
                if (idx !== -1) {
                    usersDb[idx].role = newRole;
                    DB.set('users', usersDb);
                    showToast('Role updated successfully.', 'success');
                    renderUsersTable(document.getElementById('search-users-input').value);
                }
            });
        });
    });

    document.querySelectorAll('.toggle-status-btn').forEach(btn => {
        btn.addEventListener('click', (e) => {
            const userId = e.target.getAttribute('data-id');
            const usersDb = DB.get('users');
            const user = usersDb.find(x => x.id === userId);
            const willDisable = !user.isDisabled;
            const actionVerb = willDisable ? 'disable' : 'enable';

            confirmUserAction(`${willDisable ? 'Disable' : 'Enable'} Account`, `Are you sure you want to ${actionVerb} this account? Users disabled will not be able to log in.`, () => {
                user.isDisabled = willDisable;
                DB.set('users', usersDb);
                showToast(`Account ${actionVerb}d successfully.`, 'success');
                renderUsersTable(document.getElementById('search-users-input').value);
            });
        });
    });
}

function confirmUserAction(title, message, onConfirm) {
    document.getElementById('confirm-action-title').textContent = title;
    document.getElementById('confirm-action-message').textContent = message;

    actionPending = onConfirm; // Bound to global click listener on #confirm-action-btn
    toggleModal('modal-confirm-action');
}
