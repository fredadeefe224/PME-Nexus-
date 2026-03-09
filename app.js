// --- Database Service (Antigravity Cloud Backend Integration) ---
window.dbStore = {
    users: [], projects: [], stages: [], notifications: [], delayRecords: [], lessonsLearned: [], projectReports: [], documents: []
};


const API_BASE = "https://pme-nexus.onrender.com";

// This tells the browser: "Stay on the same website I'm on, but go to /api/data"
fetch(API_BASE)
    .then(response => response.json())
    .then(data => {
        // Your code to display data
    })
    .catch(err => console.error("Fetch failed:", err));

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
                const collections = ['users', 'projects', 'stages', 'notifications', 'delayRecords', 'lessonsLearned', 'projectReports', 'documents'];
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
        const collections = ['users', 'projects', 'stages', 'notifications', 'delayRecords', 'lessonsLearned', 'projectReports', 'documents'];
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

    // --- Cloudinary Upload Widget ---
    document.getElementById('upload-document-btn')?.addEventListener('click', () => {
        if (!currentUser || (currentUser.role !== 'Project Manager' && currentUser.role !== 'Admin')) {
            showToast('Only Project Managers and Admins can upload documents.', 'error');
            return;
        }

        const widget = cloudinary.createUploadWidget({
            cloudName: 'dz11ualzg',
            uploadPreset: 'PME Nexus Cloud',
            sources: ['local', 'url', 'google_drive', 'dropbox'],
            clientAllowedFormats: ['pdf', 'docx', 'xlsx'],
            maxFileSize: 10000000, // 10MB
            multiple: true,
            showAdvancedOptions: false,
            cropping: false,
            showSkipCropButton: true,
            folder: 'pme-nexus-documents',
            resourceType: 'raw',
            theme: document.body.classList.contains('light-theme') ? 'white' : 'minimal'
        }, (error, result) => {
            if (error) {
                console.error('[Cloudinary Error]', error);
                showToast('Upload failed. Please try again.', 'error');
                return;
            }

            if (result.event === 'success') {
                const info = result.info;
                const fileUrl = info.secure_url;
                const originalName = info.original_filename || info.public_id;
                const fileFormat = info.format || info.original_extension || getExtensionFromName(originalName);

                // Determine file type
                let fileType = 'other';
                if (fileFormat === 'pdf') fileType = 'pdf';
                else if (fileFormat === 'docx' || fileFormat === 'doc') fileType = 'docx';
                else if (fileFormat === 'xlsx' || fileFormat === 'xls') fileType = 'xlsx';

                const docPayload = {
                    name: originalName + (fileFormat ? '.' + fileFormat : ''),
                    url: fileUrl,
                    type: fileType,
                    uploadedBy: currentUser.username,
                    date: new Date().toISOString()
                };

                // Save to backend
                fetch(`${API_BASE}/api/documents`, {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify(docPayload)
                })
                    .then(res => res.json())
                    .then(data => {
                        if (data.success) {
                            // Add to local store
                            if (!Array.isArray(window.dbStore.documents)) {
                                window.dbStore.documents = [];
                            }
                            window.dbStore.documents.push(data.document);
                            showToast(`"${docPayload.name}" uploaded successfully!`, 'success');
                            renderDocumentLibrary();
                        } else {
                            showToast('Failed to save document record.', 'error');
                        }
                    })
                    .catch(err => {
                        console.error('[Doc Save Error]', err);
                        showToast('Failed to save document record.', 'error');
                    });
            }
        });

        widget.open();
    });
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

    // Upload Document button: only visible to Admin and Project Manager
    const uploadDocBtn = document.getElementById('upload-document-btn');
    if (uploadDocBtn) {
        uploadDocBtn.style.display = (currentUser.role === 'Project Manager' || currentUser.role === 'Admin') ? 'flex' : 'none';
    }

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

    // Render document library
    renderDocumentLibrary();
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
                    <button class="btn-download-report" onclick="downloadProjectReport('${p.id}', this)">
                        <i data-lucide="download"></i> Download
                    </button>
                    <button class="btn-download-report btn-generate" onclick="generateAndDownloadReport('${p.id}', this)" title="Regenerate report with latest data">
                        <i data-lucide="refresh-cw"></i>
                    </button>
                </div>`;
        } else {
            // No report yet — show Generate & Download button
            reportBtns = `<button class="btn-download-report btn-generate" onclick="generateAndDownloadReport('${p.id}', this)">
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
async function downloadProjectReport(projectId, btn) {
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
    if (!report) {
        showToast('No report found. Please generate one first.', 'error');
        return;
    }

    // Re-build the .docx from current data for download
    const project = DB.get('projects').find(p => p.id === projectId);
    if (!project) {
        showToast('Project not found.', 'error');
        return;
    }

    const stages = DB.get('stages').filter(s => s.projectId === projectId);
    const delays = DB.get('delayRecords').filter(d => d.projectId === projectId);
    const lessons = DB.get('lessonsLearned').filter(l => l.projectId === projectId);

    let totalProgress = stages.length ? Math.round(stages.reduce((sum, s) => sum + parseInt(s.progress), 0) / stages.length) : 0;
    let projectStatus = report.currentStageStatus || 'On Track';
    let executiveSummary = report.executiveSummary || `Project "${project.name}" report.`;

    stages.sort((a, b) => new Date(a.plannedStart) - new Date(b.plannedStart));

    // Disable button for entire operation (AI fetch + docx build + download)
    const originalLabel = btn ? btn.innerHTML : null;
    if (btn) { btn.disabled = true; btn.innerHTML = '<i data-lucide="loader" class="spin-icon"></i> Enhancing...'; lucide.createIcons(); }

    try {
        // --- AI Enhancement: fetch polished executive summary from backend ---
        try {
            const aiRes = await fetch(`${API_BASE}/api/enhance-report`, {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({
                    projectName: project.name,
                    projectDescription: project.description || '',
                    projectStatus,
                    totalProgress,
                    stages: stages.map(s => ({ name: s.name, status: s.status, progress: s.progress, plannedStart: s.plannedStart, plannedEnd: s.plannedEnd, actualStart: s.actualStart, actualEnd: s.actualEnd })),
                    delays: delays.map(d => ({ stageId: d.stageId, reason: d.reason, impact: d.impact })),
                    lessons: lessons.map(l => ({ stageId: l.stageId, lessonDesc: l.lessonDesc, recommendation: l.recommendation }))
                })
            });
            const aiData = await aiRes.json();
            if (aiData.success && aiData.aiText) {
                executiveSummary = aiData.aiText;
            }
        } catch (aiErr) {
            console.warn('[AI ENHANCE] Fallback to raw summary:', aiErr.message);
        }

        if (btn) { btn.innerHTML = '<i data-lucide="loader" class="spin-icon"></i> Generating...'; lucide.createIcons(); }

        const doc = buildDocxReport({
            project, stages, delays, lessons,
            totalProgress, projectStatus, executiveSummary
        });

        const fileName = `Project_Report_${(project?.name || 'Report').replace(/\s+/g, '_')}.docx`;
        await triggerDocxDownload(doc, fileName);
        showToast('Report downloaded successfully.', 'success');
    } catch (err) {
        console.error('[DOWNLOAD ERROR]', err);
        showToast('Failed to generate report. Please try again.', 'error');
    } finally {
        if (btn) { btn.disabled = false; btn.innerHTML = originalLabel || '<i data-lucide="download"></i> Download'; lucide.createIcons(); }
    }
}

// Generate a fresh report for a project and immediately download it
async function generateAndDownloadReport(projectId, btn) {
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

    // Disable button for entire operation (AI fetch + docx build + download)
    const originalLabel = btn ? btn.innerHTML : null;
    if (btn) { btn.disabled = true; btn.innerHTML = '<i data-lucide="loader" class="spin-icon"></i> Enhancing...'; lucide.createIcons(); }

    try {
        // --- AI Enhancement: fetch polished executive summary from backend ---
        try {
            const aiRes = await fetch(`${API_BASE}/api/enhance-report`, {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({
                    projectName: project.name,
                    projectDescription: project.description || '',
                    projectStatus,
                    totalProgress,
                    stages: stages.map(s => ({ name: s.name, status: s.status, progress: s.progress, plannedStart: s.plannedStart, plannedEnd: s.plannedEnd, actualStart: s.actualStart, actualEnd: s.actualEnd })),
                    delays: delays.map(d => ({ stageId: d.stageId, reason: d.reason, impact: d.impact })),
                    lessons: lessons.map(l => ({ stageId: l.stageId, lessonDesc: l.lessonDesc, recommendation: l.recommendation }))
                })
            });
            const aiData = await aiRes.json();
            // 👉 ADDED LINE: Let's see exactly what the backend handed us!
            console.log("THE SMOKING GUN:", aiData);
            if (aiData.success && aiData.aiText) {
                executiveSummary = aiData.aiText;
            }
        } catch (aiErr) {
            console.warn('[AI ENHANCE] Fallback to raw summary:', aiErr.message);
        }

        if (btn) { btn.innerHTML = '<i data-lucide="loader" class="spin-icon"></i> Generating...'; lucide.createIcons(); }

        // Build a real .docx document using the docx library
        const doc = buildDocxReport({
            project, stages, delays, lessons,
            totalProgress, projectStatus, executiveSummary
        });

        // Save report metadata to DB (structured data, not HTML)
        const newReport = {
            id: Date.now().toString(),
            projectId: projectId,
            currentStageStatus: projectStatus,
            overallProgress: totalProgress,
            executiveSummary: executiveSummary,
            keyDelaysSummary: numDelays > 0 ? `${numDelays} delay(s) recorded` : 'No delays recorded',
            lessonsLearnedSummary: lessons.length > 0 ? `${lessons.length} lesson(s) recorded` : 'No lessons recorded',
            content: '__docx__', // marker indicating this is a docx-format report
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
        const fileName = `Project_Report_${project.name.replace(/\s+/g, '_')}.docx`;
        await triggerDocxDownload(doc, fileName);

        showToast('Report generated and downloaded.', 'success');

        // Refresh the completed projects table to show updated button states
        renderProjectsTab();
    } catch (err) {
        console.error('[GENERATE ERROR]', err);
        showToast('Failed to generate report. Please try again.', 'error');
    } finally {
        if (btn) { btn.disabled = false; btn.innerHTML = originalLabel || '<i data-lucide="file-text"></i> Generate & Download'; lucide.createIcons(); }
    }
}

// Helper: Builds a docx.Document object from report data

function buildDocxReport({ project, stages, delays, lessons, totalProgress, projectStatus, executiveSummary }) {
    // 1. Safety Check: Did the CDN actually load?
    if (typeof window.docx === "undefined") {
        console.error("CRITICAL ERROR: The docx library failed to load from the CDN.");
        alert("The document generator failed to load. Please disable your ad-blocker or check your firewall.");
        return; // Stop running the code so it doesn't crash
    }
    const { Document, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, HeadingLevel, AlignmentType, BorderStyle, ShadingType } = docx;

    // Color constants
    const BLUE_HEADER = '2c3e50';
    const BLUE_ACCENT = '3498db';
    const GREEN = '27ae60';
    const RED = 'e74c3c';
    const DARK_GREY = '333333';
    const LIGHT_BG = 'ecf0f1';

    function statusColor(status) {
        if (status === 'Behind Schedule') return RED;
        if (status === 'Completed') return '2980b9';
        return GREEN;
    }

    // Build stage performance table rows
    const stageTableRows = [
        new TableRow({
            tableHeader: true,
            children: ['Stage Name', 'Planned Start / End', 'Actual Start / End', 'Progress', 'Status'].map(text =>
                new TableCell({
                    shading: { type: ShadingType.SOLID, color: LIGHT_BG },
                    children: [new Paragraph({ children: [new TextRun({ text, bold: true, font: 'Arial', size: 20, color: DARK_GREY })] })],
                })
            ),
        }),
        ...stages.map(s => new TableRow({
            children: [
                new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: s.name, font: 'Arial', size: 20 })] })] }),
                new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: `${s.plannedStart} to ${s.plannedEnd}`, font: 'Arial', size: 20 })] })] }),
                new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: `${s.actualStart || 'TBD'} to ${s.actualEnd || 'TBD'}`, font: 'Arial', size: 20 })] })] }),
                new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: `${s.progress}%`, font: 'Arial', size: 20 })] })] }),
                new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: s.status, bold: true, font: 'Arial', size: 20, color: statusColor(s.status) })] })] }),
            ]
        }))
    ];

    // Build delay paragraphs
    const delayParagraphs = delays.length === 0
        ? [new Paragraph({ children: [new TextRun({ text: 'No delays logged.', font: 'Arial', size: 22, italics: true, color: '7f8c8d' })] })]
        : delays.map(d => {
            const stageName = stages.find(s => s.id === d.stageId)?.name || 'Unknown Stage';
            return new Paragraph({
                bullet: { level: 0 },
                spacing: { after: 120 },
                children: [
                    new TextRun({ text: 'Stage: ', bold: true, font: 'Arial', size: 22 }),
                    new TextRun({ text: stageName, font: 'Arial', size: 22 }),
                    new TextRun({ text: '\nReason: ', bold: true, font: 'Arial', size: 22, break: 1 }),
                    new TextRun({ text: d.reason, font: 'Arial', size: 22 }),
                    new TextRun({ text: '\nImpact: ', bold: true, font: 'Arial', size: 22, break: 1 }),
                    new TextRun({ text: d.impact, font: 'Arial', size: 22 }),
                ]
            });
        });

    // Build lessons paragraphs
    const lessonParagraphs = lessons.length === 0
        ? [new Paragraph({ children: [new TextRun({ text: 'No lessons logged.', font: 'Arial', size: 22, italics: true, color: '7f8c8d' })] })]
        : lessons.map(l => {
            const stageName = l.stageId ? (stages.find(s => s.id === l.stageId)?.name || 'Unknown Stage') : 'General Project';
            return new Paragraph({
                bullet: { level: 0 },
                spacing: { after: 120 },
                children: [
                    new TextRun({ text: stageName, bold: true, font: 'Arial', size: 22 }),
                    new TextRun({ text: '\nDescription: ', bold: true, font: 'Arial', size: 22, break: 1 }),
                    new TextRun({ text: l.lessonDesc, font: 'Arial', size: 22 }),
                    new TextRun({ text: '\nRecommendation: ', bold: true, font: 'Arial', size: 22, break: 1 }),
                    new TextRun({ text: l.recommendation, font: 'Arial', size: 22 }),
                ]
            });
        });

    // Build overview bullet points
    const overviewItems = [
        { label: 'Name', value: project.name },
        { label: 'Project ID', value: project.id },
        { label: 'Description', value: project.description || 'N/A' },
        { label: 'Creation Date', value: new Date(project.createdAt).toLocaleDateString() },
        { label: 'Total Stages', value: String(stages.length) },
        { label: 'Overall Progress', value: `${totalProgress}%` },
        { label: 'Current Status', value: projectStatus },
    ];
    if (project.completionDate) {
        overviewItems.push({ label: 'Completion Date', value: new Date(project.completionDate).toLocaleDateString() });
    }

    const overviewParagraphs = overviewItems.map(item => new Paragraph({
        bullet: { level: 0 },
        spacing: { after: 60 },
        children: [
            new TextRun({ text: `${item.label}: `, bold: true, font: 'Arial', size: 22 }),
            new TextRun({ text: item.value, font: 'Arial', size: 22 }),
        ]
    }));

    return new Document({
        styles: {
            default: {
                document: {
                    run: { font: 'Arial', size: 22, color: DARK_GREY },
                    paragraph: { spacing: { line: 360 } }
                }
            }
        },
        sections: [{
            children: [
                // Title
                new Paragraph({
                    alignment: AlignmentType.CENTER,
                    spacing: { after: 200 },
                    border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: BLUE_ACCENT } },
                    children: [new TextRun({ text: `PROJECT REPORT: ${project.name}`, bold: true, font: 'Arial', size: 32, color: BLUE_HEADER })],
                }),
                // Generated timestamp
                new Paragraph({
                    alignment: AlignmentType.RIGHT,
                    spacing: { after: 300 },
                    children: [new TextRun({ text: `Generated At: ${new Date().toLocaleString()}`, font: 'Arial', size: 18, color: '7f8c8d' })],
                }),
                // 1. Executive Summary
                new Paragraph({ heading: HeadingLevel.HEADING_2, spacing: { before: 240, after: 120 }, children: [new TextRun({ text: '1. EXECUTIVE SUMMARY', bold: true, font: 'Arial', color: '34495e' })] }),
                new Paragraph({ spacing: { after: 200 }, children: [new TextRun({ text: executiveSummary, font: 'Arial', size: 22 })] }),
                // 2. Project Overview
                new Paragraph({ heading: HeadingLevel.HEADING_2, spacing: { before: 240, after: 120 }, children: [new TextRun({ text: '2. PROJECT OVERVIEW', bold: true, font: 'Arial', color: '34495e' })] }),
                ...overviewParagraphs,
                // 3. Stage-by-Stage Performance
                new Paragraph({ heading: HeadingLevel.HEADING_2, spacing: { before: 240, after: 120 }, children: [new TextRun({ text: '3. STAGE-BY-STAGE PERFORMANCE', bold: true, font: 'Arial', color: '34495e' })] }),
                new Table({
                    width: { size: 100, type: WidthType.PERCENTAGE },
                    rows: stageTableRows,
                }),
                // 4. Delay Reasons
                new Paragraph({ heading: HeadingLevel.HEADING_2, spacing: { before: 240, after: 120 }, children: [new TextRun({ text: '4. DELAY REASONS', bold: true, font: 'Arial', color: '34495e' })] }),
                ...delayParagraphs,
                // 5. Lessons Learned
                new Paragraph({ heading: HeadingLevel.HEADING_2, spacing: { before: 240, after: 120 }, children: [new TextRun({ text: '5. LESSONS LEARNED', bold: true, font: 'Arial', color: '34495e' })] }),
                ...lessonParagraphs,
            ]
        }]
    });
}

// Trigger browser download of a real .docx document using the docx library
async function triggerDocxDownload(docObject, fileName) {
    const blob = await docx.Packer.toBlob(docObject);
    saveAs(blob, fileName);
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
        // Use docx-based download
        dlBtn.href = '#';
        dlBtn.download = `Project_Report_${project.name.replace(/\s+/g, '_')}.docx`;
        dlBtn.onclick = (e) => {
            e.preventDefault();
            downloadProjectReport(currentProjectId, dlBtn);
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

    const newReport = {
        id: Date.now().toString(),
        projectId: currentProjectId,
        currentStageStatus: projectStatus,
        overallProgress: totalProgress,
        executiveSummary: executiveSummary,
        keyDelaysSummary: numDelays > 0 ? `${numDelays} delay(s) recorded` : 'No delays recorded',
        lessonsLearnedSummary: lessons.length > 0 ? `${lessons.length} lesson(s) recorded` : 'No lessons recorded',
        content: '__docx__', // marker indicating this is a docx-format report
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

// ============================================================
// Document Management — Library Rendering & Smart Viewing
// ============================================================

function getExtensionFromName(filename) {
    if (!filename) return '';
    const parts = filename.split('.');
    return parts.length > 1 ? parts.pop().toLowerCase() : '';
}

function getDocTypeClass(type) {
    if (type === 'pdf') return 'doc-pdf';
    if (type === 'docx' || type === 'doc') return 'doc-docx';
    if (type === 'xlsx' || type === 'xls') return 'doc-xlsx';
    return 'doc-other';
}

function getDocIcon(type) {
    if (type === 'pdf') return 'file-text';
    if (type === 'docx' || type === 'doc') return 'file-type';
    if (type === 'xlsx' || type === 'xls') return 'file-spreadsheet';
    return 'file';
}

function getDocTypeLabel(type) {
    if (type === 'pdf') return 'PDF';
    if (type === 'docx' || type === 'doc') return 'Word';
    if (type === 'xlsx' || type === 'xls') return 'Excel';
    return type.toUpperCase();
}

function openDocument(doc) {
    const url = doc.url;
    const type = (doc.type || '').toLowerCase();

    if (type === 'pdf') {
        // PDFs — open directly in a new tab
        window.open(url, '_blank');
    } else if (type === 'docx' || type === 'doc' || type === 'xlsx' || type === 'xls') {
        // Word / Excel — use Google Docs Viewer
        const viewerUrl = `https://docs.google.com/viewer?url=${encodeURIComponent(url)}&embedded=true`;
        window.open(viewerUrl, '_blank');
    } else {
        // Fallback — open in new tab
        window.open(url, '_blank');
    }
}

function renderDocumentLibrary() {
    const container = document.getElementById('documents-container');
    const countBadge = document.getElementById('doc-count-badge');
    if (!container) return;

    const canDelete = currentUser && (currentUser.role === 'Admin' || currentUser.role === 'Project Manager');

    // Fetch fresh documents from backend
    fetch(`${API_BASE}/api/documents`)
        .then(res => res.json())
        .then(data => {
            const documents = data.documents || [];

            // Update local store
            window.dbStore.documents = documents;

            // Update count badge
            if (countBadge) {
                countBadge.textContent = `${documents.length} file${documents.length !== 1 ? 's' : ''}`;
            }

            if (documents.length === 0) {
                container.innerHTML = `
                    <div class="documents-empty">
                        <i data-lucide="folder-open"></i>
                        <p>No documents uploaded yet.</p>
                        <p style="font-size: 0.82rem; margin-top: 0.5rem; opacity: 0.7;">
                            Project Managers and Admins can upload PDF, Word, and Excel files.
                        </p>
                    </div>
                `;
                lucide.createIcons();
                return;
            }

            // Sort newest first
            documents.sort((a, b) => new Date(b.date) - new Date(a.date));

            container.innerHTML = '';

            documents.forEach(doc => {
                const typeClass = getDocTypeClass(doc.type);
                const iconName = getDocIcon(doc.type);
                const typeLabel = getDocTypeLabel(doc.type);
                const uploadDate = doc.date ? new Date(doc.date).toLocaleDateString('en-US', { month: 'short', day: 'numeric', year: 'numeric' }) : 'Unknown date';

                const card = document.createElement('div');
                card.className = `document-card ${typeClass}`;
                card.setAttribute('title', `Click to view: ${doc.name}`);
                card.setAttribute('data-doc-id', doc.id);
                card.innerHTML = `
                    <div class="doc-icon">
                        <i data-lucide="${iconName}"></i>
                    </div>
                    <div class="doc-info">
                        <div class="doc-name">${doc.name}</div>
                        <div class="doc-meta">
                            <span class="doc-type-badge">${typeLabel}</span>
                            <span>•</span>
                            <span>${doc.uploadedBy || 'Unknown'}</span>
                            <span>•</span>
                            <span>${uploadDate}</span>
                        </div>
                    </div>
                    ${canDelete ? `<button class="doc-delete-btn" data-doc-id="${doc.id}" title="Delete document"><i data-lucide="trash-2" style="width: 16px; height: 16px;"></i></button>` : ''}
                    <div class="doc-open-icon">
                        <i data-lucide="external-link" style="width: 18px; height: 18px;"></i>
                    </div>
                `;

                card.addEventListener('click', (e) => {
                    // Don't open document if the delete button was clicked
                    if (e.target.closest('.doc-delete-btn')) return;
                    openDocument(doc);
                });
                container.appendChild(card);
            });

            // Attach delete handlers
            if (canDelete) {
                container.querySelectorAll('.doc-delete-btn').forEach(btn => {
                    btn.addEventListener('click', (e) => {
                        e.stopPropagation();
                        const docId = btn.getAttribute('data-doc-id');
                        const docCard = btn.closest('.document-card');
                        const docName = docCard?.querySelector('.doc-name')?.textContent || 'this document';

                        if (!confirm(`Are you sure you want to delete "${docName}"?`)) return;

                        fetch(`${API_BASE}/api/documents?id=${encodeURIComponent(docId)}`, {
                            method: 'DELETE'
                        })
                            .then(res => res.json())
                            .then(data => {
                                if (data.success) {
                                    // Remove from DOM instantly
                                    if (docCard) {
                                        docCard.style.transition = 'opacity 0.3s ease, transform 0.3s ease';
                                        docCard.style.opacity = '0';
                                        docCard.style.transform = 'scale(0.95)';
                                        setTimeout(() => {
                                            docCard.remove();
                                            // Update local store
                                            window.dbStore.documents = (window.dbStore.documents || []).filter(d => d.id !== docId);
                                            // Update count badge
                                            const remaining = container.querySelectorAll('.document-card').length;
                                            if (countBadge) {
                                                countBadge.textContent = `${remaining} file${remaining !== 1 ? 's' : ''}`;
                                            }
                                            // Show empty state if no documents left
                                            if (remaining === 0) {
                                                container.innerHTML = `
                                                <div class="documents-empty">
                                                    <i data-lucide="folder-open"></i>
                                                    <p>No documents uploaded yet.</p>
                                                    <p style="font-size: 0.82rem; margin-top: 0.5rem; opacity: 0.7;">
                                                        Project Managers and Admins can upload PDF, Word, and Excel files.
                                                    </p>
                                                </div>
                                            `;
                                                lucide.createIcons();
                                            }
                                        }, 300);
                                    }
                                    showToast(`"${docName}" deleted successfully.`, 'success');
                                } else {
                                    showToast('Failed to delete document.', 'error');
                                }
                            })
                            .catch(err => {
                                console.error('[Doc Delete Error]', err);
                                showToast('Failed to delete document.', 'error');
                            });
                    });
                });
            }

            lucide.createIcons();
        })
        .catch(err => {
            console.error('[Document Library Error]', err);
            // Fallback to local store
            const documents = window.dbStore.documents || [];
            if (countBadge) {
                countBadge.textContent = `${documents.length} file${documents.length !== 1 ? 's' : ''}`;
            }
            if (documents.length === 0) {
                container.innerHTML = `
                    <div class="documents-empty">
                        <i data-lucide="folder-open"></i>
                        <p>No documents uploaded yet.</p>
                    </div>
                `;
            }
            lucide.createIcons();
        });
}
