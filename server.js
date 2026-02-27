const http = require('http');
const fs = require('fs');
const path = require('path');
const url = require('url');

// ============================================================
// Antigravity Cloud — GridCo PME Backend Server
// ============================================================

const DB_FILE = path.join(__dirname, 'database.json');
const PORT = process.env.PORT || 3000;

// Initialize DB file if not exists
if (!fs.existsSync(DB_FILE)) {
    fs.writeFileSync(DB_FILE, JSON.stringify({
        users: [],
        projects: [],
        stages: [],
        notifications: [],
        delayRecords: [],
        lessonsLearned: [],
        projectReports: []
    }, null, 2));
}

// ============================================================
// Concurrency-safe DB helpers (write queue prevents race conditions)
// ============================================================
let writeQueue = Promise.resolve();

function readDB() {
    try {
        return JSON.parse(fs.readFileSync(DB_FILE, 'utf8'));
    } catch (e) {
        console.error('[DB READ ERROR]', e.message);
        return null;
    }
}

function writeDB(dbData) {
    return new Promise((resolve, reject) => {
        writeQueue = writeQueue.then(() => {
            try {
                fs.writeFileSync(DB_FILE, JSON.stringify(dbData, null, 2));
                resolve(true);
            } catch (e) {
                console.error('[DB WRITE ERROR]', e.message);
                reject(e);
            }
        });
    });
}

// ============================================================
// Business logic helpers
// ============================================================

function isProjectCompleted(projectId, stages) {
    const projectStages = stages.filter(s => s.projectId === projectId);
    if (projectStages.length === 0) return false;
    return projectStages.every(s => parseInt(s.progress) === 100);
}

function evaluateProjectCompletionStatus(dbData) {
    const projects = dbData.projects || [];
    const stages = dbData.stages || [];
    let updated = false;

    projects.forEach(project => {
        const completed = isProjectCompleted(project.id, stages);

        if (completed && !project.completionDate) {
            project.completionDate = new Date().toISOString();
            updated = true;
        } else if (!completed && project.completionDate) {
            project.completionDate = null;
            updated = true;
        }
    });

    if (updated) {
        dbData.projects = projects;
    }

    return updated;
}

// ============================================================
// Response helpers
// ============================================================

function sendJSON(res, statusCode, data) {
    const body = JSON.stringify(data);
    res.writeHead(statusCode, {
        'Content-Type': 'application/json',
        'Content-Length': Buffer.byteLength(body)
    });
    res.end(body);
}

function setCORS(res) {
    res.setHeader('Access-Control-Allow-Origin', '*');
    res.setHeader('Access-Control-Allow-Methods', 'GET, POST, PUT, DELETE, OPTIONS');
    res.setHeader('Access-Control-Allow-Headers', 'Content-Type, Authorization');
}

function logRequest(method, pathname, statusCode) {
    const timestamp = new Date().toISOString();
    console.log(`[${timestamp}] ${method} ${pathname} → ${statusCode}`);
}

// ============================================================
// HTTP Server
// ============================================================

const server = http.createServer(async (req, res) => {
    setCORS(res);
    const parsedUrl = url.parse(req.url, true);
    const pathname = parsedUrl.pathname;
    const query = parsedUrl.query;

    if (pathname === '/') {
        const htmlPath = path.join(__dirname, 'index.html'); // Make sure this matches your file name
        fs.readFile(htmlPath, (err, data) => {
            if (err) {
                res.writeHead(500);
                res.end('Error loading index.html');
                return;
            }
            res.writeHead(200, { 'Content-Type': 'text/html' });
            res.end(data);
        });
        return;
    }

    // Preflight
    if (req.method === 'OPTIONS') {
        res.writeHead(204);
        res.end();
        return;
    }



    try {
        // ==============================================================
        // GET /health — Health check endpoint
        // ==============================================================
        if (pathname === '/health' && req.method === 'GET') {
            logRequest(req.method, pathname, 200);
            sendJSON(res, 200, {
                status: 'ok',
                service: 'GridCo PME Backend — Antigravity Cloud',
                uptime: process.uptime(),
                timestamp: new Date().toISOString()
            });

            // ==============================================================
            // GET /api/data — Return full database
            // ==============================================================
        } else if (pathname === '/api/data' && req.method === 'GET') {
            const dbData = readDB();
            if (!dbData) {
                logRequest(req.method, pathname, 500);
                sendJSON(res, 500, { error: 'Failed to read database' });
                return;
            }
            logRequest(req.method, pathname, 200);
            sendJSON(res, 200, dbData);

            // ==============================================================
            // POST /api/sync — Sync a collection
            // Auto-evaluates project completion when stages change
            // ==============================================================
        } else if (pathname === '/api/sync' && req.method === 'POST') {
            let body = '';
            req.on('data', chunk => body += chunk.toString());
            req.on('end', async () => {
                try {
                    const { key, data } = JSON.parse(body);

                    if (!key || !Array.isArray(data)) {
                        logRequest(req.method, pathname, 400);
                        sendJSON(res, 400, { error: 'Invalid payload: key (string) and data (array) required' });
                        return;
                    }

                    const dbData = readDB();
                    if (!dbData) {
                        logRequest(req.method, pathname, 500);
                        sendJSON(res, 500, { error: 'Failed to read database' });
                        return;
                    }

                    dbData[key] = data;

                    // Auto-evaluate project completion whenever stages are synced
                    if (key === 'stages') {
                        evaluateProjectCompletionStatus(dbData);
                    }

                    await writeDB(dbData);
                    logRequest(req.method, pathname, 200);
                    sendJSON(res, 200, { success: true });
                } catch (e) {
                    console.error('[SYNC ERROR]', e.message);
                    logRequest(req.method, pathname, 500);
                    sendJSON(res, 500, { error: 'Sync failed: ' + e.message });
                }
            });

            // ==============================================================
            // GET /api/projects/completed — Fetch completed projects
            // Optional query params: ?month=MM&year=YYYY
            // ==============================================================
        } else if (pathname === '/api/projects/completed' && req.method === 'GET') {
            const dbData = readDB();
            if (!dbData) {
                logRequest(req.method, pathname, 500);
                sendJSON(res, 500, { error: 'Failed to read database' });
                return;
            }

            // Re-evaluate to ensure fresh status
            evaluateProjectCompletionStatus(dbData);
            await writeDB(dbData);

            const stages = dbData.stages || [];
            let completedProjects = (dbData.projects || []).filter(p => {
                return isProjectCompleted(p.id, stages) && p.completionDate;
            });

            // Apply month/year filtering on completionDate
            const filterMonth = query.month ? parseInt(query.month) : null;
            const filterYear = query.year ? parseInt(query.year) : null;

            if (filterMonth !== null || filterYear !== null) {
                completedProjects = completedProjects.filter(p => {
                    const completionDate = new Date(p.completionDate);
                    const matchMonth = filterMonth !== null
                        ? (completionDate.getMonth() + 1) === filterMonth
                        : true;
                    const matchYear = filterYear !== null
                        ? completionDate.getFullYear() === filterYear
                        : true;
                    return matchMonth && matchYear;
                });
            }

            // Enrich with computed progress
            const enriched = completedProjects.map(p => {
                const pStages = stages.filter(s => s.projectId === p.id);
                const avgProgress = pStages.length
                    ? Math.round(pStages.reduce((sum, s) => sum + parseInt(s.progress), 0) / pStages.length)
                    : 0;
                return {
                    ...p,
                    status: 'Completed',
                    totalProgress: avgProgress,
                    stageCount: pStages.length
                };
            });

            logRequest(req.method, pathname, 200);
            sendJSON(res, 200, {
                count: enriched.length,
                filters: { month: filterMonth, year: filterYear },
                projects: enriched
            });

            // ==============================================================
            // GET /api/projects/in-progress — Fetch in-progress projects
            // ==============================================================
        } else if (pathname === '/api/projects/in-progress' && req.method === 'GET') {
            const dbData = readDB();
            if (!dbData) {
                logRequest(req.method, pathname, 500);
                sendJSON(res, 500, { error: 'Failed to read database' });
                return;
            }

            evaluateProjectCompletionStatus(dbData);
            await writeDB(dbData);

            const stages = dbData.stages || [];
            const inProgressProjects = (dbData.projects || []).filter(p => {
                return !isProjectCompleted(p.id, stages);
            });

            const enriched = inProgressProjects.map(p => {
                const pStages = stages.filter(s => s.projectId === p.id);
                const avgProgress = pStages.length
                    ? Math.round(pStages.reduce((sum, s) => sum + parseInt(s.progress), 0) / pStages.length)
                    : 0;
                return {
                    ...p,
                    status: 'In Progress',
                    totalProgress: avgProgress,
                    stageCount: pStages.length
                };
            });

            logRequest(req.method, pathname, 200);
            sendJSON(res, 200, {
                count: enriched.length,
                projects: enriched
            });

            // ==============================================================
            // GET /api/projects/evaluate — Trigger re-evaluation
            // ==============================================================
        } else if (pathname === '/api/projects/evaluate' && req.method === 'GET') {
            const dbData = readDB();
            if (!dbData) {
                logRequest(req.method, pathname, 500);
                sendJSON(res, 500, { error: 'Failed to read database' });
                return;
            }

            const updated = evaluateProjectCompletionStatus(dbData);
            if (updated) {
                await writeDB(dbData);
            }

            const stages = dbData.stages || [];
            const summary = (dbData.projects || []).map(p => ({
                id: p.id,
                name: p.name,
                completed: isProjectCompleted(p.id, stages),
                completionDate: p.completionDate || null
            }));

            logRequest(req.method, pathname, 200);
            sendJSON(res, 200, {
                evaluated: true,
                updated,
                projects: summary
            });

            // ==============================================================
            // 404 — Not found
            // ==============================================================
        } else {
            logRequest(req.method, pathname, 404);
            sendJSON(res, 404, { error: 'Not found', path: pathname });
        }
    } catch (err) {
        console.error('[SERVER ERROR]', err);
        logRequest(req.method, pathname, 500);
        sendJSON(res, 500, { error: 'Internal server error' });
    }
});


// ============================================================
// Start server
// ============================================================
server.listen(PORT, () => {
    console.log('');
    console.log('╔══════════════════════════════════════════════════════╗');
    console.log('║   GridCo PME Backend — Antigravity Cloud            ║');
    console.log('╠══════════════════════════════════════════════════════╣');
    console.log(`║   Status  : RUNNING                                 ║`);
    console.log(`║   Port    : ${String(PORT).padEnd(41)}║`);
    console.log(`║   URL     : http://localhost:${String(PORT).padEnd(25)}║`);
    console.log(`║   Health  : http://localhost:${String(PORT).padEnd(1)}/health${' '.repeat(Math.max(0, 17 - String(PORT).length))}║`);
    console.log('║   DB File : database.json                           ║');
    console.log('╚══════════════════════════════════════════════════════╝');
    console.log('');
    console.log('Endpoints:');
    console.log('  GET  /health                    → Health check');
    console.log('  GET  /api/data                  → Full database');
    console.log('  POST /api/sync                  → Sync collection');
    console.log('  GET  /api/projects/completed    → Completed projects');
    console.log('  GET  /api/projects/in-progress  → In-progress projects');
    console.log('  GET  /api/projects/evaluate     → Re-evaluate statuses');
    console.log('');
    console.log('Waiting for requests...');
    console.log('');
});

// Graceful shutdown
process.on('SIGTERM', () => {
    console.log('\n[SHUTDOWN] Received SIGTERM — closing server...');
    server.close(() => {
        console.log('[SHUTDOWN] Server closed.');
        process.exit(0);
    });
});

process.on('SIGINT', () => {
    console.log('\n[SHUTDOWN] Received SIGINT — closing server...');
    server.close(() => {
        console.log('[SHUTDOWN] Server closed.');
        process.exit(0);
    });
});
