const http = require('http');
const fs = require('fs');
const path = require('path');
const url = require('url');
const { MongoClient } = require('mongodb');

// ============================================================
// Antigravity Cloud — GridCo PME Backend Server (MongoDB)
// ============================================================

const PORT = process.env.PORT || 3000;
const MONGO_URI = process.env.MONGO_URI || '';
const DB_NAME = process.env.DB_NAME || 'pme_nexus';
const GEMINI_API_KEY = process.env.GEMINI_API_KEY || '';

// All collection names (must match the keys in the old database.json structure)
const COLLECTIONS = [
    'users', 'projects', 'stages', 'notifications',
    'delayRecords', 'lessonsLearned', 'projectReports', 'documents'
];

// ============================================================
// MongoDB Connection
// ============================================================
let db = null;
let mongoClient = null;

async function connectToDB() {
    if (!MONGO_URI) {
        console.error('[MONGO] ❌ MONGO_URI environment variable is not set!');
        console.error('[MONGO] Set it in your Render dashboard or .env file.');
        console.error('[MONGO] Example: mongodb+srv://user:pass@cluster.mongodb.net/?retryWrites=true&w=majority');
        process.exit(1);
    }

    try {
        mongoClient = new MongoClient(MONGO_URI);
        await mongoClient.connect();

        db = mongoClient.db(DB_NAME);

        // Verify connection
        await db.command({ ping: 1 });
        console.log(`[MONGO] ✅ Connected to MongoDB Atlas — database: "${DB_NAME}"`);

        return db;
    } catch (err) {
        console.error('[MONGO] ❌ Connection failed:', err.message);
        process.exit(1);
    }
}

// ============================================================
// DB Helpers — Same interface as the old JSON file system
// readDB()  → returns { users: [], projects: [], ... }
// writeDB() → persists the full object back to MongoDB
// ============================================================

async function readDB() {
    try {
        const result = {};

        // Read all collections in parallel
        const reads = COLLECTIONS.map(async (col) => {
            const docs = await db.collection(col).find({}).toArray();
            // Strip MongoDB's internal _id field so the data looks identical
            // to the old database.json format the frontend expects
            result[col] = docs.map(doc => {
                const { _id, ...rest } = doc;
                return rest;
            });
        });

        await Promise.all(reads);
        return result;
    } catch (e) {
        console.error('[DB READ ERROR]', e.message);
        return null;
    }
}

async function writeDB(dbData) {
    try {
        // Write each collection that exists in the data object
        const writes = COLLECTIONS.map(async (col) => {
            if (!Array.isArray(dbData[col])) return;

            const collection = db.collection(col);

            // Bulk replace: clear existing, insert new
            await collection.deleteMany({});

            if (dbData[col].length > 0) {
                await collection.insertMany(dbData[col]);
            }
        });

        await Promise.all(writes);
        return true;
    } catch (e) {
        console.error('[DB WRITE ERROR]', e.message);
        throw e;
    }
}

// ============================================================
// Single-collection helpers (more efficient for targeted ops)
// ============================================================

async function readCollection(collectionName) {
    try {
        const docs = await db.collection(collectionName).find({}).toArray();
        return docs.map(doc => {
            const { _id, ...rest } = doc;
            return rest;
        });
    } catch (e) {
        console.error(`[DB READ ${collectionName} ERROR]`, e.message);
        return [];
    }
}

async function writeCollection(collectionName, data) {
    try {
        const collection = db.collection(collectionName);
        await collection.deleteMany({});
        if (data.length > 0) {
            await collection.insertMany(data);
        }
        return true;
    } catch (e) {
        console.error(`[DB WRITE ${collectionName} ERROR]`, e.message);
        throw e;
    }
}

// ============================================================
// Business logic helpers
// ============================================================

// --- Gemini AI Helper ---
async function callGeminiAPI(reportData) {
    const https = require('https');

    const systemPrompt = `You are a senior Project Monitoring and Evaluation (PM&E) specialist. 
You will receive raw project data including: project name, description, overall status, progress percentage, 
stage-by-stage performance, delay records, and lessons learned.

Your task is to produce a polished, professional Executive Summary suitable for a formal PM&E report. 
The summary must:
- Open with the project title and current status in a formal tone.
- Provide a concise performance overview referencing the overall progress percentage.
- Highlight any stages that are behind schedule, citing specific stage names and planned end dates.
- Summarize delay root causes and their impacts if delay records exist.
- Reference key lessons learned and recommendations if they exist.
- Close with a forward-looking statement on risk mitigation or next steps.
- Be between 150 and 350 words.
- Use third-person, formal language appropriate for stakeholder distribution.
- Do NOT use markdown formatting, bullet points, or headers — return flowing paragraph text only.`;

    const userContent = JSON.stringify(reportData, null, 2);

    const requestBody = JSON.stringify({
        contents: [{
            parts: [
                { text: systemPrompt },
                { text: `Here is the raw project data to enhance:\n\n${userContent}` }
            ]
        }],
        generationConfig: {
            temperature: 0.7,
            maxOutputTokens: 1024
        }
    });

    const apiUrl = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:generateContent?key=${GEMINI_API_KEY}`;

    return new Promise((resolve, reject) => {
        const parsedUrl = new URL(apiUrl);

        const options = {
            hostname: parsedUrl.hostname,
            path: parsedUrl.pathname + parsedUrl.search,
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
                'Content-Length': Buffer.byteLength(requestBody)
            }
        };

        const req = https.request(options, (res) => {
            let data = '';
            res.on('data', chunk => data += chunk);
            res.on('end', () => {
                try {
                    const parsed = JSON.parse(data);

                    if (parsed.error) {
                        reject(new Error(parsed.error.message || 'Gemini API error'));
                        return;
                    }

                    const text = parsed?.candidates?.[0]?.content?.parts?.[0]?.text;
                    if (!text) {
                        reject(new Error('No text returned from Gemini API'));
                        return;
                    }

                    resolve(text.trim());
                } catch (e) {
                    reject(new Error('Failed to parse Gemini response: ' + e.message));
                }
            });
        });

        req.on('error', (e) => reject(new Error('Gemini request failed: ' + e.message)));
        req.write(requestBody);
        req.end();
    });
}

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
// Server-side delay detection & notification generation
// Scans all stages, marks overdue ones as "Behind Schedule",
// and creates per-user notifications for Admins & PMs.
// Deduplicates by stageId to avoid spamming.
// ============================================================
function generateDelayNotifications(dbData) {
    const stages = dbData.stages || [];
    const projects = dbData.projects || [];
    const users = dbData.users || [];
    const notifications = dbData.notifications || [];

    const todayStr = new Date().toISOString().split('T')[0];
    const targetUsers = users.filter(u => u.role === 'Admin' || u.role === 'Project Manager');

    let stagesUpdated = false;
    let notifsCreated = 0;

    stages.forEach(stage => {
        const progress = parseInt(stage.progress) || 0;

        // --- Step 1: Evaluate & update stage status ---
        let correctStatus = 'On Track';
        if (progress === 100) {
            correctStatus = 'Completed';
        } else if (todayStr > stage.plannedEnd) {
            correctStatus = 'Behind Schedule';
        }

        if (stage.status !== correctStatus) {
            stage.status = correctStatus;
            stagesUpdated = true;
        }

        // --- Step 2: Generate notifications for behind-schedule stages ---
        if (correctStatus === 'Behind Schedule') {
            const project = projects.find(p => p.id === stage.projectId);
            const projectName = project ? project.name : 'Unknown Project';

            targetUsers.forEach(user => {
                // Deduplicate: check if this user already has a notification for this stage
                const alreadyExists = notifications.find(
                    n => n.stageId === stage.id &&
                        n.userId === user.id &&
                        n.message.includes('behind schedule')
                );

                if (!alreadyExists) {
                    notifications.push({
                        id: Date.now().toString() + '-' + stage.id + '-' + user.id,
                        userId: user.id,
                        projectId: stage.projectId,
                        stageId: stage.id,
                        message: `Stage "${stage.name}" in project "${projectName}" is behind schedule.`,
                        read: false,
                        createdAt: new Date().toISOString()
                    });
                    notifsCreated++;
                }
            });
        }
    });

    // Write changes back into the dbData object (caller handles persistence)
    if (stagesUpdated) {
        dbData.stages = stages;
    }
    if (notifsCreated > 0) {
        dbData.notifications = notifications;
    }

    return { stagesUpdated, notifsCreated };
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

    if (pathname === '/health') {
        res.writeHead(200, { 'Content-Type': 'text/plain' });
        res.end('Server is awake');
        return;
    }
    // ============================================================
    // Serve Static Files (CSS, JS, Images)
    // ============================================================
    if (pathname.match(/\.(css|js|png|jpg|jpeg|gif|svg)$/)) {
        const filePath = path.join(__dirname, pathname);
        const extname = path.extname(filePath);

        // Decide the correct Content-Type based on the file extension
        let contentType = 'text/plain';
        switch (extname) {
            case '.css': contentType = 'text/css'; break;
            case '.js': contentType = 'application/javascript'; break;
            case '.png': contentType = 'image/png'; break;
            case '.jpg': case '.jpeg': contentType = 'image/jpeg'; break;
            case '.svg': contentType = 'image/svg+xml'; break;
        }

        fs.readFile(filePath, (err, data) => {
            if (err) {
                res.writeHead(404);
                res.end('File not found');
                return;
            }
            res.writeHead(200, { 'Content-Type': contentType });
            res.end(data);
        });
        return; // Important: stops the server from running your API code below
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
                service: 'GridCo PME Backend — MongoDB Atlas',
                uptime: process.uptime(),
                timestamp: new Date().toISOString()
            });

            // ==============================================================
            // GET /api/data — Return full database
            // ==============================================================
        } else if (pathname === '/api/data' && req.method === 'GET') {
            const dbData = await readDB();
            if (!dbData) {
                logRequest(req.method, pathname, 500);
                sendJSON(res, 500, { error: 'Failed to read database' });
                return;
            }

            // --- Evaluate delays & generate notifications before responding ---
            const delayResult = generateDelayNotifications(dbData);
            evaluateProjectCompletionStatus(dbData);

            // Persist if anything changed
            if (delayResult.stagesUpdated || delayResult.notifsCreated) {
                await writeDB(dbData);
                if (delayResult.notifsCreated > 0) {
                    console.log(`[DELAY CHECK] Generated ${delayResult.notifsCreated} new delay notification(s)`);
                }
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

                    // Write only the synced collection for efficiency
                    await writeCollection(key, data);

                    // Auto-evaluate project completion whenever stages are synced
                    if (key === 'stages') {
                        const dbData = await readDB();
                        if (dbData) {
                            const updated = evaluateProjectCompletionStatus(dbData);
                            if (updated) {
                                await writeCollection('projects', dbData.projects);
                            }
                        }
                    }

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
            const dbData = await readDB();
            if (!dbData) {
                logRequest(req.method, pathname, 500);
                sendJSON(res, 500, { error: 'Failed to read database' });
                return;
            }

            // Re-evaluate to ensure fresh status
            const updated = evaluateProjectCompletionStatus(dbData);
            if (updated) {
                await writeCollection('projects', dbData.projects);
            }

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
            const dbData = await readDB();
            if (!dbData) {
                logRequest(req.method, pathname, 500);
                sendJSON(res, 500, { error: 'Failed to read database' });
                return;
            }

            const updated = evaluateProjectCompletionStatus(dbData);
            if (updated) {
                await writeCollection('projects', dbData.projects);
            }

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
            const dbData = await readDB();
            if (!dbData) {
                logRequest(req.method, pathname, 500);
                sendJSON(res, 500, { error: 'Failed to read database' });
                return;
            }

            const updated = evaluateProjectCompletionStatus(dbData);
            if (updated) {
                await writeCollection('projects', dbData.projects);
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
            // GET /api/documents — Retrieve all documents
            // ==============================================================
        } else if (pathname === '/api/documents' && req.method === 'GET') {
            const documents = await readCollection('documents');
            logRequest(req.method, pathname, 200);
            sendJSON(res, 200, { documents });

            // ==============================================================
            // POST /api/documents — Save a new document record
            // ==============================================================
        } else if (pathname === '/api/documents' && req.method === 'POST') {
            let body = '';
            req.on('data', chunk => body += chunk.toString());
            req.on('end', async () => {
                try {
                    const { name, url: fileUrl, type, uploadedBy, date } = JSON.parse(body);

                    if (!name || !fileUrl) {
                        logRequest(req.method, pathname, 400);
                        sendJSON(res, 400, { error: 'Missing required fields: name, url' });
                        return;
                    }

                    const newDoc = {
                        id: Date.now().toString(),
                        name: name,
                        url: fileUrl,
                        type: type || 'unknown',
                        uploadedBy: uploadedBy || 'Unknown',
                        date: date || new Date().toISOString()
                    };

                    await db.collection('documents').insertOne(newDoc);

                    logRequest(req.method, pathname, 201);
                    sendJSON(res, 201, { success: true, document: newDoc });
                } catch (e) {
                    console.error('[DOC SAVE ERROR]', e.message);
                    logRequest(req.method, pathname, 500);
                    sendJSON(res, 500, { error: 'Failed to save document: ' + e.message });
                }
            });

            // ==============================================================
            // DELETE /api/documents — Remove a document by ID
            // ==============================================================
        } else if (pathname === '/api/documents' && req.method === 'DELETE') {
            const docId = query.id;

            if (!docId) {
                logRequest(req.method, pathname, 400);
                sendJSON(res, 400, { error: 'Missing required query parameter: id' });
                return;
            }

            const result = await db.collection('documents').deleteOne({ id: docId });

            if (result.deletedCount === 0) {
                logRequest(req.method, pathname, 404);
                sendJSON(res, 404, { error: 'Document not found', id: docId });
                return;
            }

            logRequest(req.method, pathname, 200);
            sendJSON(res, 200, { success: true, message: 'Document deleted', id: docId });

            // ==============================================================
            // POST /api/enhance-report — AI-enhanced report via Gemini
            // ==============================================================
        } else if (pathname === '/api/enhance-report' && req.method === 'POST') {
            let body = '';
            req.on('data', chunk => body += chunk.toString());
            req.on('end', async () => {
                try {
                    const reportData = JSON.parse(body);

                    if (!GEMINI_API_KEY) {
                        logRequest(req.method, pathname, 500);
                        sendJSON(res, 500, { success: false, error: 'GEMINI_API_KEY is not configured on the server.' });
                        return;
                    }

                    const aiText = await callGeminiAPI(reportData);

                    logRequest(req.method, pathname, 200);
                    sendJSON(res, 200, { success: true, aiText });
                } catch (e) {
                    console.error('[AI ENHANCE ERROR]', e.message);
                    logRequest(req.method, pathname, 500);
                    sendJSON(res, 500, { success: false, error: 'AI enhancement failed: ' + e.message });
                }
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
// Start server (connects to MongoDB first, then listens)
// ============================================================
async function startServer() {
    // 1. Connect to MongoDB
    await connectToDB();

    // 2. Start HTTP server
    server.listen(PORT, () => {
        console.log('');
        console.log('╔══════════════════════════════════════════════════════╗');
        console.log('║   GridCo PME Backend — MongoDB Atlas                ║');
        console.log('╠══════════════════════════════════════════════════════╣');
        console.log(`║   Status  : RUNNING                                 ║`);
        console.log(`║   Port    : ${String(PORT).padEnd(41)}║`);
        console.log(`║   URL     : http://localhost:${String(PORT).padEnd(25)}║`);
        console.log(`║   Health  : http://localhost:${String(PORT).padEnd(1)}/health${' '.repeat(Math.max(0, 17 - String(PORT).length))}║`);
        console.log('║   DB      : MongoDB Atlas                           ║');
        console.log('╚══════════════════════════════════════════════════════╝');
        console.log('');
        console.log('Endpoints:');
        console.log('  GET  /health                    → Health check');
        console.log('  GET  /api/data                  → Full database');
        console.log('  POST /api/sync                  → Sync collection');
        console.log('  GET  /api/projects/completed    → Completed projects');
        console.log('  GET  /api/projects/in-progress  → In-progress projects');
        console.log('  GET  /api/projects/evaluate     → Re-evaluate statuses');
        console.log('  GET  /api/documents             → List documents');
        console.log('  POST /api/documents             → Save document');
        console.log('  DELETE /api/documents?id=...     → Delete document');
        console.log('  POST /api/enhance-report        → AI-enhanced report (Gemini)');
        console.log('');
        console.log('Waiting for requests...');
        console.log('');
    });
}

startServer().catch(err => {
    console.error('[STARTUP FATAL]', err);
    process.exit(1);
});

// Graceful shutdown — close MongoDB connection
process.on('SIGTERM', () => {
    console.log('\n[SHUTDOWN] Received SIGTERM — closing server...');
    server.close(async () => {
        if (mongoClient) await mongoClient.close();
        console.log('[SHUTDOWN] Server and MongoDB connection closed.');
        process.exit(0);
    });
});

process.on('SIGINT', () => {
    console.log('\n[SHUTDOWN] Received SIGINT — closing server...');
    server.close(async () => {
        if (mongoClient) await mongoClient.close();
        console.log('[SHUTDOWN] Server and MongoDB connection closed.');
        process.exit(0);
    });
});

