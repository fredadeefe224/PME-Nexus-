import json
import os
from http.server import BaseHTTPRequestHandler, HTTPServer
from urllib.parse import urlparse, parse_qs
from datetime import datetime
from typing import Any

DB_FILE = os.path.join(os.path.dirname(__file__), 'database.json')

if not os.path.exists(DB_FILE):
    with open(DB_FILE, 'w') as f:
        json.dump({
            "users": [], "projects": [], "stages": [], "notifications": [],
            "delayRecords": [], "lessonsLearned": [], "projectReports": []
        }, f, indent=2)


# --- Helper: Read database from file ---
def read_db():
    try:
        with open(DB_FILE, 'r') as f:
            return json.load(f)
    except Exception as e:
        print(f"Failed to read DB: {e}")
        return None


# --- Helper: Write database to file ---
def write_db(db_data):
    try:
        with open(DB_FILE, 'w') as f:
            json.dump(db_data, f, indent=2)
        return True
    except Exception as e:
        print(f"Failed to write DB: {e}")
        return False


# --- Helper: Determine if a project is completed (all stages at 100%) ---
def is_project_completed(project_id, stages):
    project_stages = [s for s in stages if s.get('projectId') == project_id]
    if len(project_stages) == 0:
        return False  # No stages = not completed
    return all(int(s.get('progress', 0)) == 100 for s in project_stages)


# --- Helper: Evaluate and update completionDate for all projects ---
def evaluate_project_completion_status(db_data):
    projects = db_data.get('projects', [])
    stages = db_data.get('stages', [])
    updated = False

    for project in projects:
        completed = is_project_completed(project['id'], stages)

        if completed and not project.get('completionDate'):
            # Project just became completed — set completionDate to now
            project['completionDate'] = datetime.utcnow().isoformat() + 'Z'
            updated = True
        elif not completed and project.get('completionDate'):
            # Project was completed but now a stage went below 100% — clear completionDate
            project['completionDate'] = None
            updated = True

    if updated:
        db_data['projects'] = projects

    return updated


class RequestHandler(BaseHTTPRequestHandler):
    def _send_json(self, status_code, data):
        self.send_response(status_code)
        self.send_header('Access-Control-Allow-Origin', '*')
        self.send_header('Content-Type', 'application/json')
        self.end_headers()
        self.wfile.write(json.dumps(data).encode('utf-8'))

    def do_OPTIONS(self):
        self.send_response(204)
        self.send_header('Access-Control-Allow-Origin', '*')
        self.send_header('Access-Control-Allow-Methods', 'GET, POST, OPTIONS')
        self.send_header('Access-Control-Allow-Headers', 'Content-Type')
        self.end_headers()

    def do_GET(self):
        parsed = urlparse(self.path)
        pathname = parsed.path
        query = parse_qs(parsed.query)

        # ============================================================
        # GET /api/data — Return full database (existing endpoint)
        # ============================================================
        if pathname == '/api/data':
            try:
                with open(DB_FILE, 'r') as f:
                    data = f.read()
                self.send_response(200)
                self.send_header('Access-Control-Allow-Origin', '*')
                self.send_header('Content-Type', 'application/json')
                self.end_headers()
                self.wfile.write(data.encode('utf-8'))
            except Exception as e:
                self._send_json(500, {'error': 'Failed to read db'})

        # ============================================================
        # GET /api/projects/completed — Fetch all completed projects
        # Optional query params: ?month=MM&year=YYYY
        # ============================================================
        elif pathname == '/api/projects/completed':
            db_data = read_db()
            if not db_data:
                self._send_json(500, {'error': 'Failed to read db'})
                return

            # Re-evaluate to ensure fresh status
            evaluate_project_completion_status(db_data)
            write_db(db_data)

            stages = db_data.get('stages', [])
            completed_projects = [
                p for p in db_data.get('projects', [])
                if is_project_completed(p['id'], stages) and p.get('completionDate')
            ]

            # Apply month/year filtering on completionDate if provided
            filter_month = int(query['month'][0]) if 'month' in query else None
            filter_year = int(query['year'][0]) if 'year' in query else None

            if filter_month is not None or filter_year is not None:
                filtered = []
                for p in completed_projects:
                    try:
                        cd = datetime.fromisoformat(p['completionDate'].replace('Z', '+00:00'))
                        match_month = (cd.month == filter_month) if filter_month is not None else True
                        match_year = (cd.year == filter_year) if filter_year is not None else True
                        if match_month and match_year:
                            filtered.append(p)
                    except (ValueError, KeyError):
                        pass
                completed_projects = filtered

            # Enrich each project with computed progress info
            enriched: list[dict[str, Any]] = []
            for p in completed_projects:
                p_stages = [s for s in stages if s.get('projectId') == p['id']]
                avg_progress = (
                    round(sum(int(s.get('progress', 0)) for s in p_stages) / len(p_stages))
                    if p_stages else 0
                )
                enriched.append({
                    **p,
                    'status': 'Completed',
                    'totalProgress': avg_progress,
                    'stageCount': len(p_stages)
                })

            self._send_json(200, {
                'count': len(enriched),
                'filters': {
                    'month': filter_month,
                    'year': filter_year
                },
                'projects': enriched
            })

        # ============================================================
        # GET /api/projects/in-progress — Fetch all in-progress projects
        # ============================================================
        elif pathname == '/api/projects/in-progress':
            db_data = read_db()
            if not db_data:
                self._send_json(500, {'error': 'Failed to read db'})
                return

            # Re-evaluate to ensure fresh status
            evaluate_project_completion_status(db_data)
            write_db(db_data)

            stages = db_data.get('stages', [])
            in_progress_projects = [
                p for p in db_data.get('projects', [])
                if not is_project_completed(p['id'], stages)
            ]

            # Enrich each project with computed progress info
            enriched: list[dict[str, Any]] = []
            for p in in_progress_projects:
                p_stages = [s for s in stages if s.get('projectId') == p['id']]
                avg_progress = (
                    round(sum(int(s.get('progress', 0)) for s in p_stages) / len(p_stages))
                    if p_stages else 0
                )
                enriched.append({
                    **p,
                    'status': 'In Progress',
                    'totalProgress': avg_progress,
                    'stageCount': len(p_stages)
                })

            self._send_json(200, {
                'count': len(enriched),
                'projects': enriched
            })

        # ============================================================
        # GET /api/projects/evaluate — Trigger re-evaluation of all
        # project completion statuses and persist the results
        # ============================================================
        elif pathname == '/api/projects/evaluate':
            db_data = read_db()
            if not db_data:
                self._send_json(500, {'error': 'Failed to read db'})
                return

            updated = evaluate_project_completion_status(db_data)
            if updated:
                write_db(db_data)

            stages = db_data.get('stages', [])
            summary = [{
                'id': p['id'],
                'name': p.get('name', ''),
                'completed': is_project_completed(p['id'], stages),
                'completionDate': p.get('completionDate')
            } for p in db_data.get('projects', [])]

            self._send_json(200, {
                'evaluated': True,
                'updated': updated,
                'projects': summary
            })

        # ============================================================
        # 404 — Not found
        # ============================================================
        else:
            self.send_response(404)
            self.end_headers()
            self.wfile.write(b'Not found')

    def do_POST(self):
        parsed = urlparse(self.path)
        pathname = parsed.path

        # ============================================================
        # POST /api/sync — Sync a collection (existing endpoint)
        # Now also auto-evaluates project completion when stages change
        # ============================================================
        if pathname == '/api/sync':
            content_length = int(self.headers.get('Content-Length', 0))
            post_data = self.rfile.read(content_length)
            try:
                payload = json.loads(post_data.decode('utf-8'))
                key = payload.get('key')
                data = payload.get('data')

                db_data = read_db()
                if not db_data:
                    self._send_json(500, {'error': 'Failed to read db'})
                    return

                db_data[key] = data

                # Auto-evaluate project completion whenever stages are synced
                if key == 'stages':
                    evaluate_project_completion_status(db_data)

                if not write_db(db_data):
                    self._send_json(500, {'error': 'Failed to save db'})
                    return

                self._send_json(200, {'success': True})
            except Exception as e:
                self._send_json(500, {'error': str(e)})
        else:
            self.send_response(404)
            self.end_headers()
            self.wfile.write(b'Not found')


def run(server_class=HTTPServer, handler_class=RequestHandler, port=3000):
    server_address = ('', port)
    httpd = server_class(server_address, handler_class)
    print(f'Antigravity Cloud Backend simulating on http://localhost:{port} (Python)')
    httpd.serve_forever()

if __name__ == '__main__':
    run()

