# RL-sync (RocketLane Project Viewer)

Simple Flask application that calls RocketLane API and displays 3 fields in a web table:
- Project Name
- Close Date
- CSM Handover Date

## Setup

1. Clone / navigate to directory:
   ```bash
   cd ~/rl-sync
   ```
2. Create virtualenv and install deps:
   ```bash
   python3 -m venv venv
   source venv/bin/activate
   pip install -r requirements.txt
   ```
3. Copy .env example:
   ```bash
   cp .env.example .env
   ```
4. Set environment variables in `.env`:
   - `ROCKETLANE_API_KEY`
   - `ROCKETLANE_DOMAIN` (e.g. `yourcompany.rocketlane.com`)
   - optionally `ROCKETLANE_API_BASE_URL` if custom

5. Run app:
   ```bash
   flask run --host=0.0.0.0 --port=5000
   ```

   or

   ```bash
   python app.py
   ```

6. Browse: `http://localhost:5000`

## RocketLane API mapping

This app assumes response structure has either `data` or `projects` root container.
Adapt `fetch_projects()` in `app.py` for your actual RocketLane API response schema.

## Troubleshooting

- API auth errors: check key and domain values.
- Fields not found: modify field mapping in `fetch_projects()`.
