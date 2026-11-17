"""Run the app locally without Docker.

    python run.py              -> http://localhost:8031  (API + web interface)
    python run.py --port 9000
    python run.py --reload     -> auto-restart when code changes
"""
import argparse
import os

import uvicorn

if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--port", type=int, default=int(os.getenv("APP_PORT", 8034)))
    parser.add_argument("--host", default="127.0.0.1")
    parser.add_argument("--reload", action="store_true")
    args = parser.parse_args()
    print(f"Open http://localhost:{args.port}")
    uvicorn.run("app.main:app", host=args.host, port=args.port, reload=args.reload)
