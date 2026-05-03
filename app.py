"""
app.py
------
Entry point. Creates the Flask app and registers routes.
Run this file to start the server.
"""

import os
from flask import Flask, render_template
from routes import bp


def create_app():
    app = Flask(__name__)
    app.config["MAX_CONTENT_LENGTH"] = 50 * 1024 * 1024   # 50 MB max upload

    os.makedirs("outputs", exist_ok=True)
    os.makedirs("uploads", exist_ok=True)

    app.register_blueprint(bp)

    # ── Friendly error pages ──────────────────────────────────────────────────
    @app.errorhandler(413)
    def file_too_large(e):
        return render_template("index.html",
            error="⚠️ File is too large. Please upload a file smaller than 50 MB."), 413

    @app.errorhandler(500)
    def server_error(e):
        return render_template("index.html",
            error="Something went wrong on the server. Please try again."), 500

    return app


if __name__ == "__main__":
    app = create_app()
    print("\n🌞  Energybae Solar Calculator")
    print("   Open in browser → http://localhost:5000\n")
    app.run(debug=True, port=5000)