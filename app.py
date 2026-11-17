import sys
import traceback

try:
    from webapp.app import app
except Exception:
    traceback.print_exc(file=sys.stderr)
    from flask import Flask, jsonify
    app = Flask(__name__)

    @app.route("/", defaults={"path": ""})
    @app.route("/<path:path>")
    def catch_all(path):
        err = traceback.format_exc()
        return jsonify({
            "error": "Initialization error on cloud server",
            "details": err
        }), 500

if __name__ == "__main__":
    app.run()
