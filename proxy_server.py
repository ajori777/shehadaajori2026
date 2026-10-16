import hmac
import json
import os
import secrets
from pathlib import Path

from flask import Flask, jsonify, request
import firebase_admin
from firebase_admin import credentials, firestore


app = Flask(__name__)
db = None


def initialize_firebase():
    global db
    try:
        firebase_admin.get_app()
    except ValueError:
        pass
    else:
        db = firestore.client()
        return

    service_account_json = os.environ.get("FIREBASE_SERVICE_ACCOUNT_JSON")
    service_account_path = os.environ.get("FIREBASE_KEY_PATH")
    if service_account_json:
        credential = credentials.Certificate(json.loads(service_account_json))
    elif service_account_path:
        credential = credentials.Certificate(service_account_path)
    else:
        app.logger.error("Firebase credentials are not configured.")
        return

    firebase_admin.initialize_app(credential)
    db = firestore.client()


initialize_firebase()


@app.after_request
def add_cors_headers(response):
    response.headers["Access-Control-Allow-Origin"] = "*"
    response.headers["Access-Control-Allow-Methods"] = "GET, POST, OPTIONS"
    response.headers["Access-Control-Allow-Headers"] = "Content-Type"
    return response


def firebase_ready():
    if db is None:
        return jsonify({"status": "error", "msg": "Firebase is unavailable"}), 503
    return None


def admin_authorized():
    expected = os.environ.get("ADMIN_PASSWORD", "")
    supplied = request.args.get("pw", "")
    return bool(expected) and hmac.compare_digest(supplied, expected)


@app.route("/")
def home():
    return "License Proxy Server is Running."


@app.route("/admin_login")
def admin_page():
    if not admin_authorized():
        return "<h1>Access denied</h1>", 403
    panel_path = Path(__file__).resolve().with_name("admin_panel.html")
    try:
        return panel_path.read_text(encoding="utf-8")
    except OSError:
        app.logger.exception("Unable to load admin panel.")
        return "Admin panel is unavailable.", 500


@app.route("/check_license", methods=["GET"])
def check_license():
    unavailable = firebase_ready()
    if unavailable:
        return unavailable

    code = request.args.get("code", "").strip()
    machine_id = request.args.get("m_id", "").strip()
    if not code or not machine_id:
        return jsonify({"status": "error", "msg": "code and m_id are required"}), 400

    license_ref = db.collection("licenses").document(code)
    transaction = db.transaction()

    @firestore.transactional
    def read_and_bind(transaction):
        snapshot = license_ref.get(transaction=transaction)
        if not snapshot.exists:
            return {"status": "error", "msg": "invalid"}
        data = snapshot.to_dict() or {}
        if data.get("Status", "active") != "active":
            return {"status": "error", "msg": "blocked"}

        bound_machine = str(data.get("MachineID") or "")
        if bound_machine and bound_machine != machine_id:
            return {"status": "error", "msg": "mismatch"}
        if not bound_machine:
            transaction.update(license_ref, {"MachineID": machine_id})

        return {
            "status": "success",
            "max": data.get("Max", 0),
            "used": data.get("Used", 0),
            "user_name": data.get("Name", "User"),
            "license_status": data.get("Status", "active"),
        }

    try:
        result = read_and_bind(transaction)
        return jsonify(result)
    except Exception:
        app.logger.exception("License lookup failed.")
        return jsonify({"status": "error", "msg": "License lookup failed"}), 500


@app.route("/consume_usage", methods=["POST"])
def consume_usage():
    unavailable = firebase_ready()
    if unavailable:
        return unavailable

    code = request.args.get("code", "").strip()
    machine_id = request.args.get("m_id", "").strip()
    payload = request.get_json(silent=True) or {}
    try:
        count = int(payload.get("count", 0))
    except (TypeError, ValueError):
        count = 0
    if not code or not machine_id or count <= 0:
        return jsonify({"status": "error", "msg": "invalid_request"}), 400

    license_ref = db.collection("licenses").document(code)
    transaction = db.transaction()

    @firestore.transactional
    def consume(transaction):
        snapshot = license_ref.get(transaction=transaction)
        if not snapshot.exists:
            return {"status": "error", "msg": "invalid"}

        data = snapshot.to_dict() or {}
        if data.get("Status", "active") != "active":
            return {"status": "error", "msg": "blocked"}
        if str(data.get("MachineID") or "") != machine_id:
            return {"status": "error", "msg": "mismatch"}

        try:
            maximum = int(data.get("Max", 0))
            used = int(data.get("Used", 0))
        except (TypeError, ValueError):
            return {"status": "error", "msg": "invalid_balance"}
        if maximum < 0 or used < 0 or used + count > maximum:
            return {"status": "error", "msg": "insufficient_balance"}

        updated_used = used + count
        transaction.update(license_ref, {"Used": updated_used})
        return {
            "status": "success",
            "max": maximum,
            "used": updated_used,
            "license_status": "active",
        }

    try:
        result = consume(transaction)
        status_code = 200 if result.get("status") == "success" else 409
        return jsonify(result), status_code
    except Exception:
        app.logger.exception("Usage update failed.")
        return jsonify({"status": "error", "msg": "Usage update failed"}), 500


@app.route("/update_usage", methods=["POST"])
def reject_legacy_usage_update():
    return jsonify({
        "status": "error",
        "msg": "Use the validated /consume_usage endpoint.",
    }), 410


@app.route("/request_registration", methods=["POST"])
def request_registration():
    unavailable = firebase_ready()
    if unavailable:
        return unavailable
    payload = request.get_json(silent=True) or {}
    name = str(payload.get("name") or "").strip()
    machine_id = str(payload.get("m_id") or "").strip()
    if not name or not machine_id:
        return jsonify({"status": "error", "msg": "name and m_id are required"}), 400

    try:
        existing = (
            db.collection("licenses")
            .where("MachineID", "==", machine_id)
            .limit(1)
            .get()
        )
        if existing:
            return jsonify({"status": "exists", "code": existing[0].id})
        db.collection("registration_requests").document(machine_id).set({
            "Name": name,
            "Status": "pending",
            "Timestamp": firestore.SERVER_TIMESTAMP,
        })
        return jsonify({"status": "success"})
    except Exception:
        app.logger.exception("Registration request failed.")
        return jsonify({"status": "error", "msg": "Registration request failed"}), 500


@app.route("/check_request_status", methods=["GET"])
def check_request_status():
    unavailable = firebase_ready()
    if unavailable:
        return unavailable
    machine_id = request.args.get("m_id", "").strip()
    if not machine_id:
        return jsonify({"status": "error", "msg": "m_id is required"}), 400
    try:
        snapshot = db.collection("registration_requests").document(machine_id).get()
        data = snapshot.to_dict() or {}
        if data.get("Status") == "approved":
            return jsonify({"status": "approved", "code": data.get("GeneratedCode")})
        if data.get("Status") == "rejected":
            return jsonify({"status": "rejected"})
        return jsonify({"status": "pending"})
    except Exception:
        app.logger.exception("Registration status lookup failed.")
        return jsonify({"status": "error"}), 500


@app.route("/admin/list_requests", methods=["GET"])
def admin_list_requests():
    unavailable = firebase_ready()
    if not admin_authorized() or unavailable:
        return "Forbidden", 403
    try:
        docs = (
            db.collection("registration_requests")
            .where("Status", "==", "pending")
            .stream()
        )
        requests_data = []
        for document in docs:
            item = document.to_dict() or {}
            item["m_id"] = document.id
            requests_data.append(item)
        return jsonify(requests_data)
    except Exception:
        app.logger.exception("Unable to list registration requests.")
        return jsonify({"status": "error"}), 500


@app.route("/admin/handle_request", methods=["POST"])
def admin_handle_request():
    unavailable = firebase_ready()
    if not admin_authorized() or unavailable:
        return "Forbidden", 403
    payload = request.get_json(silent=True) or {}
    machine_id = str(payload.get("m_id") or "").strip()
    action = payload.get("action")
    if not machine_id or action not in {"approve", "reject"}:
        return jsonify({"status": "error", "msg": "invalid_request"}), 400

    request_ref = db.collection("registration_requests").document(machine_id)
    try:
        snapshot = request_ref.get()
        if not snapshot.exists:
            return jsonify({"status": "error", "msg": "request_not_found"}), 404
        if action == "reject":
            request_ref.update({"Status": "rejected"})
            return jsonify({"status": "success"})

        request_data = snapshot.to_dict() or {}
        try:
            max_runs = int(payload.get("max_runs", 5))
        except (TypeError, ValueError):
            return jsonify({"status": "error", "msg": "invalid_max_runs"}), 400
        if max_runs <= 0:
            return jsonify({"status": "error", "msg": "invalid_max_runs"}), 400

        for _ in range(10):
            code = str(secrets.randbelow(9000) + 1000)
            license_ref = db.collection("licenses").document(code)
            if license_ref.get().exists:
                continue
            license_ref.set({
                "Name": request_data.get("Name", ""),
                "Max": max_runs,
                "Used": 0,
                "Status": "active",
                "MachineID": machine_id,
            })
            request_ref.update({"Status": "approved", "GeneratedCode": code})
            return jsonify({"status": "success", "code": code})
        return jsonify({"status": "error", "msg": "unable_to_generate_code"}), 500
    except Exception:
        app.logger.exception("Unable to handle registration request.")
        return jsonify({"status": "error", "msg": "request_update_failed"}), 500


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
