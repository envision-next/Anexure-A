"""
Annexure A generator - local web app.

Flow:
  1. Type a person's name in the search bar.
  2. See every organization that person is linked to.
  3. Toggle (cancel) any rows you don't want.
  4. Click "Generate Word" -> downloads a filled Annexure A .docx.

Data source: the first .xlsx found in this folder. If none is present yet,
the app falls back to built-in SAMPLE data (from the original template) so you
can try the whole flow immediately.

Run:  python app.py   (or double-click run_annexure.bat)
Then open http://127.0.0.1:5000 in your browser.
"""

import io
import os
import re
import datetime

from flask import Flask, request, jsonify, send_file, render_template

from docx_fill import build_document
from data_source import load_people, describe_source

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

app = Flask(__name__)
# Re-read templates from disk on every request so UI edits show on refresh.
app.config["TEMPLATES_AUTO_RELOAD"] = True
app.jinja_env.auto_reload = True

# People are loaded once at startup and cached. Call /api/reload to re-read the
# Excel after you change it, without restarting the server.
_PEOPLE = load_people(BASE_DIR)


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/api/source")
def api_source():
    """Tell the UI where the data is coming from (real Excel vs sample)."""
    return jsonify(describe_source(BASE_DIR))


@app.route("/api/reload", methods=["POST"])
def api_reload():
    global _PEOPLE
    _PEOPLE = load_people(BASE_DIR)
    return jsonify({"ok": True, "count": len(_PEOPLE), **describe_source(BASE_DIR)})


@app.route("/api/search")
def api_search():
    """Return people whose name contains the query (case-insensitive)."""
    q = (request.args.get("q") or "").strip().lower()
    if len(q) < 2:
        return jsonify({"count": 0, "shown": 0, "results": [],
                        "message": "Type at least 2 letters of an organisation name."})

    CAP = 50
    results = []
    total = 0
    words = q.split()
    for idx, person in enumerate(_PEOPLE):
        # Every typed word must match the start of a word in the name,
        # so "na" finds "Nareshkumar" but not "Krishna".
        name_words = re.split(r"[^a-z0-9]+", person["name"].lower())
        if all(any(nw.startswith(w) for nw in name_words) for w in words):
            total += 1
            if len(results) < CAP:
                results.append({
                    "id": idx,
                    "name": person["name"],
                    "din": person["din"],
                    "orgs": person["orgs"],
                })
    return jsonify({"count": total, "shown": len(results), "results": results,
                    "capped": total > len(results)})


@app.route("/api/generate", methods=["POST"])
def api_generate():
    """Fill the template for one person + the org rows they kept."""
    payload = request.get_json(force=True)
    person_id = payload.get("person_id")
    keep = payload.get("keep_indexes")  # list of org indexes to include
    din_override = payload.get("din")   # user-entered DIN/DPIN (optional)
    custom_name = (payload.get("custom_name") or "").strip()
    custom_orgs = payload.get("orgs")  # merged multi-name selection sends its rows directly

    if custom_name:
        # Either a name not found in the data ("No" Annexure with NA tables),
        # or a merged selection of several name spellings with explicit rows.
        person = {"name": custom_name, "din": "", "orgs": list(custom_orgs or [])}
    elif person_id is None or person_id < 0 or person_id >= len(_PEOPLE):
        return jsonify({"error": "Unknown person."}), 400
    else:
        # Copy so an edited DIN never mutates the cached record.
        person = dict(_PEOPLE[person_id])
    if din_override is not None:
        person["din"] = str(din_override).strip()
    all_orgs = person["orgs"]
    if keep is None:
        chosen = all_orgs
    else:
        keep = set(keep)
        chosen = [o for i, o in enumerate(all_orgs) if i in keep]

    # A person with organizations must keep at least one row.
    # A person with NO organizations is allowed -> the doc shows an "NA" row.
    if all_orgs and not chosen:
        return jsonify({"error": "No organizations selected."}), 400

    today = datetime.date.today().strftime("%d/%m/%Y")
    doc_bytes = build_document(BASE_DIR, person, chosen, today)

    safe_name = "".join(c for c in person["name"] if c.isalnum() or c in " -_").strip()
    filename = f"{safe_name} - Annexure A.docx"

    return send_file(
        io.BytesIO(doc_bytes),
        as_attachment=True,
        download_name=filename,
        mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )


if __name__ == "__main__":
    src = describe_source(BASE_DIR)
    print("=" * 60)
    print(" Annexure A generator")
    print(" Data source:", src["label"])
    print(" People loaded:", len(_PEOPLE))
    print(" Open:  http://127.0.0.1:5000")
    print("=" * 60)
    app.run(host="127.0.0.1", port=5000, debug=False)
