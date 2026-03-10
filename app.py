# -*- coding: utf-8 -*-
"""Slide Creator — Flask server."""

import os, json, time, glob
from flask import Flask, render_template, request, jsonify, send_file
from werkzeug.utils import secure_filename

import template_slick
import template_colorful
import template_bold
import template_editorial
import template_noir

app = Flask(__name__)

OUTPUT_DIR = os.path.join(os.path.dirname(__file__), "output")
PROJECTS_DIR = os.path.join(os.path.dirname(__file__), "projects")
os.makedirs(OUTPUT_DIR, exist_ok=True)
os.makedirs(PROJECTS_DIR, exist_ok=True)

TEMPLATES = {
    "slick": template_slick,
    "colorful": template_colorful,
    "bold": template_bold,
    "editorial": template_editorial,
    "noir": template_noir,
}


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/build", methods=["POST"])
def build():
    data = request.json
    design = data.get("designSystem", "slick")
    slides = data.get("slides", [])

    tpl = TEMPLATES.get(design, template_slick)
    slide_configs = [(s[0], s[1]) for s in slides]

    fname = f"deck_{int(time.time())}.pptx"
    output_path = os.path.join(OUTPUT_DIR, fname)
    tpl.build_deck(slide_configs, output_path)

    deck_title = data.get("deckTitle", "presentation")
    safe_title = secure_filename(deck_title) or "presentation"
    return send_file(output_path, as_attachment=True,
                     download_name=f"{safe_title}.pptx")


@app.route("/save", methods=["POST"])
def save_project():
    data = request.json
    name = secure_filename(data.get("name", "untitled")) or "untitled"
    path = os.path.join(PROJECTS_DIR, f"{name}.json")
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data.get("project", {}), f, indent=2)
    return jsonify({"success": True, "name": name})


@app.route("/load/<name>")
def load_project(name):
    path = os.path.join(PROJECTS_DIR, f"{secure_filename(name)}.json")
    if not os.path.exists(path):
        return jsonify({"error": "Not found"}), 404
    with open(path, encoding="utf-8") as f:
        return jsonify(json.load(f))


@app.route("/projects")
def list_projects():
    files = glob.glob(os.path.join(PROJECTS_DIR, "*.json"))
    names = [os.path.splitext(os.path.basename(f))[0] for f in sorted(files)]
    return jsonify(names)


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5001))
    print(f"Slide Creator running at http://localhost:{port}")
    app.run(debug=True, port=port)
