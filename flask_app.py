from datetime import datetime, timezone
from pathlib import Path
import uuid

from flask import Flask, request, render_template, jsonify, send_from_directory, abort
from flask_sqlalchemy import SQLAlchemy
from flask_dropzone import Dropzone
from werkzeug.utils import secure_filename


# this is NOT a pdf converter. It is A WAY to work with pdf:s in a business setting.

app = Flask(__name__)
app.config.from_object("config")

BASE_DIR = Path(__file__).resolve().parent
UPLOADS_DIR = BASE_DIR / "uploads"
STATIC_CLIENT_DIR = BASE_DIR / "static" / "client"
UPLOADS_DIR.mkdir(exist_ok=True, parents=True)
STATIC_CLIENT_DIR.mkdir(exist_ok=True, parents=True)

app.config.update(
    UPLOADED_PATH=str(UPLOADS_DIR),
    CLIENT_PDF2=str(STATIC_CLIENT_DIR),
    DROPZONE_MAX_FILE_SIZE=30,
    DROPZONE_MAX_FILES=500,
    DROPZONE_UPLOAD_ON_CLICK=True,
    DROPZONE_ALLOWED_FILE_CUSTOM=True,
    DROPZONE_ALLOWED_FILE_TYPE=".pdf,.msg,.eml",
)

db = SQLAlchemy(app)
dropzone = Dropzone(app)


class Job(db.Model):
    __tablename__ = "jobs"

    id = db.Column(db.Integer, primary_key=True)
    slug = db.Column(db.String(64), nullable=False, unique=True)
    state = db.Column(db.String(20), nullable=False, default="queued")
    result = db.Column(db.Integer, nullable=False, default=0)


def read_filename_from_txt(txt_path: Path) -> str | None:
    if not txt_path.is_file():
        return None

    lines = txt_path.read_text(encoding="utf-8").splitlines()
    if not lines:
        return None

    last_line = lines[-1].strip()
    if not last_line:
        return None

    return Path(last_line).name


@app.route("/get-image/<image_name>")
def get_image(image_name: str):
    txt_path = BASE_DIR / "excel_File_and_Path.txt"
    filename = read_filename_from_txt(txt_path)

    if not filename:
        abort(404)

    file_path = STATIC_CLIENT_DIR / filename
    if not file_path.is_file():
        abort(404)

    return send_from_directory(
        directory=app.config["CLIENT_PDF2"],
        path=filename,
        as_attachment=True,
    )


@app.route("/get-pdf/<image_name>")
def get_pdf(image_name: str):
    txt_path = BASE_DIR / "summary_File_and_Path.txt"
    filename = read_filename_from_txt(txt_path)

    if not filename:
        abort(404)

    file_path = STATIC_CLIENT_DIR / filename
    if not file_path.is_file():
        abort(404)

    return send_from_directory(
        directory=app.config["CLIENT_PDF2"],
        path=filename,
        as_attachment=True,
    )


@app.route("/", methods=["GET", "POST"])
def upload():
    if request.method == "POST":
        saved_any = False

        for key, uploaded_file in request.files.items():
            if key.startswith("file") and uploaded_file.filename:
                safe_name = secure_filename(uploaded_file.filename)
                if safe_name:
                    uploaded_file.save(UPLOADS_DIR / safe_name)
                    saved_any = True
                    print(f"Uploaded: {safe_name}")

        if not saved_any:
            return jsonify({"error": "No valid files uploaded"}), 400

        job_id = uuid.uuid4().hex
        job = Job(slug=job_id, state="queued", result=0) 
        db.session.add(job)
        db.session.commit()

        return jsonify({"job_id": job_id})

    return render_template("main.html")


@app.route("/completed/<job_id>")
def completed(job_id: str):
    return render_template("complete.html", job_id=job_id)


@app.route("/query", methods=["POST"])
def query():
    job_id = request.form.get("id", "").strip()
    if not job_id:
        return jsonify({"error": "missing job id"}), 400

    data = Job.query.filter_by(slug=job_id).first()
    if data is None:
        return jsonify({"error": "job not found"}), 404

    return jsonify(
        {
            "state": data.state,
            "result": data.result,
        }
    )


with app.app_context():
    db.create_all()


if __name__ == "__main__":
    app.run(debug=True)