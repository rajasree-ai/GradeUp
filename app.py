import os
from flask import (Flask, render_template, request, redirect,
                   url_for, flash, send_file, session)
from flask_login import login_required, current_user
from werkzeug.utils import secure_filename

from analyzer import ResultAnalyzer
from database import init_db, save_upload, get_history
from auth import init_auth

UPLOAD_FOLDER = "uploads"
REPORT_FOLDER = "reports"
ALLOWED_EXT = {"xlsx", "xls", "csv", "pdf"}

app = Flask(__name__)
app.secret_key = "change-this-secret-key-in-production-12345"
app.config["UPLOAD_FOLDER"] = UPLOAD_FOLDER
app.config["MAX_CONTENT_LENGTH"] = 10 * 1024 * 1024  # 10 MB

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(REPORT_FOLDER, exist_ok=True)

init_db()
init_auth(app)


def allowed_file(fn):
    return "." in fn and fn.rsplit(".", 1)[1].lower() in ALLOWED_EXT


@app.route("/")
@login_required
def index():
    return render_template("index.html", history=get_history(current_user.id))


@app.route("/upload", methods=["POST"])
@login_required
def upload():
    if "file" not in request.files:
        flash("No file selected.", "danger")
        return redirect(url_for("index"))

    file = request.files["file"]
    pass_mark = int(request.form.get("pass_mark", 40))
    class_key = request.form.get("class_key", "").strip()

    if file.filename == "" or not allowed_file(file.filename):
        flash("Invalid file. Use .pdf, .xlsx, .xls, or .csv", "danger")
        return redirect(url_for("index"))

    filename = secure_filename(file.filename)
    user_filename = f"user{current_user.id}_{filename}"
    filepath = os.path.join(app.config["UPLOAD_FOLDER"], user_filename)
    file.save(filepath)

    try:
        analyzer = ResultAnalyzer(filepath, pass_mark)
        analyzer.load_data()
        analyzer.analyze()
        stats = analyzer.get_statistics()

        session["filepath"] = filepath
        session["pass_mark"] = pass_mark
        session["class_key"] = class_key
        for key in (
            "college_name", "class_info", "academic_year", "branch",
            "year_semester", "exam_details",
        ):
            session.pop(key, None)

        save_upload(current_user.id, filename, pass_mark, stats)
        return redirect(url_for("dashboard"))

    except Exception as e:
        import traceback
        traceback.print_exc()
        flash(f"Error: {e}", "danger")
        return redirect(url_for("index"))


@app.route("/dashboard")
@login_required
def dashboard():
    filepath = session.get("filepath")
    pass_mark = session.get("pass_mark", 40)
    if not filepath or not os.path.exists(filepath):
        flash("Please upload a file first.", "warning")
        return redirect(url_for("index"))

    analyzer = ResultAnalyzer(filepath, pass_mark)
    analyzer.load_data()
    analyzer.analyze()
    stats = analyzer.get_statistics()
    return render_template("dashboard.html", stats=stats)


@app.route("/failures")
@login_required
def failures():
    filepath = session.get("filepath")
    pass_mark = session.get("pass_mark", 40)
    if not filepath:
        return redirect(url_for("index"))

    analyzer = ResultAnalyzer(filepath, pass_mark)
    analyzer.load_data()
    analyzer.analyze()

    student_wise = analyzer.get_student_wise_failures().to_dict(orient="records")
    subject_wise = analyzer.get_subject_wise_failures()
    return render_template("failures.html",
                           student_wise=student_wise,
                           subject_wise=subject_wise)


@app.route("/round-lists")
@login_required
def round_lists():
    filepath = session.get("filepath")
    pass_mark = session.get("pass_mark", 40)
    if not filepath:
        return redirect(url_for("index"))

    analyzer = ResultAnalyzer(filepath, pass_mark)
    analyzer.load_data()
    analyzer.analyze()
    lists = analyzer.generate_round_lists()
    return render_template("round_lists.html", round_lists=lists)


@app.route("/download-report")
@login_required
def download_report():
    filepath = session.get("filepath")
    pass_mark = session.get("pass_mark", 40)
    if not filepath:
        return redirect(url_for("index"))

    analyzer = ResultAnalyzer(filepath, pass_mark)
    analyzer.load_data()
    analyzer.analyze()
    out_path = analyzer.export_report(REPORT_FOLDER)
    return send_file(out_path, as_attachment=True)


@app.route("/download-pdf")
@login_required
def download_pdf():
    filepath = session.get("filepath")
    pass_mark = session.get("pass_mark", 40)
    if not filepath:
        return redirect(url_for("index"))

    analyzer = ResultAnalyzer(filepath, pass_mark)
    analyzer.load_data()
    analyzer.analyze()
    out_path = analyzer.export_pdf(
        REPORT_FOLDER,
        college_name=session.get("college_name", ""),
        class_info=session.get("class_info", ""),
    )
    return send_file(out_path, as_attachment=True,
                     download_name=os.path.basename(out_path))


@app.route("/download-student-correction-pdf")
@login_required
def download_student_correction_pdf():
    filepath = session.get("filepath")
    pass_mark = session.get("pass_mark", 40)
    if not filepath:
        return redirect(url_for("index"))

    analyzer = ResultAnalyzer(filepath, pass_mark)
    analyzer.load_data()
    analyzer.analyze()
    out_path = analyzer.export_student_correction_pdf(
        REPORT_FOLDER,
        college_name=session.get("college_name", ""),
        class_info=session.get("class_info", ""),
    )
    return send_file(out_path, as_attachment=True,
                     download_name=os.path.basename(out_path))


@app.route("/download-counselling-report")
@login_required
def download_counselling_report():
    filepath = session.get("filepath")
    pass_mark = session.get("pass_mark", 40)
    if not filepath:
        return redirect(url_for("index"))

    analyzer = ResultAnalyzer(filepath, pass_mark)
    analyzer.load_data()
    analyzer.analyze()
    out_path = analyzer.export_counselling_pdf(
        REPORT_FOLDER,
        department=current_user.department,
        class_key=session.get("class_key", ""),
    )
    return send_file(out_path, as_attachment=True,
                     download_name=os.path.basename(out_path))


if __name__ == "__main__":
    app.run(debug=True)