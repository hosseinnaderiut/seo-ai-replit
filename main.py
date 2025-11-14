from flask import Flask, render_template, request, send_file
import os
from werkzeug.utils import secure_filename
from aicode import run_seo

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = "uploads"
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
os.makedirs("SEO_Output", exist_ok=True)

@app.route("/", methods=["GET","POST"])
def index():
    if request.method == "POST":
        # دریافت API Key
        api_key = request.form.get("api_key","").strip()

        # دریافت فایل اکسل
        file = request.files.get("excel_file")
        if not file:
            return render_template("index.html", error="فایل اکسل انتخاب نشده!")

        filename = secure_filename(file.filename)
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(filepath)

        # اجرا کردن aicode.py
        output_file, merged_count, category_count, mode = run_seo(filepath, api_key)

        return render_template("index.html",
                               success=True,
                               output_file=output_file,
                               merged_count=merged_count,
                               category_count=category_count,
                               mode=mode)

    return render_template("index.html")

@app.route("/download/<path:filename>")
def download_file(filename):
    # دانلود فایل خروجی
    return send_file(filename, as_attachment=True)

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=8080)
