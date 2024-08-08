import os
import shutil
from flask import Flask, request, render_template, send_from_directory
from werkzeug.utils import secure_filename

app = Flask(__name__)

# Configuration
UPLOAD_FOLDER = "webui_test/uploads"
TRANSLATED_FOLDER = "webui_test/translated"
ALLOWED_EXTENSIONS = {"docx", "pptx"}
app.config["UPLOAD_FOLDER"] = UPLOAD_FOLDER
app.config["TRANSLATED_FOLDER"] = TRANSLATED_FOLDER

# Ensure the upload and translated folders exist
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(TRANSLATED_FOLDER, exist_ok=True)


def allowed_file(filename) -> bool:
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXTENSIONS


@app.route("/", methods=["GET", "POST"])
def upload_file() -> str:
    if request.method == "POST":
        if "file" not in request.files:
            return render_template("upload.html", message="No file part")
        file = request.files["file"]
        if file.filename == "":
            return render_template("upload.html", message="No selected file")
        if file and allowed_file(file.filename):
            filename = secure_filename(file.filename)
            file_path = os.path.join(app.config["UPLOAD_FOLDER"], filename)
            file.save(file_path)
            translated_filename = translate_file(file_path)
            if translated_filename:
                download_url = f"/translated/{translated_filename}"
                return render_template(
                    "upload.html",
                    message="File successfully uploaded",
                    download_url=download_url,
                )
            else:
                return render_template(
                    "upload.html",
                    message="Translation failed or no translated file found",
                )
        return render_template("upload.html", message="Invalid file type")
    return render_template("upload.html")


@app.route("/translated/<filename>")
def download_file(filename) -> str:
    directory = os.path.abspath(app.config["TRANSLATED_FOLDER"])
    file_path = os.path.join(directory, filename)

    # Debugging print statements
    print(f"Trying to download file: {filename}")
    print(f"Directory: {directory}")
    print(f"Full file path: {file_path}")

    # Debug: List files in the directory
    print(f"Files in directory: {os.listdir(directory)}")

    if not os.path.isfile(file_path):
        print(f"File not found: {file_path}")
        return "File not found", 404

    return send_from_directory(directory, filename, as_attachment=True)


def translate_file(file_path) -> str:
    _original_filepath = file_path
    _destination_filepath = os.path.join(
        app.config["TRANSLATED_FOLDER"], "translated.pptx"
    )
    shutil.copy(file_path, _destination_filepath)
    print(f"File copied to: {_destination_filepath}")  # Debugging line
    return os.path.basename(_destination_filepath)


if __name__ == "__main__":
    app.run(debug=True)
