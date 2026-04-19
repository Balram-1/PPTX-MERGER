import os
import uuid
from flask import Flask, request, render_template, send_file, redirect
import pythoncom
import win32com.client

UPLOAD_FOLDER = "uploads"
OUTPUT_FOLDER = "output"

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

app = Flask(__name__)
app.config["UPLOAD_FOLDER"] = UPLOAD_FOLDER

def merge_pptx(file_paths):
    # Initialize PowerPoint via COM
    pythoncom.CoInitialize() # Needed for multi-threading in Flask
    try:
        powerpoint = win32com.client.Dispatch("PowerPoint.Application")
    except Exception as e:
        print(f"Failed to start PowerPoint via COM. Is Microsoft Office installed? Error: {e}")
        return None
        
    output_file = os.path.abspath(os.path.join(OUTPUT_FOLDER, f"merged_{uuid.uuid4().hex}.pptx"))
    
    try:
        # Open the first presentation as the base
        base_prs_path = os.path.abspath(file_paths[0])
        prs = powerpoint.Presentations.Open(base_prs_path, WithWindow=False)
        
        # Insert slides from remaining files
        for path in file_paths[1:]:
            abs_path = os.path.abspath(path)
            insert_index = prs.Slides.Count
            prs.Slides.InsertFromFile(abs_path, insert_index)
            
        prs.SaveAs(output_file)
        prs.Close()
    except Exception as e:
        print(f"Error merging files: {e}")
        output_file = None
    finally:
        pythoncom.CoUninitialize()

    return output_file


@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        files = request.files.getlist("files")

        if not files or files[0].filename == "":
            return redirect("/")

        file_paths = []

        for f in files:
            if not f.filename.lower().endswith(".pptx"):
                continue

            unique_name = f"{uuid.uuid4().hex}_{f.filename}"
            path = os.path.join(UPLOAD_FOLDER, unique_name)
            f.save(path)
            file_paths.append(path)

        if not file_paths:
            return "No valid PPTX files uploaded"

        merged_file = merge_pptx(file_paths)

        # cleanup uploads after merge
        for path in file_paths:
            try:
                os.remove(path)
            except:
                pass
                
        if not merged_file:
            return "Failed to merge presentations. Ensure PowerPoint is installed and running correctly.", 500

        return send_file(merged_file, as_attachment=True)

    return render_template("index.html")


if __name__ == "__main__":
    app.run(debug=True)