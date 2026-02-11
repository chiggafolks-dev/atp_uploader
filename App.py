# app.py
import re
from flask import Flask, render_template, request, jsonify, send_file
import os
import json
from werkzeug.utils import secure_filename
from atp_photo_insert import ATPPhotoInserter
from atp_text_insert import ATPTextReplacer
import tempfile
from atp_docx_insert import ATPDocxInserter

app = Flask(__name__)
app.config["UPLOAD_FOLDER"] = "uploads"
app.config["MAX_CONTENT_LENGTH"] = 50 * 1024 * 1024
IMAGE_EXTENSIONS = {"png", "jpg", "jpeg", "gif"}
TEMPLATE_EXTENSIONS = {"xlsx", "xls", "docx"}


def allowed_template(filename):
    return "." in filename and filename.rsplit(".", 1)[1].lower() in TEMPLATE_EXTENSIONS


def allowed_image(filename):
    return "." in filename and filename.rsplit(".", 1)[1].lower() in IMAGE_EXTENSIONS



@app.route("/")
def index():
    return render_template("atp_photo_upload.html")


# app.py (updated analyze_template endpoint)
@app.route("/analyze_template", methods=["POST"])
def analyze_template():
    """
    API endpoint to analyze uploaded template (Excel or DOCX).
    """
    if "excel_file" not in request.files:
        return jsonify({"error": "No file uploaded"}), 400

    template_file = request.files["excel_file"]  # Rename for clarity
    filename = template_file.filename.lower()

    if template_file.filename == "":
        return jsonify({"error": "No file selected"}), 400

    if not allowed_template(template_file.filename):
        return (
            jsonify({"error": "Invalid template file. Use .xlsx, .xls, or .docx"}),
            400,
        )

    temp_dir = tempfile.mkdtemp()
    template_path = os.path.join(temp_dir, secure_filename(template_file.filename))
    template_file.save(template_path)

    try:
        # Determine file type and use appropriate processor
        if filename.endswith((".xlsx", ".xls")):
            # Excel template
            inserter = ATPPhotoInserter(template_path)
            photo_slots = inserter.get_available_photo_slots()
            # Get text fields if available
            if hasattr(inserter, "get_available_text_fields"):
                text_fields = inserter.get_available_text_fields()
            else:
                text_fields = []

        elif filename.endswith(".docx"):
            # Word template
            inserter = ATPDocxInserter(template_path)
            photo_slots = inserter.get_available_photo_slots()
            text_fields = inserter.get_available_text_fields()

        else:
            return jsonify({"error": "Unsupported file type"}), 400

        # Store processor type for later use
        file_type = "excel" if filename.endswith((".xlsx", ".xls")) else "docx"

        # Clean up
        os.remove(template_path)
        os.rmdir(temp_dir)

        return jsonify(
            {
                "success": True,
                "template_name": template_file.filename,
                "file_type": file_type,
                "photo_slots": photo_slots,
                "text_fields": text_fields,
                "slots_count": len(photo_slots),
                "text_fields_count": len(text_fields),
            }
        )

    except Exception as e:
        # Clean up on error
        if "template_path" in locals() and os.path.exists(template_path):
            os.remove(template_path)
        if "temp_dir" in locals() and os.path.exists(temp_dir):
            os.rmdir(temp_dir)

        return jsonify({"error": str(e)}), 500


# app.py (updated upload_photos endpoint)
# app.py - Fix the standardize_key function

# ... (other imports and code remain the same) ...


def standardize_key(key):
    """
    Standardize placeholder keys to a consistent format.
    Example: 'siteid' -> 'site_id', 'SITEID' -> 'site_id'
    """
    # Convert to lowercase first
    key = key.lower()

    # Handle common variations
    variations = {
        "siteid": "site_id",
        "sitename": "site_name",
        "scopeofwork": "scope_of_work",
        "devicetype": "device_type",
        "projectcode": "project_code",
    }

    return variations.get(key, key)


@app.route("/upload_photos", methods=["POST"])
def upload_photos():
    """
    Main processing endpoint for both Excel and DOCX templates.
    """
    try:
        template_file = request.files["excel_file"]  # Now can be Excel or DOCX
        filename = template_file.filename.lower()
        raw_project_code = request.form.get("project_code", "UNKNOWN").strip()

        # Sanitize project code
        project_code = re.sub(r'[<>:"/\\|?*]', "", raw_project_code)
        if not project_code:
            project_code = "ATP_Project"

        temp_dir = tempfile.mkdtemp()
        template_path = os.path.join(temp_dir, secure_filename(template_file.filename))
        template_file.save(template_path)

        # Collect text values
        text_values = {}
        possible_text_keys = [
            "site_id",
            "siteid",
            "site_name",
            "sitename",
            "hostname",
            "scope_of_work",
            "scopeofwork",
            "device_type",
            "devicetype",
            "project_code",
            "projectcode",
            "siteid1",
            "site_id1",
            "sk1",
            "sk_1",
        ]

        for key in possible_text_keys:
            form_value = request.form.get(key)
            if form_value:
                standardized_key = standardize_key(key)
                text_values[standardized_key] = form_value

        # Also collect any text_ prefixed fields
        for key in request.form:
            if key.startswith("text_"):
                text_key = key.replace("text_", "")
                standardized_key = standardize_key(text_key)
                text_values[standardized_key] = request.form.get(key)

        # Determine file type and use appropriate processor
        if filename.endswith((".xlsx", ".xls")):
            # Process Excel template
            inserter = ATPPhotoInserter(template_path)

            # Get photo mappings
            photo_mappings = json.loads(request.form.get("photo_mappings", "[]"))

            # Insert photos
            for mapping in photo_mappings:
                if (
                    "field_name" not in mapping
                    or mapping["field_name"] not in request.files
                ):
                    continue

                photo_file = request.files[mapping["field_name"]]
                if photo_file.filename == "" or not allowed_image(photo_file.filename):
                    continue

                photo_path = os.path.join(
                    temp_dir, secure_filename(photo_file.filename)
                )
                photo_file.save(photo_path)

                # Use appropriate insertion method
                if "slot_index" in mapping:
                    # Old method by placeholder
                    inserter.insert_photo_by_placeholder(
                        mapping["photo_type"],
                        photo_path,
                        resize_width=300,
                        resize_height=200,
                    )
                elif "target_cell" in mapping:
                    # New method by cell reference
                    inserter.insert_photo_by_cell(
                        mapping["sheet"],
                        mapping["target_cell"],
                        photo_path,
                        resize_width=300,
                        resize_height=200,
                    )

            # Replace text
            text_replacer = ATPTextReplacer(inserter.wb)
            text_replacer.replace(text_values)

            # Generate output filename
            template_basename = os.path.splitext(
                secure_filename(template_file.filename)
            )[0]
            output_filename = f"{project_code}_ATP_Photos_{template_basename}.xlsx"

        elif filename.endswith(".docx"):
            # Process DOCX template
            inserter = ATPDocxInserter(template_path)

            # Get photo mappings
            photo_mappings = json.loads(request.form.get("photo_mappings", "[]"))

            # Insert photos
            for mapping in photo_mappings:
                if (
                    "field_name" not in mapping
                    or mapping["field_name"] not in request.files
                ):
                    continue

                photo_file = request.files[mapping["field_name"]]
                if photo_file.filename == "" or not allowed_image(photo_file.filename):
                    continue

                photo_path = os.path.join(
                    temp_dir, secure_filename(photo_file.filename)
                )
                photo_file.save(photo_path)

                # Insert photo by mapping index
                if "slot_index" in mapping:
                    inserter.insert_photo(mapping["slot_index"], photo_path)

            # Replace text
            inserter.replace_all_text(text_values)

            # Generate output filename
            template_basename = os.path.splitext(
                secure_filename(template_file.filename)
            )[0]
            output_filename = f"{project_code}_ATP_Photos_{template_basename}.docx"

        else:
            return jsonify({"error": "Unsupported file type"}), 400

        # Save the modified document
        output_path = os.path.join(temp_dir, output_filename)
        inserter.save(output_path)

        return jsonify(
            {
                "success": True,
                "message": f"Successfully processed template with {len(photo_mappings)} photos and replaced {len(text_values)} text fields",
                "download_url": f"/download/{os.path.basename(temp_dir)}/{output_filename}",
                "temp_dir": temp_dir,
            }
        )

    except Exception as e:
        import traceback

        print(f"Error in upload_photos: {str(e)}")
        print(traceback.format_exc())
        return jsonify({"error": str(e)}), 500


@app.route("/download/<temp_dir>/<filename>")
def download_file(temp_dir, filename):
    file_path = (
        os.path.join("/tmp", temp_dir, filename)
        if "tmp" in temp_dir
        else os.path.join(temp_dir, filename)
    )
    return send_file(file_path, as_attachment=True)


if __name__ == "__main__":
    os.makedirs(app.config["UPLOAD_FOLDER"], exist_ok=True)
    app.run(debug=True)
