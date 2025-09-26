from flask import Flask, request, render_template_string, send_file
import pandas as pd
import io

app = Flask(__name__)

HTML = """
<!doctype html>
<title>XLS → TXT Converter</title>
<h2>Upload Excel → Download Custom TXT (UTF-8)</h2>
<form method=post enctype=multipart/form-data>
  <input type=file name=file accept=".xls,.xlsx">
  <input type=submit value=Convert>
</form>
"""

# Mapping for รหัสชั้น
class_map = {
    "อ.1": 1, "อ.2": 2, "อ.3": 3,
    "ป.1": 4, "ป.2": 5, "ป.3": 6,
    "ป.4": 7, "ป.5": 8, "ป.6": 9,
    "ม.1": 10, "ม.2": 11, "ม.3": 12
}

# Mapping for เพศ
gender_map = {"ชาย": 1, "หญิง": 2}

@app.route("/", methods=["GET", "POST"])
def upload_file():
    if request.method == "POST":
        f = request.files["file"]
        if f and (f.filename.endswith(".xlsx") or f.filename.endswith(".xls")):
            # Read Excel
            df = pd.read_excel(f)

            # Convert รหัสชั้น
            df["รหัสชั้น_num"] = df["ชั้น"].map(class_map)

            # Convert เพศ
            df["เพศรหัส"] = df["เพศ"].map(gender_map)

            # Sort by class + room + student ID
            df = df.sort_values(by=["รหัสชั้น_num", "ห้อง", "รหัสนักเรียน"])

            # Generate เลขที่ (reset per class+room)
            df["เลขที่"] = df.groupby(["รหัสชั้น_num", "ห้อง"]).cumcount() + 1

            # Column mapping
            col_map = {
                "รหัสนักเรียน": "เลขประจำตัวนักเรียน",
                "เลขประจำตัวประชาชน": "เลขประจำตัวประชาชน",
                "เลขที่": "เลขที่",
                "คำนำหน้าชื่อ": "คำนำหน้าชื่อ",
                "ชื่อ": "ชื่อ",
                "นามสกุล": "นามสกุล",
                "เพศรหัส": "รหัสเพศ",
                "รหัสชั้น_num": "รหัสชั้น",
                "ห้อง": "ห้องที่"
            }

            # Create output DataFrame
            df_out = df[list(col_map.keys())].rename(columns=col_map)

            # Save to UTF-8 txt (tab-separated)
            output = io.StringIO()
            df_out.to_csv(output, sep="\t", index=False, encoding="utf-8")

            mem = io.BytesIO()
            mem.write(output.getvalue().encode("utf-8"))
            mem.seek(0)

            return send_file(
                mem,
                as_attachment=True,
                download_name=f.filename.rsplit(".", 1)[0] + "_filtered.txt",
                mimetype="text/plain; charset=utf-8"
            )

    return render_template_string(HTML)

if __name__ == "__main__":
    app.run(debug=True)
