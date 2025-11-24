from flask import Flask, render_template, request, send_file
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from openpyxl import Workbook
import io

app = Flask(__name__)

# store last result for export
last_result = {}


@app.route("/", methods=["GET", "POST"])
def index():
    global last_result
    result = {}

    # Location → UCS map
    location_map = {
        "kg": ("KG Basin", 12000),
        "bombay": ("Bombay High", 18000),
        "rajasthan": ("Rajasthan", 25000),
        "assam": ("Assam", 15000),
        "gulf": ("Middle East / Gulf", 28000),
    }

    # Formation → UCS map
    formation_map = {
        "shale": ("Shale", 8000),
        "sandstone": ("Sandstone", 14000),
        "limestone": ("Limestone", 20000),
        "dolomite": ("Dolomite", 26000),
        "granite": ("Granite / Basement", 35000),
    }

    if request.method == "POST":
        # Inputs
        WOB = float(request.form.get("wob", 0))
        RPM = float(request.form.get("rpm", 0))
        depth = float(request.form.get("depth", 1))
        rig_cost = float(request.form.get("rig_cost", 0))
        bit_cost = float(request.form.get("bit_cost", 0))
        abrasiveness = float(request.form.get("abrasiveness", 1))

        location_key = request.form.get("location", "")
        formation_key = request.form.get("formation", "")
        ucs_input = request.form.get("ucs", "")

        # ---- UCS SELECTION PRIORITY ----
        source_used = ""
        location_name = "N/A"
        formation_name = "N/A"

        if location_key in location_map:
            location_name, UCS = location_map[location_key]
            source_used = "Location based"
        elif ucs_input:
            UCS = float(ucs_input)
            source_used = "User UCS"
        elif formation_key in formation_map:
            formation_name, UCS = formation_map[formation_key]
            source_used = "Formation based"
        else:
            UCS = 15000
            source_used = "Default"

        # ---- AI-STYLE ROP MODEL ----
        UCS_kpsi = UCS / 1000.0
        ai_factor = 1 + (abrasiveness * 0.05) + (WOB * 0.005)

        ROP = (0.25 * (WOB ** 1.1) * (RPM ** 0.9) / max(UCS_kpsi, 1)) * ai_factor
        if ROP <= 0:
            ROP = 0.1

        drilling_time = depth / ROP
        cost_per_meter = ((rig_cost * drilling_time) + bit_cost) / depth
        bit_life = 400000 / (WOB * RPM * max(abrasiveness, 0.1))

        # ---- COST & ROP vs WOB CURVES + OPTIMUM POINT ----
        wob_curve = []
        cost_curve = []
        rop_curve = []

        min_cost = None
        best_wob = None

        for w in range(5, 65, 5):
            temp_rop = (0.25 * (w ** 1.1) * (RPM ** 0.9) / max(UCS_kpsi, 1)) * ai_factor
            if temp_rop <= 0:
                temp_rop = 0.1

            temp_time = depth / temp_rop
            temp_cost = ((rig_cost * temp_time) + bit_cost) / depth

            wob_curve.append(w)
            rop_curve.append(round(temp_rop, 2))
            cost_curve.append(round(temp_cost, 2))

            if min_cost is None or temp_cost < min_cost:
                min_cost = temp_cost
                best_wob = w

        best_rpm = RPM  # for now, RPM optimum = current RPM

        result = {
            "source": source_used,
            "location": location_name,
            "formation_name": formation_name,
            "ucs": UCS,
            "rop": round(ROP, 2),
            "time": round(drilling_time, 2),
            "cost": round(cost_per_meter, 2),
            "bitlife": round(bit_life, 2),

            "best_wob": best_wob,
            "best_rpm": best_rpm,
            "min_cost": round(min_cost, 2),

            "wob_curve": wob_curve,
            "cost_curve": cost_curve,
            "rop_curve": rop_curve,
        }

        last_result = result

    return render_template("index.html", result=result)


# ------------ PDF EXPORT ------------
@app.route("/export/pdf")
def export_pdf():
    if not last_result:
        return "No calculation found. Run a calculation first.", 400

    buffer = io.BytesIO()
    c = canvas.Canvas(buffer, pagesize=letter)
    c.setTitle("Drilling AI Optimizer Report")

    y = 750
    c.setFont("Helvetica-Bold", 16)
    c.drawString(100, y, "Drilling AI Optimizer Report")
    y -= 40

    c.setFont("Helvetica", 11)
    for key, value in last_result.items():
        if isinstance(value, (list, dict)):
            continue
        c.drawString(80, y, f"{key}: {value}")
        y -= 18
        if y < 80:
            c.showPage()
            y = 750

    c.showPage()
    c.save()
    buffer.seek(0)

    return send_file(
        buffer,
        as_attachment=True,
        download_name="drilling_report.pdf",
        mimetype="application/pdf",
    )


# ------------ EXCEL EXPORT ------------
@app.route("/export/excel")
def export_excel():
    if not last_result:
        return "No calculation found. Run a calculation first.", 400

    wb = Workbook()
    ws = wb.active
    ws.title = "Drilling Report"

    ws.append(["Parameter", "Value"])
    for key, value in last_result.items():
        if isinstance(value, (list, dict)):
            continue
        ws.append([key, value])

    file_stream = io.BytesIO()
    wb.save(file_stream)
    file_stream.seek(0)

    return send_file(
        file_stream,
        as_attachment=True,
        download_name="drilling_report.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


if __name__ == "__main__":
    app.run(debug=True)
