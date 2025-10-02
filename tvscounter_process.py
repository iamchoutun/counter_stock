from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string
from collections import defaultdict

def tvscounter_process(input_file, output_file):
    wb = load_workbook(input_file)
    ws = wb.active

    tvs_model_col = column_index_from_string("C") - 1  # Model
    tvs_status_col = column_index_from_string("F") - 1 # Status

    summary = {
        "Hybrid": defaultdict(lambda: {"Good": 0, "Defect": 0}),
        "SKWAMX3 : TRUEIDTVGEN2 SKY TICC": {"Good": 0, "Defect": 0},
        "SKWAMX5M : SMARTTRUEIDTVGEN3 T3 SKY TICC": {"Good": 0, "Defect": 0},
        "SKWAMX5M-NO : SMARTTRUETDTVGEN3.1 T3 SKY TICC": {"Good": 0, "Defect": 0}
    }

    hybrid_models = [
        "C-HD ATV : SMC HD SKY TVG CATV",
        "HBSK-NET100C : STD HYBRID SKY100 TVG"
    ]

    for row in ws.iter_rows(min_row=2, values_only=True):
        model = row[tvs_model_col]
        status = str(row[tvs_status_col] or "").strip().upper()
        is_defect = status == "DEFECT"

        if model in hybrid_models:
            summary["Hybrid"][model]["Defect" if is_defect else "Good"] += 1
        elif model == "SKWAMX3 : TRUEIDTVGEN2 SKY TICC":
            summary[model]["Defect" if is_defect else "Good"] += 1
        elif model == "SKWAMX5M : SMARTTRUEIDTVGEN3 T3 SKY TICC":
            summary[model]["Defect" if is_defect else "Good"] += 1
        elif model == "SKWAMX5M-NO : SMARTTRUETDTVGEN3.1 T3 SKY TICC":
            summary[model]["Defect" if is_defect else "Good"] += 1

    # สร้าง Total สำหรับ Hybrid
    total_good = sum(v["Good"] for v in summary["Hybrid"].values())
    total_defect = sum(v["Defect"] for v in summary["Hybrid"].values())
    summary["Hybrid"]["Total"] = {"Good": total_good, "Defect": total_defect}

    return summary
