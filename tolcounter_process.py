from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string

def tolcounter_process(input_file, output_file):
    wb = load_workbook(input_file)
    ws = wb.active

    # แปลง column letter → index (0-based)
    tol_model_col = column_index_from_string("I") - 1  # Model
    tol_status_col = column_index_from_string("M") - 1  # Status
    tol_dongle_col = column_index_from_string("H") - 1  # Dongle flag

    summary = {}
    good_sum = 0
    defect_sum = 0

    # ใช้ values_only=True → row จะเป็น tuple ของค่า (ไม่ใช่ cell)
    for row in ws.iter_rows(min_row=2, values_only=True):
        model = row[tol_model_col]
        status = row[tol_status_col]
        dongle_flag = row[tol_dongle_col]

        # แยก T626ProV2 ตาม Dongle
        if model == "T626ProV2":
            if dongle_flag and "dongle" in str(dongle_flag).lower():
                model_key = "T626ProV2_Dongle"
            else:
                model_key = "T626ProV2"
        else:
            model_key = model

        # ถ้า model_key ยังไม่มีใน summary ให้สร้าง dict ใหม่
        if model_key not in summary:
            summary[model_key] = {"Good": 0, "Defect": 0}

        # นับ Good/Defect
        if status and "DEFECT" in str(status).upper():
            summary[model_key]["Defect"] += 1
            defect_sum += 1
        else:
            summary[model_key]["Good"] += 1
            good_sum += 1

    return summary
