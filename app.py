import tkinter as tk
from tkinter import filedialog, ttk, messagebox
from pathlib import Path
from tolcounter_process import tolcounter_process as tol_process
from tvscounter_process import tvscounter_process as tvs_process



class StockApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Counter Stock 2.0")
        self.root.geometry("800x400")

        icon_path = Path("icon.ico").resolve()
        if icon_path.exists():
            try:
                root.iconbitmap(str(icon_path))  # Set window icon (works on Windows with .ico files)
            except tk.TclError:
                pass  # Ignore icon errors on non-Windows platforms or invalid icon files

        self.tol_file = ""
        self.tvs_file = ""

        frame = tk.Frame(root)
        frame.pack(pady=10)

        self.tol_label = tk.Label(frame, text="TOL File: None", width=50, anchor="w")
        self.tol_label.grid(row=0, column=0)
        self.tvs_label = tk.Label(frame, text="TVS File: None", width=50, anchor="w")
        self.tvs_label.grid(row=1, column=0)

        tk.Button(frame, text="Browse TOL", command=self.browse_tol).grid(row=0, column=1)
        tk.Button(frame, text="Browse TVS", command=self.browse_tvs).grid(row=1, column=1)
        tk.Button(frame, text="Run Process", command=self.run_process).grid(row=0, column=2, rowspan=2)

        # Treeviews
        tree_frame = tk.Frame(root)
        tree_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        tol_frame = tk.Frame(tree_frame)
        tol_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5)
        tk.Label(tol_frame, text="TOL Results").pack()
        self.tree_tol = self.create_treeview(tol_frame)

        tvs_frame = tk.Frame(tree_frame)
        tvs_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=5)
        tk.Label(tvs_frame, text="TVS Results").pack()
        self.tree_tvs = self.create_treeview(tvs_frame)

    def create_treeview(self, parent):
        frame = tk.Frame(parent)
        frame.pack(fill=tk.BOTH, expand=True)
        scroll = tk.Scrollbar(frame)
        scroll.pack(side=tk.RIGHT, fill=tk.Y)
        tree = ttk.Treeview(frame, yscrollcommand=scroll.set)

        # กำหนดคอลัมน์
        tree["columns"] = ("Model", "Good", "Defect", "Total")
        tree["show"] = "headings"  # ซ่อนคอลัมน์ #0

        tree.heading("Model", text="Model / Category")
        tree.heading("Good", text="Good")
        tree.heading("Defect", text="Defect")
        tree.heading("Total", text="Total")

        tree.column("Model", width=60, anchor="w")
        tree.column("Good", width=20, anchor="center")
        tree.column("Defect", width=20, anchor="center")
        tree.column("Total", width=20, anchor="center")

        tree.pack(fill=tk.BOTH, expand=True)
        scroll.config(command=tree.yview)
        return tree

    def browse_tol(self):
        path = filedialog.askopenfilename(filetypes=[("Excel files","*.xlsx *.xlsm")])
        if path:
            self.tol_file = path
            self.tol_label.config(text=f"TOL File: {self.tol_file}")

    def browse_tvs(self):
        path = filedialog.askopenfilename(filetypes=[("Excel files","*.xlsx *.xlsm")])
        if path:
            self.tvs_file = path
            self.tvs_label.config(text=f"TVS File: {self.tvs_file}")

    def run_process(self):
        for tree in [self.tree_tol, self.tree_tvs]:
            for item in tree.get_children():
                tree.delete(item)

        # TOL
        if self.tol_file:
            try:
                summary = tol_process(self.tol_file, None)  # ไม่ต้อง export แล้ว
                self.fill_treeview_tol(self.tree_tol, summary)
            except Exception as e:
                messagebox.showerror("TOL Error", str(e))

        # TVS
        if self.tvs_file:
            try:
                summary = tvs_process(self.tvs_file, None)  # ไม่ต้อง export แล้ว
                self.fill_treeview_tvs(self.tree_tvs, summary, is_tvs=True)
            except Exception as e:
                messagebox.showerror("TVS Error", str(e))

        messagebox.showinfo("Done", "Processing completed!")

    # ================== เติมข้อมูล TOL ==================
    def fill_treeview_tol(self, tree, summary):
        # ล้างข้อมูลเก่า
        for row in tree.get_children():
            tree.delete(row)

        # เรียง Model A-Z
        for model, counts in sorted(summary.items()):
            display_model = "T626Pro_AI" if model == "T626Pro" else model
            good = counts.get("Good", 0)
            defect = counts.get("Defect", 0)
            total = good + defect
            tree.insert("", "end", values=(display_model, good, defect, total))

        # รวมผลทั้งหมด
        total_good = sum(c["Good"] for c in summary.values())
        total_defect = sum(c["Defect"] for c in summary.values())
        grand_total = total_good + total_defect
        tree.insert("", "end", values=("GRAND TOTAL", total_good, total_defect, grand_total))

    # ================== เติมข้อมูล TVS ==================
    def fill_treeview_tvs(self, tree, summary, is_tvs=False):
        total_good = 0
        total_defect = 0
        tvs_name_map = {"SKWAMX3 : TRUEIDTVGEN2 SKY TICC":"True ID Gen2","SKWAMX5M : SMARTTRUEIDTVGEN3 T3 SKY TICC":"True ID Gen3",
                        "SKWAMX5M-NO : SMARTTRUETDTVGEN3.1 T3 SKY TICC":"True ID Gen3.1 (ม่วง)"} if is_tvs else {}

        for key,val in summary.items():
            display_key = key if not is_tvs else tvs_name_map.get(key,key)

            # Hybrid กรณีพิเศษ
            if is_tvs and key=="Hybrid":
                parent = tree.insert("", tk.END, values=("Hybrid", "", "", ""), open=True)
                for sub_key, sub_counts in val.items():
                    if sub_key=="Total":
                        continue
                    g = sub_counts.get("Good",0)
                    d = sub_counts.get("Defect",0)
                    tree.insert(parent, tk.END, values=(sub_key, g, d, g+d))
                    total_good += g
                    total_defect += d
                if "Total" in val:
                    t=val["Total"]
                    tree.insert(parent, tk.END, values=("TOTAL", t.get("Good",0), t.get("Defect",0), t.get("Good",0)+t.get("Defect",0)))
            else:
                g = val.get("Good",0)
                d = val.get("Defect",0)
                tree.insert("", tk.END, values=(display_key, g, d, g+d))
                total_good += g
                total_defect += d

        tree.insert("", tk.END, values=("GRAND TOTAL", total_good, total_defect, total_good+total_defect))


if __name__=="__main__":
    root = tk.Tk()
    app = StockApp(root)
    root.mainloop()
