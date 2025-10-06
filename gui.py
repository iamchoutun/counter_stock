import tkinter as tk
from tkinter import filedialog, ttk, messagebox
from pathlib import Path
from tolcounter_process import tolcounter_process as tol_process
from tvscounter_process import tvscounter_process as tvs_process

class StockApp:
    """
    GUI application for counting TOL and TVS stock items from Excel files.
    """

    def __init__(self, root: tk.Tk):
        """
        Initialize the main GUI window and widgets.

        Args:
            root (tk.Tk): The root Tkinter window.
        """
        self.root = root
        self.root.title("Counter Stock 2.0")
        self.root.geometry("800x400")

        # Set window icon if exists (Windows only)
        icon_path = Path("icon.ico").resolve()
        if icon_path.exists():
            try:
                root.iconbitmap(str(icon_path))
            except tk.TclError:
                pass  # Ignore errors on non-Windows platforms

        self.tol_file = ""
        self.tvs_file = ""

        # Create GUI components
        self.create_file_selection_frame()
        self.create_treeviews()  # Correctly create both Treeviews

    # ================= GUI CREATION =================
    def create_file_selection_frame(self):
        """Create file selection buttons and labels."""
        frame = tk.Frame(self.root)
        frame.pack(pady=10)

        self.tol_label = tk.Label(frame, text="TOL File: None", width=50, anchor="w")
        self.tol_label.grid(row=0, column=0)
        self.tvs_label = tk.Label(frame, text="TVS File: None", width=50, anchor="w")
        self.tvs_label.grid(row=1, column=0)

        tk.Button(frame, text="Browse TOL", command=self.browse_tol).grid(row=0, column=1)
        tk.Button(frame, text="Browse TVS", command=self.browse_tvs).grid(row=1, column=1)
        tk.Button(frame, text="Run Process", command=self.run_process).grid(row=0, column=2, rowspan=2)

    def create_treeviews(self):
        """Create both TOL and TVS Treeview widgets inside the main window."""
        tree_frame = tk.Frame(self.root)
        tree_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # TOL Treeview
        tol_frame = tk.Frame(tree_frame)
        tol_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5)
        tk.Label(tol_frame, text="TOL Results").pack()
        self.tree_tol = self.create_treeview(tol_frame)  # Pass parent frame

        # TVS Treeview
        tvs_frame = tk.Frame(tree_frame)
        tvs_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=5)
        tk.Label(tvs_frame, text="TVS Results").pack()
        self.tree_tvs = self.create_treeview(tvs_frame)  # Pass parent frame

    def create_treeview(self, parent: tk.Frame) -> ttk.Treeview:
        """
        Create a single Treeview with vertical scrollbar.

        Args:
            parent (tk.Frame): The parent frame for the Treeview.

        Returns:
            ttk.Treeview: Configured Treeview widget.
        """
        frame = tk.Frame(parent)
        frame.pack(fill=tk.BOTH, expand=True)
        scroll = tk.Scrollbar(frame)
        scroll.pack(side=tk.RIGHT, fill=tk.Y)

        tree = ttk.Treeview(frame, yscrollcommand=scroll.set)
        tree["columns"] = ("Model", "Good", "Defect", "Total")
        tree["show"] = "headings"  # Hide default first column

        # Define headings
        tree.heading("Model", text="Model / Category")
        tree.heading("Good", text="Good")
        tree.heading("Defect", text="Defect")
        tree.heading("Total", text="Total")

        # Define column widths
        tree.column("Model", width=60, anchor="w")
        tree.column("Good", width=20, anchor="center")
        tree.column("Defect", width=20, anchor="center")
        tree.column("Total", width=20, anchor="center")

        tree.pack(fill=tk.BOTH, expand=True)
        scroll.config(command=tree.yview)

        return tree

    # ================= FILE BROWSING =================
    def browse_tol(self):
        """Browse and select TOL Excel file."""
        path = filedialog.askopenfilename(filetypes=[("Excel files","*.xlsx *.xlsm")])
        if path:
            self.tol_file = path
            self.tol_label.config(text=f"TOL File: {self.tol_file}")

    def browse_tvs(self):
        """Browse and select TVS Excel file."""
        path = filedialog.askopenfilename(filetypes=[("Excel files","*.xlsx *.xlsm")])
        if path:
            self.tvs_file = path
            self.tvs_label.config(text=f"TVS File: {self.tvs_file}")

    # ================= RUN PROCESS =================
    def run_process(self):
        """Process selected Excel files and populate Treeviews."""
        # Clear previous data
        for tree in [self.tree_tol, self.tree_tvs]:
            for item in tree.get_children():
                tree.delete(item)

        # Process TOL file
        if self.tol_file:
            try:
                summary = tol_process(self.tol_file, None)
                self.fill_treeview_tol(self.tree_tol, summary)
            except Exception as e:
                messagebox.showerror("TOL Error", str(e))

        # Process TVS file
        if self.tvs_file:
            try:
                summary = tvs_process(self.tvs_file, None)
                self.fill_treeview_tvs(self.tree_tvs, summary, is_tvs=True)
            except Exception as e:
                messagebox.showerror("TVS Error", str(e))

        messagebox.showinfo("Done", "Processing completed!")

    # ================= FILL TREEVIEW =================
    def fill_treeview_tol(self, tree: ttk.Treeview, summary: dict):
        """Fill TOL Treeview with summarized data."""
        # Sort models A-Z and insert rows
        for model, counts in sorted(summary.items()):
            display_model = "T626Pro_AI" if model == "T626Pro" else model
            good = counts.get("Good", 0)
            defect = counts.get("Defect", 0)
            total = good + defect
            tree.insert("", "end", values=(display_model, good, defect, total))

        # Add grand total row
        total_good = sum(c["Good"] for c in summary.values())
        total_defect = sum(c["Defect"] for c in summary.values())
        tree.insert("", "end", values=("GRAND TOTAL", total_good, total_defect, total_good + total_defect))

    def fill_treeview_tvs(self, tree: ttk.Treeview, summary: dict, is_tvs: bool = False):
        """Fill TVS Treeview with summarized data, including Hybrid models."""
        total_good = 0
        total_defect = 0
        tvs_name_map = {
            "SKWAMX3 : TRUEIDTVGEN2 SKY TICC": "True ID Gen2",
            "SKWAMX5M : SMARTTRUEIDTVGEN3 T3 SKY TICC": "True ID Gen3",
            "SKWAMX5M-NO : SMARTTRUETDTVGEN3.1 T3 SKY TICC": "True ID Gen3.1 (Purple)"
        } if is_tvs else {}

        for key, val in summary.items():
            display_key = key if not is_tvs else tvs_name_map.get(key, key)

            # Special handling for Hybrid models
            if is_tvs and key == "Hybrid":
                parent = tree.insert("", tk.END, values=("Hybrid", "", "", ""), open=True)
                for sub_key, sub_counts in val.items():
                    if sub_key == "Total":
                        continue
                    g = sub_counts.get("Good", 0)
                    d = sub_counts.get("Defect", 0)
                    tree.insert(parent, tk.END, values=(sub_key, g, d, g + d))
                    total_good += g
                    total_defect += d
                # Insert Total for Hybrid
                if "Total" in val:
                    t = val["Total"]
                    tree.insert(parent, tk.END, values=("TOTAL", t.get("Good", 0), t.get("Defect", 0), t.get("Good",0) + t.get("Defect",0)))
            else:
                g = val.get("Good", 0)
                d = val.get("Defect", 0)
                tree.insert("", tk.END, values=(display_key, g, d, g + d))
                total_good += g
                total_defect += d

        # Insert grand total row
        tree.insert("", tk.END, values=("GRAND TOTAL", total_good, total_defect, total_good + total_defect))
