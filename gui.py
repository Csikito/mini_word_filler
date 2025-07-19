import tkinter as tk
from tkinter import filedialog, messagebox
from pathlib import Path
from core import process_input, get_unique_output_folder, fill_files
from tkinter import ttk
import webbrowser
import os

class WordFillerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Word Template Auto-Filler")
        self.root.geometry("450x350")
        self.root.minsize(450, 350)

        self.excel_path = None
        self.template_paths = []
        self.output_dir = None
        self.generated_paths = []

        style = ttk.Style()
        style.theme_use("clam")
        style.configure("TButton", font=("Arial", 10))

        # Excel file
        frame_excel = tk.Frame(root)
        frame_excel.pack(pady=(20,5), fill="x", padx=20)
        self.excel_label = tk.Label(frame_excel, text="No Excel file selected")
        self.excel_label.pack(side="left")
        ttk.Button(frame_excel, text="📄", command=self.select_excel, cursor="hand2").pack(side="right")

        # Word templates
        frame_templates = tk.Frame(root)
        frame_templates.pack(pady=5, fill="x", padx=20)
        self.templates_label = tk.Label(frame_templates, text="No Word templates selected")
        self.templates_label.pack(side="left")
        ttk.Button(frame_templates, text="📂", command=self.select_templates, cursor="hand2").pack(side="right")

        # Output folder
        frame_output = tk.Frame(root)
        frame_output.pack(pady=5, fill="x", padx=20)
        self.output_label = tk.Label(frame_output, text="No output folder selected (will auto-create)")
        self.output_label.pack(side="left")
        ttk.Button(frame_output, text="📁", command=self.select_output_dir, cursor="hand2").pack(side="right")

        # Process button
        tk.Button(root, text="▶ Fill Templates", bg="#28a745", fg="white", font=("Arial", 12, "bold"), command=self.process_files, cursor="hand2").pack(pady=20)

        # Result
        self.result_frame = tk.Frame(root)
        self.result_frame.pack(padx=20, pady=10)

        self.status_container = tk.Frame(self.result_frame)
        canvas = tk.Canvas(self.status_container, height=100, width=340, bg=self.root["bg"], highlightthickness=0)
        scrollbar = tk.Scrollbar(self.status_container, orient="vertical", command=canvas.yview)
        self.status_frame = tk.Frame(canvas, bg=self.root["bg"])

        self.status_label = tk.Label(self.status_frame, text="", fg="green", wraplength=600, justify="left")
        self.status_label.pack()

        self.status_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )

        canvas.create_window((0, 0), window=self.status_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # Open Output folder
        self.open_folder_button = ttk.Button(self.result_frame, text="📂", width=3, command=self.open_output_folder,
                                             state="disabled", cursor="hand2")
        self.open_folder_button.pack(side="left", anchor="w", padx=2)

    def select_excel(self):
        path = filedialog.askopenfilename(filetypes=[["Excel files", "*.xlsx"]])
        if path:
            self.excel_path = Path(path)
            self.excel_label.config(text=f"Excel: {self.excel_path.name}")

    def select_templates(self):
        paths = filedialog.askopenfilenames(filetypes=[["Word files", "*.docx"]])
        if paths:
            self.template_paths = list(map(Path, paths))
            self.templates_label.config(text=f"{len(self.template_paths)} template(s) selected")

    def select_output_dir(self):
        path = filedialog.askdirectory()
        if path:
            self.output_dir = Path(path)
            self.output_label.config(text=f"Output: {self.output_dir}")

    def open_file(self, filepath):
        try:
            webbrowser.open(filepath)
        except Exception as e:
            messagebox.showerror("Open File Error", f"Failed to open file:\n{e}")

    def open_output_folder(self):
        if hasattr(self, 'generated_output_dir') and self.generated_output_dir:
            os.startfile(self.generated_output_dir)

    def process_files(self):
        if not self.excel_path or not self.template_paths:
            messagebox.showerror("Missing input", "Please select both Excel file and Word templates.")
            return

        for widget in self.status_frame.winfo_children():
            widget.destroy()

        data = process_input(self.excel_path)
        output_dir = get_unique_output_folder(self.output_dir)
        self.generated_output_dir = output_dir
        generated = fill_files(data, self.template_paths, output_dir)

        if generated:
            files = list(output_dir.glob("*.docx"))

            self.status_label = tk.Label(self.status_frame, text=f"✅ {generated} file(s) generated:", fg="green",
                                         bg=self.root["bg"])
            self.status_label.pack()
            for f in files:
                link = tk.Label(self.status_frame, text=f"→ {f.name}", fg="blue", bg=self.root["bg"], cursor="hand2")
                link.pack(anchor="w")
                link.bind("<Button-1>", lambda e, path=f: self.open_file(path))

            self.status_container.pack(fill="both", padx=(10, 0), expand=True,)
            self.open_folder_button.config(state="normal")

        else:
            self.status_label = tk.Label(self.status_frame, text="⚠ No files were generated.", fg="red",
                                         bg=self.root["bg"])
            self.status_label.pack()
            self.open_folder_button.pack_forget()
