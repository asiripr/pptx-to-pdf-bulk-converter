# main.py - PPTX -> PDF using PowerPoint (Windows)
import os
import threading
import traceback
import comtypes.client
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

MAX_FILES = 10

def pptx_to_pdf_com(pptx_path, pdf_path):
    import os
    powerpoint = comtypes.client.CreateObject("PowerPoint.Application")
    powerpoint.Visible = 1

    # Ensure absolute path and proper escaping
    pptx_path = os.path.abspath(pptx_path)
    pdf_path = os.path.abspath(pdf_path)

    presentation = powerpoint.Presentations.Open(pptx_path, WithWindow=False)
    presentation.SaveAs(pdf_path, 32)
    presentation.Close()
    powerpoint.Quit()


class App:
    def __init__(self, root):
        self.root = root
        root.title("PPTX → PDF (Windows PowerPoint)")
        root.geometry("600x380")

        frm = tk.Frame(root, padx=10, pady=10)
        frm.pack(fill=tk.BOTH, expand=True)

        tk.Label(frm, text=f"Select up to {MAX_FILES} .pptx files").pack(anchor="w")

        btn_frame = tk.Frame(frm)
        btn_frame.pack(fill="x", pady=(6,6))
        tk.Button(btn_frame, text="Browse", command=self.select_files).pack(side="left")
        tk.Button(btn_frame, text="Select Output Folder", command=self.select_out_dir).pack(side="left", padx=6)
        tk.Button(btn_frame, text="Convert to PDF", command=self.start_convert).pack(side="right")

        self.listbox = tk.Listbox(frm, height=10, width=80)
        self.listbox.pack(pady=6, fill="both", expand=True)

        self.progress = ttk.Progressbar(frm, orient="horizontal", mode="determinate")
        self.progress.pack(fill="x", pady=(6,2))

        self.status_lbl = tk.Label(frm, text="Ready")
        self.status_lbl.pack(anchor="w")

        self.selected = []
        self.out_dir = None

    def select_files(self):
        files = filedialog.askopenfilenames(title="Select PPTX files", filetypes=[("PowerPoint Files", "*.pptx")])
        files = list(files)[:MAX_FILES]
        self.selected = files
        self.listbox.delete(0, tk.END)
        for f in files:
            self.listbox.insert(tk.END, f)
        self.status_lbl.config(text=f"{len(files)} file(s) selected")

    def select_out_dir(self):
        d = filedialog.askdirectory(title="Select output folder (PDFs will be saved here)")
        if d:
            self.out_dir = d
            self.status_lbl.config(text=f"Output folder: {d}")

    def start_convert(self):
        if not self.selected:
            messagebox.showwarning("No files", "Please select up to 10 PPTX files first.")
            return
        # disable UI
        self.progress["maximum"] = len(self.selected)
        self.progress["value"] = 0
        self.status_lbl.config(text="Starting conversion...")
        threading.Thread(target=self.convert_thread, daemon=True).start()

    def convert_thread(self):
        try:
            for idx, pptx in enumerate(self.selected, start=1):
                self.status_lbl.config(text=f"Converting ({idx}/{len(self.selected)}): {os.path.basename(pptx)}")
                out_folder = self.out_dir or os.path.dirname(pptx)
                pdf_path = os.path.join(out_folder, os.path.splitext(os.path.basename(pptx))[0] + ".pdf")
                pptx_to_pdf_com(pptx, pdf_path)
                self.progress["value"] = idx
            self.status_lbl.config(text="All conversions completed.")
            messagebox.showinfo("Done", "All files converted successfully.")
        except Exception as e:
            traceback.print_exc()
            messagebox.showerror("Error", f"Conversion failed:\n{e}")
            self.status_lbl.config(text="Error occurred")
        finally:
            self.progress["value"] = 0

if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()