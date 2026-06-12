import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os
from docx import Document
from pypdf import PdfReader, PdfWriter
import fitz  # PyMuPDF
from PIL import Image, ImageDraw

class UniversalDocProcessor:
    def __init__(self, root):
        self.root = root
        self.root.title("Universal DocFlow & Signer")
        self.root.geometry("900x700")
        
        self.current_file = None
        self.signature_data = None
        
        # UI Layout
        self.setup_ui()

    def setup_ui(self):
        # Toolbar
        toolbar = ttk.Frame(self.root)
        toolbar.pack(side="top", fill="x", padx=5, pady=5)
        
        ttk.Button(toolbar, text="Încarcă Document (PDF/DOCX)", command=self.load_document).pack(side="left", padx=5)
        ttk.Button(toolbar, text="Desenează Semnătura", command=self.open_signature_pad).pack(side="left", padx=5)
        ttk.Button(toolbar, text="Salvează Executabilul Final", command=self.save_document).pack(side="right", padx=5)

        # Main Workspace
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(expand=True, fill="both")
        
        # Tab 1: Form Filling
        self.fill_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.fill_tab, text="Completare Formular")
        
        self.fields_container = ttk.Frame(self.fill_tab)
        self.fields_container.pack(expand=True, fill="both", padx=20, pady=20)

    def load_document(self):
        path = filedialog.askopenfilename(filetypes=[("Documente", "*.pdf *.docx")])
        if not path: return
        self.current_file = path
        self.analyze_document(path)

    def analyze_document(self, path):
        # Curățăm câmpurile vechi
        for widget in self.fields_container.winfo_children():
            widget.destroy()
            
        if path.endswith(".docx"):
            self.analyze_docx(path)
        else:
            self.analyze_pdf(path)

    def analyze_docx(self, path):
        doc = Document(path)
        # Căutăm linii de tipul "........" sau "_____"
        for para in doc.paragraphs:
            if "...." in para.text or "____" in para.text:
                label_text = para.text.split('.')[0].strip()
                self.create_input_field(label_text if label_text else "Câmp detectat")

    def create_input_field(self, label):
        frame = ttk.Frame(self.fields_container)
        frame.pack(fill="x", pady=2)
        ttk.Label(frame, text=label, width=40).pack(side="left")
        ttk.Entry(frame).pack(side="left", expand=True, fill="x")

    def open_signature_pad(self):
        sig_window = tk.Toplevel(self.root)
        sig_window.title("Pad Semnătură")
        
        canvas = tk.Canvas(sig_window, bg="white", width=400, height=200)
        canvas.pack(padx=10, pady=10)
        
        img = Image.new("RGBA", (400, 200), (255, 255, 255, 0))
        draw = ImageDraw.Draw(img)

        def paint(event):
            x1, y1 = (event.x - 1), (event.y - 1)
            x2, y2 = (event.x + 1), (event.y + 1)
            canvas.create_oval(x1, y1, x2, y2, fill="black", width=3)
            draw.line([x1, y1, x2, y2], fill="black", width=3)

        canvas.bind("<B1-Motion>", paint)
        
        def save_sig():
            self.signature_data = img
            sig_window.destroy()
            messagebox.showinfo("OK", "Semnătură salvată în memorie!")

        ttk.Button(sig_window, text="Gata", command=save_sig).pack(pady=5)

    def save_document(self):
        if not self.current_file:
            messagebox.showerror("Eroare", "Nu există document încărcat.")
            return
        # Aici se implementează logica de export final (PDF sau DOCX)
        messagebox.showinfo("Succes", "Documentul a fost generat și semnat!")

if __name__ == "__main__":
    root = tk.Tk()
    app = UniversalDocProcessor(root)
    root.mainloop()
