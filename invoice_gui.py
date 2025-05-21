import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
import os
import re

from main import RecordProcessor
from invoice_generato import InvoiceDocGeneratorXML
from mail_invoice import send_invoices_for_records          # still calls docx_to_pdf internally


class InvoiceApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("Invoice Generator & Mailer")
        self.root.geometry("600x400")
        self.root.resizable(False, False)

        # ── Tkinter StringVars ───────────────────────────────────────────
        self.excel_path      = tk.StringVar()
        self.template_path   = tk.StringVar()
        self.output_dir      = tk.StringVar()
        self.sender_email    = tk.StringVar()
        self.sender_password = tk.StringVar()

        # Build the widgets
        self._build_gui()

    # ────────────────────────────────────────────────────────────────────
    #  GUI layout
    # ────────────────────────────────────────────────────────────────────
    def _build_gui(self):
        pad = 10
        row = 0

        # Excel file
        ttk.Label(self.root, text="Select Excel File:").grid(row=row, column=0,
                                                             sticky="e", padx=pad, pady=pad)
        ttk.Entry(self.root, textvariable=self.excel_path, width=40).grid(row=row, column=1, padx=pad)
        ttk.Button(self.root, text="Browse", command=self.browse_excel).grid(row=row, column=2, padx=pad)
        row += 1

        # Template DOCX
        ttk.Label(self.root, text="Select Template DOCX:").grid(row=row, column=0,
                                                                sticky="e", padx=pad, pady=pad)
        ttk.Entry(self.root, textvariable=self.template_path, width=40).grid(row=row, column=1, padx=pad)
        ttk.Button(self.root, text="Browse", command=self.browse_template).grid(row=row, column=2, padx=pad)
        row += 1
        ttk.Label(self.root,
                  text="Template DOCX: This is the Word file used as the invoice format.",
                  foreground="gray").grid(row=row, column=1, sticky="w", padx=pad)
        row += 1

        # Output directory
        ttk.Label(self.root, text="Select Output Directory:").grid(row=row, column=0,
                                                                   sticky="e", padx=pad, pady=pad)
        ttk.Entry(self.root, textvariable=self.output_dir, width=40).grid(row=row, column=1, padx=pad)
        ttk.Button(self.root, text="Browse", command=self.browse_output_dir).grid(row=row, column=2, padx=pad)
        row += 1

        # Sender email
        ttk.Label(self.root, text="Sender Email:").grid(row=row, column=0,
                                                        sticky="e", padx=pad, pady=pad)
        ttk.Entry(self.root, textvariable=self.sender_email, width=40).grid(row=row, column=1, padx=pad)
        row += 1

        # Sender app password
        ttk.Label(self.root, text="Sender App Password:").grid(row=row, column=0,
                                                               sticky="e", padx=pad, pady=pad)
        ttk.Entry(self.root, textvariable=self.sender_password, show="*", width=40).grid(row=row, column=1, padx=pad)
        row += 1

        # Run button  ➜  now runs _synchronously_ on the main thread
        ttk.Button(self.root, text="Run", command=self.run_process).grid(row=row, column=1, pady=pad)
        row += 1

        # Progress / log box
        self.progress = tk.Text(self.root, height=8, width=70,
                                state='disabled', bg="#f7f7f7")
        self.progress.grid(row=row, column=0, columnspan=3, padx=pad, pady=pad)

    # ────────────────────────────────────────────────────────────────────
    #  File/directory pickers
    # ────────────────────────────────────────────────────────────────────
    def browse_excel(self):
        path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if path:
            self.excel_path.set(path)

    def browse_template(self):
        path = filedialog.askopenfilename(filetypes=[("Word files", "*.docx")])
        if path:
            self.template_path.set(path)

    def browse_output_dir(self):
        path = filedialog.askdirectory()
        if path:
            self.output_dir.set(path)

    # ────────────────────────────────────────────────────────────────────
    #  Logging helpers
    # ────────────────────────────────────────────────────────────────────
    def log(self, msg: str):
        self.progress.config(state='normal')
        self.progress.insert('end', msg + '\n')
        self.progress.see('end')
        self.progress.config(state='disabled')
        self.root.update()

    # ────────────────────────────────────────────────────────────────────
    #  Main process (runs on the main thread to keep docx2pdf happy)
    # ────────────────────────────────────────────────────────────────────
    def run_process(self):
        try:
            # Clear log
            self.progress.config(state='normal')
            self.progress.delete('1.0', 'end')
            self.progress.config(state='disabled')

            # Collect user input
            excel    = self.excel_path.get()
            template = self.template_path.get()
            output   = self.output_dir.get()
            email    = self.sender_email.get()
            password = self.sender_password.get()

            if not all([excel, template, output, email, password]):
                self.log("Please fill in all fields.")
                return

            # 1. Read Excel
            self.log("Reading and processing Excel file …")
            processor = RecordProcessor()
            processor.read_excel_file(excel)
            processed_records = processor.process_records()
            self.log(f"Processed {len(processor.original_records)} records.")

            # 2. Generate DOCX invoices
            self.log("Generating invoice DOCX files …")
            invoice_generator = InvoiceDocGeneratorXML(
                template_path=template,
                output_dir=output,
                data_list=processed_records
            )
            invoice_generator.generate_documents()
            self.log("DOCX invoices generated.")

            # 3. Send emails with DOCX attachments
            self.log("Sending DOCX invoices via email …")
            send_invoices_for_records(
                processed_records,
                output,
                email,
                password
            )

            self.log("✅ All done! Invoices generated and emails sent.")

        except Exception as e:
            self.log(f"Error: {e}")
            messagebox.showerror("Error", str(e))


# ────────────────────────────────────────────────────────────────────────
#  Launch the app
# ────────────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    root = tk.Tk()
    app = InvoiceApp(root)
    root.mainloop()
