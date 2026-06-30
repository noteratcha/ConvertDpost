import os
import sys
import glob
import threading
from datetime import datetime
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from tkinter.scrolledtext import ScrolledText
import pandas as pd

# Import conversion functions and version from convert_dpost
try:
    from convert_dpost import process_pdf, records_to_dataframe, __version__
except ImportError:
    __version__ = "2026.0630.1453"
    def process_pdf(path): return []
    def records_to_dataframe(records): return pd.DataFrame()

# Colors for modern UI
COLOR_PRIMARY = "#1e293b"      # Slate 800 (Header)
COLOR_SECONDARY = "#0f766e"    # Teal 700 (Action buttons)
COLOR_ACCENT = "#0284c7"       # Sky 600 (Selection buttons)
COLOR_BG = "#f1f5f9"           # Slate 100 (Window background)
COLOR_CARD = "#ffffff"         # White (Frames background)
COLOR_TEXT = "#0f172a"         # Slate 900 (Main text)
COLOR_TEXT_MUTED = "#64748b"   # Slate 500 (Subtext)
COLOR_SUCCESS = "#16a34a"      # Green 600 (Export button)
COLOR_DANGER = "#dc2626"       # Red 600 (Clear button)

class HoverButton(tk.Button):
    """Custom button with hover effect and modern flat style."""
    def __init__(self, master, active_bg=None, **kw):
        # Default flat button configuration
        kw.setdefault("relief", "flat")
        kw.setdefault("bd", 0)
        kw.setdefault("cursor", "hand2")
        kw.setdefault("font", ("Segoe UI", 10, "bold"))
        
        super().__init__(master, **kw)
        self.default_bg = self["background"]
        self.active_bg = active_bg if active_bg else self.default_bg
        
        self.bind("<Enter>", self.on_enter)
        self.bind("<Leave>", self.on_leave)

    def on_enter(self, e):
        self.config(background=self.active_bg)

    def on_leave(self, e):
        self.config(background=self.default_bg)

class StdoutRedirector:
    """Redirects stdout to a tkinter ScrolledText widget."""
    def __init__(self, text_widget):
        self.text_widget = text_widget

    def write(self, string):
        self.text_widget.configure(state='normal')
        self.text_widget.insert('end', string)
        self.text_widget.see('end')
        self.text_widget.configure(state='disabled')

    def flush(self):
        pass

class DPostConverterGUI:
    def __init__(self, root):
        self.root = root
        self.root.title(f"DPost PDF Converter v{__version__}")
        self.root.geometry("1000x720")
        self.root.minsize(900, 650)
        self.root.configure(bg=COLOR_BG)
        
        self.selected_files = []
        self.parsed_records = []
        self.dataframe = None

        # Setup Styles for ttk widgets (Treeview, Scrollbars)
        self.setup_styles()
        
        # Build UI layout
        self.create_layout()
        
        # Redirect stdout
        sys.stdout = StdoutRedirector(self.log_text)

    def setup_styles(self):
        style = ttk.Style()
        style.theme_use('clam')
        
        # Treeview styling
        style.configure('Treeview', 
                        background=COLOR_CARD, 
                        foreground=COLOR_TEXT, 
                        rowheight=26, 
                        fieldbackground=COLOR_CARD, 
                        font=('Segoe UI', 9))
        style.map('Treeview', 
                  background=[('selected', '#e2e8f0')], 
                  foreground=[('selected', COLOR_TEXT)])
        
        style.configure('Treeview.Heading', 
                        background='#e2e8f0', 
                        foreground=COLOR_TEXT, 
                        font=('Segoe UI', 9, 'bold'),
                        relief='flat')
        
        # Progressbar styling
        style.configure('TProgressbar', 
                        troughcolor='#e2e8f0', 
                        background=COLOR_SECONDARY, 
                        thickness=8)

    def create_layout(self):
        # 1. Header Banner
        header = tk.Frame(self.root, bg=COLOR_PRIMARY, height=75)
        header.pack(fill='x', side='top')
        header.pack_propagate(False)
        
        title_lbl = tk.Label(header, text="DPost Mailing Label PDF Converter", 
                             font=("Segoe UI", 16, "bold"), fg="#ffffff", bg=COLOR_PRIMARY)
        title_lbl.pack(anchor='w', padx=20, pady=(12, 0))
        
        version_lbl = tk.Label(header, text=f"เวอร์ชัน {__version__} | สำหรับสกัดข้อมูลใบนำส่งเพื่อนำเข้า DPost", 
                               font=("Segoe UI", 9), fg="#94a3b8", bg=COLOR_PRIMARY)
        version_lbl.pack(anchor='w', padx=20, pady=(0, 10))

        # Main Scrollable Container (just in case screen is too small)
        container = tk.Frame(self.root, bg=COLOR_BG)
        container.pack(fill='both', expand=True, padx=15, pady=15)
        
        # Grid Configuration for main container
        container.columnconfigure(0, weight=1)
        container.rowconfigure(0, weight=0) # File Selection Card
        container.rowconfigure(1, weight=3) # Preview Table Card
        container.rowconfigure(2, weight=2) # Log & Export Card

        # --- Card 1: File Selection & Conversion Controls ---
        card_files = tk.Frame(container, bg=COLOR_CARD, bd=1, relief='flat')
        card_files.grid(row=0, column=0, sticky='nsew', pady=(0, 10))
        
        # Layout inside Card 1
        tk.Label(card_files, text="1. เลือกแหล่งข้อมูลเอกสาร PDF", font=("Segoe UI", 11, "bold"), 
                 bg=COLOR_CARD, fg=COLOR_PRIMARY).pack(anchor='w', padx=15, pady=(10, 5))
        
        btn_frame = tk.Frame(card_files, bg=COLOR_CARD)
        btn_frame.pack(fill='x', padx=15, pady=(5, 10))
        
        self.btn_select_files = HoverButton(btn_frame, text=" เลือกไฟล์ PDF... ", bg=COLOR_ACCENT, fg="#ffffff", 
                                            active_bg="#0284c7", command=self.select_files, height=1, width=15)
        self.btn_select_files.pack(side='left', padx=(0, 10))
        
        self.btn_select_dir = HoverButton(btn_frame, text=" เลือกโฟลเดอร์... ", bg=COLOR_ACCENT, fg="#ffffff", 
                                          active_bg="#0284c7", command=self.select_directory, height=1, width=15)
        self.btn_select_dir.pack(side='left', padx=(0, 10))
        
        self.btn_clear = HoverButton(btn_frame, text="ล้างข้อมูล", bg="#e2e8f0", fg=COLOR_TEXT, 
                                     active_bg="#cbd5e1", command=self.clear_selection, height=1, width=10)
        self.btn_clear.pack(side='left', padx=(0, 20))
        
        self.lbl_status = tk.Label(btn_frame, text="ยังไม่ได้เลือกไฟล์", font=("Segoe UI", 10, "italic"), 
                                   bg=COLOR_CARD, fg=COLOR_TEXT_MUTED)
        self.lbl_status.pack(side='left', fill='x', expand=True, anchor='w')
        
        self.btn_convert = HoverButton(btn_frame, text=" เริ่มแปลงข้อมูล ", bg=COLOR_SECONDARY, fg="#ffffff", 
                                       active_bg="#0d9488", command=self.start_conversion, state='disabled', height=1, width=15)
        self.btn_convert.config(disabledforeground="#94a3b8")
        self.btn_convert.pack(side='right')

        # Progress bar
        self.progress = ttk.Progressbar(card_files, style='TProgressbar', mode='determinate')
        self.progress.pack(fill='x', padx=15, pady=(0, 5))

        # --- Card 2: Preview Table ---
        card_preview = tk.Frame(container, bg=COLOR_CARD, bd=1, relief='flat')
        card_preview.grid(row=1, column=0, sticky='nsew', pady=(0, 10))
        
        tk.Label(card_preview, text="2. ตารางตัวอย่างข้อมูลหลังสกัด (Preview)", font=("Segoe UI", 11, "bold"), 
                 bg=COLOR_CARD, fg=COLOR_PRIMARY).pack(anchor='w', padx=15, pady=(10, 5))
        
        table_frame = tk.Frame(card_preview, bg=COLOR_CARD)
        table_frame.pack(fill='both', expand=True, padx=15, pady=(0, 10))
        
        # Scrollbars for Treeview
        vsb = ttk.Scrollbar(table_frame, orient="vertical")
        hsb = ttk.Scrollbar(table_frame, orient="horizontal")
        
        # Columns
        self.preview_cols = ["NO", "REF NO", "SHIPPER NAME", "RECEIVER", "RECEIVER ADDRESS", "RECEIVER ZIPCODE"]
        col_widths = {"NO": 40, "REF NO": 120, "SHIPPER NAME": 200, "RECEIVER": 150, "RECEIVER ADDRESS": 300, "RECEIVER ZIPCODE": 90}
        col_titles = {"NO": "ลำดับ", "REF NO": "เลขที่อ้างอิง", "SHIPPER NAME": "ผู้ส่ง", "RECEIVER": "ผู้รับ", "RECEIVER ADDRESS": "ที่อยู่ผู้รับ", "RECEIVER ZIPCODE": "รหัสไปรษณีย์"}
        
        self.tree = ttk.Treeview(table_frame, columns=self.preview_cols, show="headings", 
                                 yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        
        vsb.config(command=self.tree.yview)
        hsb.config(command=self.tree.xview)
        
        for col in self.preview_cols:
            self.tree.heading(col, text=col_titles[col], anchor='w')
            self.tree.column(col, width=col_widths[col], minwidth=40, anchor='w')
            
        vsb.pack(side='right', fill='y')
        hsb.pack(side='bottom', fill='x')
        self.tree.pack(side='left', fill='both', expand=True)

        # --- Card 3: Log console & Save Action ---
        card_footer = tk.Frame(container, bg=COLOR_BG)
        card_footer.grid(row=2, column=0, sticky='nsew')
        card_footer.columnconfigure(0, weight=2) # Log console
        card_footer.columnconfigure(1, weight=1) # Export Panel
        card_footer.rowconfigure(0, weight=1)
        
        # Card 3a: Logs
        log_frame = tk.Frame(card_footer, bg=COLOR_CARD, bd=1, relief='flat')
        log_frame.grid(row=0, column=0, sticky='nsew', padx=(0, 10))
        
        tk.Label(log_frame, text="3. รายละเอียดการทำงาน (Log)", font=("Segoe UI", 11, "bold"), 
                 bg=COLOR_CARD, fg=COLOR_PRIMARY).pack(anchor='w', padx=15, pady=(10, 5))
        
        self.log_text = ScrolledText(log_frame, state='disabled', height=6, bg="#1e293b", fg="#f8fafc", 
                                     insertbackground="white", font=("Courier New", 9))
        self.log_text.pack(fill='both', expand=True, padx=15, pady=(0, 10))
        
        # Card 3b: Export panel
        export_frame = tk.Frame(card_footer, bg=COLOR_CARD, bd=1, relief='flat')
        export_frame.grid(row=0, column=1, sticky='nsew')
        
        tk.Label(export_frame, text="4. นำออกไฟล์ Excel", font=("Segoe UI", 11, "bold"), 
                 bg=COLOR_CARD, fg=COLOR_PRIMARY).pack(anchor='w', padx=15, pady=(10, 5))
        
        export_inner = tk.Frame(export_frame, bg=COLOR_CARD)
        export_inner.pack(fill='both', expand=True, padx=15, pady=10)
        
        self.btn_export = HoverButton(export_inner, text=" บันทึกไฟล์ Excel... ", bg=COLOR_SUCCESS, fg="#ffffff", 
                                      active_bg="#15803d", command=self.export_excel, state='disabled', height=2)
        self.btn_export.config(disabledforeground="#94a3b8")
        self.btn_export.pack(fill='x', pady=(15, 10))
        
        self.lbl_export_status = tk.Label(export_inner, text="กรุณาแปลงข้อมูลก่อนบันทึก", 
                                          font=("Segoe UI", 9, "italic"), bg=COLOR_CARD, fg=COLOR_TEXT_MUTED, wraplength=220)
        self.lbl_export_status.pack(fill='x', side='bottom', pady=10)

    # --- Command Event Handlers ---
    
    def select_files(self):
        files = filedialog.askopenfilenames(
            title="เลือกไฟล์ PDF ใบนำส่ง DPost",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")]
        )
        if files:
            self.selected_files = list(files)
            self.update_file_selection_status()

    def select_directory(self):
        directory = filedialog.askdirectory(title="เลือกโฟลเดอร์ที่เก็บไฟล์ PDF")
        if directory:
            # Find all PDFs in the selected directory
            files = glob.glob(os.path.join(directory, "*.pdf"))
            if files:
                self.selected_files = files
                self.update_file_selection_status()
            else:
                messagebox.showwarning("ไม่พบไฟล์", "ไม่พบไฟล์ PDF ในโฟลเดอร์ที่เลือก")

    def update_file_selection_status(self):
        count = len(self.selected_files)
        if count == 1:
            name = os.path.basename(self.selected_files[0])
            self.lbl_status.config(text=f"เลือกไฟล์: {name}", fg=COLOR_TEXT)
        else:
            self.lbl_status.config(text=f"เลือกไฟล์ทั้งหมด {count} ไฟล์", fg=COLOR_TEXT)
        
        self.btn_convert.config(state='normal')
        self.progress['value'] = 0
        self.btn_export.config(state='disabled')
        self.lbl_export_status.config(text="พร้อมเริ่มแปลงข้อมูล", fg=COLOR_TEXT_MUTED)

    def clear_selection(self):
        self.selected_files = []
        self.parsed_records = []
        self.dataframe = None
        self.lbl_status.config(text="ยังไม่ได้เลือกไฟล์", fg=COLOR_TEXT_MUTED)
        self.btn_convert.config(state='disabled')
        self.btn_export.config(state='disabled')
        self.lbl_export_status.config(text="กรุณาแปลงข้อมูลก่อนบันทึก", fg=COLOR_TEXT_MUTED)
        self.progress['value'] = 0
        
        # Clear preview table
        for item in self.tree.get_children():
            self.tree.delete(item)
            
        # Clear log text
        self.log_text.configure(state='normal')
        self.log_text.delete('1.0', tk.END)
        self.log_text.configure(state='disabled')

    def start_conversion(self):
        # Disable buttons during work
        self.btn_select_files.config(state='disabled')
        self.btn_select_dir.config(state='disabled')
        self.btn_convert.config(state='disabled')
        self.btn_clear.config(state='disabled')
        
        self.progress['value'] = 0
        
        # Clear preview table
        for item in self.tree.get_children():
            self.tree.delete(item)
            
        # Run conversion in background thread
        thread = threading.Thread(target=self.run_conversion_task)
        thread.daemon = True
        thread.start()

    def run_conversion_task(self):
        try:
            print(f"=== เริ่มการแปลงข้อมูล ({datetime.now().strftime('%H:%M:%S')}) ===")
            self.parsed_records = []
            
            total_files = len(self.selected_files)
            for index, file_path in enumerate(self.selected_files, 1):
                # Process individual PDF
                records = process_pdf(file_path)
                self.parsed_records.extend(records)
                
                # Update progress in GUI thread safely
                percent = int((index / total_files) * 100)
                self.root.after(0, self.update_progress, percent)
                
            print(f"สกัดข้อมูลเสร็จสิ้น ค้นพบข้อมูลผู้รับทั้งหมด {len(self.parsed_records)} รายการ\n")
            
            if self.parsed_records:
                self.dataframe = records_to_dataframe(self.parsed_records)
                self.root.after(0, self.conversion_success)
            else:
                self.root.after(0, self.conversion_failed, "ไม่พบข้อมูลใบนำส่งที่ถูกต้องในไฟล์ PDF ที่เลือก")
                
        except Exception as e:
            print(f"เกิดข้อผิดพลาด: {str(e)}\n")
            self.root.after(0, self.conversion_failed, str(e))

    def update_progress(self, val):
        self.progress['value'] = val

    def conversion_success(self):
        # Re-enable buttons
        self.btn_select_files.config(state='normal')
        self.btn_select_dir.config(state='normal')
        self.btn_convert.config(state='normal')
        self.btn_clear.config(state='normal')
        
        # Populate Treeview preview
        for idx, row in self.dataframe.iterrows():
            values = [
                row.get("NO", idx+1),
                row.get("REF NO", ""),
                row.get("SHIPPER NAME", ""),
                row.get("RECEIVER", ""),
                row.get("RECEIVER ADDRESS", ""),
                row.get("RECEIVER ZIPCODE", "")
            ]
            self.tree.insert("", "end", values=values)
            
        self.btn_export.config(state='normal')
        self.lbl_export_status.config(text=f"แปลงข้อมูลสำเร็จ ค้นพบทั้งหมด {len(self.dataframe)} รายการ พร้อมนำออกไฟล์", fg=COLOR_SECONDARY)
        messagebox.showinfo("เสร็จสิ้น", f"แปลงข้อมูลสำเร็จทั้งหมด {len(self.dataframe)} รายการ")

    def conversion_failed(self, error_msg):
        self.btn_select_files.config(state='normal')
        self.btn_select_dir.config(state='normal')
        self.btn_convert.config(state='normal')
        self.btn_clear.config(state='normal')
        
        self.btn_export.config(state='disabled')
        self.lbl_export_status.config(text="การแปลงข้อมูลล้มเหลว", fg=COLOR_DANGER)
        messagebox.showerror("เกิดข้อผิดพลาด", f"ไม่สามารถแปลงข้อมูลได้:\n{error_msg}")

    def export_excel(self):
        if self.dataframe is None or self.dataframe.empty:
            messagebox.showwarning("ไม่มีข้อมูล", "ไม่มีข้อมูลสำหรับส่งออก")
            return
            
        # Default filename with datetime suffix
        suffix = datetime.now().strftime("%Y%m%d%H%M")
        default_name = f"dpost_import_{suffix}.xlsx"
        
        output_path = filedialog.asksaveasfilename(
            title="บันทึกไฟล์ Excel สำหรับ DPost",
            initialfile=default_name,
            filetypes=[("Excel files", "*.xlsx")],
            defaultextension=".xlsx"
        )
        
        if output_path:
            try:
                # Save with sheet_name="New Order Data"
                self.dataframe.to_excel(output_path, sheet_name="New Order Data", index=False)
                print(f"บันทึกไฟล์ Excel สำเร็จ: {os.path.basename(output_path)}")
                self.lbl_export_status.config(text=f"บันทึกไฟล์เรียบร้อยแล้วที่:\n{os.path.basename(output_path)}", fg=COLOR_SUCCESS)
                messagebox.showinfo("บันทึกสำเร็จ", f"บันทึกไฟล์เรียบร้อยแล้วที่:\n{output_path}")
            except Exception as e:
                print(f"ไม่สามารถบันทึกไฟล์ได้: {str(e)}")
                messagebox.showerror("เกิดข้อผิดพลาด", f"ไม่สามารถบันทึกไฟล์ได้:\n{str(e)}")

def main():
    root = tk.Tk()
    
    # Set Windows window icon if exists (fallback to default)
    try:
        # Avoid issues if run on non-windows
        if os.name == 'nt':
            root.iconbitmap(default=None)
    except Exception:
        pass
        
    app = DPostConverterGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main()
