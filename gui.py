import os
import sys
import glob
import threading
from datetime import datetime
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import customtkinter as ctk
import pandas as pd

# Import conversion functions and version from convert_dpost
try:
    from convert_dpost import process_pdf, records_to_dataframe, __version__
except ImportError:
    __version__ = "2026.0630.1700"
    def process_pdf(path): return []
    def records_to_dataframe(records): return pd.DataFrame()

# Set ctk options
ctk.set_appearance_mode("dark")  # Modes: "System", "Dark", "Light"
ctk.set_default_color_theme("blue")  # Themes: "blue", "green", "dark-blue"

class StdoutRedirector:
    """Redirects stdout to a customtkinter CTkTextbox widget with pretty icon prefixing."""
    def __init__(self, text_widget):
        self.text_widget = text_widget

    def write(self, string):
        if not string.strip():
            return
            
        # Decorate logs dynamically for a premium developer feel
        decorated = string
        if "เริ่ม" in string:
            decorated = f"🚀 {string}"
        elif "เสร็จสิ้น" in string or "สำเร็จ" in string:
            decorated = f"✅ {string}"
        elif "เกิดข้อผิดพลาด" in string or "ล้มเหลว" in string:
            decorated = f"❌ {string}"
        elif "ประมวลผลไฟล์" in string:
            decorated = f"📂 {string}"
        else:
            decorated = f"ℹ️ {string}"
            
        self.text_widget.configure(state='normal')
        self.text_widget.insert('end', decorated + "\n")
        self.text_widget.see('end')
        self.text_widget.configure(state='disabled')

    def flush(self):
        pass

class DPostConverterGUI(ctk.CTk):
    def __init__(self):
        super().__init__()
        
        self.title(f"Convert PDF To Excel v{__version__}")
        self.geometry("1100x780")
        self.minsize(950, 720)
        
        self.selected_files = []
        self.parsed_records = []
        self.dataframe = None

        # Build UI layout
        self.create_layout()
        
        # Set initial active step
        self.set_current_step(1)
        
        # Redirect stdout
        sys.stdout = StdoutRedirector(self.log_text)
        
        # Setup keyboard shortcuts for advanced UX
        self.bind("<Control-o>", lambda e: self.select_files())
        self.bind("<Control-O>", lambda e: self.select_files())
        self.bind("<Control-s>", lambda e: self.export_excel_shortcut())
        self.bind("<Control-S>", lambda e: self.export_excel_shortcut())
        self.bind("<Escape>", self.on_escape_press)

    def create_layout(self):
        # 1. Header Banner
        header = ctk.CTkFrame(self, corner_radius=0, fg_color="#1e293b", height=65)
        header.pack(fill='x', side='top')
        header.pack_propagate(False)
        
        # Center container in header for nice padding
        header_content = ctk.CTkFrame(header, fg_color="transparent")
        header_content.pack(fill='both', expand=True, padx=25, pady=10)
        
        title_lbl = ctk.CTkLabel(header_content, text=f"Convert PDF To Excel v{__version__}", 
                                 font=("Segoe UI", 18, "bold"), text_color="#f8fafc")
        title_lbl.pack(anchor='w', side='left')
        
        # Theme toggle switch on the right side of header
        self.switch_theme = ctk.CTkSwitch(header_content, text="โหมดมืด (Dark Mode)", 
                                           font=("Segoe UI", 10, "bold"), text_color="#cbd5e1",
                                           command=self.toggle_theme)
        self.switch_theme.select() # Default to dark mode selected
        self.switch_theme.pack(anchor='e', side='right', padx=10)

        # 2. Sub-header Instruction Bar (Step-by-step guidance)
        instruction_bar = ctk.CTkFrame(self, corner_radius=8, border_width=1, border_color=("#cbd5e1", "#334155"))
        instruction_bar.pack(fill='x', padx=20, pady=(15, 0))
        
        center_frame = ctk.CTkFrame(instruction_bar, fg_color="transparent")
        center_frame.pack(anchor='center', pady=8)
        
        icon_lbl = ctk.CTkLabel(center_frame, text="💡 ขั้นตอนการทำงาน:", font=("Segoe UI", 11, "bold"), text_color="#38bdf8")
        icon_lbl.pack(side='left', padx=(0, 10))
        
        self.lbl_step1 = ctk.CTkLabel(center_frame, text=" [1] เลือกไฟล์ PDF (หรือโฟลเดอร์) ", font=("Segoe UI", 11, "bold"), text_color=("#64748b", "#94a3b8"), padx=8, pady=3)
        self.lbl_step1.pack(side='left')
        
        arrow1 = ctk.CTkLabel(center_frame, text=" ➔ ", font=("Segoe UI", 11), text_color="#64748b")
        arrow1.pack(side='left')
        
        self.lbl_step2 = ctk.CTkLabel(center_frame, text=" [2] เริ่มแปลงข้อมูล (ขวาบน) ", font=("Segoe UI", 11, "bold"), text_color=("#64748b", "#94a3b8"), padx=8, pady=3)
        self.lbl_step2.pack(side='left')
        
        arrow2 = ctk.CTkLabel(center_frame, text=" ➔ ", font=("Segoe UI", 11), text_color="#64748b")
        arrow2.pack(side='left')
        
        self.lbl_step3 = ctk.CTkLabel(center_frame, text=" [3] บันทึกไฟล์ Excel... (ขวาล่าง) ", font=("Segoe UI", 11, "bold"), text_color=("#64748b", "#94a3b8"), padx=8, pady=3)
        self.lbl_step3.pack(side='left')

        # 3. Main Container
        container = ctk.CTkFrame(self, fg_color="transparent")
        container.pack(fill='both', expand=True, padx=20, pady=15)
        
        # Grid config
        container.columnconfigure(0, weight=1)
        container.rowconfigure(0, weight=0) # File Selection Card
        container.rowconfigure(1, weight=3) # Preview Table Card
        container.rowconfigure(2, weight=2) # Log & Export Card

        # --- Card 1: File Selection & Controls ---
        card_files = ctk.CTkFrame(container, corner_radius=10, border_width=1, border_color=("#cbd5e1", "#334155"))
        card_files.grid(row=0, column=0, sticky='nsew', pady=(0, 10))
        
        ctk.CTkLabel(card_files, text="1. เลือกแหล่งข้อมูลเอกสาร PDF", font=("Segoe UI", 12, "bold"), 
                     text_color=("#0f172a", "#f8fafc")).pack(anchor='w', padx=20, pady=(12, 5))
        
        btn_frame = ctk.CTkFrame(card_files, fg_color="transparent")
        btn_frame.pack(fill='x', padx=20, pady=(5, 10))
        
        self.btn_select_files = ctk.CTkButton(btn_frame, text=" 📄 เลือกไฟล์ PDF... ", fg_color="#0284c7", hover_color="#0369a1",
                                              font=("Segoe UI", 11, "bold"), command=self.select_files, width=170)
        self.btn_select_files.pack(side='left', padx=(0, 10))
        
        self.btn_select_dir = ctk.CTkButton(btn_frame, text=" 📁 เลือกโฟลเดอร์... ", fg_color="#0284c7", hover_color="#0369a1",
                                            font=("Segoe UI", 11, "bold"), command=self.select_directory, width=170)
        self.btn_select_dir.pack(side='left', padx=(0, 10))
        
        self.btn_clear = ctk.CTkButton(btn_frame, text=" 🧹 ล้างข้อมูล ", fg_color="#475569", hover_color="#334155",
                                       font=("Segoe UI", 11, "bold"), command=self.clear_selection, width=110)
        self.btn_clear.pack(side='left', padx=(0, 20))
        
        self.lbl_status = ctk.CTkLabel(btn_frame, text="ยังไม่ได้เลือกไฟล์", font=("Segoe UI", 11, "italic"), text_color=("#475569", "#94a3b8"))
        self.lbl_status.pack(side='left', fill='x', expand=True, anchor='w')
        
        self.btn_convert = ctk.CTkButton(btn_frame, text=" ⚡ เริ่มแปลงข้อมูล ", fg_color="#0f766e", hover_color="#0d9488",
                                         font=("Segoe UI", 11, "bold"), command=self.start_conversion, state='disabled', width=170)
        self.btn_convert.pack(side='right')

        # Progress bar
        self.progress = ctk.CTkProgressBar(card_files, progress_color="#0f766e", height=8)
        self.progress.pack(fill='x', padx=20, pady=(0, 12))
        self.progress.set(0)

        # --- Card 2: Preview Table ---
        card_preview = ctk.CTkFrame(container, corner_radius=10, border_width=1, border_color=("#cbd5e1", "#334155"))
        card_preview.grid(row=1, column=0, sticky='nsew', pady=(0, 10))
        
        # Header frame with real-time search box
        preview_header = ctk.CTkFrame(card_preview, fg_color="transparent")
        preview_header.pack(fill='x', padx=20, pady=(10, 5))
        
        title_lbl = ctk.CTkLabel(preview_header, text="2. ตารางตัวอย่างข้อมูลหลังสกัด (Preview)", font=("Segoe UI", 12, "bold"), 
                                 text_color=("#0f172a", "#f8fafc"))
        title_lbl.pack(side='left')
        
        self.search_entry = ctk.CTkEntry(preview_header, placeholder_text=" 🔍 ค้นหาผู้รับ / ผู้ส่ง / เลขอ้างอิง... ", 
                                         width=300, height=28, font=("Segoe UI", 11))
        self.search_entry.pack(side='right')
        self.search_entry.bind("<KeyRelease>", self.filter_treeview)
        self.search_entry.bind("<Escape>", self.clear_search)

        table_frame = ctk.CTkFrame(card_preview, fg_color="transparent")
        table_frame.pack(fill='both', expand=True, padx=20, pady=(0, 15))
        
        # Scrollbars for Treeview
        vsb = ttk.Scrollbar(table_frame, orient="vertical")
        hsb = ttk.Scrollbar(table_frame, orient="horizontal")
        
        # Setup Styles for treeview
        self.preview_cols = ["NO", "INV_NO", "SHIPPER_NAME", "RECEIVER", "RECEIVER_ADDRESS", "RECEIVER_ZIPCODE"]
        col_widths = {"NO": 50, "INV_NO": 130, "SHIPPER_NAME": 200, "RECEIVER": 160, "RECEIVER_ADDRESS": 320, "RECEIVER_ZIPCODE": 100}
        col_titles = {"NO": "ลำดับ", "INV_NO": "เลขที่อ้างอิง", "SHIPPER_NAME": "ผู้ส่ง", "RECEIVER": "ผู้รับ", "RECEIVER_ADDRESS": "ที่อยู่ผู้รับ", "RECEIVER_ZIPCODE": "รหัสไปรษณีย์"}
        
        self.tree = ttk.Treeview(table_frame, columns=self.preview_cols, show="headings", 
                                 yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        
        vsb.config(command=self.tree.yview)
        hsb.config(command=self.tree.xview)
        
        for col in self.preview_cols:
            self.tree.heading(col, text=col_titles[col], anchor='w')
            self.tree.column(col, width=col_widths[col], minwidth=50, anchor='w')
            
        vsb.pack(side='right', fill='y')
        hsb.pack(side='bottom', fill='x')
        self.tree.pack(side='left', fill='both', expand=True)

        # Style Treeview initially for dark mode
        self.style_treeview("dark")

        # --- Card 3: Log console & Save Action ---
        card_footer = ctk.CTkFrame(container, fg_color="transparent")
        card_footer.grid(row=2, column=0, sticky='nsew')
        card_footer.columnconfigure(0, weight=2) # Log console
        card_footer.columnconfigure(1, weight=1) # Export Panel
        card_footer.rowconfigure(0, weight=1)
        
        # Card 3a: Logs
        log_frame = ctk.CTkFrame(card_footer, corner_radius=10, border_width=1, border_color=("#cbd5e1", "#334155"))
        log_frame.grid(row=0, column=0, sticky='nsew', padx=(0, 10))
        
        ctk.CTkLabel(log_frame, text="3. รายละเอียดการทำงาน (Log)", font=("Segoe UI", 12, "bold"), 
                     text_color=("#0f172a", "#f8fafc")).pack(anchor='w', padx=20, pady=(12, 5))
        
        self.log_text = ctk.CTkTextbox(log_frame, state='disabled', fg_color=("#f1f5f9", "#0f172a"), text_color=("#0f172a", "#38bdf8"), 
                                       font=("Consolas", 11), corner_radius=6)
        self.log_text.pack(fill='both', expand=True, padx=20, pady=(0, 15))
        
        # Card 3b: Export panel
        export_frame = ctk.CTkFrame(card_footer, corner_radius=10, border_width=1, border_color=("#cbd5e1", "#334155"))
        export_frame.grid(row=0, column=1, sticky='nsew')
        
        ctk.CTkLabel(export_frame, text="4. นำออกไฟล์ Excel", font=("Segoe UI", 12, "bold"), 
                     text_color=("#0f172a", "#f8fafc")).pack(anchor='w', padx=20, pady=(12, 5))
        
        export_inner = ctk.CTkFrame(export_frame, fg_color="transparent")
        export_inner.pack(fill='both', expand=True, padx=20, pady=10)
        
        # Stats summary dashboard frame
        self.stats_frame = ctk.CTkFrame(export_inner, fg_color=("#f1f5f9", "#1e293b"), corner_radius=6)
        self.stats_frame.pack(fill='x', pady=(0, 10))
        
        self.lbl_stat_files = ctk.CTkLabel(self.stats_frame, text="📂 ไฟล์ PDF: 0 ไฟล์", font=("Segoe UI", 11, "bold"), text_color=("#475569", "#cbd5e1"))
        self.lbl_stat_files.pack(anchor='w', padx=15, pady=(8, 3))
        
        self.lbl_stat_records = ctk.CTkLabel(self.stats_frame, text="👥 รายการผู้รับ: 0 รายการ", font=("Segoe UI", 11, "bold"), text_color=("#475569", "#cbd5e1"))
        self.lbl_stat_records.pack(anchor='w', padx=15, pady=3)
        
        self.lbl_stat_prov = ctk.CTkLabel(self.stats_frame, text="📍 ส่งบ่อยสุด: -", font=("Segoe UI", 11, "bold"), text_color=("#475569", "#cbd5e1"))
        self.lbl_stat_prov.pack(anchor='w', padx=15, pady=(3, 8))
        
        self.btn_export = ctk.CTkButton(export_inner, text=" 📥 บันทึกไฟล์ Excel... ", fg_color="#16a34a", hover_color="#15803d",
                                        font=("Segoe UI", 12, "bold"), command=self.export_excel, state='disabled', height=45)
        self.btn_export.pack(fill='x', pady=(5, 5))
        
        self.lbl_export_status = ctk.CTkLabel(export_inner, text="กรุณาแปลงข้อมูลก่อนบันทึก", 
                                              font=("Segoe UI", 11, "italic"), text_color=("#475569", "#94a3b8"))
        self.lbl_export_status.pack(fill='x', side='bottom', pady=5)

    def style_treeview(self, mode):
        """Dynamic styling for the standard ttk.Treeview depending on the active theme."""
        style = ttk.Style()
        style.theme_use('clam')
        
        if mode == "dark":
            bg = "#1e293b"          # Dark slate
            fg = "#f8fafc"          # Light text
            headings_bg = "#334155" # Slate 700
            selected_bg = "#475569" # Slate 600
        else:
            bg = "#ffffff"          # White
            fg = "#0f172a"          # Dark slate text
            headings_bg = "#cbd5e1" # Slate 200
            selected_bg = "#94a3b8" # Slate 400
            
        style.configure('Treeview', 
                        background=bg, 
                        foreground=fg, 
                        rowheight=26, 
                        fieldbackground=bg, 
                        font=('Segoe UI', 9))
        style.map('Treeview', 
                  background=[('selected', selected_bg)], 
                  foreground=[('selected', '#ffffff')])
        
        style.configure('Treeview.Heading', 
                        background=headings_bg, 
                        foreground=fg, 
                        font=('Segoe UI', 9, 'bold'),
                        relief='flat')

    def set_current_step(self, step):
        """Highlights the active step label and dims inactive steps."""
        for s_num, lbl in {1: self.lbl_step1, 2: self.lbl_step2, 3: self.lbl_step3}.items():
            if s_num == step:
                # Active style: Light blue capsule background, dark text
                lbl.configure(text_color="#0f172a", fg_color="#38bdf8", corner_radius=6)
            else:
                # Inactive style: Muted gray text, transparent background
                lbl.configure(text_color=("#64748b", "#94a3b8"), fg_color="transparent")

    def toggle_theme(self):
        """Toggles between Dark and Light appearance modes."""
        if self.switch_theme.get() == 1:
            ctk.set_appearance_mode("dark")
            self.style_treeview("dark")
            self.switch_theme.configure(text="โหมดมืด (Dark Mode)")
        else:
            ctk.set_appearance_mode("light")
            self.style_treeview("light")
            self.switch_theme.configure(text="โหมดสว่าง (Light Mode)")

    # --- UX Dynamic Helpers ---
    
    def update_stats(self):
        """Updates the mini-dashboard stats panel."""
        file_count = len(self.selected_files)
        rec_count = len(self.parsed_records)
        
        top_prov = "-"
        if self.dataframe is not None and not self.dataframe.empty:
            prov_series = self.dataframe['RECEIVER_PROVINCE'].dropna()
            prov_series = prov_series[prov_series != ""]
            if not prov_series.empty:
                top_prov = prov_series.mode().iloc[0]
                
        self.lbl_stat_files.configure(text=f"📂 ไฟล์ PDF: {file_count} ไฟล์")
        self.lbl_stat_records.configure(text=f"👥 รายการผู้รับ: {rec_count} รายการ")
        self.lbl_stat_prov.configure(text=f"📍 ส่งบ่อยสุด: {top_prov}")

    def filter_treeview(self, event=None):
        """Filters the Treeview preview list in real-time based on the search query."""
        query = self.search_entry.get().strip().lower()
        
        # Clear current rows
        for item in self.tree.get_children():
            self.tree.delete(item)
            
        if self.dataframe is None or self.dataframe.empty:
            return
            
        for idx, row in self.dataframe.iterrows():
            inv_no = str(row.get("INV_NO", "")).lower()
            shipper = str(row.get("SHIPPER_NAME", "")).lower()
            receiver = str(row.get("RECEIVER", "")).lower()
            address = str(row.get("RECEIVER_ADDRESS", "")).lower()
            zipcode = str(row.get("RECEIVER_ZIPCODE", "")).lower()
            
            if (not query) or (query in inv_no or query in shipper or query in receiver or query in address or query in zipcode):
                values = [
                    row.get("NO", idx+1),
                    row.get("INV_NO", ""),
                    row.get("SHIPPER_NAME", ""),
                    row.get("RECEIVER", ""),
                    row.get("RECEIVER_ADDRESS", ""),
                    row.get("RECEIVER_ZIPCODE", "")
                ]
                self.tree.insert("", "end", values=values)

    def clear_search(self, event=None):
        """Clears the search entry content and resets view."""
        self.search_entry.delete(0, tk.END)
        self.filter_treeview()

    def export_excel_shortcut(self):
        """Shortcut triggered via Ctrl+S."""
        if self.dataframe is not None and not self.dataframe.empty:
            self.export_excel()

    def on_escape_press(self, event=None):
        """Clears search if active, otherwise clears the workspace."""
        if self.focus_get() == self.search_entry:
            self.clear_search()
            self.focus_set() # Unfocus search entry
        else:
            self.clear_selection()

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
            self.lbl_status.configure(text=f"เลือกไฟล์: {name}", text_color=("#0284c7", "#38bdf8"))
        else:
            self.lbl_status.configure(text=f"เลือกไฟล์ทั้งหมด {count} ไฟล์", text_color=("#0284c7", "#38bdf8"))
        
        self.btn_convert.configure(state='normal')
        self.progress.set(0)
        self.btn_export.configure(state='disabled')
        self.lbl_export_status.configure(text="พร้อมเริ่มแปลงข้อมูล", text_color=("#475569", "#94a3b8"))
        
        # Update dynamic stats panel
        self.update_stats()
        
        # Transition to step 2 (ready to convert)
        self.set_current_step(2)

    def clear_selection(self):
        self.selected_files = []
        self.parsed_records = []
        self.dataframe = None
        self.lbl_status.configure(text="ยังไม่ได้เลือกไฟล์", text_color="#94a3b8")
        self.btn_convert.configure(state='disabled')
        self.btn_export.configure(state='disabled')
        self.lbl_export_status.configure(text="กรุณาแปลงข้อมูลก่อนบันทึก", text_color=("#475569", "#94a3b8"))
        self.progress.set(0)
        
        # Clear search box
        self.search_entry.delete(0, tk.END)
        
        # Reset stats
        self.update_stats()
        
        # Clear preview table
        for item in self.tree.get_children():
            self.tree.delete(item)
            
        # Clear log text
        self.log_text.configure(state='normal')
        self.log_text.delete('1.0', tk.END)
        self.log_text.configure(state='disabled')
        
        # Reset to step 1
        self.set_current_step(1)

    def start_conversion(self):
        # Disable buttons during work
        self.btn_select_files.configure(state='disabled')
        self.btn_select_dir.configure(state='disabled')
        self.btn_convert.configure(state='disabled')
        self.btn_clear.configure(state='disabled')
        
        self.progress.set(0)
        
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
                records = process_pdf(file_path)
                self.parsed_records.extend(records)
                
                percent = float(index / total_files)
                self.after(0, self.update_progress, percent)
                
            print(f"สกัดข้อมูลเสร็จสิ้น ค้นพบข้อมูลผู้รับทั้งหมด {len(self.parsed_records)} รายการ\n")
            
            if self.parsed_records:
                self.dataframe = records_to_dataframe(self.parsed_records)
                self.after(0, self.conversion_success)
            else:
                self.after(0, self.conversion_failed, "ไม่พบข้อมูลใบนำส่งที่ถูกต้องในไฟล์ PDF ที่เลือก")
                
        except Exception as e:
            print(f"เกิดข้อผิดพลาด: {str(e)}\n")
            self.after(0, self.conversion_failed, str(e))

    def update_progress(self, val):
        self.progress.set(val)

    def conversion_success(self):
        # Re-enable buttons
        self.btn_select_files.configure(state='normal')
        self.btn_select_dir.configure(state='normal')
        self.btn_convert.configure(state='normal')
        self.btn_clear.configure(state='normal')
        
        # Refresh Treeview preview
        self.filter_treeview()
            
        self.btn_export.configure(state='normal')
        self.lbl_export_status.configure(text=f"แปลงข้อมูลสำเร็จ ค้นพบทั้งหมด {len(self.dataframe)} รายการ พร้อมนำออกไฟล์", text_color="#16a34a")
        
        # Update dynamic stats panel
        self.update_stats()
        
        messagebox.showinfo("เสร็จสิ้น", f"แปลงข้อมูลสำเร็จทั้งหมด {len(self.dataframe)} รายการ")
        
        # Transition to step 3 (ready to export)
        self.set_current_step(3)

    def conversion_failed(self, error_msg):
        self.btn_select_files.configure(state='normal')
        self.btn_select_dir.configure(state='normal')
        self.btn_convert.configure(state='normal')
        self.btn_clear.configure(state='normal')
        
        self.btn_export.configure(state='disabled')
        self.lbl_export_status.configure(text="การแปลงข้อมูลล้มเหลว", text_color="#dc2626")
        
        # Reset stats
        self.update_stats()
        
        messagebox.showerror("เกิดข้อผิดพลาด", f"ไม่สามารถแปลงข้อมูลได้:\n{error_msg}")

    def export_excel(self):
        if self.dataframe is None or self.dataframe.empty:
            messagebox.showwarning("ไม่มีข้อมูล", "ไม่มีข้อมูลสำหรับส่งออก")
            return
            
        suffix = datetime.now().strftime("%Y%m%d%H%M")
        default_name = f"dpost_import_{suffix}.xlsx"
        
        downloads_dir = os.path.join(os.path.expanduser("~"), "Downloads")
        if not os.path.exists(downloads_dir):
            downloads_dir = os.path.expanduser("~")
            
        output_path = filedialog.asksaveasfilename(
            title="บันทึกไฟล์ Excel สำหรับ DPost",
            initialdir=downloads_dir,
            initialfile=default_name,
            filetypes=[("Excel files", "*.xlsx")],
            defaultextension=".xlsx"
        )
        
        if output_path:
            try:
                self.dataframe.to_excel(output_path, sheet_name="New Order Data", index=False)
                print(f"บันทึกไฟล์ Excel สำเร็จ: {os.path.basename(output_path)}")
                messagebox.showinfo("บันทึกสำเร็จ", f"บันทึกไฟล์เรียบร้อยแล้วที่:\n{output_path}")
                self.clear_selection()
            except Exception as e:
                print(f"ไม่สามารถบันทึกไฟล์ได้: {str(e)}")
                messagebox.showerror("เกิดข้อผิดพลาด", f"ไม่สามารถบันทึกไฟล์ได้:\n{str(e)}")

def main():
    app = DPostConverterGUI()
    app.mainloop()

if __name__ == "__main__":
    main()
