import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from excel_processor import ExcelProcessor
from word_generator import WordGenerator
import os
import threading
import pandas as pd
import re
import json
import subprocess
import requests
import webbrowser

class StudentApp:
    def __init__(self, root):
        self.root = root
        
        # Load Local Version
        self.version = "3.0.0"
        """
        if os.path.exists("ver.txt"):
            try:
                with open("ver.txt", "r") as f:
                    self.version = f.read().strip()
            except: pass
        """

        self.root.title(f"مستخرج علامات الطلاب - v{self.version}")
        self.root.geometry("1250x800")
        
        # Update Check
        self.update_url = "https://drive.google.com/uc?export=download&id=1Hjzt3OGXYG6dppIjJLvThPWpRWuyXAa5"
        if not self.check_for_updates():
            self.root.destroy()
            return

        self.processor = ExcelProcessor()
        self.all_classes = []
        self.select_all_state = False
        
        # Trial/Activation Configuration
        self.is_activated = False
        self.license_info = {}
        self.total_trial_usage = 0
        self.trial_max_total = 30
        self.trial_per_file_limit = 5
        self.hwid = self.get_hwid()
        # IMPORTANT: Replace with your actual Render URL
        self.render_base_url = "https://shehadaajori2026-2.onrender.com" 

        # Load Config
        self.config_path = "config.json"
        self.load_config()

        # Styles for larger fonts
        self.label_font = ("Arial", 14, "bold")
        self.entry_font = ("Arial", 14)
        self.table_font = ("Arial", 13)
        
        style = ttk.Style()
        style.configure("Treeview", font=self.table_font, rowheight=35)
        style.configure("Treeview.Heading", font=("Arial", 14, "bold"))
        style.map("Treeview", background=[('selected', '#347083')])
        style.configure("TProgressbar", thickness=25)

        self.setup_ui()
        self.setup_menu()
        
        # Set default values after UI is setup
        self.school_entries['actual_days'].insert(0, self.default_actual_days)
        
        # Check license status on startup
        threading.Thread(target=self.check_activation_on_startup, daemon=True).start()
        # Start connection heartbeat
        threading.Thread(target=self.connection_heartbeat, daemon=True).start()

    def connection_heartbeat(self):
        while True:
            try:
                # Use a simple health check or the status endpoint
                response = requests.get(f"{self.render_base_url}/trial_status/health", timeout=45)
                status = "connected" if response.status_code in [200, 404] else "error"
            except:
                status = "disconnected"
            
            self.root.after(0, lambda s=status: self.update_connection_ui(s))
            import time
            time.sleep(30) # Check every 30 seconds

    def update_connection_ui(self, status):
        if not hasattr(self, 'lbl_conn'):
            self.lbl_conn = tk.Label(self.status_frame, text="● جاري الاتصال...", font=("Arial", 10, "bold"), bg="#f8f9fa")
            self.lbl_conn.pack(side=tk.LEFT, padx=20)
        
        if status == "connected":
            self.lbl_conn.config(text="● متصل بالسيرفر", fg="#27ae60")
        else:
            self.lbl_conn.config(text="○ غير متصل", fg="#e74c3c")

    def check_for_updates(self):
        try:
            # Try to fetch version from Google Drive (Direct Download Link)
            response = requests.get(self.update_url, timeout=45)
            if response.status_code == 200:
                content = response.text.strip().split('\n')
                remote_version = content[0].strip()
                download_link = content[1].strip() if len(content) > 1 else ""
                
                # Check if remote version is newer
                if remote_version > self.version:
                    msg = f"يتوفر إصدار جديد ({remote_version}). الإصدار الحالي ({self.version}).\n\nهل ترغب في فتح رابط التحميل الآن؟\n\n*ملاحظة: البرنامج لن يعمل بدون تحديث."
                    ans = messagebox.askyesno("تحديث جديد متوفر", msg)
                    if ans:
                        if download_link:
                            webbrowser.open(download_link)
                        else:
                            messagebox.showinfo("تنبيه", "يرجى تحميل النسخة الجديدة من الرابط المعتمد لإكمال العمل.")
                    return False # Mandatory update: stop execution regardless
            return True # No update or server error
        except:
            # If no internet or drive is down, allow usage for now
            return True

    def get_hwid(self):
        try:
            cmd = "wmic csproduct get uuid"
            uuid = subprocess.check_output(cmd, shell=True).decode().split('\n')[1].strip()
            return uuid
        except:
            try: return os.popen("vol C:").read().split()[-1]
            except: return "UNKNOWN_DEVICE"

    def check_activation_on_startup(self):
        # Always check server for activation status by HWID
        try:
            response = requests.get(f"{self.render_base_url}/status/{self.hwid}", timeout=45)
            if response.status_code == 200:
                data = response.json()
                if data.get('status') == 'active':
                    self.is_activated = True
                    self.license_info = data
                    self.save_config_to_file()
                else:
                    self.is_activated = False
            elif response.status_code == 404:
                try:
                    trial_resp = requests.get(f"{self.render_base_url}/trial_status/{self.hwid}", timeout=45)
                    if trial_resp.status_code == 200:
                        t_data = trial_resp.json()
                        server_used = t_data.get('used', 0)
                        if server_used > self.total_trial_usage:
                            self.total_trial_usage = server_used
                            self.save_config_to_file()
                except: pass
                if self.is_activated:
                    self.is_activated = False
                    self.save_config_to_file()
        except: pass
        self.root.after(0, self.update_trial_status_ui)

    def load_config(self):
        default_mapping = {
            'الأساسية': ["الأول", "الثاني", "الثالث", "الرابع", "الخامس", "السادس", "السابع", "الثامن", "التاسع"],
            'الأكاديمية': ["العاشر", "الحادي عشر", "الثاني عشر"]
        }
        if os.path.exists(self.config_path):
            try:
                with open(self.config_path, 'r', encoding='utf-8') as f:
                    config = json.load(f)
                    self.stage_mapping = config.get('stage_mapping', default_mapping)
                    self.default_actual_days = config.get('default_actual_days', "180")
                    self.is_activated = config.get('is_activated', False)
                    self.total_trial_usage = config.get('total_trial_usage', 0)
            except:
                self.stage_mapping = default_mapping
                self.default_actual_days = "180"
        else:
            self.stage_mapping = default_mapping
            self.default_actual_days = "180"

    def save_config_to_file(self):
        config = {
            'stage_mapping': self.stage_mapping,
            'default_actual_days': self.default_actual_days,
            'is_activated': self.is_activated,
            'total_trial_usage': self.total_trial_usage
        }
        with open(self.config_path, 'w', encoding='utf-8') as f:
            json.dump(config, f, ensure_ascii=False, indent=4)

    def setup_menu(self):
        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)
        
        settings_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="إعدادات", menu=settings_menu)
        settings_menu.add_command(label="الإعدادات العامة", command=self.open_settings)
        settings_menu.add_command(label="تفعيل البرنامج", command=self.open_activation_window)

    def open_activation_window(self):
        act_win = tk.Toplevel(self.root)
        act_win.title("تفعيل البرنامج")
        act_win.geometry("500x550")
        
        main_frame = tk.Frame(act_win, padx=20, pady=20)
        main_frame.pack(fill=tk.BOTH, expand=True)

        tk.Label(main_frame, text="رقم الجهاز المميز:", font=("Arial", 10, "bold")).pack(pady=(10,0))
        hwid_ent = tk.Entry(main_frame, font=("Arial", 10), justify="center", width=40)
        hwid_ent.insert(0, self.hwid)
        hwid_ent.config(state='readonly')
        hwid_ent.pack(pady=5)

        tk.Label(main_frame, text="الاسم المسجل:", font=("Arial", 10, "bold")).pack(pady=(10,0))
        name_ent = tk.Entry(main_frame, font=("Arial", 11), justify="center", width=40)
        if self.is_activated:
            name_ent.insert(0, self.license_info.get('Name', ''))
            name_ent.config(state='readonly')
        name_ent.pack(pady=5)

        tk.Label(main_frame, text="كود التفعيل:", font=("Arial", 10, "bold")).pack(pady=(10,0))
        code_ent = tk.Entry(main_frame, font=("Arial", 12, "bold"), justify="center", width=25, fg="blue")
        if self.is_activated:
            # Show actual code if available, or placeholder
            code_ent.insert(0, self.license_info.get('Code', '******'))
            code_ent.config(state='readonly')
        code_ent.pack(pady=5)

        def attempt_activate():
            if self.is_activated:
                messagebox.showinfo("تنبيه", "البرنامج مفعل بالفعل ولا يمكن تغيير الكود.")
                return

            name = name_ent.get().strip()
            code = code_ent.get().strip()
            if not name or not code:
                messagebox.showwarning("تنبيه", "يرجى إدخال الاسم وكود التفعيل")
                return
            
            try:
                payload = {'name': name, 'code': code, 'hwid': self.hwid}
                response = requests.post(f"{self.render_base_url}/activate", json=payload, timeout=15)
                
                if response.status_code == 200:
                    data = response.json()
                    # Store the code locally too for display
                    data['Code'] = code
                    self.is_activated = True
                    self.license_info = data
                    self.save_config_to_file()
                    messagebox.showinfo("نجاح", f"تم تفعيل البرنامج بنجاح!\nالمشترك: {name}")
                    act_win.destroy()
                    self.update_trial_status_ui()
                else:
                    err = response.json().get('detail', 'كود غير صحيح')
                    messagebox.showerror("خطأ", f"فشل التفعيل: {err}")
            except Exception as e:
                messagebox.showerror("خطأ", f"فشل الاتصال بخادم التفعيل: {str(e)}")

        btn_text = "البرنامج مفعل ✅" if self.is_activated else "تفعيل الآن 🚀"
        btn_state = tk.DISABLED if self.is_activated else tk.NORMAL
        tk.Button(main_frame, text=btn_text, command=attempt_activate, bg="#4CAF50", fg="white", font=self.label_font, state=btn_state).pack(pady=20)
        
        # Support Info
        support_frame = tk.Frame(main_frame, pady=10)
        support_frame.pack(fill=tk.X)
        
        tk.Label(support_frame, text="للحصول على كود التفعيل أو الدعم الفني:", font=("Arial", 10, "bold")).pack()
        tk.Label(support_frame, text="📞 هاتف: 0781155386", font=("Arial", 11)).pack()
        tk.Label(support_frame, text="💬 واتساب: 0777638137", font=("Arial", 11)).pack()
        
        tk.Label(main_frame, text="إعداد وتطوير: محمود العجوري", font=("Arial", 10, "italic"), fg="#555").pack(pady=(20, 0))

    def update_trial_status_ui(self):
        if self.is_activated:
            name = self.license_info.get('Name', '')
            used = self.license_info.get('Used', 0)
            max_val = self.license_info.get('Max', 0)
            
            self.root.title(f"مستخرج علامات الطلاب - [{name}] - إعداد محمود العجوري (0781155386)")
            self.btn_word.config(state=tk.NORMAL)
            
            msg = f"الرصيد: {used} / {max_val} شهادة"
            if not hasattr(self, 'lbl_usage'):
                self.lbl_usage = tk.Label(self.status_frame, text=msg, font=("Arial", 10, "bold"), fg="#2c3e50", bg="#f8f9fa")
                self.lbl_usage.pack(side=tk.RIGHT, padx=20)
            else:
                self.lbl_usage.config(text=msg, fg="#2c3e50")
            
            if hasattr(self, 'lbl_trial'): self.lbl_trial.pack_forget()
        else:
            self.root.title("مستخرج علامات الطلاب - [الوضع التجريبي] - إعداد محمود العجوري (0781155386)")
            self.btn_word.config(state=tk.DISABLED)
            rem = max(0, self.trial_max_total - self.total_trial_usage)
            msg = f"الوضع التجريبي: متبقي {rem} شهادة"
            
            if hasattr(self, 'lbl_usage'): self.lbl_usage.pack_forget()
            
            if not hasattr(self, 'lbl_trial'):
                self.lbl_trial = tk.Label(self.status_frame, text=msg, font=("Arial", 10, "bold"), fg="red", bg="#f8f9fa")
                self.lbl_trial.pack(side=tk.RIGHT, padx=20)
            else:
                self.lbl_trial.config(text=msg)

    def open_settings(self):
        settings_win = tk.Toplevel(self.root)
        settings_win.title("الإعدادات")
        settings_win.geometry("600x600")
        
        main_frame = tk.Frame(settings_win)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)

        # 1. Default Days
        days_frame = tk.LabelFrame(main_frame, text="الإعدادات العامة", font=self.label_font, labelanchor="ne")
        days_frame.pack(fill=tk.X, pady=10)
        
        tk.Label(days_frame, text="أيام الدوام الافتراضية:", font=("Arial", 12)).pack(side=tk.RIGHT, padx=5, pady=10)
        days_ent = tk.Entry(days_frame, font=("Arial", 12), width=10, justify="center")
        days_ent.insert(0, self.default_actual_days)
        days_ent.pack(side=tk.RIGHT, padx=5, pady=10)

        # 2. Stage Mapping
        tk.Label(main_frame, text="تخصيص المراحل والصفوف التابعة لها", font=self.label_font).pack(pady=10)
        
        container = tk.Frame(main_frame)
        container.pack(fill=tk.BOTH, expand=True)

        rows = []
        
        def add_row_ui(stage_name="", grades_str=""):
            row_frame = tk.Frame(container)
            row_frame.pack(fill=tk.X, pady=2)
            
            name_ent = tk.Entry(row_frame, font=("Arial", 11), width=15, justify="right")
            name_ent.insert(0, stage_name)
            name_ent.pack(side=tk.RIGHT, padx=5)
            
            grades_ent = tk.Entry(row_frame, font=("Arial", 11), justify="right")
            grades_ent.insert(0, grades_str)
            grades_ent.pack(side=tk.RIGHT, fill=tk.X, expand=True, padx=5)
            
            def delete_row():
                row_frame.destroy()
                rows.remove((name_ent, grades_ent))
                
            tk.Button(row_frame, text="حذف", command=delete_row, bg="#ffcccc", fg="red").pack(side=tk.LEFT)
            rows.append((name_ent, grades_ent))

        # Initial Load
        for stage, grades in self.stage_mapping.items():
            add_row_ui(stage, ", ".join(grades))

        tk.Button(main_frame, text="+ إضافة مرحلة جديدة", command=lambda: add_row_ui("مرحلة جديدة", ""), bg="lightblue").pack(pady=10)

        def save():
            # Save mapping
            new_mapping = {}
            for name_ent, grades_ent in rows:
                name = name_ent.get().strip()
                if not name: continue
                grades_list = [g.strip() for g in grades_ent.get().split(',') if g.strip()]
                new_mapping[name] = grades_list
            self.stage_mapping = new_mapping
            
            # Save default days
            self.default_actual_days = days_ent.get().strip()
            
            # Sync with main UI
            self.school_entries['actual_days'].delete(0, tk.END)
            self.school_entries['actual_days'].insert(0, self.default_actual_days)
            self.save_school_info(quiet=True)

            self.save_config_to_file()
            settings_win.destroy()
            messagebox.showinfo("نجاح", "تم حفظ الإعدادات بنجاح")

        tk.Button(main_frame, text="حفظ الإعدادات", command=save, bg="lightgreen", font=self.label_font, width=20).pack(pady=10)

    def get_stage_for_grade(self, grade_name):
        if not grade_name: return ""
        
        matches = []
        for stage, grades in self.stage_mapping.items():
            for g in grades:
                if g in grade_name:
                    # Store (length of match, stage)
                    matches.append((len(g), stage))
        
        if matches:
            # Sort by length descending and take the longest match
            # This ensures "الثاني عشر" (length 11) is picked over "الثاني" (length 6)
            matches.sort(key=lambda x: x[0], reverse=True)
            return matches[0][1]
            
        return "أساسية" # Fallback

    def setup_ui(self):
        # 1. School Info Frame (Now at the very top)
        self.info_frame = tk.LabelFrame(self.root, text="بيانات المدرسة العامة", labelanchor="ne", font=self.label_font)
        self.info_frame.pack(pady=5, padx=10, fill=tk.X)
        
        self.school_entries = {}
        info_keys = [
            ('school_name', "اسم المدرسة"), 
            ('directorate', "المديرية"), 
            ('district', "اللواء"), 
            ('town', "البلدة"), 
            ('year', "العام الدراسي"),
            ('principal', "مدير المدرسة"),
            ('actual_days', "أيام الدوام")
        ]
        
        for i, (key, label) in enumerate(info_keys):
            col = 3 - (i % 4)
            row_base = (i // 4) * 2
            
            lbl = tk.Label(self.info_frame, text=" : " + label, font=self.label_font)
            lbl.grid(row=row_base, column=col, padx=10, pady=(5, 0), sticky="e")
            
            ent = tk.Entry(self.info_frame, justify=tk.RIGHT, fg="blue", font=self.entry_font, width=20)
            ent.grid(row=row_base+1, column=col, padx=10, pady=(0, 10), sticky="e")
            ent.bind("<KeyRelease>", lambda e: self.save_school_info(quiet=True))
            self.school_entries[key] = ent

        # 2. Dedicated Progress Frame (Below Info, Above Buttons)
        self.status_frame = tk.Frame(self.root, bg="#f8f9fa", bd=1, relief=tk.GROOVE)
        self.status_frame.pack(fill=tk.X, padx=10, pady=5)
        
        self.canvas_progress = tk.Canvas(self.status_frame, height=35, bg="#e6e6e6", highlightthickness=0)
        self.canvas_progress.pack(fill=tk.X, padx=20, pady=10)
        
        self.progress_rect = self.canvas_progress.create_rectangle(0, 0, 0, 35, fill="#4CAF50", outline="")
        self.progress_text = self.canvas_progress.create_text(0, 0, text="جاهز", font=("Arial", 12, "bold"), fill="#333333")
        
        self.canvas_progress.bind("<Configure>", lambda e: self.update_progress_ui())
        self.current_progress = 0
        self.current_msg = "جاهز"

        # 3. Buttons Frame (Below progress)
        btn_frame = tk.Frame(self.root)
        btn_frame.pack(pady=5, fill=tk.X)

        self.btn_load = tk.Button(btn_frame, text="فتح ملفات إكسيل", command=self.load_files, bg="lightblue", font=self.label_font)
        self.btn_load.pack(side=tk.LEFT, padx=10)

        self.btn_pdf = tk.Button(btn_frame, text="شهادات pdf", command=lambda: self.start_export('pdf'), bg="lightgreen", font=self.label_font)
        self.btn_pdf.pack(side=tk.LEFT, padx=10)

        self.btn_word = tk.Button(btn_frame, text="شهادات word", command=lambda: self.start_export('word'), bg="lightyellow", font=self.label_font)
        self.btn_word.pack(side=tk.LEFT, padx=10)

        self.lbl_total_students = tk.Label(btn_frame, text="إجمالي الطلاب: 0 | المحددة: 0", font=self.label_font, fg="darkblue")
        self.lbl_total_students.pack(side=tk.RIGHT, padx=20)

        # Main Container for Table
        self.main_container = tk.Frame(self.root)
        self.main_container.pack(pady=5, padx=10, fill=tk.BOTH, expand=True)

        # Main Table
        self.table_frame = tk.Frame(self.main_container)
        self.table_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)
        self.table_frame.grid_rowconfigure(0, weight=1)
        self.table_frame.grid_columnconfigure(0, weight=1)

        columns = ("select", "class", "section", "teacher", "principal", "students", "file", "open")
        self.tree = ttk.Treeview(self.table_frame, columns=columns, show="headings")
        self.tree.grid(row=0, column=0, sticky="nsew")

        # Scrollbars
        v_scroll = ttk.Scrollbar(self.table_frame, orient=tk.VERTICAL, command=self.tree.yview)
        v_scroll.grid(row=0, column=1, sticky="ns")
        
        h_scroll = ttk.Scrollbar(self.table_frame, orient=tk.HORIZONTAL, command=self.tree.xview)
        h_scroll.grid(row=1, column=0, sticky="ew")

        self.tree.configure(yscrollcommand=v_scroll.set, xscrollcommand=h_scroll.set)

        self.tree.heading("select", text="☐ الكل", command=self.toggle_select_all)
        self.tree.heading("class", text="الصف")
        self.tree.heading("section", text="شعبة")
        self.tree.heading("teacher", text="المربي")
        self.tree.heading("principal", text="المدير")
        self.tree.heading("students", text="طلاب")
        self.tree.heading("file", text="الملف")
        self.tree.heading("open", text="فتح")

        self.tree.column("select", width=50, anchor="center")
        self.tree.column("class", width=70, anchor="center")
        self.tree.column("section", width=50, anchor="center")
        self.tree.column("teacher", width=120, anchor="e")
        self.tree.column("principal", width=120, anchor="e")
        self.tree.column("students", width=50, anchor="center")
        self.tree.column("file", width=140, anchor="e")
        self.tree.column("open", width=80, anchor="center")

        self.tree.tag_configure('oddrow', background="white")
        self.tree.tag_configure('evenrow', background="#f0f0f0")
        self.tree.tag_configure('warning', background="#ffcccc")
        self.tree.tag_configure('section_missing', background="#fff0e0")
        self.tree.tag_configure('selected_row', foreground="#27ae60", font=("Arial", 13, "bold"))

        self.tree.bind("<Button-1>", self.on_click)
        self.tree.bind("<Double-1>", self.on_double_click)

    def update_progress_ui(self, val=None, msg=None):
        if val is not None: self.current_progress = val
        if msg is not None: self.current_msg = msg
        
        self.root.update_idletasks()
        w = self.canvas_progress.winfo_width()
        h = self.canvas_progress.winfo_height()
        
        # Update rectangle
        self.canvas_progress.coords(self.progress_rect, 0, 0, (self.current_progress / 100) * w, h)
        # Update text
        self.canvas_progress.itemconfig(self.progress_text, text=self.current_msg)
        self.canvas_progress.coords(self.progress_text, w/2, h/2)

    def toggle_select_all(self):
        self.select_all_state = not self.select_all_state
        mark = "✓ الكل" if self.select_all_state else "☐ الكل"
        self.tree.heading("select", text=mark)
        
        for class_data in self.all_classes:
            if self.select_all_state:
                if len(class_data['students']) > 0 and class_data['school_info']['section']:
                    class_data['selected'] = True
                else:
                    class_data['selected'] = False
            else:
                class_data['selected'] = False
        self.refresh_ui()

    def on_click(self, event):
        item_id = self.tree.identify_row(event.y)
        column = self.tree.identify_column(event.x)
        if not item_id: return

        idx = int(item_id)
        class_data = self.all_classes[idx]

        if column == "#1": # Selection Column
            if len(class_data['students']) == 0:
                 return
            class_data['selected'] = not class_data.get('selected', False)
            self.refresh_ui()
        
        elif column == "#8": # Open Column
            # Try to open the export directory or the file itself
            current_info = {key: env_val.get() if hasattr(env_val, 'get') else env_val for key, env_val in self.school_entries.items()}
            school_name = self.sanitize_filename(current_info.get('school_name', 'الشهادات').strip() or 'الشهادات')
            
            grade = self.sanitize_filename(class_data['school_info']['grade'])
            section = self.sanitize_filename(class_data['school_info']['section'])
            
            # Prefer PDF folder if it exists, else Word, else base export folder
            pdf_path = os.path.join(os.getcwd(), "شهادات pdf", school_name, f"شهادات_{grade}_{section}.pdf")
            word_path = os.path.join(os.getcwd(), "شهادات word", school_name, f"شهادات_{grade}_{section}.docx")
            
            target = None
            if os.path.exists(pdf_path): target = pdf_path
            elif os.path.exists(word_path): target = word_path
            else:
                # Try opening the school folder at least
                base_pdf = os.path.join(os.getcwd(), "شهادات pdf", school_name)
                base_word = os.path.join(os.getcwd(), "شهادات word", school_name)
                if os.path.exists(base_pdf): target = base_pdf
                elif os.path.exists(base_word): target = base_word

            if target:
                os.startfile(target)
            else:
                messagebox.showinfo("تنبيه", "لم يتم تصدير هذا الصف بعد أو الملف غير موجود")

    def load_files(self):
        files = filedialog.askopenfilenames(filetypes=[("Excel files", "*.xlsx")])
        if not files:
            return
        
        self.btn_load.config(state=tk.DISABLED)
        self.all_classes = []
        self.processor = ExcelProcessor()
        
        def run_load():
            total_files = len(files)
            for i, f in enumerate(files):
                filename = os.path.basename(f)
                count_info = f"({i+1}/{total_files})"
                msg = f"جاري تحميل {filename} {count_info}..."
                self.root.after(0, lambda m=msg: self.update_progress_ui(msg=m))
                try:
                    results = self.processor.process_file(f)
                    for r in results:
                        r['selected'] = False
                    self.all_classes.extend(results)
                except Exception as e:
                    self.root.after(0, lambda name=filename, ex=e: messagebox.showerror("خطأ", f"فشل قراءة الملف {name}: {str(ex)}"))
                
                prog = ((i + 1) / total_files) * 100
                self.root.after(0, lambda v=prog: self.update_progress_ui(val=v))
            
            self.root.after(0, self.finish_load)

        threading.Thread(target=run_load, daemon=True).start()

    def finish_load(self):
        self.btn_load.config(state=tk.NORMAL)
        self.update_progress_ui(val=100, msg="تم تحميل الملفات بنجاح")
        self.refresh_ui()
        # Reset progress after a delay
        self.root.after(3000, lambda: self.update_progress_ui(val=0, msg="جاهز"))

    def refresh_ui(self):
        y_pos = self.tree.yview()[0]
        agg_info = self.processor.get_aggregated_school_info()
        
        # Avoid updating entries if they are currently being edited by the user
        focused_widget = self.root.focus_get()
        
        for key, ent in self.school_entries.items():
            if ent == focused_widget:
                continue # Don't interrupt user typing
                
            ent.delete(0, tk.END)
            if key == 'principal':
                ent.insert(0, agg_info.get('principal_name', ""))
            else:
                ent.insert(0, agg_info.get(key, ""))

        for item in self.tree.get_children():
            self.tree.delete(item)

        total_students = 0
        selected_students = 0

        # Current school name for path checking
        current_info = {key: ent.get() for key, ent in self.school_entries.items()}
        school_name = self.sanitize_filename(current_info.get('school_name', 'الشهادات').strip() or 'الشهادات')

        for i, class_data in enumerate(self.all_classes):
            num_students = len(class_data['students'])
            total_students += num_students
            
            is_selected = class_data.get('selected', False)
            if is_selected:
                selected_students += num_students

            is_empty = (num_students == 0)
            is_missing_section = not class_data['school_info']['section']
            
            # Base tags
            tags = []
            if is_selected: tags.append('selected_row')
            
            if is_empty: tags.append('warning')
            elif is_missing_section: tags.append('section_missing')
            else: tags.append('evenrow' if i % 2 == 0 else 'oddrow')
            
            # Selection mark logic
            select_mark = "✅" if is_selected else "☐"
            if is_empty: select_mark = "✕"
            elif is_missing_section: select_mark = "✅" if is_selected else "⚠"

            # Check if file exists to show 'Open'
            grade = self.sanitize_filename(class_data['school_info']['grade'])
            section = self.sanitize_filename(class_data['school_info']['section'])
            pdf_path = os.path.join(os.getcwd(), "شهادات pdf", school_name, f"شهادات_{grade}_{section}.pdf")
            word_path = os.path.join(os.getcwd(), "شهادات word", school_name, f"شهادات_{grade}_{section}.docx")
            
            open_text = ""
            if os.path.exists(pdf_path) or os.path.exists(word_path):
                open_text = "📂 فتح"

            self.tree.insert("", tk.END, iid=i, values=(
                select_mark,
                class_data['school_info']['grade'],
                class_data['school_info']['section'],
                class_data['teacher_name'],
                class_data['principal_name'],
                num_students,
                os.path.basename(class_data['file_path']),
                open_text
            ), tags=tuple(tags))
        
        self.lbl_total_students.config(text=f"إجمالي الطلاب: {total_students} | المحددة: {selected_students}")
        self.tree.yview_moveto(y_pos)

    def save_school_info(self, quiet=False):
        if not self.all_classes: return
        new_info = {key: ent.get() for key, ent in self.school_entries.items()}
        for class_data in self.all_classes:
            # Sync all keys including actual_days
            for key in ['school_name', 'directorate', 'district', 'town', 'year', 'actual_days']:
                class_data['school_info'][key] = new_info.get(key, "")
            class_data['principal_name'] = new_info['principal']
            class_data['school_info']['principal_name'] = new_info['principal']
        self.refresh_ui()
        if not quiet:
            messagebox.showinfo("نجاح", "تم تحديث البيانات العامة لجميع الصفوف")

    def on_double_click(self, event):
        item_id = self.tree.identify_row(event.y)
        column = self.tree.identify_column(event.x)
        if not item_id or not column: return
        class_idx = int(item_id)
        class_data = self.all_classes[class_idx]
        if len(class_data['students']) > 0:
            class_data['selected'] = not class_data.get('selected', False)
            self.refresh_ui()
        col_idx = int(column.replace("#", ""))
        if col_idx in [4, 5]: 
            x, y, w, h = self.tree.bbox(item_id, column)
            old_val = self.tree.item(item_id, 'values')[col_idx-1]
            entry = tk.Entry(self.tree, font=self.table_font)
            entry.insert(0, old_val)
            entry.select_range(0, tk.END)
            entry.focus_set()
            def save_edit(e=None):
                new_val = entry.get()
                class_idx = int(item_id)
                if col_idx == 4: self.all_classes[class_idx]['teacher_name'] = new_val
                elif col_idx == 5:
                    self.all_classes[class_idx]['principal_name'] = new_val
                    if class_idx == 0:
                        self.school_entries['principal'].delete(0, tk.END)
                        self.school_entries['principal'].insert(0, new_val)
                entry.destroy()
                self.refresh_ui()
            entry.bind("<Return>", save_edit)
            entry.bind("<FocusOut>", lambda e: entry.destroy())
            entry.place(x=x, y=y, width=w, height=h)

    def sanitize_filename(self, filename):
        # Remove characters that are illegal in Windows filenames
        return re.sub(r'[<>:"/\\|?*]', '_', str(filename))

    def start_export(self, format_type='both'):
        selected_classes = [c for c in self.all_classes if c.get('selected')]
        if not selected_classes:
            messagebox.showwarning("تنبيه", "يرجى تحديد صفوف صالحة للتصدير أولاً")
            return
        if not os.path.exists("m.docx"):
            messagebox.showerror("خطأ", "ملف النموذج m.docx غير موجود")
            return

        # Trial Mode Restrictions
        export_limit = None
        if not self.is_activated:
            if format_type == 'word' or format_type == 'both':
                messagebox.showwarning("الوضع التجريبي", "تصدير ملفات Word متاح فقط في النسخة المفعلة")
                return
            if self.total_trial_usage >= self.trial_max_total:
                messagebox.showerror("الوضع التجريبي", "لقد استنفدت الحد المسموح به في الوضع التجريبي (30 شهادة). يرجى تفعيل البرنامج.")
                return
            export_limit = self.trial_per_file_limit
            messagebox.showinfo("الوضع التجريبي", f"سيتم تصدير أول {export_limit} طلاب فقط من كل صف في الوضع التجريبي.")

        current_info = {key: env_val.get() if hasattr(env_val, 'get') else env_val for key, env_val in self.school_entries.items()}
        school_name = self.sanitize_filename(current_info.get('school_name', 'الشهادات').strip() or 'الشهادات')
        
        # Paths for both formats
        pdf_base_dir = os.path.join(os.getcwd(), "شهادات pdf", school_name)
        word_base_dir = os.path.join(os.getcwd(), "شهادات word", school_name)
        
        if format_type in ['pdf', 'both']:
            if not os.path.exists(pdf_base_dir):
                os.makedirs(pdf_base_dir)
        if format_type in ['word', 'both']:
            if not os.path.exists(word_base_dir):
                os.makedirs(word_base_dir)

        school_info = {
            'school_name': current_info['school_name'], 'directorate': current_info['directorate'],
            'district': current_info['district'], 'town': current_info['town'],
            'year': current_info['year'], 'principal_name': current_info['principal'],
            'actual_days': current_info['actual_days'],
            'national_id': self.processor.get_aggregated_school_info().get('national_id', '')
        }

        self.btn_pdf.config(state=tk.DISABLED)
        self.btn_word.config(state=tk.DISABLED)

        def run_export():
            word_gen = WordGenerator("m.docx")
            total_classes = len(selected_classes)
            exported_count = 0
            
            try:
                for i, class_data in enumerate(selected_classes):
                    if not self.is_activated and (self.total_trial_usage + exported_count) >= self.trial_max_total:
                        break # Stop if total limit reached

                    grade = self.sanitize_filename(class_data['school_info']['grade'])
                    section = self.sanitize_filename(class_data['school_info']['section'])
                    
                    pdf_filename = f"شهادات_{grade}_{section}.pdf"
                    word_filename = f"شهادات_{grade}_{section}.docx"
                    
                    pdf_path = os.path.join(pdf_base_dir, pdf_filename) if format_type in ['pdf', 'both'] else None
                    word_path = os.path.join(word_base_dir, word_filename) if format_type in ['word', 'both'] else None
                    
                    msg = f"جاري تصدير {grade} {section}..."
                    self.root.after(0, lambda m=msg: self.update_progress_ui(msg=m))
                    
                    students_to_process = []
                    stage = self.get_stage_for_grade(grade)
                    for student in class_data['students']:
                        s_copy = student.copy()
                        s_copy['الصف'] = grade
                        s_copy['الشعبة'] = section
                        s_copy['مربي_الصف'] = class_data['teacher_name']
                        s_copy['مدير_المدرسة'] = school_info['principal_name']
                        s_copy['mar'] = stage
                        students_to_process.append(s_copy)
                    
                    current_school_info = school_info.copy()
                    current_school_info['grade'] = grade

                    def update_prog(curr, tot):
                        class_base = (i / total_classes) * 100
                        class_weight = (1 / total_classes) * 100
                        overall = class_base + (curr / tot) * class_weight
                        m = f"جاري تصدير {grade} {section} | طالب {curr} من {tot}"
                        self.root.after(0, lambda v=overall, m_text=m: self.update_progress_ui(val=v, msg=m_text))

                    # Apply limit and track usage
                    actual_limit = export_limit
                    if actual_limit:
                        # Ensure we don't exceed the total 30 certificates limit
                        remaining_trial = self.trial_max_total - (self.total_trial_usage + exported_count)
                        actual_limit = min(actual_limit, remaining_trial)
                    
                    num_exported_in_this_file = min(len(students_to_process), actual_limit) if actual_limit else len(students_to_process)
                    
                    word_gen.generate_section_file(students_to_process, current_school_info, 
                                                   pdf_path=pdf_path, word_path=word_path, 
                                                   progress_callback=update_prog, limit=actual_limit)
                    
                    exported_count += num_exported_in_this_file
                
                if not self.is_activated:
                    self.total_trial_usage += exported_count
                    self.save_config_to_file()
                    # Sync with server
                    try:
                        requests.post(f"{self.render_base_url}/sync_trial", 
                                      json={'hwid': self.hwid, 'count': exported_count}, 
                                      timeout=45)
                    except: pass
                    self.root.after(0, self.update_trial_status_ui)
                else:
                    # Sync real activation usage with server
                    try:
                        requests.post(f"{self.render_base_url}/sync_usage", 
                                      json={'hwid': self.hwid, 'count': exported_count}, 
                                      timeout=45)
                        # Re-fetch license info to update UI counter
                        self.check_activation_on_startup()
                    except: pass

                success_msg = "تم تصدير الشهادات بنجاح"
                if format_type == 'pdf': success_msg = f"تم تصدير ملفات PDF بنجاح في:\n{pdf_base_dir}"
                elif format_type == 'word': success_msg = f"تم تصدير ملفات Word بنجاح في:\n{word_base_dir}"
                else: success_msg = f"تم تصدير الملفات بنجاح في:\n{school_name}"
                
                self.root.after(0, lambda: messagebox.showinfo("نجاح", success_msg))
            except Exception as e:
                self.root.after(0, lambda ex=e: messagebox.showerror("خطأ", f"فشل التصدير: {str(ex)}"))
            finally:
                self.root.after(0, self.cleanup_export)

        threading.Thread(target=run_export, daemon=True).start()

    def cleanup_export(self):
        self.update_progress_ui(val=0, msg="جاهز")
        self.btn_pdf.config(state=tk.NORMAL)
        self.btn_word.config(state=tk.NORMAL)

if __name__ == "__main__":
    root = tk.Tk()
    app = StudentApp(root)
    root.mainloop()
