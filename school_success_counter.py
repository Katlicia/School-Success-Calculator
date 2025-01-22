import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import sqlite3
import pandas as pd
import json

class SuccessCalculationApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Başarı Hesaplama Arayüzü")
        self.root.geometry("800x600")
        self.criteria= []
        self.weights = {}

        # Veritabanı bağlantısı oluştur
        self.conn = sqlite3.connect("courses.db")
        self.cursor = self.conn.cursor()
        self.create_database()

        self.create_main_menu()

    def create_database(self):
        # Veritabanı tablolarını oluştur
        self.cursor.execute("""
            CREATE TABLE IF NOT EXISTS courses (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                course_code TEXT NOT NULL UNIQUE,
                course_name TEXT NOT NULL
            )
        """)
        self.cursor.execute("""
            CREATE TABLE IF NOT EXISTS program_outcomes (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                outcome TEXT NOT NULL
            )
        """)
        self.cursor.execute("""
            CREATE TABLE IF NOT EXISTS course_outcomes (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                course_id INTEGER NOT NULL,
                outcome TEXT NOT NULL,
                FOREIGN KEY(course_id) REFERENCES courses(id) ON DELETE CASCADE
            )
        """)
        self.cursor.execute("""
            CREATE TABLE IF NOT EXISTS criteria (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                course_id INTEGER NOT NULL,
                criterion TEXT NOT NULL,
                weight REAL NOT NULL,
                FOREIGN KEY(course_id) REFERENCES courses(id) ON DELETE CASCADE
            )
        """)
        self.cursor.execute("""
            CREATE TABLE IF NOT EXISTS relationship_values (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                course_id INTEGER NOT NULL,
                program_outcome_id INTEGER NOT NULL,
                course_outcome_id INTEGER NOT NULL,
                value REAL NOT NULL CHECK (value >= 0 AND value <= 1),
                FOREIGN KEY(course_id) REFERENCES courses(id) ON DELETE CASCADE,
                FOREIGN KEY(program_outcome_id) REFERENCES program_outcomes(id) ON DELETE CASCADE,
                FOREIGN KEY(course_outcome_id) REFERENCES course_outcomes(id) ON DELETE CASCADE
            )
        """)
        self.cursor.execute("""
            CREATE TABLE IF NOT EXISTS course_outcome_evaluations (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                course_id INTEGER NOT NULL,
                course_outcome_id INTEGER NOT NULL,
                criterion_id INTEGER NOT NULL,
                value REAL NOT NULL CHECK (value >= 0 AND value <= 1),
                FOREIGN KEY(course_id) REFERENCES courses(id) ON DELETE CASCADE,
                FOREIGN KEY(course_outcome_id) REFERENCES course_outcomes(id) ON DELETE CASCADE,
                FOREIGN KEY(criterion_id) REFERENCES criteria(id) ON DELETE CASCADE
            )
        """)
        self.cursor.execute("""
            CREATE TABLE IF NOT EXISTS student_grades (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                course_id INTEGER NOT NULL,
                student_name TEXT NOT NULL,
                grades TEXT NOT NULL,
                FOREIGN KEY(course_id) REFERENCES courses(id) ON DELETE CASCADE,
                UNIQUE(course_id, student_name)
            )
        """)
        self.cursor.execute("""
            CREATE TABLE IF NOT EXISTS outcome_relations (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                course_id INTEGER,
                program_outcome_id INTEGER,
                course_outcome_id INTEGER,
                contribution_level REAL,
                FOREIGN KEY (course_id) REFERENCES courses (id),
                FOREIGN KEY (program_outcome_id) REFERENCES program_outcomes (id),
                FOREIGN KEY (course_outcome_id) REFERENCES course_outcomes (id)
            )
        """)
        # Tablo 4 için başarı oranları tablosu
        self.cursor.execute("""
            CREATE TABLE IF NOT EXISTS success_rates (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                course_id INTEGER,
                student_name TEXT,
                course_outcome_id INTEGER,
                success_rate REAL,
                FOREIGN KEY (course_id) REFERENCES courses (id),
                FOREIGN KEY (course_outcome_id) REFERENCES course_outcomes (id)
            )
        """)
        self.conn.commit()

    def create_main_menu(self):
        self.clear_frame()
        self.main_frame = ttk.Frame(self.root)
        self.main_frame.pack(fill="both", expand=True, padx=20, pady=20)

        ttk.Label(self.main_frame, text="Başarı Hesaplama Arayüzüne Hoşgeldiniz", 
                 font=("Arial", 16)).pack(pady=20)

        # Standart buton genişliği
        button_width = 25

        ttk.Button(self.main_frame, text="Ders Ekle", 
                   command=self.add_course, width=button_width).pack(pady=5)
        ttk.Button(self.main_frame, text="Program Çıktısı Ekle", 
                   command=self.add_program_outcomes, width=button_width).pack(pady=5)
        ttk.Button(self.main_frame, text="Ders Çıktısı Ekle", 
                   command=self.add_course_outcomes, width=button_width).pack(pady=5)
        ttk.Button(self.main_frame, text="Değerlendirme Kriterleri", 
                   command=self.criteria_screen, width=button_width).pack(pady=5)
        ttk.Button(self.main_frame, text="Tablo 1 Giriş", 
                   command=self.table1_screen, width=button_width).pack(pady=5)
        ttk.Button(self.main_frame, text="Tablo 2 Giriş", 
                   command=self.table2_screen, width=button_width).pack(pady=5)
        ttk.Button(self.main_frame, text="Öğrenci Notları", 
                   command=self.student_grades_screen, width=button_width).pack(pady=5)
        ttk.Button(self.main_frame, text="Tablo 4", 
                   command=self.table4_screen, width=button_width).pack(pady=5)
        ttk.Button(self.main_frame, text="Tablo 5", 
                   command=self.table5_screen, width=button_width).pack(pady=5)
        ttk.Button(self.main_frame, text="Çıkış", 
                   command=self.root.quit, width=button_width).pack(pady=5)

    def add_course(self):
        self.clear_frame()
        self.course_frame = ttk.Frame(self.root)
        self.course_frame.pack(fill="both", expand=True, padx=20, pady=20)

        # Başlık
        ttk.Label(self.course_frame, text="Ders Ekle", 
                 font=("Arial", 16)).pack(pady=20)

        # İçerik frame
        content_frame = ttk.Frame(self.course_frame)
        content_frame.pack(fill="both", expand=True, padx=10)

        ttk.Label(content_frame, text="Ders Kodu:").pack(pady=5)
        self.course_code_entry = ttk.Entry(content_frame, width=30)
        self.course_code_entry.pack(pady=5)

        ttk.Label(content_frame, text="Ders Adı:").pack(pady=5)
        self.course_name_entry = ttk.Entry(content_frame, width=30)
        self.course_name_entry.pack(pady=5)

        # Buton frame
        button_frame = ttk.Frame(self.course_frame)
        button_frame.pack(fill="x", pady=10, side="bottom")
        
        ttk.Button(button_frame, text="Kaydet", 
                   command=self.save_course).pack(side="left", padx=5)
        ttk.Button(button_frame, text="Ana Menüye Dön", 
                   command=self.create_main_menu).pack(side="left", padx=5)

    def save_course(self):
        course_code = self.course_code_entry.get()
        course_name = self.course_name_entry.get()

        if course_code and course_name:
            self.cursor.execute("SELECT * FROM courses WHERE course_code = ?", (course_code,))
            existing_course = self.cursor.fetchone()

            if existing_course:
                messagebox.showwarning("Uyarı", f"{course_code} kodlu ders zaten kayıtlı.")
            else:
                self.cursor.execute("INSERT INTO courses (course_code, course_name) VALUES (?, ?)", (course_code, course_name))
                self.conn.commit()
                messagebox.showinfo("Başarılı", f"{course_code} - {course_name} dersi eklendi.")
                self.create_main_menu()
        else:
            messagebox.showerror("Hata", "Lütfen tüm alanları doldurun.")

    def add_program_outcomes(self):
        self.clear_frame()
        self.program_outcomes_frame = ttk.Frame(self.root)
        self.program_outcomes_frame.pack(fill="both", expand=True, padx=20, pady=20)

        # Başlık
        ttk.Label(self.program_outcomes_frame, text="Program Çıktısı Ekle", 
                 font=("Arial", 16)).pack(pady=20)

        # İçerik frame
        content_frame = ttk.Frame(self.program_outcomes_frame)
        content_frame.pack(fill="both", expand=True, padx=10)

        # Sol frame - Giriş alanı
        left_frame = ttk.Frame(content_frame)
        left_frame.pack(side="left", fill="both", expand=True, padx=5)

        ttk.Label(left_frame, text="Kaç tane program çıktısı ekleyeceksiniz?").pack(pady=5)
        self.num_outcomes_entry = ttk.Entry(left_frame, width=30)
        self.num_outcomes_entry.pack(pady=5)

        # Sağ frame - Mevcut çıktılar listesi
        right_frame = ttk.Frame(content_frame)
        right_frame.pack(side="right", fill="both", expand=True, padx=5)

        ttk.Label(right_frame, text="Mevcut Program Çıktıları:").pack(pady=5)
        self.outcomes_listbox = tk.Listbox(right_frame, width=50, height=15)
        self.outcomes_listbox.pack(pady=5)

        # Mevcut program çıktılarını yükle
        self.cursor.execute("SELECT outcome FROM program_outcomes ORDER BY id")
        outcomes = self.cursor.fetchall()
        for i, outcome in enumerate(outcomes, 1):
            self.outcomes_listbox.insert(tk.END, f"PÇ{i}: {outcome[0]}")

        # Buton frame
        button_frame = ttk.Frame(self.program_outcomes_frame)
        button_frame.pack(fill="x", pady=10, side="bottom")

        ttk.Button(button_frame, text="Devam", 
                   command=self.get_outcomes).pack(side="left", padx=5)
        ttk.Button(button_frame, text="Excel'den Yükle", 
                   command=self.load_program_outcomes).pack(side="left", padx=5)
        ttk.Button(button_frame, text="Ana Menüye Dön", 
                   command=self.create_main_menu).pack(side="left", padx=5)

    def get_outcomes(self):
        try:
            num_outcomes = int(self.num_outcomes_entry.get())
            if num_outcomes <= 0:
                raise ValueError
        except ValueError:
            messagebox.showerror("Hata", "Lütfen geçerli bir sayı girin.")
            return

        self.clear_frame()
        self.manual_outcomes_frame = ttk.Frame(self.root)
        self.manual_outcomes_frame.pack(fill="both", expand=True, padx=20, pady=20)

        # Başlık
        ttk.Label(self.manual_outcomes_frame, text="Program Çıktıları", 
                 font=("Arial", 16)).pack(pady=20)

        # İçerik frame
        content_frame = ttk.Frame(self.manual_outcomes_frame)
        content_frame.pack(fill="both", expand=True, padx=10)

        self.outcomes_entries = []
        for i in range(num_outcomes):
            ttk.Label(content_frame, text=f"Program Çıktısı {i+1}:").pack(pady=5)
            entry = ttk.Entry(content_frame, width=30)
            entry.pack(pady=5)
            self.outcomes_entries.append(entry)

        # Buton frame
        button_frame = ttk.Frame(self.manual_outcomes_frame)
        button_frame.pack(fill="x", pady=10, side="bottom")

        ttk.Button(button_frame, text="Kaydet", 
                   command=self.save_outcomes_to_db).pack(side="left", padx=5)
        ttk.Button(button_frame, text="Ana Menüye Dön", 
                   command=self.create_main_menu).pack(side="left", padx=5)

    def save_outcomes_to_db(self):
        outcomes = [entry.get() for entry in self.outcomes_entries if entry.get().strip()]
        if not outcomes:
            messagebox.showerror("Hata", "Lütfen tüm alanları doldurun.")
            return

        try:
            for outcome in outcomes:
                self.cursor.execute("INSERT INTO program_outcomes (outcome) VALUES (?)", (outcome,))
            self.conn.commit()
            messagebox.showinfo("Başarılı", "Program çıktıları veritabanına kaydedildi.")
            self.create_main_menu()
        except Exception as e:
            messagebox.showerror("Hata", f"Veritabanına kaydedilemedi: {e}")

    def load_program_outcomes(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx"), ("All Files", "*.*")])
        if file_path:
            try:
                df = pd.read_excel(file_path, header=None)  # Ensure no header row is assumed
                outcomes = df.iloc[:, 0].tolist()

                for outcome in outcomes:
                    self.cursor.execute("INSERT INTO program_outcomes (outcome) VALUES (?)", (outcome,))
                self.conn.commit()

                messagebox.showinfo("Başarılı", "Program çıktıları Excel'den yüklenip veritabanına kaydedildi.")
            except Exception as e:
                messagebox.showerror("Hata", f"Dosya yüklenemedi veya kaydedilemedi: {e}")

    def add_course_outcomes(self):
        self.clear_frame()
        self.course_outcomes_frame = ttk.Frame(self.root)
        self.course_outcomes_frame.pack(fill="both", expand=True, padx=20, pady=20)

        # Başlık
        ttk.Label(self.course_outcomes_frame, text="Ders Çıktısı Ekle", 
                 font=("Arial", 16)).pack(pady=20)

        # İçerik frame
        content_frame = ttk.Frame(self.course_outcomes_frame)
        content_frame.pack(fill="both", expand=True, padx=10)

        # Sol frame - Giriş alanı
        left_frame = ttk.Frame(content_frame)
        left_frame.pack(side="left", fill="both", expand=True, padx=5)

        # Ders seçimi
        self.cursor.execute("SELECT id, course_name FROM courses")
        courses = self.cursor.fetchall()
        if not courses:
            messagebox.showerror("Hata", "Kayıtlı bir ders bulunamadı.")
            self.create_main_menu()
            return

        ttk.Label(left_frame, text="Ders Seçiniz:").pack(pady=5)
        self.selected_course = tk.StringVar()
        course_names = [f"{course[1]}" for course in courses]
        self.course_combobox = ttk.Combobox(left_frame, textvariable=self.selected_course, 
                                           values=course_names, width=30)
        self.course_combobox.pack(pady=5)
        self.course_combobox.bind('<<ComboboxSelected>>', self.on_course_select)

        ttk.Label(left_frame, text="Kaç tane ders çıktısı ekleyeceksiniz?").pack(pady=5)
        self.num_outcomes_entry = ttk.Entry(left_frame, width=30)
        self.num_outcomes_entry.pack(pady=5)

        # Sağ frame - Mevcut çıktılar listesi
        right_frame = ttk.Frame(content_frame)
        right_frame.pack(side="right", fill="both", expand=True, padx=5)

        ttk.Label(right_frame, text="Mevcut Ders Çıktıları:").pack(pady=5)
        self.course_outcomes_listbox = tk.Listbox(right_frame, width=50, height=15)
        self.course_outcomes_listbox.pack(pady=5)

        # Buton frame
        button_frame = ttk.Frame(self.course_outcomes_frame)
        button_frame.pack(fill="x", pady=10, side="bottom")

        ttk.Button(button_frame, text="Devam", 
                   command=self.get_course_outcomes).pack(side="left", padx=5)
        ttk.Button(button_frame, text="Excel'den Yükle", 
                   command=self.load_course_outcomes).pack(side="left", padx=5)
        ttk.Button(button_frame, text="Ana Menüye Dön", 
                   command=self.create_main_menu).pack(side="left", padx=5)

    def get_course_outcomes(self):
        try:
            num_outcomes = int(self.num_outcomes_entry.get())
            if num_outcomes <= 0:
                raise ValueError
        except ValueError:
            messagebox.showerror("Hata", "Lütfen geçerli bir sayı girin.")
            return

        self.clear_frame()
        self.manual_course_outcomes_frame = ttk.Frame(self.root)
        self.manual_course_outcomes_frame.pack(fill="both", expand=True)

        ttk.Label(self.manual_course_outcomes_frame, text="Ders Çıktıları", font=("Arial", 16)).pack(pady=20)

        self.outcomes_entries = []
        for i in range(num_outcomes):
            ttk.Label(self.manual_course_outcomes_frame, text=f"Ders Çıktısı {i+1}:").pack(pady=5)
            entry = ttk.Entry(self.manual_course_outcomes_frame)
            entry.pack(pady=5)
            self.outcomes_entries.append(entry)

        ttk.Button(self.manual_course_outcomes_frame, text="Kaydet", command=self.save_course_outcomes).pack(pady=10)
        ttk.Button(self.manual_course_outcomes_frame, text="Ana Menüye Dön", command=self.create_main_menu).pack(pady=10)

    def save_course_outcomes(self):
        selected_course_name = self.selected_course.get()
        if not selected_course_name:
            messagebox.showerror("Hata", "Lütfen bir ders seçin.")
            return

        self.cursor.execute("SELECT id FROM courses WHERE course_name = ?", (selected_course_name,))
        course_id = self.cursor.fetchone()
        if not course_id:
            messagebox.showerror("Hata", "Ders bulunamadı.")
            return

        course_id = course_id[0]
        outcomes = [entry.get() for entry in self.outcomes_entries if entry.get().strip()]
        if not outcomes:
            messagebox.showerror("Hata", "Lütfen tüm alanları doldurun.")
            return

        try:
            # Yeni çıktıları ekle
            for outcome in outcomes:
                self.cursor.execute("INSERT INTO course_outcomes (course_id, outcome) VALUES (?, ?)", (course_id, outcome))
            self.conn.commit()
            messagebox.showinfo("Başarılı", "Ders çıktıları kaydedildi.")
            self.create_main_menu()
        except Exception as e:
            messagebox.showerror("Hata", f"Veritabanına kaydedilemedi: {e}")

    def load_course_outcomes(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx"), ("All Files", "*.*")])
        if file_path:
            try:
                df = pd.read_excel(file_path, header=None)  # Ensure no header row is assumed
                outcomes = df.iloc[:, 0].tolist()
                selected_course_name = self.selected_course.get()

                if not selected_course_name:
                    messagebox.showerror("Hata", "Lütfen bir ders seçin.")
                    return

                self.cursor.execute("SELECT id FROM courses WHERE course_name = ?", (selected_course_name,))
                course_id = self.cursor.fetchone()
                if not course_id:
                    messagebox.showerror("Hata", "Ders bulunamadı.")
                    return

                course_id = course_id[0]
                
                # Yeni çıktıları ekle
                for outcome in outcomes:
                    self.cursor.execute("INSERT INTO course_outcomes (course_id, outcome) VALUES (?, ?)", (course_id, outcome))
                self.conn.commit()

                messagebox.showinfo("Başarılı", "Ders çıktıları Excel'den yüklenip veritabanına kaydedildi.")
                self.create_main_menu()
            except Exception as e:
                messagebox.showerror("Hata", f"Dosya yüklenemedi veya kaydedilemedi: {e}")

    def on_course_select(self, event=None):
        """Seçilen derse ait çıktıları listbox'ta göster"""
        self.course_outcomes_listbox.delete(0, tk.END)
        if not self.selected_course.get():
            return

        self.cursor.execute("""
            SELECT co.outcome 
            FROM course_outcomes co 
            JOIN courses c ON co.course_id = c.id 
            WHERE c.course_name = ?
            ORDER BY co.id
        """, (self.selected_course.get(),))
        
        outcomes = self.cursor.fetchall()
        for i, outcome in enumerate(outcomes, 1):
            self.course_outcomes_listbox.insert(tk.END, f"ÇK{i}: {outcome[0]}")

    def criteria_screen(self):
        self.clear_frame()
        self.criteria_frame = ttk.Frame(self.root)
        self.criteria_frame.pack(fill="both", expand=True, padx=20, pady=20)

        # Başlık
        ttk.Label(self.criteria_frame, text="Değerlendirme Kriterleri ve Ağırlıkları", 
                 font=("Arial", 16)).pack(pady=20)

        # İçerik frame
        content_frame = ttk.Frame(self.criteria_frame)
        content_frame.pack(fill="both", expand=True, padx=10)

        # Ders seçimi ekle
        course_frame = ttk.Frame(content_frame)
        course_frame.pack(pady=5)
        
        ttk.Label(course_frame, text="Ders Seçin:").pack(side="left", padx=5)
        self.selected_course = tk.StringVar()
        self.cursor.execute("SELECT course_name FROM courses")
        courses = [course[0] for course in self.cursor.fetchall()]
        self.course_combobox = ttk.Combobox(course_frame, textvariable=self.selected_course, 
                                           values=courses, width=30)
        self.course_combobox.pack(side="left", padx=5)
        self.course_combobox.bind('<<ComboboxSelected>>', self.load_criteria_from_db)

        # Kriter Seçimi
        criteria_container = ttk.Frame(content_frame)
        criteria_container.pack(pady=5)

        ttk.Label(criteria_container, text="Kriter Seçin:").pack(side="left", padx=5)
        self.criteria_var = tk.StringVar()
        self.criteria_combobox = ttk.Combobox(criteria_container, textvariable=self.criteria_var, width=25)
        self.criteria_combobox['values'] = ["Ödev", "Proje", "Sunum", "Rapor", "KPL", 
                                          "Quiz", "Vize", "Final", "Diğer"]
        self.criteria_combobox.pack(side="left", padx=5)
        self.criteria_combobox.bind("<<ComboboxSelected>>", self.check_other_selection)

        self.other_entry_label = ttk.Label(criteria_container, text="Diğer Kriter Adı:")
        self.other_entry = ttk.Entry(criteria_container)

        weight_container = ttk.Frame(content_frame)
        weight_container.pack(pady=5)
        ttk.Label(weight_container, text="Ağırlık (%):").pack(side="left", padx=5)
        self.weight_entry = ttk.Entry(weight_container, width=10)
        self.weight_entry.pack(side="left", padx=5)

        # Kriter Listesi
        self.criteria_listbox = tk.Listbox(content_frame, width=50, height=10)
        self.criteria_listbox.pack(pady=10)

        # Buton frame
        button_frame = ttk.Frame(self.criteria_frame)
        button_frame.pack(fill="x", pady=10, side="bottom")

        ttk.Button(button_frame, text="Ekle", 
                   command=self.add_criteria).pack(side="left", padx=5)
        ttk.Button(button_frame, text="Seçili Kriteri Düzenle", 
                   command=self.edit_selected_criteria).pack(side="left", padx=5)
        ttk.Button(button_frame, text="Kaydet", 
                   command=self.save_criteria_to_db).pack(side="left", padx=5)
        ttk.Button(button_frame, text="Ana Menüye Dön", 
                   command=self.create_main_menu).pack(side="left", padx=5)

        # Kayıtlı kriterleri yükle
        self.load_criteria_from_db()

    def check_other_selection(self, event):
        if self.criteria_var.get() == "Diğer":
            self.other_entry_label.pack(side="left", padx=5)
            self.other_entry.pack(side="left", padx=5)
        else:
            self.other_entry_label.pack_forget()
            self.other_entry.pack_forget()

    def add_criteria(self):
        criterion = self.criteria_var.get()
        if criterion == "Diğer":
            criterion = self.other_entry.get().strip()
            if not criterion:
                messagebox.showerror("Hata", "Lütfen 'Diğer' için bir kriter adı girin.")
                return
        
        try:
            weight = float(self.weight_entry.get())
            if weight <= 0 or weight > 100:
                raise ValueError
        except ValueError:
            messagebox.showerror("Hata", "Geçerli bir ağırlık girin (1-100 arası).")
            return

        # Eğer aynı isimde kriter varsa, numaralandır
        base_criterion = criterion
        counter = 1
        while criterion in self.criteria:
            counter += 1
            criterion = f"{base_criterion} {counter}"

        self.criteria.append(criterion)
        self.weights[criterion] = weight
        self.criteria_listbox.insert(tk.END, f"{criterion}: {weight}%")

        self.criteria_var.set("")
        self.weight_entry.delete(0, tk.END)
        self.other_entry.delete(0, tk.END)
        self.other_entry_label.pack_forget()
        self.other_entry.pack_forget()

    def edit_selected_criteria(self):
        try:
            selected_index = self.criteria_listbox.curselection()[0]
            selected_item = self.criteria_listbox.get(selected_index)
            criterion, weight = selected_item.split(": ")
            weight = float(weight.strip('%'))

            # Kriteri combobox'a yerleştir
            base_criterion = criterion.split()[0] if len(criterion.split()) > 1 else criterion
            self.criteria_var.set(base_criterion if base_criterion in self.criteria_combobox['values'] else "Diğer")
            
            # Eğer "Diğer" seçiliyse, other_entry'ye yerleştir
            if self.criteria_var.get() == "Diğer":
                self.other_entry_label.pack(side="left", padx=5)
                self.other_entry.pack(side="left", padx=5)
                self.other_entry.delete(0, tk.END)
                self.other_entry.insert(0, criterion)

            # Ağırlığı weight_entry'ye yerleştir
            self.weight_entry.delete(0, tk.END)
            self.weight_entry.insert(0, weight)

            # Kriteri listeden ve kayıtlardan sil
            del self.weights[criterion]
            self.criteria.remove(criterion)
            self.criteria_listbox.delete(selected_index)

        except IndexError:
            messagebox.showerror("Hata", "Düzenlemek için bir kriter seçin.")

    def save_criteria_to_db(self):
        if not self.selected_course.get():
            messagebox.showerror("Hata", "Lütfen bir ders seçin.")
            return

        total_weight = sum(self.weights.values())
        if total_weight != 100:
            messagebox.showerror("Hata", "Ağırlıkların toplamı 100 olmalıdır. Şu anki toplam: {}%".format(total_weight))
            return

        try:
            # Ders ID'sini al
            self.cursor.execute("SELECT id FROM courses WHERE course_name = ?", (self.selected_course.get(),))
            course_id = self.cursor.fetchone()[0]

            # Seçili dersin mevcut kriterlerini sil
            self.cursor.execute("DELETE FROM criteria WHERE course_id = ?", (course_id,))
            
            # Yeni kriterleri ekle
            for criterion, weight in self.weights.items():
                self.cursor.execute("""
                    INSERT INTO criteria (course_id, criterion, weight) 
                    VALUES (?, ?, ?)
                """, (course_id, criterion, weight))
            
            self.conn.commit()
            messagebox.showinfo("Başarılı", "Değerlendirme kriterleri başarıyla kaydedildi.")
            self.create_main_menu()
        except Exception as e:
            messagebox.showerror("Hata", f"Kriterler kaydedilemedi: {e}")

    def load_criteria_from_db(self, event=None):
        """Seçili derse ait kriterleri yükle"""
        self.criteria_listbox.delete(0, tk.END)
        self.criteria = []
        self.weights = {}

        if not self.selected_course.get():
            return

        try:
            self.cursor.execute("""
                SELECT criterion, weight 
                FROM criteria 
                WHERE course_id = (SELECT id FROM courses WHERE course_name = ?)
                ORDER BY id
            """, (self.selected_course.get(),))
            
            for criterion, weight in self.cursor.fetchall():
                self.criteria.append(criterion)
                self.weights[criterion] = weight
                self.criteria_listbox.insert(tk.END, f"{criterion}: {weight}%")
        except Exception as e:
            messagebox.showerror("Hata", f"Kriterler yüklenirken hata oluştu: {e}")

    def clear_frame(self):
        for widget in self.root.winfo_children():
            widget.destroy()

    def upload_data(self):
        self.clear_frame()
        self.upload_frame = ttk.Frame(self.root)
        self.upload_frame.pack(fill="both", expand=True)

        ttk.Label(self.upload_frame, text="Veri Yükle", font=("Arial", 16)).pack(pady=20)

        ttk.Button(self.upload_frame, text="Öğrenci Listesi Yükle", command=self.load_student_list).pack(pady=10)
        ttk.Button(self.upload_frame, text="Program Çıktıları Yükle", command=self.load_program_outcomes).pack(pady=10)
        ttk.Button(self.upload_frame, text="Ana Menüye Dön", command=self.create_main_menu).pack(pady=10)

    def load_student_list(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx"), ("All Files", "*.*")])
        if file_path:
            try:
                self.student_list = pd.read_excel(file_path)
                messagebox.showinfo("Başarılı", "Öğrenci listesi yüklendi.")
            except Exception as e:
                messagebox.showerror("Hata", f"Dosya yüklenemedi: {e}")

    def create_tables(self):
        self.clear_frame()
        self.table_frame = ttk.Frame(self.root)
        self.table_frame.pack(fill="both", expand=True)

        ttk.Label(self.table_frame, text="Tablo 4 ve 5 Oluştur", font=("Arial", 16)).pack(pady=20)

        ttk.Button(self.table_frame, text="Tablo 4 Oluştur", command=self.generate_table_4).pack(pady=10)
        ttk.Button(self.table_frame, text="Tablo 5 Oluştur", command=self.generate_table_5).pack(pady=10)
        ttk.Button(self.table_frame, text="Ana Menüye Dön", command=self.create_main_menu).pack(pady=10)

    def generate_table_4(self):
        if hasattr(self, 'student_list'):
            try:
                table_4 = self.student_list.copy()
                table_4.to_excel("Tablo_4.xlsx", index=False)
                messagebox.showinfo("Başarılı", "Tablo 4 oluşturuldu ve kaydedildi.")
            except Exception as e:
                messagebox.showerror("Hata", f"Tablo oluşturulamadı: {e}")
        else:
            messagebox.showerror("Hata", "Öncelikle öğrenci listesini yükleyin.")

    def generate_table_5(self):
        if hasattr(self, 'student_list'):
            try:
                table_5 = self.student_list.copy()
                table_5.to_excel("Tablo_5.xlsx", index=False)
                messagebox.showinfo("Başarılı", "Tablo 5 oluşturuldu ve kaydedildi.")
            except Exception as e:
                messagebox.showerror("Hata", f"Tablo oluşturulamadı: {e}")
        else:
            messagebox.showerror("Hata", "Öncelikle öğrenci listesini yükleyin.")

    def clear_frame(self):
        for widget in self.root.winfo_children():
            widget.destroy()

    def validate_input(self, P):
        if P == "":
            return True
        try:
            value = float(P)
            return 0 <= value <= 1
        except ValueError:
            return False

    def table1_screen(self):
        self.clear_frame()
        self.table1_frame = ttk.Frame(self.root)
        self.table1_frame.pack(fill="both", expand=True, padx=20, pady=20)

        # Başlık
        ttk.Label(self.table1_frame, text="Tablo 1: Program Çıktıları - Ders Çıktıları İlişkisi", 
                 font=("Arial", 16)).pack(pady=20)

        # Ders seçimi
        select_frame = ttk.Frame(self.table1_frame)
        select_frame.pack(fill="x", padx=10)
        
        ttk.Label(select_frame, text="Ders:").pack(side="left", padx=5)
        self.selected_course = tk.StringVar()
        self.cursor.execute("SELECT course_name FROM courses")
        courses = [course[0] for course in self.cursor.fetchall()]
        course_combo = ttk.Combobox(select_frame, textvariable=self.selected_course, 
                                   values=courses, width=30)
        course_combo.pack(side="left", padx=5)
        course_combo.bind('<<ComboboxSelected>>', self.update_table1)

        # Tablo frame
        self.table1_grid_frame = ttk.Frame(self.table1_frame)
        self.table1_grid_frame.pack(fill="both", expand=True, pady=10, padx=10)

        # Buton frame
        button_frame = ttk.Frame(self.table1_frame)
        button_frame.pack(fill="x", pady=10)
        ttk.Button(button_frame, text="Kaydet", 
                   command=self.save_table1).pack(side="left", padx=5)
        ttk.Button(button_frame, text="İçeri Aktar",
                   command=self.load_table1_excel).pack(side="left", padx=5)
        ttk.Button(button_frame, text="Dışarı Aktar",
                   command=self.export_table1_excel).pack(side="left", padx=5)
        ttk.Button(button_frame, text="Ana Menüye Dön", 
                   command=self.create_main_menu).pack(side="left", padx=5)

    def update_table1(self, event=None):
        # Mevcut grid'i temizle
        for widget in self.table1_grid_frame.winfo_children():
            widget.destroy()

        if not self.selected_course.get():
            return

        # Program çıktılarını al
        self.cursor.execute("SELECT id, outcome FROM program_outcomes")
        program_outcomes = self.cursor.fetchall()

        # Seçilen dersin ders çıktılarını al
        self.cursor.execute("""
            SELECT co.id, co.outcome 
            FROM course_outcomes co 
            JOIN courses c ON co.course_id = c.id 
            WHERE c.course_name = ?
        """, (self.selected_course.get(),))
        course_outcomes = self.cursor.fetchall()

        if not course_outcomes or not program_outcomes:
            messagebox.showwarning("Uyarı", "Ders çıktıları veya program çıktıları bulunamadı.")
            return

        # Grid başlangıç pozisyonu
        start_row = 0

        # Validation register
        vcmd = (self.root.register(self.validate_input), '%P')

        # Başlık satırı - Program Çıktıları ve İlişki Değeri
        ttk.Label(self.table1_grid_frame, text="Program Çıktıları / Ders Çıktıları", 
                 font=("Arial", 10, "bold")).grid(row=start_row, column=0, padx=5, pady=5)
        
        # İlişki Değeri başlığı
        ttk.Label(self.table1_grid_frame, text="İlişki Değ.", 
                 font=("Arial", 10, "bold")).grid(row=start_row, column=len(course_outcomes)+1, padx=5, pady=5)
        
        # Program Çıktıları başlıkları ve ilişki değeri etiketleri
        self.relation_labels = []
        for row, po in enumerate(program_outcomes, start=1):
            ttk.Label(self.table1_grid_frame, text=f"PÇ{row}", 
                     font=("Arial", 10, "bold")).grid(row=start_row+row, column=0, padx=5, pady=2)
            
            # Her program çıktısı için ilişki değeri etiketi
            relation_label = ttk.Label(self.table1_grid_frame, text="0.0")
            relation_label.grid(row=start_row+row, column=len(course_outcomes)+1, padx=5, pady=2)
            self.relation_labels.append(relation_label)

        # Ders Çıktıları başlıkları ve giriş matrisi
        self.table1_entries = []
        for col, co in enumerate(course_outcomes, start=1):
            ttk.Label(self.table1_grid_frame, text=f"ÇK{col}").grid(row=start_row, column=col, padx=5, pady=2)
            
            column_entries = []
            for row, po in enumerate(program_outcomes, start=1):
                entry = ttk.Entry(self.table1_grid_frame, width=8, validate="key", validatecommand=vcmd)
                entry.grid(row=start_row+row, column=col, padx=2, pady=2)
                entry.bind('<KeyRelease>', lambda e, r=row-1: self.update_relation_value(r))
                
                # Mevcut değeri yükle
                self.cursor.execute("""
                    SELECT value FROM relationship_values 
                    WHERE course_id = (SELECT id FROM courses WHERE course_name = ?) 
                    AND program_outcome_id = ? 
                    AND course_outcome_id = ?
                """, (self.selected_course.get(), po[0], co[0]))
                
                value = self.cursor.fetchone()
                if value:
                    entry.insert(0, str(value[0]))
                
                column_entries.append(entry)
            self.table1_entries.append(column_entries)
            
        # İlk ilişki değerlerini hesapla
        for row in range(len(program_outcomes)):
            self.update_relation_value(row)

    def update_relation_value(self, row):
        """Her satır için ilişki değerini günceller"""
        try:
            # Satırdaki tüm değerleri topla
            values = []
            for col in range(len(self.table1_entries)):
                value = self.table1_entries[col][row].get().strip()
                if value:
                    values.append(float(value))
            
            # Ortalamayı hesapla
            if values:
                avg = sum(values) / len(self.table1_entries)  # Toplam ders çıktısı sayısına böl
                self.relation_labels[row].config(text=f"{avg:.2f}")
            else:
                self.relation_labels[row].config(text="0.00")
        except (ValueError, IndexError):
            self.relation_labels[row].config(text="0.00")

    def load_table1_excel(self):
        if not self.selected_course.get():
            messagebox.showerror("Hata", "Lütfen önce bir ders seçin.")
            return

        file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx"), ("All Files", "*.*")])
        if not file_path:
            return

        try:
            df = pd.read_excel(file_path)
            
            # Değerlerin 0-1 arasında olduğunu kontrol et
            if ((df.iloc[:, 1:] < 0) | (df.iloc[:, 1:] > 1)).any().any():
                messagebox.showerror("Hata", "Excel dosyasındaki değerler 0 ile 1 arasında olmalıdır.")
                return

            # Verileri entry'lere aktar
            for i, row in enumerate(self.table1_entries[0]):  # Her satır için
                for j, col in enumerate(self.table1_entries):  # Her sütun için
                    value = df.iloc[i, j+1]  # İlk sütun başlık olduğu için j+1
                    self.table1_entries[j][i].delete(0, tk.END)
                    self.table1_entries[j][i].insert(0, str(value))

            messagebox.showinfo("Başarılı", "Veriler Excel'den yüklendi.")
        except Exception as e:
            messagebox.showerror("Hata", f"Excel dosyası yüklenirken hata oluştu: {e}")

    def save_table1(self):
        if not self.selected_course.get():
            messagebox.showerror("Hata", "Lütfen bir ders seçin.")
            return

        try:
            # Ders ID'sini al
            self.cursor.execute("SELECT id FROM courses WHERE course_name = ?", (self.selected_course.get(),))
            course_id = self.cursor.fetchone()[0]

            # Program ve ders çıktılarının ID'lerini al
            self.cursor.execute("SELECT id FROM program_outcomes ORDER BY id")
            program_outcome_ids = [row[0] for row in self.cursor.fetchall()]

            self.cursor.execute("""
                SELECT co.id 
                FROM course_outcomes co 
                JOIN courses c ON co.course_id = c.id 
                WHERE c.course_name = ?
                ORDER BY co.id
            """, (self.selected_course.get(),))
            course_outcome_ids = [row[0] for row in self.cursor.fetchall()]

            # Mevcut ilişkileri sil
            self.cursor.execute("DELETE FROM relationship_values WHERE course_id = ?", (course_id,))

            # Yeni değerleri kaydet
            for col, co_id in enumerate(course_outcome_ids):
                for row, po_id in enumerate(program_outcome_ids):
                    value = self.table1_entries[col][row].get().strip()
                    if not value:
                        messagebox.showerror("Hata", "Lütfen tüm değerleri doldurun.")
                        return
                    try:
                        float_value = float(value)
                        if not (0 <= float_value <= 1):
                            messagebox.showerror("Hata", "Tüm değerler 0 ile 1 arasında olmalıdır.")
                            return
                        
                        self.cursor.execute("""
                            INSERT INTO relationship_values (course_id, program_outcome_id, course_outcome_id, value)
                            VALUES (?, ?, ?, ?)
                        """, (course_id, po_id, co_id, float_value))
                        
                    except ValueError:
                        messagebox.showerror("Hata", "Geçersiz değer girişi.")
                        return

            self.conn.commit()
            messagebox.showinfo("Başarılı", "İlişki değerleri başarıyla kaydedildi.")
        
            
        except Exception as e:
            messagebox.showerror("Hata", f"Kayıt sırasında hata oluştu: {e}")

    def export_to_excel(self):
        try:
            # Verileri DataFrame'e dönüştür
            self.cursor.execute("SELECT outcome FROM program_outcomes")
            program_outcomes = [f"PÇ{i+1}" for i in range(len(self.cursor.fetchall()))]
            
            self.cursor.execute("""
                SELECT co.outcome 
                FROM course_outcomes co 
                JOIN courses c ON co.course_id = c.id 
                WHERE c.course_name = ?
            """, (self.selected_course.get(),))
            course_outcomes = [f"ÇK{i+1}" for i in range(len(self.cursor.fetchall()))]

            # Ana matrix verilerini hazırla
            matrix_data = []
            relation_values = []
            for row_idx in range(len(program_outcomes)):
                row_data = [float(entry.get()) for col in self.table1_entries for entry in [col[row_idx]]]
            matrix_data.append(row_data)
            # İlişki değerini hesapla
            relation_value = sum(row_data) / len(course_outcomes)
            relation_values.append(relation_value)

            # DataFrame oluştur
            df = pd.DataFrame(matrix_data, index=program_outcomes, columns=course_outcomes)
            # İlişki değerleri sütununu ekle
            df['İlişki Değ.'] = relation_values
            
            # Excel'e kaydet
            filename = f"{self.selected_course.get()}_Tablo1.xlsx"
            df.to_excel(filename)
            messagebox.showinfo("Başarılı", f"Veriler {filename} dosyasına kaydedildi.")
        except Exception as e:
            messagebox.showerror("Hata", f"Excel'e kayıt sırasında hata oluştu: {e}")

    def table2_screen(self):
        self.clear_frame()
        self.table2_frame = ttk.Frame(self.root)
        self.table2_frame.pack(fill="both", expand=True, padx=20, pady=20)

        # Başlık
        ttk.Label(self.table2_frame, text="Tablo 2: Değerlendirmeler - Ders Çıktıları İlişkisi", 
                 font=("Arial", 16)).pack(pady=20)

        # Ders seçimi
        select_frame = ttk.Frame(self.table2_frame)
        select_frame.pack(fill="x", padx=10)
        
        ttk.Label(select_frame, text="Ders:").pack(side="left", padx=5)
        self.selected_course = tk.StringVar()
        self.cursor.execute("SELECT course_name FROM courses")
        courses = [course[0] for course in self.cursor.fetchall()]
        course_combo = ttk.Combobox(select_frame, textvariable=self.selected_course, 
                                   values=courses, width=30)
        course_combo.pack(side="left", padx=5)
        course_combo.bind('<<ComboboxSelected>>', self.update_table2)

        # Tablo frame
        self.table2_grid_frame = ttk.Frame(self.table2_frame)
        self.table2_grid_frame.pack(fill="both", expand=True, pady=10, padx=10)

        # Buton frame
        button_frame = ttk.Frame(self.table2_frame)
        button_frame.pack(fill="x", pady=10)
        ttk.Button(button_frame, text="Kaydet", 
                   command=self.save_table2).pack(side="left", padx=5)
        ttk.Button(button_frame, text="İçeri Aktar", 
                   command=self.load_table2_excel).pack(side="left", padx=5)
        ttk.Button(button_frame, text="Dışarı Aktar", 
                   command=self.export_table2).pack(side="left", padx=5)
        ttk.Button(button_frame, text="Ana Menüye Dön", 
                   command=self.create_main_menu).pack(side="left", padx=5)

    def load_table2_excel(self):
        """Excel'den Tablo 2 verilerini yükle"""
        if not self.selected_course.get():
            messagebox.showerror("Hata", "Lütfen önce bir ders seçin.")
            return

        file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx"), ("All Files", "*.*")])
        if not file_path:
            return

        try:
            df = pd.read_excel(file_path)
            
            # Kriterlerin olduğu sütunları al (ilk iki sütun ve son sütun hariç)
            criteria_columns = df.iloc[:, 2:-1]  
            
            # Sadece kriter sütunlarında değerlerin 0-1 arasında olduğunu kontrol et
            if ((criteria_columns < 0) | (criteria_columns > 1)).any().any():
                messagebox.showerror("Hata", "Excel dosyasındaki değerler 0 ile 1 arasında olmalıdır.")
                return

            # Verileri entry'lere aktar
            row = 0  # Excel satır sayacı
            for outcome_id in self.table2_entries.keys():
                criterion_col = 0  # Excel sütun sayacı
                for criterion_id in self.table2_entries[outcome_id].keys():
                    try:
                        value = df.iloc[row, criterion_col + 2]  # +2 ilk iki sütunu atla
                        if pd.notna(value):  # NaN değilse
                            entry = self.table2_entries[outcome_id][criterion_id]
                            entry.delete(0, tk.END)
                            entry.insert(0, str(value))
                        criterion_col += 1
                    except IndexError:
                        continue
                row += 1

            messagebox.showinfo("Başarılı", "Veriler Excel'den yüklendi.")
            
            # Her satır için toplamları güncelle
            for outcome_id in self.table2_entries.keys():
                self.update_row_total(outcome_id)
                
        except Exception as e:
            messagebox.showerror("Hata", f"Excel dosyası yüklenirken hata oluştu: {e}")

    def update_table2(self, event=None):
        """Tablo 2'yi güncelle"""
        if not self.selected_course.get():
            return

        try:
            # Grid'i temizle
            for widget in self.table2_grid_frame.winfo_children():
                widget.destroy()

            # Seçili dersin kriterlerini al
            self.cursor.execute("""
                SELECT id, criterion, weight 
                FROM criteria 
                WHERE course_id = (SELECT id FROM courses WHERE course_name = ?)
                ORDER BY id
            """, (self.selected_course.get(),))
            criteria = self.cursor.fetchall()

            # Seçili dersin çıktılarını al
            self.cursor.execute("""
                SELECT id, outcome 
                FROM course_outcomes 
                WHERE course_id = (SELECT id FROM courses WHERE course_name = ?)
                ORDER BY id
            """, (self.selected_course.get(),))
            outcomes = self.cursor.fetchall()

            # Mevcut değerleri al
            self.cursor.execute("""
                SELECT course_outcome_id, criterion_id, value
                FROM course_outcome_evaluations
                WHERE course_id = (SELECT id FROM courses WHERE course_name = ?)
            """, (self.selected_course.get(),))
            evaluations = {(row[0], row[1]): row[2] for row in self.cursor.fetchall()}

            # Başlıkları oluştur
            ttk.Label(self.table2_grid_frame, text="Ders Çıktısı", 
                     font=("Arial", 10, "bold")).grid(row=0, column=0, padx=5, pady=5)

            # Değerlendirme kriterleri başlıkları
            for col, (_, criterion, weight) in enumerate(criteria, start=1):
                ttk.Label(self.table2_grid_frame, text=criterion).grid(row=0, column=col, padx=5, pady=5)
                ttk.Label(self.table2_grid_frame, text=f"{weight}%").grid(row=1, column=col, padx=5, pady=2)

            # TOPLAM sütunu
            ttk.Label(self.table2_grid_frame, text="TOPLAM", 
                     font=("Arial", 10, "bold")).grid(row=0, column=len(criteria)+1, padx=5, pady=5)

            # Ders çıktıları ve değerlendirme kutucukları
            self.table2_entries = {}
            self.total_labels = {}  # Toplam değerler için label'lar
            
            for row, (outcome_id, outcome) in enumerate(outcomes, start=2):
                ttk.Label(self.table2_grid_frame, text=f"ÇK{outcome_id}").grid(
                    row=row, column=0, padx=5, pady=2)
                
                row_entries = {}
                row_total = 0
                
                # Her kriter için değeri ekle
                for col, (criterion_id, _, _) in enumerate(criteria, start=1):
                    value = evaluations.get((outcome_id, criterion_id), 0)
                    row_total += value
                    
                    entry = ttk.Entry(self.table2_grid_frame, width=10)
                    entry.insert(0, str(value))
                    entry.grid(row=row, column=col, padx=2, pady=2)
                    entry.bind('<KeyRelease>', lambda e, oid=outcome_id: self.update_row_total(oid))
                    row_entries[criterion_id] = entry
                
                # Toplam label
                total_label = ttk.Label(self.table2_grid_frame, text=f"{row_total:.2f}")
                total_label.grid(row=row, column=len(criteria)+1, padx=5, pady=2)
                self.total_labels[outcome_id] = total_label
                
                self.table2_entries[outcome_id] = row_entries

        except Exception as e:
            messagebox.showerror("Hata", f"Tablo güncellenirken hata oluştu: {str(e)}")

    def update_row_total(self, outcome_id):
        """Satır toplamını güncelle"""
        try:
            total = 0
            for entry in self.table2_entries[outcome_id].values():
                value = entry.get().strip()
                if value:
                    total += float(value)
            self.total_labels[outcome_id].config(text=f"{total:.2f}")
        except ValueError:
            self.total_labels[outcome_id].config(text="0.00")

    def save_table2(self):
        if not self.selected_course.get():
            messagebox.showerror("Hata", "Lütfen önce bir ders seçin.")
            return

        try:
            # Seçili dersin ID'sini al
            self.cursor.execute("SELECT id FROM courses WHERE course_name = ?", 
                              (self.selected_course.get(),))
            course_id = self.cursor.fetchone()[0]

            # Mevcut değerlendirmeleri sil
            self.cursor.execute("""
                DELETE FROM course_outcome_evaluations 
                WHERE course_id = ?
            """, (course_id,))

            # Yeni değerleri kaydet
            for outcome_id, criterion_entries in self.table2_entries.items():
                for criterion_id, entry in criterion_entries.items():
                    value = entry.get().strip()
                    if not value:
                        messagebox.showerror("Hata", "Lütfen tüm değerleri doldurun.")
                        self.conn.rollback()
                        return
                    
                    try:
                        value = float(value)
                        if not (0 <= value <= 1):
                            messagebox.showerror("Hata", "Değerler 0 ile 1 arasında olmalıdır.")
                            self.conn.rollback()
                            return
                    except ValueError:
                        messagebox.showerror("Hata", "Lütfen geçerli sayısal değerler girin.")
                        self.conn.rollback()
                        return

                    self.cursor.execute("""
                        INSERT INTO course_outcome_evaluations 
                        (course_id, course_outcome_id, criterion_id, value)
                        VALUES (?, ?, ?, ?)
                    """, (course_id, outcome_id, criterion_id, value))

            self.conn.commit()
            messagebox.showinfo("Başarılı", "Değerlendirmeler kaydedildi.")

        except Exception as e:
            self.conn.rollback()
            messagebox.showerror("Hata", f"Kayıt sırasında hata oluştu: {str(e)}")

    def export_table2(self):
        """Tablo 2'yi Excel'e aktar"""
        if not self.selected_course.get():
            messagebox.showerror("Hata", "Lütfen önce bir ders seçin.")
            return

        try:
            # Kriterleri al
            self.cursor.execute("""
                SELECT id, criterion, weight 
                FROM criteria 
                WHERE course_id = (SELECT id FROM courses WHERE course_name = ?)
                ORDER BY id
            """, (self.selected_course.get(),))
            criteria = self.cursor.fetchall()

            # Ders çıktılarını al
            self.cursor.execute("""
                SELECT id, outcome 
                FROM course_outcomes 
                WHERE course_id = (SELECT id FROM courses WHERE course_name = ?)
                ORDER BY id
            """, (self.selected_course.get(),))
            outcomes = self.cursor.fetchall()

            # Tablo 2 değerlerini al
            self.cursor.execute("""
                SELECT course_outcome_id, criterion_id, value
                FROM course_outcome_evaluations
                WHERE course_id = (SELECT id FROM courses WHERE course_name = ?)
            """, (self.selected_course.get(),))
            evaluations = {(row[0], row[1]): row[2] for row in self.cursor.fetchall()}

            # Excel için verileri hazırla
            excel_data = []
            
            # Her ders çıktısı için
            for outcome_id, outcome_text in outcomes:
                row_data = [f"ÇK{outcome_id}", outcome_text]
                row_total = 0
                
                # Her kriter için değeri ekle
                for criterion_id, _, _ in criteria:
                    value = evaluations.get((outcome_id, criterion_id), 0)
                    row_data.append(value)
                    row_total += value
                
                # Toplam değeri ekle
                row_data.append(row_total)
                excel_data.append(row_data)

            # Başlıkları hazırla
            headers = ["ÇK No", "Ders Çıktısı"] + [c[1] for c in criteria] + ["TOPLAM"]

            # DataFrame oluştur
            df = pd.DataFrame(excel_data, columns=headers)
            
            # Excel'e kaydet
            filename = f"{self.selected_course.get()}_Tablo2.xlsx"
            df.to_excel(filename, index=False)
            messagebox.showinfo("Başarılı", f"Tablo 2 verileri {filename} dosyasına aktarıldı.")

        except Exception as e:
            messagebox.showerror("Hata", f"Dışa aktarma sırasında hata oluştu: {str(e)}")

    def student_grades_screen(self):
        self.clear_frame()
        self.grades_frame = ttk.Frame(self.root)
        self.grades_frame.pack(fill="both", expand=True, padx=20, pady=20)

        # Başlık
        ttk.Label(self.grades_frame, text="Öğrenci Notları", 
                 font=("Arial", 16)).pack(pady=20)

        # Ders seçimi
        select_frame = ttk.Frame(self.grades_frame)
        select_frame.pack(fill="x", padx=10)
        
        ttk.Label(select_frame, text="Ders Seçin:").pack(side="left", padx=5)
        self.selected_course = tk.StringVar()
        self.cursor.execute("SELECT course_name FROM courses")
        courses = [course[0] for course in self.cursor.fetchall()]
        course_combo = ttk.Combobox(select_frame, textvariable=self.selected_course, 
                                   values=courses, width=30)
        course_combo.pack(side="left", padx=5)
        course_combo.bind('<<ComboboxSelected>>', self.load_grades_grid)

        # Tablo grid için frame
        self.grades_grid_frame = ttk.Frame(self.grades_frame)
        self.grades_grid_frame.pack(fill="both", expand=True, pady=10, padx=10)

        # Buton frame
        button_frame = ttk.Frame(self.grades_frame)
        button_frame.pack(fill="x", pady=10, side="bottom")
        ttk.Button(button_frame, text="Öğrenci Ekle", 
                   command=self.add_student_row).pack(side="left", padx=5)
        ttk.Button(button_frame, text="Excel'den Yükle", 
                   command=self.load_grades_excel).pack(side="left", padx=5)
        ttk.Button(button_frame, text="Kaydet", 
                   command=self.save_grades).pack(side="left", padx=5)
        ttk.Button(button_frame, text="Dışarı Aktar", 
                   command=self.export_grades_excel).pack(side="left", padx=5)
        ttk.Button(button_frame, text="Ana Menüye Dön", 
                   command=self.create_main_menu).pack(side="left", padx=5)

    def load_grades_grid(self, event=None):
        # Grid'i temizle
        for widget in self.grades_grid_frame.winfo_children():
            widget.destroy()

        if not self.selected_course.get():
            return

        # Değerlendirme kriterlerini al
        self.cursor.execute("""
            SELECT id, criterion, weight 
            FROM criteria 
            WHERE course_id = (SELECT id FROM courses WHERE course_name = ?)
            ORDER BY id
        """, (self.selected_course.get(),))
        self.criteria = self.cursor.fetchall()

        if not self.criteria:
            messagebox.showwarning("Uyarı", "Bu ders için değerlendirme kriterleri bulunamadı.")
            return

        # Başlıklar
        ttk.Label(self.grades_grid_frame, text="Öğrenci", 
                 font=("Arial", 10, "bold")).grid(row=0, column=0, padx=5, pady=5)
        
        # Kriter başlıkları
        for col, (_, criterion, _) in enumerate(self.criteria, start=1):
            ttk.Label(self.grades_grid_frame, text=criterion).grid(row=0, column=col, padx=5, pady=5)
        
        # ORT sütunu
        ttk.Label(self.grades_grid_frame, text="ORT", 
                 font=("Arial", 10, "bold")).grid(row=0, column=len(self.criteria)+1, padx=5, pady=5)

        # Değişkenleri hazırla
        self.student_entries = []
        self.grade_entries = []
        self.average_labels = []
        self.current_row = 1

        # Veritabanından mevcut öğrenci ve notları yükle
        try:
            self.cursor.execute("""
                SELECT student_name, grades 
                FROM student_grades 
                WHERE course_id = (SELECT id FROM courses WHERE course_name = ?)
                ORDER BY student_name
            """, (self.selected_course.get(),))
            
            students = self.cursor.fetchall()
            
            # Her öğrenci için notları yükle
            for student_name, grades_json in students:
                self.add_student_row()  # Yeni satır ekle
                
                # Öğrenci adını yaz
                self.student_entries[-1].insert(0, student_name)
                
                # Notları yükle
                grades_dict = json.loads(grades_json)
                
                # Her kriter için notu yerleştir
                for col, (crit_id, _, _) in enumerate(self.criteria):
                    grade = grades_dict.get(str(crit_id), "")
                    if grade:
                        self.grade_entries[-1][col].insert(0, str(grade))
                
                # Ortalamayı hesapla
                self.update_student_average(len(self.student_entries)-1)
                
        except Exception as e:
            messagebox.showerror("Hata", f"Notlar yüklenirken hata oluştu: {str(e)}")

    def add_student_row(self):
        """Yeni öğrenci satırı ekle"""
        if not hasattr(self, 'criteria') or not self.criteria:
            messagebox.showerror("Hata", "Lütfen önce bir ders seçin.")
            return

        # Öğrenci adı
        name_entry = ttk.Entry(self.grades_grid_frame, width=20)
        name_entry.grid(row=self.current_row, column=0, padx=5, pady=2)
        
        # Not girişleri
        row_entries = []
        vcmd = (self.root.register(self.validate_grade), '%P')
        
        for col in range(len(self.criteria)):
            entry = ttk.Entry(self.grades_grid_frame, width=8, validate="key", validatecommand=vcmd)
            entry.grid(row=self.current_row, column=col+1, padx=2, pady=2)
            entry.bind('<KeyRelease>', lambda e, r=len(self.student_entries): self.update_student_average(r))
            row_entries.append(entry)

        # Ortalama label
        avg_label = ttk.Label(self.grades_grid_frame, text="0.00")
        avg_label.grid(row=self.current_row, column=len(self.criteria)+1, padx=5, pady=2)

        self.student_entries.append(name_entry)
        self.grade_entries.append(row_entries)
        self.average_labels.append(avg_label)
        self.current_row += 1

    def validate_grade(self, value):
        """Not girişi için validation (0-100 arası)"""
        if value == "":
            return True
        try:
            grade = float(value)
            return 0 <= grade <= 100
        except ValueError:
            return False

    def update_student_average(self, student_index):
        """Öğrenci not ortalamasını hesapla"""
        try:
            total_weight = 0
            weighted_sum = 0
            
            # Her kriter için
            for i, (_, _, weight) in enumerate(self.criteria):
                grade = self.grade_entries[student_index][i].get().strip()
                if grade:  # Eğer not girilmişse
                    grade_value = float(grade)
                    weighted_sum += grade_value * (weight/100)  # Ağırlıklı toplam
                    total_weight += weight/100  # Toplam ağırlık
            
            # Ortalamayı hesapla
            if total_weight > 0:
                average = weighted_sum
                self.average_labels[student_index].config(text=f"{average:.2f}")
            else:
                self.average_labels[student_index].config(text="0.00")
                
        except ValueError:
            self.average_labels[student_index].config(text="0.00")

    def save_grades(self):
        """Notları veritabanına kaydet"""
        if not self.selected_course.get():
            messagebox.showerror("Hata", "Lütfen bir ders seçin.")
            return

        try:
            # Ders ID'sini al
            self.cursor.execute("SELECT id FROM courses WHERE course_name = ?", 
                              (self.selected_course.get(),))
            course_id = self.cursor.fetchone()[0]

            # Mevcut notları sil
            self.cursor.execute("DELETE FROM student_grades WHERE course_id = ?", (course_id,))

            # Yeni notları kaydet
            for i, name_entry in enumerate(self.student_entries):
                student_name = name_entry.get().strip()
                if not student_name:
                    continue

                # Öğrencinin tüm notlarını bir sözlükte topla
                grades_dict = {}
                for j, (criterion_id, criterion, _) in enumerate(self.criteria):
                    grade = self.grade_entries[i][j].get().strip()
                    if grade:
                        grades_dict[str(criterion_id)] = float(grade)

                # Notları JSON formatında kaydet
                grades_json = json.dumps(grades_dict)
                
                self.cursor.execute("""
                    INSERT INTO student_grades (course_id, student_name, grades)
                    VALUES (?, ?, ?)
                """, (course_id, student_name, grades_json))

            self.conn.commit()
            messagebox.showinfo("Başarılı", "Notlar başarıyla kaydedildi.")
        except Exception as e:
            messagebox.showerror("Hata", f"Notlar kaydedilirken hata oluştu: {str(e)}")

    def load_grades_excel(self):
        """Excel'den notları yükle"""
        if not self.selected_course.get():
            messagebox.showerror("Hata", "Lütfen önce bir ders seçin.")
            return

        file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx"), ("All Files", "*.*")])
        if not file_path:
            return

        try:
            df = pd.read_excel(file_path)
            df = df.fillna(0)  # Boş değerleri 0 ile doldur
            
            # Başlıkları kontrol et
            required_columns = ["Öğrenci"] + [criterion for _, criterion, _ in self.criteria]
            missing_columns = [col for col in required_columns if col not in df.columns]
            
            if missing_columns:
                messagebox.showerror("Hata", 
                                   f"Excel dosyasında eksik sütunlar var: {', '.join(missing_columns)}")
                return

            # Mevcut öğrenci listesini al
            existing_students = [entry.get().strip() for entry in self.student_entries]

            # Excel'deki her öğrenci için
            for _, row in df.iterrows():
                student_name = str(row["Öğrenci"]).strip()
                
                # Öğrenci zaten varsa, notlarını güncelle
                if student_name in existing_students:
                    student_index = existing_students.index(student_name)
                    
                    # Notları güncelle
                    for col, (_, criterion, _) in enumerate(self.criteria):
                        try:
                            grade = float(row[criterion])
                            if 0 <= grade <= 100:
                                self.grade_entries[student_index][col].delete(0, tk.END)
                                self.grade_entries[student_index][col].insert(0, str(grade))
                        except (ValueError, TypeError):
                            continue
                    
                    # Ortalamayı güncelle
                    self.update_student_average(student_index)
                
                # Yeni öğrenci ise, en alta ekle
                else:
                    self.add_student_row()
                    self.student_entries[-1].insert(0, student_name)
                    
                    for col, (_, criterion, _) in enumerate(self.criteria):
                        try:
                            grade = float(row[criterion])
                            if 0 <= grade <= 100:
                                self.grade_entries[-1][col].insert(0, str(grade))
                        except (ValueError, TypeError):
                            continue
                    
                    # Yeni öğrencinin ortalamasını hesapla
                    self.update_student_average(len(self.student_entries)-1)
                    
                    # Mevcut öğrenci listesini güncelle
                    existing_students.append(student_name)

            messagebox.showinfo("Başarılı", "Notlar Excel'den başarıyla yüklendi.")
        except Exception as e:
            messagebox.showerror("Hata", f"Excel dosyası yüklenirken hata oluştu: {str(e)}")

    def export_grades_excel(self):
        """Notları Excel dosyasına aktar"""
        if not self.selected_course.get():
            messagebox.showerror("Hata", "Lütfen önce bir ders seçin.")
            return

        try:
            # Başlıkları hazırla
            headers = ["Öğrenci"] + [criterion for _, criterion, _ in self.criteria] + ["Ortalama"]
            
            # Verileri hazırla
            data = []
            for i in range(len(self.student_entries)):
                row = [self.student_entries[i].get().strip()]  # Öğrenci adı
                
                # Notlar
                for j in range(len(self.criteria)):
                    grade = self.grade_entries[i][j].get().strip()
                    row.append(float(grade) if grade else 0)
                
                # Ortalama
                row.append(float(self.average_labels[i].cget("text")))
                
                data.append(row)
            
            # DataFrame oluştur
            df = pd.DataFrame(data, columns=headers)
            
            # Excel'e kaydet
            filename = f"{self.selected_course.get()}_Notlar.xlsx"
            df.to_excel(filename, index=False)
            messagebox.showinfo("Başarılı", f"Veriler {filename} dosyasına aktarıldı.")
        except Exception as e:
            messagebox.showerror("Hata", f"Dışa aktarma sırasında hata oluştu: {str(e)}")

    def table4_screen(self):
        """Tablo 4 ekranını oluştur"""
        self.clear_frame()
        self.table4_frame = ttk.Frame(self.root)
        self.table4_frame.pack(fill="both", expand=True, padx=20, pady=20)

        # Başlık
        ttk.Label(self.table4_frame, text="Tablo 4: Ders Çıktıları Başarı Oranları", 
                 font=("Arial", 16)).pack(pady=20)

        # Ders ve öğrenci seçimi
        select_frame = ttk.Frame(self.table4_frame)
        select_frame.pack(fill="x", padx=10)
        
        # Ders seçimi
        ttk.Label(select_frame, text="Ders:").pack(side="left", padx=5)
        self.selected_course = tk.StringVar()
        self.cursor.execute("SELECT course_name FROM courses")
        courses = [course[0] for course in self.cursor.fetchall()]
        course_combo = ttk.Combobox(select_frame, textvariable=self.selected_course, 
                                   values=courses, width=30)
        course_combo.pack(side="left", padx=5)
        course_combo.bind('<<ComboboxSelected>>', self.on_course_select_table4)

        # Öğrenci seçimi
        ttk.Label(select_frame, text="Öğrenci:").pack(side="left", padx=5)
        self.selected_student = tk.StringVar()
        self.student_combo = ttk.Combobox(select_frame, textvariable=self.selected_student, 
                                         width=30, state="readonly")
        self.student_combo.pack(side="left", padx=5)
        self.student_combo.bind('<<ComboboxSelected>>', self.calculate_table4)

        # Tablo frame
        self.table4_grid_frame = ttk.Frame(self.table4_frame)
        self.table4_grid_frame.pack(fill="both", expand=True, pady=10, padx=10)

        # Buton frame
        button_frame = ttk.Frame(self.table4_frame)
        button_frame.pack(fill="x", pady=10)
        ttk.Button(button_frame, text="Kaydet", 
                   command=self.save_table4).pack(side="left", padx=5)
        ttk.Button(button_frame, text="Ana Menüye Dön", 
                   command=self.create_main_menu).pack(side="left", padx=5)

    def on_course_select_table4(self, event=None):
        """Ders seçildiğinde öğrenci listesini güncelle"""
        if not self.selected_course.get():
            return

        # Öğrenci listesini al
        self.cursor.execute("""
            SELECT DISTINCT student_name 
            FROM student_grades 
            WHERE course_id = (SELECT id FROM courses WHERE course_name = ?)
            ORDER BY student_name
        """, (self.selected_course.get(),))
        
        students = [student[0] for student in self.cursor.fetchall()]
        self.student_combo['values'] = students
        self.student_combo['state'] = 'readonly' if students else 'disabled'
        self.selected_student.set('')

    def calculate_table4(self, event=None):
        """Tablo 4'ü hesapla ve göster"""
        if not self.selected_course.get() or not self.selected_student.get():
            return

        try:
            # Grid'i temizle
            for widget in self.table4_grid_frame.winfo_children():
                widget.destroy()

            # Kriterleri ve ağırlıklarını al
            self.cursor.execute("""
                SELECT id, criterion, weight 
                FROM criteria 
                WHERE course_id = (SELECT id FROM courses WHERE course_name = ?)
                ORDER BY id
            """, (self.selected_course.get(),))
            criteria = self.cursor.fetchall()

            # Ders çıktılarını al
            self.cursor.execute("""
                SELECT id, outcome 
                FROM course_outcomes 
                WHERE course_id = (SELECT id FROM courses WHERE course_name = ?)
                ORDER BY id
            """, (self.selected_course.get(),))
            outcomes = self.cursor.fetchall()

            # Başlıkları oluştur
            ttk.Label(self.table4_grid_frame, text="Çıktı No", 
                     font=("Arial", 10, "bold")).grid(row=0, column=0, padx=5, pady=5)
            
            # Kriter başlıkları
            for col, (_, criterion, weight) in enumerate(criteria, start=1):
                ttk.Label(self.table4_grid_frame, text=f"{criterion}\n({weight}%)").grid(
                    row=0, column=col, padx=5, pady=5)
            
            # TOPLAM, MAX ve %Başarı başlıkları
            ttk.Label(self.table4_grid_frame, text="TOPLAM").grid(
                row=0, column=len(criteria)+1, padx=5, pady=5)
            ttk.Label(self.table4_grid_frame, text="MAX").grid(
                row=0, column=len(criteria)+2, padx=5, pady=5)
            ttk.Label(self.table4_grid_frame, text="%Başarı").grid(
                row=0, column=len(criteria)+3, padx=5, pady=5)

            # Tablo 2'deki değerleri al
            self.cursor.execute("""
                SELECT course_outcome_id, criterion_id, value
                FROM course_outcome_evaluations
                WHERE course_id = (SELECT id FROM courses WHERE course_name = ?)
            """, (self.selected_course.get(),))
            table2_values = {(row[0], row[1]): row[2] for row in self.cursor.fetchall()}

            # Seçili öğrencinin notlarını al
            self.cursor.execute("""
                SELECT grades
                FROM student_grades
                WHERE course_id = (SELECT id FROM courses WHERE course_name = ?)
                AND student_name = ?
            """, (self.selected_course.get(), self.selected_student.get()))
            
            grades = json.loads(self.cursor.fetchone()[0])

            # Her çıktı için satır oluştur
            for row, (outcome_id, _) in enumerate(outcomes, start=1):
                # Çıktı numarası
                ttk.Label(self.table4_grid_frame, text=f"{row}").grid(
                    row=row, column=0, padx=5, pady=2)
                
                total = 0
                max_total = 0
                
                # Her kriter için değer hesapla
                for col, (criterion_id, _, weight) in enumerate(criteria, start=1):
                    # Tablo 2'den değeri al
                    t2_value = table2_values.get((outcome_id, criterion_id), 0)
                    # Öğrenci notunu al
                    student_grade = float(grades.get(str(criterion_id), 0))
                    
                    # Değeri hesapla (Tablo2 değeri * kriter ağırlığı/100 * öğrenci notu)
                    value = t2_value * (weight/100) * student_grade
                    total += value
                    # Maximum değer (not 100 olduğunda)
                    max_value = t2_value * (weight/100) * 100  # not 100 olduğundaki değer
                    max_total += max_value
                    
                    # Değeri göster
                    ttk.Label(self.table4_grid_frame, text=f"{value:.2f}").grid(
                        row=row, column=col, padx=2, pady=2)
                
                # TOPLAM
                ttk.Label(self.table4_grid_frame, text=f"{total:.2f}").grid(
                    row=row, column=len(criteria)+1, padx=5, pady=2)
                
                # MAX
                ttk.Label(self.table4_grid_frame, text=f"{max_total:.2f}").grid(
                    row=row, column=len(criteria)+2, padx=5, pady=2)
                
                # %Başarı
                success_rate = (total * 100 / max_total) if max_total > 0 else 0
                ttk.Label(self.table4_grid_frame, text=f"{success_rate:.1f}",
                         foreground="red").grid(row=row, column=len(criteria)+3, padx=5, pady=2)

        except Exception as e:
            messagebox.showerror("Hata", f"Hesaplama sırasında hata oluştu: {str(e)}")

    def export_table4(self):
        """Tablo 4'ü Excel'e aktar"""
        if not self.selected_course.get():
            messagebox.showerror("Hata", "Lütfen önce bir ders seçin.")
            return

        try:
            # Kriterleri al
            self.cursor.execute("""
                SELECT id, criterion, weight 
                FROM criteria 
                WHERE course_id = (SELECT id FROM courses WHERE course_name = ?)
                ORDER BY id
            """, (self.selected_course.get(),))
            criteria = self.cursor.fetchall()

            # Ders çıktılarını al
            self.cursor.execute("""
                SELECT id, outcome 
                FROM course_outcomes 
                WHERE course_id = (SELECT id FROM courses WHERE course_name = ?)
                ORDER BY id
            """, (self.selected_course.get(),))
            outcomes = self.cursor.fetchall()

            # Tablo 2'deki değerleri al
            self.cursor.execute("""
                SELECT course_outcome_id, criterion_id, value
                FROM course_outcome_evaluations
                WHERE course_id = (SELECT id FROM courses WHERE course_name = ?)
            """, (self.selected_course.get(),))
            table2_values = {(row[0], row[1]): row[2] for row in self.cursor.fetchall()}

            # Tüm öğrencilerin notlarını al
            self.cursor.execute("""
                SELECT student_name, grades
                FROM student_grades
                WHERE course_id = (SELECT id FROM courses WHERE course_name = ?)
                ORDER BY student_name
            """, (self.selected_course.get(),))
            students_grades = self.cursor.fetchall()

            # Excel için verileri hazırla
            excel_data = []
            
            # Her öğrenci için
            for student_name, grades in students_grades:
                grades_dict = json.loads(grades)
                
                # Her çıktı için
                for outcome_id, outcome in outcomes:
                    row_data = [student_name, f"ÇK{outcome_id}"]
                    total = 0
                    max_total = 0
                    
                    # Her kriter için değer hesapla
                    for criterion_id, criterion, weight in criteria:
                        t2_value = table2_values.get((outcome_id, criterion_id), 0)
                        student_grade = float(grades_dict.get(str(criterion_id), 0))
                        
                        value = t2_value * (weight/100) * student_grade
                        max_value = t2_value * (weight/100) * 100
                        
                        row_data.append(f"{value:.2f}")
                        total += value
                        max_total += max_value
                    
                    # Toplam, Max ve Başarı değerlerini ekle
                    row_data.extend([
                        f"{total:.2f}",
                        f"{max_total:.2f}",
                        f"{(total * 100 / max_total):.1f}" if max_total > 0 else "0.0"
                    ])
                    
                    excel_data.append(row_data)

            # Başlıkları hazırla
            headers = ["Öğrenci", "Çıktı No"] + [c[1] for c in criteria] + ["TOPLAM", "MAX", "%Başarı"]

            # DataFrame oluştur
            df = pd.DataFrame(excel_data, columns=headers)
            
            # Excel'e kaydet
            filename = f"{self.selected_course.get()}_Tablo4.xlsx"
            df.to_excel(filename, index=False)
            messagebox.showinfo("Başarılı", f"Tablo 4 verileri {filename} dosyasına aktarıldı.")

        except Exception as e:
            messagebox.showerror("Hata", f"Dışa aktarma sırasında hata oluştu: {str(e)}")

    def export_table1_excel(self):
        """Tablo 1'i Excel'e aktar"""
        if not self.selected_course.get():
            messagebox.showerror("Hata", "Lütfen önce bir ders seçin.")
            return

        try:
            # Program çıktılarını al
            self.cursor.execute("SELECT id FROM program_outcomes ORDER BY id")
            program_outcomes = self.cursor.fetchall()

            # Ders çıktılarını al
            self.cursor.execute("""
                SELECT id 
                FROM course_outcomes 
                WHERE course_id = (SELECT id FROM courses WHERE course_name = ?)
                ORDER BY id
            """, (self.selected_course.get(),))
            course_outcomes = self.cursor.fetchall()

            # Excel için verileri hazırla
            data = []
            headers = ["PÇ No"]  # İlk sütun başlığı
            
            # Ders çıktısı başlıklarını ekle
            for co_id, in course_outcomes:
                headers.append(f"ÇK{co_id}")
            headers.append("İlişki Değ.")  # Son sütun başlığı

            # Her program çıktısı için satır oluştur
            for i, (po_id,) in enumerate(program_outcomes):
                row = [f"PÇ{po_id}"]  # İlk sütun
                
                # Her ders çıktısı için değeri al
                for j in range(len(course_outcomes)):
                    value = self.table1_entries[j][i].get().strip()
                    row.append(float(value) if value else 0)
                
                # İlişki değerini ekle
                relation_value = self.relation_labels[i].cget("text")
                row.append(float(relation_value))
                
                data.append(row)

            # DataFrame oluştur ve Excel'e kaydet
            df = pd.DataFrame(data, columns=headers)
            filename = f"{self.selected_course.get()}_Tablo1.xlsx"
            df.to_excel(filename, index=False)
            messagebox.showinfo("Başarılı", f"Tablo 1 verileri {filename} dosyasına aktarıldı.")

        except Exception as e:
            messagebox.showerror("Hata", f"Dışa aktarma sırasında hata oluştu: {e}")

    def export_table2(self):
        """Tablo 2'yi Excel'e aktar"""
        if not self.selected_course.get():
            messagebox.showerror("Hata", "Lütfen önce bir ders seçin.")
            return

        try:
            # Kriterleri al
            self.cursor.execute("""
                SELECT id, criterion, weight 
                FROM criteria 
                WHERE course_id = (SELECT id FROM courses WHERE course_name = ?)
                ORDER BY id
            """, (self.selected_course.get(),))
            criteria = self.cursor.fetchall()

            # Ders çıktılarını al
            self.cursor.execute("""
                SELECT id, outcome 
                FROM course_outcomes 
                WHERE course_id = (SELECT id FROM courses WHERE course_name = ?)
                ORDER BY id
            """, (self.selected_course.get(),))
            outcomes = self.cursor.fetchall()

            # Tablo 2 değerlerini al
            self.cursor.execute("""
                SELECT course_outcome_id, criterion_id, value
                FROM course_outcome_evaluations
                WHERE course_id = (SELECT id FROM courses WHERE course_name = ?)
            """, (self.selected_course.get(),))
            evaluations = {(row[0], row[1]): row[2] for row in self.cursor.fetchall()}

            # Excel için verileri hazırla
            excel_data = []
            
            # Her ders çıktısı için
            for outcome_id, outcome_text in outcomes:
                row_data = [f"ÇK{outcome_id}", outcome_text]
                row_total = 0
                
                # Her kriter için değeri ekle
                for criterion_id, _, _ in criteria:
                    value = evaluations.get((outcome_id, criterion_id), 0)
                    row_data.append(value)
                    row_total += value
                
                # Toplam değeri ekle
                row_data.append(row_total)
                excel_data.append(row_data)

            # Başlıkları hazırla
            headers = ["ÇK No", "Ders Çıktısı"] + [c[1] for c in criteria] + ["TOPLAM"]

            # DataFrame oluştur
            df = pd.DataFrame(excel_data, columns=headers)
            
            # Excel'e kaydet
            filename = f"{self.selected_course.get()}_Tablo2.xlsx"
            df.to_excel(filename, index=False)
            messagebox.showinfo("Başarılı", f"Tablo 2 verileri {filename} dosyasına aktarıldı.")

        except Exception as e:
            messagebox.showerror("Hata", f"Dışa aktarma sırasında hata oluştu: {str(e)}")

    def load_student_grades(self, event=None):
        """Öğrenci notları grid'ini yükle"""
        if not self.selected_course.get():
            return
        
        try:
            # Grid'i temizle
            for widget in self.grades_grid_frame.winfo_children():
                widget.destroy()

            # Kriterleri al
            self.cursor.execute("""
                SELECT id, criterion, weight 
                FROM criteria 
                WHERE course_id = (SELECT id FROM courses WHERE course_name = ?)
                ORDER BY id
            """, (self.selected_course.get(),))
            self.criteria = self.cursor.fetchall()

            # Başlıkları oluştur
            ttk.Label(self.grades_grid_frame, text="Öğrenci", 
                     font=("Arial", 10, "bold")).grid(row=0, column=0, padx=5, pady=5)
            
            # Kriter başlıkları
            for col, (_, criterion, weight) in enumerate(self.criteria, start=1):
                ttk.Label(self.grades_grid_frame, text=f"{criterion}\n(%{weight})").grid(
                    row=0, column=col, padx=5, pady=5)
            
            # Ortalama sütunu
            ttk.Label(self.grades_grid_frame, text="Ortalama", 
                     font=("Arial", 10, "bold")).grid(row=0, column=len(self.criteria)+1, padx=5, pady=5)

            # Entry'leri ve öğrenci listesini oluştur
            self.student_entries = []
            self.grade_entries = []
            self.average_labels = []

            # Validation register
            vcmd = (self.root.register(self.validate_grade), '%P')

            # Mevcut öğrenci notlarını al
            self.cursor.execute("""
                SELECT student_name, grades 
                FROM student_grades 
                WHERE course_id = (SELECT id FROM courses WHERE course_name = ?)
                ORDER BY student_name
            """, (self.selected_course.get(),))
            existing_grades = self.cursor.fetchall()

            for row, (student_name, grades) in enumerate(existing_grades, start=1):
                # Öğrenci adı
                student_entry = ttk.Entry(self.grades_grid_frame, width=20)
                student_entry.insert(0, student_name)
                student_entry.grid(row=row, column=0, padx=5, pady=2)
                self.student_entries.append(student_entry)

                # Not girişleri
                grades_dict = json.loads(grades)
                row_entries = []
                
                for col, (crit_id, _, _) in enumerate(self.criteria, start=1):
                    grade_entry = ttk.Entry(self.grades_grid_frame, width=10, 
                                          validate="key", validatecommand=vcmd)
                    grade = grades_dict.get(str(crit_id), "")
                    grade_entry.insert(0, str(grade))
                    grade_entry.grid(row=row, column=col, padx=2, pady=2)
                    grade_entry.bind('<KeyRelease>', lambda e, r=row-1: self.update_student_average(r))
                    row_entries.append(grade_entry)
                
                self.grade_entries.append(row_entries)

                # Ortalama label
                avg_label = ttk.Label(self.grades_grid_frame, text="0.00")
                avg_label.grid(row=row, column=len(self.criteria)+1, padx=5, pady=2)
                self.average_labels.append(avg_label)
                
                # İlk ortalamayı hesapla
                self.update_student_average(row-1)

        except Exception as e:
            messagebox.showerror("Hata", f"Notlar yüklenirken hata oluştu: {str(e)}")

    def table5_screen(self):
        """Tablo 5 ekranını oluştur"""
        self.clear_frame()
        self.table5_frame = ttk.Frame(self.root)
        self.table5_frame.pack(fill="both", expand=True, padx=20, pady=20)

        # Başlık
        ttk.Label(self.table5_frame, text="Tablo 5: Program Çıktıları Başarı Oranları", 
                 font=("Arial", 16)).pack(pady=20)

        # Ders ve öğrenci seçimi
        select_frame = ttk.Frame(self.table5_frame)
        select_frame.pack(fill="x", padx=10)
        
        # Ders seçimi
        ttk.Label(select_frame, text="Ders:").pack(side="left", padx=5)
        self.selected_course = tk.StringVar()
        self.cursor.execute("SELECT course_name FROM courses")
        courses = [course[0] for course in self.cursor.fetchall()]
        course_combo = ttk.Combobox(select_frame, textvariable=self.selected_course, 
                                   values=courses, width=30)
        course_combo.pack(side="left", padx=5)
        course_combo.bind('<<ComboboxSelected>>', self.on_course_select_table5)

        # Öğrenci seçimi
        ttk.Label(select_frame, text="Öğrenci:").pack(side="left", padx=5)
        self.selected_student = tk.StringVar()
        student_combo = ttk.Combobox(select_frame, textvariable=self.selected_student, 
                                    width=30)
        student_combo.pack(side="left", padx=5)
        student_combo.bind('<<ComboboxSelected>>', self.calculate_table5)

        # Tablo frame
        self.table5_grid_frame = ttk.Frame(self.table5_frame)
        self.table5_grid_frame.pack(fill="both", expand=True, pady=10, padx=10)

        # Buton frame
        button_frame = ttk.Frame(self.table5_frame)
        button_frame.pack(fill="x", pady=10)
        ttk.Button(button_frame, text="Dışarı Aktar", 
                   command=self.export_table5).pack(side="left", padx=5)
        ttk.Button(button_frame, text="Ana Menüye Dön", 
                   command=self.create_main_menu).pack(side="left", padx=5)

    def on_course_select_table5(self, event=None):
        """Ders seçildiğinde öğrenci listesini güncelle"""
        if not self.selected_course.get():
            return

        # Öğrenci listesini güncelle
        self.cursor.execute("""
            SELECT student_name 
            FROM student_grades 
            WHERE course_id = (SELECT id FROM courses WHERE course_name = ?)
            ORDER BY student_name
        """, (self.selected_course.get(),))
        students = [row[0] for row in self.cursor.fetchall()]
        
        # Öğrenci combobox'ını güncelle
        student_combo = self.table5_frame.winfo_children()[1].winfo_children()[3]
        student_combo['values'] = students
        student_combo.set('')

    def calculate_table5(self, event=None):
        """Tablo 5'i hesapla ve göster"""
        if not self.selected_course.get() or not self.selected_student.get():
            return

        try:
            # Grid'i temizle
            for widget in self.table5_grid_frame.winfo_children():
                widget.destroy()

            # Program çıktılarını al
            self.cursor.execute("SELECT id, outcome FROM program_outcomes ORDER BY id")
            program_outcomes = self.cursor.fetchall()

            # Ders çıktılarını al
            self.cursor.execute("""
                SELECT id, outcome 
                FROM course_outcomes 
                WHERE course_id = (SELECT id FROM courses WHERE course_name = ?)
                ORDER BY id
            """, (self.selected_course.get(),))
            course_outcomes = self.cursor.fetchall()

            # Tablo 1'deki ilişki değerlerini al
            self.cursor.execute("""
                SELECT program_outcome_id, course_outcome_id, value
                FROM relationship_values
                WHERE course_id = (SELECT id FROM courses WHERE course_name = ?)
            """, (self.selected_course.get(),))
            relations = {(row[0], row[1]): row[2] for row in self.cursor.fetchall()}

            # Tablo 4'teki başarı değerlerini al
            self.cursor.execute("""
                SELECT course_outcome_id, success_rate
                FROM success_rates
                WHERE course_id = (SELECT id FROM courses WHERE course_name = ?)
                AND student_name = ?
            """, (self.selected_course.get(), self.selected_student.get()))
            success_rates = {row[0]: row[1] for row in self.cursor.fetchall()}

            if not success_rates:
                messagebox.showwarning("Uyarı", "Bu öğrenci için başarı oranları kaydedilmemiş.")
                return

            # Başlıkları oluştur
            ttk.Label(self.table5_grid_frame, text="Program Çıktısı", 
                     font=("Arial", 10, "bold")).grid(row=0, column=0, padx=5, pady=5)
            
            # Ders çıktıları başlıkları
            for col, (co_id, _) in enumerate(course_outcomes, start=1):
                ttk.Label(self.table5_grid_frame, text=f"ÇK{co_id}").grid(
                    row=0, column=col, padx=5, pady=5)

            # Başarı Oranı başlığı
            ttk.Label(self.table5_grid_frame, text="Başarı Oranı", 
                     font=("Arial", 10, "bold")).grid(
                         row=0, column=len(course_outcomes)+1, padx=5, pady=5)

            # Her program çıktısı için hesapla
            for row, (po_id, po_text) in enumerate(program_outcomes, start=1):
                ttk.Label(self.table5_grid_frame, text=f"PÇ{po_id}").grid(
                    row=row, column=0, padx=5, pady=2)
                
                total_relation = 0
                total_success = 0
                
                # Her ders çıktısı için
                for col, (co_id, _) in enumerate(course_outcomes, start=1):
                    relation = relations.get((po_id, co_id), 0)
                    success = success_rates.get(co_id, 0)
                    value = relation * success / 100  # Değeri hesapla
                    total_relation += relation
                    total_success += value
                    
                    ttk.Label(self.table5_grid_frame, text=f"{value * 100:.1f}").grid(
                        row=row, column=col, padx=2, pady=2)
                
                # Başarı oranını hesapla ve göster
                success_rate = (total_success / total_relation * 100) if total_relation > 0 else 0
                ttk.Label(self.table5_grid_frame, text=f"{success_rate:.1f}",
                         foreground="red").grid(
                             row=row, column=len(course_outcomes)+1, padx=5, pady=2)

        except Exception as e:
            messagebox.showerror("Hata", f"Hesaplama sırasında hata oluştu: {str(e)}")

    def export_table5(self):
        """Tablo 5'i Excel'e aktar"""
        if not self.selected_course.get():
            messagebox.showerror("Hata", "Lütfen önce bir ders seçin.")
            return

        try:
            # Öğrenci listesini al
            self.cursor.execute("""
                SELECT DISTINCT student_name 
                FROM student_grades 
                WHERE course_id = (SELECT id FROM courses WHERE course_name = ?)
                ORDER BY student_name
            """, (self.selected_course.get(),))
            students = [row[0] for row in self.cursor.fetchall()]

            if not students:
                messagebox.showwarning("Uyarı", "Bu derste kayıtlı öğrenci bulunamadı.")
                return

            # Program çıktılarını al
            self.cursor.execute("SELECT id, outcome FROM program_outcomes ORDER BY id")
            program_outcomes = self.cursor.fetchall()

            # Ders çıktılarını al
            self.cursor.execute("""
                SELECT id, outcome 
                FROM course_outcomes 
                WHERE course_id = (SELECT id FROM courses WHERE course_name = ?)
                ORDER BY id
            """, (self.selected_course.get(),))
            course_outcomes = self.cursor.fetchall()

            # Tablo 1'deki ilişki değerlerini al
            self.cursor.execute("""
                SELECT program_outcome_id, course_outcome_id, value
                FROM relationship_values
                WHERE course_id = (SELECT id FROM courses WHERE course_name = ?)
            """, (self.selected_course.get(),))
            relations = {(row[0], row[1]): row[2] for row in self.cursor.fetchall()}

            # Excel için verileri hazırla
            excel_data = []
            
            # Her öğrenci için
            for student_name in students:
                # Tablo 4'teki başarı değerlerini al
                self.cursor.execute("""
                    SELECT course_outcome_id, success_rate
                    FROM success_rates
                    WHERE course_id = (SELECT id FROM courses WHERE course_name = ?)
                    AND student_name = ?
                """, (self.selected_course.get(), student_name))
                success_rates = {row[0]: row[1] for row in self.cursor.fetchall()}

                if not success_rates:  # Eğer öğrencinin başarı oranları kaydedilmemişse
                    continue  # Bu öğrenciyi atla
                
                # Her program çıktısı için
                for po_id, po_text in program_outcomes:
                    row_data = [student_name, f"PÇ{po_id}"]
                    total_relation = 0
                    total_success = 0 
                    
                    # Her ders çıktısı için değerleri hesapla
                    for co_id, _ in course_outcomes:
                        relation = relations.get((po_id, co_id), 0)
                        success = success_rates.get(co_id, 0)
                        value = relation * success / 100  # Değeri hesapla
                        row_data.append(value * 100)  # Yüzde olarak göster
                        total_relation += relation
                        total_success += value
                    
                    # Başarı oranını hesapla
                    success_rate = (total_success / total_relation * 100) if total_relation > 0 else 0
                    row_data.append(success_rate)
                    
                    excel_data.append(row_data)

            if not excel_data:
                messagebox.showwarning("Uyarı", "Dışa aktarılacak veri bulunamadı.")
                return

            # Başlıkları hazırla
            headers = ["Öğrenci", "Program Çıktısı"] + [f"ÇK{co[0]}" for co in course_outcomes] + ["Başarı Oranı"]

            # DataFrame oluştur
            df = pd.DataFrame(excel_data, columns=headers)
            
            # Excel'e kaydet
            filename = f"{self.selected_course.get()}_Tablo5.xlsx"
            df.to_excel(filename, index=False)
            messagebox.showinfo("Başarılı", f"Tablo 5 verileri {filename} dosyasına aktarıldı.")

        except Exception as e:
            messagebox.showerror("Hata", f"Dışa aktarma sırasında hata oluştu: {str(e)}")

    def save_table4(self):
        """Tablo 4'teki başarı oranlarını kaydet"""
        if not self.selected_course.get():
            messagebox.showerror("Hata", "Lütfen ders seçin.")
            return

        try:
            # Ders ID'sini al
            self.cursor.execute("SELECT id FROM courses WHERE course_name = ?", 
                              (self.selected_course.get(),))
            course_id = self.cursor.fetchone()[0]

            # Dersin tüm öğrencilerini al
            self.cursor.execute("""
                SELECT student_name 
                FROM student_grades 
                WHERE course_id = ?
                ORDER BY student_name
            """, (course_id,))
            students = [row[0] for row in self.cursor.fetchall()]

            # Her öğrenci için başarı oranlarını hesapla ve kaydet
            for student_name in students:
                # Öğrenci için mevcut kayıtları sil
                self.cursor.execute("""
                    DELETE FROM success_rates 
                    WHERE course_id = ? AND student_name = ?
                """, (course_id, student_name))

                # Kriterleri al
                self.cursor.execute("""
                    SELECT id, criterion, weight 
                    FROM criteria 
                    WHERE course_id = ?
                    ORDER BY id
                """, (course_id,))
                criteria = self.cursor.fetchall()

                # Tablo 2 değerlerini al
                self.cursor.execute("""
                    SELECT course_outcome_id, criterion_id, value
                    FROM course_outcome_evaluations
                    WHERE course_id = ?
                """, (course_id,))
                evaluations = {(row[0], row[1]): row[2] for row in self.cursor.fetchall()}

                # Öğrenci notlarını al
                self.cursor.execute("""
                    SELECT grades
                    FROM student_grades
                    WHERE course_id = ? AND student_name = ?
                """, (course_id, student_name))
                
                grades = json.loads(self.cursor.fetchone()[0])

                # Her ders çıktısı için başarı oranını hesapla
                self.cursor.execute("""
                    SELECT id 
                    FROM course_outcomes 
                    WHERE course_id = ?
                    ORDER BY id
                """, (course_id,))
                
                for (outcome_id,) in self.cursor.fetchall():
                    total = 0
                    max_total = 0
                    
                    for criterion_id, _, weight in criteria:
                        t2_value = evaluations.get((outcome_id, criterion_id), 0)
                        student_grade = float(grades.get(str(criterion_id), 0))
                        
                        value = t2_value * (weight/100) * student_grade
                        max_value = t2_value * (weight/100) * 100
                        
                        total += value
                        max_total += max_value
                    
                    success_rate = (total * 100 / max_total) if max_total > 0 else 0
                    
                    # Başarı oranını kaydet
                    self.cursor.execute("""
                        INSERT INTO success_rates (course_id, student_name, course_outcome_id, success_rate)
                        VALUES (?, ?, ?, ?)
                    """, (course_id, student_name, outcome_id, success_rate))

            self.conn.commit()
            messagebox.showinfo("Başarılı", "Tüm öğrencilerin başarı oranları kaydedildi.")

        except Exception as e:
            self.conn.rollback()
            messagebox.showerror("Hata", f"Kayıt sırasında hata oluştu: {str(e)}")

    def calculate_table4_success_rates(self):
        """Tablo 4'teki başarı oranlarını hesapla"""
        success_rates = {}
        
        try:
            # Kriterleri al
            self.cursor.execute("""
                SELECT id, criterion, weight 
                FROM criteria 
                WHERE course_id = (SELECT id FROM courses WHERE course_name = ?)
                ORDER BY id
            """, (self.selected_course.get(),))
            criteria = self.cursor.fetchall()

            # Tablo 2 değerlerini al
            self.cursor.execute("""
                SELECT course_outcome_id, criterion_id, value
                FROM course_outcome_evaluations
                WHERE course_id = (SELECT id FROM courses WHERE course_name = ?)
            """, (self.selected_course.get(),))
            evaluations = {(row[0], row[1]): row[2] for row in self.cursor.fetchall()}

            # Öğrenci notlarını al
            self.cursor.execute("""
                SELECT grades
                FROM student_grades
                WHERE course_id = (SELECT id FROM courses WHERE course_name = ?)
                AND student_name = ?
            """, (self.selected_course.get(), self.selected_student.get()))
            
            grades = json.loads(self.cursor.fetchone()[0])

            # Her ders çıktısı için başarı oranını hesapla
            self.cursor.execute("""
                SELECT id 
                FROM course_outcomes 
                WHERE course_id = (SELECT id FROM courses WHERE course_name = ?)
                ORDER BY id
            """, (self.selected_course.get(),))
            
            for (outcome_id,) in self.cursor.fetchall():
                total = 0
                max_total = 0
                
                for criterion_id, _, weight in criteria:
                    t2_value = evaluations.get((outcome_id, criterion_id), 0)
                    student_grade = float(grades.get(str(criterion_id), 0))
                    
                    value = t2_value * (weight/100) * student_grade
                    max_value = t2_value * (weight/100) * 100
                    
                    total += value
                    max_total += max_value
                
                success_rate = (total * 100 / max_total) if max_total > 0 else 0
                success_rates[outcome_id] = success_rate

        except Exception as e:
            messagebox.showerror("Hata", f"Başarı oranları hesaplanırken hata oluştu: {str(e)}")

        return success_rates

if __name__ == "__main__":
    root = tk.Tk()
    app = SuccessCalculationApp(root)
    root.mainloop()