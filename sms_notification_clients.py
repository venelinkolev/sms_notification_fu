"""
Kasi Extractor v2.0 - GUI Приложение за извличане на данни от MDB и CSV
Поддържа както .mdb файлове (чрез mdbtools), така и директна работа с .csv файлове
"""

from tkinter import ttk, filedialog, messagebox
from datetime import datetime, date
import tkinter as tk
import json
import csv
import sys
import io
import os
import subprocess
import tempfile
import platform

try:
    import pandas as pd
    PANDAS_AVAILABLE = True
except ImportError:
    PANDAS_AVAILABLE = False

# Проверка дали сме на Windows и имаме mdbtools
IS_WINDOWS = platform.system().lower() == 'windows'
MDBTOOLS_AVAILABLE = False

# Проверяваме дали mdbtools са налични в системата
try:
    if IS_WINDOWS:
        # Проверка за mdbtools на Windows
        result = subprocess.run(['mdb-ver'], capture_output=True, text=True, timeout=5)
        MDBTOOLS_AVAILABLE = result.returncode == 0
    else:
        # На Linux проверяваме за mdb-tools
        result = subprocess.run(['mdb-ver'], capture_output=True, text=True, timeout=5)
        MDBTOOLS_AVAILABLE = result.returncode == 0
except (subprocess.TimeoutExpired, FileNotFoundError, subprocess.SubprocessError):
    MDBTOOLS_AVAILABLE = False

class KasiExtractor:
    def __init__(self, root):
        self.root = root
        self.root.title("SMS Notification Clients v2.0 - CSV Support")
        self.root.geometry("950x830")
        self.root.resizable(True, True)

        self.filtered_data_lines = []
        self.current_file_type = None
        self.file_path = tk.StringVar()
        self.start_date = tk.StringVar()
        self.end_date = tk.StringVar()
        
        # Задаване на начални дати
        today = date.today()
        month_ago = date(today.year, today.month-1 if today.month > 1 else 12, today.day)
        self.start_date.set(month_ago.strftime('%d.%m.%Y'))
        self.end_date.set(today.strftime('%d.%m.%Y'))
        
        self.create_widgets()
        self.set_default_dates()

    def validate_date_input(self, date_string):
        """Валидира дата в формат dd.mm.yyyy"""
        if not date_string.strip():
            return "empty"
        
        try:
            if len(date_string) != 10:
                return "invalid"
            
            if date_string.count('.') != 2:
                return "invalid"
            
            datetime.strptime(date_string, '%d.%m.%Y')
            return "valid"
        except ValueError:
            return "invalid"

    def validate_date_range(self):
        """Проверява дали крайната дата е след началната"""
        start_text = self.start_date_entry.get().strip()
        end_text = self.end_date_entry.get().strip()
        
        if not start_text or not end_text:
            return True
        
        if (self.validate_date_input(start_text) != "valid" or 
            self.validate_date_input(end_text) != "valid"):
            return True
        
        try:
            start_date = datetime.strptime(start_text, '%d.%m.%Y')
            end_date = datetime.strptime(end_text, '%d.%m.%Y')
            return start_date <= end_date
        except:
            return True

    def on_date_entry_change(self, event, entry_widget):
        """Проверява датата при промяна в Entry полето"""
        date_text = entry_widget.get()
        validation_result = self.validate_date_input(date_text)
        
        if validation_result == "valid":
            entry_widget.config(bg="lightgreen")
        elif validation_result == "empty":
            entry_widget.config(bg="white")
        else:
            entry_widget.config(bg="lightcoral")
            return
        
        if validation_result == "valid":
            if not self.validate_date_range():
                self.start_date_entry.config(bg="orange")
                self.end_date_entry.config(bg="orange")
                self.update_status_bar("ГРЕШКА: Крайната дата е преди началната!")
            else:
                if self.validate_date_input(self.start_date_entry.get()) == "valid":
                    self.start_date_entry.config(bg="lightgreen")
                if self.validate_date_input(self.end_date_entry.get()) == "valid":
                    self.end_date_entry.config(bg="lightgreen")
                self.update_status_bar("Готов за работа")
    
    def create_widgets(self):
        """Създава всички UI елементи"""
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        
        # 1. СЕКЦИЯ: ИЗБОР НА ФАЙЛ
        file_frame = ttk.LabelFrame(main_frame, text="📁 Избор на MDB или CSV файл", padding="10")
        file_frame.grid(row=0, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        file_frame.columnconfigure(1, weight=1)
        
        ttk.Button(file_frame, text="Избери файл", 
                  command=self.select_file).grid(row=0, column=0, padx=(0, 10))
        
        self.file_entry = ttk.Entry(file_frame, textvariable=self.file_path, 
                                   state="readonly", width=50)
        self.file_entry.grid(row=0, column=1, sticky=(tk.W, tk.E))
        
        info_label = ttk.Label(file_frame, 
                              text="Поддържани файлове: .mdb (Access Database), .csv (Comma Separated Values)", 
                              foreground="gray", font=("TkDefaultFont", 8))
        info_label.grid(row=1, column=0, columnspan=2, pady=(5, 0), sticky=tk.W)

        # 2. СЕКЦИЯ: ИНФОРМАЦИЯ ЗА MDBTOOLS
        self.mdb_info_frame = ttk.LabelFrame(main_frame, text="ℹ️ Информация за MDB поддръжка", padding="10")
        self.mdb_info_frame.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        self.mdb_info_frame.columnconfigure(0, weight=1)
        
        mdb_info_text = ""
        if IS_WINDOWS:
            if MDBTOOLS_AVAILABLE:
                mdb_info_text = "✅ mdbtools са инсталирани и налични в системата"
            else:
                mdb_info_text = "⚠️ За MDB файлове е необходимо да инсталирате mdbtools\n" \
                               "1. Изтеглете от: https://github.com/mdbtools/mdbtools/releases\n" \
                               "2. Добавете bin директорията в системния PATH\n" \
                               "3. Рестартирайте приложението"
        else:
            mdb_info_text = "✅ На Linux система с mdb-tools"
        
        mdb_info_label = ttk.Label(self.mdb_info_frame, text=mdb_info_text, 
                                  foreground="green" if MDBTOOLS_AVAILABLE else "orange",
                                  font=("TkDefaultFont", 9), wraplength=700)
        mdb_info_label.grid(row=0, column=0, sticky=tk.W)

        # 3. СЕКЦИЯ: СТАТУС НА ФАЙЛА
        status_frame = ttk.LabelFrame(main_frame, text="📊 Информация за файла", padding="10")
        status_frame.grid(row=2, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        status_frame.columnconfigure(0, weight=1)
        
        self.status_label = ttk.Label(status_frame, text="Няма избран файл", 
                                     foreground="gray")
        self.status_label.grid(row=0, column=0, sticky=tk.W)
        
        # 4. СЕКЦИЯ: ТЕСТ НА ВРЪЗКАТА/ФАЙЛА
        test_frame = ttk.LabelFrame(main_frame, text="🔧 Преглед на данните", padding="10")
        test_frame.grid(row=3, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        
        self.test_button = ttk.Button(test_frame, text="📋 Прегледай файла", 
                                     command=self.test_file_connection, 
                                     state="disabled")
        self.test_button.grid(row=0, column=0, padx=(0, 10))

        # 5. СЕКЦИЯ: ИЗБОР НА ДАТИ
        date_frame = ttk.LabelFrame(main_frame, text="📅 Филтриране по дати", padding="10")
        date_frame.grid(row=4, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        date_frame.columnconfigure(1, weight=1)
        date_frame.columnconfigure(3, weight=1)
        
        ttk.Label(date_frame, text="От дата:").grid(row=0, column=0, padx=(0, 5), sticky=tk.W)
        self.start_date_entry = tk.Entry(date_frame, width=12)
        self.start_date_entry.grid(row=0, column=1, padx=(0, 20), sticky=tk.W)
        self.start_date_entry.bind('<KeyRelease>', lambda e: self.on_date_entry_change(e, self.start_date_entry))

        ttk.Label(date_frame, text="До дата:").grid(row=0, column=2, padx=(0, 5), sticky=tk.W)
        self.end_date_entry = tk.Entry(date_frame, width=12)
        self.end_date_entry.grid(row=0, column=3, padx=(0, 20), sticky=tk.W)
        self.end_date_entry.bind('<KeyRelease>', lambda e: self.on_date_entry_change(e, self.end_date_entry))
        
        self.filter_button = ttk.Button(date_frame, text="📊 Филтрирай данните", 
                                       command=self.filter_data, state="disabled")
        self.filter_button.grid(row=0, column=4, padx=(20, 0))
        
        instruction_label = ttk.Label(date_frame, text="Формат: dd.mm.yyyy (например: 10.09.2025)", 
                                     foreground="gray", font=("TkDefaultFont", 8))
        instruction_label.grid(row=1, column=0, columnspan=4, pady=(5, 0), sticky=tk.W)
        
        self.filter_result_label = ttk.Label(date_frame, text="", foreground="gray")
        self.filter_result_label.grid(row=2, column=0, columnspan=5, pady=(10, 0), sticky=tk.W)

        # 6. СЕКЦИЯ: ИЗВЛИЧАНЕ НА КОНКРЕТНИ КОЛОНИ
        extract_frame = ttk.LabelFrame(main_frame, text="📋 Извличане на данни", padding="10")
        extract_frame.grid(row=5, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        extract_frame.columnconfigure(0, weight=1)
        
        info_label = ttk.Label(extract_frame, 
                              text="Колони за извличане: Number, End_Data, Model, Number_EKA, Ime_Obekt, Adres_Obekt, Dan_Number, Phone, Ime_Firma, bulst",
                              foreground="gray", font=("TkDefaultFont", 8), wraplength=500)
        info_label.grid(row=0, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        
        self.extract_button = ttk.Button(extract_frame, text="📊 Извлечи колони", 
                                        command=self.extract_specific_columns, state="disabled")
        self.extract_button.grid(row=1, column=0, padx=(0, 10))
        
        self.save_csv_button = ttk.Button(extract_frame, text="💾 Запиши CSV", 
                                         command=self.save_csv, state="disabled")
        self.save_csv_button.grid(row=1, column=1, padx=(0, 10))
        
        self.save_json_button = ttk.Button(extract_frame, text="💾 Запиши JSON", 
                                          command=self.save_json, state="disabled")
        self.save_json_button.grid(row=1, column=2)

        self.extract_result_label = ttk.Label(extract_frame, text="", foreground="gray")
        self.extract_result_label.grid(row=2, column=0, columnspan=3, pady=(10, 0), sticky=tk.W)
        
        # 7. СЕКЦИЯ: ПЪЛЕН ЕКСПОРТ
        export_frame = ttk.LabelFrame(main_frame, text="📤 Пълен експорт", padding="10")
        export_frame.grid(row=6, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        export_frame.columnconfigure(0, weight=1)
        
        export_info_label = ttk.Label(export_frame, 
                                     text="Експортиране на цялата таблица (всички колони, всички редове)",
                                     foreground="gray", font=("TkDefaultFont", 8))
        export_info_label.grid(row=0, column=0, sticky=tk.W, pady=(0, 10))
        
        self.full_export_button = ttk.Button(export_frame, text="📁 Експортирай цял файл", 
                                            command=self.export_full_table, state="disabled")
        self.full_export_button.grid(row=1, column=0, sticky=tk.W)
        
        # 8. СТАТУС БАР
        status_bar_frame = ttk.Frame(main_frame)
        status_bar_frame.grid(row=10, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(20, 0))
        status_bar_frame.columnconfigure(0, weight=1)

        self.status_bar = ttk.Label(status_bar_frame, text="Готов за работа", 
                                   relief=tk.SUNKEN, anchor=tk.W, padding="5")
        self.status_bar.grid(row=0, column=0, sticky=(tk.W, tk.E))
        
        ttk.Button(status_bar_frame, text="Изход", 
                  command=self.exit_application).grid(row=0, column=1, padx=(10, 0))

    def set_default_dates(self):
        """Задава днешна дата като период по подразбиране"""
        try:
            today = date.today()
            today_str = today.strftime('%d.%m.%Y')
            
            self.start_date_entry.delete(0, tk.END)
            self.start_date_entry.insert(0, today_str)
            
            self.end_date_entry.delete(0, tk.END)
            self.end_date_entry.insert(0, today_str)
            
            self.update_status_bar(f"Зададен е период: {today_str} - {today_str}")
            
        except Exception as e:
            print(f"Предупреждение: Не можах да задам началните дати: {e}")
            self.update_status_bar("Готов за работа")

    def select_file(self):
        """Отваря диалог за избор на MDB или CSV файл"""
        file_path = filedialog.askopenfilename(
            title="Избери MDB или CSV файл",
            filetypes=[
                ("MDB файлове", "*.mdb"),
                ("CSV файлове", "*.csv"),
                ("Всички файлове", "*.*")
            ]
        )
        
        if file_path:
            self.file_path.set(file_path)
            self.detect_file_type(file_path)
            self.update_file_status(file_path)
            self.update_status_bar(f"Избран файл: {os.path.basename(file_path)}")

    def detect_file_type(self, file_path):
        """Разпознава типа на файла и адаптира интерфейса"""
        file_extension = os.path.splitext(file_path)[1].lower()
        
        if file_extension == '.mdb':
            self.current_file_type = 'mdb'
            self.test_button.config(text="🔧 Тествай MDB файла")
            
            # Проверка дали mdbtools са налични за MDB
            if not MDBTOOLS_AVAILABLE:
                self.filter_button.config(state="disabled")
                self.full_export_button.config(state="disabled")
                self.update_status_bar("⚠️ За MDB файлове са необходими mdbtools")
            else:
                self.filter_button.config(state="normal")
                self.full_export_button.config(state="normal")
                
        elif file_extension == '.csv':
            self.current_file_type = 'csv'
            self.test_button.config(text="📋 Прегледай CSV файла")
            self.filter_button.config(state="normal")
            self.full_export_button.config(state="normal")
        else:
            self.current_file_type = 'unknown'
            self.test_button.config(text="❓ Прегледай файла")
            self.filter_button.config(state="disabled")
            self.full_export_button.config(state="disabled")
    
    def update_file_status(self, file_path):
        """Обновява статуса на избрания файл"""
        if os.path.exists(file_path):
            file_size = os.path.getsize(file_path)
            size_mb = file_size / (1024 * 1024)
            file_type = self.current_file_type.upper() if self.current_file_type else "НЕИЗВЕСТЕН"
            
            status_text = f"✅ Файл: {os.path.basename(file_path)} ({file_type}, {size_mb:.1f} MB)"
            self.status_label.config(text=status_text, foreground="green")
            
            self.test_button.config(state="normal")
            
        else:
            self.status_label.config(text="❌ Файлът не съществува", foreground="red")
            self.test_button.config(state="disabled")

    def test_file_connection(self):
        """Тества файла и показва информация за него"""
        if not self.file_path.get():
            messagebox.showerror("Грешка", "Моля изберете файл първо!")
            return
        
        self.update_status_bar("Прегледане на файла...")
        
        if self.current_file_type == 'csv':
            self._test_csv_file()
        elif self.current_file_type == 'mdb':
            self._test_mdb_file()
        else:
            messagebox.showerror("Грешка", "Неподдържан файлов формат!")

    def _test_csv_file(self):
        """Тества CSV файл"""
        try:
            if not PANDAS_AVAILABLE:
                messagebox.showerror("Грешка", "pandas не е инсталиран! Необходим е за работа с CSV файлове.")
                return
            
            df = pd.read_csv(self.file_path.get(), nrows=5, encoding='utf-8')
            total_rows = sum(1 for line in open(self.file_path.get(), 'r', encoding='utf-8')) - 1
            total_columns = len(df.columns)
            
            has_end_data = 'End_Data' in df.columns
            
            required_columns = ['Number', 'End_Data', 'Model', 'Number_EKA', 'Ime_Obekt', 
                              'Adres_Obekt', 'Dan_Number', 'Phone', 'Ime_Firma', 'bulst']
            found_columns = [col for col in required_columns if col in df.columns]
            
            messagebox.showinfo("Информация за CSV файла", 
                              f"✅ CSV файлът е четлив!\n\n"
                              f"📊 Общо редове: {total_rows:,}\n"
                              f"📋 Общо колони: {total_columns}\n"
                              f"📅 Колона 'End_Data': {'✅ Намерена' if has_end_data else '❌ Не е намерена'}\n"
                              f"🎯 Намерени нужни колони: {len(found_columns)}/{len(required_columns)}\n\n"
                              f"Първите колони:\n" + ", ".join(df.columns[:10]))
            
            self.update_status_bar("✅ CSV файлът е готов за работа")
            
        except Exception as e:
            messagebox.showerror("Грешка", f"Грешка при четене на CSV файла:\n{str(e)}")
            self.update_status_bar(f"Грешка: {str(e)}")

    def _test_mdb_file(self):
        """Тества MDB файл с mdbtools"""
        if not MDBTOOLS_AVAILABLE:
            messagebox.showerror("Грешка", 
                               "mdbtools не са налични!\n\n"
                               "Моля инсталирайте mdbtools:\n"
                               "1. Изтеглете от: https://github.com/mdbtools/mdbtools/releases\n"
                               "2. Добавете bin директорията в PATH\n"
                               "3. Рестартирайте приложението")
            return
        
        try:
            # Използваме mdb-tables за получаване на списък с таблици
            result = subprocess.run(['mdb-tables', self.file_path.get()], 
                                  capture_output=True, text=True, timeout=30)
            
            if result.returncode != 0:
                messagebox.showerror("Грешка", 
                                   f"Грешка при четене на MDB файла:\n{result.stderr}")
                return
            
            tables = result.stdout.strip().split()
            
            if "Kasi_all" in tables:
                messagebox.showinfo("Успех", 
                                f"✅ Връзката е успешна!\n\n"
                                f"Намерени таблици: {len(tables)}\n"
                                f"Таблица 'Kasi_all': ✅ Намерена\n\n"
                                f"Други таблици:\n" + "\n".join(tables))
                self.update_status_bar("✅ MDB файлът е готов за работа")
            else:
                messagebox.showwarning("Внимание", 
                                    f"Таблица 'Kasi_all' не е намерена!\n\n"
                                    f"Налични таблици:\n" + "\n".join(tables))
                self.update_status_bar("⚠️ Таблица 'Kasi_all' не е намерена")
            
        except subprocess.TimeoutExpired:
            messagebox.showerror("Грешка", "Таймаут при четене на MDB файла!")
            self.update_status_bar("Таймаут при тестване на MDB")
        except Exception as e:
            messagebox.showerror("Грешка", f"Неочаквана грешка:\n{str(e)}")
            self.update_status_bar(f"Грешка: {str(e)}")

    def filter_data(self):
        """Филтрира данните по избраните дати"""
        if not self.file_path.get():
            messagebox.showerror("Грешка", "Моля изберете файл първо!")
            return
        
        if self.current_file_type == 'csv':
            return self._filter_csv_data()
        elif self.current_file_type == 'mdb':
            return self._filter_mdb_data()
        else:
            messagebox.showerror("Грешка", "Неподдържан файлов формат!")
            return False

    def _filter_csv_data(self):
        """Филтрира CSV данни"""
        if not PANDAS_AVAILABLE:
            messagebox.showerror("Грешка", "pandas не е инсталиран!")
            return False
        
        try:
            start_date_str = self.start_date_entry.get().strip()
            end_date_str = self.end_date_entry.get().strip()
            
            if not start_date_str or not end_date_str:
                messagebox.showerror("Грешка", "Моля въведете начална и крайна дата!")
                return False
            
            if not self.validate_date_range():
                messagebox.showerror("Грешка", "Крайната дата не може да бъде преди началната дата!")
                return False
                
        except Exception as e:
            messagebox.showerror("Грешка", f"Проблем с четенето на датите:\n{str(e)}")
            return False

        self.update_status_bar(f"Филтриране от {start_date_str} до {end_date_str}...")
        
        try:
            df = pd.read_csv(self.file_path.get(), encoding='utf-8')
            start_date = datetime.strptime(start_date_str, '%d.%m.%Y')
            end_date = datetime.strptime(end_date_str, '%d.%m.%Y')
            
            if 'End_Data' not in df.columns:
                messagebox.showerror("Грешка", "Колона 'End_Data' не е намерена в CSV файла!")
                return False
            
            try:
                df['End_Data_parsed'] = pd.to_datetime(df['End_Data'], format='%m/%d/%y %H:%M:%S', errors='coerce')
            except:
                try:
                    df['End_Data_parsed'] = pd.to_datetime(df['End_Data'], format='%m/%d/%Y %H:%M:%S', errors='coerce')
                except:
                    df['End_Data_parsed'] = pd.to_datetime(df['End_Data'], errors='coerce')
            
            mask = (df['End_Data_parsed'].dt.date >= start_date.date()) & \
                (df['End_Data_parsed'].dt.date <= end_date.date())
            filtered_df = df[mask]
            
            self._save_filtered_data_as_lines(filtered_df)
            
            total_rows = len(filtered_df)
            original_rows = len(df)
            percent = (total_rows/original_rows*100) if original_rows > 0 else 0
            
            result_text = f"✅ Филтрирани {total_rows} от общо {original_rows} реда"
            self.filter_result_label.config(text=result_text, foreground="green")
            self.update_status_bar(f"Филтриране завършено: {total_rows} от {original_rows} реда ({percent:.1f}%)")
            
            messagebox.showinfo("Резултат", f"Филтрирането е завършено!\n\nПериод: {start_date_str} - {end_date_str}\nОбщо редове: {original_rows}\nФилтрирани редове: {total_rows}")
            
            self.extract_button.config(state="normal")
            return True
            
        except Exception as e:
            messagebox.showerror("Грешка", f"Неочаквана грешка:\n{str(e)}")
            self.update_status_bar(f"Грешка: {str(e)}")
            return False

    def _filter_mdb_data(self):
        """Филтрира MDB данни с mdbtools"""
        if not MDBTOOLS_AVAILABLE:
            messagebox.showerror("Грешка", "mdbtools не са налични!")
            return False
        
        try:
            start_date_str = self.start_date_entry.get().strip()
            end_date_str = self.end_date_entry.get().strip()
            
            if not start_date_str or not end_date_str:
                messagebox.showerror("Грешка", "Моля въведете начална и крайна дата!")
                return False
            
            if not self.validate_date_range():
                messagebox.showerror("Грешка", "Крайната дата не може да бъде преди началната дата!")
                return False
                
        except Exception as e:
            messagebox.showerror("Грешка", f"Проблем с четенето на датите:\n{str(e)}")
            return False

        self.update_status_bar(f"Филтриране от {start_date_str} до {end_date_str}...")
        
        try:
            # Експортираме цялата таблица временно
            with tempfile.NamedTemporaryFile(suffix='.csv', delete=False, mode='w+', encoding='utf-8') as temp_file:
                temp_csv_path = temp_file.name
            
            # Експортираме таблицата с mdb-export
            cmd = ['mdb-export', self.file_path.get(), 'Kasi_all']
            
            with open(temp_csv_path, 'w', encoding='utf-8') as output_file:
                result = subprocess.run(cmd, stdout=output_file, stderr=subprocess.PIPE, text=True, timeout=120)
            
            if result.returncode != 0:
                messagebox.showerror("Грешка", f"Грешка при експорт на MDB: {result.stderr}")
                os.unlink(temp_csv_path)
                return False
            
            # Четем CSV с pandas
            df = pd.read_csv(temp_csv_path, encoding='utf-8')
            
            # Почистваме временния файл
            os.unlink(temp_csv_path)
            
            start_date = datetime.strptime(start_date_str, '%d.%m.%Y')
            end_date = datetime.strptime(end_date_str, '%d.%m.%Y')
            
            if 'End_Data' not in df.columns:
                messagebox.showerror("Грешка", "Колона 'End_Data' не е намерена в таблицата!")
                return False
            
            try:
                df['End_Data_parsed'] = pd.to_datetime(df['End_Data'], format='%m/%d/%y %H:%M:%S', errors='coerce')
            except:
                try:
                    df['End_Data_parsed'] = pd.to_datetime(df['End_Data'], format='%m/%d/%Y %H:%M:%S', errors='coerce')
                except:
                    df['End_Data_parsed'] = pd.to_datetime(df['End_Data'], errors='coerce')
            
            mask = (df['End_Data_parsed'].dt.date >= start_date.date()) & \
                (df['End_Data_parsed'].dt.date <= end_date.date())
            filtered_df = df[mask]
            
            self._save_filtered_data_as_lines(filtered_df)
            
            total_rows = len(filtered_df)
            original_rows = len(df)
            percent = (total_rows/original_rows*100) if original_rows > 0 else 0
            
            result_text = f"✅ Филтрирани {total_rows} от общо {original_rows} реда"
            self.filter_result_label.config(text=result_text, foreground="green")
            self.update_status_bar(f"Филтриране завършено: {total_rows} от {original_rows} реда ({percent:.1f}%)")
            
            messagebox.showinfo("Резултат", f"Филтрирането е завършено!\n\nПериод: {start_date_str} - {end_date_str}\nОбщо редове: {original_rows}\nФилтрирани редове: {total_rows}")
            
            self.extract_button.config(state="normal")
            return True
            
        except subprocess.TimeoutExpired:
            messagebox.showerror("Грешка", "Таймаут при филтриране на MDB файла!")
            self.update_status_bar("Таймаут при филтриране")
            return False
        except Exception as e:
            messagebox.showerror("Грешка", f"Неочаквана грешка при филтриране:\n{str(e)}")
            self.update_status_bar(f"Грешка: {str(e)}")
            return False

    def _save_filtered_data_as_lines(self, filtered_df):
        """Запазва филтрираните данни като CSV lines"""
        self.filtered_data_lines = []
        
        columns = list(filtered_df.columns)
        if 'End_Data_parsed' in columns:
            columns.remove('End_Data_parsed')
        self.filtered_data_lines.append(','.join(f'"{col}"' for col in columns))
        
        for _, row in filtered_df.iterrows():
            csv_row = []
            for col in columns:
                value = row[col]
                if pd.isna(value):
                    csv_row.append('""')
                else:
                    str_value = str(value).replace('"', '""')
                    csv_row.append(f'"{str_value}"')
            self.filtered_data_lines.append(','.join(csv_row))

    def extract_specific_columns(self):
        """Извлича конкретните 10 колони от филтрираните данни"""
        if not hasattr(self, 'filtered_data_lines') or len(self.filtered_data_lines) < 2:
            messagebox.showerror("Грешка", "Няма филтрирани данни! Първо направете филтрация.")
            return False
        
        self.update_status_bar("Извличане на конкретни колони...")
        
        required_columns = ['Number', 'End_Data', 'Model', 'Number_EKA', 'Ime_Obekt', 
                        'Adres_Obekt', 'Dan_Number', 'Phone', 'Ime_Firma', 'bulst']
        
        try:
            header_line = self.filtered_data_lines[0]
            header_reader = csv.reader(io.StringIO(header_line))
            headers = next(header_reader)
            
            column_indices = {}
            missing_columns = []
            
            for col_name in required_columns:
                found_index = None
                for i, header in enumerate(headers):
                    if col_name.lower() in header.lower():
                        found_index = i
                        break
                
                if found_index is not None:
                    column_indices[col_name] = found_index
                else:
                    missing_columns.append(col_name)
            
            if missing_columns:
                messagebox.showwarning("Внимание", 
                                    f"Следните колони не са намерени:\n{', '.join(missing_columns)}\n\n"
                                    f"Ще бъдат извлечени само намерените колони.")
            
            new_header = [col for col in required_columns if col in column_indices]
            extracted_data = []
            extracted_data.append(','.join(f'"{col}"' for col in new_header))
            
            for line in self.filtered_data_lines[1:]:
                try:
                    reader = csv.reader(io.StringIO(line))
                    fields = next(reader)
                    
                    new_row = []
                    for col_name in new_header:
                        if column_indices[col_name] < len(fields):
                            field_value = fields[column_indices[col_name]]
                            
                            if field_value.endswith('.0') and field_value.replace('.0', '').replace('-', '').isdigit():
                                field_value = field_value[:-2]
                            
                            new_row.append(f'"{field_value}"')
                        else:
                            new_row.append('""')
                    
                    extracted_data.append(','.join(new_row))
                
                except Exception as e:
                    continue
            
            self.extracted_data_lines = extracted_data
            total_extracted = len(extracted_data) - 1
            
            result_text = f"✅ Извлечени {len(new_header)} колони от {total_extracted} реда"
            if hasattr(self, 'filtered_data_lines'):
                original_rows = len(self.filtered_data_lines) - 1
                result_text += f" (от {original_rows} филтрирани)"
            
            self.extract_result_label.config(text=result_text, foreground="green")
            self.update_status_bar(f"Извличане завършено: {total_extracted} реда с {len(new_header)} колони")
            
            self.save_csv_button.config(state="normal")
            self.save_json_button.config(state="normal")
            
            messagebox.showinfo("Успех", 
                            f"Извличането е успешно!\n\n"
                            f"Колони: {len(new_header)}\n"
                            f"Редове: {total_extracted}\n\n"
                            f"Намерени колони:\n{', '.join(new_header)}")
            
            return True
            
        except Exception as e:
            messagebox.showerror("Грешка", f"Неочаквана грешка при извличане:\n{str(e)}")
            self.update_status_bar(f"Грешка: {str(e)}")
            return False

    def export_full_table(self):
        """Експортира целия файл в CSV формат"""
        if not self.file_path.get():
            messagebox.showerror("Грешка", "Моля изберете файл първо!")
            return
        
        if self.current_file_type == 'csv':
            self._export_full_csv()
        elif self.current_file_type == 'mdb':
            self._export_full_mdb()
        else:
            messagebox.showerror("Грешка", "Неподдържан файлов формат!")

    def _export_full_csv(self):
        """Експортира целия CSV файл"""
        file_path = filedialog.asksaveasfilename(
            title="Експортирай цял CSV файл",
            defaultextension=".csv",
            filetypes=[("CSV файлове", "*.csv"), ("Всички файлове", "*.*")],
            initialfile=os.path.splitext(os.path.basename(self.file_path.get()))[0] + "_export.csv"
        )
        
        if not file_path:
            return
        
        try:
            self.update_status_bar("Експортиране на целия CSV файл...")
            
            if not PANDAS_AVAILABLE:
                import shutil
                shutil.copy2(self.file_path.get(), file_path)
            else:
                df = pd.read_csv(self.file_path.get(), encoding='utf-8')
                
                for column in df.columns:
                    if df[column].dtype == 'object':
                        df[column] = df[column].astype(str).apply(
                            lambda x: self.fix_encoding_utf8_to_windows1251(x) if x != 'nan' else ''
                        )
                
                df.to_csv(file_path, index=False, encoding='utf-8')
            
            file_size = os.path.getsize(file_path)
            
            if PANDAS_AVAILABLE:
                total_rows = len(df)
                total_columns = len(df.columns)
                stats_text = f"📊 Редове: {total_rows:,}\n📋 Колони: {total_columns}\n"
            else:
                stats_text = ""
            
            self.update_status_bar(f"Пълен експорт завършен: {os.path.basename(file_path)}")
            
            messagebox.showinfo("Успех", 
                            f"Пълният експорт е завършен успешно!\n\n"
                            f"📁 Файл: {os.path.basename(file_path)}\n"
                            f"{stats_text}"
                            f"💾 Размер: {file_size / 1024 / 1024:.1f} MB\n"
                            f"🔗 Път: {file_path}")
            
        except Exception as e:
            messagebox.showerror("Грешка", f"Грешка при пълен експорт:\n{str(e)}")
            self.update_status_bar(f"Грешка: {str(e)}")

    def _export_full_mdb(self):
        """Експортира цялата MDB таблица с mdbtools"""
        if not MDBTOOLS_AVAILABLE:
            messagebox.showerror("Грешка", "mdbtools не са налични!")
            return
        
        file_path = filedialog.asksaveasfilename(
            title="Експортирай цяла таблица като CSV",
            defaultextension=".csv",
            filetypes=[("CSV файлове", "*.csv"), ("Всички файлове", "*.*")],
            initialfile="Kasi_all_full_export.csv"
        )
        
        if not file_path:
            return
        
        try:
            self.update_status_bar("Експортиране на цялата таблица...")
            
            # Използваме mdb-export за директен експорт
            cmd = ['mdb-export', self.file_path.get(), 'Kasi_all']
            
            with open(file_path, 'w', encoding='utf-8') as output_file:
                result = subprocess.run(cmd, stdout=output_file, stderr=subprocess.PIPE, text=True, timeout=300)
            
            if result.returncode != 0:
                messagebox.showerror("Грешка", f"Грешка при експорт на MDB: {result.stderr}")
                return
            
            # Ако имаме pandas, поправяме кодировката
            if PANDAS_AVAILABLE:
                df = pd.read_csv(file_path, encoding='utf-8')
                
                for column in df.columns:
                    if df[column].dtype == 'object':
                        df[column] = df[column].astype(str).apply(
                            lambda x: self.fix_encoding_utf8_to_windows1251(x) if x != 'nan' else ''
                        )
                
                df.to_csv(file_path, index=False, encoding='utf-8')
                total_rows = len(df)
                total_columns = len(df.columns)
            else:
                # Броим редове без header
                with open(file_path, 'r', encoding='utf-8') as f:
                    total_rows = sum(1 for _ in f) - 1
                total_columns = "unknown"
            
            file_size = os.path.getsize(file_path)
            
            self.update_status_bar(f"Пълен експорт завършен: {os.path.basename(file_path)}")
            
            messagebox.showinfo("Успех", 
                            f"Пълният експорт е завършен успешно!\n\n"
                            f"📁 Файл: {os.path.basename(file_path)}\n"
                            f"📊 Редове: {total_rows:,}\n"
                            f"📋 Колони: {total_columns}\n"
                            f"💾 Размер: {file_size / 1024 / 1024:.1f} MB\n"
                            f"🔗 Път: {file_path}")
            
        except subprocess.TimeoutExpired:
            messagebox.showerror("Грешка", "Таймаут при експорт на MDB файла!")
            self.update_status_bar("Таймаут при експорт")
        except Exception as e:
            messagebox.showerror("Грешка", f"Грешка при пълен експорт:\n{str(e)}")
            self.update_status_bar(f"Грешка: {str(e)}")

    def update_status_bar(self, message):
        """Обновява статус бара"""
        self.status_bar.config(text=message)
        self.root.update_idletasks()
    
    def exit_application(self):
        """Затваря приложението"""
        self.root.quit()

    def save_csv(self):
        """Запис в CSV формат"""
        if not hasattr(self, 'extracted_data_lines') or len(self.extracted_data_lines) < 2:
            messagebox.showerror("Грешка", "Няма извлечени данни за запис!")
            return
        
        file_path = filedialog.asksaveasfilename(
            title="Запиши като CSV",
            defaultextension=".csv",
            filetypes=[("CSV файлове", "*.csv"), ("Всички файлове", "*.*")]
        )
        
        if not file_path:
            return
        
        try:
            self.update_status_bar("Записване на CSV файл...")
            
            with open(file_path, 'w', encoding='utf-8', newline='') as f:
                for line in self.extracted_data_lines:
                    f.write(line + '\n')
            
            total_rows = len(self.extracted_data_lines) - 1
            file_size = os.path.getsize(file_path)
            
            self.update_status_bar(f"CSV файл записан успешно: {os.path.basename(file_path)}")
            
            messagebox.showinfo("Успех", 
                               f"CSV файлът е записан успешно!\n\n"
                               f"📁 Файл: {os.path.basename(file_path)}\n"
                               f"📊 Редове: {total_rows}\n"
                               f"💾 Размер: {file_size / 1024:.1f} KB\n"
                               f"🔗 Път: {file_path}")
            
        except Exception as e:
            messagebox.showerror("Грешка", f"Грешка при записване на CSV:\n{str(e)}")
            self.update_status_bar("Грешка при записване на CSV")
    
    def save_json(self):
        """Запис в JSON формат"""
        if not hasattr(self, 'extracted_data_lines') or len(self.extracted_data_lines) < 2:
            messagebox.showerror("Грешка", "Няма извлечени данни за запис!")
            return
        
        file_path = filedialog.asksaveasfilename(
            title="Запиши като JSON",
            defaultextension=".json",
            filetypes=[("JSON файлове", "*.json"), ("Всички файлове", "*.*")]
        )
        
        if not file_path:
            return
        
        try:
            self.update_status_bar("Записване на JSON файл...")
            
            header_line = self.extracted_data_lines[0]
            header_reader = csv.reader(io.StringIO(header_line))
            headers = next(header_reader)
            
            json_data = []
            
            for line in self.extracted_data_lines[1:]:
                try:
                    reader = csv.reader(io.StringIO(line))
                    fields = next(reader)
                    
                    row_object = {}
                    for i, header in enumerate(headers):
                        if i < len(fields):
                            row_object[header] = fields[i]
                        else:
                            row_object[header] = ""
                    
                    json_data.append(row_object)
                
                except Exception as e:
                    continue
            
            import json
            with open(file_path, 'w', encoding='utf-8') as f:
                json.dump(json_data, f, ensure_ascii=False, indent=2)
            
            total_objects = len(json_data)
            file_size = os.path.getsize(file_path)
            
            self.update_status_bar(f"JSON файл записан успешно: {os.path.basename(file_path)}")
            
            messagebox.showinfo("Успех", 
                               f"JSON файлът е записан успешно!\n\n"
                               f"📁 Файл: {os.path.basename(file_path)}\n"
                               f"📊 Обекти: {total_objects}\n"
                               f"💾 Размер: {file_size / 1024:.1f} KB\n"
                               f"🔗 Път: {file_path}")
            
        except Exception as e:
            messagebox.showerror("Грешка", f"Грешка при записване на JSON:\n{str(e)}")
            self.update_status_bar("Грешка при записване на JSON")
    
    def fix_encoding_utf8_to_windows1251(self, text):
        """
        Поправя текст използвайки работещия метод: UTF-8→Latin-1→Windows-1251
        """
        try:
            step1 = text.encode('latin-1', errors='ignore')
            result = step1.decode('windows-1251', errors='ignore')
            return result
        except:
            return text


def main():
    """Главна функция"""
    root = tk.Tk()
    app = KasiExtractor(root)
    root.mainloop()


if __name__ == "__main__":
    main()