"""
Kasi Extractor - GUI Приложение за извличане на данни от MDB
"""

from tkinter import ttk, filedialog, messagebox
# from tkcalendar import DateEntry
from datetime import datetime, date
import tkinter as tk
import subprocess
import json
import csv
import sys
import io
import os

try:
    import pandas_access as mdb
    import pandas as pd
    PANDAS_ACCESS_AVAILABLE = True
except ImportError:
    PANDAS_ACCESS_AVAILABLE = False

class KasiExtractor:
    def __init__(self, root):
        self.root = root
        self.root.title("SMS Notification Clients v1.0")
        self.root.geometry("900x700")
        self.root.resizable(True, True)

        self.filtered_data_lines = []  # За запазване на филтрираните данни
        
        # Променливи
        self.mdb_file_path = tk.StringVar()

        # Променливи за дати
        self.start_date = tk.StringVar()
        self.end_date = tk.StringVar()
        
        # Задаване на начални дати (последните 30 дни)
        today = date.today()
        month_ago = date(today.year, today.month-1 if today.month > 1 else 12, today.day)
        self.start_date.set(month_ago.strftime('%d.%m.%Y'))
        self.end_date.set(today.strftime('%d.%m.%Y'))
        
        # Създаване на интерфейса
        self.create_widgets()

        # Задаваме днешни дати като по подразбиране
        self.set_default_dates()

    def validate_date_input(self, date_string):
        """Валидира дата в формат dd.mm.yyyy"""
        if not date_string.strip():
            return "empty"  # Празно поле
        
        try:
            # Проверка на дължината
            if len(date_string) != 10:
                return "invalid"
            
            # Проверка на формата с точки
            if date_string.count('.') != 2:
                return "invalid"
            
            # Парсиране на датата
            datetime.strptime(date_string, '%d.%m.%Y')
            return "valid"
        except ValueError:
            return "invalid"

    def validate_date_range(self):
        """Проверява дали крайната дата е след началната"""
        start_text = self.start_date_entry.get().strip()
        end_text = self.end_date_entry.get().strip()
        
        # Ако някое поле е празно, не проверяваме последователността
        if not start_text or not end_text:
            return True
        
        # Ако някоя дата е невалидна, не проверяваме последователността  
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
        
        # Първо проверяваме формата на датата
        if validation_result == "valid":
            entry_widget.config(bg="lightgreen")  # Зелен фон за валидна дата
        elif validation_result == "empty":
            entry_widget.config(bg="white")  # Бял фон за празно поле
        else:
            entry_widget.config(bg="lightcoral")  # Червен фон за невалидна дата
            return
        
        # Ако датата е валидна, проверяваме последователността
        if validation_result == "valid":
            if not self.validate_date_range():
                # Ако крайната дата е преди началната, правим фона оранжев
                self.start_date_entry.config(bg="orange")
                self.end_date_entry.config(bg="orange")
                self.update_status_bar("ГРЕШКА: Крайната дата е преди началната!")
            else:
                # Ако всичко е наред, възстановяваме зеления цвят
                if self.validate_date_input(self.start_date_entry.get()) == "valid":
                    self.start_date_entry.config(bg="lightgreen")
                if self.validate_date_input(self.end_date_entry.get()) == "valid":
                    self.end_date_entry.config(bg="lightgreen")
                self.update_status_bar("Готов за работа")
            
    def create_widgets(self):
        """Създава всички UI елементи"""
        
        # Главна рамка
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Конфигурация на grid
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        
        # 1. СЕКЦИЯ: ИЗБОР НА MDB ФАЙЛ
        file_frame = ttk.LabelFrame(main_frame, text="📁 Избор на MDB файл", padding="10")
        file_frame.grid(row=0, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        file_frame.columnconfigure(1, weight=1)
        
        # Бутон за избор на файл
        ttk.Button(file_frame, text="Избери MDB файл", 
                  command=self.select_mdb_file).grid(row=0, column=0, padx=(0, 10))
        
        # Поле за показване на избрания файл
        self.file_entry = ttk.Entry(file_frame, textvariable=self.mdb_file_path, 
                                   state="readonly", width=50)
        self.file_entry.grid(row=0, column=1, sticky=(tk.W, tk.E))
        
        # 2. СЕКЦИЯ: СТАТУС НА ФАЙЛА
        status_frame = ttk.LabelFrame(main_frame, text="📊 Информация за файла", padding="10")
        status_frame.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        status_frame.columnconfigure(0, weight=1)
        
        # Статус лейбъл
        self.status_label = ttk.Label(status_frame, text="Няма избран файл", 
                                     foreground="gray")
        self.status_label.grid(row=0, column=0, sticky=tk.W)
        
        # 3. СЕКЦИЯ: ТЕСТ НА ВРЪЗКАТА (временно за тестване)
        test_frame = ttk.LabelFrame(main_frame, text="🔧 Тестване", padding="10")
        test_frame.grid(row=2, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        
        # Бутон за тест на таблици
        self.test_button = ttk.Button(test_frame, text="Тествай връзка с базата", 
                                     command=self.test_database_connection, 
                                     state="disabled")
        self.test_button.grid(row=0, column=0, padx=(0, 10))
        
        # 4. СЕКЦИЯ: СТАТУС БАР (долу)
        status_bar_frame = ttk.Frame(main_frame)
        status_bar_frame.grid(row=10, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(20, 0))
        status_bar_frame.columnconfigure(0, weight=1)

        # 5. СЕКЦИЯ: ИЗБОР НА ДАТИ ЗА ФИЛТРИРАНЕ
        date_frame = ttk.LabelFrame(main_frame, text="📅 Филтриране по дати", padding="10")
        date_frame.grid(row=3, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        date_frame.columnconfigure(1, weight=1)
        date_frame.columnconfigure(3, weight=1)
        
        # От дата
        ttk.Label(date_frame, text="От дата:").grid(row=0, column=0, padx=(0, 5), sticky=tk.W)
        self.start_date_entry = tk.Entry(date_frame, width=12)
        self.start_date_entry.grid(row=0, column=1, padx=(0, 20), sticky=tk.W)
        # Добави event binding за real-time validation
        self.start_date_entry.bind('<KeyRelease>', lambda e: self.on_date_entry_change(e, self.start_date_entry))

        # До дата
        ttk.Label(date_frame, text="До дата:").grid(row=0, column=2, padx=(0, 5), sticky=tk.W)
        self.end_date_entry = tk.Entry(date_frame, width=12)
        self.end_date_entry.grid(row=0, column=3, padx=(0, 20), sticky=tk.W)
        # Добави event binding за real-time validation
        self.end_date_entry.bind('<KeyRelease>', lambda e: self.on_date_entry_change(e, self.end_date_entry))
        
        # Бутон за филтриране
        self.filter_button = ttk.Button(date_frame, text="📊 Филтрирай данните", 
                                       command=self.filter_data, state="disabled")
        self.filter_button.grid(row=0, column=4, padx=(20, 0))
        
        # Инструкции с пример
        instruction_label = ttk.Label(date_frame, text="Формат: dd.mm.yyyy (например: 10.09.2025)", 
                                     foreground="gray", font=("TkDefaultFont", 8))
        instruction_label.grid(row=1, column=0, columnspan=4, pady=(5, 0), sticky=tk.W)
        
        # Резултат от филтрирането
        self.filter_result_label = ttk.Label(date_frame, text="", foreground="gray")
        self.filter_result_label.grid(row=2, column=0, columnspan=5, pady=(10, 0), sticky=tk.W)

        # 6. СЕКЦИЯ: ИЗВЛИЧАНЕ НА КОНКРЕТНИ КОЛОНИ
        extract_frame = ttk.LabelFrame(main_frame, text="📋 Извличане на данни", padding="10")
        extract_frame.grid(row=4, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        extract_frame.columnconfigure(0, weight=1)
        
        # Информация за колоните
        info_label = ttk.Label(extract_frame, 
                              text="Колони за извличане: Number, End_Data, Model, Number_EKA, Ime_Obekt, Adres_Obekt, Dan_Number, Phone, Ime_Firma, bulst",
                              foreground="gray", font=("TkDefaultFont", 8), wraplength=500)
        info_label.grid(row=0, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        
        # Бутон за извличане
        self.extract_button = ttk.Button(extract_frame, text="📊 Извлечи колони", 
                                        command=self.extract_specific_columns, state="disabled")
        self.extract_button.grid(row=1, column=0, padx=(0, 10))
        
        # Бутони за запис (неактивни до извличане)
        self.save_csv_button = ttk.Button(extract_frame, text="💾 Запиши CSV", 
                                         command=self.save_csv, state="disabled")
        self.save_csv_button.grid(row=1, column=1, padx=(0, 10))
        
        self.save_json_button = ttk.Button(extract_frame, text="💾 Запиши JSON", 
                                          command=self.save_json, state="disabled")
        self.save_json_button.grid(row=1, column=2)

        # 7. СЕКЦИЯ: ПЪЛЕН ЕКСПОРТ НА ТАБЛИЦА
        export_frame = ttk.LabelFrame(main_frame, text="📤 Пълен експорт", padding="10")
        export_frame.grid(row=5, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        export_frame.columnconfigure(0, weight=1)
        
        # Информация
        export_info_label = ttk.Label(export_frame, 
                                     text="Експортиране на цялата таблица Kasi_all (всички колони, всички редове)",
                                     foreground="gray", font=("TkDefaultFont", 8))
        export_info_label.grid(row=0, column=0, sticky=tk.W, pady=(0, 10))
        
        # Бутон за пълен експорт
        self.full_export_button = ttk.Button(export_frame, text="📁 Експортирай цяла таблица", 
                                            command=self.export_full_table, state="disabled")
        self.full_export_button.grid(row=1, column=0, sticky=tk.W)
        
        # Резултат от извличането
        self.extract_result_label = ttk.Label(extract_frame, text="", foreground="gray")
        self.extract_result_label.grid(row=2, column=0, columnspan=3, pady=(10, 0), sticky=tk.W)
        
        # Статус бар
        self.status_bar = ttk.Label(status_bar_frame, text="Готов за работа", 
                                   relief=tk.SUNKEN, anchor=tk.W, padding="5")
        self.status_bar.grid(row=0, column=0, sticky=(tk.W, tk.E))
        
        # Бутон за изход
        ttk.Button(status_bar_frame, text="Изход", 
                  command=self.exit_application).grid(row=0, column=1, padx=(10, 0))
    
    def set_default_dates(self):
        """Задава днешна дата като период по подразбиране"""
        try:
            from datetime import date
            today = date.today()
            today_str = today.strftime('%d.%m.%Y')
            
            # Задаваме днешната дата в двете полета
            self.start_date_entry.delete(0, tk.END)
            self.start_date_entry.insert(0, today_str)
            
            self.end_date_entry.delete(0, tk.END)
            self.end_date_entry.insert(0, today_str)
            
            self.update_status_bar(f"Зададен е период: {today_str} - {today_str}")
            
        except Exception as e:
            print(f"Предупреждение: Не можах да задам началните дати: {e}")
            self.update_status_bar("Готов за работа")

    def select_mdb_file(self):
        """Отваря диалог за избор на MDB файл"""
        file_path = filedialog.askopenfilename(
            title="Избери MDB файл",
            filetypes=[
                ("MDB файлове", "*.mdb"),
                ("Всички файлове", "*.*")
            ]
        )
        
        if file_path:
            self.mdb_file_path.set(file_path)
            self.update_file_status(file_path)
            self.update_status_bar(f"Избран файл: {os.path.basename(file_path)}")
    
    def update_file_status(self, file_path):
        """Обновява статуса на избрания файл"""
        if os.path.exists(file_path):
            file_size = os.path.getsize(file_path)
            size_mb = file_size / (1024 * 1024)
            
            status_text = f"✅ Файл: {os.path.basename(file_path)} ({size_mb:.1f} MB)"
            self.status_label.config(text=status_text, foreground="green")
            
            # Активираме бутоните
            self.test_button.config(state="normal")
            self.filter_button.config(state="normal")
            self.full_export_button.config(state="normal")
        else:
            self.status_label.config(text="❌ Файлът не съществува", foreground="red")
            self.test_button.config(state="disabled")
    
    def test_database_connection(self):
        """Тества връзката с базата данни и показва таблиците"""
        if not self.mdb_file_path.get():
            messagebox.showerror("Грешка", "Моля изберете MDB файл първо!")
            return
        
        if not PANDAS_ACCESS_AVAILABLE:
            messagebox.showerror("Грешка", "pandas_access не е инсталиран! Моля инсталирайте го с: pip install pandas_access")
            return
        
        self.update_status_bar("Тестване на връзката с базата...")
        
        try:
            # Използваме pandas_access
            tables = list(mdb.list_tables(self.mdb_file_path.get()))
            self._show_tables_result(tables)
            
        except Exception as e:
            messagebox.showerror("Грешка", f"Неочаквана грешка:\n{str(e)}")
            self.update_status_bar(f"Грешка: {str(e)}")

    def _show_tables_result(self, tables):
        """Показва резултата от намерените таблици"""
        if "Kasi_all" in tables:
            messagebox.showinfo("Успех", 
                            f"✅ Връзката е успешна!\n\n"
                            f"Намерени таблици: {len(tables)}\n"
                            f"Таблица 'Kasi_all': ✅ Намерена")
            self.update_status_bar("✅ Базата е готова за работа")
        else:
            messagebox.showwarning("Внимание", 
                                f"Таблица 'Kasi_all' не е намерена!\n\n"
                                f"Налични таблици:\n" + "\n".join(tables))
            self.update_status_bar("⚠️ Таблица 'Kasi_all' не е намерена")
    
    def filter_data(self):
        """Филтрира данните по избраните дати"""
        if not self.mdb_file_path.get():
            messagebox.showerror("Грешка", "Моля изберете MDB файл първо!")
            return
        
        if not PANDAS_ACCESS_AVAILABLE:
            messagebox.showerror("Грешка", "pandas_access не е инсталиран!")
            return
        
        try:
            start_date_str = self.start_date_entry.get().strip()
            end_date_str = self.end_date_entry.get().strip()
            
            if not start_date_str or not end_date_str:
                messagebox.showerror("Грешка", "Моля въведете начална и крайна дата!")
                return
            
            # Проверка на последователността на датите
            if not self.validate_date_range():
                messagebox.showerror("Грешка", "Крайната дата не може да бъде преди началната дата!")
                return
                
        except Exception as e:
            messagebox.showerror("Грешка", f"Проблем с четенето на датите:\n{str(e)}")
            return

        self.update_status_bar(f"Филтриране от {start_date_str} до {end_date_str}...")
        self.root.update_idletasks()
        
        try:
            # Четене на цялата таблица с pandas_access
            df = mdb.read_table(self.mdb_file_path.get(), "Kasi_all")
            
            # Парсиране на датите за филтриране
            start_date = datetime.strptime(start_date_str, '%d.%m.%Y')
            end_date = datetime.strptime(end_date_str, '%d.%m.%Y')
            
            # Филтриране по End_Data колоната
            if 'End_Data' not in df.columns:
                messagebox.showerror("Грешка", "Колона 'End_Data' не е намерена в таблицата!")
                return False
            
            # Конвертиране на End_Data към datetime
            # Опитваме различни формати дати
            try:
                df['End_Data_parsed'] = pd.to_datetime(df['End_Data'], format='%m/%d/%y %H:%M:%S', errors='coerce')
            except:
                try:
                    df['End_Data_parsed'] = pd.to_datetime(df['End_Data'], format='%m/%d/%Y %H:%M:%S', errors='coerce')
                except:
                    df['End_Data_parsed'] = pd.to_datetime(df['End_Data'], errors='coerce')
            
            # Филтриране по дати
            mask = (df['End_Data_parsed'].dt.date >= start_date.date()) & \
                (df['End_Data_parsed'].dt.date <= end_date.date())
            filtered_df = df[mask]
            
            # Запазване на филтрираните данни като CSV lines
            self.filtered_data_lines = []
            
            # Header
            columns = list(filtered_df.columns)
            if 'End_Data_parsed' in columns:
                columns.remove('End_Data_parsed')  # Премахваме помощната колона
            self.filtered_data_lines.append(','.join(f'"{col}"' for col in columns))
            
            # Данни
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
            
            total_rows = len(filtered_df)
            original_rows = len(df)
            percent = (total_rows/original_rows*100) if original_rows > 0 else 0
            
            result_text = f"✅ Филтрирани {total_rows} от общо {original_rows} реда"
            detailed_result = f"{result_text} ({percent:.1f}%)"
            self.filter_result_label.config(text=detailed_result, foreground="green")
            self.update_status_bar(f"Филтриране завършено: {total_rows} от {original_rows} реда ({percent:.1f}%)")
            
            messagebox.showinfo("Резултат", f"Филтрирането е завършено!\n\nПериод: {start_date_str} - {end_date_str}\nОбщо редове: {original_rows}\nФилтрирани редове: {total_rows}")
            
            # Активираме бутона за извличане
            self.extract_button.config(state="normal")
            return True
            
        except Exception as e:
            messagebox.showerror("Грешка", f"Неочаквана грешка:\n{str(e)}")
            self.update_status_bar(f"Грешка: {str(e)}")
            return False

    def extract_specific_columns(self):
        """Извлича конкретните 10 колони от филтрираните данни"""
        if not hasattr(self, 'filtered_data_lines') or len(self.filtered_data_lines) < 2:
            messagebox.showerror("Грешка", "Няма филтрирани данни! Първо направете филтрация.")
            return False
        
        self.update_status_bar("Извличане на конкретни колони...")
        
        # Колоните които ни трябват
        required_columns = ['Number', 'End_Data', 'Model', 'Number_EKA', 'Ime_Obekt', 
                        'Adres_Obekt', 'Dan_Number', 'Phone', 'Ime_Firma', 'bulst']
        
        try:
            # Намираме индексите на колоните
            header_line = self.filtered_data_lines[0]
            header_reader = csv.reader(io.StringIO(header_line))
            headers = next(header_reader)
            
            # Мапинг на колони към индекси
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
            
            # Създаваме новия header
            new_header = [col for col in required_columns if col in column_indices]
            
            # Извличаме данните
            extracted_data = []
            extracted_data.append(','.join(f'"{col}"' for col in new_header))  # Header
            
            for line in self.filtered_data_lines[1:]:
                try:
                    reader = csv.reader(io.StringIO(line))
                    fields = next(reader)
                    
                    # Извличаме нужните полета
                    new_row = []
                    for col_name in new_header:
                        if column_indices[col_name] < len(fields):
                            field_value = fields[column_indices[col_name]]
                            # Поправяме кодировката само на Linux
                            if sys.platform != "win32":
                                fixed_value = self.fix_encoding_utf8_to_windows1251(field_value)
                            else:
                                fixed_value = field_value
                            new_row.append(f'"{fixed_value}"')
                        else:
                            new_row.append('""')  # Празно поле ако няма данни
                    
                    extracted_data.append(','.join(new_row))
                
                except Exception as e:
                    # Прескачаме проблемни редове
                    continue
            
            # Запазваме извлечените данни
            self.extracted_data_lines = extracted_data
            
            # Показваме резултата
            total_extracted = len(extracted_data) - 1  # Без header-а
            
            result_text = f"✅ Извлечени {len(new_header)} колони от {total_extracted} реда"
            if hasattr(self, 'filtered_data_lines'):
                original_rows = len(self.filtered_data_lines) - 1
                result_text += f" (от {original_rows} филтрирани)"
            
            self.extract_result_label.config(text=result_text, foreground="green")
            self.update_status_bar(f"Извличане завършено: {total_extracted} реда с {len(new_header)} колони")
            
            # Активираме бутоните за запис
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
        """Експортира цялата таблица Kasi_all в CSV формат"""
        if not self.mdb_file_path.get():
            messagebox.showerror("Грешка", "Моля изберете MDB файл първо!")
            return
        
        if not PANDAS_ACCESS_AVAILABLE:
            messagebox.showerror("Грешка", "pandas_access не е инсталиран!")
            return
        
        # Избор на файл за запис
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
            
            # Четене на цялата таблица с pandas_access
            df = mdb.read_table(self.mdb_file_path.get(), "Kasi_all")
            
            # Поправяме кодировката на всички string колони
            for column in df.columns:
                if df[column].dtype == 'object':  # string колони
                    df[column] = df[column].astype(str).apply(
                        lambda x: self.fix_encoding_utf8_to_windows1251(x) if x != 'nan' else ''
                    )
            
            # Записваме директно с pandas
            df.to_csv(file_path, index=False, encoding='utf-8')
            
            # Статистики
            total_rows = len(df)
            total_columns = len(df.columns)
            file_size = os.path.getsize(file_path)
            
            self.update_status_bar(f"Пълен експорт завършен: {os.path.basename(file_path)}")
            
            messagebox.showinfo("Успех", 
                            f"Пълният експорт е завършен успешно!\n\n"
                            f"📁 Файл: {os.path.basename(file_path)}\n"
                            f"📊 Редове: {total_rows:,}\n"
                            f"📋 Колони: {total_columns}\n"
                            f"💾 Размер: {file_size / 1024 / 1024:.1f} MB\n"
                            f"🔗 Път: {file_path}")
            
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
        
        # Избор на файл за запис
        file_path = filedialog.asksaveasfilename(
            title="Запиши като CSV",
            defaultextension=".csv",
            filetypes=[("CSV файлове", "*.csv"), ("Всички файлове", "*.*")]
        )
        
        if not file_path:
            return
        
        try:
            self.update_status_bar("Записване на CSV файл...")
            
            # Записваме данните
            with open(file_path, 'w', encoding='utf-8', newline='') as f:
                for line in self.extracted_data_lines:
                    f.write(line + '\n')
            
            # Статистики
            total_rows = len(self.extracted_data_lines) - 1  # Без header
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
        """Запис в JSON формат като масив от обекти"""
        if not hasattr(self, 'extracted_data_lines') or len(self.extracted_data_lines) < 2:
            messagebox.showerror("Грешка", "Няма извлечени данни за запис!")
            return
        
        # Избор на файл за запис
        file_path = filedialog.asksaveasfilename(
            title="Запиши като JSON",
            defaultextension=".json",
            filetypes=[("JSON файлове", "*.json"), ("Всички файлове", "*.*")]
        )
        
        if not file_path:
            return
        
        try:
            self.update_status_bar("Записване на JSON файл...")
            
            # Парсираме header-а
            header_line = self.extracted_data_lines[0]
            header_reader = csv.reader(io.StringIO(header_line))
            headers = next(header_reader)
            
            # Създаваме масив от обекти
            json_data = []
            
            for line in self.extracted_data_lines[1:]:
                try:
                    reader = csv.reader(io.StringIO(line))
                    fields = next(reader)
                    
                    # Създаваме обект за този ред
                    row_object = {}
                    for i, header in enumerate(headers):
                        if i < len(fields):
                            row_object[header] = fields[i]
                        else:
                            row_object[header] = ""
                    
                    json_data.append(row_object)
                
                except Exception as e:
                    # Прескачаме проблемни редове
                    continue
            
            # Записваме JSON файла
            import json
            with open(file_path, 'w', encoding='utf-8') as f:
                json.dump(json_data, f, ensure_ascii=False, indent=2)
            
            # Статистики
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
    
    # ========== ОРИГИНАЛНИ ФУНКЦИИ ЗА КОДИРОВКА ==========
    
    def fix_encoding_utf8_to_windows1251(self, text):
        """
        Поправя текст използвайки работещия метод: UTF-8→Latin-1→Windows-1251
        (Запазена оригинална функция)
        """
        try:
            # Работещия метод от теста
            step1 = text.encode('latin-1', errors='ignore')
            result = step1.decode('windows-1251', errors='ignore')
            return result
        except:
            return text  # Ако има проблем, връща оригинала


def main():
    """Главна функция"""
    # Проверяваме за mdb-tools при стартиране
    try:
        subprocess.run(["mdb-tables", "--help"], capture_output=True)
    except FileNotFoundError:
        print("⚠️ ВНИМАНИЕ: mdb-tools не е намерен!")
        print("Моля инсталирайте mdb-tools за да работи приложението.")
    
    # Стартираме GUI
    root = tk.Tk()
    app = KasiExtractor(root)
    root.mainloop()


if __name__ == "__main__":
    main()