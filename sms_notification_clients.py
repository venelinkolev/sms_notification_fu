#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Kasi Extractor - GUI Приложение за извличане на данни от MDB
Етап 1: Избор на MDB файл + основна GUI структура
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

# Условен import за pyodbc (само на Windows)
PYODBC_AVAILABLE = False
if sys.platform == "win32":
    try:
        import pyodbc
        PYODBC_AVAILABLE = True
    except ImportError:
        PYODBC_AVAILABLE = False

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

        # # 5. СЕКЦИЯ: ИЗБОР НА ДАТИ ЗА ФИЛТРИРАНЕ С КАЛЕНДАРИ
        # date_frame = ttk.LabelFrame(main_frame, text="📅 Филтриране по дати", padding="10")
        # date_frame.grid(row=3, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        
        # # Създаваме вътрешна рамка за по-добро подреждане
        # inner_frame = ttk.Frame(date_frame)
        # inner_frame.grid(row=0, column=0, sticky=(tk.W, tk.E))
        # inner_frame.columnconfigure(1, weight=1)
        # inner_frame.columnconfigure(3, weight=1)
        
        # # От дата
        # ttk.Label(inner_frame, text="От дата:").grid(row=0, column=0, padx=(0, 10), sticky=tk.W)
        # try:
        #     self.start_date_entry = DateEntry(inner_frame, width=12, 
        #                                      date_pattern='dd.mm.yyyy',
        #                                      state='readonly')
        #     self.start_date_entry.grid(row=0, column=1, padx=(0, 30), sticky=tk.W)
        # except Exception as e:
        #     print(f"Грешка при създаване на първия календар: {e}")
        
        # # До дата  
        # ttk.Label(inner_frame, text="До дата:").grid(row=0, column=2, padx=(0, 10), sticky=tk.W)
        # try:
        #     self.end_date_entry = DateEntry(inner_frame, width=12,
        #                                    date_pattern='dd.mm.yyyy',
        #                                    state='readonly')
        #     self.end_date_entry.grid(row=0, column=3, padx=(0, 30), sticky=tk.W)
        # except Exception as e:
        #     print(f"Грешка при създаване на втория календар: {e}")
        
        # # Бутон за филтриране на нов ред
        # button_frame = ttk.Frame(date_frame)
        # button_frame.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=(10, 0))
        
        # self.filter_button = ttk.Button(button_frame, text="📊 Филтрирай данните", 
        #                                command=self.filter_data, state="disabled")
        # self.filter_button.grid(row=0, column=0, sticky=tk.W)
        
        # # Инструкции
        # instruction_label = ttk.Label(button_frame, text="← Натиснете на календара за избор на дата", 
        #                              foreground="gray", font=("TkDefaultFont", 8))
        # instruction_label.grid(row=0, column=1, padx=(20, 0), sticky=tk.W)
        
        # # Резултат от филтрирането
        # self.filter_result_label = ttk.Label(date_frame, text="", foreground="gray")
        # self.filter_result_label.grid(row=2, column=0, pady=(10, 0), sticky=tk.W)

        # # 5. СЕКЦИЯ: ИЗБОР НА ДАТИ ЗА ФИЛТРИРАНЕ
        # date_frame = ttk.LabelFrame(main_frame, text="📅 Филтриране по дати", padding="10")
        # date_frame.grid(row=3, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        # date_frame.columnconfigure(1, weight=1)
        # date_frame.columnconfigure(3, weight=1)
        
        # # От дата
        # ttk.Label(date_frame, text="От дата (dd.mm.yyyy):").grid(row=0, column=0, padx=(0, 5), sticky=tk.W)
        # self.start_date_entry = ttk.Entry(date_frame, width=12)
        # self.start_date_entry.grid(row=0, column=1, padx=(0, 20), sticky=tk.W)
        # self.start_date_entry.insert(0, "01.04.2009")  # Примерна начална дата
        
        # # До дата
        # ttk.Label(date_frame, text="До дата (dd.mm.yyyy):").grid(row=0, column=2, padx=(0, 5), sticky=tk.W)
        # self.end_date_entry = ttk.Entry(date_frame, width=12)
        # self.end_date_entry.grid(row=0, column=3, padx=(0, 20), sticky=tk.W)
        # self.end_date_entry.insert(0, "31.12.2009")  # Примерна крайна дата
        
        # # Бутон за филтриране
        # self.filter_button = ttk.Button(date_frame, text="📊 Филтрирай данните", 
        #                                command=self.filter_data, state="disabled")
        # self.filter_button.grid(row=0, column=4, padx=(20, 0))
        
        # # Инструкции
        # instruction_label = ttk.Label(date_frame, text="Формат: dd.mm.yyyy (например: 01.08.2025)", 
        #                              foreground="gray", font=("TkDefaultFont", 8))
        # instruction_label.grid(row=1, column=0, columnspan=4, pady=(5, 0), sticky=tk.W)
        
        # # Резултат от филтрирането
        # self.filter_result_label = ttk.Label(date_frame, text="", foreground="gray")
        # self.filter_result_label.grid(row=2, column=0, columnspan=5, pady=(10, 0), sticky=tk.W)

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
        
        self.update_status_bar("Тестване на връзката с базата...")
        
        # На Windows използваме pyodbc, на Linux - mdb-tools
        if sys.platform == "win32" and PYODBC_AVAILABLE:
            self._test_with_pyodbc()
        else:
            self._test_with_mdb_tools()

    def _test_with_pyodbc(self):
        import pyodbc  # Local import само когато е нужно
        """Тест с pyodbc за Windows"""
        try:
            # Опитваме различни драйвери по реда на приоритет
            driver_options = [
                'Microsoft Access Driver (*.mdb, *.accdb)',
                'Microsoft Access Driver (*.mdb)',
                'Microsoft Access Driver (*.accdb)',
                'Microsoft Office Access Driver (*.mdb, *.accdb)',
                'Microsoft Office 16.0 Access Database Engine OLE DB Provider'
            ]

            conn = None
            for driver in driver_options:
                try:
                    conn_str = f'DRIVER={{{driver}}};DBQ={self.mdb_file_path.get()};'
                    conn = pyodbc.connect(conn_str)
                    break  # Ако успее, спираме тук
                except:
                    continue  # Опитваме следващия драйвер

            if conn is None:
                raise Exception("Не може да се намери подходящ Access драйвер")
            conn = pyodbc.connect(conn_str)
            cursor = conn.cursor()
            
            tables = [table_info.table_name for table_info in cursor.tables(tableType='TABLE')]
            conn.close()
            
            self._show_tables_result(tables)
            
        except Exception as e:
            messagebox.showerror("Грешка", f"Не можах да се свържа с базата:\n{str(e)}")
            self.update_status_bar("Грешка при свързване с базата")

    def _test_with_mdb_tools(self):
        """Тест с mdb-tools за Linux"""
        try:
            result = subprocess.run(["mdb-tables", self.mdb_file_path.get()], 
                                capture_output=True, text=True)
            
            if result.returncode != 0:
                messagebox.showerror("Грешка", f"Не можах да чета базата:\n{result.stderr}")
                return
            
            tables = result.stdout.strip().split()
            self._show_tables_result(tables)
            
        except FileNotFoundError:
            messagebox.showerror("Грешка", "mdb-tools не е намерен!")
            self.update_status_bar("Грешка: mdb-tools не е инсталиран")

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

    # def test_database_connection(self):
    #     """Тества връзката с базата данни и показва таблиците"""
    #     if not self.mdb_file_path.get():
    #         messagebox.showerror("Грешка", "Моля изберете MDB файл първо!")
    #         return
        
    #     self.update_status_bar("Тестване на връзката с базата...")
        
    #     try:
    #         # Проверяваме дали mdb-tables е инсталиран
    #         result = subprocess.run(["mdb-tables", "--help"], 
    #                                capture_output=True, text=True)
    #         if result.returncode != 0:
    #             raise FileNotFoundError("mdb-tables не работи")
    #     except FileNotFoundError:
    #         messagebox.showerror("Грешка", 
    #                            "mdb-tools не е намерен!\n\n"
    #                            "Моля инсталирайте mdb-tools:\n"
    #                            "- Windows: choco install mdb-tools\n"
    #                            "- Ubuntu: sudo apt-get install mdb-tools")
    #         self.update_status_bar("Грешка: mdb-tools не е инсталиран")
    #         return
        
    #     try:
    #         # Списък на таблиците
    #         result = subprocess.run(["mdb-tables", self.mdb_file_path.get()], 
    #                                capture_output=True, text=True)
            
    #         if result.returncode != 0:
    #             messagebox.showerror("Грешка", f"Не можах да чета базата:\n{result.stderr}")
    #             self.update_status_bar("Грешка при четене на базата")
    #             return
            
    #         tables = result.stdout.strip().split()
            
    #         if "Kasi_all" in tables:
    #             messagebox.showinfo("Успех", 
    #                                f"✅ Връзката е успешна!\n\n"
    #                                f"Намерени таблици: {len(tables)}\n"
    #                                f"Таблица 'Kasi_all': ✅ Намерена\n\n"
    #                                f"Всички таблици:\n" + "\n".join(tables[:10]) + 
    #                                (f"\n... и още {len(tables)-10}" if len(tables) > 10 else ""))
    #             self.update_status_bar("✅ Базата е готова за работа")
    #         else:
    #             messagebox.showwarning("Внимание", 
    #                                   f"Таблица 'Kasi_all' не е намерена!\n\n"
    #                                   f"Налични таблици:\n" + "\n".join(tables))
    #             self.update_status_bar("⚠️ Таблица 'Kasi_all' не е намерена")
                
    #     except Exception as e:
    #         messagebox.showerror("Грешка", f"Неочаквана грешка:\n{str(e)}")
    #         self.update_status_bar(f"Грешка: {str(e)}")
    
    def filter_data(self):
        """Филтрира данните по избраните дати"""
        if not self.mdb_file_path.get():
            messagebox.showerror("Грешка", "Моля изберете MDB файл първо!")
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
        
        # Използваме подходящия метод за платформата
        if sys.platform == "win32" and PYODBC_AVAILABLE:
            success = self._filter_data_with_pyodbc(start_date_str, end_date_str)
        else:
            success = self._filter_data_with_mdb_tools(start_date_str, end_date_str)
        
        if success:
            self.extract_button.config(state="normal")

    def _filter_data_with_pyodbc(self, start_date_str, end_date_str):
        import pyodbc  # Local import само когато е нужно
        """Филтриране с pyodbc за Windows"""
        try:
            start_date = datetime.strptime(start_date_str, '%d.%m.%Y')
            end_date = datetime.strptime(end_date_str, '%d.%m.%Y')
            
            # Опитваме различни драйвери по реда на приоритет
            driver_options = [
                'Microsoft Access Driver (*.mdb, *.accdb)',
                'Microsoft Access Driver (*.mdb)',
                'Microsoft Access Driver (*.accdb)',
                'Microsoft Office Access Driver (*.mdb, *.accdb)',
                'Microsoft Office 16.0 Access Database Engine OLE DB Provider'
            ]

            conn = None
            for driver in driver_options:
                try:
                    conn_str = f'DRIVER={{{driver}}};DBQ={self.mdb_file_path.get()};'
                    conn = pyodbc.connect(conn_str)
                    break  # Ако успее, спираме тук
                except:
                    continue  # Опитваме следващия драйвер

            if conn is None:
                raise Exception("Не може да се намери подходящ Access драйвер")
            conn = pyodbc.connect(conn_str)
            cursor = conn.cursor()
            
            # SQL заявка с филтриране по дата
            query = """
            SELECT * FROM Kasi_all 
            WHERE End_Data >= ? AND End_Data <= ?
            """
            
            cursor.execute(query, start_date, end_date)
            rows = cursor.fetchall()
            
            # Получаваме имената на колоните
            columns = [column[0] for column in cursor.description]
            
            # Конвертираме в CSV формат
            self.filtered_data_lines = []
            # Header
            self.filtered_data_lines.append(','.join(f'"{col}"' for col in columns))
            
            # Данни
            for row in rows:
                csv_row = []
                for value in row:
                    if value is None:
                        csv_row.append('""')
                    else:
                        # Конвертираме към string и escape-ваме кавички
                        str_value = str(value).replace('"', '""')
                        csv_row.append(f'"{str_value}"')
                self.filtered_data_lines.append(','.join(csv_row))
            
            conn.close()
            
            total_rows = len(self.filtered_data_lines) - 1
            percent = 100.0  # При SQL заявка всички редове са филтрирани
            
            result_text = f"✅ Филтрирани {total_rows} реда"
            detailed_result = f"{result_text} (100%)"
            self.filter_result_label.config(text=detailed_result, foreground="green")
            self.update_status_bar(f"Филтриране завършено: {total_rows} реда")
            
            messagebox.showinfo("Резултат", f"Филтрирането е завършено!\n\nПериод: {start_date_str} - {end_date_str}\nФилтрирани редове: {total_rows}")
            return True
            
        except Exception as e:
            messagebox.showerror("Грешка", f"Грешка при филтриране:\n{str(e)}")
            self.update_status_bar(f"Грешка: {str(e)}")
            return False

    def _filter_data_with_mdb_tools(self, start_date_str, end_date_str):
        """Филтриране с mdb-tools за Linux (запазва оригиналния код)"""
        try:
            start_date = datetime.strptime(start_date_str, '%d.%m.%Y')
            end_date = datetime.strptime(end_date_str, '%d.%m.%Y')
            
            # Извличане на данните от таблицата
            result = subprocess.run(["mdb-export", self.mdb_file_path.get(), "Kasi_all"], 
                                capture_output=True, text=False)
            
            if result.returncode != 0:
                messagebox.showerror("Грешка", f"Не можах да извлека данните:\n{result.stderr}")
                self.update_status_bar("Грешка при извличане на данни")
                return False
            
            # Декодиране като UTF-8
            raw_content = result.stdout.decode('utf-8', errors='ignore')
            lines = raw_content.strip().split('\n')
            
            if len(lines) < 2:
                messagebox.showwarning("Внимание", "Таблицата е празна или няма данни")
                return False
            
            # [Запазва целия оригинален код за филтриране с mdb-tools...]
            # Намираме индекса на End_Data колоната
            header_line = lines[0]
            header_reader = csv.reader(io.StringIO(header_line))
            headers = next(header_reader)
            
            end_data_index = None
            for i, header in enumerate(headers):
                if 'End_Data' in header:
                    end_data_index = i
                    break
            
            if end_data_index is None:
                messagebox.showerror("Грешка", "Колона 'End_Data' не е намерена в таблицата!")
                return False
            
            # Филтриране на данните
            filtered_lines = [lines[0]]  # Добавяме header-а
            total_rows = 0
            filtered_rows = 0
            
            for line in lines[1:]:
                total_rows += 1
                try:
                    reader = csv.reader(io.StringIO(line))
                    fields = next(reader)
                    
                    if len(fields) > end_data_index:
                        end_data_str = fields[end_data_index].strip()
                        
                        if end_data_str and len(end_data_str) >= 8:
                            date_part = end_data_str.split()[0]
                            
                            row_date = None
                            try:
                                temp_date = datetime.strptime(date_part, '%m/%d/%y')
                                if temp_date.year < 1950:
                                    temp_date = temp_date.replace(year=temp_date.year + 100)
                                row_date = temp_date
                            except ValueError:
                                for date_format in ['%m.%d.%Y', '%d.%m.%Y', '%Y-%m-%d', '%m/%d/%Y']:
                                    try:
                                        row_date = datetime.strptime(date_part, date_format)
                                        break
                                    except ValueError:
                                        continue
                            
                            if row_date:
                                if start_date.date() <= row_date.date() <= end_date.date():
                                    filtered_lines.append(line)
                                    filtered_rows += 1
                except Exception as e:
                    continue
            
            self.filtered_data_lines = filtered_lines
            
            percent = (filtered_rows/total_rows*100) if total_rows > 0 else 0
            result_text = f"✅ Филтрирани {filtered_rows} от общо {total_rows} реда"
            detailed_result = f"{result_text} ({percent:.1f}%)"
            self.filter_result_label.config(text=detailed_result, foreground="green")
            self.update_status_bar(f"Филтриране завършено: {filtered_rows} от {total_rows} реда ({percent:.1f}%)")
            
            messagebox.showinfo("Резултат", f"Филтрирането е завършено!\n\nПериод: {start_date_str} - {end_date_str}\nОбщо редове: {total_rows}\nФилтрирани редове: {filtered_rows}")
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
            
            # Използваме подходящия метод за платформата
            if sys.platform == "win32" and PYODBC_AVAILABLE:
                success = self._export_full_table_with_pyodbc(file_path)
            else:
                success = self._export_full_table_with_mdb_tools(file_path)
            
            if success:
                self.update_status_bar(f"Пълен експорт завършен: {os.path.basename(file_path)}")
            
        except Exception as e:
            messagebox.showerror("Грешка", f"Грешка при пълен експорт:\n{str(e)}")
            self.update_status_bar(f"Грешка: {str(e)}")

    def _export_full_table_with_pyodbc(self, file_path):
        import pyodbc  # Local import само когато е нужно
        """Пълен експорт с pyodbc за Windows"""
        try:
            # Опитваме различни драйвери по реда на приоритет
            driver_options = [
                'Microsoft Access Driver (*.mdb, *.accdb)',
                'Microsoft Access Driver (*.mdb)',
                'Microsoft Access Driver (*.accdb)',
                'Microsoft Office Access Driver (*.mdb, *.accdb)',
                'Microsoft Office 16.0 Access Database Engine OLE DB Provider'
            ]

            conn = None
            for driver in driver_options:
                try:
                    conn_str = f'DRIVER={{{driver}}};DBQ={self.mdb_file_path.get()};'
                    conn = pyodbc.connect(conn_str)
                    break  # Ако успее, спираме тук
                except:
                    continue  # Опитваме следващия драйвер

            if conn is None:
                raise Exception("Не може да се намери подходящ Access драйвер")
            conn = pyodbc.connect(conn_str)
            cursor = conn.cursor()
            
            cursor.execute("SELECT * FROM Kasi_all")
            rows = cursor.fetchall()
            columns = [column[0] for column in cursor.description]
            
            # Записваме файла
            with open(file_path, 'w', encoding='utf-8', newline='') as f:
                writer = csv.writer(f)
                writer.writerow(columns)  # Header
                
                for row in rows:
                    writer.writerow(row)
            
            conn.close()
            
            # Статистики
            total_rows = len(rows)
            file_size = os.path.getsize(file_path)
            total_columns = len(columns)
            
            messagebox.showinfo("Успех", 
                            f"Пълният експорт е завършен успешно!\n\n"
                            f"📁 Файл: {os.path.basename(file_path)}\n"
                            f"📊 Редове: {total_rows:,}\n"
                            f"📋 Колони: {total_columns}\n"
                            f"💾 Размер: {file_size / 1024 / 1024:.1f} MB\n"
                            f"🔗 Път: {file_path}")
            return True
            
        except Exception as e:
            messagebox.showerror("Грешка", f"Грешка при експорт с pyodbc:\n{str(e)}")
            return False

    def _export_full_table_with_mdb_tools(self, file_path):
        """Пълен експорт с mdb-tools за Linux (запазва оригиналния код)"""
        try:
            # Извличане на всички данни от таблицата
            result = subprocess.run(["mdb-export", self.mdb_file_path.get(), "Kasi_all"], 
                                capture_output=True, text=False)
            
            if result.returncode != 0:
                messagebox.showerror("Грешка", f"Не можах да експортирам таблицата:\n{result.stderr}")
                return False
            
            # Декодиране като UTF-8
            raw_content = result.stdout.decode('utf-8', errors='ignore')
            lines = raw_content.strip().split('\n')
            
            if len(lines) < 1:
                messagebox.showwarning("Внимание", "Таблицата е празна")
                return False
            
            # Поправяме кодировката на всички редове
            fixed_lines = []
            
            for i, line in enumerate(lines):
                if i == 0:
                    # Header остава както е
                    fixed_lines.append(line)
                else:
                    # Поправяме българските текстове
                    fixed_line = self.fix_encoding_utf8_to_windows1251(line)
                    fixed_lines.append(fixed_line)
            
            # Записваме файла
            with open(file_path, 'w', encoding='utf-8', newline='') as f:
                for line in fixed_lines:
                    f.write(line + '\n')
            
            # Статистики
            total_rows = len(fixed_lines) - 1  # Без header
            file_size = os.path.getsize(file_path)
            
            # Броим колоните
            header_reader = csv.reader(io.StringIO(fixed_lines[0]))
            headers = next(header_reader)
            total_columns = len(headers)
            
            messagebox.showinfo("Успех", 
                            f"Пълният експорт е завършен успешно!\n\n"
                            f"📁 Файл: {os.path.basename(file_path)}\n"
                            f"📊 Редове: {total_rows:,}\n"
                            f"📋 Колони: {total_columns}\n"
                            f"💾 Размер: {file_size / 1024 / 1024:.1f} MB\n"
                            f"🔗 Път: {file_path}")
            return True
            
        except Exception as e:
            messagebox.showerror("Грешка", f"Грешка при експорт с mdb-tools:\n{str(e)}")
            return False

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