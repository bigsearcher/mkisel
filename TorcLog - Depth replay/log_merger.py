"""
Парсер для объединения двух логов (Encoder.txt и Time.txt) в единый LAS файл.
Синхронизирует данные по времени с шагом 1 секунда и использует интерполяцию
для значений, которые попадают между тиками в исходных файлах.
"""

import re
import os
import tkinter as tk
from tkinter import messagebox
from datetime import datetime, timedelta
from typing import Dict, List, Tuple, Optional
from dataclasses import dataclass
from collections import defaultdict
import bisect


@dataclass
class EncoderRecord:
    """Запись из Encoder.txt"""
    time: datetime
    depth: float
    speed: float
    weight: float
    pressure: float


@dataclass
class TimeRecord:
    """Запись из Time.txt"""
    time: datetime
    pressure: float
    temperature: float
    head_tension: float
    acceleration_x: float
    acceleration_y: float
    acceleration_z: float
    encoder_depth: float
    speed: float
    surface_tension: float
    ccl_depth: float
    element_depth: float
    # Добавить другие поля по необходимости


@dataclass
class MergedRecord:
    """Объединенная запись с синхронизированными данными"""
    time: datetime
    depth: float  # Из Encoder или Time
    # Данные из Encoder.txt
    encoder_depth: Optional[float] = None
    encoder_speed: Optional[float] = None
    encoder_weight: Optional[float] = None
    encoder_pressure: Optional[float] = None
    # Данные из Time.txt
    time_pressure: Optional[float] = None
    temperature: Optional[float] = None
    head_tension: Optional[float] = None
    acceleration_x: Optional[float] = None
    acceleration_y: Optional[float] = None
    acceleration_z: Optional[float] = None
    time_speed: Optional[float] = None
    surface_tension: Optional[float] = None
    ccl_depth: Optional[float] = None
    element_depth: Optional[float] = None


class LogParser:
    """Парсер для чтения и объединения логов"""
    
    def __init__(self, encoder_files: List[str], time_files: List[str], time_step_seconds: float = 0.1):
        """
        Args:
            encoder_files: список путей к файлам Encoder.txt
            time_files: список путей к файлам Time.txt
            time_step_seconds: шаг времени для синхронизации (по умолчанию 0.1 секунды = 100 мс)
        """
        self.encoder_files = encoder_files if isinstance(encoder_files, list) else [encoder_files]
        self.time_files = time_files if isinstance(time_files, list) else [time_files]
        self.time_step = timedelta(seconds=time_step_seconds)
        self.time_step_seconds = time_step_seconds  # Сохраняем шаг времени в секундах
        self.encoder_records: List[EncoderRecord] = []
        self.time_records: List[TimeRecord] = []
        self.null_value = -999.25  # Стандартное null значение для LAS
        # Частоты дискретизации (средний шаг времени между записями)
        self.encoder_sampling_rate = None  # в секундах
        self.time_sampling_rate = None  # в секундах
        # Максимальное время окончания файлов другого типа (для продолжения записи при решении не сшивать)
        self.max_time_end_time = None  # Максимальное время окончания Time файлов
        self.max_encoder_end_time = None  # Максимальное время окончания Encoder файлов
        # Флаг для отслеживания начала новой группы файлов
        self.new_group_started = False
        # Индексы текущих файлов для каждой группы (начинаем с 0)
        self.current_encoder_file_index = 0
        self.current_time_file_index = 0
        
    def _extract_date_from_file_content(self, filepath: str) -> Optional[datetime]:
        """
        Извлекает дату из содержимого файла.
        Ищет дату в первых строках файла или в заголовках.
        Возвращает datetime или None если не найдено.
        """
        date_patterns = [
            # ISO 8601 формат
            (r'(\d{4}-\d{2}-\d{2})', '%Y-%m-%d'),
            # DD.MM.YYYY
            (r'(\d{2}\.\d{2}\.\d{4})', '%d.%m.%Y'),
            # DD/MM/YYYY
            (r'(\d{2}/\d{2}/\d{4})', '%d/%m/%Y'),
            # YYYY.MM.DD
            (r'(\d{4}\.\d{2}\.\d{2})', '%Y.%m.%d'),
            # DDMMYY в заголовке
            (r'(\d{6})', None),  # Специальная обработка для DDMMYY
        ]
        
        try:
            with open(filepath, 'r', encoding='utf-8') as f:
                # Читаем первые 50 строк для поиска даты
                for line_num, line in enumerate(f, 1):
                    if line_num > 50:  # Ограничиваем поиск первыми 50 строками
                        break
                    
                    line = line.strip()
                    if not line:
                        continue
                    
                    # Ищем дату в строке
                    for pattern, date_format in date_patterns:
                        match = re.search(pattern, line)
                        if match:
                            date_str = match.group(1)
                            
                            # Специальная обработка для DDMMYY
                            if date_format is None and len(date_str) == 6:
                                try:
                                    day = int(date_str[:2])
                                    month = int(date_str[2:4])
                                    year = 2000 + int(date_str[4:6])
                                    # Проверяем валидность даты
                                    parsed_date = datetime(year, month, day)
                                    # Если дата разумная (не слишком старая и не в будущем)
                                    if 2000 <= year <= 2100:
                                        return parsed_date
                                except (ValueError, IndexError):
                                    continue
                            elif date_format:
                                try:
                                    parsed_date = datetime.strptime(date_str, date_format)
                                    # Проверяем валидность даты
                                    if 2000 <= parsed_date.year <= 2100:
                                        return parsed_date
                                except ValueError:
                                    continue
        except Exception as e:
            # Если ошибка при чтении, просто возвращаем None
            pass
        
        return None
    
    def _extract_date_from_filename(self, filepath: str) -> Optional[datetime]:
        """
        Извлекает дату из названия файла.
        Формат: DDMMYY_HHMMSS или DDMMYY_HHMMSS - ...
        Возвращает datetime или None если не найдено.
        """
        filename = os.path.basename(filepath)
        # Ищем паттерн DDMMYY_HHMMSS в начале имени файла
        match = re.match(r'(\d{6})_(\d{6})', filename)
        if match:
            date_str = match.group(1)  # DDMMYY
            time_str = match.group(2)  # HHMMSS
            try:
                # Парсим дату: DDMMYY -> DD-MM-20YY
                day = int(date_str[:2])
                month = int(date_str[2:4])
                year = 2000 + int(date_str[4:6])
                # Парсим время: HHMMSS
                hour = int(time_str[:2])
                minute = int(time_str[2:4])
                second = int(time_str[4:6])
                return datetime(year, month, day, hour, minute, second)
            except (ValueError, IndexError):
                pass
        return None
    
    def _ask_user_merge_decision(self, time_diff_seconds: float, prev_file: str, next_file: str) -> bool:
        """
        Показывает GUI окно для выбора действия при большой разнице времени между файлами.
        Возвращает True если нужно сшивать, False если создавать новый файл.
        """
        root = tk.Tk()
        root.withdraw()  # Скрываем главное окно
        
        # Вычисляем разницу в читаемом формате
        hours = int(time_diff_seconds // 3600)
        minutes = int((time_diff_seconds % 3600) // 60)
        seconds = int(time_diff_seconds % 60)
        
        time_str = ""
        if hours > 0:
            time_str += f"{hours} час(ов) "
        if minutes > 0:
            time_str += f"{minutes} минут(ы) "
        if seconds > 0 or not time_str:
            time_str += f"{seconds} секунд(ы)"
        
        # Создаем диалоговое окно
        dialog = tk.Toplevel(root)
        dialog.title("Разница во времени между файлами")
        dialog.geometry("500x250")
        dialog.transient(root)
        
        # Центрируем окно
        dialog.update_idletasks()
        x = (dialog.winfo_screenwidth() // 2) - (dialog.winfo_width() // 2)
        y = (dialog.winfo_screenheight() // 2) - (dialog.winfo_height() // 2)
        dialog.geometry(f"+{x}+{y}")
        
        result = [False]  # По умолчанию создавать новый файл
        
        def on_merge():
            result[0] = True
            dialog.destroy()
        
        def on_new_file():
            result[0] = False
            dialog.destroy()
        
        # Текст с информацией
        message = f"Обнаружена разница во времени между файлами:\n\n"
        message += f"Предыдущий файл: {os.path.basename(prev_file)}\n"
        message += f"Следующий файл: {os.path.basename(next_file)}\n\n"
        message += f"Разница: {time_str}\n\n"
        message += f"Выберите действие:"
        
        label = tk.Label(dialog, text=message, wraplength=450, justify='left', font=('Arial', 10))
        label.pack(pady=15)
        
        # Кнопки
        button_frame = tk.Frame(dialog)
        button_frame.pack(pady=15)
        
        merge_button = tk.Button(button_frame, text="Сшивать файлы", command=on_merge, width=20, height=2)
        merge_button.pack(side=tk.LEFT, padx=10)
        
        new_file_button = tk.Button(button_frame, text="Создать новый файл", command=on_new_file, width=20, height=2)
        new_file_button.pack(side=tk.LEFT, padx=10)
        
        # Ждем закрытия окна
        dialog.wait_window()
        root.destroy()
        
        return result[0]
    
    def _ask_user_for_date(self, filename: str) -> Optional[datetime]:
        """
        Показывает GUI окно для ввода даты пользователем.
        Возвращает datetime или None если пользователь отменил.
        """
        root = tk.Tk()
        root.withdraw()  # Скрываем главное окно
        
        # Создаем диалоговое окно
        dialog = tk.Toplevel(root)
        dialog.title("Укажите дату для файла")
        dialog.geometry("400x200")
        dialog.transient(root)
        
        # Центрируем окно
        dialog.update_idletasks()
        x = (dialog.winfo_screenwidth() // 2) - (dialog.winfo_width() // 2)
        y = (dialog.winfo_screenheight() // 2) - (dialog.winfo_height() // 2)
        dialog.geometry(f"+{x}+{y}")
        
        result = [None]  # Используем список для передачи значения из замыкания
        
        def on_ok():
            try:
                date_str = date_entry.get().strip()
                if not date_str:
                    messagebox.showerror("Ошибка", "Пожалуйста, введите дату")
                    return
                
                # Пробуем разные форматы даты
                formats = ['%Y-%m-%d', '%d.%m.%Y', '%d/%m/%Y', '%Y.%m.%d']
                parsed_date = None
                for fmt in formats:
                    try:
                        parsed_date = datetime.strptime(date_str, fmt)
                        break
                    except ValueError:
                        continue
                
                if parsed_date is None:
                    messagebox.showerror("Ошибка", "Неверный формат даты. Используйте YYYY-MM-DD, DD.MM.YYYY или DD/MM/YYYY")
                    return
                
                result[0] = parsed_date
                dialog.destroy()
            except Exception as e:
                messagebox.showerror("Ошибка", f"Ошибка при парсинге даты: {e}")
        
        def on_cancel():
            dialog.destroy()
        
        # Текст с инструкцией
        label = tk.Label(dialog, text=f"Файл: {os.path.basename(filename)}\n\nДата не найдена ни в содержимом файла, ни в названии.\nПожалуйста, укажите дату (YYYY-MM-DD, DD.MM.YYYY или DD/MM/YYYY):", 
                        wraplength=350, justify='left')
        label.pack(pady=10)
        
        # Поле ввода даты
        date_entry = tk.Entry(dialog, width=20, font=('Arial', 12))
        date_entry.pack(pady=10)
        date_entry.focus()
        
        # Кнопки
        button_frame = tk.Frame(dialog)
        button_frame.pack(pady=10)
        
        ok_button = tk.Button(button_frame, text="OK", command=on_ok, width=10)
        ok_button.pack(side=tk.LEFT, padx=5)
        
        cancel_button = tk.Button(button_frame, text="Отмена", command=on_cancel, width=10)
        cancel_button.pack(side=tk.LEFT, padx=5)
        
        # Привязываем Enter к OK
        date_entry.bind('<Return>', lambda e: on_ok())
        
        # Ждем закрытия окна
        dialog.wait_window()
        root.destroy()
        
        return result[0]
    
    def _get_file_date(self, filepath: str) -> datetime:
        """
        Определяет дату для файла по приоритету:
        1. Дата из содержимого файла (если есть)
        2. Дата из названия файла (если есть)
        3. Запрос даты у пользователя через GUI (если нет ни в файле, ни в названии)
        """
        # Приоритет 1: Пробуем извлечь дату из содержимого файла
        file_content_date = self._extract_date_from_file_content(filepath)
        if file_content_date:
            print(f"Дата из содержимого файла: {file_content_date.date()}")
            return file_content_date
        
        # Приоритет 2: Пробуем извлечь дату из названия
        filename_date = self._extract_date_from_filename(filepath)
        if filename_date:
            print(f"Дата из названия файла: {filename_date.date()}")
            return filename_date
        
        # Приоритет 3: Если нет ни в файле, ни в названии, запрашиваем у пользователя
        print(f"Дата не найдена ни в содержимом, ни в названии файла: {os.path.basename(filepath)}")
        user_date = self._ask_user_for_date(filepath)
        
        if user_date is None:
            raise ValueError("Дата не указана пользователем. Операция отменена.")
        
        print(f"Дата указана пользователем: {user_date.date()}")
        return user_date
    
    def _determine_sampling_rate_encoder(self, encoder_file: str, file_date, channel_mapping: dict, found_service_line: bool) -> Optional[float]:
        """Определяет частоту дискретизации по первым записям файла"""
        header_pattern = re.compile(r'^##')
        sample_records = []
        sample_size = 100  # Анализируем первые 100 записей
        
        with open(encoder_file, 'r', encoding='utf-8') as f:
            for line in f:
                if len(sample_records) >= sample_size:
                    break
                    
                line = line.strip()
                if not line:
                    continue
                
                if line.startswith('##!^$5') or (header_pattern.match(line) and not found_service_line):
                    continue
                if found_service_line and header_pattern.match(line):
                    continue
                
                parts = line.split()
                if len(parts) < 2:
                    continue
                
                try:
                    time_str = parts[0]
                    time_only = datetime.strptime(time_str, '%H:%M:%S').time()
                    time_obj = datetime.combine(file_date, time_only)
                    sample_records.append(time_obj)
                except (ValueError, IndexError):
                    continue
        
        if len(sample_records) < 2:
            return None
        
        # Вычисляем средний шаг времени
        time_diffs = []
        for i in range(1, len(sample_records)):
            diff = (sample_records[i] - sample_records[i-1]).total_seconds()
            if diff > 0:
                time_diffs.append(diff)
        
        if time_diffs:
            return sum(time_diffs) / len(time_diffs)
        return None
    
    def parse_encoder_file(self, encoder_file: str) -> List[EncoderRecord]:
        """Парсит один файл Encoder.txt с поддержкой служебной строки ##!^$5 и оптимизацией чтения"""
        records = []
        header_pattern = re.compile(r'^##')
        
        # Определяем дату для файла
        file_datetime = self._get_file_date(encoder_file)
        file_date = file_datetime.date()
        
        print(f"Используемая дата для {os.path.basename(encoder_file)}: {file_date}")
        
        # Маппинг каналов: название канала -> индекс в данных (без учета времени)
        # ACQ_TIME_STRING - это время, не данные, поэтому его не включаем в маппинг данных
        channel_mapping = {}
        found_service_line = False
        data_started = False  # Флаг, что начались данные после служебной строки
        
        # Сначала определяем частоту дискретизации (читаем первые строки для анализа)
        # Для этого нужно сначала найти служебную строку
        with open(encoder_file, 'r', encoding='utf-8') as f:
            for line in f:
                line = line.strip()
                if not line:
                    continue
                if line.startswith('##!^$5'):
                    found_service_line = True
                    parts = line.split()
                    channel_names = []
                    for part in parts[1:]:
                        if part.startswith('#') and part != '#NULL':
                            break
                        if part in ['FT', 'F/MN', 'LBF', 'PSI', 'BAR', 'M', 'M/MN', 'KN', 'MG', 'C', '#NULL']:
                            break
                        channel_names.append(part)
                    for idx, channel_name in enumerate(channel_names):
                        if channel_name != 'ACQ_TIME_STRING':
                            channel_mapping[channel_name] = idx
                    break
        
        # Определяем частоту дискретизации
        sampling_rate = self._determine_sampling_rate_encoder(encoder_file, file_date, channel_mapping, found_service_line)
        
        # Вычисляем шаг для пропуска строк (если данные чаще, чем целевой шаг)
        skip_lines = 1
        if sampling_rate and sampling_rate < self.time_step_seconds:
            # Если данные каждые 0.05 сек, а нужен шаг 0.1 сек, читаем каждую 2-ю строку
            skip_lines = max(1, int(self.time_step_seconds / sampling_rate))
            print(f"Оптимизация: частота {sampling_rate:.3f} сек < целевой шаг {self.time_step_seconds:.3f} сек, пропуск {skip_lines-1} строк из {skip_lines}")
        
        # Теперь читаем файл с оптимизацией
        with open(encoder_file, 'r', encoding='utf-8') as f:
            line_count = 0
            last_time = None
            check_interval = 1000  # Проверяем время каждые 1000 строк
            
            for line in f:
                line = line.strip()
                if not line:
                    continue
                
                # Ищем служебную строку ##!^$5
                if line.startswith('##!^$5'):
                    if not found_service_line:
                        found_service_line = True
                        data_started = False
                        parts = line.split()
                        channel_names = []
                        for part in parts[1:]:
                            if part.startswith('#') and part != '#NULL':
                                break
                            if part in ['FT', 'F/MN', 'LBF', 'PSI', 'BAR', 'M', 'M/MN', 'KN', 'MG', 'C', '#NULL']:
                                break
                            channel_names.append(part)
                        for idx, channel_name in enumerate(channel_names):
                            if channel_name != 'ACQ_TIME_STRING':
                                channel_mapping[channel_name] = idx
                        print(f"Найдена служебная строка. Каналы: {channel_names}")
                        print(f"Маппинг каналов данных: {channel_mapping}")
                    continue
                
                # Пропускаем другие служебные строки
                if header_pattern.match(line) and not found_service_line:
                    continue
                if found_service_line and header_pattern.match(line):
                    continue
                
                # Оптимизация: пропускаем строки, если данные чаще целевого шага
                if skip_lines > 1:
                    line_count += 1
                    if line_count % skip_lines != 0:
                        continue
                
                # Парсим данные
                parts = line.split()
                if len(parts) < 2:
                    continue
                
                try:
                    # Первый элемент всегда время
                    time_str = parts[0]
                    
                    # Если есть маппинг каналов, используем его
                    if channel_mapping:
                        depth_idx = channel_mapping.get('CT_CORR_DEPTH', 1)
                        speed_idx = channel_mapping.get('CT_SPEED', 2)
                        weight_idx = channel_mapping.get('CT_WEIGHT', 3)
                        pressure_idx = channel_mapping.get('WH_PRESS', 4)
                        
                        max_idx = max(depth_idx, speed_idx, weight_idx, pressure_idx)
                        if len(parts) <= max_idx:
                            continue
                        
                        depth = float(parts[depth_idx]) if depth_idx < len(parts) else 0.0
                        speed = float(parts[speed_idx]) if speed_idx < len(parts) else 0.0
                        weight = float(parts[weight_idx]) if weight_idx < len(parts) else 0.0
                        pressure = float(parts[pressure_idx]) if pressure_idx < len(parts) else 0.0
                    else:
                        if len(parts) < 5:
                            continue
                        depth = float(parts[1])
                        speed = float(parts[2])
                        weight = float(parts[3])
                        pressure = float(parts[4])
                    
                    # Парсим время
                    time_only = datetime.strptime(time_str, '%H:%M:%S').time()
                    time_obj = datetime.combine(file_date, time_only)
                    
                    # Периодически проверяем время и корректируем шаг (если пропускаем строки)
                    if skip_lines > 1 and last_time and len(records) % check_interval == 0:
                        expected_time = last_time + timedelta(seconds=sampling_rate * skip_lines)
                        time_error = (time_obj - expected_time).total_seconds()  # Может быть положительным или отрицательным
                        time_diff = abs(time_error)
                        
                        if time_diff > sampling_rate * 1.5:  # Если разница больше 1.5 шагов, корректируем
                            old_skip = skip_lines
                            
                            if time_error < 0:
                                # Время раньше ожидаемого - мы пропустили слишком много строк, уменьшаем шаг
                                skip_lines = max(1, skip_lines - 1)
                                print(f"Корректировка: время опережает на {time_diff:.3f} сек, уменьшаем шаг: {old_skip} -> {skip_lines}")
                            else:
                                # Время позже ожидаемого - мы пропустили слишком мало строк, можно увеличить шаг
                                max_skip = int(self.time_step_seconds / sampling_rate) if sampling_rate else skip_lines
                                if skip_lines < max_skip:
                                    skip_lines = min(max_skip, skip_lines + 1)
                                    print(f"Корректировка: время отстает на {time_diff:.3f} сек, увеличиваем шаг: {old_skip} -> {skip_lines}")
                    
                    records.append(EncoderRecord(
                        time=time_obj,
                        depth=depth,
                        speed=speed,
                        weight=weight,
                        pressure=pressure
                    ))
                    last_time = time_obj
                    data_started = True
                except (ValueError, IndexError) as e:
                    if not line.startswith('##'):
                        if len(records) < 5:
                            print(f"Ошибка парсинга строки Encoder: {line[:50]}, {e}")
                    continue
        
        print(f"Загружено {len(records)} записей из {os.path.basename(encoder_file)}")
        if records:
            print(f"Диапазон времени: {records[0].time} - {records[-1].time}")
        return records
    
    def _encoder_file_generator(self):
        """
        Генератор, который парсит Encoder файлы по мере необходимости.
        При окончании файла проверяет разрыв с следующим и принимает решение.
        Если разрыв < 10 сек или пользователь выбрал сшивать - продолжает выдавать последние значения до начала нового файла.
        Если пользователь выбрал не сшивать - продолжает выдавать последние значения до окончания Time файлов.
        Начинает с current_encoder_file_index для поддержки обработки групп.
        Yields: EncoderRecord
        """
        if not self.encoder_files:
            return
        
        # Определяем максимальное время окончания Time файлов (если они есть)
        if self.max_time_end_time is None and self.time_files:
            # Быстро определяем максимальное время окончания Time файлов
            max_time_end = None
            for time_file in self.time_files:
                try:
                    records = self.parse_time_file(time_file)
                    if records:
                        file_end = records[-1].time
                        if max_time_end is None or file_end > max_time_end:
                            max_time_end = file_end
                except:
                    pass
            self.max_time_end_time = max_time_end
        
        # Начинаем с текущего индекса файла (для поддержки групп)
        current_file_index = self.current_encoder_file_index
        current_records = None
        current_record_index = 0
        last_record = None  # Последняя запись из предыдущего файла
        
        while current_file_index < len(self.encoder_files):
            encoder_file = self.encoder_files[current_file_index]
            
            # Если это новый файл, парсим его
            if current_records is None:
                print(f"\nПарсинг Encoder файла {current_file_index+1}/{len(self.encoder_files)}: {os.path.basename(encoder_file)}")
                current_records = self.parse_encoder_file(encoder_file)
                
                if not current_records:
                    print(f"Пропуск пустого файла: {encoder_file}")
                    current_file_index += 1
                    continue
                
                file_start_time = current_records[0].time
                
                # Проверяем разрыв с предыдущим файлом
                if last_record is not None:
                    last_end_time = last_record.time
                    time_diff = (file_start_time - last_end_time).total_seconds()
                    
                    should_continue_last_values = False
                    continue_until_time = None  # До какого времени продолжать
                    should_merge = True  # По умолчанию сшиваем
                    
                    if abs(time_diff) < 10:  # Разница меньше 10 секунд - автоматически продолжаем
                        should_continue_last_values = True
                        continue_until_time = file_start_time
                        should_merge = True
                        print(f"Разрыв {time_diff:.1f} сек < 10 сек, продолжаем последние значения до начала нового файла")
                    elif abs(time_diff) >= 10:  # Разница >= 10 секунд - создаем новую группу
                        print(f"\nВНИМАНИЕ: Разница во времени между Encoder файлами {time_diff:.1f} секунд")
                        print(f"  Предыдущий: {os.path.basename(self.encoder_files[current_file_index-1])}")
                        print(f"  Текущий: {os.path.basename(encoder_file)}")
                        print(f"  Автоматически создаем новую группу (разрыв >= 10 сек)")

                        # Не сшиваем - создаем новую группу
                        should_merge = False
                        should_continue_last_values = True
                        if self.max_time_end_time:
                            continue_until_time = self.max_time_end_time
                            print(f"Остановка сшивания: продолжаем последние значения до окончания Time файлов ({self.max_time_end_time})")
                        else:
                            # Если нет Time файлов, продолжаем до начала нового Encoder файла
                            continue_until_time = file_start_time
                            print("Остановка сшивания: продолжаем последние значения до начала нового файла (нет Time файлов)")
                    
                    # Если нужно продолжать последние значения, выдаем их до указанного времени
                    if should_continue_last_values and continue_until_time:
                        current_time = last_end_time + self.time_step
                        while current_time < continue_until_time:
                            # Создаем запись с последними значениями, но новым временем
                            continued_record = EncoderRecord(
                                time=current_time,
                                depth=last_record.depth,
                                speed=last_record.speed,
                                weight=last_record.weight,
                                pressure=last_record.pressure
                            )
                            yield continued_record
                            current_time += self.time_step
                        
                        # Если не сшиваем, устанавливаем флаг начала новой группы
                        if not should_merge and abs(time_diff) >= 10:
                            self.new_group_started = True
                            last_record = None
                
                current_record_index = 0
            
            # Выдаем записи из текущего файла
            while current_record_index < len(current_records):
                record = current_records[current_record_index]
                yield record
                last_record = record  # Сохраняем последнюю запись
                current_record_index += 1
            
            # Файл закончился, сохраняем последнюю запись
            if current_records:
                last_record = current_records[-1]
            
            # Обновляем индекс для следующей группы
            self.current_encoder_file_index = current_file_index + 1
            
            # Останавливаем генератор - следующий файл будет прочитан только после принятия решения
            # о сшивании в следующей итерации группы (если она будет)
            print(f"Encoder файл {current_file_index+1} закончился, останавливаем генератор")
            break
    
    def parse_all_encoder_files(self) -> List[EncoderRecord]:
        """Парсит все файлы Encoder.txt для обратной совместимости"""
        all_records = []
        for record in self._encoder_file_generator():
            all_records.append(record)
        
        self.encoder_records = all_records
        print(f"\nВсего загружено {len(all_records)} записей из всех Encoder файлов")
        if all_records:
            print(f"Общий диапазон времени: {all_records[0].time} - {all_records[-1].time}")
            # Вычисляем частоту дискретизации (средний шаг времени)
            if len(all_records) > 1:
                time_diffs = []
                for i in range(1, min(100, len(all_records))):  # Анализируем первые 100 записей
                    diff = (all_records[i].time - all_records[i-1].time).total_seconds()
                    if diff > 0:  # Игнорируем нулевые или отрицательные разницы
                        time_diffs.append(diff)
                if time_diffs:
                    self.encoder_sampling_rate = sum(time_diffs) / len(time_diffs)
                    print(f"Частота дискретизации Encoder: {self.encoder_sampling_rate:.3f} сек")
        return all_records
    
    def _determine_sampling_rate_time(self, time_file: str) -> Optional[float]:
        """Определяет частоту дискретизации по первым записям файла Time.txt"""
        sample_records = []
        sample_size = 100
        header_skipped = False
        
        with open(time_file, 'r', encoding='utf-8') as f:
            for line in f:
                if len(sample_records) >= sample_size:
                    break
                    
                line = line.strip()
                if not line:
                    continue
                
                if line.startswith('TorcLog'):
                    continue
                
                if not header_skipped:
                    if 'DateTime' in line:
                        header_skipped = True
                    continue
                
                parts = line.split('\t')
                if len(parts) < 15:
                    continue
                
                try:
                    time_str = parts[0]
                    if '+03:00' in time_str:
                        time_str = time_str.replace('+03:00', '')
                    time_obj = datetime.fromisoformat(time_str)
                    sample_records.append(time_obj)
                except (ValueError, IndexError):
                    continue
        
        if len(sample_records) < 2:
            return None
        
        time_diffs = []
        for i in range(1, len(sample_records)):
            diff = (sample_records[i] - sample_records[i-1]).total_seconds()
            if diff > 0:
                time_diffs.append(diff)
        
        if time_diffs:
            return sum(time_diffs) / len(time_diffs)
        return None
    
    def parse_time_file(self, time_file: str, max_records: Optional[int] = None) -> List[TimeRecord]:
        """Парсит один файл Time.txt с оптимизацией чтения"""
        records = []
        header_skipped = False
        
        # Определяем частоту дискретизации
        sampling_rate = self._determine_sampling_rate_time(time_file)
        
        # Вычисляем шаг для пропуска строк
        skip_lines = 1
        if sampling_rate and sampling_rate < self.time_step_seconds:
            skip_lines = max(1, int(self.time_step_seconds / sampling_rate))
            print(f"Оптимизация: частота {sampling_rate:.3f} сек < целевой шаг {self.time_step_seconds:.3f} сек, пропуск {skip_lines-1} строк из {skip_lines}")
        
        def safe_float(value, default=0.0):
            """Безопасное преобразование в float"""
            try:
                return float(value) if value else default
            except (ValueError, TypeError):
                return default
        
        with open(time_file, 'r', encoding='utf-8') as f:
            line_count = 0
            last_time = None
            check_interval = 1000
            
            for line_num, line in enumerate(f, 1):
                if max_records and len(records) >= max_records:
                    break
                    
                line = line.strip()
                if not line:
                    continue
                
                # Пропускаем заголовок версии
                if line.startswith('TorcLog'):
                    continue
                
                # Пропускаем строку заголовков
                if not header_skipped:
                    if 'DateTime' in line:
                        header_skipped = True
                    continue
                
                # Оптимизация: пропускаем строки, если данные чаще целевого шага
                if skip_lines > 1:
                    line_count += 1
                    if line_count % skip_lines != 0:
                        continue
                
                # Разделяем по табуляции
                parts = line.split('\t')
                if len(parts) < 15:
                    continue
                
                try:
                    # Парсим ISO 8601 время
                    time_str = parts[0]
                    # Убираем временную зону для упрощения
                    if '+03:00' in time_str:
                        time_str = time_str.replace('+03:00', '')
                    time_obj = datetime.fromisoformat(time_str)
                    
                    # Периодически проверяем время и корректируем шаг (если пропускаем строки)
                    if skip_lines > 1 and last_time and len(records) % check_interval == 0:
                        expected_time = last_time + timedelta(seconds=sampling_rate * skip_lines)
                        time_error = (time_obj - expected_time).total_seconds()  # Может быть положительным или отрицательным
                        time_diff = abs(time_error)
                        
                        if time_diff > sampling_rate * 1.5:  # Если разница больше 1.5 шагов, корректируем
                            old_skip = skip_lines
                            
                            if time_error < 0:
                                # Время раньше ожидаемого - мы пропустили слишком много строк, уменьшаем шаг
                                skip_lines = max(1, skip_lines - 1)
                                print(f"Корректировка: время опережает на {time_diff:.3f} сек, уменьшаем шаг: {old_skip} -> {skip_lines}")
                            else:
                                # Время позже ожидаемого - мы пропустили слишком мало строк, можно увеличить шаг
                                max_skip = int(self.time_step_seconds / sampling_rate) if sampling_rate else skip_lines
                                if skip_lines < max_skip:
                                    skip_lines = min(max_skip, skip_lines + 1)
                                    print(f"Корректировка: время отстает на {time_diff:.3f} сек, увеличиваем шаг: {old_skip} -> {skip_lines}")
                    
                    # Парсим основные поля
                    pressure = safe_float(parts[2])
                    temperature = safe_float(parts[3])
                    head_tension = safe_float(parts[4])
                    acceleration_x = safe_float(parts[5])
                    acceleration_y = safe_float(parts[6])
                    acceleration_z = safe_float(parts[7])
                    encoder_depth = safe_float(parts[10])
                    speed = safe_float(parts[11])
                    surface_tension = safe_float(parts[12])
                    ccl_depth = safe_float(parts[13])
                    element_depth = safe_float(parts[14])
                    
                    records.append(TimeRecord(
                        time=time_obj,
                        pressure=pressure,
                        temperature=temperature,
                        head_tension=head_tension,
                        acceleration_x=acceleration_x,
                        acceleration_y=acceleration_y,
                        acceleration_z=acceleration_z,
                        encoder_depth=encoder_depth,
                        speed=speed,
                        surface_tension=surface_tension,
                        ccl_depth=ccl_depth,
                        element_depth=element_depth
                    ))
                    last_time = time_obj
                except (ValueError, IndexError) as e:
                    if line_num < 10:
                        print(f"Ошибка парсинга строки Time (строка {line_num}): {line[:100]}, {e}")
                    continue
        
        print(f"Загружено {len(records)} записей из {os.path.basename(time_file)}")
        if records:
            print(f"Диапазон времени: {records[0].time} - {records[-1].time}")
        return records
    
    def _time_file_generator(self, max_records: Optional[int] = None):
        """
        Генератор, который парсит Time файлы по мере необходимости.
        При окончании файла проверяет разрыв с следующим и принимает решение.
        Если разрыв < 10 сек или пользователь выбрал сшивать - продолжает выдавать последние значения до начала нового файла.
        Если пользователь выбрал не сшивать - продолжает выдавать последние значения до окончания Encoder файлов.
        Начинает с current_time_file_index для поддержки обработки групп.
        Yields: TimeRecord
        """
        if not self.time_files:
            return
        
        # Определяем максимальное время окончания Encoder файлов (если они есть)
        if self.max_encoder_end_time is None and self.encoder_files:
            # Быстро определяем максимальное время окончания Encoder файлов
            max_encoder_end = None
            for encoder_file in self.encoder_files:
                try:
                    records = self.parse_encoder_file(encoder_file)
                    if records:
                        file_end = records[-1].time
                        if max_encoder_end is None or file_end > max_encoder_end:
                            max_encoder_end = file_end
                except:
                    pass
            self.max_encoder_end_time = max_encoder_end
        
        # Начинаем с текущего индекса файла (для поддержки групп)
        current_file_index = self.current_time_file_index
        current_records = None
        current_record_index = 0
        last_record = None  # Последняя запись из предыдущего файла
        total_records = 0
        
        while current_file_index < len(self.time_files):
            time_file = self.time_files[current_file_index]
            
            # Если это новый файл, парсим его
            if current_records is None:
                print(f"\nПарсинг Time файла {current_file_index+1}/{len(self.time_files)}: {os.path.basename(time_file)}")
                # Ограничиваем max_records для текущего файла
                remaining_records = max_records - total_records if max_records else None
                current_records = self.parse_time_file(time_file, remaining_records)
                
                if not current_records:
                    print(f"Пропуск пустого файла: {time_file}")
                    current_file_index += 1
                    continue
                
                file_start_time = current_records[0].time
                
                # Проверяем разрыв с предыдущим файлом
                if last_record is not None:
                    last_end_time = last_record.time
                    time_diff = (file_start_time - last_end_time).total_seconds()
                    
                    should_continue_last_values = False
                    continue_until_time = None  # До какого времени продолжать
                    should_merge = True  # По умолчанию сшиваем
                    
                    if abs(time_diff) < 10:  # Разница меньше 10 секунд - автоматически продолжаем
                        should_continue_last_values = True
                        continue_until_time = file_start_time
                        should_merge = True
                        print(f"Разрыв {time_diff:.1f} сек < 10 сек, продолжаем последние значения до начала нового файла")
                    elif abs(time_diff) >= 10:  # Разница >= 10 секунд - создаем новую группу
                        print(f"\nВНИМАНИЕ: Разница во времени между Time файлами {time_diff:.1f} секунд")
                        print(f"  Предыдущий: {os.path.basename(self.time_files[current_file_index-1])}")
                        print(f"  Текущий: {os.path.basename(time_file)}")
                        print(f"  Автоматически создаем новую группу (разрыв >= 10 сек)")

                        # Не сшиваем - создаем новую группу
                        should_merge = False
                        should_continue_last_values = True
                        if self.max_encoder_end_time:
                            continue_until_time = self.max_encoder_end_time
                            print(f"Остановка сшивания: продолжаем последние значения до окончания Encoder файлов ({self.max_encoder_end_time})")
                        else:
                            # Если нет Encoder файлов, продолжаем до начала нового Time файла
                            continue_until_time = file_start_time
                            print("Остановка сшивания: продолжаем последние значения до начала нового файла (нет Encoder файлов)")
                    
                    # Если нужно продолжать последние значения, выдаем их до указанного времени
                    if should_continue_last_values and continue_until_time:
                        current_time = last_end_time + self.time_step
                        while current_time < continue_until_time:
                            if max_records and total_records >= max_records:
                                return
                            # Создаем запись с последними значениями, но новым временем
                            continued_record = TimeRecord(
                                time=current_time,
                                pressure=last_record.pressure,
                                temperature=last_record.temperature,
                                head_tension=last_record.head_tension,
                                acceleration_x=last_record.acceleration_x,
                                acceleration_y=last_record.acceleration_y,
                                acceleration_z=last_record.acceleration_z,
                                encoder_depth=last_record.encoder_depth,
                                speed=last_record.speed,
                                surface_tension=last_record.surface_tension,
                                ccl_depth=last_record.ccl_depth,
                                element_depth=last_record.element_depth
                            )
                            yield continued_record
                            total_records += 1
                            current_time += self.time_step
                        
                        # Если не сшиваем, устанавливаем флаг начала новой группы
                        if not should_merge and abs(time_diff) >= 10:
                            self.new_group_started = True
                            last_record = None
            
            # Выдаем записи из текущего файла
            while current_record_index < len(current_records):
                if max_records and total_records >= max_records:
                    return
                
                record = current_records[current_record_index]
                yield record
                last_record = record  # Сохраняем последнюю запись
                current_record_index += 1
                total_records += 1
            
            # Файл закончился, сохраняем последнюю запись
            if current_records:
                last_record = current_records[-1]
            
            # Обновляем индекс для следующей группы
            self.current_time_file_index = current_file_index + 1
            
            # Останавливаем генератор - следующий файл будет прочитан только после принятия решения
            # о сшивании в следующей итерации группы (если она будет)
            print(f"Time файл {current_file_index+1} закончился, останавливаем генератор")
            break
    
    def parse_all_time_files(self, max_records: Optional[int] = None) -> List[TimeRecord]:
        """Парсит все файлы Time.txt для обратной совместимости"""
        all_records = []
        for record in self._time_file_generator(max_records):
            all_records.append(record)
        
        self.time_records = all_records
        print(f"\nВсего загружено {len(all_records)} записей из всех Time файлов")
        if all_records:
            print(f"Общий диапазон времени: {all_records[0].time} - {all_records[-1].time}")
            # Вычисляем частоту дискретизации (средний шаг времени)
            if len(all_records) > 1:
                time_diffs = []
                for i in range(1, min(100, len(all_records))):  # Анализируем первые 100 записей
                    diff = (all_records[i].time - all_records[i-1].time).total_seconds()
                    if diff > 0:  # Игнорируем нулевые или отрицательные разницы
                        time_diffs.append(diff)
                if time_diffs:
                    self.time_sampling_rate = sum(time_diffs) / len(time_diffs)
                    print(f"Частота дискретизации Time: {self.time_sampling_rate:.3f} сек")
        return all_records
    
    def normalize_dates(self):
        """
        Нормализует даты в обоих файлах для синхронизации.
        Encoder.txt уже имеет дату из времени изменения файла.
        Проверяем и корректируем при необходимости.
        """
        # Если нет одного из типов файлов, нормализация не требуется
        if not self.encoder_records or not self.time_records:
            return
        
        # Берем дату и время из первого файла Time.txt
        base_datetime = self.time_records[0].time
        base_date = base_datetime.date()
        
        # Берем дату из первого файла Encoder.txt (уже установлена из времени файла)
        encoder_date = self.encoder_records[0].time.date()
        encoder_time_only = self.encoder_records[0].time.time()
        
        print(f"\nНормализация дат:")
        print(f"  Time.txt дата: {base_date}, первое время: {base_datetime}")
        print(f"  Encoder.txt дата: {encoder_date}, первое время: {self.encoder_records[0].time}")
        
        # Если даты разные, нужно синхронизировать
        if encoder_date != base_date:
            # Вычисляем разницу в днях
            days_diff = (base_date - encoder_date).days
            
            # Если разница больше 1 дня, возможно файлы из разных периодов
            if abs(days_diff) > 1:
                print(f"  ВНИМАНИЕ: Разница в датах {days_diff} дней!")
            
            # Создаем datetime для Encoder с датой из Time.txt
            encoder_base_datetime = datetime.combine(base_date, encoder_time_only)
            
            # Вычисляем разницу между первым временем Encoder и первым временем Time
            time_diff = base_datetime - encoder_base_datetime
            
            # Если разница слишком большая (больше 12 часов), возможно время в Encoder
            # относится к следующему дню или предыдущему
            if abs(time_diff.total_seconds()) > 12 * 3600:
                print(f"  Разница во времени: {time_diff.total_seconds()/3600:.1f} часов")
                # Пробуем следующий день
                encoder_base_datetime_next = datetime.combine(
                    base_date + timedelta(days=1), 
                    encoder_time_only
                )
                time_diff_next = base_datetime - encoder_base_datetime_next
                
                # Пробуем предыдущий день
                encoder_base_datetime_prev = datetime.combine(
                    base_date - timedelta(days=1), 
                    encoder_time_only
                )
                time_diff_prev = base_datetime - encoder_base_datetime_prev
                
                # Выбираем минимальную разницу
                if abs(time_diff_next.total_seconds()) < abs(time_diff.total_seconds()):
                    time_diff = time_diff_next
                    base_date = base_date + timedelta(days=1)
                    print(f"  Используем следующий день")
                elif abs(time_diff_prev.total_seconds()) < abs(time_diff.total_seconds()):
                    time_diff = time_diff_prev
                    base_date = base_date - timedelta(days=1)
                    print(f"  Используем предыдущий день")
            
            # Применяем разницу ко всем записям Encoder
            for record in self.encoder_records:
                # Обновляем дату записи
                new_datetime = datetime.combine(base_date, record.time.time()) + time_diff
                record.time = new_datetime
            
            print(f"  Синхронизировано: Encoder.txt теперь использует дату {base_date}")
        else:
            print(f"  Даты совпадают, синхронизация не требуется")
    
    def interpolate_value(self, target_time: datetime,
                         records: List,
                         time_field: str,
                         value_field: str,
                         sampling_rate: Optional[float] = None,
                         times_cache: Optional[List[datetime]] = None) -> Optional[float]:
        """
        Интерполирует значение для заданного времени.
        Если частота дискретизации источника < частоты итоговой сетки (данные чаще),
        то в каждый момент итоговой сетки есть точное значение - берем ближайшее (не интерполируем).
        Если частота дискретизации источника >= частоты итоговой сетки (данные реже),
        то нужно интерполировать между предыдущей и следующей записью.

        Args:
            target_time: целевое время
            records: список записей (EncoderRecord или TimeRecord)
            time_field: имя поля времени ('time')
            value_field: имя поля значения
            sampling_rate: частота дискретизации источника в секундах (если None, вычисляется автоматически)
            times_cache: предварительно созданный список времен (для оптимизации)
        """
        if not records:
            return None

        # Определяем частоту дискретизации, если не указана
        if sampling_rate is None:
            # Пытаемся определить по первым записям
            if len(records) > 1:
                time_diffs = []
                for i in range(1, min(10, len(records))):
                    diff = (getattr(records[i], time_field) - getattr(records[i-1], time_field)).total_seconds()
                    if diff > 0:
                        time_diffs.append(diff)
                if time_diffs:
                    sampling_rate = sum(time_diffs) / len(time_diffs)

        # Используем бинарный поиск для нахождения ближайших записей
        # Используем кэш времен, если передан, иначе создаем
        if times_cache is None:
            times = [getattr(record, time_field) for record in records]
        else:
            times = times_cache

        # Находим индекс элемента, который должен быть вставлен справа от target_time
        idx = bisect.bisect_right(times, target_time)

        before = None
        after = None

        # before - последний элемент <= target_time (idx - 1)
        if idx > 0:
            before = records[idx - 1]

        # after - первый элемент > target_time (idx)
        if idx < len(records):
            after = records[idx]
        
        # Если есть точное совпадение
        if before and getattr(before, time_field) == target_time:
            return getattr(before, value_field, None)
        
        # Если нет данных до или после
        if not before:
            if after:
                return getattr(after, value_field, None)
            return None
        
        if not after:
            return getattr(before, value_field, None)
        
        before_time = getattr(before, time_field)
        after_time = getattr(after, time_field)
        before_val = getattr(before, value_field, None)
        after_val = getattr(after, value_field, None)
        
        if before_val is None or after_val is None:
            return before_val or after_val
        
        # Если шаг времени >= 1 секунды, всегда берем ближайшее значение без интерполяции
        # Если частота дискретизации источника < частоты итоговой сетки (данные чаще), 
        # то в каждый момент итоговой сетки есть точное значение - берем ближайшее (не интерполируем)
        if self.time_step_seconds >= 1.0 or (sampling_rate is not None and sampling_rate < self.time_step_seconds):
            # Берем ближайшее значение (before или after)
            time_diff_before = abs((target_time - before_time).total_seconds())
            time_diff_after = abs((target_time - after_time).total_seconds())
            
            if time_diff_before <= time_diff_after:
                return before_val
            else:
                return after_val
        
        # Если частота дискретизации источника >= частоты итоговой сетки (данные реже, например 1.0 сек >= 0.1 сек),
        # то нужно интерполировать между предыдущей и следующей записью
        # Линейная интерполяция
        if before_time == after_time:
            return before_val
        
        time_diff = (target_time - before_time).total_seconds()
        total_diff = (after_time - before_time).total_seconds()
        
        if total_diff == 0:
            return before_val
        
        weight = time_diff / total_diff
        interpolated = before_val + (after_val - before_val) * weight
        
        return interpolated
    
    def merge_logs(self, max_records: Optional[int] = None, get_output_filename_callback=None):
        """
        Объединяет логи с синхронизацией по времени.
        Работает с генераторами файлов, подгружая их по мере необходимости.
        При окончании одного файла дописывает null значения до окончания другого.
        При начале новой группы (когда пользователь выбрал не сшивать) обрабатывает группы по отдельности.
        
        Args:
            max_records: максимальное количество записей для обработки
            get_output_filename_callback: функция, которая вызывается для получения имени файла для каждой группы.
                Принимает номер группы (начиная с 0) и возвращает путь к файлу или None для пропуска группы.
                Если None, используется один файл для всех групп.
        
        Yields: генераторы записей для каждой группы (или один генератор, если callback не указан)
        """
        if not self.encoder_files and not self.time_files:
            raise ValueError("Нужно загрузить хотя бы один файл (Encoder или Time)")
        
        # Сбрасываем флаг новой группы и индексы файлов
        self.new_group_started = False
        self.current_encoder_file_index = 0
        self.current_time_file_index = 0
        
        # Если callback не указан, обрабатываем все группы в один файл
        if get_output_filename_callback is None:
            # Стандартная обработка - все в один файл
            yield from self._merge_logs_single_group(max_records)
            return
        
        # Обрабатываем группы по отдельности
        group_index = 0
        while True:
            print(f"\n=== Начало обработки группы {group_index} ===")
            print(f"Текущие индексы: encoder={self.current_encoder_file_index}/{len(self.encoder_files)}, time={self.current_time_file_index}/{len(self.time_files)}")
            
            # Сбрасываем флаг новой группы перед обработкой группы
            self.new_group_started = False
            
            # Проверяем, есть ли еще файлы для обработки
            if self.current_encoder_file_index >= len(self.encoder_files) and self.current_time_file_index >= len(self.time_files):
                # Все файлы обработаны
                print("Все файлы обработаны, выходим из цикла")
                break
            
            # Получаем имя файла для текущей группы
            print(f"\n>>> ВЫЗЫВАЕМ CALLBACK ДЛЯ ГРУППЫ {group_index} <<<")
            try:
                output_filename = get_output_filename_callback(group_index)
                print(f">>> CALLBACK ВЕРНУЛ: {output_filename} <<<")
            except Exception as e:
                print(f">>> ОШИБКА В CALLBACK: {e} <<<")
                import traceback
                traceback.print_exc()
                output_filename = None
            if output_filename is None:
                # Пропускаем эту группу - пользователь отменил
                print("Пользователь отменил, выходим из цикла")
                break
            
            # Обрабатываем текущую группу
            # Создаем wrapper-генератор, который будет следить за состоянием
            print(f"Создаем генератор для группы {group_index}...")

            def group_generator_wrapper():
                """Wrapper, который потребляет генератор и yield-ит записи"""
                gen = self._merge_logs_single_group(max_records)
                try:
                    for record in gen:
                        yield record
                except StopIteration:
                    pass
                # После завершения генератора флаг new_group_started уже установлен (или нет)
                print(f"Генератор группы {group_index} завершен, new_group_started={self.new_group_started}")

            # Передаем wrapper-генератор
            # LAS генератор начнет запись заголовка сразу, потом данные чанками
            print(f"Передаем генератор для группы {group_index} (потоковая обработка)")
            yield (group_index, output_filename, group_generator_wrapper())

            # ВАЖНО: После yield управление вернется только когда генератор будет полностью обработан
            # (когда LAS генератор прочитает все записи)
            # К этому моменту флаг new_group_started уже будет установлен (или нет)

            print(f"Генератор группы {group_index} обработан")
            print(f"Состояние: new_group_started={self.new_group_started}, encoder_idx={self.current_encoder_file_index}/{len(self.encoder_files)}, time_idx={self.current_time_file_index}/{len(self.time_files)}")

            # Если установлен флаг новой группы - продолжаем цикл для следующей группы
            if self.new_group_started:
                print(f"Начинается группа {group_index + 1}")
                group_index += 1
                continue

            # Если флаг не установлен - все файлы обработаны
            print("Все файлы обработаны, завершаем")
            break
    
    def _merge_logs_single_group(self, max_records: Optional[int] = None):
        """
        Объединяет логи одной группы с синхронизацией по времени.
        Возвращает генератор записей для потоковой обработки.
        Останавливается, когда начинается новая группа (new_group_started = True).
        """
        # Создаем генераторы для файлов
        encoder_gen = self._encoder_file_generator()
        time_gen = self._time_file_generator()
        
        # Получаем первые записи для определения start_time
        encoder_records_buffer = []  # Буфер для Encoder записей
        time_records_buffer = []  # Буфер для Time записей
        
        # Загружаем первые записи из каждого генератора
        try:
            encoder_record = next(encoder_gen)
            encoder_records_buffer.append(encoder_record)
        except StopIteration:
            encoder_gen = None
        
        try:
            time_record = next(time_gen)
            time_records_buffer.append(time_record)
        except StopIteration:
            time_gen = None
        
        if not encoder_records_buffer and not time_records_buffer:
            return
        
        # Определяем start_time (приоритет Encoder)
        if encoder_records_buffer:
            start_time = encoder_records_buffer[0].time
            if time_records_buffer:
                start_time = min(start_time, time_records_buffer[0].time)
        else:
            start_time = time_records_buffer[0].time
        
        # Округляем до начала секунды
        start_time = start_time.replace(microsecond=0)
        
        # Нормализуем даты (нужно сделать это на основе первых записей)
        # Для упрощения, используем существующую логику normalize_dates
        # Но сначала нужно загрузить все записи в буферы для нормализации
        # Временно загружаем все записи для нормализации
        all_encoder_records = []
        all_time_records = []
        
        # Загружаем все Encoder записи из текущего файла (генератор останавливается при окончании файла)
        encoder_file_ended = False
        if encoder_records_buffer:
            all_encoder_records.extend(encoder_records_buffer)
            try:
                while True:
                    record = next(encoder_gen)
                    all_encoder_records.append(record)
                    # Проверяем флаг после каждой записи (генератор может установить его)
                    if self.new_group_started:
                        print("Обнаружен флаг new_group_started во время загрузки Encoder записей")
                        break
            except StopIteration:
                # Генератор остановился (файл закончился)
                encoder_file_ended = True
                print("Encoder генератор остановился (файл закончился)")
                pass
        
        # Сохраняем флаг перед загрузкой Time записей
        encoder_new_group = self.new_group_started
        
        # Загружаем все Time записи из текущего файла (генератор останавливается при окончании файла)
        time_file_ended = False
        if time_records_buffer:
            all_time_records.extend(time_records_buffer)
            try:
                while True:
                    record = next(time_gen)
                    all_time_records.append(record)
                    # Проверяем флаг после каждой записи (генератор может установить его)
                    if self.new_group_started:
                        print("Обнаружен флаг new_group_started во время загрузки Time записей")
                        break
            except StopIteration:
                # Генератор остановился (файл закончился)
                time_file_ended = True
                print("Time генератор остановился (файл закончился)")
                pass
        
        # Сначала обрабатываем текущие записи и начинаем генерировать выходные записи
        # Проверку разрывов делаем только после того, как хотя бы один файл закончился
        
        # Нормализуем даты для текущих записей
        old_encoder_records = self.encoder_records
        old_time_records = self.time_records
        self.encoder_records = all_encoder_records
        self.time_records = all_time_records
        self.normalize_dates()
        all_encoder_records = self.encoder_records
        all_time_records = self.time_records
        self.encoder_records = old_encoder_records
        self.time_records = old_time_records
        
        # Определяем end_time для текущих записей
        end_time = start_time
        if all_encoder_records:
            encoder_end = max([r.time for r in all_encoder_records])
            end_time = max(end_time, encoder_end)
        if all_time_records:
            time_end = max([r.time for r in all_time_records])
            end_time = max(end_time, time_end)
        
        # Вычисляем частоты дискретизации
        if len(all_encoder_records) > 1:
            time_diffs = []
            for i in range(1, min(100, len(all_encoder_records))):
                diff = (all_encoder_records[i].time - all_encoder_records[i-1].time).total_seconds()
                if diff > 0:
                    time_diffs.append(diff)
            if time_diffs:
                self.encoder_sampling_rate = sum(time_diffs) / len(time_diffs)
        
        if len(all_time_records) > 1:
            time_diffs = []
            for i in range(1, min(100, len(all_time_records))):
                diff = (all_time_records[i].time - all_time_records[i-1].time).total_seconds()
                if diff > 0:
                    time_diffs.append(diff)
            if time_diffs:
                self.time_sampling_rate = sum(time_diffs) / len(time_diffs)
        
        # Генерируем временные метки с заданным шагом и yield'им записи
        current_time = start_time
        record_count = 0
        
        print(f"Начинаем генерацию записей для текущих файлов: {len(all_encoder_records)} encoder, {len(all_time_records)} time записей")
        print(f"Диапазон времени для генерации: {start_time} - {end_time}")
        expected_records = int((end_time - start_time).total_seconds() / self.time_step_seconds)
        print(f"Ожидаемое количество записей: ~{expected_records}")

        # Создаем кэш времен один раз для оптимизации (вместо создания на каждый вызов)
        encoder_times_cache = [r.time for r in all_encoder_records] if all_encoder_records else []
        time_times_cache = [r.time for r in all_time_records] if all_time_records else []
        print(f"Кэш времен создан: {len(encoder_times_cache)} encoder, {len(time_times_cache)} time")

        while current_time <= end_time:
            if max_records and record_count >= max_records:
                break

            # Отладочный вывод каждые 5000 записей
            if record_count > 0 and record_count % 5000 == 0:
                progress_pct = (record_count / expected_records * 100) if expected_records > 0 else 0
                print(f"  Прогресс: {record_count}/{expected_records} записей ({progress_pct:.1f}%)")

            # Интерполируем данные из Encoder.txt (с кэшем времен)
            encoder_depth = self.interpolate_value(
                current_time, all_encoder_records, 'time', 'depth', self.encoder_sampling_rate, encoder_times_cache
            )
            encoder_speed = self.interpolate_value(
                current_time, all_encoder_records, 'time', 'speed', self.encoder_sampling_rate, encoder_times_cache
            )
            encoder_weight = self.interpolate_value(
                current_time, all_encoder_records, 'time', 'weight', self.encoder_sampling_rate, encoder_times_cache
            )
            encoder_pressure = self.interpolate_value(
                current_time, all_encoder_records, 'time', 'pressure', self.encoder_sampling_rate, encoder_times_cache
            )

            # Интерполируем данные из Time.txt (с кэшем времен)
            time_pressure = self.interpolate_value(
                current_time, all_time_records, 'time', 'pressure', self.time_sampling_rate, time_times_cache
            )
            temperature = self.interpolate_value(
                current_time, all_time_records, 'time', 'temperature', self.time_sampling_rate, time_times_cache
            )
            head_tension = self.interpolate_value(
                current_time, all_time_records, 'time', 'head_tension', self.time_sampling_rate, time_times_cache
            )
            acceleration_x = self.interpolate_value(
                current_time, all_time_records, 'time', 'acceleration_x', self.time_sampling_rate, time_times_cache
            )
            acceleration_y = self.interpolate_value(
                current_time, all_time_records, 'time', 'acceleration_y', self.time_sampling_rate, time_times_cache
            )
            acceleration_z = self.interpolate_value(
                current_time, all_time_records, 'time', 'acceleration_z', self.time_sampling_rate, time_times_cache
            )
            time_speed = self.interpolate_value(
                current_time, all_time_records, 'time', 'speed', self.time_sampling_rate, time_times_cache
            )
            surface_tension = self.interpolate_value(
                current_time, all_time_records, 'time', 'surface_tension', self.time_sampling_rate, time_times_cache
            )
            ccl_depth = self.interpolate_value(
                current_time, all_time_records, 'time', 'ccl_depth', self.time_sampling_rate, time_times_cache
            )
            element_depth = self.interpolate_value(
                current_time, all_time_records, 'time', 'element_depth', self.time_sampling_rate, time_times_cache
            )
            
            # Глубина берется из Encoder, если есть, иначе из Time
            depth = encoder_depth if encoder_depth is not None else (self.interpolate_value(
                current_time, all_time_records, 'time', 'encoder_depth', self.time_sampling_rate, time_times_cache
            ) if all_time_records else None)
            
            if depth is None:
                depth = self.null_value
            
            # Заполняем отсутствующие данные null значением
            merged_record = MergedRecord(
                time=current_time,
                depth=depth if depth != self.null_value else None,
                encoder_depth=encoder_depth if encoder_depth is not None else None,
                encoder_speed=encoder_speed if encoder_speed is not None else None,
                encoder_weight=encoder_weight if encoder_weight is not None else None,
                encoder_pressure=encoder_pressure if encoder_pressure is not None else None,
                time_pressure=time_pressure if time_pressure is not None else None,
                temperature=temperature if temperature is not None else None,
                head_tension=head_tension if head_tension is not None else None,
                acceleration_x=acceleration_x if acceleration_x is not None else None,
                acceleration_y=acceleration_y if acceleration_y is not None else None,
                acceleration_z=acceleration_z if acceleration_z is not None else None,
                time_speed=time_speed if time_speed is not None else None,
                surface_tension=surface_tension if surface_tension is not None else None,
                ccl_depth=ccl_depth if ccl_depth is not None else None,
                element_depth=element_depth if element_depth is not None else None
            )
            
            yield merged_record
            current_time += self.time_step
            record_count += 1
            
            # Проверяем, не началась ли новая группа
            if self.new_group_started:
                print(f"Группа завершена. Сгенерировано {record_count} объединенных записей")
                return
        
        print(f"Обработано {record_count} записей из текущих файлов")
        
        # Цикл для обработки следующих файлов, если текущие закончились
        while True:
            # Только теперь, когда хотя бы один файл закончился, проверяем разрывы между файлами
            # Проверяем разрыв для Encoder файлов (только если файл закончился)
            encoder_loaded_more = False
            if encoder_file_ended and self.current_encoder_file_index < len(self.encoder_files):
                # Есть следующий Encoder файл - проверяем разрыв
                next_encoder_file = self.encoder_files[self.current_encoder_file_index]
                # Быстро читаем первую запись следующего файла для проверки разрыва
                try:
                    next_encoder_records = self.parse_encoder_file(next_encoder_file)
                    if next_encoder_records and all_encoder_records:
                        last_end_time = all_encoder_records[-1].time
                        next_start_time = next_encoder_records[0].time
                        time_diff = (next_start_time - last_end_time).total_seconds()
                        
                        if abs(time_diff) < 10:  # Разница < 10 секунд - автоматически продолжаем
                            print(f"Разрыв {time_diff:.1f} сек < 10 сек, автоматически загружаем следующий Encoder файл")
                            all_encoder_records.extend(next_encoder_records)
                            self.current_encoder_file_index += 1
                            # Продолжаем читать файлы, пока они не закончатся или не будет разрыва
                            while self.current_encoder_file_index < len(self.encoder_files):
                                next_file = self.encoder_files[self.current_encoder_file_index]
                                next_records = self.parse_encoder_file(next_file)
                                if next_records:
                                    last_end = all_encoder_records[-1].time
                                    next_start = next_records[0].time
                                    next_diff = (next_start - last_end).total_seconds()
                                    if abs(next_diff) < 10:
                                        # Автоматически продолжаем
                                        all_encoder_records.extend(next_records)
                                        self.current_encoder_file_index += 1
                                    else:
                                        # Разрыв - останавливаемся, будет обработано в следующей группе
                                        break
                                else:
                                    self.current_encoder_file_index += 1
                        elif abs(time_diff) >= 10:  # Разница >= 10 секунд - создаем новую группу
                            print(f"\nВНИМАНИЕ: Разница во времени между Encoder файлами {time_diff:.1f} секунд")
                            print(f"  Предыдущий: {os.path.basename(self.encoder_files[self.current_encoder_file_index-1])}")
                            print(f"  Текущий: {os.path.basename(next_encoder_file)}")
                            print(f"  Автоматически создаем новую группу (разрыв >= 10 сек)")

                            # Не сшиваем - устанавливаем флаг новой группы
                            self.new_group_started = True
                            print("Остановка сшивания: начинается новая группа")
                            break  # Выходим из цикла while True
                except Exception as e:
                    print(f"Ошибка при проверке разрыва Encoder файлов: {e}")
            
            # Проверяем разрыв для Time файлов (только если файл закончился)
            time_loaded_more = False
            if time_file_ended and self.current_time_file_index < len(self.time_files):
                # Есть следующий Time файл - проверяем разрыв
                next_time_file = self.time_files[self.current_time_file_index]
                # Быстро читаем первую запись следующего файла для проверки разрыва
                try:
                    next_time_records = self.parse_time_file(next_time_file)
                    if next_time_records and all_time_records:
                        last_end_time = all_time_records[-1].time
                        next_start_time = next_time_records[0].time
                        time_diff = (next_start_time - last_end_time).total_seconds()
                        
                        if abs(time_diff) < 10:  # Разница < 10 секунд - автоматически продолжаем
                            print(f"Разрыв {time_diff:.1f} сек < 10 сек, автоматически загружаем следующий Time файл")
                            all_time_records.extend(next_time_records)
                            self.current_time_file_index += 1
                            # Продолжаем читать файлы, пока они не закончатся или не будет разрыва
                            while self.current_time_file_index < len(self.time_files):
                                next_file = self.time_files[self.current_time_file_index]
                                next_records = self.parse_time_file(next_file)
                                if next_records:
                                    last_end = all_time_records[-1].time
                                    next_start = next_records[0].time
                                    next_diff = (next_start - last_end).total_seconds()
                                    if abs(next_diff) < 10:
                                        # Автоматически продолжаем
                                        all_time_records.extend(next_records)
                                        self.current_time_file_index += 1
                                    else:
                                        # Разрыв - останавливаемся, будет обработано в следующей группе
                                        break
                                else:
                                    self.current_time_file_index += 1
                        elif abs(time_diff) >= 10:  # Разница >= 10 секунд - создаем новую группу
                            print(f"\nВНИМАНИЕ: Разница во времени между Time файлами {time_diff:.1f} секунд")
                            print(f"  Предыдущий: {os.path.basename(self.time_files[self.current_time_file_index-1])}")
                            print(f"  Текущий: {os.path.basename(next_time_file)}")
                            print(f"  Автоматически создаем новую группу (разрыв >= 10 сек)")

                            # Не сшиваем - устанавливаем флаг новой группы
                            self.new_group_started = True
                            print("Остановка сшивания: начинается новая группа")
                            break  # Выходим из цикла while True
                except Exception as e:
                    print(f"Ошибка при проверке разрыва Time файлов: {e}")
            
            # Если загрузили новые файлы, обновляем данные и продолжаем генерацию
            if encoder_loaded_more or time_loaded_more:
                # Обновляем normalize_dates для новых записей
                old_encoder_records = self.encoder_records
                old_time_records = self.time_records
                self.encoder_records = all_encoder_records
                self.time_records = all_time_records
                self.normalize_dates()
                all_encoder_records = self.encoder_records
                all_time_records = self.time_records
                self.encoder_records = old_encoder_records
                self.time_records = old_time_records
                
                # Обновляем end_time
                new_end_time = start_time
                if all_encoder_records:
                    encoder_end = max([r.time for r in all_encoder_records])
                    new_end_time = max(new_end_time, encoder_end)
                if all_time_records:
                    time_end = max([r.time for r in all_time_records])
                    new_end_time = max(new_end_time, time_end)
                end_time = new_end_time
                
                # Обновляем частоты дискретизации
                if len(all_encoder_records) > 1:
                    time_diffs = []
                    for i in range(1, min(100, len(all_encoder_records))):
                        diff = (all_encoder_records[i].time - all_encoder_records[i-1].time).total_seconds()
                        if diff > 0:
                            time_diffs.append(diff)
                    if time_diffs:
                        self.encoder_sampling_rate = sum(time_diffs) / len(time_diffs)
                
                if len(all_time_records) > 1:
                    time_diffs = []
                    for i in range(1, min(100, len(all_time_records))):
                        diff = (all_time_records[i].time - all_time_records[i-1].time).total_seconds()
                        if diff > 0:
                            time_diffs.append(diff)
                    if time_diffs:
                        self.time_sampling_rate = sum(time_diffs) / len(time_diffs)
                
                print(f"Загружены дополнительные файлы, продолжаем генерацию до {end_time}")

                # Пересоздаем кэш времен для обновленных списков записей
                encoder_times_cache = [r.time for r in all_encoder_records] if all_encoder_records else []
                time_times_cache = [r.time for r in all_time_records] if all_time_records else []
                print(f"Кэш времен обновлен: {len(encoder_times_cache)} encoder, {len(time_times_cache)} time")

                # Продолжаем генерацию записей с текущего времени (current_time уже установлен)
                # Генерируем записи для новых данных
                while current_time <= end_time:
                    if max_records and record_count >= max_records:
                        break

                    # Интерполируем данные из Encoder.txt (с кэшем)
                    encoder_depth = self.interpolate_value(
                        current_time, all_encoder_records, 'time', 'depth', self.encoder_sampling_rate, encoder_times_cache
                    )
                    encoder_speed = self.interpolate_value(
                        current_time, all_encoder_records, 'time', 'speed', self.encoder_sampling_rate, encoder_times_cache
                    )
                    encoder_weight = self.interpolate_value(
                        current_time, all_encoder_records, 'time', 'weight', self.encoder_sampling_rate, encoder_times_cache
                    )
                    encoder_pressure = self.interpolate_value(
                        current_time, all_encoder_records, 'time', 'pressure', self.encoder_sampling_rate, encoder_times_cache
                    )

                    # Интерполируем данные из Time.txt (с кэшем)
                    time_pressure = self.interpolate_value(
                        current_time, all_time_records, 'time', 'pressure', self.time_sampling_rate, time_times_cache
                    )
                    temperature = self.interpolate_value(
                        current_time, all_time_records, 'time', 'temperature', self.time_sampling_rate, time_times_cache
                    )
                    head_tension = self.interpolate_value(
                        current_time, all_time_records, 'time', 'head_tension', self.time_sampling_rate, time_times_cache
                    )
                    acceleration_x = self.interpolate_value(
                        current_time, all_time_records, 'time', 'acceleration_x', self.time_sampling_rate, time_times_cache
                    )
                    acceleration_y = self.interpolate_value(
                        current_time, all_time_records, 'time', 'acceleration_y', self.time_sampling_rate, time_times_cache
                    )
                    acceleration_z = self.interpolate_value(
                        current_time, all_time_records, 'time', 'acceleration_z', self.time_sampling_rate, time_times_cache
                    )
                    time_speed = self.interpolate_value(
                        current_time, all_time_records, 'time', 'speed', self.time_sampling_rate, time_times_cache
                    )
                    surface_tension = self.interpolate_value(
                        current_time, all_time_records, 'time', 'surface_tension', self.time_sampling_rate, time_times_cache
                    )
                    ccl_depth = self.interpolate_value(
                        current_time, all_time_records, 'time', 'ccl_depth', self.time_sampling_rate, time_times_cache
                    )
                    element_depth = self.interpolate_value(
                        current_time, all_time_records, 'time', 'element_depth', self.time_sampling_rate, time_times_cache
                    )
                    
                    # Глубина берется из Encoder, если есть, иначе из Time
                    depth = encoder_depth if encoder_depth is not None else (self.interpolate_value(
                        current_time, all_time_records, 'time', 'encoder_depth', self.time_sampling_rate, time_times_cache
                    ) if all_time_records else None)
                    
                    if depth is None:
                        depth = self.null_value
                    
                    # Заполняем отсутствующие данные null значением
                    merged_record = MergedRecord(
                        time=current_time,
                        depth=depth if depth != self.null_value else None,
                        encoder_depth=encoder_depth if encoder_depth is not None else None,
                        encoder_speed=encoder_speed if encoder_speed is not None else None,
                        encoder_weight=encoder_weight if encoder_weight is not None else None,
                        encoder_pressure=encoder_pressure if encoder_pressure is not None else None,
                        time_pressure=time_pressure if time_pressure is not None else None,
                        temperature=temperature if temperature is not None else None,
                        head_tension=head_tension if head_tension is not None else None,
                        acceleration_x=acceleration_x if acceleration_x is not None else None,
                        acceleration_y=acceleration_y if acceleration_y is not None else None,
                        acceleration_z=acceleration_z if acceleration_z is not None else None,
                        time_speed=time_speed if time_speed is not None else None,
                        surface_tension=surface_tension if surface_tension is not None else None,
                        ccl_depth=ccl_depth if ccl_depth is not None else None,
                        element_depth=element_depth if element_depth is not None else None
                    )
                    
                    yield merged_record
                    current_time += self.time_step
                    record_count += 1
                    
                    # Проверяем, не началась ли новая группа
                    if self.new_group_started:
                        print(f"Группа завершена. Сгенерировано {record_count} объединенных записей")
                        return
                
                print(f"Обработано еще {record_count} записей после загрузки новых файлов")
                # Возвращаемся к началу цикла while True для проверки следующих файлов
                continue
            else:
                # Нет новых файлов - выходим из цикла
                break
        
        print(f"Сгенерировано {record_count} объединенных записей")


if __name__ == "__main__":
    # Тестирование парсера
    parser = LogParser(
        encoder_files=["Samples/Encoder.txt"],
        time_files=["Samples/Time.txt"],
        time_step_seconds=1
    )
    
    print("Парсинг Encoder файлов...")
    encoder_records = parser.parse_all_encoder_files()
    
    print("Парсинг Time файлов (первые 10000 записей для теста)...")
    time_records = parser.parse_all_time_files(max_records=10000)
    
    print("Объединение логов (первые 100 записей для теста)...")
    merged_generator = parser.merge_logs(max_records=100)
    
    print(f"\nПервые 5 объединенных записей:")
    for i, record in enumerate(merged_generator):
        if i >= 5:
            break
        print(f"{i+1}. {record.time}: depth={record.depth}, "
              f"encoder_depth={record.encoder_depth}, "
              f"temperature={record.temperature}")
