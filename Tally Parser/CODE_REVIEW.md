# ОТЧЕТ ПО ПРОВЕРКЕ КОДА TALLY PARSER

## КРИТИЧЕСКИЕ ПРОБЛЕМЫ

### 1. ❌ УТЕЧКА EXCEL COM ПРОЦЕССОВ (calculate_formulas.py)

**Проблема:**
- Если exception происходит в строке 186 (`wb.Close`), переменная `wb` остается != None
- В finally блоке произойдет повторная попытка закрыть уже закрытый workbook
- Excel процесс может остаться в памяти

**Код (строки 185-208):**
```python
wb.Close(SaveChanges=False)
wb = None  # ❌ ЭТОГО НЕТ!

except Exception as e:
    if wb:
        try:
            wb.Close(SaveChanges=False)  # Может закрыть уже закрытый wb
```

**Решение:**
```python
wb.Close(SaveChanges=False)
wb = None  # ✅ Добавить обнуление после закрытия
```

---

### 2. ❌ НЕНАДЕЖНОЕ ОЖИДАНИЕ ЗАПИСИ ФАЙЛА (calculate_formulas.py)

**Проблема:**
- Используется `time.sleep(0.5)` для ожидания записи файла (строки 125, 179)
- На медленных дисках файл может не успеть записаться
- Нет гарантии что файл действительно записан

**Код (строки 178-183):**
```python
time.sleep(0.5)  # ❌ Ненадежно!

if not output_file.exists():
    raise RuntimeError(f"Failed to save file: {output_file}")
```

**Решение:**
```python
# Ждем появления файла с таймаутом
timeout = 10
elapsed = 0
while not output_file.exists() and elapsed < timeout:
    time.sleep(0.1)
    elapsed += 0.1

if not output_file.exists():
    raise RuntimeError(f"Failed to save file: {output_file}")
```

---

### 3. ❌ УТЕЧКА ВРЕМЕННЫХ ФАЙЛОВ (manual_column_selector.py)

**Проблема 1 - при закрытии через Cancel:**
- Создаются временные файлы: `_converted.xlsx`, `_fixed_temp.xlsx`
- Если пользователь нажимает Cancel, эти файлы НЕ удаляются
- При закрытии окна через X файлы тоже остаются

**Код (строки 362-369):**
```python
self.cancel_button = ttk.Button(
    button_container,
    text="Cancel",
    command=self.dialog.destroy,  # ❌ Просто уничтожает окно, не удаляет файлы
    width=30,
    style='Large.TButton'
)
```

**Проблема 2 - переменная может быть None:**
- `self._temp_file_to_cleanup` может не установиться если exception произошел рано
- В строках 1491-1496 проверка есть, но она не покрывает все случаи

**Решение:**
Добавить метод cleanup и вызывать его всегда:
```python
def __init__(self, parent, file_path):
    ...
    self._temp_files = []  # Список всех временных файлов

    # Привязать cleanup к закрытию окна
    self.dialog.protocol("WM_DELETE_WINDOW", self._cleanup_and_close)

def _cleanup_and_close(self):
    """Очистка временных файлов и закрытие окна"""
    for temp_file in self._temp_files:
        try:
            if Path(temp_file).exists():
                Path(temp_file).unlink()
        except Exception:
            pass
    self.dialog.destroy()
```

---

### 4. ❌ WORKBOOK ОБЪЕКТЫ НЕ ЗАКРЫВАЮТСЯ

**Проблема:**
Во всех файлах используется `load_workbook()` без явного закрытия:
- manual_column_selector.py (строка 157): `self.workbook = load_workbook()`
- tallyconverter.py (строки 539, 651): `wb = load_workbook()`
- process_tally.py через parse_tally_file() и generate_excel()

**Последствия:**
- Файлы Excel остаются открытыми в памяти
- При обработке многих файлов - утечка памяти
- На Windows файлы могут оставаться заблокированными

**Решение для всех случаев:**
```python
# Вместо:
wb = load_workbook(file_path)
# ... работа с wb ...

# Использовать:
wb = None
try:
    wb = load_workbook(file_path)
    # ... работа с wb ...
finally:
    if wb:
        wb.close()
```

---

### 5. ❌ МЕРТВЫЙ КОД (manual_column_selector.py)

**Около 500 строк неиспользуемого кода:**

| Функция | Строки | Статус |
|---------|--------|--------|
| `_show_column_selection_dialog_OLD` | 832-1050 | ❌ НЕ ИСПОЛЬЗУЕТСЯ |
| `_show_header_row_selection_dialog_OLD` | 1052-1247 | ❌ НЕ ИСПОЛЬЗУЕТСЯ |
| `_highlight_selections` | 1249-1292 | ❌ НЕ ИСПОЛЬЗУЕТСЯ |
| `_on_cell_click` | 1293-1296 | ❌ Legacy метод |
| `_scroll_to_row` | 810-819 | ❌ НЕ ИСПОЛЬЗУЕТСЯ |
| `_scroll_to_column` | 821-830 | ❌ НЕ ИСПОЛЬЗУЕТСЯ |

**Рекомендация:** Удалить весь мертвый код для улучшения читаемости и поддержки.

---

### 6. ❌ НЕТ CLEANUP ПРИ ЗАКРЫТИИ GUI (tallyconverter.py)

**Проблема:**
- При закрытии главного окна через X не происходит очистка
- Временные файлы `_converted.xlsx` остаются на диске (строка 32)
- Нет обработчика `WM_DELETE_WINDOW`

**Решение:**
```python
class TallyParserGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Tally Parser")
        self.root.geometry("500x400")

        # Временные файлы для очистки
        self.temp_files = []

        # Привязать cleanup к закрытию окна
        self.root.protocol("WM_DELETE_WINDOW", self._on_closing)

        # Current file
        self.current_file = None
        self.cleaned_file = None
        self.converted_file = None

        # Setup UI
        self._setup_ui()

    def _on_closing(self):
        """Очистка при закрытии приложения"""
        # Удалить временные файлы
        for temp_file in self.temp_files:
            try:
                if Path(temp_file).exists():
                    Path(temp_file).unlink()
            except Exception:
                pass

        # Удалить converted файл если он существует
        if self.converted_file:
            try:
                if Path(self.converted_file).exists():
                    Path(self.converted_file).unlink()
            except Exception:
                pass

        self.root.destroy()
```

---

### 7. ⚠️ SUBPROCESS БЕЗ ОТСЛЕЖИВАНИЯ (tallyconverter.py)

**Проблема:**
- Открытие Excel/Notepad через subprocess.Popen без сохранения ссылки на процесс
- Если процесс зависнет, никто его не закроет
- При закрытии GUI дочерние процессы могут остаться

**Код (строки 132-136):**
```python
creation_flags = subprocess.CREATE_NO_WINDOW if hasattr(subprocess, 'CREATE_NO_WINDOW') else 0
subprocess.Popen(['excel.exe', str(file_path)], shell=False, creationflags=creation_flags)
# ❌ Процесс не отслеживается
```

**Решение:**
```python
class TallyParserGUI:
    def __init__(self, root):
        ...
        self.child_processes = []  # Список дочерних процессов

    def _open_file_in_excel(self, file_path):
        try:
            ...
            process = subprocess.Popen(['excel.exe', str(file_path)], ...)
            self.child_processes.append(process)
        except Exception:
            pass

    def _on_closing(self):
        # Закрыть дочерние процессы
        for process in self.child_processes:
            try:
                if process.poll() is None:  # Процесс еще работает
                    process.terminate()
            except Exception:
                pass
        ...
```

**Примечание:** Для Excel/Notepad это не критично, т.к. пользователь сам их закроет, но это не чистое решение.

---

## СРЕДНИЕ ПРОБЛЕМЫ

### 8. ⚠️ ОТСУТСТВИЕ TIMEOUT ДЛЯ EXCEL ОПЕРАЦИЙ

**Проблема:**
- Excel операции могут зависнуть навсегда
- Нет таймаута для `wb.Application.CalculateFull()` (строка 114)
- Пользователь не сможет отменить операцию

**Рекомендация:**
Использовать threading с таймаутом для Excel операций.

---

### 9. ⚠️ MAGIC NUMBERS В КОДЕ

**Примеры:**
- `FileFormat=51` (строка 154) - не понятно что это
- `FileFormat=56` (строка 165)
- `UpdateLinks=0` (строка 72)
- `CorruptLoad=0` (строка 74)

**Рекомендация:**
Добавить константы с говорящими именами:
```python
EXCEL_FORMAT_XLSX = 51  # xlOpenXMLWorkbook
EXCEL_FORMAT_XLS = 56   # xlExcel8
```

---

## СТРУКТУРНЫЕ ПРОБЛЕМЫ

### 10. ⚠️ ДУБЛИРОВАНИЕ КОДА

**Проблема:**
Логика конвертации .xls → .xlsx дублируется в 3 местах:
1. tallyconverter.py: `_convert_xls_to_xlsx` (строки 177-241)
2. manual_column_selector.py: `_load_excel_data` (строки 116-153)
3. ModeSelectionDialog: `use_auto_mode` (строки 983-1021)

**Рекомендация:**
Вынести в utils/file_converter.py:
```python
def convert_xls_to_xlsx(xls_file):
    """Конвертация .xls в .xlsx"""
    ...
```

---

### 11. ⚠️ ОТСУТСТВИЕ LOGGING

**Проблема:**
- Используется print() для вывода сообщений
- Нет логирования ошибок в файл
- При автоматической обработке 20 файлов сложно понять где именно произошла ошибка

**Рекомендация:**
Использовать модуль logging:
```python
import logging

logging.basicConfig(
    filename='tally_parser.log',
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)

logger = logging.getLogger(__name__)
logger.info("Processing file: %s", file_path)
logger.error("Failed to parse file: %s", error_msg)
```

---

## РЕКОМЕНДАЦИИ ПО ПРИОРИТЕТАМ

### 🔴 КРИТИЧНО (исправить в первую очередь):
1. Добавить cleanup временных файлов при закрытии окон
2. Добавить `wb = None` после `wb.Close()` в calculate_formulas.py
3. Закрывать workbook объекты после использования

### 🟡 ВАЖНО (исправить при возможности):
4. Удалить 500 строк мертвого кода из manual_column_selector.py
5. Добавить обработчик `WM_DELETE_WINDOW` для GUI
6. Заменить time.sleep на цикл с проверкой файла

### 🟢 ЖЕЛАТЕЛЬНО (улучшения):
7. Добавить логирование
8. Вынести дублирующийся код в отдельные функции
9. Добавить константы вместо magic numbers
10. Добавить timeout для Excel операций

---

## ИТОГОВАЯ ОЦЕНКА

| Категория | Оценка | Комментарий |
|-----------|--------|-------------|
| Утечка памяти | ⚠️ ЕСТЬ | Workbook объекты не закрываются |
| Утечка процессов | ⚠️ ВОЗМОЖНА | Excel COM процессы могут зависнуть |
| Утечка файлов | ❌ ЕСТЬ | Временные файлы не удаляются |
| Мертвый код | ❌ ЕСТЬ | 500 строк неиспользуемого кода |
| Обработка ошибок | ✅ ХОРОШО | Try/except блоки присутствуют |
| Структура кода | ⚠️ СРЕДНЕ | Дублирование, нет логирования |

**Общий вердикт:** Код работает, но имеет серьезные проблемы с утечками ресурсов и требует рефакторинга.
