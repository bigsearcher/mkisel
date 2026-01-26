"""
Основной скрипт для объединения логов и создания LAS файла.
Использование:
    python main.py --encoder Samples/Encoder.txt --time Samples/Time.txt --output output.las
"""

import argparse
import sys
import os
from datetime import datetime
from log_merger import LogParser
from las_generator import LASGenerator


def setup_directories():
    """Создает необходимые папки для работы программы"""
    os.makedirs('log', exist_ok=True)
    os.makedirs('las', exist_ok=True)

def setup_logging():
    """Настраивает логирование в файл"""
    log_filename = f"log/log_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
    
    class Tee:
        """Класс для дублирования вывода в консоль и файл"""
        def __init__(self, *files):
            self.files = files
        
        def write(self, obj):
            for f in self.files:
                f.write(obj)
                f.flush()
        
        def flush(self):
            for f in self.files:
                f.flush()
    
    log_file = open(log_filename, 'w', encoding='utf-8')
    sys.stdout = Tee(sys.stdout, log_file)
    sys.stderr = Tee(sys.stderr, log_file)
    
    return log_file

def main():
    # Создаем необходимые папки
    setup_directories()
    
    # Настраиваем логирование
    log_file = None
    try:
        log_file = setup_logging()
    except Exception as e:
        print(f"Предупреждение: не удалось настроить логирование: {e}")
    
    parser = argparse.ArgumentParser(
        description='Объединение логов Encoder.txt и Time.txt в LAS файл'
    )
    parser.add_argument(
        '--encoder',
        type=str,
        nargs='+',
        default=['Samples/Encoder.txt'],
        help='Путь(и) к файлу(ам) Encoder.txt (можно указать несколько)'
    )
    parser.add_argument(
        '--time',
        type=str,
        nargs='+',
        default=['Samples/Time.txt'],
        help='Путь(и) к файлу(ам) Time.txt (можно указать несколько)'
    )
    parser.add_argument(
        '--output',
        type=str,
        default='las/output.las',
        help='Путь к выходному LAS файлу (по умолчанию las/output.las)'
    )
    parser.add_argument(
        '--time-step',
        type=float,
        default=0.1,
        help='Шаг времени для синхронизации (секунды, по умолчанию 0.1 = 100 мс)'
    )
    parser.add_argument(
        '--depth-step',
        type=float,
        default=0.1,
        help='Шаг глубины для LAS файла (метры, по умолчанию 0.1)'
    )
    parser.add_argument(
        '--well-name',
        type=str,
        default='UNKNOWN',
        help='Название скважины'
    )
    parser.add_argument(
        '--company',
        type=str,
        default='UNKNOWN',
        help='Название компании'
    )
    parser.add_argument(
        '--max-records',
        type=int,
        default=None,
        help='Максимальное количество записей для обработки (для тестирования)'
    )
    
    args = parser.parse_args()
    
    # Убеждаемся, что выходной файл находится в папке las
    if not args.output.startswith('las/'):
        output_filename = os.path.basename(args.output)
        args.output = f"las/{output_filename}"
    
    print("=" * 60)
    print("Объединение логов и создание LAS файла")
    print("=" * 60)
    print(f"Логи записываются в папку: log/")
    print(f"LAS файлы сохраняются в папку: las/")
    print(f"Encoder файл(ы): {', '.join(args.encoder)}")
    print(f"Time файл(ы): {', '.join(args.time)}")
    print(f"Выходной файл: {args.output}")
    print(f"Шаг времени: {args.time_step} сек ({args.time_step*1000:.0f} мс)")
    print(f"Шаг глубины: {args.depth_step} м")
    print()
    
    try:
        # Создаем парсер
        log_parser = LogParser(
            encoder_files=args.encoder,
            time_files=args.time,
            time_step_seconds=args.time_step
        )
        
        # Парсим файлы
        print("Шаг 1: Парсинг Encoder файлов...")
        encoder_records = log_parser.parse_all_encoder_files()
        if not encoder_records:
            print("ОШИБКА: Не удалось загрузить данные из Encoder файлов")
            sys.exit(1)
        
        print("\nШаг 2: Парсинг Time файлов...")
        time_records = log_parser.parse_all_time_files(max_records=args.max_records)
        if not time_records:
            print("ОШИБКА: Не удалось загрузить данные из Time файлов")
            sys.exit(1)
        
        # Объединяем логи
        print("Шаг 3: Объединение логов с синхронизацией по времени...")
        merged_records = log_parser.merge_logs(max_records=args.max_records)
        if not merged_records:
            print("ОШИБКА: Не удалось объединить логи")
            sys.exit(1)
        
        # Получаем дату изменения первого файла Time.txt
        time_file_date = None
        time_file_timestamp = None
        if args.time:
            try:
                time_file_path = args.time[0]
                if os.path.exists(time_file_path):
                    time_file_timestamp = os.path.getmtime(time_file_path)
                    time_file_date = datetime.fromtimestamp(time_file_timestamp)
                    print(f"Дата из файла Time.txt ({os.path.basename(time_file_path)}): {time_file_date}")
            except Exception as e:
                print(f"Предупреждение: не удалось получить дату файла Time.txt: {e}")

        # Генерируем LAS файл
        print("Шаг 4: Генерация time-based LAS файла...")
        las_generator = LASGenerator(
            well_name=args.well_name,
            company=args.company,
            time_step_ms=args.time_step * 1000  # Конвертируем секунды в миллисекунды
        )
        las_generator.generate_las(
            merged_records=merged_records,
            output_file=args.output,
            depth_step=args.depth_step,
            date=time_file_date
        )
        
        # Обновляем дату изменения выходного файла, чтобы она совпадала с исходным файлом
        if time_file_timestamp and os.path.exists(args.output):
            try:
                os.utime(args.output, (time_file_timestamp, time_file_timestamp))
                print(f"Дата изменения файла {args.output} обновлена до {time_file_date}")
            except Exception as e:
                print(f"Предупреждение: не удалось обновить дату файла: {e}")
        
        print()
        print("=" * 60)
        print("ГОТОВО!")
        print(f"LAS файл создан: {args.output}")
        print(f"Всего объединено записей: {len(merged_records)}")
        print("=" * 60)
        
    except Exception as e:
        print(f"\nОШИБКА: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)
    finally:
        # Восстанавливаем stdout и закрываем файл лога
        if log_file:
            try:
                sys.stdout = sys.__stdout__
                sys.stderr = sys.__stderr__
                log_file.close()
            except:
                pass


if __name__ == "__main__":
    main()
