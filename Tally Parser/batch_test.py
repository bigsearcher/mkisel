"""
Batch Test Script - Test tally parser on multiple files
"""
import sys
import os
from pathlib import Path
from datetime import datetime

# Fix encoding for Windows console
if sys.platform == 'win32':
    import io
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')

# Add current directory to path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from utils.calculate_formulas import calculate_formulas
from utils.clean_excel import clean_excel
from tally_parser import parse_tally_file
from excel_generator import generate_excel


def test_file(input_file, output_dir):
    """
    Test processing a single file

    Returns:
        dict with results: {
            'file': filename,
            'status': 'success'/'error',
            'rows': number of rows,
            'error': error message if failed,
            'output': output file path
        }
    """
    input_path = Path(input_file)
    result = {
        'file': input_path.name,
        'status': 'error',
        'rows': 0,
        'error': None,
        'output': None
    }

    try:
        # Create output filename
        output_file = output_dir / f"{input_path.stem}_output.xlsx"

        # Step 0: Calculate formulas for all xls* file types (.xls, .xlsx, .xlsm)
        # Excel will handle conversion and formula calculation
        file_to_process = str(input_path)
        converted_file = None
        
        if input_path.suffix.lower() in ['.xls', '.xlsx', '.xlsm']:
            try:
                # For .xls files, Excel will convert to .xlsx during SaveAs
                # For .xlsm files, Excel will save as .xlsx (macros removed)
                if input_path.suffix.lower() == '.xls':
                    # Create temporary .xlsx file path for converted .xls
                    converted_file = input_path.parent / f"{input_path.stem}_converted.xlsx"
                    print(f"  Calculating formulas and converting .xls to .xlsx...")
                    calculate_formulas(str(input_path), str(converted_file))
                    file_to_process = str(converted_file)
                    print(f"  Converted to: {converted_file.name}")
                elif input_path.suffix.lower() == '.xlsm':
                    # Convert .xlsm to .xlsx (removes macros)
                    converted_file = input_path.parent / f"{input_path.stem}_converted.xlsx"
                    print(f"  Calculating formulas and converting .xlsm to .xlsx (macros removed)...")
                    calculate_formulas(str(input_path), str(converted_file))
                    file_to_process = str(converted_file)
                    print(f"  Converted to: {converted_file.name}")
                else:
                    # .xlsx - just calculate formulas
                    print(f"  Calculating formulas...")
                    calculate_formulas(file_to_process)
            except Exception as e:
                print(f"  Warning: Could not calculate formulas: {e}")
                # Fallback: try pandas conversion for .xls only
                if input_path.suffix.lower() == '.xls' and converted_file is None:
                    try:
                        import pandas as pd
                        print(f"  Fallback: Converting .xls to .xlsx using pandas...")
                        xlsx_file = input_path.parent / f"{input_path.stem}_converted.xlsx"
                        if xlsx_file.exists():
                            try:
                                xlsx_file.unlink()
                            except Exception:
                                pass
                        all_sheets = pd.read_excel(str(input_path), sheet_name=None, engine='xlrd')
                        with pd.ExcelWriter(str(xlsx_file), engine='openpyxl') as writer:
                            for sheet_name, df in all_sheets.items():
                                sheet_name_truncated = sheet_name[:31] if len(sheet_name) > 31 else sheet_name
                                df.to_excel(writer, sheet_name=sheet_name_truncated, index=False)
                        file_to_process = str(xlsx_file)
                        print(f"  Converted to: {xlsx_file.name}")
                    except Exception as e2:
                        print(f"  Warning: Could not convert .xls to .xlsx: {e2}")
                # Continue anyway - file might still be processable

        # Step 2: Clean Excel
        cleaned_file = output_dir / f"{input_path.stem}_cleaned.xlsx"
        try:
            clean_excel(file_to_process, str(cleaned_file))
            input_for_parsing = str(cleaned_file)
        except Exception as e:
            print(f"  Warning: Could not clean file: {e}")
            input_for_parsing = file_to_process

        # Step 3: Parse tally data
        data = parse_tally_file(input_for_parsing)

        if not data:
            result['error'] = "No data extracted"
            return result

        # Step 4: Generate output Excel
        generate_excel(data, str(output_file))

        # Success
        result['status'] = 'success'
        result['rows'] = len(data)
        result['output'] = str(output_file)

    except Exception as e:
        result['error'] = str(e)

    return result


def main():
    """Main batch testing function"""
    # Find all test files
    base_dir = Path(__file__).parent
    sample_dir1 = base_dir / "Sample"
    sample_dir2 = base_dir / "Sample" / "1"

    # Collect all files matching pattern 01-20
    test_files = []

    for sample_dir in [sample_dir1, sample_dir2]:
        if not sample_dir.exists():
            continue

        for i in range(1, 21):
            # Try different extensions
            for ext in ['.xlsx', '.xls', '.xlsm']:
                filename = f"{i:02d}{ext}"
                filepath = sample_dir / filename
                if filepath.exists():
                    test_files.append(filepath)
                    break  # Found this number, move to next

    if not test_files:
        print("No test files found!")
        return

    # Create output directory
    output_dir = base_dir / "test_results"
    output_dir.mkdir(exist_ok=True)

    # Run tests
    print("=" * 80)
    print(f"BATCH TEST - {len(test_files)} files")
    print("=" * 80)
    print(f"Output directory: {output_dir}")
    print("=" * 80)
    print()

    results = []

    for idx, file_path in enumerate(test_files, 1):
        print(f"[{idx}/{len(test_files)}] Testing: {file_path.name}")
        result = test_file(file_path, output_dir)
        results.append(result)

        if result['status'] == 'success':
            print(f"  [OK] Success - {result['rows']} rows extracted")
        else:
            print(f"  [FAIL] Failed - {result['error']}")
        print()

    # Generate summary report
    print("=" * 80)
    print("TEST SUMMARY")
    print("=" * 80)

    success_count = sum(1 for r in results if r['status'] == 'success')
    fail_count = len(results) - success_count

    print(f"Total files tested: {len(results)}")
    print(f"Successful: {success_count}")
    print(f"Failed: {fail_count}")
    print()

    if success_count > 0:
        print("Successful files:")
        for r in results:
            if r['status'] == 'success':
                print(f"  [OK] {r['file']:<25} - {r['rows']:>4} rows")
        print()

    if fail_count > 0:
        print("Failed files:")
        for r in results:
            if r['status'] == 'error':
                print(f"  [FAIL] {r['file']:<25} - {r['error']}")
        print()

    # Save report to file
    report_file = output_dir / f"test_report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
    with open(report_file, 'w', encoding='utf-8') as f:
        f.write("TALLY PARSER - BATCH TEST REPORT\n")
        f.write("=" * 80 + "\n")
        f.write(f"Date: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
        f.write(f"Total files: {len(results)}\n")
        f.write(f"Successful: {success_count}\n")
        f.write(f"Failed: {fail_count}\n")
        f.write("=" * 80 + "\n\n")

        f.write("DETAILED RESULTS:\n")
        f.write("-" * 80 + "\n")
        for r in results:
            f.write(f"File: {r['file']}\n")
            f.write(f"Status: {r['status']}\n")
            if r['status'] == 'success':
                f.write(f"Rows: {r['rows']}\n")
                f.write(f"Output: {r['output']}\n")
            else:
                f.write(f"Error: {r['error']}\n")
            f.write("-" * 80 + "\n")

    print(f"Report saved to: {report_file}")
    print("=" * 80)


if __name__ == "__main__":
    main()
