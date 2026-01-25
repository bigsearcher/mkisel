"""
Utility for converting Excel files between formats
"""
from pathlib import Path
import pandas as pd


def convert_xls_to_xlsx(xls_file, output_file=None, delete_original=False):
    """
    Convert .xls file to .xlsx format using pandas + xlrd

    Args:
        xls_file: Path to input .xls file
        output_file: Path for output .xlsx file (optional, defaults to _converted.xlsx)
        delete_original: If True, delete original .xls file after conversion

    Returns:
        Path to converted .xlsx file

    Raises:
        ImportError: If xlrd package is not installed
        Exception: If conversion fails
    """
    xls_path = Path(xls_file)

    if not xls_path.exists():
        raise FileNotFoundError(f"File not found: {xls_file}")

    # Determine output file path
    if output_file is None:
        output_file = xls_path.parent / f"{xls_path.stem}_converted.xlsx"
    else:
        output_file = Path(output_file)

    # Delete existing output file if it exists
    if output_file.exists():
        try:
            output_file.unlink()
        except Exception:
            pass

    try:
        # Read .xls file using pandas with xlrd engine
        try:
            all_sheets = pd.read_excel(str(xls_path), sheet_name=None, engine='xlrd')
        except ImportError:
            raise ImportError(
                "xlrd package is required to convert .xls files.\n"
                "Please install it: pip install xlrd"
            )

        # Write to .xlsx using openpyxl engine
        with pd.ExcelWriter(str(output_file), engine='openpyxl') as writer:
            for sheet_name, df in all_sheets.items():
                # Truncate sheet name if too long (Excel limit is 31 characters)
                sheet_name_truncated = sheet_name[:31] if len(sheet_name) > 31 else sheet_name
                df.to_excel(writer, sheet_name=sheet_name_truncated, index=False)

        # Verify output file was created
        if not output_file.exists():
            raise RuntimeError(f"Failed to create output file: {output_file}")

        # Delete original file if requested
        if delete_original:
            try:
                xls_path.unlink()
            except Exception:
                pass

        return output_file

    except Exception as e:
        # Clean up output file if conversion failed
        if output_file.exists():
            try:
                output_file.unlink()
            except Exception:
                pass
        raise


if __name__ == "__main__":
    import sys

    if len(sys.argv) < 2:
        print("Usage: python file_converter.py <input.xls> [output.xlsx]")
        print()
        print("Converts .xls file to .xlsx format")
        sys.exit(1)

    input_file = sys.argv[1]
    output_file = sys.argv[2] if len(sys.argv) >= 3 else None

    try:
        result = convert_xls_to_xlsx(input_file, output_file)
        print(f"Success! Converted to: {result}")
    except Exception as e:
        print(f"Error: {e}")
        sys.exit(1)
