"""
Main script to process tally files
"""
import sys
import os
from pathlib import Path

# Add current directory to path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from utils.clean_excel import clean_excel
from utils.calculate_formulas import calculate_formulas
from tally_parser import parse_tally_file
from excel_generator import generate_excel


def process_tally(input_file, output_file=None, skip_clean=False):
    """
    Process tally file: clean, parse, and generate output

    Args:
        input_file: Path to input tally Excel file
        output_file: Path for output file (optional)
        skip_clean: If True, assume input is already cleaned
    """
    input_path = Path(input_file)

    # Generate output filename if not provided
    if output_file is None:
        output_file = input_path.parent / f"{input_path.stem}_output.xlsx"

    print("="*80)
    print("TALLY PARSER")
    print("="*80)
    print(f"Input file:  {input_file}")
    print(f"Output file: {output_file}")
    print("="*80)

    # Step 1: Calculate formulas (if needed)
    if not skip_clean:
        print("\nStep 1a: Calculating formulas in Excel file...")
        try:
            # Calculate formulas using Excel COM
            # This will overwrite the input file with calculated values
            calculate_formulas(str(input_path))
            print("Formulas calculated successfully")
        except Exception as e:
            print(f"Warning: Could not calculate formulas: {e}")
            print("Continuing without formula calculation (may result in missing data)")

    # Step 2: Clean Excel (if needed)
    if not skip_clean:
        print("\nStep 2: Cleaning Excel file...")
        cleaned_file = input_path.parent / f"{input_path.stem}_cleaned.xlsx"

        try:
            clean_excel(str(input_path), str(cleaned_file))
            input_for_parsing = str(cleaned_file)
        except Exception as e:
            print(f"Error during cleaning: {e}")
            print("Trying to parse original file...")
            input_for_parsing = str(input_path)
    else:
        print("\nStep 1: Skipping cleaning (using input file directly)")
        input_for_parsing = str(input_path)

    # Step 3: Parse tally data
    print("\nStep 3: Parsing tally data...")
    try:
        data = parse_tally_file(input_for_parsing)

        if not data:
            print("Warning: No data was extracted!")
            return

        print(f"Successfully parsed {len(data)} rows")

        # Show sample
        print("\nSample data (first 3 rows):")
        for i, row in enumerate(data[:3], 1):
            print(f"  {i}. Item#={row['item_num']}, Length={row['length']}, "
                  f"Depth={row['depth']}, Comment={row['comment'][:40]}...")

    except Exception as e:
        print(f"Error during parsing: {e}")
        import traceback
        traceback.print_exc()
        return

    # Step 4: Generate output Excel
    print("\nStep 4: Generating output Excel...")
    try:
        generate_excel(data, str(output_file))
    except Exception as e:
        print(f"Error during Excel generation: {e}")
        import traceback
        traceback.print_exc()
        return

    print("\n" + "="*80)
    print("PROCESSING COMPLETE!")
    print("="*80)
    print(f"Output file created: {output_file}")
    print(f"Total rows: {len(data)}")
    print("="*80)


def main():
    """Main entry point"""
    if len(sys.argv) < 2:
        print("Usage: python process_tally.py <input_file> [output_file] [--skip-clean]")
        print()
        print("Arguments:")
        print("  input_file   : Path to input tally Excel file")
        print("  output_file  : Path for output file (optional, default: <input>_output.xlsx)")
        print("  --skip-clean : Skip cleaning step (use if file is already cleaned)")
        print()
        print("Example:")
        print("  python process_tally.py Sample/Lower.xlsx")
        print("  python process_tally.py Sample/Lower_cleaned.xlsx output.xlsx --skip-clean")
        sys.exit(1)

    input_file = sys.argv[1]
    output_file = None
    skip_clean = False

    # Parse arguments
    for i in range(2, len(sys.argv)):
        arg = sys.argv[i]
        if arg == '--skip-clean':
            skip_clean = True
        elif not output_file:
            output_file = arg

    # Process
    try:
        process_tally(input_file, output_file, skip_clean)
    except KeyboardInterrupt:
        print("\nInterrupted by user")
        sys.exit(1)
    except Exception as e:
        print(f"\nUnexpected error: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)


if __name__ == "__main__":
    main()
