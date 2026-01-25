"""
Calculate Excel formulas using Microsoft Excel COM
"""
import win32com.client
import os
import sys
from pathlib import Path


def calculate_formulas(input_file, output_file=None):
    """
    Open Excel file in Microsoft Excel, calculate formulas, and save

    Args:
        input_file: Path to input Excel file
        output_file: Path for output file (optional, defaults to same as input)
    """
    input_path = Path(input_file).resolve()

    if not input_path.exists():
        raise FileNotFoundError(f"File not found: {input_file}")

    if output_file is None:
        # If input is .xlsm or .xls, convert to .xlsx
        input_ext = input_path.suffix.lower()
        if input_ext == '.xlsm':
            output_file = input_path.with_suffix('.xlsx')
            print(f"  Note: .xlsm file will be saved as .xlsx (macros removed)")
        elif input_ext == '.xls':
            output_file = input_path.with_suffix('.xlsx')
            print(f"  Note: .xls file will be saved as .xlsx format")
        else:
            output_file = input_path
    else:
        output_file = Path(output_file).resolve()

    print(f"Opening {input_path.name} in Excel...")

    # Create Excel application
    excel = None
    wb = None
    try:
        # Try to create Excel application
        try:
            excel = win32com.client.Dispatch("Excel.Application")
        except Exception as e:
            raise RuntimeError(f"Failed to start Excel application. Make sure Microsoft Excel is installed.\nError: {str(e)}")
        
        # Make Excel invisible to avoid interrupting user
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.ScreenUpdating = False
        excel.EnableEvents = False
        
        # Disable macros globally before opening any file
        try:
            excel.AutomationSecurity = 2  # msoAutomationSecurityForceDisable - disables all macros
        except Exception:
            # Fallback: try to set macro security via registry-like approach
            try:
                excel.Application.AutomationSecurity = 2
            except Exception:
                pass  # Some Excel versions may not support this

        # Open workbook
        try:
            # Open with UpdateLinks=False to avoid updating external links
            # Disable macros: Notify=False means don't run macros automatically
            # xlOpenXMLWorkbookMacroEnabled (52) files will open but macros disabled
            wb = excel.Workbooks.Open(
                str(input_path),
                UpdateLinks=0,  # Don't update external links
                ReadOnly=False,
                CorruptLoad=0,  # xlNormalLoad
                Notify=False,  # Don't notify about macros (effectively disables them)
                AddToMru=False  # Don't add to recent files
            )
            
            # Explicitly disable macros if workbook has them
            try:
                # Set macro security to disable all macros
                excel.AutomationSecurity = 2  # msoAutomationSecurityForceDisable
            except Exception:
                pass  # Some Excel versions may not support this
                
        except Exception as e:
            raise RuntimeError(f"Failed to open file in Excel: {str(e)}")

        # Break external links (remove connections to other workbooks)
        print("Breaking external links...")
        try:
            # Get all external links
            links = wb.LinkSources()
            if links:
                for link in links:
                    try:
                        # Break the link and convert to values
                        wb.BreakLink(link, 1)  # xlLinkTypeExcelLinks = 1
                    except Exception as e:
                        print(f"  Warning: Could not break link {link}: {e}")
        except Exception as e:
            print(f"  Warning: Could not process external links: {e}")

        # Switch all worksheets to normal view (not page layout)
        print("Switching to normal view...")
        for ws in wb.Worksheets:
            try:
                ws.Activate()
                excel.ActiveWindow.View = 1  # xlNormalView = 1
            except Exception:
                pass

        print("Calculating formulas...")
        wb.Application.CalculateFull()
        
        # Force recalculation of all sheets
        for ws in wb.Worksheets:
            try:
                ws.Calculate()
            except Exception:
                pass
        
        # Wait a bit for calculations to complete
        import time
        time.sleep(0.5)

        # Save - use SaveAs with explicit format to ensure clean file
        print(f"Saving to {output_file.name}...")
        
        # Delete output file if it exists
        if output_file.exists():
            try:
                output_file.unlink()
            except Exception:
                pass
        
        # Save as new file to ensure clean format
        # Use CreateBackup=False to avoid creating backup files
        # Convert .xls and .xlsm to .xlsx format
        file_ext = output_file.suffix.lower()
        input_ext = input_path.suffix.lower()
        
        # If input is .xls or .xlsm, always save as .xlsx
        if input_ext == '.xls' or input_ext == '.xlsm' or file_ext == '.xlsx' or file_ext == '.xlsm':
            # Change extension to .xlsx if needed
            if file_ext != '.xlsx':
                output_file = output_file.with_suffix('.xlsx')
                if input_ext == '.xls':
                    print(f"  Note: Converting .xls to .xlsx format")
                elif input_ext == '.xlsm':
                    print(f"  Note: Converting .xlsm to .xlsx (macros removed)")
            
            # Save as xlsx (Excel 2007-2019 format) - use xlOpenXMLWorkbook (51)
            wb.SaveAs(
                str(output_file), 
                FileFormat=51,  # xlOpenXMLWorkbook (no macros)
                ConflictResolution=2,  # xlLocalSessionChanges
                CreateBackup=False
            )
        elif file_ext == '.xls':
            # If output explicitly requested as .xls, save in old format
            # (but this shouldn't happen in normal workflow)
            wb.SaveAs(
                str(output_file), 
                FileFormat=56,  # xlExcel8
                ConflictResolution=2,
                CreateBackup=False
            )
        else:
            # Default to xlsx (no macros)
            wb.SaveAs(
                str(output_file), 
                FileFormat=51,  # xlOpenXMLWorkbook
                ConflictResolution=2,
                CreateBackup=False
            )
        
        # Wait for file to be written
        time.sleep(0.5)
        
        # Verify file was created
        if not output_file.exists():
            raise RuntimeError(f"Failed to save file: {output_file}")

        # Close workbook without saving (already saved with SaveAs)
        wb.Close(SaveChanges=False)
        wb = None

        print(f"Done! Formulas calculated and saved to: {output_file}")

    except Exception as e:
        print(f"Error: {e}")
        # Try to close workbook if it's open
        if wb:
            try:
                wb.Close(SaveChanges=False)
            except:
                pass
        raise

    finally:
        # Clean up Excel application
        if wb:
            try:
                wb.Close(SaveChanges=False)
            except:
                pass
            wb = None
        
        if excel:
            try:
                # Re-enable screen updating before quitting
                excel.ScreenUpdating = True
                excel.EnableEvents = True
                excel.Quit()
                # Release COM object
                del excel
            except:
                pass
            excel = None
        
        # Force garbage collection to release COM objects
        import gc
        gc.collect()


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python calculate_formulas.py <input_file> [output_file]")
        print()
        print("This utility opens the Excel file in Microsoft Excel,")
        print("calculates all formulas, and saves the result.")
        print()
        print("If output_file is not specified, overwrites the input file.")
        sys.exit(1)

    input_file = sys.argv[1]
    output_file = sys.argv[2] if len(sys.argv) >= 3 else None

    try:
        calculate_formulas(input_file, output_file)
    except Exception as e:
        print(f"Failed: {e}")
        sys.exit(1)
