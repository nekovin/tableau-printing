import subprocess
import time
import pyautogui
from pathlib import Path
import xml.etree.ElementTree as ET
import shutil

def change_selected_pac(workbook_path, new_pac_value):
    """Change the Selected PAC parameter to a new value"""
    
    # Make a backup copy first
    original = Path(workbook_path)
    backup = original.with_suffix('.twb.backup')
    shutil.copy2(original, backup)
    print(f"Backup created: {backup}")
    
    # Parse the XML
    tree = ET.parse(workbook_path)
    root = tree.getroot()
    
    changes_made = 0
    
    # Find ALL Selected PAC parameters (there are multiple instances)
    for column in root.findall(".//column"):
        if column.get('caption') == 'Selected PAC':
            # Change the alias and value attributes
            old_alias = column.get('alias')
            old_value = column.get('value')
            
            # Update alias to match the new PAC
            column.set('alias', new_pac_value)
            
            # Update value with proper format
            new_value = f'"PAC - {new_pac_value}"'
            column.set('value', new_value)
            
            # Also update the calculation formula
            calculation = column.find('calculation')
            if calculation is not None:
                old_formula = calculation.get('formula')
                new_formula = f'"PAC - {new_pac_value}"'
                calculation.set('formula', new_formula)
                print(f"  Updated calculation formula: {old_formula} → {new_formula}")
            
            print(f"Changed Selected PAC instance {changes_made + 1}:")
            print(f"  Alias: {old_alias} → {new_pac_value}")
            print(f"  Value: {old_value} → {new_value}")
            
            changes_made += 1
    
    if changes_made > 0:
        # Save the modified file
        tree.write(workbook_path, encoding='utf-8', xml_declaration=True)
        print(f"Made {changes_made} changes and saved to: {workbook_path}")
        return True
    else:
        print("Selected PAC parameter not found!")
        return False

def export_workbook_to_pdf(workbook_path, tableau_exe_path, output_dir, pac_name):
    """
    Open Tableau workbook and print all pages as a single PDF using Print to PDF
    """
    output_path = Path(output_dir)
    output_path.mkdir(exist_ok=True)
    
    print(f"Opening workbook: {workbook_path}")
    
    # Open Tableau with the workbook
    cmd = f'"{tableau_exe_path}" "{workbook_path}"'
    subprocess.Popen(cmd, shell=True)
    
    # Wait for Tableau to load
    print("Waiting for Tableau to load (25 seconds)...")
    time.sleep(25)
    
    # Click to ensure Tableau window is active
    pyautogui.click(500, 300)
    time.sleep(2)
    
    # Configure pyautogui
    pyautogui.PAUSE = 2
    pyautogui.FAILSAFE = True
    
    print("Opening Print to PDF dialog...")
    
    # File -> Print to PDF...
    pyautogui.hotkey('alt', 'f')  # Open File menu
    time.sleep(2)
    
    # Click on "Print to PDF..." option
    # Based on your screenshot, it's near the bottom of the File menu
    pyautogui.press('r')  # Press 'r' for "Print to PDF..." (the underlined letter)
    time.sleep(3)
    
    # Now we have the Print to PDF dialog open
    print("Selecting 'Entire workbook' to get all 4 pages...")
    
    # Click on "Entire workbook" radio button
    # Based on your dialog screenshot, it's the first option
    pyautogui.click(30, 67)  # Approximate location of "Entire workbook" radio button
    time.sleep(1)
    
    # Make sure "View PDF after printing" is checked (it should be by default)
    # We can see it's already checked in your screenshot
    
    # Create output filename
    pdf_filename = f"{pac_name} - Walkin.pdf"
    output_file = output_path / pdf_filename
    
    print(f"Setting filename to: {pdf_filename}")
    
    # Click OK to proceed with print
    pyautogui.click(365, 210)  # OK button location from your screenshot
    time.sleep(3)
    
    # Now Windows "Save Print Output As" dialog should appear
    # Type the filename
    pyautogui.write(pdf_filename)
    time.sleep(1)
    
    # Press Enter to save
    pyautogui.press('enter')
    
    # Wait for PDF generation to complete
    print(f"Generating PDF: {output_file}")
    time.sleep(15)  # Wait for PDF generation
    
    print("PDF generation completed!")
    
    # Close Tableau
    print("Closing Tableau...")
    pyautogui.hotkey('alt', 'f4')
    time.sleep(2)

def create_auburn_reports(workbook_path, tableau_exe_path, output_dir):
    """Complete workflow: change parameter and export single PDF with all 4 pages"""
    
    print("="*50)
    print("TABLEAU AUTOMATION - AUBURN WALKIN REPORT")
    print("="*50)
    
    # Step 1: Change the parameter to Auburn
    print("Step 1: Changing Selected PAC parameter to Auburn...")
    success = change_selected_pac(workbook_path, "Auburn")
    
    if not success:
        print("Failed to change parameter. Exiting.")
        return False
    
    # Step 2: Export all pages as single PDF
    print("Step 2: Printing all 4 pages to single PDF...")
    export_workbook_to_pdf(workbook_path, tableau_exe_path, output_dir, "Auburn")
    
    print("="*50)
    print("COMPLETE! Auburn - Walkin.pdf generated with all 4 pages.")
    print("="*50)
    
    return True

# Main execution
if __name__ == "__main__":
    # Configuration - UPDATE THESE PATHS
    workbook_file = r"C:\Users\CL-11\OneDrive - Lonergan Research\Repos\tableau-automation\walkin\NSW Police Service Assessment Walk-in Auburn;South West Metro.twb"
    tableau_executable = r"C:\Program Files\Tableau\Tableau 2024.3\bin\tableau.exe"
    output_directory = r"C:\Users\CL-11\OneDrive - Lonergan Research\Repos\tableau-automation\output"
    
    print("Starting Tableau automation...")
    print(f"Workbook: {workbook_file}")
    print(f"Tableau: {tableau_executable}")
    print(f"Output: {output_directory}")
    print()
    
    # Run the complete workflow
    try:
        success = create_auburn_reports(workbook_file, tableau_executable, output_directory)
        
        if success:
            print("✅ Automation completed successfully!")
            print("📄 Check your output folder for 'Auburn - Walkin.pdf'")
        else:
            print("❌ Automation failed!")
            
    except Exception as e:
        print(f"❌ Error during automation: {e}")
        print("Make sure:")
        print("- File paths are correct")
        print("- Tableau is not already running")
        print("- Don't use mouse/keyboard while script runs")
    
    finally:
        print("\nScript finished.")