import subprocess
import time
import pyautogui
from pathlib import Path
import xml.etree.ElementTree as ET
import shutil
import pandas as pd


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

def export_single_pdf(workbook_path, tableau_exe_path, output_dir, area, region):
    """
    Export PDF for a single area/region combination
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
    pyautogui.PAUSE = 1
    pyautogui.FAILSAFE = True

    # Maximize window before selecting pages
    pyautogui.hotkey('alt', 'space')  # Open window menu
    time.sleep(0.5)
    pyautogui.press('x')  # Maximize
    time.sleep(1)

    # Select all 4 pages (Page 1 to Page 4) using Shift+Click FIRST
    print("Selecting all 4 pages...")
    pyautogui.click(2211, 1403)  # Click Page 1 tab
    time.sleep(0.5)
    pyautogui.keyDown('shift')  # Hold Shift
    pyautogui.click(2418, 1403)  # Click Page 4 tab while holding Shift
    pyautogui.keyUp('shift')  # Release Shift
    time.sleep(2)

    # NOW open Print to PDF dialog
    #print("Opening Print to PDF dialog...")
    ##pyautogui.hotkey('alt', 'f')  # Open File menu
    #time.sleep(2)
    #pyautogui.press('r')  # Press 'r' for "Print to PDF..."
    #time.sleep(3)

    pyautogui.click(19, 29)
    time.sleep(1)

    # Click on "View PDF after printing" checkbox  
    pyautogui.click(54, 432)
    time.sleep(1)

    # Select "Selected sheets" radio button
    pyautogui.click(1158, 718)
    time.sleep(1)

    # Click OK button
    pyautogui.click(1300, 811)
    time.sleep(3)

    # Type the PDF filename
    #pdf_filename = f"Auburn - Walkin.pdf"
    pdf_filename = f"{area} - {region}.pdf"
    pyautogui.write(pdf_filename)

    # Click Save button
    #pyautogui.click(1533, 1303)
    #time.sleep(8)  # Wait for PDF generation

    # enter
    pyautogui.press('enter')
    time.sleep(8)  # Wait for PDF generation

    print(f"✅ Generated: {pdf_filename}")

    # Close Tableau
    pyautogui.hotkey('alt', 'f4')
    time.sleep(3)

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
            # Check if this area is PAC or PD
            area_type = "PAC"
            for member in root.findall(".//member[@alias]"):
                if member.get('alias') == new_pac_value:
                    if member.get('value', '').startswith('"PD - '):
                        area_type = "PD"
                    break
            
            # Change the alias and value attributes
            old_alias = column.get('alias')
            old_value = column.get('value')
            
            # Update alias to match the new PAC/PD
            column.set('alias', new_pac_value)
            
            # Update value with proper format
            new_value = f'"{area_type} - {new_pac_value}"'
            column.set('value', new_value)
            
            # Also update the calculation formula
            calculation = column.find('calculation')
            if calculation is not None:
                old_formula = calculation.get('formula')
                new_formula = f'"{area_type} - {new_pac_value}"'
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

def validate_filter_data(filter_file_path, workbook_path):
    """Validate that Excel values exist in TWB file"""
    # Read Excel
    df = pd.read_excel(filter_file_path)
    
    # Parse TWB XML
    tree = ET.parse(workbook_path)
    root = tree.getroot()
    
    # Extract valid PAC and PD values (Area column)
    valid_areas = set()
    for member in root.findall(".//member[@alias]"):
        value = member.get('value', '')
        if value.startswith('"PAC - ') or value.startswith('"PD - '):
            valid_areas.add(member.get('alias'))
    
    # Extract valid Region values
    valid_regions = set()
    for member in root.findall(".//member"):
        value = member.get('value', '')
        if value and not value.startswith('"PAC - ') and not value.startswith('"PD - '):
            # Remove quotes and extract region name
            region = value.strip('"')
            if region in ['Central Metropolitan', 'North West Metropolitan', 'Northern', 
                         'South West Metropolitan', 'Southern', 'Western']:
                valid_regions.add(region)
    
    # Find failures
    invalid_areas = set(df['Area']) - valid_areas
    invalid_regions = set(df['Region']) - valid_regions
    
    # Show all failures with reasons
    if invalid_areas or invalid_regions:
        print("❌ VALIDATION FAILURES:")
        
        if invalid_areas:
            print(f"Invalid Areas: {list(invalid_areas)}")
            print("Reason: These areas don't have matching PAC or PD entries")
            print("Looking for pattern: <member alias='AreaName' value='\"PAC - AreaName\"' /> OR <member alias='AreaName' value='\"PD - AreaName\"' />")
            print(f"Valid Areas found: {sorted(valid_areas)}")
            
            # Show what was actually found for failed areas
            for area in invalid_areas:
                found_entries = []
                for member in root.findall(".//member"):
                    if area.lower() in str(member.get('alias', '')).lower() or area.lower() in str(member.get('value', '')).lower():
                        found_entries.append(f"alias='{member.get('alias')}' value='{member.get('value')}'")
                for alias_elem in root.findall(".//alias"):
                    if area.lower() in str(alias_elem.get('key', '')).lower() or area.lower() in str(alias_elem.get('value', '')).lower():
                        found_entries.append(f"alias key='{alias_elem.get('key')}' value='{alias_elem.get('value')}'")
                
                if found_entries:
                    print(f"  '{area}' found as: {found_entries[:3]}")  # Show first 3 matches
                else:
                    print(f"  '{area}' not found anywhere in TWB")
        
        if invalid_regions:
            print(f"Invalid Regions: {list(invalid_regions)}")
            print(f"Valid Regions: {sorted(valid_regions)}")
        
        return False
    
    print("✅ All values validated")
    return True

def process_filter_file(filter_file_path, workbook_path, tableau_exe_path, output_dir):
    """
    Read filter.xlsx and process each row
    """
    # Read the Excel file
    df = pd.read_excel(filter_file_path)

    for index, row in df.iterrows():
        area = row['Area']
        region = row['Region']
        
        # Skip rows with NSW region temporarily
        if region == 'NSW':
            print(f"⏭️ Skipping {index + 1}/{len(df)}: {area} - {region} (NSW region)")
            continue
        
        print(f"\n🔄 Processing {index + 1}/{len(df)}: {area} - {region}")

    if validate_filter_data(filter_file_path, workbook_path) is False:
        print("❌ Validation failed. Please fix the filter.xlsx file.")
        return
    
    print(f"Found {len(df)} rows to process:")
    print(df.to_string(index=False))
    print("="*50)
    
    # Process each row
    for index, row in df.iterrows():
        area = row['Area']
        region = row['Region']
        
        print(f"\n🔄 Processing {index + 1}/{len(df)}: {area} - {region}")
        print("-" * 40)
        
        # Step 1: Change the parameter in the workbook
        print(f"Step 1: Changing Selected PAC parameter to {area}...")
        success = change_selected_pac(workbook_path, area)
        
        if not success:
            print(f"❌ Failed to change parameter for {area}. Skipping...")
            continue
        
        # Step 2: Export PDF
        print(f"Step 2: Exporting PDF for {area} - {region}...")
        try:
            export_single_pdf(workbook_path, tableau_exe_path, output_dir, area, region)
        except Exception as e:
            print(f"❌ Error exporting {area}: {e}")
            continue
        
        print(f"✅ Completed: {area} - {region}")
    
    print("\n" + "="*50)
    print("🎉 ALL REPORTS GENERATED!")
    print(f"📁 Check your output folder: {output_dir}")
    print("="*50)

# Main execution
if __name__ == "__main__":
    # Configuration - UPDATE THESE PATHS
    filter_file = r"C:\Users\CL-11\OneDrive - Lonergan Research\Repos\tableau-automation\filter.xlsx"
    workbook_file = r"C:\Users\CL-11\OneDrive - Lonergan Research\Repos\tableau-automation\walkin\NSW Police Service Assessment Walk-in Auburn;South West Metro.twb"
    tableau_executable = r"C:\Program Files\Tableau\Tableau 2024.3\bin\tableau.exe"
    #output_directory = r"C:\Users\CL-11\OneDrive - Lonergan Research\Repos\tableau-automation\output"
    output_directory = r"output"
    
    
    print("BATCH TABLEAU AUTOMATION")
    print("=" * 50)
    print(f"Filter file: {filter_file}")
    print(f"Workbook: {workbook_file}")
    print(f"Output: {output_directory}")
    print()
    
    # Install pandas if needed: pip install pandas openpyxl
    
    try:
        process_filter_file(filter_file, workbook_file, tableau_executable, output_directory)
    except Exception as e:
        print(f"❌ Critical error: {e}")
        print("Make sure:")
        print("- filter.xlsx exists and has 'Area' and 'Region' columns")
        print("- All file paths are correct") 
        print("- pandas and openpyxl are installed: pip install pandas openpyxl")