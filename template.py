import xml.etree.ElementTree as ET
import shutil
from pathlib import Path

def create_region_templates(source_twb_path):
    """Create 6 template files with different regions"""
    
    regions = [
        "Central Metropolitan",
        "Northern", 
        "North West Metropolitan",
        "Southern",
        "South West Metropolitan", 
        "Western"
    ]
    
    source_path = Path(source_twb_path)
    
    for region in regions:
        # Create output filename
        safe_region = region.replace(" ", "_")
        # Save to output directory
        output_dir = source_path.parent.parent / "output"
        output_dir.mkdir(exist_ok=True)  # Create output folder if it doesn't exist
        output_path = output_dir / f"{source_path.stem}_{safe_region}.twb"

        print("Path:", output_path)
        
        # Copy source file
        shutil.copy2(source_path, output_path)
        
        # Parse and modify
        tree = ET.parse(output_path)
        root = tree.getroot()
        
        # Update Areas Region filter
        for filter_elem in root.findall(".//filter[@class='categorical']"):
            column_attr = filter_elem.get('column')
            if column_attr and 'Areas Region' in column_attr:
                groupfilter = filter_elem.find('groupfilter')
                if groupfilter is not None:
                    groupfilter.set('member', f'"{region}"')
        
        # Update all region tuple values
        for value_elem in root.findall(".//value"):
            if value_elem.text and "Metropolitan" in value_elem.text or value_elem.text in ['"Northern"', '"Southern"', '"Western"']:
                value_elem.text = f'"{region}"'
        
        # Update Parameter 2 (Selected Region) columns
        region_aliases = {
            "Central Metropolitan": "Central Metro",
            "North West Metropolitan": "North West Metro", 
            "South West Metropolitan": "South West Metro",
            "Northern": "Northern",
            "Southern": "Southern",
            "Western": "Western"
        }
        
        for column_elem in root.findall(".//column[@name='[Parameter 2]']"):
            if column_elem.get('caption') == 'Selected Region':
                # Update the default value
                column_elem.set('value', f'"{region}"')
                
                # Update the alias (display name)
                if region in region_aliases:
                    column_elem.set('alias', region_aliases[region])
                else:
                    column_elem.set('alias', region)
                
                # Update the calculation formula
                calculation_elem = column_elem.find('calculation')
                if calculation_elem is not None:
                    calculation_elem.set('formula', f'"{region}"')
        
        # Save
        tree.write(output_path, encoding='utf-8', xml_declaration=True)
        print(f"Created: {output_path}")

        break

# Usage
source_file = r"telephone\templates\NSW Police Service Assessment Telephone.twb"
#source_file = r"walkin\NSW Police Service Assessment Walk-in.twb"
create_region_templates(source_file)