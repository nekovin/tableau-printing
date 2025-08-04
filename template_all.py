import xml.etree.ElementTree as ET
import shutil
from pathlib import Path
import os

def extract_pac_names_from_pdfs():
    """Extract PAC names from PDF files in Export directories"""
    pac_names = set()
    region_names = set()
    
    # Define base export directory
    export_base = Path("Export")
    
    # Define known region names to avoid treating them as PACs
    known_regions = {
        "Central Metro", "North West Metro", "South West Metro", 
        "Northern", "Southern", "Western"
    }
    
    # Check both Walk In and Telephone directories
    for service_type in ["Walk In", "Telephone"]:
        service_dir = export_base / service_type
        if service_dir.exists():
            # First check for region-level PDFs in the main service directory
            for pdf_file in service_dir.glob("*.pdf"):
                pdf_name = pdf_file.stem
                
                # These are region-level files (e.g., "Central Metro - Telephone.pdf")
                if " - Walk In" in pdf_name or " - Walkin" in pdf_name:
                    region_name = pdf_name.replace(" - Walk In", "").replace(" - Walkin", "").strip()
                elif " - Telephone" in pdf_name:
                    region_name = pdf_name.replace(" - Telephone", "").replace("- Telephone", "").strip()
                else:
                    region_name = pdf_name.strip()
                
                # Clean up region name variations
                region_name = region_name.replace("South West Metro- Telephone", "South West Metro")
                region_name = region_name.replace("Southern West Metro", "South West Metro")
                
                if region_name and region_name not in ["", " "]:
                    region_names.add(region_name)
            
            # Then look in region subdirectories for PAC-level PDFs
            for region_dir in service_dir.iterdir():
                if region_dir.is_dir():
                    # Extract PACs from PDF files in region directories
                    for pdf_file in region_dir.glob("*.pdf"):
                        # Extract PAC name (everything before " - Walk In.pdf" or " - Telephone.pdf")
                        pdf_name = pdf_file.stem
                        
                        # Clean up the PAC name by removing service type suffixes and variations
                        pac_name = pdf_name
                        
                        # Remove common suffixes
                        suffixes_to_remove = [
                            " - Walk In",
                            " - Telephone", 
                            " - Walkin",
                            " -Walk In",
                            " - Central Metropolitan - Walkin",
                            "- Telephone",
                            " Telephone",
                            " - old",
                            "_WRONG"
                        ]
                        
                        for suffix in suffixes_to_remove:
                            if pac_name.endswith(suffix):
                                pac_name = pac_name[:-len(suffix)]
                                break
                        
                        # Handle specific naming inconsistencies and PDF to Tableau name mappings
                        pac_name = pac_name.replace("Campelltown", "Campbelltown")
                        pac_name = pac_name.replace("Nepan", "Nepean")
                        pac_name = pac_name.replace("Parammatta", "Parramatta")
                        pac_name = pac_name.replace("Hunter Valley- Telephone", "Hunter Valley")
                        pac_name = pac_name.replace("Northern Beaches- Telephone", "Northern Beaches")
                        pac_name = pac_name.replace("Riverstone Telephone", "Riverstone")
                        
                        # Map PDF names to exact Tableau data names
                        if pac_name == "Liverpool":
                            pac_name = "Liverpool City"
                        elif pac_name == "Campbelltown":
                            pac_name = "Campbelltown City"
                        elif pac_name == "Fairfield":
                            pac_name = "Fairfield City"
                        elif pac_name == "Tweed and Byron":
                            pac_name = "Tweed/Byron"
                        elif pac_name == "TweedByron":
                            pac_name = "Tweed/Byron"
                        elif pac_name == "Port Stephens and Hunter":
                            pac_name = "Port Stephens/Hunter"
                        elif pac_name == "Port StephensHunter":
                            pac_name = "Port Stephens/Hunter"
                        elif pac_name == "CoffsClarence":
                            pac_name = "Coffs/Clarence"
                        
                        # Clean up any extra formatting
                        pac_name = pac_name.strip()
                        
                        # Only add if it's not empty and not a known region name
                        if pac_name and pac_name not in ["", " "] and pac_name not in known_regions:
                            pac_names.add(pac_name)
    
    return sorted(list(pac_names)), sorted(list(region_names))

def create_region_templates(source_twb_path, region_names):
    """Create template files for each region (comparing region to NSW)"""
    
    # Use the specific region template file as the base
    source_path = Path(source_twb_path)
    region_template_path = source_path.parent / f"{source_path.stem} Region.twb"
    
    if not region_template_path.exists():
        print(f"Warning: Region template not found at {region_template_path}")
        return
    
    # Map region display names to full region names
    region_display_to_full = {
        "Central Metro": "Central Metropolitan",
        "North West Metro": "North West Metropolitan",
        "South West Metro": "South West Metropolitan",
        "Northern": "Northern",
        "Southern": "Southern",
        "Western": "Western"
    }
    
    for region_display in region_names:
        # Get the full region name
        full_region = region_display_to_full.get(region_display, region_display)
        
        # Create output filename
        safe_region_name = region_display.replace(" ", "_")
        # Save to output directory with region subdirectory
        output_base_dir = source_path.parent.parent / "output"
        output_base_dir.mkdir(exist_ok=True)
        region_output_dir = output_base_dir / safe_region_name
        region_output_dir.mkdir(exist_ok=True)
        output_path = region_output_dir / f"{source_path.stem}_Region_{safe_region_name}.twb"

        print("Region Template Path:", output_path)
        
        # Copy the region template file (not the main template)
        shutil.copy2(region_template_path, output_path)
        
        # Parse and modify
        tree = ET.parse(output_path)
        root = tree.getroot()
        
        # Update Areas Region filter to match the specific region
        for filter_elem in root.findall(".//filter[@class='categorical']"):
            column_attr = filter_elem.get('column')
            if column_attr and 'Areas Region' in column_attr:
                groupfilter = filter_elem.find('groupfilter')
                if groupfilter is not None:
                    groupfilter.set('member', f'"{full_region}"')
        
        # Save
        tree.write(output_path, encoding='utf-8', xml_declaration=True)
        print(f"Created region template: {output_path}")

def create_pac_templates(source_twb_path):
    """Create template files for each PAC and region"""
    
    # Extract PAC names and region names from PDF files
    pac_names, region_names = extract_pac_names_from_pdfs()
    print(f"Found {len(pac_names)} PACs: {pac_names}")
    print(f"Found {len(region_names)} Regions: {region_names}")
    
    # Create region templates first
    create_region_templates(source_twb_path, region_names)
    
    # Then create PAC templates
    
    # Define region mapping for PACs based on exact names from original template
    pac_to_region = {
        # Central Metro PACs (use PAC prefix)
        "Eastern Beaches": "Central Metropolitan",
        "Eastern Suburbs": "Central Metropolitan", 
        "Inner West": "Central Metropolitan",
        "Kings Cross": "Central Metropolitan",
        "Leichhardt": "Central Metropolitan",
        "South Sydney": "Central Metropolitan",
        "St George": "Central Metropolitan",
        "Surry Hills": "Central Metropolitan",
        "Sutherland Shire": "Central Metropolitan",
        "Sydney City": "Central Metropolitan",
        
        # North West Metro PACs (use PAC prefix)
        "Blacktown": "North West Metropolitan",
        "Blue Mountains": "North West Metropolitan",
        "Hawkesbury": "North West Metropolitan",
        "Kuring-gai": "North West Metropolitan",
        "Mount Druitt": "North West Metropolitan",
        "Nepean": "North West Metropolitan",
        "North Shore": "North West Metropolitan",
        "Northern Beaches": "North West Metropolitan",
        "Parramatta": "North West Metropolitan",
        "Riverstone": "North West Metropolitan",
        "Ryde": "North West Metropolitan",
        "The Hills": "North West Metropolitan",
        
        # South West Metro PACs (use PAC prefix)
        "Auburn": "South West Metropolitan",
        "Bankstown": "South West Metropolitan",
        "Burwood": "South West Metropolitan",
        "Camden": "South West Metropolitan",
        "Campbelltown City": "South West Metropolitan",
        "Campsie": "South West Metropolitan",
        "Cumberland": "South West Metropolitan",
        "Fairfield City": "South West Metropolitan",
        "Liverpool City": "South West Metropolitan",
        
        # Northern PACs (use PD prefix)
        "Brisbane Water": "Northern",
        "Coffs/Clarence": "Northern",
        "Hunter Valley": "Northern",
        "Lake Macquarie": "Northern",
        "Manning Great Lakes": "Northern",
        "Mid North Coast": "Northern",
        "Newcastle City": "Northern",
        "Port Stephens/Hunter": "Northern",
        "Richmond": "Northern",
        "Tuggerah Lakes": "Northern",
        "Tweed/Byron": "Northern",
        
        # Southern PACs (use PD prefix)
        "Lake Illawarra": "Southern",
        "Monaro": "Southern",
        "Murray River": "Southern",
        "Murrumbidgee": "Southern",
        "Riverina": "Southern",
        "South Coast": "Southern",
        "The Hume": "Southern",
        "Wollongong": "Southern",
        
        # Western PACs (use PD prefix)
        "Barrier": "Western",
        "Central North": "Western",
        "Central West": "Western",
        "Chifley": "Western",
        "New England": "Western",
        "Orana Mid-Western": "Western",
        "Oxley": "Western",
    }
    
    regions = [
        "Central Metropolitan",
        "Northern", 
        "North West Metropolitan",
        "Southern",
        "South West Metropolitan", 
        "Western"
    ]
    
    source_path = Path(source_twb_path)
    
    for pac_name in pac_names:
        # Determine which region this PAC belongs to
        region = pac_to_region.get(pac_name, "Unknown Region")
        if region == "Unknown Region":
            print(f"Warning: No region mapping found for PAC '{pac_name}', skipping...")
            continue
            
        # Create output filename
        safe_pac_name = pac_name.replace(" ", "_").replace("&", "and").replace("/", "_")
        
        # Determine which region directory this PAC belongs to
        region_dir_mapping = {
            "Central Metropolitan": "Central_Metro",
            "North West Metropolitan": "North_West_Metro",
            "South West Metropolitan": "South_West_Metro",
            "Northern": "Northern",
            "Southern": "Southern",
            "Western": "Western"
        }
        
        region_dir_name = region_dir_mapping.get(region, region.replace(" ", "_"))
        
        # Save to output directory with region subdirectory
        output_base_dir = source_path.parent.parent / "output"
        output_base_dir.mkdir(exist_ok=True)
        region_output_dir = output_base_dir / region_dir_name
        region_output_dir.mkdir(exist_ok=True)
        output_path = region_output_dir / f"{source_path.stem}_{safe_pac_name}.twb"

        print("Path:", output_path)
        
        # Copy source file
        shutil.copy2(source_path, output_path)
        
        # Parse and modify
        tree = ET.parse(output_path)
        root = tree.getroot()
        
        # Update Areas Region filter to match the PAC's region
        for filter_elem in root.findall(".//filter[@class='categorical']"):
            column_attr = filter_elem.get('column')
            if column_attr and 'Areas Region' in column_attr:
                groupfilter = filter_elem.find('groupfilter')
                if groupfilter is not None:
                    groupfilter.set('member', f'"{region}"')
        
        # Determine the correct prefix based on region
        # Metropolitan areas use "PAC - " prefix, regional areas use "PD - " prefix
        if region in ["Central Metropolitan", "North West Metropolitan", "South West Metropolitan"]:
            prefix = "PAC"
        else:
            prefix = "PD"
        
        # Update Areas PAC filter to the specific PAC
        for filter_elem in root.findall(".//filter[@class='categorical']"):
            column_attr = filter_elem.get('column')
            if column_attr and 'Areas PAC' in column_attr:
                groupfilter = filter_elem.find('groupfilter')
                if groupfilter is not None:
                    groupfilter.set('member', f'"{prefix} - {pac_name}"')
        
        # Get all PACs for the region for populating the parameter domain
        region_pacs = [pac for pac, reg in pac_to_region.items() if reg == region]
        
        # Determine the correct prefix based on region
        # Metropolitan areas use "PAC - " prefix, regional areas use "PD - " prefix
        if region in ["Central Metropolitan", "North West Metropolitan", "South West Metropolitan"]:
            prefix = "PAC"
        else:
            prefix = "PD"
        
        # Update Parameter 1 (Selected PAC) columns
        for column_elem in root.findall(".//column[@name='[Parameter 1]']"):
            # Update the default value to the current PAC
            column_elem.set('value', f'"{prefix} - {pac_name}"')
            
            # Update the alias (display name) 
            column_elem.set('alias', pac_name)
            
            # Update the calculation formula
            calculation_elem = column_elem.find('calculation')
            if calculation_elem is not None:
                calculation_elem.set('formula', f'"{prefix} - {pac_name}"')
            
            # Update aliases - include all PACs in this region
            aliases_elem = column_elem.find('aliases')
            if aliases_elem is not None:
                aliases_elem.clear()
                # Add all PACs in the region to the aliases
                for region_pac in sorted(region_pacs):
                    alias_elem = ET.SubElement(aliases_elem, 'alias')
                    alias_elem.set('key', f'"{prefix} - {region_pac}"')
                    alias_elem.set('value', region_pac)
            
            # Update members - include all PACs in this region  
            members_elem = column_elem.find('members')
            if members_elem is not None:
                members_elem.clear()
                # Add all PACs in the region to the members
                for region_pac in sorted(region_pacs):
                    member_elem = ET.SubElement(members_elem, 'member')
                    member_elem.set('alias', region_pac)
                    member_elem.set('value', f'"{prefix} - {region_pac}"')
        
        # Update all region tuple values to match the PAC's region
        for value_elem in root.findall(".//value"):
            if value_elem.text and ("Metropolitan" in value_elem.text or value_elem.text in ['"Northern"', '"Southern"', '"Western"']):
                value_elem.text = f'"{region}"'
        
        # Update Parameter 2 (Selected Region) columns to match the PAC's region
        region_aliases = {
            "Central Metropolitan": "Central Metro",
            "North West Metropolitan": "North West Metro", 
            "South West Metropolitan": "South West Metro",
            "Northern": "Northern",
            "Southern": "Southern",
            "Western": "Western"
        }
        
        # Get all regions for the dropdown
        all_regions = list(region_aliases.keys())
        
        for column_elem in root.findall(".//column[@name='[Parameter 2]']"):
            # Update the default value to the PAC's region
            column_elem.set('value', f'"{region}"')
            
            # Update the alias (display name)
            if region in region_aliases:
                column_elem.set('alias', region_aliases[region])
                display_region = region_aliases[region]
            else:
                column_elem.set('alias', region)
                display_region = region
            
            # Update the calculation formula
            calculation_elem = column_elem.find('calculation')
            if calculation_elem is not None:
                calculation_elem.set('formula', f'"{region}"')
            
            # Update aliases - include all regions
            aliases_elem = column_elem.find('aliases')
            if aliases_elem is not None:
                aliases_elem.clear()
                # Add all regions to the aliases
                for reg in sorted(all_regions):
                    alias_elem = ET.SubElement(aliases_elem, 'alias')
                    alias_elem.set('key', f'"{reg}"')
                    alias_elem.set('value', region_aliases.get(reg, reg))
            
            # Update members - include all regions
            members_elem = column_elem.find('members')
            if members_elem is not None:
                members_elem.clear()
                # Add all regions to the members
                for reg in sorted(all_regions):
                    member_elem = ET.SubElement(members_elem, 'member')
                    member_elem.set('alias', region_aliases.get(reg, reg))
                    member_elem.set('value', f'"{reg}"')
        
        # Save
        tree.write(output_path, encoding='utf-8', xml_declaration=True)
        print(f"Created: {output_path}")

# Usage
source_file = r"telephone\templates\NSW Police Service Assessment Telephone.twb"
create_pac_templates(source_file)

source_file = r"walkin\templates\NSW Police Service Assessment Walk-in.twb"
create_pac_templates(source_file)