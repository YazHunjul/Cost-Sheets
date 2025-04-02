import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.worksheet.datavalidation import DataValidation
from datetime import datetime
import shutil
import os
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docxtpl import DocxTemplate
import math
import zipfile

# Add this dictionary at the top of the file, after imports
contacts = {
    'Marc Byford': '(07974 403322)',
    'Karl Nicholson': '(07791 397866)',
    'Dan Butler': '(07703 729686)',
    'Chris Mannus': '(07870 263280)',
    'Dean Griffiths': '(07814 784352)',
    'Kent Phillips': '(07949 016501)'
}

estimators = {
    'Simon Still': 'Lead Estimator',
    'Nick Soton': 'Estimator',
    'Chris Davis': 'Estimator'
}

# Add this at the top of the file with other constants
VALID_CANOPY_MODELS = [
    'KVF', 'UVF', 'CMWF', 'CMWI', 'KVI', 'UVI', 'CXW', 'KVS', 'UVS', 'KSW'
]  # Add any other valid canopy models

# Add this function after the imports
def create_project_sidebar(project_data):
    """Create a sidebar showing project progress and structure"""
    with st.sidebar:
        st.header("📋 Project Overview")
        
        # Show General Info Status
        st.subheader("General Information")
        
        # Check required fields
        required_fields = {
            'Project Name': project_data.get('Project Name', ''),
            'Project Number': project_data.get('Project Number', ''),
            'Customer': project_data.get('Customer', ''),
            'Company': project_data.get('Company', ''),
            'Location': project_data.get('Location', ''),
            'Address': project_data.get('Address', ''),
            'Sales Contact': project_data.get('Sales Contact', ''),
            'Estimator': project_data.get('Estimator', '')
        }
        
        # Show status of each required field
        for field, value in required_fields.items():
            if value and value != "Select...":
                st.markdown(f"✅ {field}")
            else:
                st.markdown(f"❌ {field} *required*")
        
        # Show Levels and Areas
        st.markdown("---")
        st.subheader("Project Structure")
        
        levels = project_data.get('Levels', [])
        if not levels:
            st.info("No levels added yet")
            return
            
        for level in levels:
            if isinstance(level, dict):  # Make sure level is a dictionary
                level_name = level.get('level_name', '')
                areas = level.get('areas', [])
                
                if level_name:
                    st.markdown(f"### 🏢 {level_name}")
                    
                    if not areas:
                        st.markdown("- *No areas added*")
                        continue
                        
                    for area in areas:
                        if isinstance(area, dict):  # Make sure area is a dictionary
                            area_name = area.get('area_name', '')
                            canopies = area.get('canopies', [])
                            
                            if area_name:
                                st.markdown(f"#### 📍 {area_name}")
                                
                                if not canopies:
                                    st.markdown("- *No canopies added*")
                                    continue
                                    
                                valid_canopies = [
                                    canopy for canopy in canopies 
                                    if isinstance(canopy, dict) and  # Make sure canopy is a dictionary
                                    canopy.get('reference_number') and 
                                    canopy.get('model') != "Select..." and 
                                    canopy.get('configuration') != "Select..."
                                ]
                                
                                if valid_canopies:
                                    for canopy in valid_canopies:
                                        ref = canopy.get('reference_number', '')
                                        model = canopy.get('model', '')
                                        config = canopy.get('configuration', '')
                                        st.markdown(f"- 🔹 {ref}: {model} ({config})")
                                        
                                        # Show wall cladding if present
                                        if (isinstance(canopy.get('wall_cladding'), dict) and 
                                            canopy['wall_cladding'].get('type') and 
                                            canopy['wall_cladding']['type'] != "Select..."):
                                            positions = '/'.join(canopy['wall_cladding'].get('positions', []))
                                            st.markdown(f"  - 🧱 Wall Cladding: {positions}")
                                else:
                                    st.markdown("- *No valid canopies configured*")

def create_general_info_form():
    # Initialize session state if needed
    if 'project_data' not in st.session_state:
        st.session_state.project_data = {
            'Project Name': '',
            'Project Number': '',
            'Customer': '',
            'Company': '',
            'Location': '',
            'Address': '',
            'Sales Contact': '',
            'Estimator': '',
            'Levels': []
        }
    
    st.title("📋 Project Information")
    
    # General Information Section
    col1, col2, col3 = st.columns(3)
    
    with col1:
        project_name = st.text_input("Project Name")
        st.session_state.project_data['Project Name'] = project_name
    with col2:
        project_number = st.text_input("Project Number")
        st.session_state.project_data['Project Number'] = project_number
    with col3:
        date = st.date_input("Date", datetime.now())
        st.session_state.project_data['Date'] = date.strftime("%d/%m/%Y")
    
    col4, col5, col6 = st.columns(3)
    
    with col4:
        customer = st.text_input("Customer")
        st.session_state.project_data['Customer'] = customer
    with col5:
        company = st.text_input("Company")
        st.session_state.project_data['Company'] = company
    with col6:
        location = st.text_input("Location")
        st.session_state.project_data['Location'] = location
    
    col7, col8, col9 = st.columns(3)
    
    with col7:
        address = st.text_input("Address")
        st.session_state.project_data['Address'] = address
    with col8:
        sales_contact = st.selectbox("Sales Contact", ["Select..."] + list(contacts.keys()))
        st.session_state.project_data['Sales Contact'] = sales_contact
    with col9:
        estimator = st.selectbox("Estimator", ["Select..."] + list(estimators.keys()))
        st.session_state.project_data['Estimator'] = estimator
    
    cost_sheet = st.selectbox("Select Cost Sheet to Report", ["Select...", "Canopy", "Other Options"])
    st.session_state.project_data['Cost Sheet'] = cost_sheet
    
    # Levels Configuration
    if cost_sheet != "Select...":
        st.markdown("---")
        st.subheader("Levels Configuration")
        
        num_levels = st.number_input("Enter Number of Levels", min_value=1, value=1, step=1)
        
        levels_data = []
        
        for level_idx in range(num_levels):
            with st.expander(f"Level {level_idx + 1}", expanded=False):
                level_name = st.text_input(f"Enter Level {level_idx + 1} Name", 
                                         key=f"level_name_{level_idx}")
                
                if level_name:
                    num_areas = st.number_input(
                        f"Enter the number of areas in {level_name}",
                        min_value=1,
                        value=1,
                        step=1,
                        key=f"num_areas_{level_idx}"
                    )
                    
                    areas_data = []
                    
                    for area_idx in range(num_areas):
                        area_name = st.text_input(
                            f"Enter area {area_idx + 1} Name",
                            key=f"area_name_{level_idx}_{area_idx}"
                        )
                        
                        if area_name:
                            # Add UV-C Control Schedule radio button for the area
                            include_uvc = st.radio(
                                "Include UV-C Control Schedule",
                                options=["No", "Yes"],
                                key=f"uvc_{level_idx}_{area_idx}"
                            )
                            
                            num_canopies = st.number_input(
                                f"Enter Number of Canopies in {area_name}",
                                min_value=1,
                                value=1,
                                step=1,
                                key=f"num_canopies_{level_idx}_{area_idx}"
                            )
                            
                            canopies_data = []
                            
                            for canopy_idx in range(num_canopies):
                                st.write(f"Processing canopy {canopy_idx + 1} for {area_name}")
                                col1, col2 = st.columns(2)
                                with col1:
                                    reference_number = st.text_input(
                                        "Reference Number",
                                        key=f"ref_{level_idx}_{area_idx}_{canopy_idx}"
                                    )
                                    configuration = st.selectbox(
                                        "Configuration",
                                        options=["Select...", "WALL", "ISLAND", "OTHER"],
                                        key=f"config_{level_idx}_{area_idx}_{canopy_idx}"
                                    )
                                    # Add fire suppression radio button right after configuration
                                    fire_suppression = st.radio(
                                        "Include Fire Suppression",
                                        options=["No", "Yes"],
                                        key=f"fire_suppression_{level_idx}_{area_idx}_{canopy_idx}"
                                    )
                                
                                with col2:
                                    model = st.selectbox(
                                        "Model",
                                        options=["Select..."] + VALID_CANOPY_MODELS,
                                        key=f"model_{level_idx}_{area_idx}_{canopy_idx}"
                                    )
                                    
                                    wall_cladding = st.selectbox(
                                        "Wall Cladding",
                                        options=["Select...", "2M² (HFL)"],
                                        key=f"cladding_{level_idx}_{area_idx}_{canopy_idx}"
                                    )
                                    
                                    if wall_cladding != "Select...":
                                        col1, col2 = st.columns(2)
                                        with col1:
                                            cladding_width = st.number_input(
                                                "Cladding Width (mm)",
                                                min_value=0,
                                                value=0,
                                                step=100,
                                                key=f"cladding_width_{level_idx}_{area_idx}_{canopy_idx}"
                                            )
                                        with col2:
                                            cladding_height = st.number_input(
                                                "Cladding Height (mm)",
                                                min_value=0,
                                                value=0,
                                                step=100,
                                                key=f"cladding_height_{level_idx}_{area_idx}_{canopy_idx}"
                                            )
                                        
                                        wall_positions = st.multiselect(
                                            "Wall Positions",
                                            options=["Rear", "Left", "Right"],
                                            key=f"wall_positions_{level_idx}_{area_idx}_{canopy_idx}"
                                        )
                                
                                if reference_number:
                                    canopy_data = {
                                        'reference_number': reference_number,
                                        'configuration': configuration,
                                        'model': model,
                                        'fire_suppression': fire_suppression == "Yes",  # Store as boolean
                                        'level_name': level_name,  # Store level name
                                        'area_name': area_name,    # Store area name
                                        'wall_cladding': {
                                            'type': wall_cladding,
                                            'width': cladding_width if wall_cladding != "Select..." else 0,
                                            'height': cladding_height if wall_cladding != "Select..." else 0,
                                            'positions': wall_positions if wall_cladding != "Select..." else []
                                        }
                                    }
                                    canopies_data.append(canopy_data)
                                
                                # Add MUA VOL input for F-type canopies
                                if 'F' in str(model).upper():
                                    mua_vol = st.text_input(
                                        "MUA Volume (m³/h)",
                                        key=f"mua_vol_{level_idx}_{area_idx}_{canopy_idx}"
                                    )
                                    if canopy_data:
                                        canopy_data['mua_vol'] = mua_vol
                            
                            if canopies_data:
                                areas_data.append({
                                    'area_name': area_name,
                                    'canopies': canopies_data,
                                    'include_uvc': include_uvc == "Yes"  # Store UV-C Control Schedule preference
                                })
                    
                    if areas_data:
                        levels_data.append({
                            'level_name': level_name,
                            'areas': areas_data
                        })
        
        # Update session state with levels data
        st.session_state.project_data['Levels'] = levels_data
        
        # Update sidebar with current data
        create_project_sidebar(st.session_state.project_data)
        
        # Add save button and check if all required fields are filled
        required_fields = {
            'Project Name': st.session_state.project_data.get('Project Name', ''),
            'Project Number': st.session_state.project_data.get('Project Number', ''),
            'Customer': st.session_state.project_data.get('Customer', ''),
            'Company': st.session_state.project_data.get('Company', ''),
            'Location': st.session_state.project_data.get('Location', ''),
            'Address': st.session_state.project_data.get('Address', ''),
            'Sales Contact': st.session_state.project_data.get('Sales Contact', ''),
            'Estimator': st.session_state.project_data.get('Estimator', '')
        }
        
        # Check if all required fields are filled and at least one canopy is added
        all_fields_filled = all(value and value != "Select..." for value in required_fields.values())
        has_canopies = any(level.get('areas', []) for level in st.session_state.project_data.get('Levels', []))
        
        st.markdown("---")
        col1, col2 = st.columns(2)
        
        with col1:
            if st.button("Save Project", key="save_project_button", disabled=not (all_fields_filled and has_canopies)):
                try:
                    save_to_excel(st.session_state.project_data)
                except Exception as e:
                    st.error(f"Error saving project: {str(e)}")
        
        with col2:
            # Add upload section
            create_upload_section(col2)

def copy_template_sheet(workbook, source_sheet_name, new_sheet_name):
    """Copy a sheet and rename it"""
    source = workbook[source_sheet_name]
    target = workbook.copy_worksheet(source)
    target.title = new_sheet_name
    return target

def get_initials(name):
    """Extract initials from a name, skipping 'Select...'"""
    if name == "Select...":
        return ""
    # Split the name and take first letter of each part
    words = name.split()
    initials = ''.join(word[0].upper() for word in words if word)
    return initials

def write_to_sheet(sheet, data, level_name, area_name, canopies, fs_sheet=None, is_edge_box=False):
    """Write data to specific cells in the sheet"""
    # Write sheet title
    sheet_title = f"{level_name} - {area_name}"
    if is_edge_box:
        sheet['C1'] = area_name  # Write just area name to C1 for EDGE BOX sheets
        # Write general info in columns D and H
        sheet['D3'] = data['Project Number']
        customer_company = f"{data['Customer']} ({data['Company']})" if data['Company'] else data['Customer']
        sheet['D5'] = customer_company
        sheet['D7'] = f"{get_initials(data['Sales Contact'])}/{get_initials(data['Estimator'])}"
        sheet['H3'] = data['Project Name']
        sheet['H5'] = data['Location']
        sheet['H7'] = data['Date']
        return  # Don't write any other data to EDGE BOX sheets
    else:
        sheet['B1'] = sheet_title  # Write to B1 for other sheets
    
    # Write general info to sheets
    def write_general_info(target_sheet, is_edge_box=False):
        # Skip writing to EDGE BOX sheets
        if is_edge_box:
            return
            
        # For regular sheets
        target_sheet['C3'] = data['Project Number']
        customer_company = f"{data['Customer']} ({data['Company']})" if data['Company'] else data['Customer']
        target_sheet['C5'] = customer_company
        target_sheet['C7'] = f"{get_initials(data['Sales Contact'])}/{get_initials(data['Estimator'])}"
        target_sheet['G3'] = data['Project Name']
        target_sheet['G5'] = data['Location']
        target_sheet['G7'] = data['Date']
        target_sheet['O7'] = 'A'  # Initial revision
    
    write_general_info(sheet, is_edge_box)
    
    # Handle fire suppression sheet if needed
    if fs_sheet:
        # Write the same title as the canopy sheet
        fs_sheet['B1'] = sheet_title
        
        # Get the sheet number from the canopy sheet name
        sheet_number = sheet.title.split('(')[-1].strip(')')
        
        # Rename fire suppression sheet to match canopy sheet format
        fs_sheet.title = f"FIRE SUPP - {level_name} ({sheet_number})"
        
        # Write general info
        write_general_info(fs_sheet)
        
        # Write fire suppression canopies starting from first row
        fs_row = 12  # Start at first row
        for canopy in canopies:
            if canopy.get('fire_suppression', False):
                fs_sheet[f'B{fs_row}'] = canopy['reference_number']
                fs_row += 17  # Move to next row
    
    # Write canopy data
    for idx, canopy in enumerate(canopies):
        base_row = 12 + (idx * 17)
        
        # Write to sheet
        if canopy['reference_number'] != "Select...":
            sheet[f'B{base_row}'] = canopy['reference_number']
        if canopy['configuration'] != "Select...":
            sheet[f'C{base_row + 2}'] = canopy['configuration']
        if canopy['model'] != "Select...":
            sheet[f'D{base_row + 2}'] = canopy['model']
        
        # Write standard entries one by one
        sheet[f'C{base_row + 3}'] = "LIGHT SELECTION"
        sheet[f'C{base_row + 4}'] = "SELECT WORKS"
        sheet[f'C{base_row + 5}'] = "SELECT WORKS"
        sheet[f'C{base_row + 6}'] = "BIM/ REVIT per CANOPY"
        sheet[f'D{base_row + 6}'] = "1"
        
        # Wall cladding - only write if not "Select..."
        if canopy['wall_cladding']['type'] and canopy['wall_cladding']['type'] != "Select...":
            cladding_row = base_row + 7
            # Write wall cladding type to C column
            sheet[f'C{cladding_row}'] = canopy['wall_cladding']['type']
            
            # Write dimensions and positions to hidden cells for storage
            sheet[f'Q{cladding_row}'] = canopy['wall_cladding']['width']
            sheet[f'R{cladding_row}'] = canopy['wall_cladding']['height']
            sheet[f'S{cladding_row}'] = ','.join(canopy['wall_cladding']['positions'])
            sheet[f'T{cladding_row}'] = f"2M² (HFL) - {canopy['wall_cladding']['width']}x{canopy['wall_cladding']['height']}mm ({'/'.join(canopy['wall_cladding']['positions'])})"
        
        # Write MUA VOL to Q22 if it's an F-type canopy
        model = str(canopy.get('model', ''))  # Convert to string first
        if 'F' in model.upper():
            print(f"Writing MUA VOL for {model}: {canopy.get('mua_vol', '-')}")
            sheet[f'Q{22 + (idx * 17)}'] = f"MUA VOL: {canopy.get('mua_vol', '-')}"
    
    # Get fire suppression canopies for data structure
    fire_suppression_canopies = [c for c in canopies if c.get('fire_suppression', False)]

def add_dropdowns_to_sheet(wb, sheet, start_row):
    """Add data validation (dropdowns) to specific cells"""
    if 'Lists' not in wb.sheetnames:
        list_sheet = wb.create_sheet('Lists')
    else:
        list_sheet = wb['Lists']

    # Define all dropdown options
    dropdowns = {
        'lights': {
            'options': [
                'LED STRIP L6 Inc DALI',
                'LED STRIP L12 inc DALI', 
                'LED STRIP L18 Inc DALI',
                'Small LED Spots inc DALI',
                'LARGE LED Spots inc DALI'
            ],
            'column': 'A',
            'target_col': 'C',
            'row_offset': 3  # C15 for first canopy
        },
        'special_works_1': {
            'options': [
                'ROUND CORNERS',
                'CUT OUT',
                'CASTELLE LOCKING',
                'HEADER DUCT S/S',
                'HEADER DUCT',
                'PAINT FINSH',
                'UV ON DEMAND',
                'E/over for emergency strip light',
                'E/over for small emer. spot light',
                'E/over for large emer. spot light',
                'COLD MIST ON DEMAND',
                'CMW  PIPEWORK HWS/CWS',
                'CANOPY GROUND SUPPORT',
                '2nd EXTRACT PLENUM',
                'SUPPLY AIR PLENUM',
                'CAPTUREJET PLENUM',
                'COALESCER'
            ],
            'column': 'B',
            'target_col': 'C',
            'row_offset': 4  # C16 for first special work
        },
        'special_works_2': {
            'options': [
                'SELECT WORKS',  # Add default option
                'ROUND CORNERS',
                'CUT OUT',
                'CASTELLE LOCKING',
                'HEADER DUCT S/S',
                'HEADER DUCT',
                'PAINT FINSH',
                'UV ON DEMAND',
                'E/over for emergency strip light',
                'E/over for small emer. spot light',
                'E/over for large emer. spot light',
                'COLD MIST ON DEMAND',
                'CMW  PIPEWORK HWS/CWS',
                'CANOPY GROUND SUPPORT',
                '2nd EXTRACT PLENUM',
                'SUPPLY AIR PLENUM',
                'CAPTUREJET PLENUM',
                'COALESCER'
            ],
            'column': 'B',  # Use same column as they're the same options
            'target_col': 'C',
            'row_offset': 5  # C17 for second special work
        },
        'wall_cladding': {
            'options': ['', '2M² (HFL)'],
            'column': 'C',
            'target_col': 'C',
            'row_offset': 7  # C19 for first canopy
        },
        'control_panel': {
            'options': ['CP1S', 'CP2S', 'CP3S', 'CP4S'],
            'column': 'D',
            'target_col': 'C',
            'row_offset': 13  # C25 for first canopy
        },
        'ww_pods': {
            'options': [
                '1000-S', '1500-S', '2000-S', '2500-S', '3000-S',
                '1000-D', '1500-D', '2000-D', '2500-D', '3000-D'
            ],
            'column': 'E',
            'target_col': 'C',
            'row_offset': 14  # C26 for first canopy
        },
        'delivery_location': {
            'options': [
                "",  # Empty option first
                "ABERDEEN 590",
                "ABINGDON 110",
                "ALDEBURGH 112",
                "ALDERSHOT 110",
                "ALNWICK 342",
                "ANDOVER 110",
                "ASHFORD 25",
                "AYLESBURY 86",
                "BANBURY 102",
                "BANGOR 324",
                "BARKING 32",
                "BARNET 55",
                "BARNSLEY 209",
                "BARNSTABLE 227",
                "BARROW-IN-FURNESS 348",
                "BASILDON 38",
                "BASINGSTOKE 82",
                "BATH 154",
                "BEDFORD 103",
                "BERWICK-UPON-TWEED 371",
                "BILLERICAY 37",
                "BIRKENHEAD 277",
                "BIRMINGHAM 168",
                "BLACKBURN 283",
                "BLACKPOOL 289",
                "BLANDFORD FORUM 144",
                "BODMIN 273",
                "BOGNOR REGIS 88",
                "BOLTON 259",
                "BOOTLE 272",
                "BOURNEMOUTH 140",
                "BRADFORD 234",
                "BRAINTREE 60",
                "BRIDGEND 205",
                "BRIDLINGTON 244",
                "BRIGHTON 68",
                "BRISTOL 157",
                "BUCKINGHAMSHIRE 109",
                "BURNLEY 296",
                "BURTON UPON TRENT 175",
                "BURY ST EDMUNDS 98",
                "CAMBRIDGE 85",
                "CANNOCK 175",
                "CANTERBURY 30",
                "CARDIFF 192",
                "CARLISLE 356",
                "CARMARTHEN 252",
                "CHELTENHAM 148",
                "CHESTER 268",
                "COVENTRY 146",
                "CHIPPENHAM 136",
                "COLCHESTER 78",
                "CORBY 128",
                "DARTMOUTH 245",
                "DERBY 178",
                "DONCASTER 203",
                "DORCHESTER 160",
                "DORKING 46",
                "DOVER 45",
                "DURHAM 299",
                "EASTBOURNE 57",
                "EASTLEIGH 109",
                "EDINBURGH 428",
                "ENFIELD 49",
                "EXETER 205",
                "EXMOUTH 207",
                "FELIXSTOWE 103",
                "GATWICK 44",
                "GLASGOW 456",
                "GLASTONBURY 164",
                "GLOUCESTER 151",
                "GRANTHAM 143",
                "GREAT YARMOUTH 147",
                "GRIMSBY 215",
                "GUILDFORD 59",
                "HARLOW 47",
                "HARROGATE 236",
                "HARTLEPOOL 286",
                "HASTINGS 40",
                "HEXHAM 325",
                "HEREFORD 184",
                "HIGH WYCOMBE 80",
                "HIGHBRIDGE 187",
                "HONITON 190",
                "HORSHAM 55",
                "HOUNSLOW 55",
                "HUDDERSFIELD 239",
                "HULL 247",
                "HUNTINGDON 94",
                "INVERNESS 619",
                "IPSWICH 94",
                "IRELAND",
                "KENDAL 321",
                "KETTERING 127",
                "KIDDERMINSTER 179",
                "KILMARNOCK 449",
                "KINGSTON UPON HULL 220",
                "KINGSTON UPON THAMES 52",
                "LANCASTER 290",
                "LAUNCESTON 251",
                "LEAMINGTON SPA 146",
                "LEEDS 231",
                "LEICESTER 151",
                "LEIGH ON SEA 45",
                "LEWISHAM 29",
                "LINCOLN 179",
                "LIVERPOOL 258",
                "LLANDUDNO 309",
                "LONDON in FORS GOLD(varies)",
                "LUTON 80",
                "MABLETHORPE 182",
                "MACCLESFIELD 244",
                "MANCHESTER 251",
                "MARGATE 46",
                "MIDDLESBROUGH 286",
                "MILFORD HAVEN 289",
                "MILTON KEYNES 101",
                "MORPETH 327",
                "NANTWICH 232",
                "NEWBURY 101",
                "NEWCASTLE 308",
                "NEWPORT 178",
                "NEWQUAY 178",
                "NORTHAMPTON 116",
                "NORTHUMBERLAND 341",
                "NORWICH 136",
                "NOTTINGHAM 177",
                "OKEHAMPTON 232",
                "OXFORD 106",
                "PENRITH 316",
                "PENZANCE 318",
                "PERTH 477",
                "PETERBOROUGH 124",
                "PETERSFIELD 87",
                "PETWORTH 71",
                "PLYMOUTH 247",
                "PONTEFRACT 221",
                "POOLE 144",
                "PORTSMOUTH 102",
                "READING 88",
                "REIGATE 39",
                "RINGWOOD 130",
                "ROSS-ON-WYE 171",
                "ROTHERHAM 203",
                "SALISBURY 120",
                "SCARBOROUGH 277",
                "SCUNTHORPE 204",
                "SHEFFIELD 205",
                "SHREWSBURY 207",
                "SHROPSHIRE 218",
                "SLOUGH 72",
                "SOUTH SHIELDS 310",
                "SOUTHAMPTON 112",
                "SOUTHEND 52",
                "SOUTHPORT 279",
                "SPALDING 143",
                "ST ALBANS 62",
                "ST IVES 317",
                "STAFFORD 187",
                "STAINES 61",
                "STEVENAGE 72",
                "STIRLING 445",
                "STOCKPORT 257",
                "STOCKTON 278",
                "STOKE-ON-TRENT 205",
                "STRATFORD UPON AVON 151",
                "SUNDERLAND 309",
                "SWINDON 121",
                "TAMWORTH 180",
                "TAUNTON 185",
                "TELFORD 193",
                "TILBURY 34",
                "TORQUAY 227",
                "TUNBRIDGE WELLS 26",
                "UXBRIDGE 74",
                "WAKEFIELD 214",
                "WARMISTER 137",
                "WARWICK 148",
                "WATFORD 67",
                "WELSHPOOL 238",
                "WEMBLEY 55",
                "WEYMOUTH 173",
                "WHITBY 282",
                "WIGAN 252",
                "WINCANTON 149",
                "WINCHESTER 100",
                "WOKING 60",
                "WOLVERHAMPTON 175",
                "WORCESTER 160",
                "WREXHAM 250",
                "YEOVIL 163",
                "YORK 243"
            ],
            'column': 'F',
            'target_col': 'D',
            'target_row': 183  # Fixed row for delivery location
        },
        'plant_hire_1': {
            'options': [
                "", "SL10 GENIE", "EXTENSION FORKS", "2.5M COMBI LADDER",
                "1.5M PODIUM", "3M TOWER", "COMBI LADDER", "PECO LIFT",
                "3M YOUNGMAN BOARD", "GS1930 SCISSOR LIFT",
                "4-6 SHERASCOPIC", "7-9 SHERASCOPIC"
            ],
            'column': 'G',
            'target_col': 'D',
            'target_row': 184  # Fixed row for first plant hire
        },
        'plant_hire_2': {
            'options': [
                "", "SL10 GENIE", "EXTENSION FORKS", "2.5M COMBI LADDER",
                "1.5M PODIUM", "3M TOWER", "COMBI LADDER", "PECO LIFT",
                "3M YOUNGMAN BOARD", "GS1930 SCISSOR LIFT",
                "4-6 SHERASCOPIC", "7-9 SHERASCOPIC"
            ],
            'column': 'G',  # Can use same column as plant_hire_1
            'target_col': 'D',
            'target_row': 185  # Fixed row for second plant hire
        }
    }

    # Write options to Lists sheet and create validations
    for name, config in dropdowns.items():
        # Write options to Lists sheet
        for i, option in enumerate(config['options'], 1):
            list_sheet[f"{config['column']}{i}"] = option
        
        # Create range reference
        range_ref = f"Lists!${config['column']}$1:${config['column']}${len(config['options'])}"
        
        # Create validation
        dv = DataValidation(
            type="list",
            formula1=range_ref,
            allow_blank=True
        )
        sheet.add_data_validation(dv)
        
        if 'target_row' in config:
            # Fixed position dropdown (delivery/installation)
            dv.add(f"{config['target_col']}{config['target_row']}")
        else:
            # Repeating dropdowns (canopy-related)
            current_row = start_row + config['row_offset']
            while current_row <= sheet.max_row:
                dv.add(f"{config['target_col']}{current_row}")
                current_row += 17

def get_canopy_description(model, count):
    """Generate appropriate description based on canopy model"""
    model = model.upper()
    count_str = f"{count}no"  # Format the count properly
    
    # CMWF/CMWI type canopies (Water Wash)
    if model.startswith('CMW'):
        if 'F' in model:
            return f"{count_str} Extract/Supply Canopy c/w Capture Jet Tech and Water Wash Function"
        else:
            return f"{count_str} Extract Canopy c/w Capture Jet Tech and Water Wash Function"
    
    # UV type canopies
    elif model.startswith('UV'):
        if 'F' in model:
            return f"{count_str} Extract/Supply Canopies c/w Capture Jet Tech and UV-c Filtration"
        else:
            return f"{count_str} Extract Canopies c/w Capture Jet Tech and UV-c Filtration"
    
    # CXW type canopies
    elif model.startswith('CXW'):
        return f"{count_str} Condense Canopies c/w Extract and Led Light"
    
    # Standard canopies (KV, etc)
    else:
        if 'F' in model:
            return f"{count_str} Extract/Supply Canopies c/w Capture Jet Tech"
        else:
            return f"{count_str} Extract Canopies c/w Capture Jet Tech"

def format_price(value):
    """Format price with commas and 2 decimal places"""
    try:
        # Convert to float and round up
        num = math.ceil(float(str(value).replace(',', '')))
        # Format with commas and 2 decimal places
        return f"{num:,.2f}"
    except (ValueError, TypeError):
        return "0.00"

def write_to_word_doc(data, project_data, output_path="output.docx"):
    """Write data to Word document"""
    areas = []
    
    st.write("Debug: Starting to process areas")
    st.write(f"Debug: Number of sheets in project_data: {len(project_data['sheets'])}")
    
    # Process each sheet
    for sheet in project_data['sheets']:
        st.write(f"Debug: Processing sheet {sheet.get('sheet_name', 'Unknown')}")
        
        # Get canopies from sheet
        canopies = sheet['canopies']
        
        # Skip if no valid canopies
        if not any(canopy['reference_number'] != 'ITEM' and 
                  canopy['model'] != 'CANOPY TYPE' and
                  canopy['reference_number'] != 'DELIVERY & INSTALLATION'
                  for canopy in canopies):
            st.write(f"Debug: Skipping sheet {sheet.get('sheet_name', 'Unknown')} - no valid canopies")
            continue
        
        # Get display name
        display_name = sheet['sheet_name']
        st.write(f"Debug: Processing area {display_name}")
        
        # Check for UV canopy
        has_uv_canopy = any(str(canopy.get('model', '')).upper().startswith('UV') for canopy in canopies)
        st.write(f"Debug: Area {display_name} has UV canopy: {has_uv_canopy}")
        
        # Get fire suppression data
        has_fire_suppression = any(canopy.get('has_fire_suppression', False) for canopy in canopies)
        
        # Calculate totals
        area_total = float(str(sheet['total_price']).replace(',', ''))
        delivery_total = float(str(sheet.get('delivery_install', 0)).replace(',', ''))
        commissioning_total = float(str(sheet.get('commissioning_price', 0)).replace(',', ''))
        sheet_total = area_total
        
        # Get MUA calculations
        mua_calcs = sheet.get('mua_calculations', {})
        total_extract = mua_calcs.get('total_extract_volume', 0)
        required_mua = mua_calcs.get('required_mua', 0)
        total_mua = mua_calcs.get('total_mua_volume', 0)
        mua_shortfall = mua_calcs.get('mua_shortfall', 0)
        important_note = sheet.get('important_note', '')
        
        # Get UV price from project data
        uv_price = project_data.get('uv_control_data', {}).get(display_name, {}).get('price', 0)
        st.write(f"Debug: UV price for {display_name}: {uv_price}")
        
        # Create area data
        area_data = {
            'name': display_name,
            'canopies': [{**canopy, 'base_price': format_price(canopy['base_price'])} for canopy in canopies],
            'has_uv': has_uv_canopy,
            'uv_price': format_price(uv_price) if has_uv_canopy else None,
            'has_fire_suppression': has_fire_suppression,
            'area_total': format_price(area_total),
            'delivery_total': format_price(delivery_total),
            'commissioning_total': format_price(commissioning_total),
            'sheet_total': format_price(sheet_total),
            'mua_calculations': {
                'total_extract_volume': round(total_extract, 3),
                'required_mua': round(required_mua, 3),
                'total_mua_volume': round(total_mua, 3),
                'mua_shortfall': round(mua_shortfall, 3)
            },
            'important_note': important_note
        }
        
        areas.append(area_data)
        st.write(f"Debug: Added area {display_name} to areas list")
    
    st.write(f"Debug: Final number of areas: {len(areas)}")
    st.write("Debug: Areas list content:", areas)
    
    # Create Word document
    template_path = "resources/template.docx"
    
    if not os.path.exists(template_path):
        return "Template file not found"
    
    # Load template
    doc = DocxTemplate(template_path)
    
    # Prepare context
    context = {
        'areas': areas,
        'project_info': data,
        'has_fire_suppression': any(area['has_fire_suppression'] for area in areas),
        'has_uv': any(area['has_uv'] for area in areas),
        'job_total': format_price(sum(float(str(area['sheet_total']).replace(',', '')) for area in areas))
    }
    
    st.write("Debug: Context prepared with areas:", context['areas'])
    
    # Render template
    doc.render(context)
    doc.save(output_path)
    
    return "Word document generated successfully"

def format_price(value):
    """Format price with commas and 2 decimal places"""
    try:
        # Convert to float and round up
        num = math.ceil(float(str(value).replace(',', '')))
        # Format with commas and 2 decimal places
        return f"{num:,.2f}"
    except (ValueError, TypeError):
        return "0.00"

def write_to_word_doc(data, project_data, output_path="output.docx"):
    """Write project information to a Word document using Jinja template"""
    f_canopy_mua_vols = []
    ww_canopies = []
    has_water_wash_canopies = False
    has_uv_canopy = False  # Also initialize if used
    for sheet in project_data['sheets']:
        for canopy in sheet['canopies']:
            # Skip non-canopy rows
            if (canopy['reference_number'] == 'ITEM' or 
                canopy['model'] == 'CANOPY TYPE' or
                canopy['reference_number'] == 'DELIVERY & INSTALLATION' or
                not any(valid_model in str(canopy['model']).upper() 
                       for valid_model in VALID_CANOPY_MODELS)):  # Check against valid models
                continue
            
            # Convert model and configuration to strings and handle None values
            model = str(canopy.get('model', '')) if canopy.get('model') is not None else ''
            config = str(canopy.get('configuration', '')) if canopy.get('configuration') is not None else ''
            
            # Check for UV canopies
            if ('UV' in model.upper() or 'UV' in config.upper()):
                has_uv_canopy = True
            
            # Check for water wash canopies
            if 'CMWI' in model.upper() or 'CMWF' in model.upper():
                has_water_wash_canopies = True
                ww_canopies.append(canopy)
            
            # Check for F-type canopies
            if 'F' in model.upper() and canopy.get('mua_vol'):
                f_canopy_mua_vols.append({
                    'item_no': canopy['reference_number'],
                    'model': model,
                    'mua_vol': canopy['mua_vol']
                })

    # Load the template
    doc = DocxTemplate("resources/Halton Quote Feb 2024 (1).docx")
    
    # Get the full sales contact name from initials
    sales_initials = data['Sales Contact']
    sales_contact_name = next((name for name in contacts.keys() if name.startswith(sales_initials)), None)
    
    # If we couldn't find the full name, try to get it from the contacts dictionary
    if not sales_contact_name:
        # Try to find a matching contact by initials
        for full_name in contacts.keys():
            if get_initials(full_name) == sales_initials:
                sales_contact_name = full_name
                break
    
    # If we still don't have a name, use the initials
    if not sales_contact_name:
        st.error(f"Could not find contact information for {sales_initials}")
        sales_contact_name = sales_initials
        contact_number = ""
    else:
        contact_number = contacts[sales_contact_name]
    
    # Extract customer's first name
    customer_full = data['Customer']
    customer_first_name = customer_full.split()[0] if customer_full else ""
    
    # Get estimator initials from the sales_estimator field (format: "sales/estimator")
    estimator_initials = data.get('Estimator', '').split('/')[-1] if '/' in data.get('Estimator', '') else ''
    
    # Find estimator name and role from initials
    estimator_name = ''
    estimator_role = ''
    for full_name, role in estimators.items():
        if get_initials(full_name) == estimator_initials:
            estimator_name = full_name
            estimator_role = role
            break
    
    # Get current revision
    current_revision = project_data['sheets'][0].get('revision', 'A')
    
    # Create Halton reference with project number, initials and revision
    halton_ref = f"{data['Project Number']}/{get_initials(sales_contact_name)}/{estimator_initials}/{current_revision}"
    
    # Create quote title based on revision
    quote_title = "QUOTATION - Revision A" if current_revision == 'A' else f"QUOTATION - Revision {current_revision}"
    
    # Initialize global flags and arrays
    has_water_wash_canopies = False
    ww_canopies = []
    
    # Collect all cladding items first
    cladding_items = []
    
    for sheet in project_data['sheets']:
        for canopy in sheet['canopies']:
            # Skip non-canopy rows
            if (canopy['reference_number'] == 'ITEM' or 
                canopy['model'] == 'CANOPY TYPE' or
                canopy['reference_number'] == 'DELIVERY & INSTALLATION' or
                not any(valid_model in str(canopy['model']).upper() 
                       for valid_model in VALID_CANOPY_MODELS)):  # Check against valid models
                continue
                
            # Convert model and configuration to strings
            model = str(canopy.get('model', '')) if canopy.get('model') is not None else ''
            config = str(canopy.get('configuration', '')) if canopy.get('configuration') is not None else ''
            
            # Check for UV canopies
            if ('UV' in model.upper() or 'UV' in config.upper()):
                has_uv_canopy = True
            
            # Check for water wash canopies
            if 'CMWI' in model.upper() or 'CMWF' in model.upper():
                has_water_wash_canopies = True
                ww_canopies.append(canopy)
            
            # Check for F-type canopies
            if 'F' in model.upper() and canopy.get('mua_vol'):
                f_canopy_mua_vols.append({
                    'item_no': canopy['reference_number'],
                    'model': model,
                    'mua_vol': canopy['mua_vol']
                })
            
            # Create description based on selected positions
            positions = canopy['wall_cladding']['positions']
            position_str = '/'.join(positions).lower() if positions else ''
            
            if position_str:  # Only add if there are positions selected
                cladding_items.append({
                    'item_no': canopy['reference_number'],
                    'description': f"Cladding below Item {canopy['reference_number']}, supplied and installed",
                    'width': canopy['wall_cladding']['width'],
                    'height': canopy['wall_cladding']['height'],
                    'price': math.ceil(canopy['wall_cladding']['price'])  # Round up the price
                })
    
    # Prepare areas data with technical specifications
    areas = []
    
    for sheet in project_data['sheets']:
        display_name = sheet['sheet_name']
        canopies = []
        has_uv_canopy = False
        has_fire_suppression = False
        fire_suppression_canopies = []
        fs_canopy_count = 0  # Track count of fire suppression canopies
        fs_base_total = 0    # Track total of N12 prices
        
        # Get MUA calculations for this area
        mua_calcs = sheet.get('mua_calculations', {})
        total_extract = mua_calcs.get('total_extract_volume', 0)
        required_mua = mua_calcs.get('required_mua', 0)
        total_mua = mua_calcs.get('total_mua_volume', 0)
        mua_shortfall = mua_calcs.get('mua_shortfall', 0)
        
        # Format the important note for this area
        important_note = (
            f"Important Note: - The make-up air flows shown above are the maximum that we can introduce through the "
            f"canopy. This should be equal to approximately 85% of the extract i.e. {required_mua}m³/s\n"
            f"In this instance it only totals {total_mua}m³/s therefore the shortfall of "
            f"{mua_shortfall}m³/s must be introduced through ceiling grilles or diffusers, by others.\n"
            f"If you require further guidance on this, please do not hesitate to contact us."
        ) if total_extract > 0 else ""
        
        # Get shared installation price from N182 safely
        try:
            fs_install_price = float(sheet.get('fire_suppression_install', 0) or 0)
        except (TypeError, ValueError):
            fs_install_price = 0
        
        # Get cladding items for this area
        area_cladding_items = [item for item in cladding_items 
                             if any(c['reference_number'] == item['item_no'] 
                                   for c in sheet['canopies'])]
        
        for idx, canopy in enumerate(sheet['canopies']):
            if (canopy['reference_number'] != 'ITEM' and 
                canopy['model'] != 'CANOPY TYPE'):
                
                # Convert model and configuration to strings
                model = str(canopy.get('model', '')) if canopy.get('model') is not None else ''
                config = str(canopy.get('configuration', '')) if canopy.get('configuration') is not None else ''
                
                # Check for UV canopies
                if ('UV' in model.upper() or 'UV' in config.upper()):
                    has_uv_canopy = True
                
                # Handle fire suppression
                if canopy.get('has_fire_suppression'):
                    has_fire_suppression = True
                    fs_canopy_count += 1
                    
                    # Get fire suppression base price from the canopy's fire suppression data
                    try:
                        fs_base_price = float(canopy.get('fire_suppression_data', {}).get('base_price', 0) or 0)
                        #st.write(f"Base price for {canopy['reference_number']}: {fs_base_price}")
                    except (TypeError, ValueError):
                        fs_base_price = 0
                        #st.write(f"Failed to get base price for {canopy['reference_number']}, using default: 0")
                    
                    fs_base_total += fs_base_price
                    
                    # Get installation share from N182 in the fire suppression sheet
                    try:
                        # Get the installation share from the fire suppression data
                        install_share = float(canopy['fire_suppression_data'].get('install_price', 0) or 0)
                        # Count total number of fire suppression canopies in this area, excluding the header row
                        total_fs_canopies = sum(1 for c in sheet['canopies'] 
                                              if c.get('has_fire_suppression') and 
                                              c.get('reference_number') and 
                                              c.get('reference_number') != 'ITEM' and
                                              c.get('model') and 
                                              c.get('model') != 'CANOPY TYPE')
                        # Divide the installation share by the total number of fire suppression canopies
                        install_share = install_share / total_fs_canopies
                        #st.write(f"Found installation share in fire suppression data: {install_share}")
                        #st.write(f"Divided by {total_fs_canopies} total canopies with fire suppression in this area (excluding header)")
                    except (TypeError, ValueError) as e:
                        install_share = 0
                       # st.write(f"Failed to get installation share: {str(e)}")
                        #st.write(f"Using default: 0")
                    
                    # Calculate total price with ceiling rounding
                    total_price = math.ceil(fs_base_price + install_share)
                    
                    fire_suppression_canopies.append({
                        'item_number': canopy['reference_number'],
                        'model': canopy['model'],
                        'system_description': canopy['fire_suppression_data']['system_description'],
                        'tank_quantity': canopy['fire_suppression_data']['tank_quantity'],
                        'manual_release': canopy['fire_suppression_data']['manual_release'],
                        'base_price': format_price(fs_base_price+install_share),
                        'install_share': format_price(install_share),
                        'total_price': format_price(fs_base_price + install_share)
                    })
                    #st.write(f"Fire suppression data for {canopy['reference_number']}:")
                   # st.write(f"Base price: {fs_base_price}")
                   # st.write(f"Install share (divided by {total_fs_canopies} total canopies): {install_share}")
                    #st.write(f"Total price: {total_price}")
                
                canopies.append({
                    'item_no': canopy['reference_number'],
                    'model': canopy['model'],
                    'configuration': canopy['configuration'],
                    'length': canopy['length'] or 0,
                    'width': canopy['width'] or 0,
                    'height': canopy['height'] or 0,
                    'sections': canopy['sections'] or 1,
                    'ext_vol': canopy['ext_vol'] or 0,
                    'ext_static': canopy['ext_static'] or 0,
                    'mua_vol': canopy['mua_vol'] or 0,
                    'supply_static': canopy['supply_static'] or 0,
                    'lighting': canopy.get('lighting', 'LED Strip'),
                    'cws_2bar': canopy['cws_2bar'] or '-',
                    'hws_2bar': canopy['hws_2bar'] or '-',
                    'hws_storage': canopy['hws_storage'] or '-',
                    'ww_price': canopy['ww_price'] or 0,
                    'ww_control_price': canopy['ww_control_price'] or 0,
                    'ww_install_price': canopy['ww_install_price'] or 0,
                    'base_price': canopy['base_price'] or 0,
                    'has_fire_suppression': canopy.get('has_fire_suppression', False),
                    'fire_suppression_data': canopy.get('fire_suppression_data')
                })
                #st.write(canopy.get('fire_suppression_data'))
                print(canopy.get('fire_suppression_data'))
        
        if canopies:
            # Calculate area totals with rounding at each step
            canopy_total = sum(math.ceil(float(str(canopy['base_price']).replace(',', ''))) for canopy in canopies)
            delivery_total = math.ceil(float(str(sheet['delivery_install']).replace(',', '')))
            commissioning_total = math.ceil(float(str(sheet['commissioning_price']).replace(',', '')))
            
            # Calculate fire suppression totals
            fs_total = 0
            if has_fire_suppression:
                # Get total from the fire suppression sheet
              #  st.write(f"Looking for fire suppression sheet for area: {sheet['sheet_name']}")
                area_name = sheet['sheet_name']  # Store the area name we're looking for
                
                # First find all FIRE SUPP sheets
                fire_supp_sheets = [fs_sheet for fs_sheet in project_data['sheets'] 
                                  if 'FIRE SUPP' in fs_sheet['sheet_name'] and 'F24' not in fs_sheet['sheet_name']]
                
             #   st.write(f"Found {len(fire_supp_sheets)} FIRE SUPP sheets")
                
                # Extract area name from the sheet name (e.g., "Second Kitchen - First Floor")
                area_parts = area_name.split(' - ')
                if len(area_parts) >= 2:
                    area_to_match = area_parts[0]  # Get the first part (e.g., "Second Kitchen")
                    floor_to_match = area_parts[1]  # Get the second part (e.g., "First Floor")
                    
                   # st.write(f"Looking for area '{area_to_match}' on floor '{floor_to_match}'")
                    
                    # Then look for a matching area name in the fire suppression sheet name
                    for fs_sheet in fire_supp_sheets:
                        fs_sheet_name = fs_sheet['sheet_name']
                      # st.write(f"Checking FIRE SUPP sheet name: {fs_sheet_name}")
                        
                        # Extract area name from fire suppression sheet (e.g., "FIRE SUPP - Second Kitchen (1)")
                        if ' - ' in fs_sheet_name:
                            fs_area = fs_sheet_name.split(' - ')[1]  # Get "Second Kitchen (1)"
                            fs_area = fs_area.split('(')[0].strip()  # Remove the "(1)" part
                            
                        #    st.write(f"Comparing areas: '{area_to_match}' with '{fs_area}'")
                            if area_to_match == fs_area:
                               # st.write(f"Found matching FIRE SUPP sheet: {fs_sheet_name}")
                                try:
                                    # Get N9 value from the fire suppression sheet
                                    fs_total = math.ceil(float(str(fs_sheet.get('N9', 0) or 0).replace(',', '')))
                                  #  st.write(f"Found N9 value: {fs_total}")
                                except (TypeError, ValueError):
                                    fs_total = 0
                                break
                    else:
                        #st.write(f"No matching FIRE SUPP sheet found for area: {area_to_match} on floor: {floor_to_match}")
                        fs_total = 0
                else:
                   # st.write(f"Could not parse area name from: {area_name}")
                    fs_total = 0
                
                # Add fire suppression total to area total
                area_total = math.ceil(canopy_total + delivery_total + commissioning_total + fs_total)
                
                # Calculate total fire suppression price from canopies
                fs_canopies_total = sum(float(str(canopy.get('total_price', 0) or 0).replace(',', '')) for canopy in fire_suppression_canopies) if fire_suppression_canopies else 0
                
                # Add both fire suppression totals to the sheet total
                sheet_total = math.ceil(float(str(sheet['total_price'] or 0).replace(',', ''))) + math.ceil(fs_canopies_total)
                
                areas.append({
                    'name': display_name,
                    'canopies': [{**canopy, 'base_price': format_price(canopy['base_price'])} for canopy in canopies],
                    'has_uv': has_uv_canopy,
                    'has_fire_suppression': has_fire_suppression,
                    'fire_suppression_canopies': [{
                        **fs_canopy,
                        'base_price': format_price(fs_canopy['base_price']),
                        'install_share': format_price(fs_canopy['install_share']),
                        'total_price': format_price(fs_canopy['total_price'])
                    } for fs_canopy in fire_suppression_canopies],
                    'fire_suppression_total': format_price(fs_total),
                    'fire_suppression_data': {
                        'system_description': 'SELECT WORKS',
                        'tank_quantity': 'SELECT WORKS',
                        'manual_release': '1no station',
                        'base_price': format_price(fs_base_total),
                        'install_price': format_price(fs_install_price),
                        'total_price': format_price(fs_canopies_total)
                    } if has_fire_suppression else None,
                    'fire_suppression_base_total': format_price(fs_base_total),
                    'fire_suppression_install': format_price(fs_install_price),
                    'has_cladding': bool(area_cladding_items),
                    'cladding_items': [{**item, 'price': format_price(item['price'])} for item in area_cladding_items],
                    'canopy_total': format_price(canopy_total),
                    'delivery_total': format_price(delivery_total),
                    'commissioning_total': format_price(commissioning_total),
                    'cladding_total': format_price(sum(math.ceil(float(str(item['price']).replace(',', ''))) for item in area_cladding_items)),
                    'uv_total': format_price(1040.00 if has_uv_canopy else 0),
                    'area_total': format_price(area_total),
                    'total_price': format_price(sheet_total),
                    # Add MUA calculations
                    'mua_calculations': {
                        'total_extract_volume': round(total_extract, 3),
                        'required_mua': round(required_mua, 3),
                        'total_mua_volume': round(total_mua, 3),
                        'mua_shortfall': round(mua_shortfall, 3)
                    },
                    'important_note': important_note
                })
    
    # Calculate totals from all sheets
    job_total = 0.0
    k9_total = 0.0
    commissioning_total = 0.0  # Add commissioning total tracking
    
    for sheet in project_data['sheets']:
        # Only add to total if sheet has canopies (is used)
        if any(canopy['reference_number'] != 'ITEM' and 
               canopy['model'] != 'CANOPY TYPE' and
               canopy['reference_number'] != 'DELIVERY & INSTALLATION'
               for canopy in sheet['canopies']):
            try:
                # Convert total price to float and add to job total
                sheet_total = float(str(sheet['total_price']).replace(',', ''))
                sheet_k9 = float(str(sheet['k9_total']).replace(',', ''))
                sheet_commissioning = float(str(sheet['commissioning_price']).replace(',', ''))
                
                job_total += sheet_total
                k9_total += sheet_k9
                commissioning_total += sheet_commissioning
            except (ValueError, TypeError):
                st.write(f"Warning: Could not convert total price for sheet {sheet.get('sheet_name', 'Unknown')}")
                continue
    
    # Add fire suppression totals to job total
    try:
        fs_n9_total = float(str(project_data['global_fs_n9_total'] or 0).replace(',', ''))
        job_total += fs_n9_total
    except (ValueError, TypeError):
        st.write("Warning: Could not convert fire suppression N9 total")
    
    # Format totals with commas and 2 decimal places
    job_total_formatted = f"{math.ceil(job_total):,.2f}"
    k9_total_formatted = f"{math.ceil(k9_total):,.2f}"
    commissioning_total_formatted = f"{math.ceil(commissioning_total):,.2f}"
    
    # Add to context
    context = {
        'date': data['Date'],
        'project_number': halton_ref,  # Updated reference format
        'quote_title': quote_title,    # New quote title
        'sales_contact_name': sales_contact_name,
        'contact_number': contact_number,
        'customer': data['Customer'],
        'customer_first_name': customer_first_name,
        'company': data.get('Company', ''),
        'estimator_name': estimator_name,  # Use extracted name
        'estimator_role': estimator_role,  # Use extracted role
        'scope_items': [],  # Add scope items to context
        'project_data': project_data,  # Add project_data to context first
        'project_name': project_data['sheets'][0]['project_info']['project_name'],
        'location': project_data['sheets'][0]['project_info']['location'],
        'areas': areas,
        'has_water_wash': has_water_wash_canopies,
        'has_cladding': bool(cladding_items),
        'cladding_items': cladding_items,
        'cladding_total': sum(math.ceil(item['price']) for item in cladding_items),
        'f_canopy_mua_vols': f_canopy_mua_vols,
        'ww_canopies': ww_canopies,
        'job_total': job_total_formatted,
        'k9_total': k9_total_formatted,
        'commissioning_total': commissioning_total_formatted,  # Add to context with proper formatting
    }
    print(context.get('estimator_name'))
    # Collect and organize canopy data for scope of works
    canopy_counts = {}
    for sheet in project_data['sheets']:
        for canopy in sheet['canopies']:
            # Skip invalid or placeholder canopies
            if (canopy['reference_number'] == 'ITEM' or 
                canopy['model'] == 'CANOPY TYPE' or 
                not canopy['model'] or 
                not canopy['configuration'] or 
                canopy['model'] == "Select..." or 
                canopy['configuration'] == "Select..."):
                continue
                
            key = f"{canopy['model']} {canopy['configuration']}"
            canopy_counts[key] = canopy_counts.get(key, 0) + 1
    
    # Format scope of works
    scope_items = []
    for canopy_type, count in canopy_counts.items():
        scope_items.append(get_canopy_description(canopy_type, count))
    
    # Add any wall cladding from valid canopies only
    wall_cladding_count = sum(1 for sheet in project_data['sheets'] 
                             for canopy in sheet['canopies'] 
                             if (canopy['wall_cladding']['type'] and 
                                 canopy['wall_cladding']['type'] != "Select..." and 
                                 canopy['wall_cladding']['type'] == "2M² (HFL)" and
                                 canopy['reference_number'] != 'ITEM' and
                                 canopy['model'] != 'CANOPY TYPE'))
    if wall_cladding_count > 0:
        scope_items.append(f"{wall_cladding_count}no Areas with Stainless Steel Cladding")
    
    # Add scope items to context
    context['scope_items'] = scope_items
    
    # Inside write_to_word_doc function, before processing areas
    # Create a map of fire suppression data by item number
    fs_data = []
    for sheet in project_data['sheets']:
        for canopy in sheet['canopies']:
            if canopy.get('has_fire_suppression') and canopy.get('fire_suppression_data'):
                fs_data.append({
                    'item_number': canopy['reference_number'],
                    'model': canopy['model'],
                    'system_description': canopy['fire_suppression_data']['system_description'],
                    'tank_quantity': canopy['fire_suppression_data']['tank_quantity'],
                    'manual_release': canopy['fire_suppression_data']['manual_release']
                })

    # Add fire suppression data to context
    context.update({
        'has_fire_suppression': bool(fs_data),
        'fire_suppression_data': fs_data,
        'fire_suppression_description': """The restaurant fire suppression system is a pre-engineered, wet chemical, cartridge-operated, regulated pressure type with a fixed nozzle agent distribution network. The system is capable of automatic detection and actuation and / or remote manual actuation to provide fire protection to the cooking appliances, canopy exhaust duct and canopy filter plenum. Fire alarm / BMS connections to the release module to be carried out by others."""
    })
    
    # Render the template with the context
    doc.render(context)
    
    # Save the document
    doc.save(output_path)
    return output_path

# Modify save_to_excel function to remove Word document generation
def save_to_excel(data):
    try:
        # Define paths using os.path for cross-platform compatibility
        template_path = os.path.join("resources", "Halton Cost Sheet Jan 2025.xlsx")
        output_path = "output.xlsx"
        
        st.write("🔍 Starting Excel file generation...")
        
        # Check if template exists
        if not os.path.exists(template_path):
            st.error(f"Template file not found. Please ensure '{template_path}' exists in the resources folder.")
            return
            
        # Load workbook
        workbook = openpyxl.load_workbook(template_path, read_only=False, data_only=False)
        
        # Get all sheets once
        all_sheets = workbook.sheetnames
        canopy_sheets = [sheet for sheet in all_sheets if 'CANOPY' in sheet]
        fire_supp_sheets = [sheet for sheet in all_sheets if 'FIRE SUPP' in sheet]
        edge_box_sheets = [sheet for sheet in all_sheets if 'EDGE BOX' in sheet]
        
        st.write(f"📑 Found template sheets: {len(canopy_sheets)} CANOPY, {len(fire_supp_sheets)} FIRE SUPP, {len(edge_box_sheets)} EBOX")
        
        sheet_count = 0
        fs_sheet_count = 0
        ebox_sheet_count = 0
        
        # Create a template cache for fire suppression sheets
        fs_template = None
        for fs_name in fire_supp_sheets:
            temp_sheet = workbook[fs_name]
            b1_value = temp_sheet['B1'].value
            if b1_value and "F24 - 19" in b1_value and "CANOPY COST SHEET" in b1_value:
                fs_template = temp_sheet
                fire_supp_sheets.remove(fs_name)
                st.write("✅ Found FIRE SUPP template sheet")
                break
        
        # Check for available EDGE BOX sheets
        if not edge_box_sheets:
            st.warning("⚠️ No EDGE BOX sheets found - UV-C Control Schedule will not be included")
        else:
            st.write("✅ Found EDGE BOX sheets available")
        
        # Process each level and area
        for level in data['Levels']:
            st.write(f"📝 Processing level: {level['level_name']}")
            for area in level['areas']:
                st.write(f"  📍 Processing area: {area['area_name']}")
                
                # Debug canopy models
                for canopy in area['canopies']:
                    model = str(canopy.get('model', ''))
                    if model.upper().startswith('UV'):
                        st.write(f"  🔎 Found UV canopy: {model}")
                
                # Check if any canopy in this area has fire suppression
                has_fire_suppression = any(canopy.get('fire_suppression', False) for canopy in area['canopies'])
                
                # Check if this area needs an EDGE BOX sheet
                needs_edge_box = area.get('include_uvc', False)
                
                if needs_edge_box:
                    st.write(f"  ✨ Area has UV-C Control Schedule - will create EDGE BOX sheet")
                
                # Handle CANOPY sheet
                if sheet_count >= len(canopy_sheets):
                    st.error(f"Not enough CANOPY sheets in template!")
                    break
                
                sheet_name = canopy_sheets[sheet_count]
                current_sheet = workbook[sheet_name]
                current_sheet.sheet_state = 'visible'
                
                # Rename the CANOPY sheet
                new_sheet_name = f"CANOPY - {level['level_name']} ({sheet_count + 1})"
                current_sheet.title = new_sheet_name
                st.write(f"  📄 Created sheet: {new_sheet_name}")
                
                # Handle FIRE SUPPRESSION sheet if needed
                fs_sheet = None
                if has_fire_suppression:
                    if fire_supp_sheets:
                        # Use next available FIRE SUPP sheet
                        fs_sheet_name = fire_supp_sheets.pop(0)
                        fs_sheet = workbook[fs_sheet_name]
                        new_fs_name = f"FIRE SUPP - {level['level_name']} ({sheet_count + 1})"
                        fs_sheet.title = new_fs_name
                        fs_sheet.sheet_state = 'visible'
                        fs_sheet_count += 1
                        st.write(f"  🔥 Using FIRE SUPP sheet: {new_fs_name}")
                    else:
                        st.error("Not enough FIRE SUPP sheets in template!")
                        break
                
                # Handle EDGE BOX sheet if needed
                if needs_edge_box and edge_box_sheets:
                    # Get next available EDGE BOX sheet
                    ebox_sheet_name = edge_box_sheets.pop(0)
                    ebox_sheet = workbook[ebox_sheet_name]
                    new_ebox_name = f"EBOX - {level['level_name']} ({ebox_sheet_count + 1})"
                    ebox_sheet.title = new_ebox_name
                    ebox_sheet.sheet_state = 'visible'
                    # Write general info to EDGE BOX sheet
                    write_to_sheet(ebox_sheet, data, level['level_name'], area['area_name'], area['canopies'], None, True)
                    ebox_sheet_count += 1
                    st.write(f"  📦 Created EDGE BOX sheet: {new_ebox_name}")
                elif needs_edge_box:
                    st.warning(f"  ⚠️ No more EDGE BOX sheets available for UV-C Control Schedule in {area['area_name']}")
                
                # Write data to the CANOPY sheet
                write_to_sheet(current_sheet, data, level['level_name'], area['area_name'], area['canopies'], fs_sheet)
                
                # Add dropdowns to sheets
                add_dropdowns_to_sheet(workbook, current_sheet, 12)
                if fs_sheet:
                    add_fire_suppression_dropdown(fs_sheet)
                
                sheet_count += 1
        
        # Save the workbook
        workbook.save(output_path)
        st.success(f"✅ Successfully created:\n- {sheet_count} CANOPY sheets\n- {fs_sheet_count} FIRE SUPPRESSION sheets\n- {ebox_sheet_count} EDGE BOX sheets")
        
        # Provide download button
        with open(output_path, "rb") as file:
            st.download_button(
                label="📥 Download Excel file",
                data=file,
                file_name="project_data.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
    except FileNotFoundError:
        st.error("❌ Template file not found in resources folder!")
    except Exception as e:
        st.error(f"❌ An error occurred: {str(e)}")
        raise e

def extract_sheet_data(sheet):
    """Extract data from a single sheet"""
    # Get estimator info from Lists sheet
    list_sheet = sheet.parent['Lists']
    estimator_name = list_sheet['Z1'].value
    estimator_role = list_sheet['Z2'].value
    
    # Get sheet name from B1 (Floor Name - Area)
    display_name = sheet['B1'].value or sheet.title
    
    # Parse customer and company from C5
    customer_value = sheet['C5'].value or ''
    if '(' in customer_value and ')' in customer_value:
        # Split into customer and company if in format "Customer (Company)"
        customer = customer_value.split('(')[0].strip()
        company = customer_value.split('(')[1].rstrip(')').strip()
    else:
        customer = customer_value
        company = ''
    
    # Helper function to safely convert values to float
    def safe_convert_to_float(value):
        if not value:
            return 0.0
        if isinstance(value, (int, float)):
            return float(value)
        # Convert to string and clean up
        str_value = str(value).upper().strip()
        # Check for special values
        if 'SELECT' in str_value or str_value == '-':
            return 0.0
        try:
            # Remove commas and convert to float
            return float(str_value.replace(',', ''))
        except (ValueError, TypeError):
            return 0.0
    
    # Get delivery and installation price from P182
    delivery_install_price = safe_convert_to_float(sheet['P182'].value)
    
    # Get K9 total from sheet
    k9_total = math.ceil(safe_convert_to_float(sheet['K9'].value))
    
    # Get commissioning price from N193
    commissioning_price = safe_convert_to_float(sheet['N193'].value)
    
    data = {
        'sheet_name': display_name,  # Use B1 value instead of sheet title
        'revision': sheet['O7'].value,
        'delivery_install': delivery_install_price,
        'total_price': math.ceil(safe_convert_to_float(sheet['N9'].value)),
        'k9_total': k9_total,
        'commissioning_price': commissioning_price,  # Store raw value
        'project_info': {
            'project_number': sheet['C3'].value,
            'customer': customer,
            'company': company,
            'sales_estimator': sheet['C7'].value,
            'project_name': sheet['G3'].value,
            'location': sheet['G5'].value,
            'date': sheet['G7'].value,
            'estimator_name': estimator_name,
            'estimator_role': estimator_role
        },
        'canopies': [],
        'fire_suppression_items': []  # Add new list for fire suppression items
    }
    
    # Extract canopy data - only process up to row 181 for canopies
    row = 12  # Starting row for first canopy
    while row <= 181:  # Processing canopies
        reference_number = sheet[f'B{row}'].value
        
        # Skip this canopy if reference number contains "ITEM" or is empty
        if not reference_number or 'ITEM' in str(reference_number).upper().strip():
            row += 17
            continue
            
        model = sheet[f'D{row + 2}'].value
        configuration = sheet[f'C{row + 2}'].value
        
        # Only process if we have valid model and configuration
        if (model and model != "Select..." and 
            configuration and configuration != "Select..."):
            
            dim_row = row + 2
            pa_row = row + 10  # F22 for first canopy
            cladding_row = row + 7
            water_row = row + 13  # F25 for first canopy water wash values
            
            # Get K9 value for this canopy
            canopy_k9 = safe_convert_to_float(sheet[f'K{row}'].value)
            
            # Get fire suppression data
            fs_system = sheet[f'C{row + 4}'].value  # C16 - System description
            fs_tank_qty = sheet[f'C{row + 4}'].value  # C16 - Tank quantity
            
            # Format fire suppression data
            has_fire_suppression = False
            fire_suppression_base_price = 0
            fire_suppression_install = 0
            fire_suppression_n9 = 0
            
            # Look for matching fire suppression sheet
            for fs_sheet_name in sheet.parent.sheetnames:
                if 'FIRE SUPP' in fs_sheet_name and 'F24' not in fs_sheet_name:
                    fs_sheet = sheet.parent[fs_sheet_name]
                    # Check if this is the correct fire suppression sheet for this area
                    if fs_sheet['B1'].value == sheet['B1'].value:
                        # Find matching item in fire suppression sheet
                        fs_row = 12
                        while fs_row <= 181:
                            if fs_sheet[f'B{fs_row}'].value == sheet[f'B{row}'].value:
                                has_fire_suppression = True
                                fire_suppression_base_price = safe_convert_to_float(fs_sheet[f'N{fs_row}'].value)
                                fire_suppression_base_price = math.ceil(fire_suppression_base_price)
                                
                                # Get tank quantity from C16 in fire suppression sheet
                                system_desc = fs_sheet[f'C{fs_row + 4}'].value or ""
                                if system_desc == "FIRE SUPPRESSION":
                                    tank_quantity = "-"
                                else:
                                    # Try to extract number from system description
                                    try:
                                        # Split the string and get the first part which should be the number
                                        first_word = system_desc.split()[0]
                                        if first_word.isdigit():
                                            tank_quantity = first_word
                                        else:
                                            tank_quantity = "1"
                                    except (AttributeError, IndexError):
                                        tank_quantity = "1"
                                
                                # Get N182 value from the fire suppression sheet
                                fire_suppression_install = safe_convert_to_float(fs_sheet['N182'].value)
                                
                                # Get N9 value from the fire suppression sheet
                                fire_suppression_n9 = safe_convert_to_float(fs_sheet['N9'].value)
                                break
                            fs_row += 17
                        if has_fire_suppression:
                            break

            fire_suppression_data = {
                'system_description': fs_system or f"Ansul R102 System to cover item {sheet[f'B{row}'].value}",
                'tank_quantity': tank_quantity,  # Use the extracted tank quantity
                'manual_release': '1no station',
                'base_price': math.ceil(fire_suppression_base_price),
                'install_price': math.ceil(fire_suppression_install),
                'total_price': math.ceil(fire_suppression_n9)
            } if has_fire_suppression else None
            
            canopy_price = safe_convert_to_float(sheet[f'P{row}'].value)  # Base canopy price
            canopy_price = math.ceil(canopy_price)  # Round up to nearest pound
            
            cladding_price = safe_convert_to_float(sheet[f'N{row + 7}'].value)  # Cladding price if exists
            emergency_lighting = sheet[f'P{row + 1}'].value or "2No, @ £100.00 if required"  # Emergency lighting text
            
            # Get water wash values
            cws_2bar = sheet[f'F{row + 13}'].value  # F25 - Cold water supply at 2 bar
            hws_2bar = sheet[f'F{row + 14}'].value  # F26 - Hot water supply at 2 bar
            hws_storage = sheet[f'F{row + 15}'].value  # F27 - Hot water storage
            
            # Get water wash prices
            ww_price = safe_convert_to_float(sheet[f'P{row + 13}'].value)  # P25 - Water wash price
            ww_control_price = safe_convert_to_float(sheet[f'P{row + 14}'].value)  # P26 - Control panel price
            ww_install_price = safe_convert_to_float(sheet[f'P{row + 15}'].value)  # P27 - Installation price
            
            wall_cladding = {
                'type': sheet[f'C{cladding_row}'].value,
                'width': sheet[f'Q{cladding_row}'].value or 0,
                'height': sheet[f'R{cladding_row}'].value or 0,
                'positions': (sheet[f'S{cladding_row}'].value or '').split(',') if sheet[f'S{cladding_row}'].value else [],
                'price': cladding_price
            }
            
            # Get MUA VOL from Q22 for F-type canopies
            mua_vol = '-'
            model_value = sheet[f'D{row + 2}'].value
            if model_value:
                model_str = str(model_value)  # Convert to string first
                if 'F' in model_str.upper():
                    # Get extract volume
                    ext_vol = sheet[f'I{dim_row}'].value or 0
                    # Get MAX SUPPLY value from H22 + offset
                    max_supply_cell = sheet[f'H{22 + ((row - 12) // 17) * 17}'].value or ''
                    
                    # Calculate 85% of extract volume and round to 2 decimal places
                    calculated_mua = round(float(ext_vol) * 0.85, 2) if ext_vol != '-' else 0
                    
                    # Parse MAX SUPPLY value (remove 'MAX' if present)
                    if isinstance(max_supply_cell, str) and '(MAX)' in max_supply_cell:
                        max_supply = float(max_supply_cell.replace('(MAX)', '').strip())
                        # If calculated MUA is less than MAX, use calculated MUA value
                        mua_vol = str(calculated_mua) if calculated_mua < max_supply else str(round(max_supply, 2))
                    else:
                        # If no MAX value, use calculated MUA
                        mua_vol = str(calculated_mua)
            
            # Inside the canopy processing loop, update the ext_static and lighting handling:
            model = sheet[f'D{row + 2}'].value or '-'
            lighting_value = sheet[f'C{row + 3}'].value or '-'
            print(lighting_value)
            if lighting_value == 'LIGHT SELECTION':
                lighting = '-'
            else:
                if 'LED STRIP' in str(lighting_value):
                    lighting = 'LED Strip'
                else:
                    lighting = 'LED Spots'
            print(lighting)
            # Handle ext_static based on model
            ext_static_value = sheet[f'F{pa_row}'].value or '-'
            if model == 'CXW':
                ext_static = 45  # Always 45 for CXW models
            elif ext_static_value != '-':
                # Remove 'Pa' and convert to string
                ext_static_str = str(ext_static_value).replace('Pa', '').strip()
                try:
                    # Convert to float and round
                    ext_static = round(float(ext_static_str))
                except (ValueError, TypeError):
                    ext_static = '-'
            else:
                ext_static = '-'
            
            # Handle supply_static based on model
            supply_static_value = sheet[f'L{dim_row}'].value
            if 'F' in str(model).upper():
                # Only process supply static for F-type canopies
                if supply_static_value and supply_static_value != '-':
                    try:
                        supply_static = round(float(str(supply_static_value).replace('Pa', '').strip()))
                    except (ValueError, TypeError):
                        supply_static = '-'
                else:
                    supply_static = '-'
            else:
                supply_static = '-'  # All non-F canopies get '-'
            
            # Handle control panel and water wash pod values (C25-C27)
            control_panel = sheet[f'C{row + 13}'].value  # C25
            ww_pods = sheet[f'C{row + 14}'].value  # C26
            ww_control = sheet[f'C{row + 15}'].value  # C27
            
            # Check for SELECT in values
            if isinstance(control_panel, str) and 'SELECT' in control_panel.upper():
                control_panel = '-'
            if isinstance(ww_pods, str) and 'SELECT' in ww_pods.upper():
                ww_pods = '-'
            if isinstance(ww_control, str) and 'SELECT' in ww_control.upper():
                ww_control = '-'
            
            canopy = {
                'reference_number': reference_number,
                'model': model,
                'configuration': configuration,
                'length': sheet[f'E{dim_row}'].value or '-',
                'width': sheet[f'F{dim_row}'].value or '-',
                'height': sheet[f'G{dim_row}'].value or '-',
                'sections': sheet[f'H{dim_row}'].value or '-',
                'ext_vol': sheet[f'I{dim_row}'].value or '-',
                'ext_static': ext_static,
                'mua_vol': mua_vol,
                'supply_static': supply_static,
                'lighting': lighting,
                'special_works_1': sheet[f'C{row + 4}'].value or '-',
                'special_works_2': sheet[f'C{row + 5}'].value or '-',
                'bim_revit': sheet[f'C{row + 6}'].value or '-',
                'wall_cladding': wall_cladding,
                'price': canopy_price,
                'base_price': canopy_price,
                'k9_value': canopy_k9,  # Add K9 value for individual canopy
                'emergency_lighting': emergency_lighting,
                'cws_2bar': cws_2bar,
                'hws_2bar': hws_2bar,
                'hws_storage': hws_storage,
                'ww_price': ww_price,
                'ww_control_price': ww_control_price,
                'ww_install_price': ww_install_price,
                'has_fire_suppression': has_fire_suppression,
                'fire_suppression_data': fire_suppression_data,
                'control_panel': control_panel,
                'ww_pods': ww_pods,
                'ww_control': ww_control
            }
            print(canopy)
            data['canopies'].append(canopy)
        row += 17  # Move to next canopy section
    
    # Calculate sheet K9 total from individual canopy K9 values
    calculated_k9_total = sum(math.ceil(canopy['k9_value']) for canopy in data['canopies'])
    # Use calculated total if sheet total is 0 or missing
    if k9_total == 0:
        k9_total = calculated_k9_total
    data['k9_total'] = k9_total
    
    # Calculate MUA requirements and shortfall
    total_extract_volume = 0
    total_mua_volume = 0
    
    # Sum up all extract volumes and MUA volumes
    for canopy in data['canopies']:
        # Get extract volume
        ext_vol = canopy['ext_vol']
        if ext_vol and ext_vol != '-':
            try:
                # Convert to float and add to total
                ext_vol_float = float(str(ext_vol).replace(',', ''))
                total_extract_volume += ext_vol_float
            except (ValueError, TypeError):
                pass
        
        # Get MUA volume
        mua_vol = canopy['mua_vol']
        if mua_vol and mua_vol != '-':
            try:
                # Convert to float and add to total
                mua_vol_float = float(str(mua_vol).replace(',', ''))
                total_mua_volume += mua_vol_float
            except (ValueError, TypeError):
                pass
    
    # Calculate required MUA (85% of total extract)
    required_mua = round(total_extract_volume * 0.85, 3)
    
    # Calculate shortfall (required MUA minus actual MUA volume)
    mua_shortfall = round(required_mua - total_mua_volume, 3)
    
    # Add MUA calculations to data structure
    data['mua_calculations'] = {
        'total_extract_volume': round(total_extract_volume, 3),
        'required_mua': required_mua,
        'total_mua_volume': round(total_mua_volume, 3),
        'mua_shortfall': mua_shortfall
    }
    
    # Format the important note
    data['important_note'] = (
        f"Important Note: - The make-up air flows shown above are the maximum that we can introduce through the "
        f"canopy. This should be equal to approximately 85% of the extract i.e. {required_mua}m³/s\n"
        f"In this instance it only totals {round(total_mua_volume, 3)}m³/s therefore the shortfall of "
        f"{mua_shortfall}m³/s must be introduced through ceiling grilles or diffusers, by others.\n"
        f"If you require further guidance on this, please do not hesitate to contact us."
    )
    
    # Get sheet totals using safe conversion
    data['total_price'] = math.ceil(safe_convert_to_float(sheet['N9'].value))
    data['delivery_install'] = math.ceil(safe_convert_to_float(sheet['P182'].value))
    data['commissioning_price'] = safe_convert_to_float(sheet['N193'].value)
    
    # Extract fire suppression data if this is a FIRE SUPP sheet
    if 'FIRE SUPP' in sheet.title:
        st.write(f"\n🔍 Processing FIRE SUPP sheet: {sheet.title}")
        
        # Initialize sheet totals
        fs_k9_total = 0
        fs_n9_total = 0
        
        # Get K9 and N9 values from cells first
        k9_cell_value = safe_convert_to_float(sheet['K9'].value)
        st.write(f"📊 K9 cell value: {k9_cell_value}")
        
        n9_cell_value = safe_convert_to_float(sheet['N9'].value)
        st.write(f"📊 N9 cell value: {n9_cell_value}")
        
        row = 12  # Starting row for first item
        while row <= 181:
            item_number = sheet[f'B{row}'].value
            
            # Clean up item number by stripping ALL whitespace and converting to string
            if item_number:
                item_number = str(item_number).strip()
            
            # Skip processing if item_number is empty, "ITEM", or "DELIVERY & INSTALLATION"
            if not item_number or item_number.upper().strip() == 'ITEM' or item_number.strip() == 'DELIVERY & INSTALLATION':
                row += 17
                continue
            
            # Get K9 and N9 values for this item using safe conversion
            k9_value = safe_convert_to_float(sheet[f'K{row}'].value)
            fs_k9_total += k9_value
            st.write(f"💰 K9 value for {item_number}: {k9_value}")
            
            n9_value = safe_convert_to_float(sheet[f'N{row}'].value)
            fs_n9_total += n9_value
            st.write(f"💰 N9 value for {item_number}: {n9_value}")
            
            # Get fire suppression installation price
            fs_install = safe_convert_to_float(sheet[f'N182'].value)
            st.write(f"💰 Fire suppression install price: {fs_install}")
            
            fs_item = {
                'item_number': item_number,
                'system_description': 'Ansul R 102 System',
                'manual_release': '1no station',
                'tank_quantity': sheet[f'P{row}'].value or '2',
                'fire_suppression_install': fs_install,
                'k9_value': k9_value,
                'n9_value': n9_value
            }
            st.write(f"✅ Adding fire suppression item: {fs_item}")
            data['fire_suppression_items'].append(fs_item)
            row += 17
        
        # Store sheet totals
        data['fs_k9_total'] = k9_cell_value or fs_k9_total
        data['fs_n9_total'] = n9_cell_value or fs_n9_total
        st.write(f"\n📊 Final totals for sheet {sheet.title}:")
        st.write(f"💰 K9 total: {data['fs_k9_total']}")
        st.write(f"💰 N9 total: {data['fs_n9_total']}")
        st.write(f"📝 Number of fire suppression items: {len(data['fire_suppression_items'])}")

    return data

def read_excel_file(uploaded_file):
    """Read and process uploaded Excel file"""
    workbook = openpyxl.load_workbook(uploaded_file, data_only=True)
    
    # Get all CANOPY and FIRE SUPP sheets, excluding F24 sheets
    canopy_sheets = [sheet for sheet in workbook.sheetnames if 'CANOPY' in sheet and 'F24' not in sheet]
    fire_supp_sheets = [sheet for sheet in workbook.sheetnames if 'FIRE SUPP' in sheet and 'F24' not in sheet]
    
    project_data = {
        'sheets': [],
        'has_water_wash': False,
        'fire_suppression_data': {},
        'global_fs_k9_total': 0,  # Initialize global K9 total
        'global_fs_n9_total': 0   # Initialize global N9 total
    }
    
    # Process each sheet
    for sheet_name in canopy_sheets + fire_supp_sheets:
        sheet = workbook[sheet_name]
        sheet_data = extract_sheet_data(sheet)
        
        # If this is a fire suppression sheet, check if it has valid items before adding to totals
        if 'FIRE SUPP' in sheet_name:
            # Check if the sheet has any valid fire suppression items
            has_valid_items = any(
                item.get('item_number') and 
                str(item['item_number']).strip().upper() != 'ITEM' and
                str(item['item_number']).strip() != 'DELIVERY & INSTALLATION'
                for item in sheet_data['fire_suppression_items']
            )
            
            if has_valid_items:
                project_data['global_fs_k9_total'] += sheet_data.get('fs_k9_total', 0)
                project_data['global_fs_n9_total'] += sheet_data.get('fs_n9_total', 0)
                st.write(f"📊 Adding to global totals from {sheet_name} (has valid items):")
                st.write(f"💰 K9: {sheet_data.get('fs_k9_total', 0):,.2f}")
                st.write(f"💰 N9: {sheet_data.get('fs_n9_total', 0):,.2f}")
            else:
                st.write(f"⚠️ Skipping {sheet_name} - no valid fire suppression items found")
        
        # Only add sheet if it has canopies or is a used fire suppression sheet
        has_canopies = any(
            canopy.get('reference_number') and 
            canopy['reference_number'] != 'ITEM' and 
            canopy.get('model') and 
            canopy['model'] != 'CANOPY TYPE' and
            canopy['reference_number'] != 'DELIVERY & INSTALLATION'
            for canopy in sheet_data['canopies']
        )
        
        if has_canopies or ('FIRE SUPP' in sheet_name and sheet_data['fire_suppression_items']):
            project_data['sheets'].append(sheet_data)
    
    st.write(f"\n📊 Final global fire suppression totals:")
    st.write(f"💰 Total K9: {project_data['global_fs_k9_total']:,.2f}")
    st.write(f"💰 Total N9: {project_data['global_fs_n9_total']:,.2f}")
    
    return project_data

def create_upload_section(col2, key_suffix=""):
    """Handle file upload and data extraction"""
    st.markdown("### Or Upload Existing Project")
    uploaded_file = st.file_uploader(
        "Upload Excel file to extract data", 
        type=['xlsx'],
        key=f"file_uploader_{key_suffix}"
    )
    
    if uploaded_file is not None:
        try:
            project_data = read_excel_file(uploaded_file)
            
            # Generate Word document from uploaded data
            if st.button("Generate Documents", key=f"generate_docs_{key_suffix}"):
                try:
                    # Get data from first sheet for Word doc
                    first_sheet = project_data['sheets'][0]
                    word_data = {
                        'Date': first_sheet['project_info']['date'],
                        'Project Number': first_sheet['project_info']['project_number'],
                        'Sales Contact': first_sheet['project_info']['sales_estimator'].split('/')[0],
                        'Estimator': first_sheet['project_info']['sales_estimator'],
                        'Estimator_Name': first_sheet['project_info']['estimator_name'],
                        'Estimator_Role': first_sheet['project_info']['estimator_role'],
                        'Customer': first_sheet['project_info']['customer'],
                        'Company': first_sheet['project_info']['company']
                    }
                    print(word_data)
                    # Generate Word document
                    word_doc_path = write_to_word_doc(word_data, project_data)
                    
                    # Save uploaded Excel temporarily to modify it
                    temp_excel_path = "temp_" + uploaded_file.name
                    with open(temp_excel_path, "wb") as f:
                        f.write(uploaded_file.getvalue())
                    
                    # Open workbook, write totals, and save
                    wb = openpyxl.load_workbook(temp_excel_path)
                    write_job_total(wb, project_data)
                    wb.save(temp_excel_path)
                    
                    # Create zip with Word doc and modified Excel
                    zip_path = create_download_zip(temp_excel_path, word_doc_path)
                    
                    # Show success message with balloons
                    st.balloons()
                    st.success("🎉 Documents generated successfully! 🎉")
                    
                    # Create download button for zip
                    with open(zip_path, "rb") as fp:
                        zip_contents = fp.read()
                        st.download_button(
                            label="⬇️ Download Documents",
                            data=zip_contents,
                            file_name="project_documents.zip",
                            mime="application/zip",
                            key=f"download_docs_{key_suffix}"
                        )
                    
                    # Clean up temp Excel file
                    os.remove(temp_excel_path)
                    
                except Exception as e:
                    st.error(f"Error generating documents: {str(e)}")
            
            st.success("File uploaded successfully!")
        except Exception as e:
            st.error(f"Error processing file: {str(e)}")

def create_download_zip(excel_file, word_file):
    """Create a zip file containing both Excel and Word documents"""
    # Get project number and date from the Excel file
    wb = openpyxl.load_workbook(excel_file)
    project_number = None
    project_date = None
    
    # Look for project number and date in the first sheet that has 'CANOPY' in its title
    for sheet_name in wb.sheetnames:
        if 'CANOPY' in sheet_name:
            sheet = wb[sheet_name]
            # Get project number from cell C3
            project_number = sheet['C3'].value
            # Get date from cell G7
            project_date = sheet['G7'].value
            break
    
    if not project_number or not project_date:
        raise ValueError("Could not find project number or date in Excel file")
    
    # Format the date from DD/MM/YYYY to DD.MM.YYYY
    try:
        date_obj = datetime.strptime(project_date, "%d/%m/%Y")
        formatted_date = date_obj.strftime("%d.%m.%Y")
    except:
        formatted_date = project_date.replace('/', '.')
    
    # Create base filename
    base_filename = f"{project_number} Cost Sheet {formatted_date}"
    
    # Create zip file
    zip_filename = f"{base_filename}.zip"
    with zipfile.ZipFile(zip_filename, 'w') as zipf:
        # Add Excel file with consistent naming
        zipf.write(excel_file, f"{base_filename}.xlsx")
        # Add Word file with consistent naming
        zipf.write(word_file, f"{base_filename}.docx")
    
    return zip_filename

def write_job_total(workbook, project_data):
    """Write total prices and project info to JOB TOTAL sheet"""
    if 'JOB TOTAL' in workbook.sheetnames:
        total_sheet = workbook['JOB TOTAL']
        
        # Get first sheet data for project info
        first_sheet = project_data['sheets'][0]
        project_info = first_sheet['project_info']
        
        # Write project information
        total_sheet['C3'] = project_info['project_number']
        customer_company = f"{project_info['customer']} ({project_info['company']})" if project_info['company'] else project_info['customer']
        total_sheet['C5'] = customer_company
        total_sheet['C7'] = project_info['sales_estimator']
        
        total_sheet['G3'] = project_info['project_name']
        total_sheet['G5'] = project_info['location']
        total_sheet['G7'] = project_info['date']
        
        # Get revision from first used sheet
        revision = next(sheet['revision'] for sheet in project_data['sheets']
                       if any(canopy['reference_number'] != 'ITEM' 
                             for canopy in sheet['canopies']))
        total_sheet['O7'] = revision
        
        # Initialize totals
        job_total = 0.0
        k9_total = 0.0
        
        st.write("\n🔄 Processing sheets for totals:")
        # Process each sheet
        for sheet_data in project_data['sheets']:
            st.write(f"\n📄 Processing sheet: {sheet_data.get('sheet_name', 'Unknown')}")
            # Only add to total if sheet has canopies (is used)
            if any(canopy['reference_number'] != 'ITEM' and 
                  canopy['model'] != 'CANOPY TYPE' and
                  canopy['reference_number'] != 'DELIVERY & INSTALLATION'
                  for canopy in sheet_data['canopies']):
                
                # Convert string prices to numbers
                try:
                    sheet_total = float(str(sheet_data['total_price']).replace(',', ''))
                    sheet_k9 = float(str(sheet_data['k9_total']).replace(',', ''))
                except (ValueError, TypeError):
                    sheet_total = 0.0
                    sheet_k9 = 0.0
                
                st.write(f"💰 Total price from sheet: {sheet_total}")
                st.write(f"💰 K9 total from sheet: {sheet_k9}")
                
                job_total += sheet_total
                k9_total += sheet_k9
        
        # Convert fire suppression totals to numbers
        try:
            fs_k9_total = float(str(project_data['global_fs_k9_total']).replace(',', ''))
            fs_n9_total = float(str(project_data['global_fs_n9_total']).replace(',', ''))
        except (ValueError, TypeError):
            fs_k9_total = 0.0
            fs_n9_total = 0.0
        
        st.write("\n📊 Final totals:")
        st.write(f"💰 Job total: {job_total}")
        st.write(f"💰 K9 total: {k9_total}")
        st.write(f"🔥 Fire suppression K9 total: {fs_k9_total}")
        st.write(f"🔥 Fire suppression N9 total: {fs_n9_total}")
        
        # Write the totals as numbers
        total_sheet['T16'] = job_total  # Main job total
        total_sheet['S16'] = k9_total   # Main K9 total
        total_sheet['S17'] = fs_k9_total  # Fire suppression K9 total
        total_sheet['T17'] = fs_n9_total  # Fire suppression N9 total

def create_revision_tab():
    st.title("📝 Revise Cost Sheet")
    
    # Create two columns
    col1, col2 = st.columns(2)
    
    with col1:
        st.subheader("Upload Original File")
        uploaded_file = st.file_uploader(
            "Upload Excel file to revise", 
            type=['xlsx'],
            key="revision_tab_uploader"
        )
        
        if uploaded_file is not None:
            try:
                # Read the Excel file
                project_data = read_excel_file(uploaded_file)
                
                # Get project name and number from first sheet
                first_sheet = project_data['sheets'][0]
                project_name = first_sheet['project_info']['project_name']
                project_number = first_sheet['project_info']['project_number']
                current_revision = first_sheet['revision']
                
                # Add new floor/area section
                st.markdown("---")
                st.subheader("Add New Floor/Area")
                new_floor = st.text_input("New Floor Name (e.g., Ground Floor)")
                new_area = st.text_input("New Area Name (e.g., Kitchen)")
                
                if st.button("Add Floor/Area", key="add_floor_area_button") and new_floor and new_area:
                    try:
                        new_filename = add_new_floor_area(uploaded_file, new_floor, new_area, current_revision)
                        st.success(f"Added new floor/area: {new_floor} - {new_area}")
                        
                        # Add download button for the updated file
                        with open(new_filename, "rb") as file:
                            st.download_button(
                                label="Download Updated Excel",
                                data=file,
                                file_name=os.path.basename(new_filename),
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                key="download_updated_excel_button"
                            )
                    except Exception as e:
                        st.error(f"Error adding floor/area: {str(e)}")
                
                # Create new revision section
                st.markdown("---")
                st.subheader("Create New Revision")
                if st.button("Create New Revision", key="create_revision_button"):
                    next_rev = chr(ord(current_revision) + 1)
                    try:
                        new_filename = create_new_revision(uploaded_file, project_name, project_number, next_rev)
                        st.success(f"Created Revision {next_rev}!")
                        
                        # Add download button for the new revision
                        with open(new_filename, "rb") as file:
                            st.download_button(
                                label=f"Download Revision {next_rev}",
                                data=file,
                                file_name=os.path.basename(new_filename),
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                key="download_revision_button"
                            )
                    except Exception as e:
                        st.error(f"Error creating revision: {str(e)}")
                
            except Exception as e:
                st.error(f"Error processing file: {str(e)}")
    
    with col2:
        st.subheader("Generate Documents")
        st.write("Upload an Excel file to generate documents")
        create_upload_section(col2, "revision_tab")

def create_new_revision(uploaded_file, project_name, project_number, next_rev):
    """Create a new revision of the Excel file"""
    # Create temporary file
    temp_excel_path = "temp_" + uploaded_file.name
    with open(temp_excel_path, "wb") as f:
        f.write(uploaded_file.getvalue())
    
    # Open workbook
    wb = openpyxl.load_workbook(temp_excel_path)
    
    # Update revision for each sheet and add dropdowns
    for sheet_name in [s for s in wb.sheetnames if 'CANOPY' in s or s == 'JOB TOTAL']:
        sheet = wb[sheet_name]
        sheet['O7'] = next_rev
        
        # Add dropdowns to CANOPY sheets
        if 'CANOPY' in sheet_name:
            add_dropdowns_to_sheet(wb, sheet, 12)
    
    # Create folder name with project details
    folder_name = f"{project_name} - {project_number} (Revision {next_rev})"
    if not os.path.exists(folder_name):
        os.makedirs(folder_name)
    
    # Save as new file in revision folder
    new_filename = os.path.join(folder_name, uploaded_file.name.replace('.xlsx', f'_Rev{next_rev}.xlsx'))
    wb.save(new_filename)
    
    # Clean up temp file
    os.remove(temp_excel_path)
    
    return new_filename

def add_new_floor_area(uploaded_file, new_floor, new_area, current_revision):
    """Add a new floor/area sheet to the Excel file"""
    temp_excel_path = "temp_" + uploaded_file.name
    with open(temp_excel_path, "wb") as f:
        f.write(uploaded_file.getvalue())
    
    wb = openpyxl.load_workbook(temp_excel_path)
    
    # Find first empty CANOPY sheet
    canopy_sheets = [s for s in wb.sheetnames if 'CANOPY' in s]
    new_sheet = None
    for sheet_name in canopy_sheets:
        sheet = wb[sheet_name]
        if not sheet['B12'].value:  # Check if sheet is empty (no canopies)
            new_sheet = sheet
            break
    
    if new_sheet:
        # Update sheet name and title
        new_sheet_name = f"CANOPY - {new_floor} ({len(canopy_sheets) + 1})"
        new_sheet.title = new_sheet_name
        new_sheet['B1'] = f"{new_floor} - {new_area}"
        new_sheet['O7'] = current_revision
        
        # Add dropdowns to the new sheet
        add_dropdowns_to_sheet(wb, new_sheet, 12)
        
        # Save workbook
        new_filename = uploaded_file.name.replace('.xlsx', '_updated.xlsx')
        wb.save(new_filename)
        os.remove(temp_excel_path)
        return new_filename
    else:
        os.remove(temp_excel_path)
        raise Exception("No empty CANOPY sheets available")

def edit_floor_area_name(uploaded_file, selected_sheet, new_name, current_revision):
    """Edit the name of an existing floor/area"""
    temp_excel_path = "temp_" + uploaded_file.name
    with open(temp_excel_path, "wb") as f:
        f.write(uploaded_file.getvalue())
    
    wb = openpyxl.load_workbook(temp_excel_path)
    sheet = wb[selected_sheet]
    
    # Update sheet title and name
    sheet['B1'] = new_name
    sheet.title = f"CANOPY - {new_name}"
    
    # Save workbook
    new_filename = uploaded_file.name.replace('.xlsx', '_updated.xlsx')
    wb.save(new_filename)
    os.remove(temp_excel_path)
    return new_filename

def add_new_canopy(uploaded_file, selected_sheet, ref_number, model, config, current_revision):
    """Add a new canopy to an existing area"""
    temp_excel_path = "temp_" + uploaded_file.name
    with open(temp_excel_path, "wb") as f:
        f.write(uploaded_file.getvalue())
    
    wb = openpyxl.load_workbook(temp_excel_path)
    sheet = wb[selected_sheet]
    
    # Find first empty canopy slot
    row = 12
    while row <= 181:
        if not sheet[f'B{row}'].value:
            # Write canopy data
            sheet[f'B{row}'] = ref_number
            sheet[f'C{row + 2}'] = config
            sheet[f'D{row + 2}'] = model
            
            # Add standard entries
            sheet[f'C{row + 3}'] = "LIGHT SELECTION"
            sheet[f'C{row + 4}'] = "SELECT WORKS"
            sheet[f'C{row + 5}'] = "SELECT WORKS"
            sheet[f'C{row + 6}'] = "BIM/ REVIT per CANOPY"
            sheet[f'D{row + 6}'] = "1"
            break
        row += 17
    
    # Save workbook
    new_filename = uploaded_file.name.replace('.xlsx', '_updated.xlsx')
    wb.save(new_filename)
    os.remove(temp_excel_path)
    return new_filename

def edit_canopy(uploaded_file, selected_sheet, selected_canopy, new_model, new_config, current_revision):
    """Edit an existing canopy's properties"""
    temp_excel_path = "temp_" + uploaded_file.name
    with open(temp_excel_path, "wb") as f:
        f.write(uploaded_file.getvalue())
    
    wb = openpyxl.load_workbook(temp_excel_path)
    sheet = wb[selected_sheet]
    
    # Find the canopy
    row = 12
    while row <= 181:
        if sheet[f'B{row}'].value == selected_canopy:
            # Update canopy data
            sheet[f'C{row + 2}'] = new_config
            sheet[f'D{row + 2}'] = new_model
            break
        row += 17
    
    # Save workbook
    new_filename = uploaded_file.name.replace('.xlsx', '_updated.xlsx')
    wb.save(new_filename)
    os.remove(temp_excel_path)
    return new_filename

def update_cladding(uploaded_file, selected_sheet, selected_canopy, width, height, positions, current_revision):
    """Update wall cladding for a canopy"""
    temp_excel_path = "temp_" + uploaded_file.name
    with open(temp_excel_path, "wb") as f:
        f.write(uploaded_file.getvalue())
    
    wb = openpyxl.load_workbook(temp_excel_path)
    sheet = wb[selected_sheet]
    
    # Find the canopy
    row = 12
    while row <= 181:
        if sheet[f'B{row}'].value == selected_canopy:
            cladding_row = row + 7
            # Update cladding data
            sheet[f'C{cladding_row}'] = "2M² (HFL)"
            sheet[f'Q{cladding_row}'] = width
            sheet[f'R{cladding_row}'] = height
            sheet[f'S{cladding_row}'] = ','.join(positions)
            
            # Create formatted display string
            positions_str = '/'.join(positions)
            cladding_info = f"2M² (HFL) - {width}x{height}mm ({positions_str})"
            sheet[f'T{cladding_row}'] = cladding_info
            break
        row += 17
    
    # Save workbook
    new_filename = uploaded_file.name.replace('.xlsx', '_updated.xlsx')
    wb.save(new_filename)
    os.remove(temp_excel_path)
    return new_filename

def delete_canopy(uploaded_file, selected_sheet, selected_canopy, current_revision):
    """Delete a canopy from an area"""
    temp_excel_path = "temp_" + uploaded_file.name
    with open(temp_excel_path, "wb") as f:
        f.write(uploaded_file.getvalue())
    
    wb = openpyxl.load_workbook(temp_excel_path)
    sheet = wb[selected_sheet]
    
    # Find the canopy
    row = 12
    while row <= 181:
        if sheet[f'B{row}'].value == selected_canopy:
            # Clear canopy data (17 rows)
            for r in range(row, row + 17):
                for col in ['B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T']:
                    sheet[f'{col}{r}'].value = None
            break
        row += 17
    
    # Save workbook
    new_filename = uploaded_file.name.replace('.xlsx', '_updated.xlsx')
    wb.save(new_filename)
    os.remove(temp_excel_path)
    return new_filename

def reorder_canopies(uploaded_file, selected_sheet, new_order, current_revision):
    """Reorder canopies in an area"""
    temp_excel_path = "temp_" + uploaded_file.name
    with open(temp_excel_path, "wb") as f:
        f.write(uploaded_file.getvalue())
    
    wb = openpyxl.load_workbook(temp_excel_path)
    sheet = wb[selected_sheet]
    
    # Store current canopy data
    canopy_data = {}
    row = 12
    while row <= 181:
        ref = sheet[f'B{row}'].value
        if ref and ref != 'ITEM':
            # Store all values for this canopy section (17 rows)
            canopy_data[ref] = {}
            for r in range(row, row + 17):
                canopy_data[ref][r-row] = {}
                for col in ['B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T']:
                    canopy_data[ref][r-row][col] = sheet[f'{col}{r}'].value
        row += 17
    
    # Clear all canopy data
    row = 12
    while row <= 181:
        for r in range(row, row + 17):
            for col in ['B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T']:
                sheet[f'{col}{r}'].value = None
        row += 17
    
    # Write canopies in new order
    row = 12
    for ref in new_order:
        if ref in canopy_data:
            for r_offset, col_data in canopy_data[ref].items():
                for col, value in col_data.items():
                    sheet[f'{col}{row + r_offset}'].value = value
        row += 17
    
    # Save workbook
    new_filename = uploaded_file.name.replace('.xlsx', '_updated.xlsx')
    wb.save(new_filename)
    os.remove(temp_excel_path)
    return new_filename

def copy_area_to_new_floor(uploaded_file, source_sheet, new_floor, new_area, current_revision):
    """Copy an area to a new floor"""
    temp_excel_path = "temp_" + uploaded_file.name
    with open(temp_excel_path, "wb") as f:
        f.write(uploaded_file.getvalue())
    
    wb = openpyxl.load_workbook(temp_excel_path)
    source = wb[source_sheet]
    
    # Find first empty CANOPY sheet
    canopy_sheets = [s for s in wb.sheetnames if 'CANOPY' in s]
    target = None
    for sheet_name in canopy_sheets:
        sheet = wb[sheet_name]
        if not sheet['B12'].value:  # Check if sheet is empty
            target = sheet
            break
    
    if target:
        # Copy all cells from source to target
        for row in source.rows:
            for cell in row:
                target[cell.coordinate].value = cell.value
        
        # Update sheet name and title
        target.title = f"CANOPY - {new_floor} ({len(canopy_sheets) + 1})"
        target['B1'] = f"{new_floor} - {new_area}"
        target['O7'] = current_revision
        
        # Save workbook
        new_filename = uploaded_file.name.replace('.xlsx', '_updated.xlsx')
        wb.save(new_filename)
        os.remove(temp_excel_path)
        return new_filename
    else:
        os.remove(temp_excel_path)
        raise Exception("No empty CANOPY sheets available")

def generate_word_document(project_data):
    """Generate Word document from project data"""
    # Get data from first sheet
    first_sheet = project_data['sheets'][0]
    word_data = {
        'Date': first_sheet['project_info']['date'],
        'Project Number': first_sheet['project_info']['project_number'],
        'Sales Contact': first_sheet['project_info']['sales_estimator'].split('/')[0],
        'Estimator': first_sheet['project_info']['sales_estimator'],
        'Estimator_Name': first_sheet['project_info']['estimator_name'],
        'Estimator_Role': first_sheet['project_info']['estimator_role'],
        'Customer': first_sheet['project_info']['customer'],
        'Company': first_sheet['project_info']['company']
    }
    
    # Generate Word document
    return write_to_word_doc(word_data, project_data)

def add_fire_suppression_dropdown(sheet):
    """Add dropdown validation for fire suppression options"""
    if 'Lists' not in sheet.parent.sheetnames:
        list_sheet = sheet.parent.create_sheet('Lists')
    else:
        list_sheet = sheet.parent['Lists']
    
    # Add fire suppression options to Lists sheet (column H)
    fire_suppression_options = [
        "SELECT TANK SYSTEM",
        "1 TANK SYSTEM",
        "1 TANK TRAVEL HUB",
        "1 TANK DISTANCE",
        "NOBEL",
        "AMAREX",
        "OTHER",
        "2 TANK SYSTEM",
        "2 TANK TRAVEL HUB",
        "2 TANK DISTANCE",
        "3 TANK SYSTEM",
        "3 TANK TRAVEL HUB",
        "3 TANK DISTANCE",
        "4 TANK SYSTEM",
        "4 TANK TRAVEL HUB",
        "4 TANK DISTANCE",
        "5 TANK SYSTEM",
        "5 TANK TRAVEL HUB",
        "5 TANK DISTANCE",
        "6 TANK SYSTEM",
        "6 TANK TRAVEL HUB",
        "6 TANK DISTANCE"
    ]
    
    # Add tank installation options to Lists sheet (column I)
    tank_install_options = [
        "1 TANK",
        "1 TANK DISTANCE",
        "2 TANK",
        "2 TANK DISTANCE",
        "3 TANK",
        "3 TANK DISTANCE",
        "4 TANK",
        "4 TANK DISTANCE",
        "5 TANK",
        "5 TANK DISTANCE",
        "6 TANK",
        "6 TANK DISTANCE"
    ]
    
    # Write fire suppression options to Lists sheet
    for i, option in enumerate(fire_suppression_options, 1):
        list_sheet[f'H{i}'] = option
        
    # Write tank installation options to Lists sheet
    for i, option in enumerate(tank_install_options, 1):
        list_sheet[f'I{i}'] = option
    
    # Create validation for fire suppression
    dv_fs = DataValidation(
        type="list",
        formula1=f"Lists!$H$1:$H${len(fire_suppression_options)}",
        allow_blank=True
    )
    sheet.add_data_validation(dv_fs)
    
    # Create validation for tank installation
    dv_tank = DataValidation(
        type="list",
        formula1=f"Lists!$I$1:$I${len(tank_install_options)}",
        allow_blank=True
    )
    sheet.add_data_validation(dv_tank)
    
    # Add fire suppression validation to C16 and every 17 rows after
    row = 16
    while row <= sheet.max_row:
        dv_fs.add(f"C{row}")
        row += 17
    
    # Add tank installation validation to C17 and every 17 rows after
    row = 17
    while row <= sheet.max_row:
        dv_tank.add(f"C{row}")
        row += 17

def main():
    st.set_page_config(page_title="Project Information Form", layout="wide")
    
    # Create tabs
    tab1, tab2 = st.tabs(["Create New Project", "Revise Project"])
    
    with tab1:
        create_general_info_form()
    
    with tab2:
        create_revision_tab()

if __name__ == "__main__":
    main() 