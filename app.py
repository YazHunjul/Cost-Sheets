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
                                    'canopies': canopies_data
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
            if st.button("Save Project", disabled=not (all_fields_filled and has_canopies)):
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

def write_to_sheet(sheet, data, level_name, area_name, canopies):
    """Write data to specific cells in the sheet"""
    # Write sheet title
    sheet['B1'] = f"{level_name} - {area_name}"
    
    # Write general info (only if cells don't contain formulas)
    def safe_write(cell, value):
        # Skip writing if value is "Select..." or None/empty
        if value == "Select..." or not value:
            return
        # Skip writing if cell contains formula
        if not cell.value or not str(cell.value).startswith('='):
            cell.value = value
    
    safe_write(sheet['C3'], data['Project Number'])
    # Combine customer and company in C5
    customer_company = f"{data['Customer']} ({data['Company']})" if data['Company'] else data['Customer']
    safe_write(sheet['C5'], customer_company)
    
    sales_initials = get_initials(data['Sales Contact'])
    estimator_initials = get_initials(data['Estimator'])
    if sales_initials and estimator_initials:  # Only write if both have valid values
        safe_write(sheet['C7'], f"{sales_initials}/{estimator_initials}")
    
    safe_write(sheet['G3'], data['Project Name'])
    safe_write(sheet['G5'], data['Location'])
    safe_write(sheet['G7'], data['Date'])
    
    # Add initial revision 'A'
    safe_write(sheet['O7'], 'A')
    
    # Store estimator info in Lists sheet
    if 'Lists' not in sheet.parent.sheetnames:
        list_sheet = sheet.parent.create_sheet('Lists')
    else:
        list_sheet = sheet.parent['Lists']
    
    # Store estimator name and role in Lists sheet
    estimator_name = data['Estimator']
    if estimator_name in estimators:
        list_sheet['Z1'] = estimator_name  # Store full name
        list_sheet['Z2'] = estimators[estimator_name]  # Store role
    
    # Write canopy data
    for idx, canopy in enumerate(canopies):
        base_row = 12 + (idx * 17)
        
        # Only write if values are not "Select..."
        if canopy['reference_number'] != "Select...":
            safe_write(sheet[f'B{base_row}'], canopy['reference_number'])
        if canopy['configuration'] != "Select...":
            safe_write(sheet[f'C{base_row + 2}'], canopy['configuration'])
        if canopy['model'] != "Select...":
            safe_write(sheet[f'D{base_row + 2}'], canopy['model'])
        
        # Standard entries
        safe_write(sheet[f'C{base_row + 3}'], "LIGHT SELECTION")
        safe_write(sheet[f'C{base_row + 4}'], "SELECT WORKS")
        safe_write(sheet[f'C{base_row + 5}'], "SELECT WORKS")
        safe_write(sheet[f'C{base_row + 6}'], "BIM/ REVIT per CANOPY")
        safe_write(sheet[f'D{base_row + 6}'], "1")
        
        # Wall cladding - only write if not "Select..."
        if canopy['wall_cladding']['type'] and canopy['wall_cladding']['type'] != "Select...":
            cladding_row = base_row + 7
            # Write wall cladding type to C column
            safe_write(sheet[f'C{cladding_row}'], canopy['wall_cladding']['type'])
            
            # Write dimensions and positions to hidden cells for storage
            # Using columns Q, R, S, T for storage (these are typically hidden)
            safe_write(sheet[f'Q{cladding_row}'], canopy['wall_cladding']['width'])  # Width in Q
            safe_write(sheet[f'R{cladding_row}'], canopy['wall_cladding']['height'])  # Height in R
            safe_write(sheet[f'S{cladding_row}'], ','.join(canopy['wall_cladding']['positions']))  # Positions in S
            
            # Create a formatted display string for the visible cell
            positions_str = '/'.join(canopy['wall_cladding']['positions'])
            cladding_info = (f"{canopy['wall_cladding']['type']} - "
                            f"{canopy['wall_cladding']['width']}x{canopy['wall_cladding']['height']}mm "
                            f"({positions_str})")
            safe_write(sheet[f'T{cladding_row}'], cladding_info)  # Store full formatted string
        
        # Write MUA VOL to Q22 if it's an F-type canopy
        model = str(canopy.get('model', ''))  # Convert to string first
        if 'F' in model.upper():
            print(f"Writing MUA VOL for {model}: {canopy.get('mua_vol', '-')}")
            sheet[f'Q{22 + (idx * 17)}'] = f"MUA VOL: {canopy.get('mua_vol', '-')}"

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
    
    # Create Halton reference with project number and both sets of initials
    halton_ref = f"{data['Project Number']} / {get_initials(sales_contact_name)} / {estimator_initials}"
    
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
        # Use sheet name from B1 that we stored earlier
        display_name = sheet['sheet_name']  # This now contains the B1 value
        
        # Collect canopy data for technical table
        canopies = []
        has_uv_canopy = False
        
        # Get cladding items for this area
        area_cladding_items = [item for item in cladding_items 
                             if any(c['reference_number'] == item['item_no'] 
                                   for c in sheet['canopies'])]
        
        for canopy in sheet['canopies']:
            if (canopy['reference_number'] != 'ITEM' and 
                canopy['model'] != 'CANOPY TYPE'):
                
                # Convert model and configuration to strings
                model = str(canopy.get('model', '')) if canopy.get('model') is not None else ''
                config = str(canopy.get('configuration', '')) if canopy.get('configuration') is not None else ''
                
                # Check for UV canopies
                if ('UV' in model.upper() or 'UV' in config.upper()):
                    has_uv_canopy = True
                
                # Check for water wash canopies
               # if 'CMWI' in model.upper() or 'CMWF' in model.upper():
                   # has_water_wash_canopies = True
                   # ww_canopies.append(canopy)
                
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
                    'lighting': canopy.get('lights', 'LED Strip'),
                    'cws_2bar': canopy['cws_2bar'] or '-',
                    'hws_2bar': canopy['hws_2bar'] or '-',
                    'hws_storage': canopy['hws_storage'] or '-',
                    'ww_price': canopy['ww_price'] or 0,
                    'ww_control_price': canopy['ww_control_price'] or 0,
                    'ww_install_price': canopy['ww_install_price'] or 0,
                    'base_price': canopy['base_price'] or 0  # Keep base_price
                })
        
        if canopies:
            # Calculate area totals with rounding at each step
            canopy_total = sum(math.ceil(canopy['base_price']) for canopy in canopies)  # Round each canopy price
            delivery_total = math.ceil(float(sheet['delivery_install']))  # Round delivery
            commissioning_total = float(sheet['commissioning_price'])  # Get raw value
            area_total = math.ceil(canopy_total + delivery_total + commissioning_total)  # Include commissioning in total
            
            areas.append({
                'name': display_name,
                'canopies': [{**canopy, 'base_price': f"{math.ceil(canopy['base_price']):,}"} for canopy in canopies],
                'has_uv': has_uv_canopy,
                'has_cladding': bool(area_cladding_items),
                'cladding_items': [{**item, 'price': f"{math.ceil(item['price']):,}"} for item in area_cladding_items],
                'canopy_total': f"{canopy_total:,}",
                'delivery_total': f"{delivery_total:,}",
                'commissioning_total': f"{commissioning_total:,.2f}",  # Format with 2 decimal places
                'cladding_total': f"{sum(math.ceil(item['price']) for item in area_cladding_items):,}",
                'uv_total': "1,040.00" if has_uv_canopy else "0",
                'area_total': f"{area_total:,}",
                'total_price': f"{math.ceil(float(sheet['total_price'] or 0)):,}"
            })
    
    # Calculate totals from all sheets
    job_total = 0
    k9_total = 0
    commissioning_total = 0  # Add commissioning total tracking
    
    for sheet in project_data['sheets']:
        # Only add to total if sheet has canopies (is used)
        if any(canopy['reference_number'] != 'ITEM' and 
               canopy['model'] != 'CANOPY TYPE' and
               canopy['reference_number'] != 'DELIVERY & INSTALLATION'
               for canopy in sheet['canopies']):
            job_total += sheet['total_price']  # N9 value
            k9_total += sheet['k9_total']  # K9 value
            commissioning_total += sheet['commissioning_price']  # Add commissioning to total
    
    # Format totals with commas and 2 decimal places
    job_total_formatted = f"{job_total:,.2f}"
    k9_total_formatted = f"{k9_total:,.2f}"
    commissioning_total_formatted = f"{commissioning_total:,.2f}"  # Format with 2 decimal places
    
    # Add to context
    context = {
        'date': data['Date'],
        'project_number': halton_ref,
        'sales_contact_name': sales_contact_name,
        'contact_number': contact_number,
        'customer': data['Customer'],
        'customer_first_name': customer_first_name,
        'company': data.get('Company', ''),
        'estimator_name': project_data['sheets'][0]['project_info'].get('estimator_name', ''),
        'estimator_role': project_data['sheets'][0]['project_info'].get('estimator_role', ''),
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
        scope_items.append(f"{count} x {canopy_type} Ventilation Canopies")
    
    # Add any wall cladding from valid canopies only
    wall_cladding_count = sum(1 for sheet in project_data['sheets'] 
                             for canopy in sheet['canopies'] 
                             if (canopy['wall_cladding']['type'] and 
                                 canopy['wall_cladding']['type'] != "Select..." and 
                                 canopy['wall_cladding']['type'] == "2M² (HFL)" and
                                 canopy['reference_number'] != 'ITEM' and
                                 canopy['model'] != 'CANOPY TYPE'))
    if wall_cladding_count > 0:
        scope_items.append(f"{wall_cladding_count} x Areas with Stainless Steel Cladding")
    
    # Add scope items to context
    context['scope_items'] = scope_items
    
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
        
        if not os.path.exists(template_path):
            st.error(f"Template file not found at: {template_path}")
            return
            
        # Load workbook with data_only=False to preserve formulas
        workbook = openpyxl.load_workbook(template_path, read_only=False, data_only=False)
        
        # Get all CANOPY sheets
        canopy_sheets = [sheet for sheet in workbook.sheetnames if 'CANOPY' in sheet]
        sheet_count = 0
        
        # Process each level and area
        for level in data['Levels']:
            for area in level['areas']:
                # Get the appropriate CANOPY sheet
                if sheet_count >= len(canopy_sheets):
                    st.error(f"Not enough CANOPY sheets in template! Need {sheet_count + 1}, but only have {len(canopy_sheets)}")
                    break
                
                sheet_name = canopy_sheets[sheet_count]
                current_sheet = workbook[sheet_name]
                
                # Make sure the sheet is visible
                current_sheet.sheet_state = 'visible'
                
                # Rename the sheet with new format: CANOPY - FLOOR NAME (#)
                new_sheet_name = f"CANOPY - {level['level_name']} ({sheet_count + 1})"
                current_sheet.title = new_sheet_name
                
                # Write data to the existing sheet
                write_to_sheet(current_sheet, data, level['level_name'], area['area_name'], area['canopies'])
                
                # Add dropdowns to the sheet
                add_dropdowns_to_sheet(workbook, current_sheet, 12)
                
                sheet_count += 1
        
        # Save the workbook
        workbook.save(output_path)
        st.success(f"Successfully updated {sheet_count} CANOPY sheets!")
        
        # Provide download button
        with open(output_path, "rb") as file:
            st.download_button(
                label="Download Excel file",
                data=file,
                file_name="project_data.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
    except FileNotFoundError:
        st.error("Template file not found in resources folder!")
    except Exception as e:
        st.error(f"An error occurred: {str(e)}")
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
    
    # Get delivery and installation price from P182
    delivery_install_price = sheet['P182'].value or 0
    
    # Get K9 total from sheet
    k9_total = math.ceil(float(sheet['K9'].value or 0)) if sheet['K9'].value is not None else 0
    
    # Get commissioning price from N193 (using same pattern as delivery price)
    commissioning_price = sheet['N193'].value or 0
    
    data = {
        'sheet_name': display_name,  # Use B1 value instead of sheet title
        'revision': sheet['O7'].value,
        'delivery_install': delivery_install_price,
        'total_price': math.ceil(float(sheet['N9'].value or 0)) if sheet['N9'].value is not None else 0,
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
        'canopies': []
    }
    
    # Extract canopy data - only process up to row 181 for canopies
    row = 12  # Starting row for first canopy
    while row <= 181:  # Processing canopies
        if (sheet[f'B{row}'].value and 
            sheet[f'C{row + 2}'].value and sheet[f'C{row + 2}'].value != "Select..." and
            sheet[f'D{row + 2}'].value and sheet[f'D{row + 2}'].value != "Select..."):
            
            dim_row = row + 2
            pa_row = row + 10  # F22 for first canopy
            cladding_row = row + 7
            water_row = row + 13  # F25 for first canopy water wash values
            
            # Get prices
            canopy_price = sheet[f'P{row}'].value or 0  # Base canopy price  # This is P12 for first canopy
            canopy_price = math.ceil(float(canopy_price))  # Round up to nearest pound
            # Don't convert to string yet - keep as number for calculations
            
            cladding_price = sheet[f'N{row + 7}'].value or 0  # Cladding price if exists
            emergency_lighting = sheet[f'P{row + 1}'].value or "2No, @ £100.00 if required"  # Emergency lighting text
            
            # Get water wash values
            cws_2bar = sheet[f'F{row + 13}'].value  # F25 - Cold water supply at 2 bar
            hws_2bar = sheet[f'F{row + 14}'].value  # F26 - Hot water supply at 2 bar
            hws_storage = sheet[f'F{row + 15}'].value  # F27 - Hot water storage
            
            # Get water wash prices
            ww_price = sheet[f'P{row + 13}'].value or 0  # P25 - Water wash price
            ww_control_price = sheet[f'P{row + 14}'].value or 0  # P26 - Control panel price
            ww_install_price = sheet[f'P{row + 15}'].value or 0  # P27 - Installation price
            
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
                    mua_vol_cell = sheet[f'Q{22 + ((row - 12) // 17) * 17}'].value or ''
                    if mua_vol_cell and 'MUA VOL:' in mua_vol_cell:
                        mua_vol = mua_vol_cell.replace('MUA VOL:', '').strip()
                    else:
                        mua_vol = '-'
            
            canopy = {
                'reference_number': sheet[f'B{row}'].value or '-',
                'model': sheet[f'D{row + 2}'].value or '-',
                'configuration': sheet[f'C{row + 2}'].value or '-',
                'length': sheet[f'E{dim_row}'].value or '-',
                'width': sheet[f'F{dim_row}'].value or '-',
                'height': sheet[f'G{dim_row}'].value or '-',
                'sections': sheet[f'H{dim_row}'].value or '-',
                'ext_vol': sheet[f'I{dim_row}'].value or '-',
                'ext_static': sheet[f'F{pa_row}'].value or '-',
                'mua_vol': mua_vol,  # Set MUA VOL from Excel for F-type canopies
                'supply_static': sheet[f'L{dim_row}'].value or '-',
                'lighting': sheet[f'C{row + 3}'].value or '-',
                'special_works_1': sheet[f'C{row + 4}'].value or '-',
                'special_works_2': sheet[f'C{row + 5}'].value or '-',
                'bim_revit': sheet[f'C{row + 6}'].value or '-',
                'wall_cladding': wall_cladding,
                'price': canopy_price,
                'base_price': canopy_price,
                'emergency_lighting': emergency_lighting,
                'cws_2bar': cws_2bar,
                'hws_2bar': hws_2bar,
                'hws_storage': hws_storage,
                'ww_price': ww_price,
                'ww_control_price': ww_control_price,
                'ww_install_price': ww_install_price
            }
            data['canopies'].append(canopy)
        row += 17  # Move to next canopy section
    
    # Get sheet totals
    try:
        total_price_cell = sheet['N9'].value
        k9_total_cell = sheet['K9'].value
        delivery_install_cell = sheet['P182'].value
        commissioning_price_cell = sheet['N193'].value
        
        data['total_price'] = math.ceil(float(total_price_cell)) if total_price_cell is not None else 0
        data['k9_total'] = math.ceil(float(k9_total_cell)) if k9_total_cell is not None else 0
        data['delivery_install'] = math.ceil(float(delivery_install_cell)) if delivery_install_cell is not None else 0
        data['commissioning_price'] = float(commissioning_price_cell) if commissioning_price_cell is not None else 0
        
    except (ValueError, TypeError) as e:
        data['total_price'] = 0
        data['k9_total'] = 0
        data['delivery_install'] = 0
        data['commissioning_price'] = 0
    
    return data

def read_excel_file(uploaded_file):
    """Read and process uploaded Excel file"""
    workbook = openpyxl.load_workbook(uploaded_file, data_only=True)
    
    # Get all CANOPY sheets
    canopy_sheets = [sheet for sheet in workbook.sheetnames if 'CANOPY' in sheet]
    
    project_data = {
        'sheets': [],
        'has_water_wash': False  # Initialize global flag
    }
    
    # Process each CANOPY sheet
    for sheet_name in canopy_sheets:
        sheet = workbook[sheet_name]
        sheet_data = extract_sheet_data(sheet)
        
        # Check for water wash canopies
        for canopy in sheet_data['canopies']:
            if any(model in canopy['model'].upper() for model in ['CMWI', 'CMWF']):
                project_data['has_water_wash'] = True  # Set global flag
                break
        
        project_data['sheets'].append(sheet_data)
    
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
            if st.button("Generate Documents"):
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
                    zip_path = create_download_zip(word_doc_path, temp_excel_path, uploaded_file.name)
                    
                    # Create download button for zip
                    with open(zip_path, "rb") as fp:
                        zip_contents = fp.read()
                        st.download_button(
                            label="Download Documents",
                            data=zip_contents,
                            file_name="project_documents.zip",
                            mime="application/zip"
                        )
                    
                    # Clean up temp Excel file
                    os.remove(temp_excel_path)
                    
                except Exception as e:
                    st.error(f"Error generating documents: {str(e)}")
            
            st.success("File uploaded successfully!")
        except Exception as e:
            st.error(f"Error processing file: {str(e)}")

def create_download_zip(word_doc_path, excel_path, excel_filename, zip_path="output.zip"):
    """Create a zip file containing the Word document and Excel file"""
    with zipfile.ZipFile(zip_path, 'w') as zipf:
        # Add Word document
        zipf.write(word_doc_path, os.path.basename(word_doc_path))
        # Add Excel file
        zipf.write(excel_path, excel_filename)
    return zip_path

def write_job_total(workbook, project_data):
    """Write total prices and project info to JOB TOTAL sheet"""
    if 'JOB TOTAL' in workbook.sheetnames:
        total_sheet = workbook['JOB TOTAL']
        
        # Get first sheet data for project info
        first_sheet = project_data['sheets'][0]
        project_info = first_sheet['project_info']
        
        # Write project information
        total_sheet['C3'] = project_info['project_number']
        # Combine customer and company
        customer_company = f"{project_info['customer']} ({project_info['company']})" if project_info['company'] else project_info['customer']
        total_sheet['C5'] = customer_company
        total_sheet['C7'] = project_info['sales_estimator']  # Sales/Estimator initials
        
        total_sheet['G3'] = project_info['project_name']
        total_sheet['G5'] = project_info['location']
        total_sheet['G7'] = project_info['date']
        
        # Get revision from first used sheet
        revision = next(sheet['revision'] for sheet in project_data['sheets']
                       if any(canopy['reference_number'] != 'ITEM' 
                             for canopy in sheet['canopies']))
        total_sheet['O7'] = revision  # Write revision to O7 cell
        
        # Calculate total from all used sheets (N9 values)
        job_total = 0
        k9_total = 0  # Add K9 total tracking
        
        for sheet_data in project_data['sheets']:
            # Only add to total if sheet has canopies (is used)
            if any(canopy['reference_number'] != 'ITEM' and 
                  canopy['model'] != 'CANOPY TYPE' and
                  canopy['reference_number'] != 'DELIVERY & INSTALLATION'
                  for canopy in sheet_data['canopies']):
                job_total += sheet_data['total_price']  # N9 value already stored in total_price
                
                # Get K9 value from sheet and add to k9_total
                if 'k9_total' in sheet_data:
                    k9_total += sheet_data['k9_total']
        
        # Write the totals
        total_sheet['T16'] = job_total
        total_sheet['S16'] = k9_total  # Write K9 total to S16

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
                
                if st.button("Add Floor/Area") and new_floor and new_area:
                    try:
                        new_filename = add_new_floor_area(uploaded_file, new_floor, new_area, current_revision)
                        st.success(f"Added new floor/area: {new_floor} - {new_area}")
                        
                        # Add download button for the updated file
                        with open(new_filename, "rb") as file:
                            st.download_button(
                                label="Download Updated Excel",
                                data=file,
                                file_name=os.path.basename(new_filename),
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                            )
                    except Exception as e:
                        st.error(f"Error adding floor/area: {str(e)}")
                
                # Create new revision section
                st.markdown("---")
                st.subheader("Create New Revision")
                if st.button("Create New Revision"):
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
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
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
    
    # Update revision for each sheet
    for sheet_name in [s for s in wb.sheetnames if 'CANOPY' in s or s == 'JOB TOTAL']:
        sheet = wb[sheet_name]
        sheet['O7'] = next_rev
    
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