"""
IFC Parser - Extracts ALL parameters from IFC elements WITHOUT unit conversion and WITHOUT subtracting openings.
Generates Excel report with TWO sheets:
  1. Summary - Model info (project name, stories count, building height, address, total elements count)
  2. Elements - All elements with each parameter in a separate column.
     Format: Columns = "Element Specific:Long Name", "Element Specific:Name", "Ifc Class", 
             "PSetName:PropertyName" for all properties.
     Rows = values for each element, plus rows for IfcBuilding and IfcBuildingStorey.
"""

import ifcopenshell
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from datetime import datetime
import os


# ----------------------------------------------------------------------
# Helper functions for extracting property and quantity values
# ----------------------------------------------------------------------

def get_property_value(prop):
    """Extract value from IfcProperty of any type."""
    try:
        if prop.is_a('IfcPropertySingleValue'):
            val = prop.NominalValue
            if val is None:
                return None
            return val.wrappedValue if hasattr(val, 'wrappedValue') else val

        elif prop.is_a('IfcPropertyEnumeratedValue'):
            vals = prop.EnumerationValues
            return ', '.join(str(v) for v in vals) if vals else None

        elif prop.is_a('IfcPropertyListValue'):
            vals = prop.ListValues
            return ', '.join(str(v) for v in vals) if vals else None

        elif prop.is_a('IfcPropertyBoundedValue'):
            lower = prop.LowerBoundValue
            upper = prop.UpperBoundValue
            if lower and upper:
                return f"{lower} – {upper}"
            return lower or upper

        else:   # Other property types - string representation
            if hasattr(prop, 'NominalValue') and prop.NominalValue is not None:
                val = prop.NominalValue
                return val.wrappedValue if hasattr(val, 'wrappedValue') else val
            return str(prop)
    except:
        return None


def get_quantity_value(qty):
    """Extract numeric value from IfcPhysicalQuantity WITHOUT unit conversion."""
    try:
        if qty.is_a('IfcQuantityLength'):
            val = qty.LengthValue
        elif qty.is_a('IfcQuantityArea'):
            val = qty.AreaValue
        elif qty.is_a('IfcQuantityVolume'):
            val = qty.VolumeValue
        elif qty.is_a('IfcQuantityCount'):
            val = qty.CountValue
        elif qty.is_a('IfcQuantityWeight'):
            val = qty.WeightValue
        else:
            val = getattr(qty, 'NominalValue', None)

        if val is None:
            return None
        return val.wrappedValue if hasattr(val, 'wrappedValue') else val
    except:
        return None


def get_element_storey(element, ifc_file):
    """Get the storey/building name that contains this element."""
    try:
        inverses = ifc_file.get_inverse(element)
        for inv in inverses:
            if inv.is_a() == 'IfcRelContainedInSpatialStructure':
                relating = getattr(inv, 'RelatingStructure', None)
                if relating:
                    return getattr(relating, 'Name', None) or getattr(relating, 'LongName', None)
    except:
        pass
    return None


def get_element_properties(element, ifc_file, include_type_props=True):
    """Collect all element parameters (attributes, PropertySet, Quantity, Type) WITHOUT unit conversion.
    Returns dict with keys like 'Element Specific:Long Name', 'PSetName:PropertyName', etc.
    """
    props = {}
    
    # Storey/Building containment
    storey_name = get_element_storey(element, ifc_file)
    if storey_name:
        props['Storey'] = storey_name

    # Basic IfcElement attributes with "Element Specific:" prefix
    base_attrs = ['GlobalId', 'Name', 'Description', 'Tag', 'ObjectType', 'LongName', 'PredefinedType']
    for attr in base_attrs:
        val = getattr(element, attr, None)
        if val is not None:
            props[f'Element Specific:{attr}'] = val
    
    # IFC Class
    props['Ifc Class'] = element.is_a()

    # Properties and quantities directly attached to element (IsDefinedBy)
    if hasattr(element, 'IsDefinedBy'):
        for rel in element.IsDefinedBy or []:
            prop_def = getattr(rel, 'RelatingPropertyDefinition', None)
            if not prop_def:
                continue

            if prop_def.is_a('IfcPropertySet'):
                pset_name = getattr(prop_def, 'Name', 'Unknown')
                for prop in prop_def.HasProperties or []:
                    name = getattr(prop, 'Name', None)
                    if name:
                        key = f'{pset_name}:{name}'
                        if key not in props:
                            props[key] = get_property_value(prop)

            elif prop_def.is_a('IfcElementQuantity'):
                qty_name = getattr(prop_def, 'Name', 'Unknown')
                for qty in prop_def.Quantities or []:
                    name = getattr(qty, 'Name', None)
                    if name:
                        key = f'{qty_name}:{name}'
                        if key not in props:
                            props[key] = get_quantity_value(qty)

    # Properties from element type (IsTypedBy)
    if include_type_props and hasattr(element, 'IsTypedBy'):
        for rel in element.IsTypedBy or []:
            type_obj = getattr(rel, 'RelatingType', None)
            if not type_obj:
                continue

            # Type attributes
            type_attrs = ['GlobalId', 'Name', 'Description', 'ElementType']
            for attr in type_attrs:
                val = getattr(type_obj, attr, None)
                if val is not None and f'Type:{attr}' not in props:
                    props[f'Type:{attr}'] = val

            # PropertySet and Quantity from type
            if hasattr(type_obj, 'HasPropertySets'):
                for ps in type_obj.HasPropertySets or []:
                    if ps.is_a('IfcPropertySet'):
                        pset_name = getattr(ps, 'Name', 'Unknown')
                        for prop in ps.HasProperties or []:
                            name = getattr(prop, 'Name', None)
                            if name:
                                key = f'{pset_name}:{name}'
                                if key not in props:
                                    props[key] = get_property_value(prop)
                    elif ps.is_a('IfcElementQuantity'):
                        qty_name = getattr(ps, 'Name', 'Unknown')
                        for qty in ps.Quantities or []:
                            name = getattr(qty, 'Name', None)
                            if name:
                                key = f'{qty_name}:{name}'
                                if key not in props:
                                    props[key] = get_quantity_value(qty)

    return props


# ----------------------------------------------------------------------
# Model summary data collection
# ----------------------------------------------------------------------

def get_model_summary(ifc_file):
    """Extract project info, stories, building height, address WITHOUT unit conversion."""
    summary = {
        'project_info': {},
        'stories': [],
        'stories_count': 0,
        'building_height': None,
        'address': 'N/A',
        'buildings': []
    }

    # Project
    projects = ifc_file.by_type('IfcProject')
    if projects:
        p = projects[0]
        summary['project_info'] = {
            'Name': p.Name or 'N/A',
            'Description': p.Description or 'N/A',
            'LongName': p.LongName or 'N/A',
            'Phase': p.Phase or 'N/A',
            'Schema': ifc_file.schema
        }
        if p.OwnerHistory:
            oh = p.OwnerHistory
            ts = getattr(oh, 'CreationDate', None) or getattr(oh, 'CreationTime', None)
            if ts:
                summary['project_info']['CreationDate'] = datetime.fromtimestamp(ts).strftime('%Y-%m-%d %H:%M:%S')

    # Stories
    stories = ifc_file.by_type('IfcBuildingStorey')
    summary['stories_count'] = len(stories)
    summary['stories'] = [
        {'Name': s.Name or 'N/A', 'Elevation': s.Elevation if s.Elevation is not None else 0}
        for s in stories
    ]

    # Building height (difference between highest and lowest story) - NO unit conversion
    if stories:
        elevations = [s.Elevation for s in stories if s.Elevation is not None]
        if elevations:
            summary['building_height'] = max(elevations) - min(elevations)

    # Buildings and address
    buildings = ifc_file.by_type('IfcBuilding')
    for b in buildings:
        summary['buildings'].append({
            'Name': b.Name or 'N/A',
            'LongName': b.LongName or 'N/A'
        })

    sites = ifc_file.by_type('IfcSite')
    if sites and sites[0].LongName:
        summary['address'] = sites[0].LongName
    elif buildings and buildings[0].LongName:
        summary['address'] = buildings[0].LongName

    return summary


# ----------------------------------------------------------------------
# Main parsing and Excel generation function
# ----------------------------------------------------------------------

def parse_ifc_to_excel(ifc_path, output_excel_path):
    """Main entry point: reads IFC and saves Excel report with two sheets: Сводка and Элементы.
    
    The 'Элементы' sheet has the following format:
    - Columns: "Element Specific:Long Name", "Element Specific:Name", "Ifc Class", 
               "Storey", and "PSetName:PropertyName" for all properties.
    - Rows: One row per element (IfcElement, IfcBuilding, IfcBuildingStorey)
    - Values are placed in the corresponding column based on the property key.
    """
    print(f"Opening IFC: {ifc_path}")
    ifc_file = ifcopenshell.open(ifc_path)

    # 1. Summary data
    summary = get_model_summary(ifc_file)

    # 2. Collect all elements: IfcElement + IfcBuilding + IfcBuildingStorey
    elements = ifc_file.by_type('IfcElement')
    buildings = ifc_file.by_type('IfcBuilding')
    storeys = ifc_file.by_type('IfcBuildingStorey')
    
    print(f"Found IfcElement: {len(elements)}, IfcBuilding: {len(buildings)}, IfcBuildingStorey: {len(storeys)}")

    all_elements_data = []
    all_keys = set()

    # Process IfcBuilding elements first (they appear at top level in user's example)
    for i, elem in enumerate(buildings):
        elem_props = get_element_properties(elem, ifc_file, include_type_props=True)
        all_elements_data.append(elem_props)
        all_keys.update(elem_props.keys())

    # Process IfcBuildingStorey elements second
    for i, elem in enumerate(storeys):
        elem_props = get_element_properties(elem, ifc_file, include_type_props=True)
        all_elements_data.append(elem_props)
        all_keys.update(elem_props.keys())

    # Process all other IfcElement elements
    for i, elem in enumerate(elements):
        if i % 500 == 0:
            print(f"Processing IfcElement {i} / {len(elements)}")
        elem_props = get_element_properties(elem, ifc_file, include_type_props=True)
        all_elements_data.append(elem_props)
        all_keys.update(elem_props.keys())

    all_keys = sorted(all_keys)
    print(f"Unique parameters: {len(all_keys)}")

    # 3. Create Excel workbook
    wb = openpyxl.Workbook()
    ws_summary = wb.active
    ws_summary.title = "Сводка"

    # Styles
    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="2c3e50", end_color="2c3e50", fill_type="solid")
    header_alignment = Alignment(horizontal="center", vertical="center")
    thin_border = Border(
        left=Side(style='thin'), right=Side(style='thin'),
        top=Side(style='thin'), bottom=Side(style='thin')
    )

    # ---- Sheet "Сводка" ----
    ws_summary.cell(row=1, column=1, value="Параметр").font = header_font
    ws_summary.cell(row=1, column=1).fill = header_fill
    ws_summary.cell(row=1, column=2, value="Значение").font = header_font
    ws_summary.cell(row=1, column=2).fill = header_fill

    row = 2
    for k, v in summary['project_info'].items():
        ws_summary.cell(row=row, column=1, value=k).border = thin_border
        ws_summary.cell(row=row, column=2, value=v).border = thin_border
        row += 1

    ws_summary.cell(row=row, column=1, value="Количество этажей").border = thin_border
    ws_summary.cell(row=row, column=2, value=summary['stories_count']).border = thin_border
    row += 1

    ws_summary.cell(row=row, column=1, value="Высота здания").border = thin_border
    ws_summary.cell(row=row, column=2, value=summary['building_height'] or 'N/A').border = thin_border
    row += 1

    ws_summary.cell(row=row, column=1, value="Адрес").border = thin_border
    ws_summary.cell(row=row, column=2, value=summary['address']).border = thin_border
    row += 1

    ws_summary.cell(row=row, column=1, value="Всего элементов (IfcElement)").border = thin_border
    ws_summary.cell(row=row, column=2, value=len(elements)).border = thin_border
    row += 1

    if summary['stories']:
        ws_summary.cell(row=row, column=1, value="Список этажей (отметка)").border = thin_border
        row += 1
        for s in summary['stories']:
            ws_summary.cell(row=row, column=1, value=f"  {s['Name']}").border = thin_border
            ws_summary.cell(row=row, column=2, value=s['Elevation']).border = thin_border
            row += 1

    ws_summary.column_dimensions['A'].width = 35
    ws_summary.column_dimensions['B'].width = 50

    # ---- Sheet "Элементы" ----
    ws_elements = wb.create_sheet("Элементы")
    for col_idx, key in enumerate(all_keys, 1):
        cell = ws_elements.cell(row=1, column=col_idx, value=key)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = header_alignment
        cell.border = thin_border

    for row_idx, elem_props in enumerate(all_elements_data, 2):
        for col_idx, key in enumerate(all_keys, 1):
            val = elem_props.get(key)
            if val is not None:
                ws_elements.cell(row=row_idx, column=col_idx, value=val).border = thin_border

    # Auto-width columns (max 50 chars)
    for col in ws_elements.columns:
        max_len = 0
        col_letter = col[0].column_letter
        for cell in col:
            if cell.value:
                try:
                    max_len = max(max_len, len(str(cell.value)))
                except:
                    pass
        ws_elements.column_dimensions[col_letter].width = min(max_len + 2, 50)

    wb.save(output_excel_path)
    print(f"Excel saved: {output_excel_path}")


# ----------------------------------------------------------------------
# Entry point for ifc_inspect service
# ----------------------------------------------------------------------

def parse_ifc_file(ifc_path, output_folder):
    """
    Main function for ifc_inspect service.
    Parses IFC file and generates Excel report with two sheets: Сводка and Элементы.
    NO unit conversion, NO opening subtraction.
    
    Args:
        ifc_path: Path to IFC file
        output_folder: Folder to save Excel report
        
    Returns:
        Dict with parsing results and path to Excel file
    """
    print(f"Parsing IFC file: {ifc_path}")
    ifc_file = ifcopenshell.open(ifc_path)
    
    # Get summary data
    summary = get_model_summary(ifc_file)
    
    # Collect all elements
    elements = ifc_file.by_type('IfcElement')
    
    # Generate Excel report
    excel_filename = "ifc_report.xlsx"
    excel_path = os.path.join(output_folder, excel_filename)
    parse_ifc_to_excel(ifc_path, excel_path)
    
    # Prepare summary
    elements_summary = {}
    for elem_type in set(e.is_a() for e in elements):
        count = len([e for e in elements if e.is_a() == elem_type])
        elements_summary[elem_type] = count
    
    result = {
        'success': True,
        'excel_path': excel_path,
        'excel_filename': excel_filename,
        'project_info': summary['project_info'],
        'elements_count': elements_summary,
        'total_elements': len(elements),
        'materials_count': 0,  # Not tracked in this version
        'stories_count': summary['stories_count']
    }
    
    print(f"✅ IFC parsing complete. Total elements: {result['total_elements']}")
    return result


# ----------------------------------------------------------------------
# Script entry point when run directly
# ----------------------------------------------------------------------
if __name__ == "__main__":
    import sys
    if len(sys.argv) < 2:
        print("Usage: python ifc_parser.py <path_to_IFC> [path_to_output_XLSX]")
        sys.exit(1)
    ifc_file = sys.argv[1]
    out_file = sys.argv[2] if len(sys.argv) > 2 else "ifc_parser_output.xlsx"
    parse_ifc_to_excel(ifc_file, out_file)
