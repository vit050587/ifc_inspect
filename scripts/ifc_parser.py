"""
IFC Parser - Extracts structural elements, materials, dimensions, volumes and other info
Generates: 1) Brief TXT report, 2) Detailed Excel report

Modified for ifc_inspect service - returns data as dict instead of writing files directly.
"""

import ifcopenshell
import ifcopenshell.util.element
import ifcopenshell.util.unit
from datetime import datetime
from collections import defaultdict
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
import os


def get_ifc_info(ifc_path):
    """Parse IFC file and extract all relevant information."""
    
    f = ifcopenshell.open(ifc_path)
    
    # Basic project info
    project_info = {}
    buildings = []
    stories = []
    elements_by_type = defaultdict(list)
    materials = set()
    material_elements = []
    property_sets_count = 0
    quantity_sets_count = 0
    material_relations = 0
    property_definitions = 0
    
    # Unit conversion - determine length unit scale
    length_unit_scale = 1.0  # Default to meters
    
    try:
        units = f.by_type("IfcUnitAssignment")
        if units:
            for u in units[0].Units or []:
                if u.is_a("IfcSIUnit") and u.UnitType == "LENGTHUNIT":
                    prefix = getattr(u, 'Prefix', None)
                    if prefix == "MILLI":
                        length_unit_scale = 0.001
                    elif prefix == "CENTI":
                        length_unit_scale = 0.01
                    elif prefix == "DECI":
                        length_unit_scale = 0.1
                    else:
                        length_unit_scale = 1.0
                    break
    except Exception as e:
        print(f"Warning: Could not determine units: {e}")
    
    # Get project
    projects = f.by_type("IfcProject")
    if projects:
        project = projects[0]
        project_info['name'] = project.Name or 'N/A'
        project_info['description'] = project.Description or 'N/A'
        project_info['phase'] = getattr(project, 'Phase', 'N/A') or 'N/A'
        
        # Get LongName for address/object info
        long_name = getattr(project, 'LongName', None)
        if long_name:
            project_info['long_name'] = long_name
        
        # Get creation time from owner history
        if hasattr(project, 'OwnerHistory') and project.OwnerHistory:
            oh = project.OwnerHistory
            # Try CreationDate attribute first (IFC4)
            timestamp = None
            if hasattr(oh, 'CreationDate') and oh.CreationDate:
                timestamp = oh.CreationDate
            elif hasattr(oh, 'CreationTime') and oh.CreationTime:
                timestamp = oh.CreationTime
            elif len(oh) > 7 and oh[7]:
                timestamp = oh[7]
                
            if timestamp:
                project_info['creation_date'] = datetime.fromtimestamp(timestamp).strftime('%d %B %Y г.')
            else:
                project_info['creation_date'] = 'N/A'
        else:
            project_info['creation_date'] = 'N/A'
            
        project_info['schema'] = f.schema
    
    # Get site and building
    sites = f.by_type("IfcSite")
    for site in sites:
        if hasattr(site, 'RefElevation') and site.RefElevation:
            project_info['site_elevation'] = site.RefElevation * length_unit_scale
        if hasattr(site, 'LongName') and site.LongName:
            project_info['site_long_name'] = site.LongName
            
    building_elements = f.by_type("IfcBuilding")
    for building in building_elements:
        building_info = {
            'name': building.Name or 'N/A',
            'long_name': getattr(building, 'LongName', 'N/A') or 'N/A',
            'elevation': getattr(building, 'ElevationOfRefHeight', None),
            'stories_count': 0
        }
        buildings.append(building_info)
        
    # Get building stories (levels)
    story_elements = f.by_type("IfcBuildingStorey")
    for story in story_elements:
        elevation_raw = getattr(story, 'Elevation', 0) or 0
        elevation_m = elevation_raw * length_unit_scale
        story_info = {
            'name': story.Name or 'N/A',
            'long_name': getattr(story, 'LongName', 'N/A') or 'N/A',
            'elevation': round(elevation_m, 2),
            'composition_type': getattr(story, 'CompositionType', 'ELEMENT')
        }
        stories.append(story_info)
    
    # Structural element types to extract
    structural_types = [
        "IfcWall", "IfcSlab", "IfcColumn", "IfcBeam", "IfcStair", 
        "IfcRailing", "IfcRoof", "IfcPlate", "IfcMember", "IfcFooting",
        "IfcPile", "IfcCurtainWall", "IfcOpeningElement", "IfcWindow",
        "IfcDoor", "IfcBuildingElementProxy"
    ]
    
    # Also get element types for type-level properties
    element_type_map = {
        "IfcWall": "IfcWallType",
        "IfcSlab": "IfcSlabType",
        "IfcColumn": "IfcColumnType",
        "IfcBeam": "IfcBeamType",
        "IfcStair": "IfcStairType",
        "IfcRailing": "IfcRailingType",
        "IfcRoof": "IfcRoofType",
        "IfcPlate": "IfcPlateType",
        "IfcMember": "IfcMemberType",
        "IfcFooting": "IfcFootingType",
        "IfcPile": "IfcPileType",
        "IfcCurtainWall": "IfcCurtainWallType",
        "IfcWindow": "IfcWindowType",
        "IfcDoor": "IfcDoorType",
        "IfcBuildingElementProxy": "IfcBuildingElementProxyType"
    }
    
    # Cache for type-level properties
    type_properties_cache = {}
    
    # First pass: collect all type-level properties
    for elem_type, type_name in element_type_map.items():
        types = f.by_type(type_name)
        for elem_type_obj in types:
            type_id = elem_type_obj.id()
            type_props = {}
            
            # Get property sets from type
            if hasattr(elem_type_obj, 'HasPropertySets'):
                for prop_set in elem_type_obj.HasPropertySets or []:
                    if prop_set.is_a("IfcPropertySet"):
                        if hasattr(prop_set, 'HasProperties'):
                            for prop in prop_set.HasProperties or []:
                                if prop.is_a("IfcPropertySingleValue"):
                                    prop_name = getattr(prop, 'Name', '')
                                    if prop_name:
                                        prop_value = 'N/A'
                                        if hasattr(prop, 'NominalValue') and prop.NominalValue:
                                            if hasattr(prop.NominalValue, 'wrappedValue'):
                                                prop_value = prop.NominalValue.wrappedValue
                                            else:
                                                prop_value = str(prop.NominalValue)
                                        type_props[prop_name] = prop_value
            
            # Get quantities from type
            if hasattr(elem_type_obj, 'HasPropertySets'):
                for prop_set in elem_type_obj.HasPropertySets or []:
                    if prop_set.is_a("IfcElementQuantity"):
                        if hasattr(prop_set, 'Quantities'):
                            for qty in prop_set.Quantities or []:
                                if qty.is_a("IfcPhysicalSimpleQuantity"):
                                    qty_name = getattr(qty, 'Name', '')
                                    qty_value = None
                                    if hasattr(qty, 'NominalValue') and qty.NominalValue:
                                        qty_value = qty.NominalValue.wrappedValue if hasattr(qty.NominalValue, 'wrappedValue') else None
                                    elif hasattr(qty, 'LengthValue') and qty.LengthValue:
                                        qty_value = qty.LengthValue.wrappedValue if hasattr(qty.LengthValue, 'wrappedValue') else qty.LengthValue
                                    elif hasattr(qty, 'AreaValue') and qty.AreaValue:
                                        qty_value = qty.AreaValue.wrappedValue if hasattr(qty.AreaValue, 'wrappedValue') else qty.AreaValue
                                    elif hasattr(qty, 'VolumeValue') and qty.VolumeValue:
                                        qty_value = qty.VolumeValue.wrappedValue if hasattr(qty.VolumeValue, 'wrappedValue') else qty.VolumeValue
                                    
                                    if qty_name and qty_value is not None:
                                        type_props[qty_name] = qty_value
            
            if type_props:
                type_properties_cache[type_id] = type_props
    
    # Unit conversion setup
    unit_context = None
    length_unit = 1.0  # meters by default
    
    try:
        contexts = f.by_type("IfcGeometricRepresentationContext")
        for ctx in contexts:
            if ctx.ContextType == 'Model':
                unit_context = ctx
                break
        
        if not unit_context and contexts:
            unit_context = contexts[0]
            
        # Get length units
        units = f.by_type("IfcUnitAssignment")
        if units:
            for u in units[0].Units:
                if u.is_a("IfcSIUnit") and u.UnitType == "LENGTHUNIT":
                    prefix = getattr(u, 'Prefix', None)
                    if prefix == "MILLI":
                        length_unit = 0.001
                    elif prefix == "CENTI":
                        length_unit = 0.01
                    elif prefix == "DECI":
                        length_unit = 0.1
                    break
                elif u.is_a("IfcDerivedUnit"):
                    pass  # Handle derived units if needed
    except Exception as e:
        print(f"Warning: Could not determine units: {e}")
    
    # Extract all structural elements
    for elem_type in structural_types:
        elements = f.by_type(elem_type)
        for elem in elements:
            elem_data = {
                'type': elem_type,
                'id': elem.id(),
                'name': getattr(elem, 'Name', 'N/A') or 'N/A',
                'tag': getattr(elem, 'Tag', 'N/A') or 'N/A',
                'description': getattr(elem, 'Description', 'N/A') or 'N/A',
                'dimensions': {},
                'volume': None,
                'area': None,
                'material': 'N/A',
                'layer': 'N/A',
                'properties': {},
                'opening_volumes': 0.0,  # Суммарный объем проемов в элементе
                'opening_areas': 0.0     # Суммарная площадь проемов в элементе
            }
            
            # Get placement/elevation
            if hasattr(elem, 'ObjectPlacement') and elem.ObjectPlacement:
                placement = elem.ObjectPlacement
                if hasattr(placement, 'RelativePlacement'):
                    rel_placement = placement.RelativePlacement
                    if hasattr(rel_placement, 'Location'):
                        loc = rel_placement.Location
                        if len(loc.Coordinates) >= 3:
                            elem_data['z_coordinate'] = loc.Coordinates[2] * length_unit
            
            # Get quantities from IsDefinedBy relationships
            if hasattr(elem, 'IsDefinedBy'):
                for rel in elem.IsDefinedBy or []:
                    if hasattr(rel, 'RelatingPropertyDefinition'):
                        prop_def = rel.RelatingPropertyDefinition
                        if prop_def.is_a("IfcElementQuantity"):
                            quantity_sets_count += 1
                            if hasattr(prop_def, 'Quantities'):
                                for qty in prop_def.Quantities or []:
                                    if qty.is_a("IfcPhysicalSimpleQuantity"):
                                        qty_name = getattr(qty, 'Name', '')
                                        
                                        # Get value - different quantity types have different value attributes
                                        qty_value = None
                                        if hasattr(qty, 'NominalValue') and qty.NominalValue:
                                            qty_value = qty.NominalValue.wrappedValue if hasattr(qty.NominalValue, 'wrappedValue') else None
                                        elif hasattr(qty, 'LengthValue') and qty.LengthValue:
                                            qty_value = qty.LengthValue.wrappedValue if hasattr(qty.LengthValue, 'wrappedValue') else qty.LengthValue
                                        elif hasattr(qty, 'AreaValue') and qty.AreaValue:
                                            qty_value = qty.AreaValue.wrappedValue if hasattr(qty.AreaValue, 'wrappedValue') else qty.AreaValue
                                        elif hasattr(qty, 'VolumeValue') and qty.VolumeValue:
                                            qty_value = qty.VolumeValue.wrappedValue if hasattr(qty.VolumeValue, 'wrappedValue') else qty.VolumeValue
                                        
                                        if qty_name == 'NetVolume':
                                            elem_data['volume'] = qty_value
                                        elif qty_name == 'NetArea':
                                            elem_data['area'] = qty_value
                                        elif qty_name in ['Length', 'Width', 'Height', 'Depth'] and qty_value is not None:
                                            elem_data['dimensions'][qty_name] = qty_value * length_unit
                        elif prop_def.is_a("IfcPropertySet"):
                            property_sets_count += 1
                            if hasattr(prop_def, 'HasProperties'):
                                for prop in prop_def.HasProperties or []:
                                    if prop.is_a("IfcPropertySingleValue"):
                                        prop_name = getattr(prop, 'Name', 'N/A')
                                        prop_value = 'N/A'
                                        if hasattr(prop, 'NominalValue') and prop.NominalValue:
                                            if hasattr(prop.NominalValue, 'wrappedValue'):
                                                prop_value = prop.NominalValue.wrappedValue
                                            else:
                                                prop_value = str(prop.NominalValue)
                                        elem_data['properties'][prop_name] = prop_value
            
            # Get type-level properties via IsTypedBy relationship
            if hasattr(elem, 'IsTypedBy'):
                for rel in elem.IsTypedBy or []:
                    if hasattr(rel, 'RelatingType'):
                        type_obj = rel.RelatingType
                        type_id = type_obj.id()
                        
                        # Merge cached type properties
                        if type_id in type_properties_cache:
                            for prop_name, prop_value in type_properties_cache[type_id].items():
                                # Instance properties take precedence over type properties
                                if prop_name not in elem_data['properties']:
                                    elem_data['properties'][prop_name] = prop_value
                        
                        # Also get material from type if not found yet
                        if elem_data['material'] == 'N/A':
                            type_material = ifcopenshell.util.element.get_material(type_obj)
                            if type_material and type_material.is_a("IfcMaterial"):
                                elem_data['material'] = getattr(type_material, 'Name', 'N/A') or 'N/A'
                                materials.add(elem_data['material'])
            
            # Get material
            mat_assoc = ifcopenshell.util.element.get_material(elem)
            if mat_assoc:
                if mat_assoc.is_a("IfcMaterial"):
                    elem_data['material'] = getattr(mat_assoc, 'Name', 'N/A') or 'N/A'
                    materials.add(elem_data['material'])
                elif mat_assoc.is_a("IfcMaterialLayerSetUsage"):
                    if hasattr(mat_assoc, 'ForLayerSet') and mat_assoc.ForLayerSet:
                        layer_set = mat_assoc.ForLayerSet
                        if hasattr(layer_set, 'MaterialLayers'):
                            layer_names = []
                            total_thickness = 0
                            for layer in layer_set.MaterialLayers or []:
                                if hasattr(layer, 'Material') and layer.Material:
                                    mat_name = getattr(layer.Material, 'Name', 'Unknown') or 'Unknown'
                                    layer_names.append(mat_name)
                                    materials.add(mat_name)
                                if hasattr(layer, 'LayerThickness') and layer.LayerThickness:
                                    total_thickness += layer.LayerThickness * length_unit
                            elem_data['material'] = ', '.join(layer_names) if layer_names else 'N/A'
                            elem_data['layer'] = f"{total_thickness:.3f}м" if total_thickness > 0 else 'N/A'
                            material_relations += 1
                elif mat_assoc.is_a("IfcMaterialProfileSetUsage"):
                    if hasattr(mat_assoc, 'ForProfileSet') and mat_assoc.ForProfileSet:
                        profile_set = mat_assoc.ForProfileSet
                        if hasattr(profile_set, 'MaterialProfiles'):
                            mat_names = []
                            for profile in profile_set.MaterialProfiles or []:
                                if hasattr(profile, 'Material') and profile.Material:
                                    mat_name = getattr(profile.Material, 'Name', 'Unknown') or 'Unknown'
                                    mat_names.append(mat_name)
                                    materials.add(mat_name)
                            elem_data['material'] = ', '.join(mat_names) if mat_names else 'N/A'
                            material_relations += 1
                elif mat_assoc.is_a("IfcMaterialList"):
                    if hasattr(mat_assoc, 'Materials'):
                        mat_names = []
                        for mat in mat_assoc.Materials or []:
                            mat_name = getattr(mat, 'Name', 'Unknown') or 'Unknown'
                            mat_names.append(mat_name)
                            materials.add(mat_name)
                        elem_data['material'] = ', '.join(mat_names)
                        material_relations += 1
            
            # If material not found via association, try to get from properties (Reference property)
            if elem_data['material'] == 'N/A' and elem_data['properties']:
                ref_prop = elem_data['properties'].get('Reference')
                if ref_prop and ref_prop != 'N/A':
                    elem_data['material'] = ref_prop
                    materials.add(ref_prop)
            
            # Count property definitions
            if hasattr(elem, 'IsTypedBy'):
                for rel in elem.IsTypedBy or []:
                    if hasattr(rel, 'RelatingType'):
                        prop_defs = getattr(rel.RelatingType, 'HasPropertySets', None)
                        if prop_defs:
                            property_definitions += len(prop_defs)
            
            elements_by_type[elem_type].append(elem_data)
    
    # Process openings and associate them with parent elements (walls, slabs, etc.)
    # Create a mapping of element IDs to their data for quick lookup
    element_id_map = {}
    for elem_type, elems in elements_by_type.items():
        for elem_data in elems:
            element_id_map[elem_data['id']] = elem_data
    
    # Process all openings
    openings = f.by_type("IfcOpeningElement")
    for opening in openings:
        opening_data = {
            'id': opening.id(),
            'name': getattr(opening, 'Name', 'N/A') or 'N/A',
            'tag': getattr(opening, 'Tag', 'N/A') or 'N/A',
            'description': getattr(opening, 'Description', 'N/A') or 'N/A',
            'type': 'IfcOpeningElement',
            'volume': None,
            'area': None,
            'dimensions': {},
            'parent_element_id': None,
            'material': 'N/A',
            'layer': 'N/A',
            'properties': {},
            'opening_volumes': 0.0,
            'opening_areas': 0.0
        }
        
        # Get quantities from opening
        if hasattr(opening, 'IsDefinedBy'):
            for rel in opening.IsDefinedBy or []:
                if hasattr(rel, 'RelatingPropertyDefinition'):
                    prop_def = rel.RelatingPropertyDefinition
                    if prop_def.is_a("IfcElementQuantity"):
                        if hasattr(prop_def, 'Quantities'):
                            for qty in prop_def.Quantities or []:
                                if qty.is_a("IfcPhysicalSimpleQuantity"):
                                    qty_name = getattr(qty, 'Name', '')
                                    qty_value = None
                                    if hasattr(qty, 'NominalValue') and qty.NominalValue:
                                        qty_value = qty.NominalValue.wrappedValue if hasattr(qty.NominalValue, 'wrappedValue') else None
                                    elif hasattr(qty, 'LengthValue') and qty.LengthValue:
                                        qty_value = qty.LengthValue.wrappedValue if hasattr(qty.LengthValue, 'wrappedValue') else qty.LengthValue
                                    elif hasattr(qty, 'AreaValue') and qty.AreaValue:
                                        qty_value = qty.AreaValue.wrappedValue if hasattr(qty.AreaValue, 'wrappedValue') else qty.AreaValue
                                    elif hasattr(qty, 'VolumeValue') and qty.VolumeValue:
                                        qty_value = qty.VolumeValue.wrappedValue if hasattr(qty.VolumeValue, 'wrappedValue') else qty.VolumeValue
                                    
                                    if qty_name in ['Length', 'Width', 'Height', 'Depth'] and qty_value is not None:
                                        opening_data['dimensions'][qty_name] = qty_value * length_unit
                                    elif qty_name == 'NetVolume':
                                        opening_data['volume'] = qty_value
                                    elif qty_name == 'NetArea':
                                        opening_data['area'] = qty_value
                    elif prop_def.is_a("IfcPropertySet"):
                        if hasattr(prop_def, 'HasProperties'):
                            for prop in prop_def.HasProperties or []:
                                if prop.is_a("IfcPropertySingleValue"):
                                    prop_name = getattr(prop, 'Name', 'N/A')
                                    prop_value = 'N/A'
                                    if hasattr(prop, 'NominalValue') and prop.NominalValue:
                                        if hasattr(prop.NominalValue, 'wrappedValue'):
                                            prop_value = prop.NominalValue.wrappedValue
                                        else:
                                            prop_value = str(prop.NominalValue)
                                    opening_data['properties'][prop_name] = prop_value
        
        # Calculate volume and area from dimensions if not provided directly
        # Dimensions in IFC are typically in model units (mm), already converted above
        if opening_data['volume'] is None:
            width = opening_data['dimensions'].get('Width', 0)
            height = opening_data['dimensions'].get('Height', 0)
            depth = opening_data['dimensions'].get('Depth', 0)
            
            # If we have all three dimensions, calculate volume
            if width > 0 and height > 0 and depth > 0:
                opening_data['volume'] = width * height * depth
            # If we have Width and Height but no Depth, calculate area
            elif width > 0 and height > 0:
                opening_data['area'] = width * height
        
        # Find the parent element (wall/slab/etc.) that this opening belongs to
        if hasattr(opening, 'VoidsElements') and opening.VoidsElements:
            for rel in opening.VoidsElements:
                parent_elem = rel.RelatingBuildingElement
                parent_id = parent_elem.id()
                
                if parent_id in element_id_map:
                    opening_data['parent_element_id'] = parent_id
                    parent_data = element_id_map[parent_id]
                    
                    # Add opening volume/area to parent element's opening totals
                    if opening_data['volume'] is not None and opening_data['volume'] > 0:
                        parent_data['opening_volumes'] += opening_data['volume']
                    if opening_data['area'] is not None and opening_data['area'] > 0:
                        parent_data['opening_areas'] += opening_data['area']
        
        # Store opening data in elements_by_type for IfcOpeningElement
        if 'IfcOpeningElement' not in elements_by_type:
            elements_by_type['IfcOpeningElement'] = []
        elements_by_type['IfcOpeningElement'].append(opening_data)
    
    # Count unique property sets and definitions more accurately
    all_property_sets = f.by_type("IfcPropertySet")
    property_sets_count = len(all_property_sets)
    
    all_element_quantities = f.by_type("IfcElementQuantity")
    quantity_sets_count = len(all_element_quantities)
    
    all_materials = f.by_type("IfcMaterial")
    material_relations = len(f.by_type("IfcRelAssociatesMaterial"))
    
    all_property_defs = f.by_type("IfcPropertyDefinition")
    property_definitions = len(all_property_defs)
    
    return {
        'project_info': project_info,
        'buildings': buildings,
        'stories': stories,
        'elements_by_type': dict(elements_by_type),
        'materials': sorted(materials),
        'property_sets_count': property_sets_count,
        'quantity_sets_count': quantity_sets_count,
        'material_relations': material_relations,
        'property_definitions': property_definitions
    }


def generate_excel_report(data, output_path):
    """Generate detailed Excel report with all element information."""
    
    wb = openpyxl.Workbook()
    ws_summary = wb.active
    ws_summary.title = "Сводка"
    
    # Styles
    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="2c3e50", end_color="2c3e50", fill_type="solid")
    header_alignment = Alignment(horizontal="center", vertical="center")
    border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )
    
    # Summary sheet headers
    summary_headers = ["Параметр", "Значение"]
    for col, header in enumerate(summary_headers, 1):
        cell = ws_summary.cell(row=1, column=col, value=header)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = header_alignment
        cell.border = border
    
    # Project info
    pi = data['project_info']
    summary_data = [
        ["Название проекта", pi.get('name', 'N/A')],
        ["Описание", pi.get('description', 'N/A')],
        ["Дата создания", pi.get('creation_date', 'N/A')],
        ["Схема IFC", pi.get('schema', 'N/A')],
        ["Адрес", pi.get('long_name', pi.get('site_long_name', 'N/A'))],
        ["Количество зданий", len(data['buildings'])],
        ["Количество этажей", len(data['stories'])],
        ["Количество материалов", len(data['materials'])],
        ["Наборы свойств", data['property_sets_count']],
        ["Наборы количеств", data['quantity_sets_count']],
    ]
    
    for row_idx, (param, value) in enumerate(summary_data, 2):
        ws_summary.cell(row=row_idx, column=1, value=param).border = border
        ws_summary.cell(row=row_idx, column=2, value=value).border = border
    
    ws_summary.column_dimensions['A'].width = 35
    ws_summary.column_dimensions['B'].width = 25
    
    # Collect all unique property keys across all elements for dynamic columns
    all_property_keys = set()
    for elem_type, elems in data['elements_by_type'].items():
        for elem in elems:
            all_property_keys.update(elem['properties'].keys())
    
    # Sort property keys for consistent column order
    sorted_property_keys = sorted(all_property_keys)
    
    # Create sheets for each element type
    type_names_rus = {
        "IfcWall": "Стены",
        "IfcSlab": "Перекрытия",
        "IfcColumn": "Колонны",
        "IfcBeam": "Балки",
        "IfcStair": "Лестницы",
        "IfcRailing": "Ограждения",
        "IfcRoof": "Крыши",
        "IfcOpeningElement": "Проемы",
        "IfcWindow": "Окна",
        "IfcDoor": "Двери",
        "IfcBuildingElementProxy": "Прочие элементы"
    }
    
    for elem_type, elems in data['elements_by_type'].items():
        sheet_name = type_names_rus.get(elem_type, elem_type)
        # Excel sheet names have 31 char limit
        sheet_name = sheet_name[:31]
        ws = wb.create_sheet(title=sheet_name)
        
        # Headers - base columns + dynamic property columns
        headers = [
            "ID", "Имя", "Тег", "Описание", "Материал", "Слой/Толщина",
            "Объем (м³)", "Площадь (м²)", "Длина (м)", "Ширина (м)",
            "Высота (м)", "Z-координата (м)"
        ]
        # Add individual property columns
        headers.extend(sorted_property_keys)
        
        for col, header in enumerate(headers, 1):
            cell = ws.cell(row=1, column=col, value=header)
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = header_alignment
            cell.border = border
        
        # Data rows
        for row_idx, elem in enumerate(elems, 2):
            ws.cell(row=row_idx, column=1, value=elem['id'])
            ws.cell(row=row_idx, column=2, value=elem['name'])
            ws.cell(row=row_idx, column=3, value=elem['tag'])
            ws.cell(row=row_idx, column=4, value=elem['description'])
            ws.cell(row=row_idx, column=5, value=elem['material'])
            ws.cell(row=row_idx, column=6, value=elem['layer'])
            
            # For walls/slabs/columns etc., calculate net volume (subtracting openings)
            volume = elem['volume']
            if volume is not None and elem.get('opening_volumes', 0) > 0:
                # Net volume = gross volume - opening volumes
                volume = volume - elem['opening_volumes']
            
            ws.cell(row=row_idx, column=7, value=round(volume, 3) if volume else None)
            ws.cell(row=row_idx, column=8, value=round(elem['area'], 3) if elem['area'] else None)
            ws.cell(row=row_idx, column=9, value=round(elem['dimensions'].get('Length', 0), 3) if 'Length' in elem['dimensions'] else None)
            ws.cell(row=row_idx, column=10, value=round(elem['dimensions'].get('Width', 0), 3) if 'Width' in elem['dimensions'] else None)
            ws.cell(row=row_idx, column=11, value=round(elem['dimensions'].get('Height', 0), 3) if 'Height' in elem['dimensions'] else None)
            ws.cell(row=row_idx, column=12, value=round(elem.get('z_coordinate', 0), 3) if elem.get('z_coordinate') else None)
            
            # Fill individual property columns
            for col_idx, prop_key in enumerate(sorted_property_keys, 13):
                prop_value = elem['properties'].get(prop_key, None)
                ws.cell(row=row_idx, column=col_idx, value=prop_value)
            
            # Apply border to all cells in row
            for col in range(1, len(headers) + 1):
                ws.cell(row=row_idx, column=col).border = border
        
        # Auto-adjust column widths
        for col in ws.columns:
            max_length = 0
            column = col[0].column_letter
            for cell in col:
                try:
                    if cell.value:
                        max_length = max(max_length, len(str(cell.value)))
                except:
                    pass
            adjusted_width = min(max_length + 2, 50)
            ws.column_dimensions[column].width = adjusted_width
    
    # Save workbook
    wb.save(output_path)
    print(f"Detailed Excel report saved to: {output_path}")
    return output_path


def parse_ifc_file(ifc_path, output_folder):
    """
    Main function for ifc_inspect service.
    Parses IFC file and generates Excel report.
    
    Args:
        ifc_path: Path to IFC file
        output_folder: Folder to save Excel report
        
    Returns:
        Dict with parsing results and path to Excel file
    """
    print(f"Parsing IFC file: {ifc_path}")
    data = get_ifc_info(ifc_path)
    
    # Generate Excel report
    excel_filename = "ifc_report.xlsx"
    excel_path = os.path.join(output_folder, excel_filename)
    generate_excel_report(data, excel_path)
    
    # Prepare summary
    elements_summary = {}
    for elem_type, elems in data['elements_by_type'].items():
        elements_summary[elem_type] = len(elems)
    
    result = {
        'success': True,
        'excel_path': excel_path,
        'excel_filename': excel_filename,
        'project_info': data['project_info'],
        'elements_count': elements_summary,
        'total_elements': sum(len(elems) for elems in data['elements_by_type'].values()),
        'materials_count': len(data['materials']),
        'stories_count': len(data['stories'])
    }
    
    print(f"✅ IFC parsing complete. Total elements: {result['total_elements']}")
    return result
