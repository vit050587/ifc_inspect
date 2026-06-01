import ifcopenshell
import sys

def extract_ifc_info(ifc_path):
    """Extract all information from IFC file in human-readable format."""
    
    try:
        model = ifcopenshell.open(ifc_path)
    except Exception as e:
        print(f"Error opening file: {e}")
        return
    
    output = []
    output.append("=" * 80)
    output.append("IFC FILE INFORMATION")
    output.append("=" * 80)
    output.append("")
    
    # File header information
    output.append("1. FILE HEADER INFORMATION")
    output.append("-" * 40)
    header = model.header
    output.append(f"File Name: {header.file_name.name if header.file_name else 'N/A'}")
    output.append(f"Time Stamp: {header.file_name.time_stamp if header.file_name else 'N/A'}")
    output.append(f"Author(s): {', '.join(header.file_name.author) if header.file_name and header.file_name.author else 'N/A'}")
    output.append(f"Organization: {', '.join(header.file_name.organization) if header.file_name and header.file_name.organization else 'N/A'}")
    output.append(f"Preprocessor Version: {header.file_name.preprocessor_version if header.file_name else 'N/A'}")
    output.append(f"Originating System: {header.file_name.originating_system if header.file_name else 'N/A'}")
    output.append(f"Authorization: {header.file_name.authorization if header.file_name else 'N/A'}")
    schema = header.schema if hasattr(header, 'schema') else getattr(model, 'schema', 'N/A')
    output.append(f"Schema: {schema}")
    output.append("")
    
    # Project information
    output.append("2. PROJECT INFORMATION")
    output.append("-" * 40)
    projects = model.by_type("IfcProject")
    for project in projects:
        output.append(f"Project Name: {project.Name if project.Name else 'N/A'}")
        output.append(f"Project Description: {project.Description if project.Description else 'N/A'}")
        output.append(f"Project ObjectType: {project.ObjectType if project.ObjectType else 'N/A'}")
        output.append(f"Project LongName: {project.LongName if project.LongName else 'N/A'}")
        output.append(f"Project Phase: {project.Phase if project.Phase else 'N/A'}")
        if hasattr(project, 'RepresentationContexts') and project.RepresentationContexts:
            output.append(f"Representation Contexts: {len(project.RepresentationContexts)}")
    output.append("")
    
    # Building information
    output.append("3. BUILDING INFORMATION")
    output.append("-" * 40)
    buildings = model.by_type("IfcBuilding")
    for building in buildings:
        output.append(f"Building Name: {building.Name if building.Name else 'N/A'}")
        output.append(f"Building Description: {building.Description if building.Description else 'N/A'}")
        output.append(f"Building ObjectType: {building.ObjectType if building.ObjectType else 'N/A'}")
        if hasattr(building, 'ElevationOfRefHeight') and building.ElevationOfRefHeight:
            output.append(f"Elevation of Ref Height: {building.ElevationOfRefHeight}")
        if hasattr(building, 'ElevationOfTerrain') and building.ElevationOfTerrain:
            output.append(f"Elevation of Terrain: {building.ElevationOfTerrain}")
    output.append("")
    
    # Building Storey information
    output.append("4. BUILDING STOREY INFORMATION")
    output.append("-" * 40)
    storeys = model.by_type("IfcBuildingStorey")
    output.append(f"Total Storeys: {len(storeys)}")
    for storey in storeys:
        output.append(f"  - Storey Name: {storey.Name if storey.Name else 'N/A'}")
        output.append(f"    Description: {storey.Description if storey.Description else 'N/A'}")
        if hasattr(storey, 'Elevation') and storey.Elevation:
            output.append(f"    Elevation: {storey.Elevation}")
    output.append("")
    
    # Element counts by type
    output.append("5. ELEMENT COUNTS BY TYPE")
    output.append("-" * 40)
    element_types = {}
    for entity in model:
        entity_type = entity.is_a()
        if entity_type not in element_types:
            element_types[entity_type] = 0
        element_types[entity_type] += 1
    
    sorted_types = sorted(element_types.items(), key=lambda x: x[1], reverse=True)
    for entity_type, count in sorted_types:
        if count > 0:
            output.append(f"{entity_type}: {count}")
    output.append("")
    
    # Detailed element information (structural elements)
    output.append("6. STRUCTURAL AND BUILDING ELEMENTS")
    output.append("-" * 40)
    
    structural_types = [
        "IfcBeam", "IfcColumn", "IfcSlab", "IfcWall", "IfcFooting", 
        "IfcPile", "IfcMember", "IfcPlate", "IfcStair", "IfcRailing",
        "IfcRoof", "IfcDoor", "IfcWindow", "IfcSpace"
    ]
    
    for elem_type in structural_types:
        elements = model.by_type(elem_type)
        if elements:
            output.append(f"\n{elem_type.upper()} ({len(elements)} elements)")
            output.append("~" * 40)
            for i, elem in enumerate(elements[:20]):
                name = elem.Name if elem.Name else f"Unnamed_{i+1}"
                desc = elem.Description if elem.Description else ""
                tag = elem.Tag if hasattr(elem, 'Tag') and elem.Tag else ""
                
                elem_info = f"  [{i+1}] Name: {name}"
                if desc:
                    elem_info += f", Desc: {desc}"
                if tag:
                    elem_info += f", Tag: {tag}"
                output.append(elem_info)
            
            if len(elements) > 20:
                output.append(f"  ... and {len(elements) - 20} more {elem_type} elements")
    
    output.append("")
    
    # Material information
    output.append("7. MATERIAL INFORMATION")
    output.append("-" * 40)
    materials = model.by_type("IfcMaterial")
    output.append(f"Total Materials: {len(materials)}")
    for mat in materials[:30]:
        name = mat.Name if mat.Name else "Unnamed"
        output.append(f"  - {name}")
    if len(materials) > 30:
        output.append(f"  ... and {len(materials) - 30} more materials")
    output.append("")
    
    # Property sets
    output.append("8. PROPERTY SETS")
    output.append("-" * 40)
    property_sets = model.by_type("IfcPropertySet")
    output.append(f"Total Property Sets: {len(property_sets)}")
    for pset in property_sets[:20]:
        name = pset.Name if pset.Name else "Unnamed"
        desc = pset.Description if pset.Description else ""
        output.append(f"  - {name}: {desc}")
    if len(property_sets) > 20:
        output.append(f"  ... and {len(property_sets) - 20} more property sets")
    output.append("")
    
    # Quantities
    output.append("9. QUANTITY SETS")
    output.append("-" * 40)
    quantity_sets = model.by_type("IfcElementQuantity")
    output.append(f"Total Quantity Sets: {len(quantity_sets)}")
    output.append("")
    
    # Relationships
    output.append("10. RELATIONSHIPS SUMMARY")
    output.append("-" * 40)
    relationship_types = ["IfcRelContainedInSpatialStructure", "IfcRelAggregates", 
                          "IfcRelDefinesByType", "IfcRelDefinesByProperties",
                          "IfcRelAssociatesMaterial"]
    for rel_type in relationship_types:
        rels = model.by_type(rel_type)
        output.append(f"{rel_type}: {len(rels)}")
    output.append("")
    
    for line in output:
        print(line)

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python extract_ifc_info.py <ifc_file>")
        sys.exit(1)
    
    ifc_file = sys.argv[1]
    extract_ifc_info(ifc_file)
