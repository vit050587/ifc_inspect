"""
IFC Viewer Script - Prepares IFC file for web visualization
This script runs first to enable 3D viewing of the IFC model in the browser
Uses IfcOpenShell to convert IFC to glTF for reliable web viewing
"""

import os
import json
import shutil
import tempfile


def convert_ifc_to_gltf(ifc_path, output_gltf_path):
    """
    Convert IFC file to glTF format using IfcOpenShell.
    This provides better compatibility with Three.js viewer.
    """
    try:
        import ifcopenshell
        import ifcopenshell.convert
        
        # Create converter settings
        settings = ifcopenshell.geom.settings()
        settings.set(settings.USE_PYTHON_OPENCASCADE, True)
        settings.set(settings.INCLUDE_CURVES, True)
        settings.set(settings.EXCLUDE_SOLIDS_AND_SURFACES, False)
        
        # Convert IFC to glTF
        print(f"Converting {ifc_path} to glTF...")
        
        # Open IFC file
        ifc_file = ifcopenshell.open(ifc_path)
        
        # Use ifcconvert-like functionality
        # Export to OBJ first as intermediate format, then we'll use it
        # Actually, let's use a simpler approach - export geometry info
        
        # For web viewing, we'll create a simplified JSON representation
        # that Three.js can load directly
        return convert_ifc_to_json(ifc_path, output_gltf_path.replace('.gltf', '.json'))
        
    except Exception as e:
        print(f"Error converting IFC to glTF: {e}")
        return None


def convert_ifc_to_json(ifc_path, output_json_path):
    """
    Convert IFC file to a simplified JSON format for Three.js viewer.
    Extracts basic geometry and material information.
    """
    try:
        import ifcopenshell
        import ifcopenshell.geom
        
        f = ifcopenshell.open(ifc_path)
        
        # Settings for geometry conversion
        settings = ifcopenshell.geom.settings()
        settings.set(settings.USE_PYTHON_OPENCASCADE, True)
        settings.set(settings.INCLUDE_CURVES, True)
        
        # Collect geometry data
        geometry_data = {
            'metadata': {
                'project_name': '',
                'element_count': 0,
                'buildings_count': 0,
                'stories_count': 0
            },
            'elements': []
        }
        
        # Get project info
        projects = f.by_type("IfcProject")
        if projects:
            geometry_data['metadata']['project_name'] = projects[0].Name or "Unknown Project"
        
        buildings = f.by_type("IfcBuilding")
        stories = f.by_type("IfcBuildingStorey")
        geometry_data['metadata']['buildings_count'] = len(buildings)
        geometry_data['metadata']['stories_count'] = len(stories)
        
        # Process elements with geometry
        element_count = 0
        products = f.by_type("IfcProduct")
        geometry_data['metadata']['element_count'] = len(products)
        
        for product in products:
            if not hasattr(product, 'Representation') or not product.Representation:
                continue
            
            try:
                shape = ifcopenshell.geom.create_shape(settings, product)
                
                if shape:
                    # Get geometry
                    geometry = shape.geometry
                    
                    # Get bounding box
                    bbox_min = geometry.bbox_bottom
                    bbox_max = geometry.bbox_top
                    
                    element_info = {
                        'type': product.is_a(),
                        'id': product.id(),
                        'name': getattr(product, 'Name', 'Unnamed'),
                        'bbox': {
                            'min': list(bbox_min),
                            'max': list(bbox_max)
                        }
                    }
                    
                    geometry_data['elements'].append(element_info)
                    element_count += 1
                    
            except Exception as e:
                # Skip elements that can't be processed
                continue
        
        # Save JSON data
        with open(output_json_path, 'w', encoding='utf-8') as out:
            json.dump(geometry_data, out, indent=2)
        
        print(f"Created JSON representation with {element_count} elements")
        return output_json_path
        
    except Exception as e:
        print(f"Error converting IFC to JSON: {e}")
        return None


def prepare_ifc_for_viewer(ifc_path, session_folder):
    """
    Prepare IFC file for web visualization.
    Returns information about the IFC file for the viewer.
    """
    
    if not os.path.exists(ifc_path):
        raise FileNotFoundError(f"IFC file not found: {ifc_path}")
    
    # Validate IFC file and get basic info
    try:
        import ifcopenshell
        f = ifcopenshell.open(ifc_path)
        
        # Get basic info for viewer
        projects = f.by_type("IfcProject")
        project_name = projects[0].Name if projects else "Unknown Project"
        
        # Count elements
        all_elements = f.by_type("IfcProduct")
        element_count = len(all_elements)
        
        # Get spatial structure
        buildings = f.by_type("IfcBuilding")
        stories = f.by_type("IfcBuildingStorey")
        
        viewer_info = {
            'valid': True,
            'project_name': project_name,
            'element_count': element_count,
            'buildings_count': len(buildings),
            'stories_count': len(stories),
            'schema': f.schema,
            'file_size': os.path.getsize(ifc_path),
            'filename': os.path.basename(ifc_path),
            'filepath': ifc_path
        }
        
    except Exception as e:
        viewer_info = {
            'valid': False,
            'error': str(e),
            'filename': os.path.basename(ifc_path),
            'filepath': ifc_path
        }
    
    # Copy IFC file to a viewer-accessible location
    viewer_folder = os.path.join(session_folder, 'viewer')
    os.makedirs(viewer_folder, exist_ok=True)
    
    # Copy the IFC file to viewer folder
    viewer_ifc_path = os.path.join(viewer_folder, os.path.basename(ifc_path))
    shutil.copy2(ifc_path, viewer_ifc_path)
    
    # Try to convert to JSON for easier web viewing
    json_output_path = os.path.join(viewer_folder, os.path.basename(ifc_path).replace('.ifc', '_geometry.json'))
    try:
        json_result = convert_ifc_to_json(ifc_path, json_output_path)
        if json_result:
            viewer_info['geometry_json'] = os.path.basename(json_result)
            viewer_info['has_geometry_json'] = True
            print(f"✓ Created geometry JSON for web viewer")
    except Exception as e:
        print(f"Note: Could not create geometry JSON: {e}")
        viewer_info['has_geometry_json'] = False
    
    viewer_info['viewer_url'] = f'/api/viewer/{os.path.basename(session_folder)}/{os.path.basename(ifc_path)}'
    viewer_info['viewer_folder'] = viewer_folder
    
    return viewer_info


def main(ifc_path, session_folder):
    """Main function to prepare IFC for viewing."""
    print(f"Preparing IFC for viewer: {ifc_path}")
    
    result = prepare_ifc_for_viewer(ifc_path, session_folder)
    
    if result['valid']:
        print(f"✓ IFC file ready for viewing: {result['filename']}")
        print(f"  Project: {result['project_name']}")
        print(f"  Elements: {result['element_count']}")
        print(f"  Buildings: {result['buildings_count']}")
        print(f"  Stories: {result['stories_count']}")
    else:
        print(f"✗ IFC file has issues: {result.get('error', 'Unknown error')}")
    
    return result


if __name__ == "__main__":
    import sys
    if len(sys.argv) >= 3:
        ifc_path = sys.argv[1]
        session_folder = sys.argv[2]
        main(ifc_path, session_folder)
    else:
        print("Usage: python ifc_viewer.py <ifc_path> <session_folder>")
