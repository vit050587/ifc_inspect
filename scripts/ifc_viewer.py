"""
IFC Viewer Script - Prepares IFC file for web visualization
This script runs first to enable 3D viewing of the IFC model in the browser
Uses IfcOpenShell to extract geometry data for BIMSurfer/Three.js viewer
"""

import os
import json
import shutil
import struct
import tempfile


def convert_ifc_to_threejs_json(ifc_path, output_json_path):
    """
    Convert IFC file to Three.js compatible JSON format using IfcOpenShell.
    Extracts geometry as vertices and faces for direct rendering.
    Falls back to mesh-less mode if geometry extraction fails.
    """
    try:
        import ifcopenshell
        import ifcopenshell.geom
        
        f = ifcopenshell.open(ifc_path)
        
        # Use simple settings without OpenCASCADE (faster, works without pythonocc)
        settings = ifcopenshell.geom.settings()
        
        # Collect geometry data
        geometry_data = {
            'metadata': {
                'project_name': '',
                'element_count': 0,
                'buildings_count': 0,
                'stories_count': 0,
                'has_geometry': False
            },
            'materials': [],
            'meshes': []
        }
        
        # Get project info
        projects = f.by_type("IfcProject")
        if projects:
            geometry_data['metadata']['project_name'] = projects[0].Name or "Unknown Project"
        
        buildings = f.by_type("IfcBuilding")
        stories = f.by_type("IfcBuildingStorey")
        geometry_data['metadata']['buildings_count'] = len(buildings)
        geometry_data['metadata']['stories_count'] = len(stories)
        
        # Process elements with geometry - limit to first 500 for performance
        element_count = 0
        products = f.by_type("IfcProduct")
        geometry_data['metadata']['element_count'] = len(products)
        
        material_map = {}
        material_index = 0
        max_elements = 500  # Limit for web performance
        
        for product in products:
            if element_count >= max_elements:
                break
                
            if not hasattr(product, 'Representation') or not product.Representation:
                continue
            
            try:
                shape = ifcopenshell.geom.create_shape(settings, product)
                
                if not shape:
                    continue
                
                geom = shape.geometry
                vertices = geom.verts
                faces = geom.faces
                
                if len(vertices) < 3 or len(faces) < 1:
                    continue
                
                # Get or create material
                mat_key = str(getattr(geom, 'material', 'default'))
                if mat_key not in material_map:
                    material_map[mat_key] = material_index
                    color = [0.7, 0.7, 0.8]
                    geometry_data['materials'].append({
                        'name': mat_key,
                        'color': color
                    })
                    material_index += 1
                
                mesh_data = {
                    'type': product.is_a(),
                    'name': getattr(product, 'Name', 'Unnamed'),
                    'materialIndex': material_map[mat_key],
                    'vertices': vertices,
                    'faces': faces
                }
                
                geometry_data['meshes'].append(mesh_data)
                element_count += 1
                geometry_data['metadata']['has_geometry'] = True
                
            except Exception as e:
                continue
        
        # Save JSON data
        with open(output_json_path, 'w', encoding='utf-8') as out:
            json.dump(geometry_data, out, indent=2)
        
        has_geom = element_count > 0
        print(f"Created Three.js JSON: {element_count}/{len(products)} elements with geometry ({'success' if has_geom else 'no geometry data'})")
        return output_json_path
        
    except Exception as e:
        print(f"Error converting IFC to Three.js JSON: {e}")
        # Create minimal metadata-only JSON as fallback
        try:
            import ifcopenshell
            f = ifcopenshell.open(ifc_path)
            
            projects = f.by_type("IfcProject")
            buildings = f.by_type("IfcBuilding")
            stories = f.by_type("IfcBuildingStorey")
            products = f.by_type("IfcProduct")
            
            fallback_data = {
                'metadata': {
                    'project_name': projects[0].Name if projects else "Unknown",
                    'element_count': len(products),
                    'buildings_count': len(buildings),
                    'stories_count': len(stories),
                    'has_geometry': False,
                    'error': str(e)
                },
                'materials': [],
                'meshes': []
            }
            
            with open(output_json_path, 'w', encoding='utf-8') as out:
                json.dump(fallback_data, out, indent=2)
            
            print(f"Created fallback metadata-only JSON")
            return output_json_path
        except:
            return None


def convert_ifc_to_json(ifc_path, output_json_path):
    """
    Convert IFC file to a simplified JSON format for Three.js viewer.
    Wrapper for convert_ifc_to_threejs_json.
    """
    return convert_ifc_to_threejs_json(ifc_path, output_json_path)


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
