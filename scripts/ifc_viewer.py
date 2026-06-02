"""
IFC Viewer Script - Prepares IFC file for web visualization
This script runs first to enable 3D viewing of the IFC model in the browser
"""

import os
import json
import shutil


def prepare_ifc_for_viewer(ifc_path, session_folder):
    """
    Prepare IFC file for web visualization.
    Returns information about the IFC file for the viewer.
    
    In a more advanced implementation, this could:
    - Convert IFC to glTF/Three.js format
    - Generate thumbnails
    - Extract geometry data
    
    For now, we simply validate the file and return its path for direct viewing.
    """
    
    if not os.path.exists(ifc_path):
        raise FileNotFoundError(f"IFC file not found: {ifc_path}")
    
    # Validate IFC file
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
