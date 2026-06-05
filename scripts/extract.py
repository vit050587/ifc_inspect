import ifcopenshell
import ifcopenshell.util.element
import ifcopenshell.util.unit
import json
import os
import math
from collections import defaultdict
import pandas as pd
import numpy as np
import re

# Пути к файлам
IFC_FILE = "data/ifc модель КР.ifc"
JSON_OUTPUT = "data/ifc_extracted_full.json"
EXCEL_OUTPUT = "data/summary.xlsx"

# Список всех необходимых параметров (колонки)
REQUIRED_COLUMNS = [
    # Element Specific
    "Element Specific:GlobalId",
    "Element Specific:LongName",
    "Element Specific:Name",
    "Element Specific:ObjectType",
    "Element Specific:PredefinedType",
    "Element Specific:Tag",
    # ExpCheckBuilding
    "ExpCheckBuilding:MGE_BuildingAddress",
    "ExpCheckBuilding:MGE_Customer",
    "ExpCheckBuilding:MGE_Designer",
    "ExpCheckBuilding:MGE_ElevationOfRefHeight",
    "ExpCheckBuilding:MGE_ElevationOfTerrain",
    "ExpCheckBuilding:MGE_FunctionalUse",
    "ExpCheckBuilding:MGE_Korpus",
    "ExpCheckBuilding:MGE_NumOfSection",
    "ExpCheckBuilding:MGE_ObjectName",
    "ExpCheckBuilding:MGE_ProjectCode",
    "ExpCheckBuilding:MGE_ProjectName",
    "ExpCheckBuilding:MGE_Section",
    # ExpCheckBuildingStorey
    "ExpCheckBuildingStorey:MGE_ComfortLevel",
    # ExpCheck_Assembly
    "ExpCheck_Assembly:MGE_AssemblyPlace",
    "ExpCheck_Assembly:MGE_Gost",
    "ExpCheck_Assembly:MGE_IsExternal",
    "ExpCheck_Assembly:MGE_Name",
    "ExpCheck_Assembly:MGE_Position",
    # ExpCheck_Beam
    "ExpCheck_Beam:MGE_BeamType",
    "ExpCheck_Beam:MGE_ElementCode",
    "ExpCheck_Beam:MGE_Gost",
    "ExpCheck_Beam:MGE_Material",
    "ExpCheck_Beam:MGE_MaterialCode",
    "ExpCheck_Beam:MGE_Name",
    "ExpCheck_Beam:MGE_Position",
    "ExpCheck_Beam:MGE_SeriaNumber",
    "ExpCheck_Beam:MGE_SteelGrade",
    # ExpCheck_BeamReinforcement
    "ExpCheck_BeamReinforcement:ReinforceStrengthClass",
    # ExpCheck_Column
    "ExpCheck_Column:MGE_ColumnType",
    "ExpCheck_Column:MGE_ElementCode",
    "ExpCheck_Column:MGE_Gost",
    "ExpCheck_Column:MGE_Material",
    "ExpCheck_Column:MGE_MaterialCode",
    "ExpCheck_Column:MGE_Name",
    "ExpCheck_Column:MGE_Position",
    "ExpCheck_Column:MGE_SeriaNumber",
    "ExpCheck_Column:MGE_SteelGrade",
    # ExpCheck_ColumnReinforcement
    "ExpCheck_ColumnReinforcement:MGE_ReinforceStrengthClass",
    # ExpCheck_MaterialConcrete
    "ExpCheck_MaterialConcrete:MGE_ConcreteGost",
    "ExpCheck_MaterialConcrete:MGE_ConcreteGrade",
    "ExpCheck_MaterialConcrete:MGE_FreezeDurability",
    "ExpCheck_MaterialConcrete:MGE_WaterResist",
    # ExpCheck_Ramp
    "ExpCheck_Ramp:MGE_ElementCode",
    "ExpCheck_Ramp:MGE_Gost",
    "ExpCheck_Ramp:MGE_Material",
    "ExpCheck_Ramp:MGE_MaterialCode",
    "ExpCheck_Ramp:MGE_Name",
    "ExpCheck_Ramp:MGE_Position",
    "ExpCheck_Ramp:MGE_Section",
    # ExpCheck_Slab
    "ExpCheck_Slab:MGE_ElementCode",
    "ExpCheck_Slab:MGE_Gost",
    "ExpCheck_Slab:MGE_Material",
    "ExpCheck_Slab:MGE_MaterialCode",
    "ExpCheck_Slab:MGE_Name",
    "ExpCheck_Slab:MGE_Position",
    "ExpCheck_Slab:MGE_SlabType",
    # ExpCheck_SlabReinforcement
    "ExpCheck_SlabReinforcement:MGE_ReinforceStrengthClass",
    # ExpCheck_StairFlight
    "ExpCheck_StairFlight:MGE_ElementCode",
    "ExpCheck_StairFlight:MGE_Gost",
    "ExpCheck_StairFlight:MGE_Material",
    "ExpCheck_StairFlight:MGE_MaterialCode",
    "ExpCheck_StairFlight:MGE_Name",
    "ExpCheck_StairFlight:MGE_Position",
    "ExpCheck_StairFlight:MGE_Section",
    # ExpCheck_Wall
    "ExpCheck_Wall:MGE_ElementCode",
    "ExpCheck_Wall:MGE_Gost",
    "ExpCheck_Wall:MGE_Material",
    "ExpCheck_Wall:MGE_MaterialCode",
    "ExpCheck_Wall:MGE_Name",
    "ExpCheck_Wall:MGE_Position",
    # ExpCheck_WallReinforcement
    "ExpCheck_WallReinforcement:MGE_ReinforceStrengthClass",
    # Ifc Class
    "Ifc Class",
    # Pset_BeamCommon
    "Pset_BeamCommon:IsExternal",
    "Pset_BeamCommon:LoadBearing",
    "Pset_BeamCommon:Reference",
    "Pset_BeamCommon:Roll",
    "Pset_BeamCommon:Slope",
    "Pset_BeamCommon:Span",
    # Pset_BuildingCommon
    "Pset_BuildingCommon:ConstructionMethod",
    "Pset_BuildingCommon:FireProtectionClass",
    "Pset_BuildingCommon:IsLandmarked",
    "Pset_BuildingCommon:NumberOfStoreys",
    "Pset_BuildingCommon:Reference",
    # Pset_BuildingElementProxyCommon
    "Pset_BuildingElementProxyCommon:IsExternal",
    "Pset_BuildingElementProxyCommon:Reference",
    # Pset_BuildingStoreyCommon
    "Pset_BuildingStoreyCommon:AboveGround",
    "Pset_BuildingStoreyCommon:EntranceLevel",
    "Pset_BuildingStoreyCommon:Reference",
    "Pset_BuildingStoreyCommon:SprinklerProtection",
    "Pset_BuildingStoreyCommon:SprinklerProtectionAutomatic",
    # Pset_BuildingSystemCommon
    "Pset_BuildingSystemCommon:Reference",
    # Pset_ColumnCommon
    "Pset_ColumnCommon:IsExternal",
    "Pset_ColumnCommon:LoadBearing",
    "Pset_ColumnCommon:Reference",
    "Pset_ColumnCommon:Slope",
    "Pset_ColumnCommon:ThermalTransmittance",
    # Pset_ConcreteElementGeneral
    "Pset_ConcreteElementGeneral:ConcreteCover",
    "Pset_ConcreteElementGeneral:ConcreteCoverAtLinks",
    "Pset_ConcreteElementGeneral:ConstructionMethod",
    "Pset_ConcreteElementGeneral:ReinforcementVolumeRatio",
    "Pset_ConcreteElementGeneral:StructuralClass",
    # Pset_ElementAssemblyCommon
    "Pset_ElementAssemblyCommon:Reference",
    # Pset_EnvironmentalImpactIndicators
    "Pset_EnvironmentalImpactIndicators:Reference",
    # Pset_ManufacturerTypeInformation
    "Pset_ManufacturerTypeInformation:AssemblyPlace",
    "Pset_ManufacturerTypeInformation:Manufacturer",
    # Pset_MemberCommon
    "Pset_MemberCommon:IsExternal",
    "Pset_MemberCommon:LoadBearing",
    "Pset_MemberCommon:Reference",
    # Pset_OpeningElementCommon
    "Pset_OpeningElementCommon:Reference",
    # Pset_PlateCommon
    "Pset_PlateCommon:IsExternal",
    "Pset_PlateCommon:LoadBearing",
    "Pset_PlateCommon:Reference",
    # Pset_RampCommon
    "Pset_RampCommon:IsExternal",
    "Pset_RampCommon:Reference",
    # Pset_RampFlightCommon
    "Pset_RampFlightCommon:Reference",
    "Pset_RampFlightCommon:Slope",
    # Pset_ReinforcementBarPitchOfBeam
    "Pset_ReinforcementBarPitchOfBeam:Reference",
    # Pset_ReinforcementBarPitchOfColumn
    "Pset_ReinforcementBarPitchOfColumn:Reference",
    # Pset_ReinforcementBarPitchOfSlab
    "Pset_ReinforcementBarPitchOfSlab:Reference",
    # Pset_ReinforcementBarPitchOfWall
    "Pset_ReinforcementBarPitchOfWall:Reference",
    # Pset_SlabCommon
    "Pset_SlabCommon:IsExternal",
    "Pset_SlabCommon:LoadBearing",
    "Pset_SlabCommon:PitchAngle",
    "Pset_SlabCommon:Reference",
    "Pset_SlabCommon:ThermalTransmittance",
    # Pset_StairFlightCommon
    "Pset_StairFlightCommon:IsExternal",
    "Pset_StairFlightCommon:LoadBearing",
    "Pset_StairFlightCommon:Reference",
    # Pset_WallCommon
    "Pset_WallCommon:ExtendToStructure",
    "Pset_WallCommon:IsExternal",
    "Pset_WallCommon:LoadBearing",
    "Pset_WallCommon:Reference",
    "Pset_WallCommon:ThermalTransmittance",
    # Qto_BeamBaseQuantities
    "Qto_BeamBaseQuantities:CrossSectionArea",
    "Qto_BeamBaseQuantities:GrossSurfaceArea",
    "Qto_BeamBaseQuantities:GrossVolume",
    "Qto_BeamBaseQuantities:Length",
    "Qto_BeamBaseQuantities:NetSurfaceArea",
    "Qto_BeamBaseQuantities:NetVolume",
    "Qto_BeamBaseQuantities:OuterSurfaceArea",
    # Qto_ColumnBaseQuantities
    "Qto_ColumnBaseQuantities:CrossSectionArea",
    "Qto_ColumnBaseQuantities:GrossVolume",
    "Qto_ColumnBaseQuantities:Length",
    "Qto_ColumnBaseQuantities:NetVolume",
    "Qto_ColumnBaseQuantities:OuterSurfaceArea",
    # Qto_MemberBaseQuantities
    "Qto_MemberBaseQuantities:CrossSectionArea",
    "Qto_MemberBaseQuantities:GrossSurfaceArea",
    "Qto_MemberBaseQuantities:GrossVolume",
    "Qto_MemberBaseQuantities:Length",
    "Qto_MemberBaseQuantities:NetSurfaceArea",
    "Qto_MemberBaseQuantities:NetVolume",
    "Qto_MemberBaseQuantities:OuterSurfaceArea",
    # Qto_OpeningElementBaseQuantities
    "Qto_OpeningElementBaseQuantities:Area",
    "Qto_OpeningElementBaseQuantities:Depth",
    "Qto_OpeningElementBaseQuantities:Height",
    "Qto_OpeningElementBaseQuantities:Width",
    # Qto_SlabBaseQuantities
    "Qto_SlabBaseQuantities:GrossArea",
    "Qto_SlabBaseQuantities:GrossVolume",
    "Qto_SlabBaseQuantities:NetArea",
    "Qto_SlabBaseQuantities:NetVolume",
    "Qto_SlabBaseQuantities:Perimeter",
    "Qto_SlabBaseQuantities:Width",
    # Qto_StairFlightBaseQuantities
    "Qto_StairFlightBaseQuantities:GrossVolume",
    "Qto_StairFlightBaseQuantities:Length",
    "Qto_StairFlightBaseQuantities:NetVolume",
    # Qto_WallBaseQuantities
    "Qto_WallBaseQuantities:GrossSideArea",
    "Qto_WallBaseQuantities:GrossVolume",
    "Qto_WallBaseQuantities:Height",
    "Qto_WallBaseQuantities:Length",
    "Qto_WallBaseQuantities:NetSideArea",
    "Qto_WallBaseQuantities:NetVolume",
    "Qto_WallBaseQuantities:Width",
    # Storey
    "Storey",
    # Type
    "Type:GlobalId",
    "Type:Name",
]


def get_unit_scale(ifc_file):
    """Получить масштаб единиц из IFC файла"""
    try:
        units = ifcopenshell.util.unit.get_unit_scale(ifc_file)
        return units
    except:
        return 1.0


def calculate_volume_by_geometry(shape, unit_scale):
    """Расчет объема через bounding box с учетом масштаба"""
    try:
        bbox_min = shape.bounding_box.lower_corner
        bbox_max = shape.bounding_box.upper_corner
        
        dx = (bbox_max[0] - bbox_min[0]) * unit_scale
        dy = (bbox_max[1] - bbox_min[1]) * unit_scale
        dz = (bbox_max[2] - bbox_min[2]) * unit_scale
        
        return abs(dx * dy * dz)
    except Exception as e:
        return 0.0


def calculate_area_by_geometry(shape, unit_scale):
    """Расчет площади поверхности через геометрию"""
    try:
        # Получаем площадь поверхности из shape
        area = shape.geometry.surface_area
        # Учитываем масштаб (площадь в квадрате)
        return area * (unit_scale ** 2)
    except Exception as e:
        return 0.0


def get_property_value(psets, pset_name, prop_name):
    """Получить значение свойства из psets"""
    if pset_name in psets:
        props = psets[pset_name]
        if prop_name in props:
            val = props[prop_name]
            if val is not None and prop_name != 'id':
                return val
    return None


def extract_element_properties(element, psets):
    """Извлечь все свойства элемента согласно REQUIRED_COLUMNS"""
    data = {}
    
    # Element Specific properties
    data["Element Specific:GlobalId"] = getattr(element, 'GlobalId', None)
    data["Element Specific:LongName"] = getattr(element, 'LongName', None)
    data["Element Specific:Name"] = getattr(element, 'Name', None)
    data["Element Specific:ObjectType"] = getattr(element, 'ObjectType', None)
    data["Element Specific:PredefinedType"] = getattr(element, 'PredefinedType', None)
    data["Element Specific:Tag"] = getattr(element, 'Tag', None)
    
    # ExpCheckBuilding
    data["ExpCheckBuilding:MGE_BuildingAddress"] = get_property_value(psets, "ExpCheck_Building", "MGE_BuildingAddress")
    data["ExpCheckBuilding:MGE_Customer"] = get_property_value(psets, "ExpCheck_Building", "MGE_Customer")
    data["ExpCheckBuilding:MGE_Designer"] = get_property_value(psets, "ExpCheck_Building", "MGE_Designer")
    data["ExpCheckBuilding:MGE_ElevationOfRefHeight"] = get_property_value(psets, "ExpCheck_Building", "MGE_ElevationOfRefHeight")
    data["ExpCheckBuilding:MGE_ElevationOfTerrain"] = get_property_value(psets, "ExpCheck_Building", "MGE_ElevationOfTerrain")
    data["ExpCheckBuilding:MGE_FunctionalUse"] = get_property_value(psets, "ExpCheck_Building", "MGE_FunctionalUse")
    data["ExpCheckBuilding:MGE_Korpus"] = get_property_value(psets, "ExpCheck_Building", "MGE_Korpus")
    data["ExpCheckBuilding:MGE_NumOfSection"] = get_property_value(psets, "ExpCheck_Building", "MGE_NumOfSection")
    data["ExpCheckBuilding:MGE_ObjectName"] = get_property_value(psets, "ExpCheck_Building", "MGE_ObjectName")
    data["ExpCheckBuilding:MGE_ProjectCode"] = get_property_value(psets, "ExpCheck_Building", "MGE_ProjectCode")
    data["ExpCheckBuilding:MGE_ProjectName"] = get_property_value(psets, "ExpCheck_Building", "MGE_ProjectName")
    data["ExpCheckBuilding:MGE_Section"] = get_property_value(psets, "ExpCheck_Building", "MGE_Section")
    
    # ExpCheckBuildingStorey
    data["ExpCheckBuildingStorey:MGE_ComfortLevel"] = get_property_value(psets, "ExpCheck_BuildingStorey", "MGE_ComfortLevel")
    
    # ExpCheck_Assembly
    data["ExpCheck_Assembly:MGE_AssemblyPlace"] = get_property_value(psets, "ExpCheck_Assembly", "MGE_AssemblyPlace")
    data["ExpCheck_Assembly:MGE_Gost"] = get_property_value(psets, "ExpCheck_Assembly", "MGE_Gost")
    data["ExpCheck_Assembly:MGE_IsExternal"] = get_property_value(psets, "ExpCheck_Assembly", "MGE_IsExternal")
    data["ExpCheck_Assembly:MGE_Name"] = get_property_value(psets, "ExpCheck_Assembly", "MGE_Name")
    data["ExpCheck_Assembly:MGE_Position"] = get_property_value(psets, "ExpCheck_Assembly", "MGE_Position")
    
    # ExpCheck_Beam
    data["ExpCheck_Beam:MGE_BeamType"] = get_property_value(psets, "ExpCheck_Beam", "MGE_BeamType")
    data["ExpCheck_Beam:MGE_ElementCode"] = get_property_value(psets, "ExpCheck_Beam", "MGE_ElementCode")
    data["ExpCheck_Beam:MGE_Gost"] = get_property_value(psets, "ExpCheck_Beam", "MGE_Gost")
    data["ExpCheck_Beam:MGE_Material"] = get_property_value(psets, "ExpCheck_Beam", "MGE_Material")
    data["ExpCheck_Beam:MGE_MaterialCode"] = get_property_value(psets, "ExpCheck_Beam", "MGE_MaterialCode")
    data["ExpCheck_Beam:MGE_Name"] = get_property_value(psets, "ExpCheck_Beam", "MGE_Name")
    data["ExpCheck_Beam:MGE_Position"] = get_property_value(psets, "ExpCheck_Beam", "MGE_Position")
    data["ExpCheck_Beam:MGE_SeriaNumber"] = get_property_value(psets, "ExpCheck_Beam", "MGE_SeriaNumber")
    data["ExpCheck_Beam:MGE_SteelGrade"] = get_property_value(psets, "ExpCheck_Beam", "MGE_SteelGrade")
    
    # ExpCheck_BeamReinforcement
    data["ExpCheck_BeamReinforcement:ReinforceStrengthClass"] = get_property_value(psets, "ExpCheck_BeamReinforcement", "ReinforceStrengthClass")
    
    # ExpCheck_Column
    data["ExpCheck_Column:MGE_ColumnType"] = get_property_value(psets, "ExpCheck_Column", "MGE_ColumnType")
    data["ExpCheck_Column:MGE_ElementCode"] = get_property_value(psets, "ExpCheck_Column", "MGE_ElementCode")
    data["ExpCheck_Column:MGE_Gost"] = get_property_value(psets, "ExpCheck_Column", "MGE_Gost")
    data["ExpCheck_Column:MGE_Material"] = get_property_value(psets, "ExpCheck_Column", "MGE_Material")
    data["ExpCheck_Column:MGE_MaterialCode"] = get_property_value(psets, "ExpCheck_Column", "MGE_MaterialCode")
    data["ExpCheck_Column:MGE_Name"] = get_property_value(psets, "ExpCheck_Column", "MGE_Name")
    data["ExpCheck_Column:MGE_Position"] = get_property_value(psets, "ExpCheck_Column", "MGE_Position")
    data["ExpCheck_Column:MGE_SeriaNumber"] = get_property_value(psets, "ExpCheck_Column", "MGE_SeriaNumber")
    data["ExpCheck_Column:MGE_SteelGrade"] = get_property_value(psets, "ExpCheck_Column", "MGE_SteelGrade")
    
    # ExpCheck_ColumnReinforcement
    data["ExpCheck_ColumnReinforcement:MGE_ReinforceStrengthClass"] = get_property_value(psets, "ExpCheck_ColumnReinforcement", "MGE_ReinforceStrengthClass")
    
    # ExpCheck_MaterialConcrete
    data["ExpCheck_MaterialConcrete:MGE_ConcreteGost"] = get_property_value(psets, "ExpCheck_MaterialConcrete", "MGE_ConcreteGost")
    data["ExpCheck_MaterialConcrete:MGE_ConcreteGrade"] = get_property_value(psets, "ExpCheck_MaterialConcrete", "MGE_ConcreteGrade")
    data["ExpCheck_MaterialConcrete:MGE_FreezeDurability"] = get_property_value(psets, "ExpCheck_MaterialConcrete", "MGE_FreezeDurability")
    data["ExpCheck_MaterialConcrete:MGE_WaterResist"] = get_property_value(psets, "ExpCheck_MaterialConcrete", "MGE_WaterResist")
    
    # ExpCheck_Ramp
    data["ExpCheck_Ramp:MGE_ElementCode"] = get_property_value(psets, "ExpCheck_Ramp", "MGE_ElementCode")
    data["ExpCheck_Ramp:MGE_Gost"] = get_property_value(psets, "ExpCheck_Ramp", "MGE_Gost")
    data["ExpCheck_Ramp:MGE_Material"] = get_property_value(psets, "ExpCheck_Ramp", "MGE_Material")
    data["ExpCheck_Ramp:MGE_MaterialCode"] = get_property_value(psets, "ExpCheck_Ramp", "MGE_MaterialCode")
    data["ExpCheck_Ramp:MGE_Name"] = get_property_value(psets, "ExpCheck_Ramp", "MGE_Name")
    data["ExpCheck_Ramp:MGE_Position"] = get_property_value(psets, "ExpCheck_Ramp", "MGE_Position")
    data["ExpCheck_Ramp:MGE_Section"] = get_property_value(psets, "ExpCheck_Ramp", "MGE_Section")
    
    # ExpCheck_Slab
    data["ExpCheck_Slab:MGE_ElementCode"] = get_property_value(psets, "ExpCheck_Slab", "MGE_ElementCode")
    data["ExpCheck_Slab:MGE_Gost"] = get_property_value(psets, "ExpCheck_Slab", "MGE_Gost")
    data["ExpCheck_Slab:MGE_Material"] = get_property_value(psets, "ExpCheck_Slab", "MGE_Material")
    data["ExpCheck_Slab:MGE_MaterialCode"] = get_property_value(psets, "ExpCheck_Slab", "MGE_MaterialCode")
    data["ExpCheck_Slab:MGE_Name"] = get_property_value(psets, "ExpCheck_Slab", "MGE_Name")
    data["ExpCheck_Slab:MGE_Position"] = get_property_value(psets, "ExpCheck_Slab", "MGE_Position")
    data["ExpCheck_Slab:MGE_SlabType"] = get_property_value(psets, "ExpCheck_Slab", "MGE_SlabType")
    
    # ExpCheck_SlabReinforcement
    data["ExpCheck_SlabReinforcement:MGE_ReinforceStrengthClass"] = get_property_value(psets, "ExpCheck_SlabReinforcement", "MGE_ReinforceStrengthClass")
    
    # ExpCheck_StairFlight
    data["ExpCheck_StairFlight:MGE_ElementCode"] = get_property_value(psets, "ExpCheck_Stair", "MGE_ElementCode")
    data["ExpCheck_StairFlight:MGE_Gost"] = get_property_value(psets, "ExpCheck_Stair", "MGE_Gost")
    data["ExpCheck_StairFlight:MGE_Material"] = get_property_value(psets, "ExpCheck_Stair", "MGE_Material")
    data["ExpCheck_StairFlight:MGE_MaterialCode"] = get_property_value(psets, "ExpCheck_Stair", "MGE_MaterialCode")
    data["ExpCheck_StairFlight:MGE_Name"] = get_property_value(psets, "ExpCheck_Stair", "MGE_Name")
    data["ExpCheck_StairFlight:MGE_Position"] = get_property_value(psets, "ExpCheck_Stair", "MGE_Position")
    data["ExpCheck_StairFlight:MGE_Section"] = get_property_value(psets, "ExpCheck_Stair", "MGE_Section")
    
    # ExpCheck_Wall
    data["ExpCheck_Wall:MGE_ElementCode"] = get_property_value(psets, "ExpCheck_Wall", "MGE_ElementCode")
    data["ExpCheck_Wall:MGE_Gost"] = get_property_value(psets, "ExpCheck_Wall", "MGE_Gost")
    data["ExpCheck_Wall:MGE_Material"] = get_property_value(psets, "ExpCheck_Wall", "MGE_Material")
    data["ExpCheck_Wall:MGE_MaterialCode"] = get_property_value(psets, "ExpCheck_Wall", "MGE_MaterialCode")
    data["ExpCheck_Wall:MGE_Name"] = get_property_value(psets, "ExpCheck_Wall", "MGE_Name")
    data["ExpCheck_Wall:MGE_Position"] = get_property_value(psets, "ExpCheck_Wall", "MGE_Position")
    
    # ExpCheck_WallReinforcement
    data["ExpCheck_WallReinforcement:MGE_ReinforceStrengthClass"] = get_property_value(psets, "ExpCheck_WallReinforcement", "MGE_ReinforceStrengthClass")
    
    # Ifc Class
    data["Ifc Class"] = element.is_a()
    
    # Pset_BeamCommon
    data["Pset_BeamCommon:IsExternal"] = get_property_value(psets, "Pset_BeamCommon", "IsExternal")
    data["Pset_BeamCommon:LoadBearing"] = get_property_value(psets, "Pset_BeamCommon", "LoadBearing")
    data["Pset_BeamCommon:Reference"] = get_property_value(psets, "Pset_BeamCommon", "Reference")
    data["Pset_BeamCommon:Roll"] = get_property_value(psets, "Pset_BeamCommon", "Roll")
    data["Pset_BeamCommon:Slope"] = get_property_value(psets, "Pset_BeamCommon", "Slope")
    data["Pset_BeamCommon:Span"] = get_property_value(psets, "Pset_BeamCommon", "Span")
    
    # Pset_BuildingCommon
    data["Pset_BuildingCommon:ConstructionMethod"] = get_property_value(psets, "Pset_BuildingCommon", "ConstructionMethod")
    data["Pset_BuildingCommon:FireProtectionClass"] = get_property_value(psets, "Pset_BuildingCommon", "FireProtectionClass")
    data["Pset_BuildingCommon:IsLandmarked"] = get_property_value(psets, "Pset_BuildingCommon", "IsLandmarked")
    data["Pset_BuildingCommon:NumberOfStoreys"] = get_property_value(psets, "Pset_BuildingCommon", "NumberOfStoreys")
    data["Pset_BuildingCommon:Reference"] = get_property_value(psets, "Pset_BuildingCommon", "Reference")
    
    # Pset_BuildingElementProxyCommon
    data["Pset_BuildingElementProxyCommon:IsExternal"] = get_property_value(psets, "Pset_BuildingElementProxyCommon", "IsExternal")
    data["Pset_BuildingElementProxyCommon:Reference"] = get_property_value(psets, "Pset_BuildingElementProxyCommon", "Reference")
    
    # Pset_BuildingStoreyCommon
    data["Pset_BuildingStoreyCommon:AboveGround"] = get_property_value(psets, "Pset_BuildingStoreyCommon", "AboveGround")
    data["Pset_BuildingStoreyCommon:EntranceLevel"] = get_property_value(psets, "Pset_BuildingStoreyCommon", "EntranceLevel")
    data["Pset_BuildingStoreyCommon:Reference"] = get_property_value(psets, "Pset_BuildingStoreyCommon", "Reference")
    data["Pset_BuildingStoreyCommon:SprinklerProtection"] = get_property_value(psets, "Pset_BuildingStoreyCommon", "SprinklerProtection")
    data["Pset_BuildingStoreyCommon:SprinklerProtectionAutomatic"] = get_property_value(psets, "Pset_BuildingStoreyCommon", "SprinklerProtectionAutomatic")
    
    # Pset_BuildingSystemCommon
    data["Pset_BuildingSystemCommon:Reference"] = get_property_value(psets, "Pset_BuildingSystemCommon", "Reference")
    
    # Pset_ColumnCommon
    data["Pset_ColumnCommon:IsExternal"] = get_property_value(psets, "Pset_ColumnCommon", "IsExternal")
    data["Pset_ColumnCommon:LoadBearing"] = get_property_value(psets, "Pset_ColumnCommon", "LoadBearing")
    data["Pset_ColumnCommon:Reference"] = get_property_value(psets, "Pset_ColumnCommon", "Reference")
    data["Pset_ColumnCommon:Slope"] = get_property_value(psets, "Pset_ColumnCommon", "Slope")
    data["Pset_ColumnCommon:ThermalTransmittance"] = get_property_value(psets, "Pset_ColumnCommon", "ThermalTransmittance")
    
    # Pset_ConcreteElementGeneral
    data["Pset_ConcreteElementGeneral:ConcreteCover"] = get_property_value(psets, "Pset_ConcreteElementGeneral", "ConcreteCover")
    data["Pset_ConcreteElementGeneral:ConcreteCoverAtLinks"] = get_property_value(psets, "Pset_ConcreteElementGeneral", "ConcreteCoverAtLinks")
    data["Pset_ConcreteElementGeneral:ConstructionMethod"] = get_property_value(psets, "Pset_ConcreteElementGeneral", "ConstructionMethod")
    data["Pset_ConcreteElementGeneral:ReinforcementVolumeRatio"] = get_property_value(psets, "Pset_ConcreteElementGeneral", "ReinforcementVolumeRatio")
    data["Pset_ConcreteElementGeneral:StructuralClass"] = get_property_value(psets, "Pset_ConcreteElementGeneral", "StructuralClass")
    
    # Pset_ElementAssemblyCommon
    data["Pset_ElementAssemblyCommon:Reference"] = get_property_value(psets, "Pset_ElementAssemblyCommon", "Reference")
    
    # Pset_EnvironmentalImpactIndicators
    data["Pset_EnvironmentalImpactIndicators:Reference"] = get_property_value(psets, "Pset_EnvironmentalImpactIndicators", "Reference")
    
    # Pset_ManufacturerTypeInformation
    data["Pset_ManufacturerTypeInformation:AssemblyPlace"] = get_property_value(psets, "Pset_ManufacturerTypeInformation", "AssemblyPlace")
    data["Pset_ManufacturerTypeInformation:Manufacturer"] = get_property_value(psets, "Pset_ManufacturerTypeInformation", "Manufacturer")
    
    # Pset_MemberCommon
    data["Pset_MemberCommon:IsExternal"] = get_property_value(psets, "Pset_MemberCommon", "IsExternal")
    data["Pset_MemberCommon:LoadBearing"] = get_property_value(psets, "Pset_MemberCommon", "LoadBearing")
    data["Pset_MemberCommon:Reference"] = get_property_value(psets, "Pset_MemberCommon", "Reference")
    
    # Pset_OpeningElementCommon
    data["Pset_OpeningElementCommon:Reference"] = get_property_value(psets, "Pset_OpeningElementCommon", "Reference")
    
    # Pset_PlateCommon
    data["Pset_PlateCommon:IsExternal"] = get_property_value(psets, "Pset_PlateCommon", "IsExternal")
    data["Pset_PlateCommon:LoadBearing"] = get_property_value(psets, "Pset_PlateCommon", "LoadBearing")
    data["Pset_PlateCommon:Reference"] = get_property_value(psets, "Pset_PlateCommon", "Reference")
    
    # Pset_RampCommon
    data["Pset_RampCommon:IsExternal"] = get_property_value(psets, "Pset_RampCommon", "IsExternal")
    data["Pset_RampCommon:Reference"] = get_property_value(psets, "Pset_RampCommon", "Reference")
    
    # Pset_RampFlightCommon
    data["Pset_RampFlightCommon:Reference"] = get_property_value(psets, "Pset_RampFlightCommon", "Reference")
    data["Pset_RampFlightCommon:Slope"] = get_property_value(psets, "Pset_RampFlightCommon", "Slope")
    
    # Pset_ReinforcementBarPitchOfBeam
    data["Pset_ReinforcementBarPitchOfBeam:Reference"] = get_property_value(psets, "Pset_ReinforcementBarPitchOfBeam", "Reference")
    
    # Pset_ReinforcementBarPitchOfColumn
    data["Pset_ReinforcementBarPitchOfColumn:Reference"] = get_property_value(psets, "Pset_ReinforcementBarPitchOfColumn", "Reference")
    
    # Pset_ReinforcementBarPitchOfSlab
    data["Pset_ReinforcementBarPitchOfSlab:Reference"] = get_property_value(psets, "Pset_ReinforcementBarPitchOfSlab", "Reference")
    
    # Pset_ReinforcementBarPitchOfWall
    data["Pset_ReinforcementBarPitchOfWall:Reference"] = get_property_value(psets, "Pset_ReinforcementBarPitchOfWall", "Reference")
    
    # Pset_SlabCommon
    data["Pset_SlabCommon:IsExternal"] = get_property_value(psets, "Pset_SlabCommon", "IsExternal")
    data["Pset_SlabCommon:LoadBearing"] = get_property_value(psets, "Pset_SlabCommon", "LoadBearing")
    data["Pset_SlabCommon:PitchAngle"] = get_property_value(psets, "Pset_SlabCommon", "PitchAngle")
    data["Pset_SlabCommon:Reference"] = get_property_value(psets, "Pset_SlabCommon", "Reference")
    data["Pset_SlabCommon:ThermalTransmittance"] = get_property_value(psets, "Pset_SlabCommon", "ThermalTransmittance")
    
    # Pset_StairFlightCommon
    data["Pset_StairFlightCommon:IsExternal"] = get_property_value(psets, "Pset_StairCommon", "IsExternal")
    data["Pset_StairFlightCommon:LoadBearing"] = get_property_value(psets, "Pset_StairCommon", "LoadBearing")
    data["Pset_StairFlightCommon:Reference"] = get_property_value(psets, "Pset_StairCommon", "Reference")
    
    # Pset_WallCommon
    data["Pset_WallCommon:ExtendToStructure"] = get_property_value(psets, "Pset_WallCommon", "ExtendToStructure")
    data["Pset_WallCommon:IsExternal"] = get_property_value(psets, "Pset_WallCommon", "IsExternal")
    data["Pset_WallCommon:LoadBearing"] = get_property_value(psets, "Pset_WallCommon", "LoadBearing")
    data["Pset_WallCommon:Reference"] = get_property_value(psets, "Pset_WallCommon", "Reference")
    data["Pset_WallCommon:ThermalTransmittance"] = get_property_value(psets, "Pset_WallCommon", "ThermalTransmittance")
    
    # Qto_BeamBaseQuantities
    data["Qto_BeamBaseQuantities:CrossSectionArea"] = get_property_value(psets, "Qto_BeamBaseQuantities", "CrossSectionArea")
    data["Qto_BeamBaseQuantities:GrossSurfaceArea"] = get_property_value(psets, "Qto_BeamBaseQuantities", "GrossSurfaceArea")
    data["Qto_BeamBaseQuantities:GrossVolume"] = get_property_value(psets, "Qto_BeamBaseQuantities", "GrossVolume")
    data["Qto_BeamBaseQuantities:Length"] = get_property_value(psets, "Qto_BeamBaseQuantities", "Length")
    data["Qto_BeamBaseQuantities:NetSurfaceArea"] = get_property_value(psets, "Qto_BeamBaseQuantities", "NetSurfaceArea")
    data["Qto_BeamBaseQuantities:NetVolume"] = get_property_value(psets, "Qto_BeamBaseQuantities", "NetVolume")
    data["Qto_BeamBaseQuantities:OuterSurfaceArea"] = get_property_value(psets, "Qto_BeamBaseQuantities", "OuterSurfaceArea")
    
    # Qto_ColumnBaseQuantities
    data["Qto_ColumnBaseQuantities:CrossSectionArea"] = get_property_value(psets, "Qto_ColumnBaseQuantities", "CrossSectionArea")
    data["Qto_ColumnBaseQuantities:GrossVolume"] = get_property_value(psets, "Qto_ColumnBaseQuantities", "GrossVolume")
    data["Qto_ColumnBaseQuantities:Length"] = get_property_value(psets, "Qto_ColumnBaseQuantities", "Length")
    data["Qto_ColumnBaseQuantities:NetVolume"] = get_property_value(psets, "Qto_ColumnBaseQuantities", "NetVolume")
    data["Qto_ColumnBaseQuantities:OuterSurfaceArea"] = get_property_value(psets, "Qto_ColumnBaseQuantities", "OuterSurfaceArea")
    
    # Qto_MemberBaseQuantities
    data["Qto_MemberBaseQuantities:CrossSectionArea"] = get_property_value(psets, "Qto_MemberBaseQuantities", "CrossSectionArea")
    data["Qto_MemberBaseQuantities:GrossSurfaceArea"] = get_property_value(psets, "Qto_MemberBaseQuantities", "GrossSurfaceArea")
    data["Qto_MemberBaseQuantities:GrossVolume"] = get_property_value(psets, "Qto_MemberBaseQuantities", "GrossVolume")
    data["Qto_MemberBaseQuantities:Length"] = get_property_value(psets, "Qto_MemberBaseQuantities", "Length")
    data["Qto_MemberBaseQuantities:NetSurfaceArea"] = get_property_value(psets, "Qto_MemberBaseQuantities", "NetSurfaceArea")
    data["Qto_MemberBaseQuantities:NetVolume"] = get_property_value(psets, "Qto_MemberBaseQuantities", "NetVolume")
    data["Qto_MemberBaseQuantities:OuterSurfaceArea"] = get_property_value(psets, "Qto_MemberBaseQuantities", "OuterSurfaceArea")
    
    # Qto_OpeningElementBaseQuantities
    data["Qto_OpeningElementBaseQuantities:Area"] = get_property_value(psets, "Qto_OpeningElementBaseQuantities", "Area")
    data["Qto_OpeningElementBaseQuantities:Depth"] = get_property_value(psets, "Qto_OpeningElementBaseQuantities", "Depth")
    data["Qto_OpeningElementBaseQuantities:Height"] = get_property_value(psets, "Qto_OpeningElementBaseQuantities", "Height")
    data["Qto_OpeningElementBaseQuantities:Width"] = get_property_value(psets, "Qto_OpeningElementBaseQuantities", "Width")
    
    # Qto_SlabBaseQuantities - извлекаем все динамические свойства включая материалы
    for prop_name in ["31_Утеплитель Полистирол экструдированный", "GrossArea", "GrossVolume", "NetArea", "NetVolume", "Perimeter", "Width", "Бетон В30 F150 W8", "Бетон В7.5 неармированный", "В25 F100 W6"]:
        data[f"Qto_SlabBaseQuantities:{prop_name}"] = get_property_value(psets, "Qto_SlabBaseQuantities", prop_name)
    
    # Извлекаем объемы слоев материалов из Qto_SlabBaseQuantities
    slab_materials = {}
    if "Qto_SlabBaseQuantities" in psets:
        for prop_name, prop_val in psets["Qto_SlabBaseQuantities"].items():
            if isinstance(prop_val, dict) and prop_val.get('type') == 'IfcPhysicalComplexQuantity':
                # Это слой материала - извлекаем объем из properties
                props = prop_val.get('properties', {})
                if 'NetVolume' in props or 'GrossVolume' in props:
                    vol = props.get('NetVolume', props.get('GrossVolume', 0))
                    slab_materials[prop_name] = vol
    
    data["slab_materials"] = slab_materials
    
    # Qto_StairFlightBaseQuantities
    data["Qto_StairFlightBaseQuantities:GrossVolume"] = get_property_value(psets, "Qto_StairFlightBaseQuantities", "GrossVolume")
    data["Qto_StairFlightBaseQuantities:Length"] = get_property_value(psets, "Qto_StairFlightBaseQuantities", "Length")
    data["Qto_StairFlightBaseQuantities:NetVolume"] = get_property_value(psets, "Qto_StairFlightBaseQuantities", "NetVolume")
    
    # Qto_WallBaseQuantities - извлекаем все свойства включая материалы
    for prop_name in ["GrossSideArea", "GrossVolume", "Height", "Length", "NetSideArea", "NetVolume", "Width", "Бетон В30 F150 W8", "В25 F100 W6"]:
        data[f"Qto_WallBaseQuantities:{prop_name}"] = get_property_value(psets, "Qto_WallBaseQuantities", prop_name)
    
    # Storey - получаем имя этажа
    data["Storey"] = None
    try:
        if hasattr(element, 'ContainedInStructure') and element.ContainedInStructure:
            storey = element.ContainedInStructure[0].RelatingStructure
            if storey and hasattr(storey, 'Name'):
                data["Storey"] = storey.Name
    except:
        pass
    
    # Type properties
    data["Type:GlobalId"] = None
    data["Type:Name"] = None
    try:
        if hasattr(element, 'IsTypedBy') and element.IsTypedBy:
            type_obj = element.IsTypedBy[0].RelatingType
            if type_obj:
                data["Type:GlobalId"] = getattr(type_obj, 'GlobalId', None)
                data["Type:Name"] = getattr(type_obj, 'Name', None)
        elif hasattr(element, 'ObjectType') and element.ObjectType:
            # Если нет типа, но есть ObjectType
            data["Type:Name"] = element.ObjectType
    except:
        pass
    
    return data


def classify_element(element, name, psets=None):
    """Классификация элемента на более детальные подтипы"""
    element_type = element.is_a()
    name_upper = name.upper() if name else ""
    
    # Получаем информацию о бетоне для проверки класса
    concrete_class = ""
    if psets:
        full_text = f"{name} "
        for pset_name, pset_data in psets.items():
            for prop_name, prop_val in pset_data.items():
                if isinstance(prop_val, str):
                    full_text += f"{prop_val} "
        class_match = re.search(r'[BВ]\s?(\d+(?:[.,]\d+)?)', full_text.upper())
        if class_match:
            concrete_class = class_match.group(1).replace(',', '.')
    
    if element_type == "IfcSlab":
        # Приоритет 1: Лестничные площадки (должны быть отдельно)
        # Проверяем по имени элемента и по типу
        if ("ЛЕСТНИЧ" in name_upper or "STAIR" in name_upper or 
            "ПЛОЩАДК" in name_upper or "ПЛОЩ." in name_upper or
            "ЛМ" in name_upper):
            return "Лестница_Площадка"
        
        # Приоритет 2: Подготовка фундамента (низкие классы бетона, аномальные классы или явные маркеры)
        # Классы бетона ниже В20 или аномальные (например В230) считаем подготовкой
        try:
            class_value = float(concrete_class) if concrete_class else 0
            # Считаем подготовкой: класс < 20 ИЛИ класс > 100 (аномалия типа В230)
            if class_value > 0 and (class_value < 20 or class_value > 100):
                return "Фундамент_Подготовка_Бетон"
        except ValueError:
            pass
        
        if "ПОДГОТОВКА" in name_upper or "СТЯЖКА" in name_upper:
            if "ПЕСОК" in name_upper or "SAND" in name_upper or "ЦЕМЕНТНО-ПЕСЧАНАЯ" in name_upper:
                return "Фундамент_Подготовка_Песок"
            elif "ЩЕБЕНЬ" in name_upper or "GRAVEL" in name_upper or "CRUSHED" in name_upper:
                return "Фундамент_Подготовка_Щебень"
            else:
                return "Фундамент_Подготовка_Бетон"
        
        # Приоритет 3: Фундаментные плиты (ищем явные маркеры фундамента)
        if ("ФУНДАМЕНТ" in name_upper and "ПОДГОТОВКА" not in name_upper) or \
           "FOUNDATION SLAB" in name_upper or \
           "BASE SLAB" in name_upper or \
           "ФУНДАМЕНТНАЯ ПЛИТА" in name_upper:
            return "Фундамент_Плита"
        
        if "ФП" in name_upper and "ПЕРЕКРЫТИЕ" not in name_upper and "ПОДГОТОВКА" not in name_upper:
            return "Фундамент_Плита"
        
        # Приоритет 4: Перекрытия (по умолчанию для плит)
        return "Перекрытие_Плита"
    
    if element_type == "IfcWall":
        return "Стена"
    if element_type == "IfcColumn":
        return "Колонна"
    if element_type in ["IfcStair", "IfcStairFlight"]:
        return "Лестница_Марш"
    if element_type == "IfcBeam":
        return "Балка"
        
    return element_type


def parse_concrete_properties(name, psets):
    """Парсинг свойств бетона из имени и PSets"""
    props = {
        "material": "Бетон",
        "class": "",
        "frost": "",
        "water": "",
        "reinforced": True,
        "type_detail": ""
    }
    
    full_text = f"{name} "
    for pset_name, pset_data in psets.items():
        for prop_name, prop_val in pset_data.items():
            if isinstance(prop_val, str):
                full_text += f"{prop_val} "
    
    full_text_upper = full_text.upper()
    
    # Проверка на пенополистирол/утеплитель
    if "ПЕНОПОЛИСТИРОЛ" in full_text_upper or "ЭКСТРУДИРОВАНН" in full_text_upper or "УТЕПЛИТЕЛ" in full_text_upper:
        props["material"] = "Экструдированный пенополистирол"
        props["type_detail"] = "Экструдированный пенополистирол"
        props["reinforced"] = False
        return props
    
    # Проверка на не бетонные материалы подготовки
    if "ПЕСОК" in full_text_upper or "SAND" in full_text_upper:
        props["material"] = "Песок"
        props["type_detail"] = "Песок"
        props["reinforced"] = False
    elif "ЩЕБЕНЬ" in full_text_upper or "GRAVEL" in full_text_upper or "CRUSHED STONE" in full_text_upper:
        props["material"] = "Щебень"
        props["type_detail"] = "Щебень"
        props["reinforced"] = False
    elif "ПОДГОТОВКА" in full_text_upper and ("В7.5" in full_text_upper or "B7.5" in full_text_upper or "НЕАРМИРОВАННЫЙ" in full_text_upper):
        props["type_detail"] = "Бетон неармированный (подготовка)"
        props["reinforced"] = False

    # Поиск класса бетона
    class_match = re.search(r'[BВ]\s?(\d+(?:[.,]\d+)?)', full_text_upper)
    if class_match:
        val = class_match.group(1).replace(',', '.')
        props["class"] = f"В{val}"
        if float(val) < 15:
             props["reinforced"] = False

    # Поиск морозостойкости
    frost_match = re.search(r'F\s?(\d+)', full_text_upper)
    if frost_match:
        props["frost"] = f"F{frost_match.group(1)}"
    
    # Поиск водонепроницаемости
    water_match = re.search(r'W\s?(\d+)', full_text_upper)
    if water_match:
        props["water"] = f"W{water_match.group(1)}"
        
    return props


def extract_ifc_data(ifc_path):
    ifc_file = ifcopenshell.open(ifc_path)
    unit_scale = get_unit_scale(ifc_file)
    
    elements_data = []
    all_elements = ifc_file.by_type("IfcProduct")
    
    print(f"Всего элементов в файле: {len(all_elements)}")
    
    for i, element in enumerate(all_elements):
        if not element.Representation:
            continue
            
        name = element.Name if element.Name else "Без имени"
        tag = element.Tag if element.Tag else ""
        element_type = element.is_a()
        
        psets = {}
        try:
            pset_definitions = ifcopenshell.util.element.get_psets(element)
            for pset_name, props in pset_definitions.items():
                psets[pset_name] = props
        except:
            pass
        
        # Извлекаем все параметры согласно REQUIRED_COLUMNS
        item = extract_element_properties(element, psets)
        
        # Добавляем вычисляемые значения объема и площади
        volume = 0.0
        area = 0.0
        shape = None
        
        # Приоритет №1: Объем из свойств
        found_volume_in_props = False
        for pset_name, props in psets.items():
            if "NetVolume" in props:
                vol = props["NetVolume"]
                if isinstance(vol, (int, float)):
                    volume = vol
                    found_volume_in_props = True
                    break
            if "GrossVolume" in props and not found_volume_in_props:
                vol = props["GrossVolume"]
                if isinstance(vol, (int, float)):
                    if vol > 10000: 
                        volume = vol / 1000.0
                    else:
                        volume = vol
                    found_volume_in_props = True
                    break
        
        # Если объема нет в свойствах, пробуем рассчитать по геометрии
        if not found_volume_in_props:
            try:
                settings = ifcopenshell.geom.settings()
                settings.set(settings.USE_PYTHON_OPENCASCADE, True)
                shape = ifcopenshell.geom.create_shape(settings, element)
                if shape:
                    volume = calculate_volume_by_geometry(shape, unit_scale)
            except Exception as e:
                pass
        
        # Расчет площади
        try:
            if shape:
                area = calculate_area_by_geometry(shape, unit_scale)
            else:
                settings = ifcopenshell.geom.settings()
                settings.set(settings.USE_PYTHON_OPENCASCADE, True)
                shape = ifcopenshell.geom.create_shape(settings, element)
                if shape:
                    area = calculate_area_by_geometry(shape, unit_scale)
        except:
            pass
        
        # Добавляем volume и area в item
        item["volume_m3"] = round(volume, 4)
        item["area_m2"] = round(area, 4)
        
        # Классификация и материал
        category = classify_element(element, name, psets)
        material_props = parse_concrete_properties(name, psets)
        
        item["category"] = category
        item["material"] = material_props["material"]
        item["concrete_class"] = material_props["class"]
        item["frost_resistance"] = material_props["frost"]
        item["water_permeability"] = material_props["water"]
        item["is_reinforced"] = material_props["reinforced"]
        item["material_detail"] = material_props["type_detail"]
        
        elements_data.append(item)
        
        if (i + 1) % 500 == 0:
            print(f"Обработано {i + 1} элементов...")
    
    return elements_data


def get_russian_type_name(category, ifc_type):
    """Получение русского названия типа элемента"""
    type_mapping = {
        "Балка": "Балки",
        "Колонна": "Колонны",
        "Стена": "Стены",
        "Перекрытие_Плита": "Перекрытия",
        "Фундамент_Плита": "Фундаментные_плиты",
        "Фундамент_Подготовка_Бетон": "Подготовка_фундамента",
        "Фундамент_Подготовка_Песок": "Подготовка_фундамента",
        "Фундамент_Подготовка_Щебень": "Подготовка_фундамента",
        "Лестница_Марш": "Лестничные_марши",
        "Лестница_Площадка": "Лестничные_площадки",
        "Пандус": "Пандусы",
    }
    
    # Сначала пробуем по категории
    if category in type_mapping:
        return type_mapping[category]
    
    # Затем по IFC типу
    ifc_type_mapping = {
        "IfcBeam": "Балки",
        "IfcColumn": "Колонны",
        "IfcWall": "Стены",
        "IfcSlab": "Перекрытия",
        "IfcStairFlight": "Лестничные_марши",
        "IfcStair": "Лестницы",
        "IfcRamp": "Пандусы",
        "IfcFooting": "Фундаменты",
        "IfcPlate": "Плиты",
        "IfcMember": "Элементы",
    }
    
    if ifc_type in ifc_type_mapping:
        return ifc_type_mapping[ifc_type]
    
    # Если не найдено, возвращаем категорию
    return category if category else ifc_type


def format_material_string(material_props):
    """Форматирование строки материала с учетом класса и свойств"""
    material = material_props.get("material", "")
    concrete_class = material_props.get("class", "")
    frost = material_props.get("frost", "")
    water = material_props.get("water", "")
    type_detail = material_props.get("type_detail", "")
    
    # Если есть type_detail, используем его
    if type_detail:
        return type_detail
    
    # Собираем полное описание материала
    parts = [material] if material else []
    
    if concrete_class:
        parts.append(concrete_class)
    
    if frost:
        parts.append(frost)
    
    if water:
        parts.append(water)
    
    # Если материал пустой, но есть свойства бетона
    if not parts and (concrete_class or frost or water):
        props = []
        if concrete_class:
            props.append(concrete_class)
        if frost:
            props.append(frost)
        if water:
            props.append(water)
        return " ".join(props)
    
    return " ".join(parts) if parts else "-"


def create_summary_table(data):
    """Создание сводной таблицы в требуемом формате"""
    df = pd.DataFrame(data)
    
    # Список IFC классов для исключения (неинформативные элементы)
    excluded_ifc_classes = ["IfcDiscreteAccessory", "IfcElementAssembly", "IfcOpeningElement"]
    
    # Фильтрация по IFC классам
    df = df[~df["Ifc Class"].isin(excluded_ifc_classes)]
    
    # Добавляем колонку с русским названием типа
    df["Тип (RU)"] = df.apply(
        lambda row: get_russian_type_name(row.get("category", ""), row.get("Ifc Class", "")), 
        axis=1
    )
    
    # Добавляем колонку с полным описанием материала
    df["Материал_полный"] = df.apply(
        lambda row: format_material_string({
            "material": row.get("material", ""),
            "class": row.get("concrete_class", ""),
            "frost": row.get("frost_resistance", ""),
            "water": row.get("water_permeability", ""),
            "type_detail": row.get("material_detail", "")
        }),
        axis=1
    )
    
    # Фильтрация строк с неинформативными материалами
    # Оставляем только если есть конкретный класс бетона (В25, В30 и т.д.) или detail-описание
    # ИСКЛЮЧЕНИЕ: Лестничные марши (IfcStairFlight) всегда включаем
    def is_informative_material(row):
        material_full = row["Материал_полный"]
        concrete_class = row.get("concrete_class", "")
        material_detail = row.get("material_detail", "")
        ifc_class = row.get("Ifc Class", "")
        
        # ИСКЛЮЧЕНИЕ: Лестничные марши всегда включаем в таблицу
        if ifc_class == "IfcStairFlight":
            return True
        
        # Если есть detail-описание - это информативно
        if material_detail:
            return True
        
        # Если есть класс бетона (В25, В30, В7.5 и т.п.) - это информативно
        if concrete_class and re.search(r'В\d+\.?\d*', concrete_class):
            return True
        
        # Если материал содержит конкретное название (не просто "Бетон")
        if material_full and material_full != "-" and len(material_full.split()) > 1:
            return True
            
        return False
    
    df = df[df.apply(is_informative_material, axis=1)]
    
    # Группировка по типу, IFC классу и материалу
    group_cols = ["Тип (RU)", "Ifc Class", "Материал_полный"]
    
    summary = df.groupby(group_cols, dropna=False).agg(
        count=("id" if "id" in df.columns else "Element Specific:GlobalId", "count"),
        total_volume_m3=("volume_m3", "sum"),
        total_area_m2=("area_m2", "sum")
    ).reset_index()
    
    # Переименовываем колонки в соответствии с требуемым форматом
    summary = summary.rename(columns={
        "Ifc Class": "Тип элемента",
        "Материал_полный": "Материал",
        "count": "Количество, шт",
        "total_volume_m3": "Объем, м³"
    })
    
    # Форматирование числовых значений
    summary["Объем, м³"] = summary["Объем, м³"].apply(lambda x: round(x, 3) if pd.notna(x) and x != 0 else None)
    
    # Замена None на "-" для отображения
    summary["Объем, м³"] = summary["Объем, м³"].fillna("-")
    
    # Сортировка
    summary = summary.sort_values(by=["Тип (RU)", "Тип элемента", "Материал"], ascending=[True, True, True])
    
    # Выбираем только нужные колонки в правильном порядке (без Площади)
    summary = summary[["Тип (RU)", "Тип элемента", "Материал", "Количество, шт", "Объем, м³"]]
    
    return summary


def main():
    if not os.path.exists(IFC_FILE):
        print(f"Ошибка: Файл {IFC_FILE} не найден!")
        return

    print("Извлечение данных из IFC...")
    data = extract_ifc_data(IFC_FILE)
    
    with open(JSON_OUTPUT, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    print(f"Подробные данные сохранены в {JSON_OUTPUT}")
    
    print("Формирование сводной таблицы...")
    summary_df = create_summary_table(data)
    
    os.makedirs(os.path.dirname(EXCEL_OUTPUT), exist_ok=True)
    summary_df.to_excel(EXCEL_OUTPUT, index=False, sheet_name="Сводная")
    print(f"Сводная таблица сохранена в {EXCEL_OUTPUT}")

    print("\n--- ИТОГИ ПО КАТЕГОРИЯМ ---")
    totals = summary_df.groupby("Тип (RU)").agg(
        total_count=("Количество, шт", "sum"),
        total_vol=("Объем, м³", lambda x: x.apply(lambda v: 0 if v == "-" else v).sum())
    )
    print(totals.to_string())
    
    targets = {
        "Стены": 3857.89,
        "Перекрытия": 4294.95,
        "Колонны": 94.28,
        "Фундаментные_плиты": 3042.14,
        "Подготовка_фундамента": 619.08,
        "Лестничные_марши": 0.0
    }
    
    print("\n--- СРАВНЕНИЕ С ЦЕЛЕВЫМИ ЗНАЧЕНИЯМИ ---")
    for cat, target_vol in targets.items():
        if cat in totals.index:
            actual_vol = totals.loc[cat, "total_vol"]
            diff = actual_vol - target_vol
            percent = (actual_vol / target_vol) * 100 if target_vol != 0 else 0
            print(f"{cat}: Факт={actual_vol:.2f}, Цель={target_vol:.2f}, Разница={diff:.2f} ({percent:.1f}%)")
        else:
            found_similar = [k for k in totals.index if cat.split('_')[0] in k]
            if found_similar:
                 print(f"{cat}: Не найдено точное совпадение. Найдено похожие: {found_similar}. Цель={target_vol:.2f}")
            else:
                print(f"{cat}: Не найдено в модели! Цель={target_vol:.2f}")


if __name__ == "__main__":
    main()
