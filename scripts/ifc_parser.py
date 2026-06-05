#!/usr/bin/env python3
"""
IFC Extract - Единый скрипт для извлечения данных из IFC и создания сводной таблицы материалов.
Работает внутри папки сессии.

Очередность операций:
1. extract_ifc_data() - извлекает все данные из IFC файла
2. create_summary_table() - создает сводную таблицу по типам элементов и материалам
3. Сохраняет только результирующий файл materials_summary.xlsx в папке сессии

Формат выходной таблицы:
Тип (RU) | Тип элемента | Материал (с характеристиками: Бетон В30 F150 W6) | Количество, шт | Объем, м³
"""

import ifcopenshell
import ifcopenshell.util.element
import ifcopenshell.util.unit
import json
import os
import math
import re
from collections import defaultdict
from pathlib import Path
from typing import Dict, List, Any, Optional, Tuple
import pandas as pd
import numpy as np
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side


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
        area = shape.geometry.surface_area
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
    data["ExpCheck_StairFlight:MGE_ElementCode"] = get_property_value(psets, "ExpCheck_StairFlight", "MGE_ElementCode")
    data["ExpCheck_StairFlight:MGE_Gost"] = get_property_value(psets, "ExpCheck_StairFlight", "MGE_Gost")
    data["ExpCheck_StairFlight:MGE_Material"] = get_property_value(psets, "ExpCheck_StairFlight", "MGE_Material")
    data["ExpCheck_StairFlight:MGE_MaterialCode"] = get_property_value(psets, "ExpCheck_StairFlight", "MGE_MaterialCode")
    data["ExpCheck_StairFlight:MGE_Name"] = get_property_value(psets, "ExpCheck_StairFlight", "MGE_Name")
    data["ExpCheck_StairFlight:MGE_Position"] = get_property_value(psets, "ExpCheck_StairFlight", "MGE_Position")
    data["ExpCheck_StairFlight:MGE_Section"] = get_property_value(psets, "ExpCheck_StairFlight", "MGE_Section")
    
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
    data["Pset_StairFlightCommon:IsExternal"] = get_property_value(psets, "Pset_StairFlightCommon", "IsExternal")
    data["Pset_StairFlightCommon:LoadBearing"] = get_property_value(psets, "Pset_StairFlightCommon", "LoadBearing")
    data["Pset_StairFlightCommon:Reference"] = get_property_value(psets, "Pset_StairFlightCommon", "Reference")
    
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
    
    # Qto_SlabBaseQuantities
    data["Qto_SlabBaseQuantities:GrossArea"] = get_property_value(psets, "Qto_SlabBaseQuantities", "GrossArea")
    data["Qto_SlabBaseQuantities:GrossVolume"] = get_property_value(psets, "Qto_SlabBaseQuantities", "GrossVolume")
    data["Qto_SlabBaseQuantities:NetArea"] = get_property_value(psets, "Qto_SlabBaseQuantities", "NetArea")
    data["Qto_SlabBaseQuantities:NetVolume"] = get_property_value(psets, "Qto_SlabBaseQuantities", "NetVolume")
    data["Qto_SlabBaseQuantities:Perimeter"] = get_property_value(psets, "Qto_SlabBaseQuantities", "Perimeter")
    data["Qto_SlabBaseQuantities:Width"] = get_property_value(psets, "Qto_SlabBaseQuantities", "Width")
    
    # Qto_StairFlightBaseQuantities
    data["Qto_StairFlightBaseQuantities:GrossVolume"] = get_property_value(psets, "Qto_StairFlightBaseQuantities", "GrossVolume")
    data["Qto_StairFlightBaseQuantities:Length"] = get_property_value(psets, "Qto_StairFlightBaseQuantities", "Length")
    data["Qto_StairFlightBaseQuantities:NetVolume"] = get_property_value(psets, "Qto_StairFlightBaseQuantities", "NetVolume")
    
    # Qto_WallBaseQuantities
    data["Qto_WallBaseQuantities:GrossSideArea"] = get_property_value(psets, "Qto_WallBaseQuantities", "GrossSideArea")
    data["Qto_WallBaseQuantities:GrossVolume"] = get_property_value(psets, "Qto_WallBaseQuantities", "GrossVolume")
    data["Qto_WallBaseQuantities:Height"] = get_property_value(psets, "Qto_WallBaseQuantities", "Height")
    data["Qto_WallBaseQuantities:Length"] = get_property_value(psets, "Qto_WallBaseQuantities", "Length")
    data["Qto_WallBaseQuantities:NetSideArea"] = get_property_value(psets, "Qto_WallBaseQuantities", "NetSideArea")
    data["Qto_WallBaseQuantities:NetVolume"] = get_property_value(psets, "Qto_WallBaseQuantities", "NetVolume")
    data["Qto_WallBaseQuantities:Width"] = get_property_value(psets, "Qto_WallBaseQuantities", "Width")
    
    # Storey
    data["Storey"] = None
    
    # Type
    data["Type:GlobalId"] = None
    data["Type:Name"] = None
    
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
    """Основная функция извлечения данных из IFC"""
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
        "IfcFooting": "Фундаментные_плиты",
        "IfcStairFlight": "Лестничные_марши",
        "IfcStair": "Лестницы",
        "IfcRamp": "Пандусы",
        "IfcPlate": "Плиты",
        "IfcMember": "Элементы_каркаса",
        "IfcFurnishingElement": "Мебель",
        "IfcBuildingElementProxy": "Прочие_элементы",
        "IfcOpeningElement": "Проемы",
        "IfcCurtainWall": "Навесные_стены",
        "IfcWindow": "Окна",
        "IfcDoor": "Двери",
        "IfcRoof": "Крыши",
        "IfcPile": "Сваи",
        "IfcRailing": "Ограждения",
        "IfcCovering": "Покрытия",
        "IfcReinforcingMesh": "Арматурные_сетки",
        "IfcBuildingStorey": "Этажи",
        "IfcBuilding": "Здание",
    }

    if ifc_type in ifc_type_mapping:
        return ifc_type_mapping[ifc_type]

    return "Прочие"


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
    excluded_ifc_classes = [
        "IfcDiscreteAccessory",
        "IfcElementAssembly",
        "IfcOpeningElement",
        "IfcBuildingElementProxy"
    ]

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
    def is_informative_material(row):
        material_full = row["Материал_полный"]
        concrete_class = row.get("concrete_class", "")
        material_detail = row.get("material_detail", "")
        ifc_class = row.get("Ifc Class", "")

        # ИСКЛЮЧЕНИЕ: Лестницы и лестничные марши всегда включаем в таблицу
        if ifc_class in ["IfcStair", "IfcStairFlight"]:
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


def create_summary_excel(items, output_path):
    """Создание Excel файла со сводной таблицей"""
    wb = Workbook()
    ws = wb.active
    ws.title = "Материалы"
    
    # Заголовки
    headers = ["Тип (RU)", "Тип элемента", "Материал", "Количество, шт", "Объем, м³"]
    
    # Стили
    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="2c3e50", end_color="2c3e50", fill_type="solid")
    header_alignment = Alignment(horizontal="center", vertical="center")
    thin_border = Border(
        left=Side(style='thin'), right=Side(style='thin'),
        top=Side(style='thin'), bottom=Side(style='thin')
    )
    
    # Записываем заголовки
    for col_idx, header in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col_idx, value=header)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = header_alignment
        cell.border = thin_border
    
    # Записываем данные
    for row_idx, item in enumerate(items, 2):
        ws.cell(row=row_idx, column=1, value=item.get("Тип (RU)", "")).border = thin_border
        ws.cell(row=row_idx, column=2, value=item.get("Тип элемента", "")).border = thin_border
        ws.cell(row=row_idx, column=3, value=item.get("Материал", "")).border = thin_border
        ws.cell(row=row_idx, column=4, value=item.get("Количество, шт", 0)).border = thin_border
        
        vol = item.get("Объем, м³", "-")
        ws.cell(row=row_idx, column=5, value=vol).border = thin_border
    
    # Авто-ширина колонок
    for col in ws.columns:
        max_len = 0
        col_letter = col[0].column_letter
        for cell in col:
            if cell.value:
                try:
                    max_len = max(max_len, len(str(cell.value)))
                except:
                    pass
        ws.column_dimensions[col_letter].width = min(max_len + 2, 50)
    
    wb.save(output_path)
    print(f"✅ Сводная таблица сохранена: {output_path}")


def extract_building_info(ifc_file):
    """
    Извлечь информацию об объекте строительства из открытого IFC файла.
    
    Args:
        ifc_file: Открытый IFC файл
        
    Returns:
        Словарь с информацией о здании
    """
    result = {
        "project": {},
        "building": {},
        "storeys": [],
        "summary": {}
    }
    
    # === Извлечение данных проекта ===
    projects = ifc_file.by_type("IfcProject")
    if projects:
        project = projects[0]
        result["project"]["name"] = getattr(project, 'Name', None)
        result["project"]["long_name"] = getattr(project, 'LongName', None)
        result["project"]["description"] = getattr(project, 'Description', None)
    
    # === Извлечение данных здания ===
    buildings = ifc_file.by_type("IfcBuilding")
    if buildings:
        building = buildings[0]
        
        # Базовые атрибуты
        result["building"]["name"] = getattr(building, 'Name', None)
        result["building"]["long_name"] = getattr(building, 'LongName', None)
        result["building"]["elevation_ref_height"] = getattr(building, 'ElevationOfRefHeight', None)
        result["building"]["elevation_terrain"] = getattr(building, 'ElevationOfTerrain', None)
        
        # Извлечение свойств через IsDefinedBy
        psets_data = {}
        if hasattr(building, 'IsDefinedBy'):
            for rel in building.IsDefinedBy:
                if hasattr(rel, 'RelatingPropertyDefinition'):
                    pset_def = rel.RelatingPropertyDefinition
                    pset_name = getattr(pset_def, 'Name', 'Unknown')
                    
                    if hasattr(pset_def, 'HasProperties'):
                        for prop in pset_def.HasProperties:
                            prop_name = getattr(prop, 'Name', None)
                            prop_value = getattr(prop, 'NominalValue', None)
                            
                            if prop_name and prop_value is not None:
                                # Извлекаем само значение из NominalValue
                                if hasattr(prop_value, 'wrappedValue'):
                                    actual_value = prop_value.wrappedValue
                                else:
                                    actual_value = prop_value
                                
                                if pset_name not in psets_data:
                                    psets_data[pset_name] = {}
                                psets_data[pset_name][prop_name] = actual_value
        
        # Сохраняем все найденные свойства
        result["building"]["properties"] = psets_data
        
        # Ключевые параметры для summary
        if "ExpCheck_Building" in psets_data:
            exp_props = psets_data["ExpCheck_Building"]
            result["building"]["address"] = exp_props.get("MGE_BuildingAddress")
            result["building"]["customer"] = exp_props.get("MGE_Customer")
            result["building"]["designer"] = exp_props.get("MGE_Designer")
            result["building"]["project_name"] = exp_props.get("MGE_ProjectName")
            result["building"]["object_name"] = exp_props.get("MGE_ObjectName")
            result["building"]["project_code"] = exp_props.get("MGE_ProjectCode")
            result["building"]["korpus"] = exp_props.get("MGE_Korpus")
            result["building"]["section"] = exp_props.get("MGE_Section")
            result["building"]["num_of_section"] = exp_props.get("MGE_NumOfSection")
            result["building"]["functional_use"] = exp_props.get("MGE_FunctionalUse")
            
            # Высота от нулевой отметки (разница между референсной высотой и уровнем земли)
            ref_height = exp_props.get("MGE_ElevationOfRefHeight")
            terrain_height = exp_props.get("MGE_ElevationOfTerrain")
            if ref_height is not None and terrain_height is not None:
                # Обычно MGE_ElevationOfRefHeight - это абсолютная высота в мм
                # Конвертируем в метры
                result["building"]["ref_height_m"] = float(ref_height) / 1000.0 if ref_height else None
                result["building"]["terrain_height_m"] = float(terrain_height) / 1000.0 if terrain_height else None
        
        if "Pset_BuildingCommon" in psets_data:
            common_props = psets_data["Pset_BuildingCommon"]
            result["building"]["number_of_storeys"] = common_props.get("NumberOfStoreys")
            result["building"]["construction_method"] = common_props.get("ConstructionMethod")
            result["building"]["fire_protection_class"] = common_props.get("FireProtectionClass")
            result["building"]["is_landmarked"] = common_props.get("IsLandmarked")
    
    # === Извлечение информации об этажах ===
    storeys = ifc_file.by_type("IfcBuildingStorey")
    above_ground_storeys = []
    below_ground_storeys = []
    
    for storey in storeys:
        elevation = getattr(storey, 'Elevation', None)
        storey_info = {
            "name": getattr(storey, 'Name', None),
            "elevation": float(elevation) / 1000.0 if elevation is not None else None  # Конвертация в метры
        }
        
        # Определяем, надземный или подземный этаж
        if elevation is not None:
            if elevation > 0:
                above_ground_storeys.append(storey_info)
            else:
                below_ground_storeys.append(storey_info)
        else:
            above_ground_storeys.append(storey_info)
        
        result["storeys"].append(storey_info)
    
    # Сортировка этажей по высоте
    result["storeys"].sort(key=lambda x: x["elevation"] if x["elevation"] is not None else 0)
    
    # === Формирование сводной информации (summary) ===
    total_storeys = len(above_ground_storeys) + len(below_ground_storeys)
    
    # Расчет высоты здания
    max_elevation = None
    min_elevation = None
    
    for storey in result["storeys"]:
        elev = storey["elevation"]
        if elev is not None:
            if max_elevation is None or elev > max_elevation:
                max_elevation = elev
            if min_elevation is None or elev < min_elevation:
                min_elevation = elev
    
    # Общая высота здания от нулевой отметки
    if max_elevation is not None:
        result["summary"]["total_height_from_zero_m"] = round(max_elevation, 3)
    
    # Глубина подземной части
    if min_elevation is not None and min_elevation < 0:
        result["summary"]["underground_depth_m"] = round(abs(min_elevation), 3)
    
    # Полная высота (от низа подземной части до верха)
    if max_elevation is not None and min_elevation is not None:
        result["summary"]["total_height_full_m"] = round(max_elevation - min_elevation, 3)
    
    # Этажность
    result["summary"]["above_ground_storeys"] = len(above_ground_storeys)
    result["summary"]["below_ground_storeys"] = len(below_ground_storeys)
    result["summary"]["total_storeys"] = total_storeys
    
    # Если есть данные из Pset_BuildingCommon
    if result["building"].get("number_of_storeys"):
        result["summary"]["storeys_from_pset"] = result["building"]["number_of_storeys"]
    
    # Адрес
    result["summary"]["address"] = result["building"].get("address")
    
    return result


def parse_ifc_file(ifc_path, output_folder):
    """
    Основная функция для сервиса ifc_inspect.
    Работает внутри папки сессии.
    
    Очередность:
    1. Извлекаем данные из IFC (extract_ifc_data)
    2. Создаем сводную таблицу (create_summary_table)
    3. Сохраняем только materials_summary.xlsx
    
    Args:
        ifc_path: Path to IFC file
        output_folder: Folder to save results (session folder)
        
    Returns:
        Dict with parsing results
    """
    print(f"\n{'='*60}")
    print("📊 ОБРАБОТКА IFC МОДЕЛИ")
    print(f"{'='*60}")
    print(f"IFC файл: {os.path.basename(ifc_path)}")
    print(f"Папка сессии: {output_folder}")
    
    # Шаг 1: Извлечение данных из IFC
    print("\n📋 Шаг 1: Извлечение данных из IFC...")
    data = extract_ifc_data(ifc_path)
    print(f"   ✅ Извлечено {len(data)} элементов")
    
    # Шаг 1.5: Извлечение информации о здании
    print("\n🏢 Шаг 1.5: Извлечение информации об объекте строительства...")
    try:
        ifc_file_for_building = ifcopenshell.open(ifc_path)
        building_info = extract_building_info(ifc_file_for_building)
        
        # Сохраняем building_info.json в папке сессии
        building_json_path = Path(output_folder) / "building_info.json"
        with open(building_json_path, 'w', encoding='utf-8') as f:
            json.dump({
                'success': True,
                'source_file': os.path.basename(ifc_path),
                'project': building_info.get('project', {}),
                'building': building_info.get('building', {}),
                'storeys': building_info.get('storeys', []),
                'summary': building_info.get('summary', {})
            }, f, ensure_ascii=False, indent=2)
        
        summary = building_info.get('summary', {})
        print(f"   ✅ Информация об объекте:")
        if summary.get('address'):
            print(f"      Адрес: {summary['address']}")
        if summary.get('total_height_from_zero_m'):
            print(f"      Высота от нулевой отметки: {summary['total_height_from_zero_m']} м")
        if summary.get('above_ground_storeys') is not None:
            print(f"      Этажность: {summary['above_ground_storeys']} надземных + {summary.get('below_ground_storeys', 0)} подземных")
        print(f"   ✅ building_info.json сохранен")
    except Exception as e:
        print(f"   ⚠️ Не удалось извлечь информацию об объекте: {e}")
        building_info = None
    
    # Шаг 2: Создание сводной таблицы
    print("\n📊 Шаг 2: Создание сводной таблицы...")
    summary_df = create_summary_table(data)
    items = summary_df.to_dict('records')
    print(f"   ✅ Сформировано {len(items)} строк в сводной таблице")
    
    # Шаг 3: Сохранение результата
    output_path = Path(output_folder) / "materials_summary.xlsx"
    create_summary_excel(items, str(output_path))
    
    # Сохраняем JSON для машинной обработки
    json_path = Path(output_folder) / "materials_summary.json"
    with open(json_path, 'w', encoding='utf-8') as f:
        json.dump({
            'success': True,
            'source_file': os.path.basename(ifc_path),
            'total_items': len(items),
            'output_excel': 'materials_summary.xlsx',
            'items': items
        }, f, ensure_ascii=False, indent=2)
    
    # Подсчет статистики
    total_count = sum(item.get("Количество, шт", 0) for item in items if isinstance(item.get("Количество, шт"), (int, float)))
    total_volume = sum(float(item.get("Объем, м³", 0)) for item in items if item.get("Объем, м³") != "-" and isinstance(item.get("Объем, м³"), (int, float)))
    
    result = {
        'success': True,
        'excel_path': str(output_path),
        'excel_filename': 'materials_summary.xlsx',
        'json_path': str(json_path),
        'total_elements': len(data),
        'aggregated_materials': len(items),
        'total_count': total_count,
        'total_volume': round(total_volume, 3)
    }
    
    print(f"\n💾 Результаты сохранены в: {output_folder}")
    print(f"   • materials_summary.xlsx - сводная таблица")
    print(f"   • materials_summary.json - данные для обработки")
    if building_info:
        print(f"   • building_info.json - информация об объекте строительства")
    print(f"\n✅ Обработка IFC завершена. Всего элементов: {len(data)}, Уникальных материалов: {len(items)}")
    
    return result


# Entry point when run directly
if __name__ == "__main__":
    import sys
    
    if len(sys.argv) < 2:
        print("Использование: python ifc_parser.py <path_to_IFC> [output_folder]")
        print("Пример: python ifc_parser.py /workspace/data/model.ifc /workspace/uploads/session_id")
        sys.exit(1)
    
    ifc_file = sys.argv[1]
    output_folder = sys.argv[2] if len(sys.argv) > 2 else "."
    
    parse_ifc_file(ifc_file, output_folder)
