import ifcopenshell
import ifcopenshell.util.element

f = ifcopenshell.open('data/ifc модель КР.ifc')
stairs = [e for e in f.by_type('IfcStair')][:3]

for e in stairs:
    print(f'Name: {e.Name}')
    psets = ifcopenshell.util.element.get_psets(e)
    qto = ifcopenshell.util.element.get_qtos(e)
    print(f'  QTO: {qto}')
    print()
