# -*- coding: utf-8 -*-
__title__ = "MEP Wall Openings"

from pyrevit import revit, forms, script
import clr

clr.AddReference("System")
from System.Collections.Generic import List as NetList

clr.AddReference("RevitAPI")
from Autodesk.Revit.DB import *
from Autodesk.Revit.DB.Plumbing import Pipe
from Autodesk.Revit.DB.Mechanical import Duct

doc = revit.doc


def get_opening_dimensions(element, offset_ft):
    """Zwraca szerokość i wysokość (lub średnicę) elementu z zapasem."""
    # Domyślne wartości
    w = 0
    h = 0

    # Obsługa rur i kanałów okrągłych
    if hasattr(element, "Diameter"):
        dim = element.Diameter + (offset_ft * 2)
        w = h = dim
    # Obsługa kanałów prostokątnych
    elif isinstance(element, Duct):
        try:
            w = element.Width + (offset_ft * 2)
            h = element.Height + (offset_ft * 2)
        except:
            # Fallback dla kanałów owalnych/okrągłych w kategorii Duct
            dim = element.Diameter + (offset_ft * 2)
            w = h = dim

    return w, h


def create_wall_opening(wall, pt, w, h):
    """Tworzy otwór systemowy o zadanych wymiarach."""
    wall_curve = wall.Location.Curve
    wall_dir = (wall_curve.GetEndPoint(1) - wall_curve.GetEndPoint(0)).Normalize()

    # Punkty skrajne oparte na szerokości (wzdłuż ściany) i wysokości (Z)
    half_w = w / 2.0
    half_h = h / 2.0

    p1 = pt - (wall_dir * half_w) - (XYZ.BasisZ * half_h)
    p2 = pt + (wall_dir * half_w) + (XYZ.BasisZ * half_h)

    return doc.Create.NewOpening(wall, p1, p2)


# --- START ---
# Wybieramy rury ORAZ kanały
selection = revit.get_selection()
mep_elements = [el for el in selection if isinstance(el, Pipe) or isinstance(el, Duct)]

if not mep_elements:
    forms.alert("Zaznacz rury lub kanały wentylacyjne.")
    script.exit()

with revit.Transaction("Otwory dla rur i kanałów"):
    view_3d = next(
        (v for v in FilteredElementCollector(doc).OfClass(View3D) if not v.IsTemplate),
        None,
    )
    wall_filter = ElementMulticategoryFilter(
        NetList[BuiltInCategory]([BuiltInCategory.OST_Walls])
    )
    intersector = ReferenceIntersector(wall_filter, FindReferenceTarget.Face, view_3d)

    offset_ft = 50.0 / 304.8  # Margines 50mm
    success_count = 0

    for el in mep_elements:
        line = el.Location.Curve
        start = line.GetEndPoint(0)
        direction = (line.GetEndPoint(1) - start).Normalize()

        all_hits = intersector.Find(start, direction)
        processed_walls = set()

        # Pobierz wymiary otworu dla tego konkretnego elementu
        width, height = get_opening_dimensions(el, offset_ft)

        for hit in all_hits:
            if hit.Proximity > line.Length:
                continue

            wall_id = hit.GetReference().ElementId
            if wall_id in processed_walls:
                continue

            wall = doc.GetElement(wall_id)
            point = hit.GetReference().GlobalPoint

            try:
                opening = create_wall_opening(wall, point, width, height)
                if opening:
                    success_count += 1
                    processed_walls.add(wall_id)
                    category_name = "Rura" if isinstance(el, Pipe) else "Kanał"
                    print(
                        "Wycięto: {} {} -> Ściana {}".format(
                            category_name, el.Id, wall.Id
                        )
                    )
            except:
                pass

print("\nGotowe. Stworzono {} unikalnych otworów MEP.".format(success_count))
