# -*- coding: utf-8 -*-
__title__ = "Manual Point Router"

from pyrevit import revit, forms, script
import clr
import heapq

# Importy systemowe i Revit API
clr.AddReference("System")
from System.Collections.Generic import List as NetList

clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")
from Autodesk.Revit.DB import *
from Autodesk.Revit.DB.Plumbing import Pipe, PipeType, PipingSystemType
from Autodesk.Revit.UI.Selection import ObjectSnapTypes

doc = revit.doc
uidoc = revit.uidoc


def is_point_safe(pt, intersector, margin_ft):
    for d in [XYZ.BasisX, -XYZ.BasisX, XYZ.BasisY, -XYZ.BasisY]:
        hit = intersector.FindNearest(pt, d)
        if hit and hit.Proximity < margin_ft:
            return False
    return True


def find_path_points(start_pt, end_pt, intersector, margin_ft):
    step = 400.0 / 304.8
    turn_penalty = 15.0

    start_t = (round(start_pt.X, 1), round(start_pt.Y, 1), round(start_pt.Z, 1))
    open_set = [(0, start_t, (0, 0))]

    came_from = {}
    g_score = {(start_t, (0, 0)): 0}

    while open_set:
        _, curr_t, last_dir = heapq.heappop(open_set)
        curr = XYZ(curr_t[0], curr_t[1], curr_t[2])

        if curr.DistanceTo(end_pt) < step * 1.5:
            path = [end_pt]
            temp_t, temp_d = curr_t, last_dir
            while (temp_t, temp_d) in came_from:
                path.append(XYZ(temp_t[0], temp_t[1], temp_t[2]))
                temp_t, temp_d = came_from[(temp_t, temp_d)]
            path.append(start_pt)
            return path[::-1]

        for dx, dy in [(step, 0), (-step, 0), (0, step), (0, -step)]:
            neighbor = XYZ(curr.X + dx, curr.Y + dy, curr.Z)
            neighbor_t = (
                round(neighbor.X, 1),
                round(neighbor.Y, 1),
                round(neighbor.Z, 1),
            )
            curr_dir = (round(dx / step), round(dy / step))

            if is_point_safe(neighbor, intersector, margin_ft):
                cost = step
                if last_dir != (0, 0) and last_dir != curr_dir:
                    cost += turn_penalty

                tentative_g = g_score[(curr_t, last_dir)] + cost
                state = (neighbor_t, curr_dir)

                if state not in g_score or tentative_g < g_score[state]:
                    came_from[state] = (curr_t, last_dir)
                    g_score[state] = tentative_g
                    f_score = tentative_g + neighbor.DistanceTo(end_pt)
                    heapq.heappush(open_set, (f_score, neighbor_t, curr_dir))
    return None


# --- UI & SELECTION ---
try:
    # Ustawienie snapowania do punktów końcowych ułatwi celowanie w rury
    pt1 = uidoc.Selection.PickPoint(ObjectSnapTypes.Endpoints, "Kliknij START trasy")
    pt2 = uidoc.Selection.PickPoint(ObjectSnapTypes.Endpoints, "Kliknij KONIEC trasy")
except:
    script.exit()

# Automatyczne pobieranie parametrów rur z projektu
pipe_types = FilteredElementCollector(doc).OfClass(PipeType).ToElements()
sys_types = FilteredElementCollector(doc).OfClass(PipingSystemType).ToElements()

if not pipe_types or not sys_types:
    forms.alert("Nie znaleziono typów rur w projekcie.")
    script.exit()

p_type_id = pipe_types[0].Id
s_type_id = sys_types[0].Id
level_id = (
    doc.ActiveView.GenLevel.Id
    if doc.ActiveView.GenLevel
    else FilteredElementCollector(doc).OfClass(Level).FirstElementId()
)

with revit.Transaction("Ręczne Trasowanie "):
    # 3D View dla kolizji
    view_3d = next(
        (v for v in FilteredElementCollector(doc).OfClass(View3D) if not v.IsTemplate),
        None,
    )
    if not view_3d:
        forms.alert("Wymagany widok 3D (nie-szablon) do wykrywania ścian.")
        script.exit()

    wall_filter = ElementMulticategoryFilter(
        NetList[BuiltInCategory]([BuiltInCategory.OST_Walls])
    )
    intersector = ReferenceIntersector(wall_filter, FindReferenceTarget.Face, view_3d)

    # Margines bezpieczeństwa (ok. 30cm)
    margin = 320.0 / 304.8

    path = find_path_points(pt1, pt2, intersector, margin)

    if path:
        # Wygładzanie (usuwanie zbędnych węzłów)
        nodes = [path[0]]
        for i in range(1, len(path) - 1):
            v1 = (path[i] - path[i - 1]).Normalize()
            v2 = (path[i + 1] - path[i]).Normalize()
            if not v1.IsAlmostEqualTo(v2, 0.01):
                nodes.append(path[i])
        nodes.append(path[-1])

        new_pipes = []
        for i in range(len(nodes) - 1):
            if nodes[i].DistanceTo(nodes[i + 1]) < 0.01:
                continue
            p = Pipe.Create(doc, s_type_id, p_type_id, level_id, nodes[i], nodes[i + 1])
            # Ustawienie średnicy na 50mm (można zmienić)
            p.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM).Set(50.0 / 304.8)
            new_pipes.append(p)

        doc.Regenerate()

        # Automatyczne kolana
        for i in range(len(new_pipes) - 1):
            try:
                c1 = next(
                    c
                    for c in new_pipes[i].ConnectorManager.Connectors
                    if c.Origin.DistanceTo(nodes[i + 1]) < 0.1
                )
                c2 = next(
                    c
                    for c in new_pipes[i + 1].ConnectorManager.Connectors
                    if c.Origin.DistanceTo(nodes[i + 1]) < 0.1
                )
                doc.Create.NewElbowFitting(c1, c2)
            except:
                pass

        print("Trasa wygenerowana pomyślnie.")
    else:
        print("Błąd: Algorytm nie znalazł drogi omijającej ściany.")
