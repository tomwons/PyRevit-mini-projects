# -*- coding: utf-8 -*-
__title__ = "Bypass"


from pyrevit import revit, forms, script
import clr

clr.AddReference("System")
from System.Collections.Generic import List

clr.AddReference("RevitAPI")
from Autodesk.Revit.DB import *
from Autodesk.Revit.DB.Plumbing import Pipe
from Autodesk.Revit.DB.Mechanical import Duct

doc = revit.doc


def get_conns(el):
    """Pobiera konektory fizyczne elementu."""
    try:
        return [
            c
            for c in el.ConnectorManager.Connectors
            if c.ConnectorType != ConnectorType.Logical
        ]
    except:
        return []


# --- PARAMETRY ---
gap_mm = 10.0  # Luz między ściankami (1 cm)
buffer_mm = 10.0  # Odległość od krawędzi omijanej rury

gap_ft = gap_mm / 304.8
buff_ft = buffer_mm / 304.8

selection = [el for el in revit.get_selection() if isinstance(el, (Pipe, Duct))]
if not selection:
    forms.alert("Zaznacz rurę lub kanał do ominięcia.")
    script.exit()

# FILTR: Szukamy kolizji tylko z rurami i kanałami (ignorujemy centrale)
cat_list = List[BuiltInCategory]()
cat_list.Add(BuiltInCategory.OST_PipeCurves)
cat_list.Add(BuiltInCategory.OST_DuctCurves)
multi_filter = ElementMulticategoryFilter(cat_list)

with revit.Transaction("Bypass Universal"):
    view_3d = next(
        (v for v in FilteredElementCollector(doc).OfClass(View3D) if not v.IsTemplate),
        None,
    )

    for main_el in selection:
        curve = main_el.Location.Curve
        p_start, p_end = curve.GetEndPoint(0), curve.GetEndPoint(1)
        direction = (p_end - p_start).Normalize()

        # Szukamy przeszkód na drodze rury/kanału
        intersector = ReferenceIntersector(
            multi_filter, FindReferenceTarget.Element, view_3d
        )
        all_hits = intersector.Find(p_start, direction)

        # Filtrujemy trafienia, by nie brać pod uwagę końców rury (min 0.5 ft od startu/końca)
        valid_hits = [h for h in all_hits if 0.5 < h.Proximity < (curve.Length - 0.5)]

        if not valid_hits:
            continue

        # Pierwsza przeszkoda (rura lub kanał)
        hit = sorted(valid_hits, key=lambda h: h.Proximity)[0]
        obs_el = doc.GetElement(hit.GetReference().ElementId)
        pt_hit = hit.GetReference().GlobalPoint

        # Centrowanie na osi przeszkody
        obs_curve = obs_el.Location.Curve
        projection = obs_curve.Project(pt_hit)
        pt_center = projection.XYZPoint

        # Rozmiary (Pipe vs Duct)
        def get_size(el):
            # Średnica dla rur/kanałów okrągłych
            p_dia = el.get_Parameter(
                BuiltInParameter.RBS_PIPE_DIAMETER_PARAM
            ) or el.get_Parameter(BuiltInParameter.RBS_CURVE_DIAMETER_PARAM)
            if p_dia:
                return p_dia.AsDouble()
            # Wysokość dla kanałów prostokątnych
            p_h = el.get_Parameter(BuiltInParameter.RBS_CURVE_HEIGHT_PARAM)
            if p_h:
                return p_h.AsDouble()
            return 0.5

        d_obs = get_size(obs_el)
        d_main = get_size(main_el)

        # Obliczenia 45 stopni "na styk"
        v_shift_ft = (d_main / 2) + (d_obs / 2) + gap_ft
        h_shift_ft = v_shift_ft

        # Geometria kompaktowa
        dist_to_edge = (d_obs / 2) + buff_ft
        pb = pt_center - direction * dist_to_edge + XYZ.BasisZ * v_shift_ft
        pc = pt_center + direction * dist_to_edge + XYZ.BasisZ * v_shift_ft
        pa = pb - direction * h_shift_ft - XYZ.BasisZ * v_shift_ft
        pd = pc + direction * h_shift_ft - XYZ.BasisZ * v_shift_ft

        # Parametry systemowe
        el_type = main_el.GetTypeId()
        sys_param = main_el.get_Parameter(
            BuiltInParameter.RBS_PIPING_SYSTEM_TYPE_PARAM
            if isinstance(main_el, Pipe)
            else BuiltInParameter.RBS_DUCT_SYSTEM_TYPE_PARAM
        )
        sys_id = sys_param.AsElementId() if sys_param else ElementId.InvalidElementId
        lev_id = main_el.LevelId

        # Rozmiary kanałów prostokątnych (jeśli dotyczy)
        w_param = main_el.get_Parameter(BuiltInParameter.RBS_CURVE_WIDTH_PARAM)
        h_param = main_el.get_Parameter(BuiltInParameter.RBS_CURVE_HEIGHT_PARAM)
        width = w_param.AsDouble() if w_param else None
        height = h_param.AsDouble() if h_param else None

        doc.Delete(main_el.Id)
        doc.Regenerate()

        pts = [(p_start, pa), (pa, pb), (pb, pc), (pc, pd), (pd, p_end)]
        new_segs = []

        for p1, p2 in pts:
            if p1.DistanceTo(p2) < 0.005:
                continue
            if isinstance(main_el, Pipe):
                s = Pipe.Create(doc, sys_id, el_type, lev_id, p1, p2)
                s.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM).Set(d_main)
            else:
                s = Duct.Create(doc, sys_id, el_type, lev_id, p1, p2)
                if width and height:
                    s.get_Parameter(BuiltInParameter.RBS_CURVE_WIDTH_PARAM).Set(width)
                    s.get_Parameter(BuiltInParameter.RBS_CURVE_HEIGHT_PARAM).Set(height)
                else:
                    s.get_Parameter(BuiltInParameter.RBS_CURVE_DIAMETER_PARAM).Set(
                        d_main
                    )
            new_segs.append(s)

        doc.Regenerate()
        for i in range(len(new_segs) - 1):
            c1, c2 = get_conns(new_segs[i]), get_conns(new_segs[i + 1])
            for ca in c1:
                for cb in c2:
                    if ca.Origin.DistanceTo(cb.Origin) < 0.1:
                        try:
                            doc.Create.NewElbowFitting(ca, cb)
                        except:
                            pass
