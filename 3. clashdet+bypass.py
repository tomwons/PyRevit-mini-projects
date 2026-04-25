# -*- coding: utf-8 -*-
__title__ = "Auto Bypass Boczny v59"

from pyrevit import revit, forms, script
import clr

clr.AddReference("System")
from System.Collections.Generic import List

clr.AddReference("RevitAPI")
from Autodesk.Revit.DB import *
from Autodesk.Revit.DB.Plumbing import Pipe

doc = revit.doc


def get_pipe_conns(pipe_el):
    try:
        return [c for c in pipe_el.ConnectorManager.Connectors]
    except:
        return []


# --- PARAMETRY ---
buffer_mm = 300.0  # Odstęp przed/za ścianą
safety_side_mm = 200.0  # Zapas poza krawędź boczną ściany

buff_ft = buffer_mm / 304.8
side_safety_ft = safety_side_mm / 304.8

# 1. SELEKCJA
selection = [el for el in revit.get_selection() if isinstance(el, Pipe)]
if not selection:
    forms.alert("Zaznacz rurę.")
    script.exit()

cat_list = List[BuiltInCategory]()
cat_list.Add(BuiltInCategory.OST_Walls)
multi_filter = ElementMulticategoryFilter(cat_list)

with revit.Transaction("Boczny Skok v59"):
    view_3d = next(
        (v for v in FilteredElementCollector(doc).OfClass(View3D) if not v.IsTemplate),
        None,
    )

    for pipe in selection:
        curve = pipe.Location.Curve
        p_start, p_end = curve.GetEndPoint(0), curve.GetEndPoint(1)
        direction = (p_end - p_start).Normalize()

        # WEKTOR BOCZNY (Prostopadły do rury w poziomie)
        # Cross product kierunku rury i osi Z daje nam wektor "w bok"
        side_vec = direction.CrossProduct(XYZ.BasisZ).Normalize()

        intersector = ReferenceIntersector(
            multi_filter, FindReferenceTarget.Face, view_3d
        )
        all_hits = intersector.Find(p_start, direction)
        valid_hits = [h for h in all_hits if h.Proximity <= curve.Length]

        if not valid_hits:
            continue

        valid_hits = sorted(valid_hits, key=lambda h: h.Proximity)
        hit_in, hit_out = valid_hits[0], valid_hits[-1]
        pt_in = hit_in.GetReference().GlobalPoint
        pt_out = hit_out.GetReference().GlobalPoint

        # ANALIZA BOKU ŚCIANY
        element = doc.GetElement(hit_in.GetReference().ElementId)
        bbox = element.get_BoundingBox(None)

        # Sprawdzamy, w którą stronę od osi rury ściana jest "krótsza"
        # Aby wybrać optymalną stronę ominięcia (lewą lub prawą)
        center_side_dist = (bbox.Max + bbox.Min) * 0.5
        # Decyzja: idziemy w stronę wektora side_vec
        # Obliczamy ekstremum ściany w kierunku bocznym
        # Uproszczenie: sprawdzamy Max/Min X lub Y zależnie od orientacji
        side_offset_ft = 0

        # Obliczamy jak daleko rura musi się odsunąć od swojej osi
        # Szukamy rzutu punktów narożnych bboxa na wektor boczny
        bbox_corners = [
            XYZ(bbox.Min.X, bbox.Min.Y, pt_in.Z),
            XYZ(bbox.Max.X, bbox.Min.Y, pt_in.Z),
            XYZ(bbox.Min.X, bbox.Max.Y, pt_in.Z),
            XYZ(bbox.Max.X, bbox.Max.Y, pt_in.Z),
        ]

        max_proj = 0
        for corner in bbox_corners:
            vec_to_corner = corner - pt_in
            proj = vec_to_corner.DotProduct(side_vec)
            if proj > max_proj:
                max_proj = proj

        side_jump_dist_ft = max_proj + side_safety_ft

        # 2. PUNKTY BRAMKI BOCZNEJ
        pa = pt_in - direction * buff_ft
        pb = pa + side_vec * side_jump_dist_ft

        pc = pt_out + direction * buff_ft
        pd = pc + side_vec * side_jump_dist_ft

        # 3. REKONSTRUKCJA
        p_type = pipe.PipeType.Id
        sys_id = pipe.get_Parameter(
            BuiltInParameter.RBS_PIPING_SYSTEM_TYPE_PARAM
        ).AsElementId()
        level_id = pipe.LevelId
        diam = pipe.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM).AsDouble()

        doc.Delete(pipe.Id)
        doc.Regenerate()

        # Rury: Start -> PrzedŚcianą -> ObokŚciany(In) -> ObokŚciany(Out) -> ZaŚcianą -> Koniec
        r1 = Pipe.Create(doc, sys_id, p_type, level_id, p_start, pa)
        r2 = Pipe.Create(doc, sys_id, p_type, level_id, pa, pb)
        r3 = Pipe.Create(doc, sys_id, p_type, level_id, pb, pd)
        r4 = Pipe.Create(doc, sys_id, p_type, level_id, pd, pc)
        r5 = Pipe.Create(doc, sys_id, p_type, level_id, pc, p_end)

        for r in [r1, r2, r3, r4, r5]:
            r.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM).Set(diam)

        doc.Regenerate()

        # 4. ŁĄCZENIE
        def connect(p_a, p_b):
            for ca in get_pipe_conns(p_a):
                for cb in get_pipe_conns(p_b):
                    if ca.Origin.DistanceTo(cb.Origin) < 0.01:
                        try:
                            doc.Create.NewElbowFitting(ca, cb)
                        except:
                            pass

        connect(r1, r2)
        connect(r2, r3)
        connect(r3, r4)
        connect(r4, r5)

forms.alert("Ominięcie boczne gotowe!")
