# -*- coding: utf-8 -*-
__title__ = "Seryjne Wstawianie Zaworów v49 (Safe Buffer)"

from pyrevit import revit, forms, script
import clr

clr.AddReference("RevitAPI")
from Autodesk.Revit.DB import *
from Autodesk.Revit.DB.Plumbing import Pipe
from Autodesk.Revit.DB.Structure import StructuralType

doc = revit.doc


def get_valve_conns(inst):
    try:
        if inst.MEPModel and inst.MEPModel.ConnectorManager:
            return [c for c in inst.MEPModel.ConnectorManager.Connectors]
    except:
        return []
    return []


# 1. WYBÓR ZAWORU
acc_types = (
    FilteredElementCollector(doc)
    .OfClass(FamilySymbol)
    .OfCategory(BuiltInCategory.OST_PipeAccessory)
    .ToElements()
)
acc_dict = {
    "{}: {}".format(t.Family.Name, revit.query.get_name(t)): t for t in acc_types
}
if not acc_dict:
    script.exit()

sel_name = forms.SelectFromList.show(
    sorted(acc_dict.keys()), multiselect=False, title="Wybierz zawór"
)
if not sel_name:
    script.exit()
sym = acc_dict[sel_name]

# 2. SELEKCJA RUR
selection = [el for el in revit.get_selection() if isinstance(el, Pipe)]
if not selection:
    forms.alert("Zaznacz rurę.")
    script.exit()

# Stałe parametry
spacing_ft = 1000.0 / 304.8
safe_buffer_ft = 1000.0 / 304.8  # Minimalny odstęp od końców rury

# 3. TRANSAKCJA
with revit.Transaction("Seryjne wstawianie v49 - Zabezpieczenie"):
    if not sym.IsActive:
        sym.Activate()

    for original_pipe in selection:
        try:
            # POBRANIE DANYCH BAZOWYCH
            diam_ft = original_pipe.get_Parameter(
                BuiltInParameter.RBS_PIPE_DIAMETER_PARAM
            ).AsDouble()
            radius_ft = diam_ft / 2.0
            p_type = original_pipe.PipeType.Id
            level_id = original_pipe.LevelId
            sys_id = original_pipe.get_Parameter(
                BuiltInParameter.RBS_PIPING_SYSTEM_TYPE_PARAM
            ).AsElementId()

            curve = original_pipe.Location.Curve
            p_start, p_end = curve.GetEndPoint(0), curve.GetEndPoint(1)
            full_vec = (p_end - p_start).Normalize()
            total_len = curve.Length

            # --- MECHANIZM ZABEZPIECZAJĄCY ---
            # Obliczamy dostępną długość: Całość - Margines na początku - Margines na końcu
            available_len = total_len - (2 * safe_buffer_ft)

            if available_len <= 0:
                print(
                    "Rura (ID: {}) zbyt krótka na zachowanie odstępów 1000mm od końców.".format(
                        original_pipe.Id
                    )
                )
                continue

            num_valves = int(available_len / spacing_ft) + 1
            # Jeśli dystans dostępny jest dokładnie wielokrotnością, korygujemy
            # (aby ostatni zawór nie wypadł idealnie na 1000mm przed końcem, co może być ryzykowne)

            # ZAPAMIĘTANIE POŁĄCZEŃ ZEWNĘTRZNYCH
            ext_refs_start, ext_refs_end = [], []
            for c in original_pipe.ConnectorManager.Connectors:
                if c.Origin.DistanceTo(p_start) < 0.05:
                    for r in c.AllRefs:
                        if r.Owner.Id != original_pipe.Id:
                            ext_refs_start.append(r)
                if c.Origin.DistanceTo(p_end) < 0.05:
                    for r in c.AllRefs:
                        if r.Owner.Id != original_pipe.Id:
                            ext_refs_end.append(r)

            # USUNIĘCIE STAREJ RURY
            doc.Delete(original_pipe.Id)
            doc.Regenerate()

            # --- PĘTLA WSTAWIANIA ---
            current_cursor_pt = p_start
            last_valve_out_conn = None

            for i in range(1, num_valves + 1):
                # Punkt wstawienia: Start + Margines + (Kolejne kroki)
                # i=1 -> wstawia na 1000mm, i=2 -> na 2000mm itd.
                dist_to_ins = safe_buffer_ft + (spacing_ft * (i - 1))

                # Zabezpieczenie nadmiarowe: jeśli punkt wypadnie poza bezpieczny koniec rury
                if dist_to_ins > total_len - safe_buffer_ft + 0.01:
                    break

                ins_pt = p_start + (full_vec * dist_to_ins)

                # Wstawienie zaworu
                f_inst = doc.Create.NewFamilyInstance(
                    ins_pt,
                    sym,
                    full_vec,
                    doc.GetElement(level_id),
                    StructuralType.NonStructural,
                )

                # Parametry rozmiaru
                for p in f_inst.Parameters:
                    if p.IsReadOnly:
                        continue
                    p_name = p.Definition.Name.lower()
                    if (
                        any(x in p_name for x in ["radius", "promień", "promien"])
                        or p_name == "r"
                    ):
                        p.Set(radius_ft)
                    elif any(
                        x in p_name for x in ["diameter", "średnica", "size", "dn"]
                    ):
                        p.Set(diam_ft)

                doc.Regenerate()

                v_conns = sorted(
                    get_valve_conns(f_inst), key=lambda c: c.Origin.DistanceTo(p_start)
                )
                vc_in, vc_out = v_conns[0], v_conns[1]

                # Rura ŁĄCZĄCA
                new_pipe = Pipe.Create(
                    doc, sys_id, p_type, level_id, current_cursor_pt, vc_in.Origin
                )
                new_pipe.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM).Set(
                    diam_ft
                )
                doc.Regenerate()

                # Łączenie
                for pc in new_pipe.ConnectorManager.Connectors:
                    if pc.Origin.DistanceTo(vc_in.Origin) < 0.05:
                        pc.ConnectTo(vc_in)
                    if i == 1 and pc.Origin.DistanceTo(p_start) < 0.05:
                        for r in ext_refs_start:
                            try:
                                pc.ConnectTo(r)
                            except:
                                pass
                    elif i > 1 and pc.Origin.DistanceTo(current_cursor_pt) < 0.05:
                        if last_valve_out_conn:
                            pc.ConnectTo(last_valve_out_conn)

                current_cursor_pt = vc_out.Origin
                last_valve_out_conn = vc_out

            # --- OSTATNI ODCINEK ---
            last_pipe = Pipe.Create(
                doc, sys_id, p_type, level_id, current_cursor_pt, p_end
            )
            last_pipe.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM).Set(
                diam_ft
            )
            doc.Regenerate()

            for pc in last_pipe.ConnectorManager.Connectors:
                if pc.Origin.DistanceTo(current_cursor_pt) < 0.05:
                    pc.ConnectTo(last_valve_out_conn)
                if pc.Origin.DistanceTo(p_end) < 0.05:
                    for r in ext_refs_end:
                        try:
                            pc.ConnectTo(r)
                        except:
                            pass

        except Exception as e:
            print("Błąd: {}".format(e))

forms.alert("Zakończono. Zachowano margines bezpieczeństwa 1000mm od połączeń.")
