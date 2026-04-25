# -*- coding: utf-8 -*-
import clr
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import *
from pyrevit import script
from Autodesk.Revit.DB.Plumbing import *

# --- USTAWIENIA ---
doc = __revit__.ActiveUIDocument.Document
selection = __revit__.ActiveUIDocument.Selection.GetElementIds()
output = script.get_output()


def get_geometry_from_selection(selection_ids, document):
    extracted_data = []
    for eid in selection_ids:
        el = document.GetElement(eid)
        if isinstance(el, CurveElement):
            geom = el.GeometryCurve
            if isinstance(geom, Line):
                extracted_data.append({"id": eid, "curve": geom})
    return extracted_data


def get_pipe_params(document):
    system_types = (
        FilteredElementCollector(document).OfClass(PipingSystemType).ToElements()
    )
    sys_id = system_types[0].Id if system_types else None

    pipe_types = FilteredElementCollector(document).OfClass(PipeType).ToElements()
    typ_id = pipe_types[0].Id if pipe_types else None

    all_levels = FilteredElementCollector(document).OfClass(Level).ToElements()
    lvl_id = None
    for l in all_levels:
        if "1" in l.Name:
            lvl_id = l.Id
            break
    if not lvl_id and all_levels:
        lvl_id = all_levels[0].Id

    return sys_id, typ_id, lvl_id


def analyze_intersections(data_list):
    narozniki_L = []
    trojniki_T = []
    skrzyzowania_X = []
    tol = 0.001

    for i in range(len(data_list)):
        for j in range(i + 1, len(data_list)):
            l1, l2 = data_list[i]["curve"], data_list[j]["curve"]
            res_array = clr.Reference[IntersectionResultArray]()

            if l1.Intersect(l2, res_array) == SetComparisonResult.Overlap:
                actual_res = res_array.Value
                if actual_res and not actual_res.IsEmpty:
                    pt = actual_res.Item[0].XYZPoint

                    is_at_end1 = any(
                        pt.DistanceTo(l1.GetEndPoint(idx)) < tol for idx in [0, 1]
                    )
                    is_at_end2 = any(
                        pt.DistanceTo(l2.GetEndPoint(idx)) < tol for idx in [0, 1]
                    )

                    data_packet = {
                        "id_a": data_list[i]["id"],
                        "id_b": data_list[j]["id"],
                        "point": pt,
                    }

                    if is_at_end1 and is_at_end2:
                        narozniki_L.append(data_packet)
                    elif is_at_end1 or is_at_end2:
                        data_packet["main_pipe"] = (
                            data_list[j]["id"] if is_at_end1 else data_list[i]["id"]
                        )
                        trojniki_T.append(data_packet)
                    else:
                        skrzyzowania_X.append(data_packet)

    return {"l_shapes": narozniki_L, "t_shapes": trojniki_T, "x_shapes": skrzyzowania_X}


def create_pipe(document, datalist, id_systemu, id_typu, id_poziomu):
    # KLUCZOWA ZMIANA: Zwracamy słownik dla mapowania ID
    pipe_map = {}
    t = Transaction(document, "Generowanie rur")
    t.Start()
    for item in datalist:
        line_geom = item["curve"]
        try:
            new_pipe = Pipe.Create(
                document,
                id_systemu,
                id_typu,
                id_poziomu,
                line_geom.GetEndPoint(0),
                line_geom.GetEndPoint(1),
            )
            pipe_map[item["id"]] = new_pipe
        except Exception as e:
            print("Błąd przy tworzeniu rury z linii {}: {}".format(item["id"], e))
    t.Commit()
    return pipe_map


def create_elbows(document, pipe_map, l_shapes_data):
    """Tworzy kolanka w jednej zbiorczej transakcji."""
    elbows_created = []
    t = Transaction(document, "Wstawianie kolanek MEP")
    t.Start()

    for item in l_shapes_data:
        rura_a = pipe_map.get(item["id_a"])
        rura_b = pipe_map.get(item["id_b"])

        if rura_a and rura_b:
            cons1 = rura_a.ConnectorManager.Connectors
            cons2 = rura_b.ConnectorManager.Connectors
            closest_pair = None
            min_dist = 10.0

            for c1 in cons1:
                for c2 in cons2:
                    dist = c1.Origin.DistanceTo(c2.Origin)
                    if dist < min_dist:
                        min_dist = dist
                        closest_pair = (c1, c2)

            if closest_pair and min_dist < 0.1:
                try:
                    new_elbow = document.Create.NewElbowFitting(
                        closest_pair[0], closest_pair[1]
                    )
                    elbows_created.append(new_elbow)
                except Exception as e:
                    print(
                        "Błąd kolanka (Linii {} i {}): {}".format(
                            item["id_a"], item["id_b"], e
                        )
                    )

    t.Commit()
    return elbows_created


def create_tees(document, pipe_map, t_shapes_data):
    """Tworzy trójniki w jednej zbiorczej transakcji."""
    tees_created = []
    tol = 0.01

    t_trans = Transaction(document, "Wstawianie trójników MEP")
    t_trans.Start()

    for item in t_shapes_data:
        rura_a = pipe_map.get(item["id_a"])
        rura_b = pipe_map.get(item["id_b"])

        if rura_a and rura_b:
            pt = item["point"]

            # POPRAWKA: Pobieramy geometrię rury przez .Location.Curve
            curve_a = rura_a.Location.Curve
            curve_b = rura_b.Location.Curve

            # Sprawdzamy dystans końców rury A do punktu przecięcia, aby znaleźć branch
            dist_a = min(
                pt.DistanceTo(curve_a.GetEndPoint(0)),
                pt.DistanceTo(curve_a.GetEndPoint(1)),
            )

            # Jeśli rura A kończy się w punkcie skrzyżowania, jest ona odgałęzieniem (branch)
            if dist_a < tol:
                branch_pipe = rura_a
                main_pipe = rura_b
            else:
                branch_pipe = rura_b
                main_pipe = rura_a

            try:
                # 1. Rozcinamy rurę główną w punkcie pt (zwraca ID nowej rury)
                new_main_id = PlumbingUtils.BreakCurve(document, main_pipe.Id, pt)
                pipe_main_2 = document.GetElement(new_main_id)

                # 2. Pobieramy konektory z rury odgałęźnej (branch)
                c_branch = None
                for c in branch_pipe.ConnectorManager.Connectors:
                    if c.Origin.DistanceTo(pt) < tol:
                        c_branch = c
                        break

                # 3. Pobieramy konektory z dwóch części rury głównej (main)
                c_main_1 = None
                for c in main_pipe.ConnectorManager.Connectors:
                    if c.Origin.DistanceTo(pt) < tol:
                        c_main_1 = c
                        break

                c_main_2 = None
                for c in pipe_main_2.ConnectorManager.Connectors:
                    if c.Origin.DistanceTo(pt) < tol:
                        c_main_2 = c
                        break

                # 4. Jeśli mamy komplet 3 konektorów, tworzymy trójnik
                if c_main_1 and c_main_2 and c_branch:
                    new_tee = document.Create.NewTeeFitting(
                        c_main_1, c_main_2, c_branch
                    )
                    tees_created.append(new_tee)
            except Exception as e:
                print("Błąd trójnika przy ID {}: {}".format(item["id_a"], e))

    t_trans.Commit()
    return tees_created


def create_crosses(document, pipe_map, x_shapes_data):
    """Tworzy czwórniki, dbając o aktualność geometrii rur."""
    crosses_created = []
    tol = 0.05

    # Rozpoczynamy transakcję zbiorczą
    t_trans = Transaction(document, "Wstawianie czwórników MEP")
    t_trans.Start()

    for item in x_shapes_data:
        # 1. Pobieramy rury z mapy
        rura_a = pipe_map.get(item["id_a"])
        rura_b = pipe_map.get(item["id_b"])

        if rura_a and rura_b:
            try:
                # REGENERACJA: To kluczowe przy dużej ilości linii.
                # Revit musi przeliczyć geometrię po poprzednich cięciach.
                document.Regenerate()

                # 2. Pobieramy AKTUALNĄ krzywą rury (po ewentualnych wcześniejszych podziałach)
                curve_a = rura_a.Location.Curve
                curve_b = rura_b.Location.Curve

                pt_raw = item["point"]

                # 3. Rzutujemy punkt na rury - to niweluje błędy precyzji rzędu 0.000001
                pt_a = curve_a.Project(pt_raw).XYZPoint
                pt_b = curve_b.Project(pt_raw).XYZPoint

                # 4. Rozcinamy rurę A
                new_a_id = PlumbingUtils.BreakCurve(document, rura_a.Id, pt_a)
                pipe_a_2 = document.GetElement(new_a_id)

                # 5. Rozcinamy rurę B
                new_b_id = PlumbingUtils.BreakCurve(document, rura_b.Id, pt_b)
                pipe_b_2 = document.GetElement(new_b_id)

                # Funkcja do łapania konektorów
                def get_conn_at_point(pipe_obj, search_point):
                    for c in pipe_obj.ConnectorManager.Connectors:
                        if c.Origin.DistanceTo(search_point) < tol:
                            return c
                    return None

                # 6. Zbieramy 4 konektory
                c1 = get_conn_at_point(rura_a, pt_a)
                c2 = get_conn_at_point(pipe_a_2, pt_a)
                c3 = get_conn_at_point(rura_b, pt_b)
                c4 = get_conn_at_point(pipe_b_2, pt_b)

                if all([c1, c2, c3, c4]):
                    new_cross = document.Create.NewCrossFitting(c1, c2, c3, c4)
                    crosses_created.append(new_cross)

            except Exception as e:
                # Jeśli rura została już pocięta przez inny czwórnik (wspólny węzeł),
                # tutaj możemy zalogować błąd i przejść dalej.
                print(
                    "Pominięto czwórnik (ID {}/{}): {}".format(
                        item["id_a"], item["id_b"], e
                    )
                )

    t_trans.Commit()
    return crosses_created


def generate_mep_system(doc, dane_geo):
    """Główna funkcja generująca wszystko pod jednym Ctrl+Z."""

    # Tworzymy grupę transakcji
    tg = TransactionGroup(doc, "Generuj System MEP z Linii")
    tg.Start()

    try:
        # 1. Pobranie parametrów
        sys_id, typ_id, lvl_id = get_pipe_params(doc)

        # 2. Tworzenie rur
        pipe_map = create_pipe(doc, dane_geo, sys_id, typ_id, lvl_id)

        # 3. Analiza skrzyżowań
        wyniki = analyze_intersections(dane_geo)

        # 4. Tworzenie złączek
        stworzone_x = create_crosses(doc, pipe_map, wyniki["x_shapes"])
        stworzone_t = create_tees(doc, pipe_map, wyniki["t_shapes"])
        stworzone_l = create_elbows(doc, pipe_map, wyniki["l_shapes"])

        # --- NOWA SEKCJA: KASOWANIE LINII ---
        t_del = Transaction(doc, "Usuwanie linii źródłowych")
        t_del.Start()
        for item in dane_geo:
            try:
                doc.Delete(item["id"])
            except:
                pass  # Ignoruj, jeśli linia została już usunięta lub jest niedostępna
        t_del.Commit()
        # ------------------------------------

        # Assimilate sprawia, że rury + złączki + usunięcie linii to jedna pozycja w menu Undo
        tg.Assimilate()

        return {
            "pipes": len(pipe_map),
            "elbows": len(stworzone_l),
            "tees": len(stworzone_t),
            "crosses": len(stworzone_x),
            "raw_results": wyniki,
        }

    except Exception as e:
        tg.RollBack()
        print("BŁĄD KRYTYCZNY: Operation rolled back. Info: {}".format(e))
        return None


# --- WYKONANIE ---
dane_geo = get_geometry_from_selection(selection, doc)

if not dane_geo:
    print("Zaznacz linie przed uruchomieniem.")
else:
    # Uruchamiamy wszystko w jednej paczce
    raport_dane = generate_mep_system(doc, dane_geo)

    if raport_dane:
        # --- RAPORT ---
        output.print_md("# 🛠️ Raport Systemu MEP (Pojedyncza operacja Undo)")

        print("- 📈 Stworzone rury: {}".format(raport_dane["pipes"]))
        print("- ∟ Kolanka (L): {}".format(raport_dane["elbows"]))
        print("- ⊥ Trójniki (T): {}".format(raport_dane["tees"]))
        print("- + Czwórniki (X): {}".format(raport_dane["crosses"]))

        # Wyświetlenie tabeli (korzystamy z zapisanych wyników analizy)
        table_data = []
        wyniki = raport_dane["raw_results"]
        kategorie = [
            ("l_shapes", "Narożnik L"),
            ("t_shapes", "Trójnik T"),
            ("x_shapes", "Czwórnik X"),
        ]

        for klucz, nazwa_typu in kategorie:
            for s in wyniki[klucz]:
                table_data.append(
                    [
                        output.linkify(s["id_a"]),
                        output.linkify(s["id_b"]),
                        nazwa_typu,
                        "X: {:.2f}, Y: {:.2f}".format(s["point"].X, s["point"].Y),
                    ]
                )

        if table_data:
            output.print_table(table_data, columns=["ID A", "ID B", "Typ", "Punkt"])
