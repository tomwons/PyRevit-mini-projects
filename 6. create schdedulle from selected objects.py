# -*- coding: utf-8 -*-
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import *
from pyrevit import script

output = script.get_output()
logger = script.get_logger()

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
selection_ids = uidoc.Selection.GetElementIds()

output.print_md("# 🚀 Generator Zestawień MEP v3.0 (Tylko Zaznaczone)")

if not selection_ids:
    logger.warning("BŁĄD: Nic nie zaznaczyłeś! Musisz zaznaczyć elementy, które mają trafić do tabeli.")
else:
    # 1. Przygotowanie unikalnego klucza dla zaznaczonych elementów
    # Nadajemy zaznaczonym elementom tymczasowy komentarz, aby móc je przefiltrować w zestawieniu
    timestamp = datetime.datetime.now().strftime("%Y%m%d%H%M%S")
    unique_filter_val = "Selection_" + timestamp

    # 2. Grupowanie kategorii i oznaczanie elementów
    elements_by_category = {}
    
    t_mark = Transaction(doc, "Oznaczanie elementów do zestawienia")
    t_mark.Start()
    for eid in selection_ids:
        el = doc.GetElement(eid)
        if el and el.Category:
            # Nadajemy unikalną wartość w parametrze "Komentarze", aby filtr zadziałał
            p = el.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS)
            if p:
                p.Set(unique_filter_val)
            
            cat_key = el.Category.Id.ToString()
            if cat_key not in elements_by_category:
                elements_by_category[cat_key] = el.Category
    t_mark.Commit()

    # 3. Lista parametrów
    targets = [
        (BuiltInParameter.ELEM_FAMILY_PARAM, "Rodzina"),
        (BuiltInParameter.ELEM_FAMILY_AND_TYPE_PARAM, "Rodzina i typ"),
        (None, "Liczba"), 
        (BuiltInParameter.CURVE_ELEM_LENGTH, "Długość"),
        (BuiltInParameter.RBS_PIPE_DIAMETER_PARAM, "Średnica"),
        (BuiltInParameter.RBS_CURVE_WIDTH_PARAM, "Szerokość"),
        (BuiltInParameter.RBS_CURVE_HEIGHT_PARAM, "Wysokość"),
        (BuiltInParameter.RBS_CALCULATED_SIZE, "Wielkość"),
        (BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS, "Filtr_ID") # Musi być do filtra
    ]

    t = Transaction(doc, "Generowanie Zestawień")
    t.Start()

    try:
        # Usuwanie starych zestawień
        existing_schedules = FilteredElementCollector(doc).OfClass(ViewSchedule).ToElements()
        for old_sched in existing_schedules:
            if "Zestawienie_Automatyczne_" in old_sched.Name:
                try: doc.Delete(old_sched.Id)
                except: pass

        for cat_key, category in elements_by_category.items():
            target_name = "Zestawienie_Automatyczne_{}".format(category.Name.replace(" ", "_"))
            
            # Tworzenie widoku
            try:
                new_schedule = ViewSchedule.CreateDataSchedule(doc, category.Id)
            except:
                new_schedule = ViewSchedule.CreateSchedule(doc, category.Id)
            
            new_schedule.Name = target_name
            definition = new_schedule.Definition
            definition.ShowGrandTotal = True

            # Dodawanie pól
            schedulable_fields = definition.GetSchedulableFields()
            filter_field = None

            for b_param, p_name in targets:
                found_field = None
                if b_param is not None:
                    p_id = ElementId(b_param)
                    for s_field in schedulable_fields:
                        if s_field.ParameterId == p_id:
                            found_field = definition.AddField(s_field)
                            if b_param == BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS:
                                filter_field = found_field # Zapamiętujemy pole do filtra
                            break
                else:
                    for s_field in schedulable_fields:
                        if s_field.FieldType == ScheduleFieldType.Count:
                            found_field = definition.AddField(s_field)
                            break

                if found_field:
                    if b_param == BuiltInParameter.CURVE_ELEM_LENGTH or b_param is None:
                        try: found_field.HasTotals = True
                        except: pass

            # --- KLUCZOWY FILTR: Tylko zaznaczone elementy ---
            if filter_field:
                # Tworzymy filtr: Komentarze RÓWNA SIĘ unique_filter_val
                filt = ScheduleFilter(filter_field.FieldId, ScheduleFilterType.Equal, unique_filter_val)
                definition.AddFilter(filt)
                # Ukrywamy kolumnę filtra, żeby nie szpeciła zestawienia
                filter_field.IsHidden = True

            output.print_md(" Stworzono zestawienie dla: **{}** (Tylko zaznaczone)".format(category.Name))
        
        t.Commit()
        output.print_md("---")
        output.print_md("###  GOTOWE! Zestawienia zawierają tylko wybrane elementy.")

    except Exception as e:
        if t.GetStatus() == TransactionStatus.Started:
            t.RollBack()
        logger.error("BŁĄD: {}".format(str(e)))
