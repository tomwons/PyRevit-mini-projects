# -*- coding: utf-8 -*-
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import *
from pyrevit import script
import datetime

# Inicjalizacja wyjścia pyRevit
output = script.get_output()
logger = script.get_logger()

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
selection_ids = uidoc.Selection.GetElementIds()

output.print_md("# 🚀 Generator Zestawień MEP v3.1 (Pancerny)")

if not selection_ids:
    logger.warning(
        "BŁĄD: Nic nie zaznaczyłeś! Zaznacz elementy w modelu przed uruchomieniem."
    )
else:
    # 1. Przygotowanie unikalnego klucza dla filtra
    timestamp = datetime.datetime.now().strftime("%Y%m%d%H%M%S")
    unique_filter_val = "Selection_" + timestamp

    # 2. Grupowanie kategorii i oznaczanie elementów w jednej transakcji
    elements_by_category = {}

    t_mark = Transaction(doc, "Oznaczanie elementów do zestawienia")
    t_mark.Start()
    for eid in selection_ids:
        el = doc.GetElement(eid)
        if el and el.Category:
            # Nadajemy unikalną wartość w parametrze "Komentarze" (Comments)
            p = el.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS)
            if p:
                p.Set(unique_filter_val)

            cat_key = el.Category.Id.ToString()
            if cat_key not in elements_by_category:
                elements_by_category[cat_key] = el.Category
    t_mark.Commit()

    # 3. Definicja parametrów docelowych (BuiltInParameter i nazwa)
    targets = [
        (BuiltInParameter.ELEM_FAMILY_PARAM, "Rodzina"),
        (BuiltInParameter.ELEM_FAMILY_AND_TYPE_PARAM, "Rodzina i typ"),
        (None, "Liczba"),  # Pole specjalne 'Count'
        (BuiltInParameter.CURVE_ELEM_LENGTH, "Długość"),
        (BuiltInParameter.RBS_PIPE_DIAMETER_PARAM, "Średnica"),
        (BuiltInParameter.RBS_CURVE_WIDTH_PARAM, "Szerokość"),
        (BuiltInParameter.RBS_CURVE_HEIGHT_PARAM, "Wysokość"),
        (BuiltInParameter.RBS_CALCULATED_SIZE, "Wielkość"),
        (BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS, "Filtr_ID"),
    ]

    t = Transaction(doc, "Generowanie Zestawień")
    t.Start()

    try:
        # Usuwanie starych zestawień automatycznych (czyszczenie widoków)
        existing_schedules = (
            FilteredElementCollector(doc).OfClass(ViewSchedule).ToElements()
        )
        for old_sched in existing_schedules:
            if "Zestawienie_Automatyczne_" in old_sched.Name:
                try:
                    doc.Delete(old_sched.Id)
                except:
                    pass

        for cat_key, category in elements_by_category.items():
            # --- WALIDACJA KATEGORII ---
            if not ViewSchedule.IsValidCategoryForSchedule(category.Id):
                output.print_md(
                    "⚠️ Pominięto kategorię: **{}** (Revit nie pozwala na jej zestawienie)".format(
                        category.Name
                    )
                )
                continue

            target_name = "Zestawienie_Automatyczne_{}".format(
                category.Name.replace(" ", "_")
            )

            # Próba stworzenia zestawienia (DataSchedule lub zwykłe)
            new_schedule = None
            try:
                new_schedule = ViewSchedule.CreateDataSchedule(doc, category.Id)
            except:
                try:
                    new_schedule = ViewSchedule.CreateSchedule(doc, category.Id)
                except Exception as e:
                    output.print_md(
                        "❌ Błąd krytyczny dla kategorii **{}**: {}".format(
                            category.Name, str(e)
                        )
                    )
                    continue

            if not new_schedule:
                continue

            new_schedule.Name = target_name
            definition = new_schedule.Definition
            definition.ShowGrandTotal = True

            # Dodawanie pól do zestawienia
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
                                filter_field = found_field
                            break
                else:
                    # Obsługa pola "Liczba" (Count)
                    for s_field in schedulable_fields:
                        if s_field.FieldType == ScheduleFieldType.Count:
                            found_field = definition.AddField(s_field)
                            break

                # Sumowanie wartości dla Długości i Liczby
                if found_field:
                    if b_param == BuiltInParameter.CURVE_ELEM_LENGTH or b_param is None:
                        try:
                            found_field.HasTotals = True
                        except:
                            pass

            # --- APLIKACJA FILTRA (TYLKO ZAZNACZONE) ---
            if filter_field:
                filt = ScheduleFilter(
                    filter_field.FieldId, ScheduleFilterType.Equal, unique_filter_val
                )
                definition.AddFilter(filt)
                filter_field.IsHidden = True  # Ukrywamy techniczny parametr filtra

            output.print_md("✅ Stworzono: **{}**".format(category.Name))

        t.Commit()
        output.print_md("---")
        output.print_md(
            "### 🏁 Sukces! Sprawdź przeglądarkę projektów w sekcji Zestawienia."
        )

    except Exception as e:
        if t.GetStatus() == TransactionStatus.Started:
            t.RollBack()
        logger.error("BŁĄD OGÓLNY: {}".format(str(e)))
