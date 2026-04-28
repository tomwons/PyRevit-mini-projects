# -*- coding: utf-8 -*-
__title__ = "Generator Arkuszy"

from pyrevit import revit, forms, script
import clr

clr.AddReference("RevitAPI")
from Autodesk.Revit.DB import *

doc = revit.doc

# 1. WYBÓR TABELKI
titleblocks = (
    FilteredElementCollector(doc)
    .OfCategory(BuiltInCategory.OST_TitleBlocks)
    .WhereElementIsElementType()
    .ToElements()
)
t_block_dict = {Element.Name.__get__(tb): tb.Id for tb in titleblocks}
selected_tb_name = forms.SelectFromList.show(
    t_block_dict.keys(), title="Wybierz tabelkę", multiselect=False
)
if not selected_tb_name:
    script.exit()

# 2. SKALA
target_scale = forms.ask_for_string(default="50", prompt="Skala:", title="Skala")
scale_val = int(target_scale) if target_scale.isdigit() else 50

# 3. WIDOKI
all_views = (
    FilteredElementCollector(doc)
    .OfClass(ViewPlan)
    .WhereElementIsNotElementType()
    .ToElements()
)
available_views = [v for v in all_views if not v.IsTemplate and v.CanBePrinted]
selected_view_names = forms.SelectFromList.show(
    [v.Name for v in available_views], title="Wybierz rzuty", multiselect=True
)

# 4. TRANSAKCJA
with revit.Transaction("Arkusze Centrowanie 70mm"):
    for v_name in selected_view_names:
        view = next(v for v in available_views if v.Name == v_name)

        try:
            view.Scale = scale_val
            view.CropBoxActive = True
            view.CropBoxVisible = False

            new_sheet = ViewSheet.Create(doc, t_block_dict[selected_tb_name])
            new_sheet.Name = view.Name

            # Pobieramy tabelkę na nowym arkuszu, by wyznaczyć jego środek
            sheet_tb = (
                FilteredElementCollector(doc, new_sheet.Id)
                .OfCategory(BuiltInCategory.OST_TitleBlocks)
                .FirstElement()
            )

            if sheet_tb:
                bbox = sheet_tb.get_BoundingBox(new_sheet)

                # Matematyczny środek
                center_x = (bbox.Min.X + bbox.Max.X) / 2
                center_y = (bbox.Min.Y + bbox.Max.Y) / 2

                # --- TWOJA KOREKTA ---
                # 70mm = ~0.23 stopy. Odejmujemy od X, żeby przesunąć w lewo.
                offset_left = 70 / 304.8
                center_x -= offset_left

                insertion_pt = XYZ(center_x, center_y, 0)
            else:
                insertion_pt = XYZ(1.22, 1.05, 0)  # Fallback

            if Viewport.CanAddViewToSheet(doc, new_sheet.Id, view.Id):
                Viewport.Create(doc, new_sheet.Id, view.Id, insertion_pt)
            else:
                doc.Delete(new_sheet.Id)

        except Exception as e:
            print("Błąd {}: {}".format(v_name, str(e)))

forms.alert("Gotowe! Rzuty przesunięte o 70mm w lewo od środka arkusza.")
