# -*- coding: utf-8 -*-
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import *
from Autodesk.Revit.DB.Plumbing import Pipe
from Autodesk.Revit.DB.Mechanical import Duct

doc = __revit__.ActiveUIDocument.Document
view = doc.ActiveView # Nadpisania grafiki działają na konkretnym widoku
selection = __revit__.ActiveUIDocument.Selection.GetElementIds()

if not selection:
    TaskDialog.Show("Błąd", "Najpierw zaznacz elementy!")
else:
    t = Transaction(doc, "Toggle Komentarz i Kolor")
    t.Start()
    
    text_to_handle = "Materiał do zamówienia"
    
    # 1. Przygotowanie koloru żółtego i wypełnienia
    yellow = Color(255, 255, 0) # RGB dla żółtego
    # Szukamy wzoru wypełnienia "Solid fill" w projekcie
    fill_pattern = FilteredElementCollector(doc).OfClass(FillPatternElement).FirstElement()
    
    # 2. Ustawienia dla statusu "Aktywne" (żółte)
    override_on = OverrideGraphicSettings()
    override_on.SetSurfaceForegroundPatternId(fill_pattern.Id)
    override_on.SetSurfaceForegroundPatternColor(yellow)
    
    # 3. Ustawienia dla statusu "Reset" (domyślne)
    override_off = OverrideGraphicSettings() # Puste ustawienia resetują widok
    
    count_added = 0
    count_removed = 0
    
    for el_id in selection:
        element = doc.GetElement(el_id)
        
        if isinstance(element, (Pipe, Duct, FamilyInstance)):
            p_comment = element.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS)
            
            if p_comment and not p_comment.IsReadOnly:
                current_val = p_comment.AsString()
                
                if current_val == text_to_handle:
                    # USUWANIE: Tekst -> "" oraz Reset koloru
                    p_comment.Set("")
                    view.SetElementOverrides(el_id, override_off)
                    count_removed += 1
                else:
                    # DODAWANIE: Tekst -> Komunikat oraz Kolor żółty
                    p_comment.Set(text_to_handle)
                    view.SetElementOverrides(el_id, override_on)
                    count_added += 1
    
    t.Commit()
    TaskDialog.Show("Status", "Dodano: {}\nUsunięto: {}".format(count_added, count_removed))