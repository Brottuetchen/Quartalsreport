import os
import win32com.client
from pathlib import Path

def create_template():
    base_dir = Path(__file__).parent
    vba_file = base_dir / "webapp" / "export_macro.vba"
    output_file = base_dir / "webapp" / "template.xlsm"
    wb = None
    
    if output_file.exists():
        try:
            os.remove(output_file)
        except Exception as e:
            print(f"Konnte altes Template nicht löschen: {e}")
            return

    print("Starte Excel...")
    excel = None

    try:
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False

        wb = excel.Workbooks.Add()
        
        # Sheets bereinigen
        # Ziel: Nur "Übersicht" und "Projekt-Budget-Übersicht" sollen existieren.
        
        # Sicherstellen, dass "Übersicht" existiert
        try:
            ws_overview = wb.Sheets("Übersicht")
        except:
            ws_overview = wb.Sheets.Add()
            ws_overview.Name = "Übersicht"
            
        # Sicherstellen, dass "Projekt-Budget-Übersicht" existiert
        try:
            ws_budget = wb.Sheets("Projekt-Budget-Übersicht")
        except:
            ws_budget = wb.Sheets.Add()
            ws_budget.Name = "Projekt-Budget-Übersicht"
            
        # Alle anderen löschen
        for ws in wb.Sheets:
            if ws.Name != "Übersicht" and ws.Name != "Projekt-Budget-Übersicht":
                try:
                    ws.Delete()
                except:
                    pass

        # Reihenfolge sicherstellen: 
        # 1. Übersicht (Index 1 in VBA)
        # 2. Projekt-Budget-Übersicht (Index 2 in VBA)
        wb.Sheets("Übersicht").Move(Before=wb.Sheets(1))
        if wb.Sheets.Count > 1:
            wb.Sheets("Projekt-Budget-Übersicht").Move(After=wb.Sheets(1))
        
        # VBA Import
        if vba_file.exists():
            print("Importiere VBA...")
            with open(vba_file, 'r', encoding='utf-8') as f:
                code = f.read()
            
            # Modul hinzufügen
            mod = wb.VBProject.VBComponents.Add(1) # vbext_ct_StdModule
            mod.CodeModule.AddFromString(code)
        else:
            print(f"WARNUNG: VBA Datei nicht gefunden: {vba_file}")

        # Button erstellen auf Übersicht
        print("Erstelle Button...")
        ws_cover = wb.Sheets("Übersicht")
        btn = ws_cover.Buttons().Add(
            Left=ws_cover.Range("G4").Left,
            Top=ws_cover.Range("G4").Top,
            Width=250,
            Height=30
        )
        btn.OnAction = "ExportMitarbeiterSheets"
        btn.Caption = "Export Mitarbeiter"
        btn.Font.Bold = True
        btn.Font.Size = 11
        
        # Speichern
        print(f"Speichere Template nach {output_file}...")
        wb.SaveAs(str(output_file.absolute()), FileFormat=52) # xlOpenXMLWorkbookMacroEnabled
        
    except Exception as e:
        print(f"FEHLER: {e}")
    finally:
        if wb is not None:
            try:
                wb.Close(SaveChanges=False)
            except Exception:
                pass
        try:
            excel.Quit()
        except Exception:
            pass
        print("Fertig.")

if __name__ == "__main__":
    create_template()
