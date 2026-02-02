Sub ExportMitarbeiterSheets()
    '=========================================================================
    ' Makro: Export einzelner Mitarbeiter-Sheets
    ' Trigger: Button auf "Übersicht" (G4)
    ' Funktion: Exportiert jedes Mitarbeiter-Sheet als separate Excel-Datei
    '           Behält Spalten A-H und T-U (Löscht I-S)
    '           Behält Dropdowns (Spalte T -> wird I)
    '           Formatiert auf eine Seite (FitToPagesWide = 1)
    '=========================================================================

    Dim ws As Worksheet
    Dim newWb As Workbook
    Dim exportPath As String
    Dim fso As Object
    
    ' Pfad für Export definieren (Unterordner "Mitarbeiter_Export" im gleichen Verzeichnis)
    exportPath = ThisWorkbook.Path & "\Mitarbeiter_Export\"
    
    ' Prüfen ob Ordner existiert, sonst erstellen
    If Dir(exportPath, vbDirectory) = "" Then
        MkDir exportPath
    End If

    ' Bildschirm-Aktualisierung ausschalten für bessere Performance
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    ' Durch alle Sheets iterieren
    For Each ws In ThisWorkbook.Worksheets
        ' Überspringe System-Sheets
        If ws.Name <> "Deckblatt" And ws.Name <> "Übersicht" And ws.Name <> "Projekt-Budget-Übersicht" Then

            ' Sheet in neue Arbeitsmappe kopieren
            ws.Copy
            Set newWb = ActiveWorkbook
            
            With newWb.Sheets(1)
                ' WICHTIG: Zuerst Formeln in Werte umwandeln!
                ' Wenn wir erst Spalten löschen, gehen Bezüge kaputt (#BEZUG!).
                ' Deshalb: Erst "einfrieren", dann löschen.
                .UsedRange.Value = .UsedRange.Value
            
                ' Jetzt Spalten I bis S löschen (11 Spalten)
                ' Damit rücken T und U (Original 20, 21) auf I und J (9, 10).
                ' Spalten A-H bleiben unberührt.
                .Range("I:S").Delete Shift:=xlToLeft
                
                ' Seitenlayout anpassen (Auf eine Seite breit)
                With .PageSetup
                    .Zoom = False
                    .FitToPagesWide = 1
                    .FitToPagesTall = False
                    .Orientation = xlLandscape
                End With
                
                ' Spaltenbreiten anpassen
                .Columns("I:I").ColumnWidth = 15 ' Ehemals T
                .Columns("J:J").ColumnWidth = 25 ' Ehemals U
            End With

            ' Speichern als normale XLSX
            newWb.SaveAs Filename:=exportPath & ws.Name & ".xlsx", FileFormat:=xlOpenXMLWorkbook
            newWb.Close SaveChanges:=False

        End If
    Next ws

    ' Aufräumen
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

    MsgBox "Export abgeschlossen!" & vbCrLf & "Dateien liegen in: " & exportPath, vbInformation, "Export erfolgreich"

End Sub
