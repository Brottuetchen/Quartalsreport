Sub ExportMitarbeiterSheets()
    '=========================================================================
    ' Makro: Export einzelner Mitarbeiter-Sheets
    ' Funktion: Exportiert jedes Mitarbeiter-Sheet als separate Excel-Datei
    '           ohne die Spalten H-Q (Bonus-Anpassung bis Umsatz kumuliert)
    '=========================================================================

    Dim ws As Worksheet
    Dim newWb As Workbook
    Dim newWs As Worksheet
    Dim folderPath As String
    Dim fileName As String
    Dim lastRow As Long
    Dim lastCol As Long
    Dim sourceRange As Range
    Dim destCol As Long
    Dim col As Long

    ' Ordner auswählen
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "Wähle Zielordner für Mitarbeiter-Exports"
        .AllowMultiSelect = False
        If .Show = -1 Then
            folderPath = .SelectedItems(1)
        Else
            MsgBox "Export abgebrochen.", vbInformation
            Exit Sub
        End If
    End With

    ' Sicherstellen, dass der Pfad mit Backslash endet
    If Right(folderPath, 1) <> "\" Then
        folderPath = folderPath & "\"
    End If

    ' Bildschirm-Aktualisierung ausschalten für bessere Performance
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    ' Durch alle Sheets iterieren (außer Deckblatt und Projekt-Budget-Übersicht)
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name <> "Deckblatt" And ws.Name <> "Projekt-Budget-Übersicht" Then

            ' Neues Workbook erstellen
            Set newWb = Workbooks.Add
            Set newWs = newWb.Worksheets(1)
            newWs.Name = ws.Name

            ' Letzte verwendete Zeile und Spalte ermitteln
            lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
            lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

            ' Spalten kopieren: A-G (1-7) und T-U (20-21)
            destCol = 1

            ' Kopiere Spalten A-G (1-7)
            For col = 1 To 7
                ws.Columns(col).Copy
                newWs.Columns(destCol).PasteSpecial xlPasteAll
                newWs.Columns(destCol).PasteSpecial xlPasteColumnWidths
                destCol = destCol + 1
            Next col

            ' Kopiere Spalten T-U (20-21)
            For col = 20 To 21
                ws.Columns(col).Copy
                newWs.Columns(destCol).PasteSpecial xlPasteAll
                newWs.Columns(destCol).PasteSpecial xlPasteColumnWidths
                destCol = destCol + 1
            Next col

            ' Formatierung anpassen
            newWs.Cells.EntireColumn.AutoFit

            ' Data Validations kopieren (Dropdowns)
            Dim dv As Validation
            Dim srcCell As Range
            Dim destCell As Range

            On Error Resume Next
            ' Kopiere Position-Dropdown (B2)
            If ws.Range("B2").Validation.Type <> xlValidateInputOnly Then
                Set srcCell = ws.Range("B2")
                Set destCell = newWs.Range("B2")
                srcCell.Copy
                destCell.PasteSpecial xlPasteValidation
            End If

            ' Kopiere Rechnung-Dropdowns (Spalte T → Spalte H im neuen Sheet)
            Dim r As Long
            For r = 1 To lastRow
                Set srcCell = ws.Cells(r, 20) ' Spalte T
                If srcCell.Validation.Type <> xlValidateInputOnly Then
                    Set destCell = newWs.Cells(r, 8) ' Spalte H im neuen Sheet
                    srcCell.Copy
                    destCell.PasteSpecial xlPasteValidation
                End If
            Next r
            On Error GoTo 0

            ' Dateiname erstellen
            fileName = folderPath & ws.Name & ".xlsx"

            ' Speichern
            newWb.SaveAs fileName, FileFormat:=xlOpenXMLWorkbook
            newWb.Close SaveChanges:=False

        End If
    Next ws

    ' Aufräumen
    Application.CutCopyMode = False
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

    MsgBox "Export abgeschlossen!" & vbCrLf & vbCrLf & _
           "Mitarbeiter-Dateien wurden gespeichert in:" & vbCrLf & _
           folderPath, vbInformation, "Export erfolgreich"

End Sub
