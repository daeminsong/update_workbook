Sub UpdateConfirm()

Dim wkb As Excel.Workbook
Dim wks As Excel.Worksheet
Dim inquiry As String
Dim yearstring As String

    Set wkb = ThisWorkbook

    inquiry = MsgBox("Proceeds?", vbYesNo)
    If inquiry = vbNo Then
    Exit Sub
    End If
    
    wkb.Activate

    Application.DisplayAlerts = False
    Dim varLink As Variant
    For Each varLink In wkb.LinkSources(xlExcelLinks)
    Workbooks.Open varLink, , ReadOnly:=True
'    ActiveWorkbook.Close savechanges:=False
    Next varLink

    For Each varLink In wkb.LinkSources(xlExcelLinks)
'    Workbooks.Open varLink, , ReadOnly:=True
    If ActiveWorkbook.Name <> wkb.Name Then
    ActiveWorkbook.Close savechanges:=False
    End If
    Next varLink
    Application.DisplayAlerts = True

    MsgBox "Done"
    
End Sub

