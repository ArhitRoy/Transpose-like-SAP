Attribute VB_Name = "Module1"
Sub industry_project()

    Dim i As Byte, bd As Worksheet, attr As Integer, entr As Integer, k As Long, f As Long

    For i = 1 To Worksheets.Count
        If Worksheets(i).Cells(1, 1).Value = "Parent" Then
            Worksheets(i).Delete
        End If
        Application.DisplayAlerts = False
    Next i

    Worksheets.Add after:=Worksheets(Worksheets.Count)
    ActiveSheet.Name = "Result"

    Worksheets("Result").Activate

    Range("A1").Value = "Parent"
    Range("A1").Font.Bold = True

    Range("B1").Value = "BOM"
    Range("B1").Font.Bold = True

    Range("C1").Value = "Units"
    Range("C1").Font.Bold = True

    Set bd = Worksheets("BOM Detail")
    bd.Activate

    attr = bd.Range("B2", Range("B2").End(xlToRight)).Count
    entr = bd.Range("A10", Range("A10").End(xlDown)).Count
    f = attr * CLng(entr)

    Worksheets("Result").Activate
    j = 1

    For k = 2 To f Step attr
    
        ' copying the parent
        bd.Range("A" & j + 9).Copy Range("A" & k)

        ' copying the BOM
        bd.Range("B2:AH2").Copy
        Range("B" & k).PasteSpecial Transpose:=True
    
        ' copying the units
        bd.Range(("B" & j + 9), ("AH" & j + 9)).Copy
        Range("C" & k).PasteSpecial Transpose:=True

        ' autofill
        Range("A" & k).AutoFill Range(("A" & k), ("A" & Range("B" & k).End(xlDown).Row)), xlFillCopy
        
        ' setting the interval
        j = j + 1
    
    Next k

    Range("A1").CurrentRegion.Font.Color = vbBlack
    Range("A1").CurrentRegion.Font.Name = "Times New Roman"
    Range("A1").CurrentRegion.Font.Size = 10
    Range("A1").CurrentRegion.HorizontalAlignment = xlLeft

    Range("A1").Font.Size = 12
    Range("B1").Font.Size = 12
    Range("C1").Font.Size = 12

End Sub
