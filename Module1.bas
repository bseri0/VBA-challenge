Attribute VB_Name = "Module1"
Sub MaxValue()
    Range("J2") = WorksheetFunction.Max(Range("C2:C22711"))
End Sub

Sub GetCount()
    Dim LastRow As Long
    LastRow = Cells(Rows.Count, 1)
        'Range(Range("A5"), Range("A5").End(xlDown)).Select
End Sub

Sub Dosomething()
    Dim xSh As Worksheet
    Application.ScreenUpdating = False
    For Each xSh In Worksheets
        xSh.Select
        Call RunCode
    Next
    Application.ScreenUpdating = True
End Sub
Sub RunCode()
    'your code here
    MaxValue
End Sub
Sub Pull_Data()

    Dim FinalRow As Integer
    Dim Counter As Integer
    Dim curCell As Integer
    Dim a As Variant
    
    
    
    'FinalRow = Cells(Rows.Count, 1).End(xlUp).Row
    With ThisWorkbook.Sheets("A")
        For Counter = 2 To 20
            'curCell = Worksheets("A").Cells(Counter, 1)
        'If Abs(curCell.Value) < 0.01 Then curCell.Value = 0
            a = Cells(Counter, 1).Value.Copy
            Cells(Counter, 11).Value = a.PasteSpecial
        'Range(GetCount).Formula = "=(C2-F2)"
            Next Counter
        'Range(GetCount).Formula = "=(C2-F2)"
   End With
End Sub

Sub test()
    Dim i As Integer
    Dim WS As Worksheet
    Dim FinalRow As Integer
    
    
    Set WS = ActiveSheet
    Counter = 2
    FinalRow = Cells(Rows.Count, 1).End(xlUp).Row
    For i = 2 To FinalRow
        'If i = 163 Or i = 165 Or i = 174 Then counter = counter + 1
            If WS.Cells(i, 1).Value <> "" Then
                WS.Cells(Counter, 11).Value = WS.Cells(i, 1).Value
                WS.Cells(Counter, 12).Value = WS.Cells(i, 3).Value - WS.Cells(i, 6).Value '=(C2-F2)"
                Counter = Counter + 1
            End If
    Next i

    WS.Range("C1:C" & Counter - 1).Select
    Selection.Copy
End Sub


