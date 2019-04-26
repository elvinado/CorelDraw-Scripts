Attribute VB_Name = "Module1"
Sub alvin()
    Dim OrigSel As ShapeRange
    Set OrigSel = ActiveSelectionRange
    Dim ht As Double

    For i = 1 To OrigSel.Count
        OrigSel(i).Name = i
        'MsgBox OrigSel(i).OriginalHeight
        ht = OrigSel(i).OriginalHeight + ht - 0.083335
        'MsgBox ht
        OrigSel(i).Move 0, -ht
    Next
    MsgBox "Done"
    
End Sub
