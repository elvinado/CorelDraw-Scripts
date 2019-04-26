Attribute VB_Name = "Module1"
Sub alvin()
    Dim OrigSel As ShapeRange
    Set OrigSel = ActiveSelectionRange
    Dim ht As Double
    ht = 0
    For i = 1 To OrigSel.Count
        OrigSel(i).Name = i
        'MsgBox OrigSel(i).OriginalHeight
        ht = OrigSel(i).OriginalHeight + ht '- 0.083
        'MsgBox ht
        OrigSel(i).Move 0, -ht
    Next
    
End Sub

Sub alvin2()
    Dim OrigSel As ShapeRange
    Set OrigSel = ActiveSelectionRange
    Dim ht, wt As Double
    Dim total, column, row As Integer
    Dim i, j, k As Integer
    
    'The only manual input needed
    column = 2
    'Get total images selected
    total = OrigSel.Count
    'Sampling first image as Height and Width
    ht = OrigSel(1).OriginalHeight
    wt = OrigSel(1).OriginalWidth
    'Loop for all images selected
    For i = 1 To total
        'Getting j,k position coordinate
        If 0 = i Mod column Then
            j = column
            k = (i / column)
        Else:
            j = i Mod column
            k = (i \ column) + 1
        End If
        'Renaming based on i,j,k
        OrigSel(i).Name = "No: " & i & " C: " & j & " R: " & k
        'Moving based on i,j,k
        OrigSel(i).Move wt * j, ht * -k
    Next
    
End Sub

Sub test()
    total = 9
    column = 3
    k = 0
    For i = 1 To total
        If 0 = i Mod column Then
            j = column
            k = (i / column)
        Else:
            j = i Mod column
            k = (i \ column) + 1
        End If
        MsgBox "Count:" & i & " Column:" & j & " Row:" & k
    Next
End Sub

