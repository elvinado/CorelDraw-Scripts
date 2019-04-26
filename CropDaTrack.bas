Sub CheckY()
    Dim OrigSelection As ShapeRange
    Set OrigSelection = ActiveSelectionRange
    Debug.Print OrigSelection.Shapes(1).PositionY
End Sub

Sub CropDaTrack()
    Dim OrigSel As ShapeRange
    Dim Guide As ShapeRange
    Set OrigSel = ActiveSelectionRange
    
    GuideName = "BOT"
    
    Set Guide = ActivePage.Shapes.FindShapes(Name:=GuideName)
    'get from guide PositionY
    y1 = Guide.Shapes(1).PositionY
    'y1 = -83.5831771653543
    'Constants
    y2 = 0
    dxA = 3.36108792650918
    dxB = 0.99582677165354
    'Initialize for Track 1
    track = 1
    x1 = 0
    x2 = dxA

    For i = OrigSel.Count To 1 Step -1

        Debug.Print "Track=" & track & " x1=" & x1 & " x2=" & x2
        OrigSel(i).CustomCommand "Crop", "CropRectArea", x1, y1, x2, y2
        
        x1 = x2
        If track <> 1 Then
            x2 = x1 + dxA
        Else
            x2 = x1 + dxA + dxB
            x1 = x1 + dxB
        End If
        track = track + 1
    Next
    
End Sub

Sub FindObject()
    Dim Guide As ShapeRange
    'Set Guide = ActiveSelectionRange
    Set Guide = ActivePage.Shapes.FindShapes(Name:="BOT")
    Debug.Print Guide.Shapes(1).PositionY
End Sub
