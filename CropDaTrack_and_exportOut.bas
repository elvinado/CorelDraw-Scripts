Attribute VB_Name = "Module1"
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

Sub SaveBitmap()
Dim LayerName, BitmapName As String
Dim regExLayer As New RegExp
Dim regExBitmap As New RegExp
Dim mc As MatchCollection
Dim lasfile, y_min, y_max, curve, x_min, x_max, scal, unit, outputName As String

Dim OrigSel As ShapeRange
Dim sh As Shape
Dim ex As ExportFilter
Set OrigSel = ActiveSelectionRange
Dim elem As Shape

With regExLayer
    .Global = True
    .Multiline = False
    .IgnoreCase = True
    .Pattern = "(.*.) (\d+) (\d+)"
End With

With regExBitmap
    .Global = True
    .Multiline = False
    .IgnoreCase = True
    .Pattern = "(\w+) (\w+) ([+-]?\d+(?:\.\d+)?) ([+-]?\d+(?:\.\d+)?) (\w+)"
End With

LayerName = OrigSel.Shapes(1).Layer.Name

Set mc = regExLayer.Execute(LayerName)

lasfile = mc(0).SubMatches(0)
y_min = mc(0).SubMatches(1)
y_max = mc(0).SubMatches(2)

For Each elem In OrigSel.Shapes
    BitmapName = elem.Name
    
    Set mc = regExBitmap.Execute(BitmapName)
    curve = mc(0).SubMatches(0)
    unit = mc(0).SubMatches(1)
    x_min = mc(0).SubMatches(2)
    x_max = mc(0).SubMatches(3)
    scal = mc(0).SubMatches(4)
    
    outputName = lasfile & "_" & curve & "_" & unit & "_" & y_min & "_" & y_max & "_" & x_min & "_" & x_max & "_" & scal & ".png"
    
    Set ex = elem.ConvertToBitmapEx.Bitmap.SaveAs("H:\All Alvin\py\LasDigitize\TEMP\" & outputName, cdrPNG)
    ex.Finish
Next
'
'BitmapName = OrigSel.Shapes(1).Name

'Set sh = OrigSel.Shapes(1).ConvertToBitmapEx
'Set ex = sh.Bitmap.SaveAs("H:\All Alvin\py\LasDigitize\TEMP\File2.png", cdrPNG)
'ex.Finish

End Sub
