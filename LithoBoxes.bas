Attribute VB_Name = "RecordedMacros"
Sub CreateBoxes()
    TopArray = Array(3041, 3093, 3267, 3580, 3921, 3956, 4640, 4785, 4851, 4961, 5076, 5211, 5273, 5407, 5455, 5468, 5536, 5948, 5982, 6004, 6420, 6484, 6629, 6817, 6903, 6979, 7016, 7152, 7211, 7294, 7677, 7823, 7839, 7877, 7928, 8058, 8146, 8166, 8199, 8220)
    BotArray = Array(3067, 3102, 3318, 3610, 3940, 3962, 4643, 4788, 4863, 4981, 5084, 5220, 5293, 5438, 5458, 5471, 5556, 5958, 6000, 6016, 6441, 6509, 6660, 6822, 6906, 6985, 7060, 7182, 7231, 7304, 7683, 7839, 7870, 7889, 7938, 8066, 8160, 8173, 8220, 8232)
    LitArray = Array("Water", "Water", "Water", "Water", "Oil", "Water", "Water", "Water", "Water", "Gas", "Gas", "Oil", "HC_Oil?", "Water", "Water", "Water", "Water", "Gas", "Gas", "Gas", "Water", "Gas", "Gas", "Gas", "HC", "Gas", "Gas", "Gas", "Gas", "Gas", "HC", "Gas", "Gas", "Gas", "HC", "Gas", "Gas", "Water", "Gas", "Gas")
    For i = LBound(TopArray) To UBound(TopArray)
        a = 1500 / 25.4
        b = TopArray(i) / -25.4
        c = 1750 / 25.4
        d = BotArray(i) / -25.4
        Col = LitArray(i)
        Set Rect = ActiveLayer.CreateRectangle(a, b, c, d)
        If Col = "Water" Then
            Rect.Fill.ApplyUniformFill CreateRGBColor(0, 0, 255)
        ElseIf Col = "Oil" Then
            Rect.Fill.ApplyUniformFill CreateRGBColor(0, 255, 0)
        ElseIf Col = "Gas" Then
            Rect.Fill.ApplyUniformFill CreateRGBColor(255, 0, 0)
        Else
            Rect.Fill.ApplyUniformFill CreateRGBColor(255, 200, 0)
        End If
    Next i
End Sub
