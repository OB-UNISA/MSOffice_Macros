Attribute VB_Name = "Module1"
Sub Progress_Bar()
    On Error Resume Next
    With ActivePresentation
        s_padding = 3
        s_heigth = 6
        s_left_div = 6
        s_left = (.PageSetup.SlideWidth / s_left_div) * (s_left_div - 1) - s_padding
        For X = 1 To .Slides.Count
            .Slides(X).Shapes("PB").Delete
            .Slides(X).Shapes("PB_LINE").Delete
            Set s = .Slides(X).Shapes.AddShape(msoShapeRoundedRectangle, s_left, s_padding, (X * .PageSetup.SlideWidth / .Slides.Count) / s_left_div, s_heigth)
            s.Fill.ForeColor.RGB = RGB(153, 175, 214)
            s.Line.Transparency = 1
            s.Name = "PB"
            
            Set s2 = .Slides(X).Shapes.AddShape(msoShapeRoundedRectangle, s_left, s_padding, .PageSetup.SlideWidth / s_left_div, s_heigth)
            s2.Fill.Transparency = 1
            s2.Line.Transparency = 0.5
            s2.Weight = 1
            s2.Name = "PB_LINE"
        Next X:
    End With
End Sub
Sub Delete_Progress_Bar()
    On Error Resume Next
    With ActivePresentation
        For X = 1 To .Slides.Count
            .Slides(X).Shapes("PB").Delete
            .Slides(X).Shapes("PB_LINE").Delete
        Next X:
    End With
End Sub
