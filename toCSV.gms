Private Sub CommandButton1_Click()
    Dim s As Shape
    
    ActiveDocument.Unit = cdrMillimeter
    
    TextBox1.Value = "ID" + vbTab + "Name" + vbTab + "X" + vbTab + "Y" + vbTab + "Angle" + vbLf
    
    For Each s In ActiveSelectionRange.Shapes
        If s.Type = cdrGroupShape Then
            describeCircle s
        End If
    Next s
    
    txt = TextBox1.Value
    
End Sub

Sub describeCircle(s As Shape)
    Dim txt As String
    
    txt = ""
    
    'optional inclusion of item ID and name
    txt = txt + Trim(Str(s.StaticID)) + vbTab + Trim(s.Name) + vbTab
    
    'X and Y coordinates
    txt = txt + Trim(Str(Round(s.PositionX + s.SizeWidth / 2, 3) * 1000)) + vbTab + Trim(Str(Round(s.PositionY - s.SizeHeight / 2, 3) * 1000)) + vbTab + Trim(Str(Round(s.RotationAngle - 90)))
    
    'optional addition of width and height
    'txt = txt + vbTab + Trim(Str(Round(s.SizeWidth, 3))) + vbTab + Trim(Str(Round(s.SizeHeight, 3)))
    
    TextBox1.Value = TextBox1.Value + txt + vbLf
End Sub
