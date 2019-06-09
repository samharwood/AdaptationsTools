Module ThickenLinesModule
    'TODO 
    ' .FEATURE: Increase thickness as a percentage of current width
    ' .FEATURE: Specify minimum width (for thinnest line in selection), thicken others proportionally 

    Sub ThickenLines()
        ' Thickens all Visible lines of Shapes and Table Borders to specified Width
        ' with option to choose a suitable width automatically (based on Normal style font size)

        Dim objUndo As Word.UndoRecord
        Dim r As Word.Range

        r = SelectionToRange()
        If r.Start = r.End Then Exit Sub

        objUndo = App.UndoRecord
        objUndo.StartCustomRecord("Thicken Lines")

        ThickenLines_int(r)
        r.Select()

        objUndo.EndCustomRecord()

    End Sub

    Sub ThickenLines_int(r As Word.Range)

        Dim s As String

        s = InputBox("Enter line thickness, leave blank for Auto", "Thicken Lines")
        If s = vbNullString Then s = -1 'Indicates to use Auto size later
        If Not IsNumeric(s) Then Exit Sub

        ' floating shapes
        For i = 1 To r.ShapeRange.Count
            ThickenLinesRecurse(r.ShapeRange(i), CSng(s))
        Next i

        ' inline shapes
        For i = 1 To r.InlineShapes.Count
            ThickenLinesRecurse(r.InlineShapes(i), CSng(s))
        Next i

        ' table borders
        For i = 1 To r.Tables.Count
            ThickenBordersRecurse(r.Tables(i), s)
        Next i

    End Sub

    Function ThickenBordersRecurse(t As Word.Table, lineWeight As Single)

        If Not DBG Then On Error GoTo er

        Dim lw As Word.WdLineWidth

        ' recurse
        For i = 1 To t.Tables.Count
            ThickenBordersRecurse(t.Tables(i), lineWeight)
        Next i

        If lineWeight >= 0 Then

            Select Case lineWeight
                Case Is >= 6 : lw = Word.WdLineWidth.wdLineWidth600pt
                Case Is >= 4 : lw = Word.WdLineWidth.wdLineWidth450pt
                Case Is >= 3 : lw = Word.WdLineWidth.wdLineWidth300pt
                Case Is >= 2 : lw = Word.WdLineWidth.wdLineWidth225pt
                Case Is >= 1.5 : lw = Word.WdLineWidth.wdLineWidth150pt
                Case Is >= 1 : lw = Word.WdLineWidth.wdLineWidth100pt
                Case Is >= 0.75 : lw = Word.WdLineWidth.wdLineWidth075pt
                Case Is >= 0.5 : lw = Word.WdLineWidth.wdLineWidth050pt
                Case Is >= 0 : lw = Word.WdLineWidth.wdLineWidth025pt
            End Select

        Else

            Select Case App.ActiveDocument.Styles("Normal").Font.Size
                Case Is >= 36 : lw = Word.WdLineWidth.wdLineWidth600pt
                Case Is >= 24 : lw = Word.WdLineWidth.wdLineWidth450pt
                Case Is >= 18 : lw = Word.WdLineWidth.wdLineWidth300pt
                Case Is >= 12 : lw = Word.WdLineWidth.wdLineWidth225pt
                Case Is >= 8 : lw = Word.WdLineWidth.wdLineWidth150pt
                Case Is >= 6 : lw = Word.WdLineWidth.wdLineWidth100pt
                Case Is >= 4 : lw = Word.WdLineWidth.wdLineWidth075pt
                Case Is >= 2 : lw = Word.WdLineWidth.wdLineWidth050pt
                Case Is >= 0 : lw = Word.WdLineWidth.wdLineWidth025pt
            End Select
        End If

        For i = -8 To -1 'border edges
            If t.Borders(i).Visible Then t.Borders(i).LineWidth = lw
        Next i


        Exit Function

er:
        Select Case Err.Number
            Case Else
                Err.Raise(Err.Number, Err.Source, Err.Description)
        End Select
    End Function

    Sub ThickenLinesRecurse(shp As Word.Shape, lineWeight As Single)

        If Not DBG Then On Error GoTo er

        If shp.Type = Office.MsoShapeType.msoGroup Then
            ' recurse into groups
            For i = 1 To shp.GroupItems.Count
                ThickenLinesRecurse(shp.GroupItems(i), lineWeight)
            Next i

        Else
            If shp.Line.Visible = False Then Exit Sub

            If lineWeight >= 0 Then
                shp.Line.Weight = lineWeight
            Else
                ' thicken lines to Normal/6
                ' 36=6, 24=4, 18=3, 12=2
                shp.Line.Weight = App.ActiveDocument.Styles("Normal").Font.Size / 6
            End If
        End If

        Exit Sub

er:
        Select Case Err.Number
            Case Else
                Err.Raise(Err.Number, Err.Source, Err.Description)
        End Select
    End Sub
End Module
