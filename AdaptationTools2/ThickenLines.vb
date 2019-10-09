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

        s = InputBox("Enter line thickness or 0 to scale in proportion to size of the Normal font style", "Thicken Lines", "0")
        If Not IsNumeric(s) Then Exit Sub

        ' floating shapes
        For i = 1 To r.ShapeRange.Count
            ThickenLinesRecurse(r.ShapeRange(i), CSng(s))
        Next i

        ' inline shapes
        For i = 1 To r.InlineShapes.Count
            ThickenLinesRecurse(r.InlineShapes(i), CSng(s))
        Next i

        'table borders
        ThickenBorders(r, s)


    End Sub

    Function ThickenBorders(sel As Word.Range, lineWeight As Single)

        If Not DEBUG Then On Error GoTo er

        Dim lw As Word.WdLineWidth
        Dim b As Word.Border
        Dim c As Word.Cell
        Dim t As Word.Table
        Dim i As Integer

        If lineWeight > 0 Then

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

        ' TODO use something other than Highlight

        ' NOTES: 
        ' Have to Highlight cells we want to process as .Cells collection doesn't work right
        ' Have to extend selection to make it work properly unless selection extends pass end of table

        Dim r As Word.Range
        Dim mixed As Boolean


        For Each t In sel.Tables
            r = t.Range
            mixed = False

            'If partial selection 
            If Not r.InRange(sel) Then
                ' Treat as mixed
                mixed = True
                ' Shrink working range
                If sel.Start > r.Start Then r.Start = sel.Start
                If sel.End < r.End Then r.End = sel.End
            Else
                ' Check for any mixed borders
                For Each b In r.Borders
                    If b.LineWidth = -1 Then
                        mixed = True
                        Exit For
                    End If
                Next
            End If



            If mixed Then
                ' slow path: check each border of each cell
                r.HighlightColorIndex = Word.WdColorIndex.wdTeal
                If r.End < r.Tables(1).Range.End Then r.End += 1 'Extend to include hidden Carriage Return

                For Each c In r.Cells
                    For Each b In c.Borders
                        If c.Range.HighlightColorIndex = Word.WdColorIndex.wdTeal _
                            And b.Visible Then b.LineWidth = lw
                    Next b
                Next

                r.End -= 1
                r.HighlightColorIndex = Word.WdColorIndex.wdNoHighlight

            Else
                ' fast path: set every border in whole table - will fail if any are mixed or partial selection
                For Each b In t.Borders
                    If b.Visible Then b.LineWidth = lw
                Next

            End If

        Next

        Exit Function

er:
        Select Case Err.Number
            Case Else
                Err.Raise(Err.Number, Err.Source, Err.Description)
        End Select
    End Function

    Sub ThickenLinesRecurse(shp As Word.Shape, lineWeight As Single)

        If Not DEBUG Then On Error GoTo er

        If shp.Type = Office.MsoShapeType.msoGroup Then
            ' recurse into groups
            For i = 1 To shp.GroupItems.Count
                ThickenLinesRecurse(shp.GroupItems(i), lineWeight)
            Next i

        Else
            If shp.Line.Visible = False Then Exit Sub

            If lineWeight > 0 Then
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
