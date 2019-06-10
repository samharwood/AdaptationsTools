Module Disabled
    Sub ResizeShapeTextAndLines(rng As Range, scaleby As Double, szL As Integer)
        'thicken lines and increase text size in textboxes
        On Error Resume Next
        System.Cursor = wdCursorWait

        For Each shp In rng.ShapeRange
            StatusBar = i & "/" & ActiveDocument.Shapes.Count & " shapes modified."

            shp.TextFrame.TextRange.Font.Size = shp.TextFrame.TextRange.Font.Size * scaleby
            If shp.TextFrame.TextRange.Font.Size < szL Then shp.TextFrame.TextRange.Font.Size = szL
            If shp.Line.Weight <> 0 Then shp.Line.Weight = shp.Line.Weight * scaleby

            i = i + 1
        Next shp
        StatusBar = i & "/" & ActiveDocument.Shapes.Count & " shapes modified."
        System.Cursor = wdCursorNormal
    End Sub


    ' Example of a work around to work with Multiple Selections.
    Sub changeNonContigCase()

        ' Find the non-contig selection
        If Selection.Font.Shading.BackgroundPatternColor = wdColorAutomatic Then
            Selection.Font.Shading.BackgroundPatternColor = whtcolor
        End If

        ' Find and process each range with .Font.Shading.BackgroundPatternColor = WhtColor
        ActiveDocument.Range.Select
        Selection.Collapse wdCollapseStart

    With Selection.find
            .Font.Shading.BackgroundPatternColor = whtcolor
            .Forward = True
            .Wrap = wdFindContinue

            Do While .Execute
                ' Do what you need
                Selection.Range.Case = wdTitleWord

                ' Reset shading as you go
                Selection.Font.Shading.BackgroundPatternColor = wdColorAutomatic

                ' Setup to find the next selection
                Selection.Collapse wdCollapseEnd
        Loop
        End With

    End Sub

    Sub SetCellMarginsForSelectedTables()
        ' Set cell margins for selected tables

        ' TODO: fixup and make configurable

        Dim t As Table

        For Each t In Selection.Tables

            t.TopPadding = CentimetersToPoints(0.2)
            t.BottomPadding = CentimetersToPoints(0.2)
            t.LeftPadding = CentimetersToPoints(0.2)
            t.RightPadding = CentimetersToPoints(0.2)

        Next t

    End Sub
End Module
