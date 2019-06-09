
Module AdaptationsTools
    'TODO
    ' .Adaptations tools
    ' .Design Ribbon


    Function SelectionToRange() As Word.Range
        ' Return current Selection as a Range with some tweaks

        Dim rng As Word.Range
        Dim Sel As Word.Selection
        Dim ActiveDocument As Word.Document
        Sel = App.Selection
        ActiveDocument = App.ActiveDocument

        SelectTestMath(Sel) 'Workaround for bug

        If Sel.Start >= Sel.End Then

            ' return no Sel
            rng = ActiveDocument.Range(Start:=Sel.Start, End:=Sel.Start)

            ' if no Sel return whole document
            'rng = ActiveDocument.Range(Start:=ActiveDocument.Range.Start, End:=ActiveDocument.Range.End - 1)

        Else

            ' Do not allow last character of document to be part of Sel
            ' (last char cannot be deleted and causes replaceall to loop infinitely)
            If Sel.End = ActiveDocument.Range.End Then
                rng = ActiveDocument.Range(Start:=Sel.Start, End:=Sel.End - 1)
            Else
                rng = ActiveDocument.Range(Start:=Sel.Start, End:=Sel.End)
            End If

        End If

        SelectionToRange = rng
    End Function

    Sub SelectTestMath(Sel As Word.Selection)
        ' WORD BUG WORKAROUND
        ' STEPS:    Select part of of an OMath equation. (>0 chars and < whole equation)
        '           Reference any property of the OMath object of the Sel
        ' BUG:      Word freezes. - True as of April 2019
        ' FIX:      Expand Sel to whole equation before using in code.
        ' BONUS:    Creates preferred behaviour of expanding a collapsed/partial Sel to whole equation.

        If Sel.OMaths.Count = 0 Then Exit Sub

        ' If entire Sel is contained within an OMath object
        If Sel.InRange(Sel.OMaths(1).Range) Then
            ' Expand the Sel to the containing OMath object
            Sel.Start = Sel.OMaths(1).ParentOMath.Range.Start
            Sel.End = Sel.OMaths(1).ParentOMath.Range.End
        End If

    End Sub


End Module
