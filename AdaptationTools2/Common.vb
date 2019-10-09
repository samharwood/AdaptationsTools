
Module Common
    'TODO
    ' .Copy SavedValues


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

    Sub FindReplace(find As String, replace As String, ByRef rng As Word.Range, Optional recurse As Boolean = False)
        '
        ' recurse = true: loops until no more matches are found, use when the replacement text may create new matches.
        '           false: loops once and stops, even if the changes made create a new match for the pattern.
        '
        If Not DEBUG Then On Error GoTo er

        Dim matchesfound As Boolean

        If rng.Start = rng.End Then Exit Sub ' Empty range

        ' Set up criteria
        rng.Find.ClearFormatting()
        rng.Find.Replacement.ClearFormatting()

        With rng.Find
            .Text = find
            .Replacement.Text = replace
            .Forward = True
            .Format = False
            .MatchCase = False
            .MatchWholeWord = False
            .MatchAllWordForms = False
            .MatchSoundsLike = False
            .MatchWildcards = True
            .Wrap = Word.WdFindWrap.wdFindStop ' stop at end of the selection
        End With

        If recurse = True Then
            Do
                matchesfound = rng.Find.Execute(Replace:=Word.WdReplace.wdReplaceAll)
            Loop While matchesfound = True
        Else
            matchesfound = rng.Find.Execute(Replace:=Word.WdReplace.wdReplaceAll)
        End If

        Exit Sub

er:
        Select Case Err.Number
            Case Else
                MsgBox(Err.Number & vbCrLf & Err.Source & vbCrLf & Err.Description)
                'Err.Raise Err.Number, Err.Source, Err.Description
        End Select
    End Sub

    Sub ReplaceWrongDash(ByRef r As Word.Range)
        ' Replace unwanted types of dash/hyphen.
        ' 8722 unicode minus = GOOD DASH

        ' Unwanted character Unicode values
        ' 45 hyphen
        ' 8210 figure dash
        ' 8211 en dash
        ' 8212 em dash
        ' 8213 horizontal Bar
        Dim f As String

        f = "[" & ChrW(45) & ChrW(8210) & ChrW(8211) & ChrW(8212) & ChrW(8213) & "]"

        FindReplace(f, ChrW(8722), r, False)

    End Sub

    Sub ReplaceXwithMultiply(ByRef r As Word.Range)
        ' Replace the letter X with proper multiply sign
        FindReplace(" x ", " " & ChrW(&HD7) & " ", r, False)
    End Sub

    Sub ReplaceColonSpacesWithHardspaces(ByRef r As Word.Range)
        ' Surround Ratio colon with hardspace to keep together on page
        FindReplace(" : ", ChrW(&HA0) & ":" & ChrW(&HA0), r, False)
    End Sub

    Public Function GetVersion() As String
        If (System.Deployment.Application.ApplicationDeployment.IsNetworkDeployed) Then
            Dim ver As Version
            ver = System.Deployment.Application.ApplicationDeployment.CurrentDeployment.CurrentVersion
            Return String.Format("{0}.{1}.{2}.{3}", ver.Major, ver.Minor, ver.Build, ver.Revision)
        Else
            Return "Not Published"
        End If
    End Function

End Module
