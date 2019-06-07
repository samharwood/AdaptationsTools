Attribute VB_Name = "Internals"
Option Private Module

Const DBG As Boolean = True


' Function code
' Separated from Macro names so Toolbar Macro names don't have to change when refactoring code.

Sub AutoExec()
'
' AutoExec Macro
' Runs when Word starts
'
End Sub

Sub ReplaceNonBreakingSpaces_int(r As Range)
    '
    'Replace the last space in each paragraph with a non-breaking space,
    'so you don't end up with a single hanging word on a new line.
    '
    On Error Resume Next
    Dim para As Paragraph

    For Each para In r.Paragraphs
        For i = para.Range.Words.Last.Start - 1 To para.Range.Start Step -1
            If r.Characters(i) = Chr(32) Then
                r.Characters(i) = Chr(160)
                'DEBUG r.Range.Characters(i).HighlightColorIndex = wdYellow
                Exit For
            End If
        Next i
    Next para

End Sub

Sub SelectTestMath()
    ' WORD BUG:
    ' STEPS:    Select part of of an OMath equation. (>0 chars and < whole equation)
    '           Reference any property of the OMath object of the Selection
    ' BUG:      Word freezes. - True as of April 2019
    ' FIX:      Expand Selection to whole equation before using in code.
    ' BONUS:    Creates preferred behaviour of expanding a collapsed/partial selection to whole equation.
    
    If Selection.OMaths.Count = 0 Then Exit Sub
    
    ' If entire Selection is contained within an OMath object
    If Selection.InRange(Selection.OMaths(1).Range) Then
        ' Expand the Selection to the containing OMath object
        Selection.Start = Selection.OMaths(1).ParentOMath.Range.Start
        Selection.End = Selection.OMaths(1).ParentOMath.Range.End
    End If
    
End Sub


Function SelectionToRange() As Range
    
    ' Return selection as range
    
    Dim rng As Range
    
    SelectTestMath
    
    If Selection.Start >= Selection.End Then
        
        ' return no selection
        Set rng = ActiveDocument.Range(Start:=Selection.Start, End:=Selection.Start)
        
        ' if no selection return whole document
        'Set rng = ActiveDocument.Range(Start:=ActiveDocument.Range.Start, End:=ActiveDocument.Range.End - 1)
        
    Else

        ' Do not allow last character of document to be part of selection
        ' (last char cannot be deleted and causes replaceall to loop infinitely)
        If Selection.End = ActiveDocument.Range.End Then
            Set rng = ActiveDocument.Range(Start:=Selection.Start, End:=Selection.End - 1)
        Else
            Set rng = ActiveDocument.Range(Start:=Selection.Start, End:=Selection.End)
        End If
        
    End If
    
    Set SelectionToRange = rng
End Function


Sub ThickenLines_int(r As Range)
    
    Dim s As String
    
    s = InputBox("Enter line thickness", "Thicken Lines")
    If Not IsNumeric(s) Then Exit Sub
        
    ' floating shapes
    For i = 1 To r.ShapeRange.Count
        ThickenLinesRecurse r.ShapeRange(i), CSng(s)
    Next i
    
    ' inline shapes
    For i = 1 To r.InlineShapes.Count
        ThickenLinesRecurse r.InlineShapes(i), CSng(s)
    Next i
    
    ' table borders
    For i = 1 To r.Tables.Count
        ThickenBordersRecurse r.Tables(i), CSng(s)
    Next i
    
End Sub



Function ThickenBordersRecurse(t As Table, lineWeight As Single)

    If Not DBG Then On Error GoTo er
    
    ' recurse
    For i = 1 To t.Tables.Count
        ThickenBordersRecurse t.Tables(i), lineWeight
    Next i
    
    Select Case lineWeight
        Case Is >= 6: lw = wdLineWidth600pt
        Case Is >= 4: lw = wdLineWidth450pt
        Case Is >= 3: lw = wdLineWidth300pt
        Case Is >= 2: lw = wdLineWidth225pt
        Case Is >= 1.5: lw = wdLineWidth150pt
        Case Is >= 1: lw = wdLineWidth100pt
        Case Is >= 0.75: lw = wdLineWidth075pt
        Case Is >= 0.5: lw = wdLineWidth050pt
        Case Is >= 0: lw = wdLineWidth025pt
    End Select

'    Select Case ActiveDocument.Styles("Normal").Font.Size
'        Case Is >= 36: lw = wdLineWidth600pt
'        Case Is >= 24: lw = wdLineWidth450pt
'        Case Is >= 18: lw = wdLineWidth300pt
'        Case Is >= 12: lw = wdLineWidth225pt
'        Case Is >= 8: lw = wdLineWidth150pt
'        Case Is >= 6: lw = wdLineWidth100pt
'        Case Is >= 4: lw = wdLineWidth075pt
'        Case Is >= 2: lw = wdLineWidth050pt
'        Case Is >= 0: lw = wdLineWidth025pt
'    End Select
    
    
    For i = -8 To -1 'border edges
        If t.Borders(i).Visible Then t.Borders(i).linewidth = lw
    Next i

    
    Exit Function
    
er:
    Select Case Err.number
        Case Else
            Err.Raise Err.number, Err.Source, Err.Description
    End Select
End Function

Function ThickenLinesRecurse(shp As Shape, lineWeight As Single)

    If Not DBG Then On Error GoTo er

    If shp.Type = msoGroup Then
        ' recurse into groups
        For i = 1 To shp.GroupItems.Count
            ThickenLinesRecurse shp.GroupItems(i), lineWeight
        Next i
        
    Else
        If shp.Line.Visible <> msoFalse Then
            
            ' thicken lines to Normal/6
            ' 36=6, 24=4, 18=3, 12=2
            ' shp.Line.weight = ActiveDocument.Styles("Normal").Font.Size / 6
            ' or
            shp.Line.Weight = lineWeight
            
        End If
    End If
    
    Exit Function
    
er:
    Select Case Err.number
        Case Else
            Err.Raise Err.number, Err.Source, Err.Description
    End Select
End Function


Function FindReplace(find As String, replace As String, ByRef rng As Range, Optional recurse As Boolean = False)
    '
    ' recurse = true: loops until no more matches are found, used when the replacement text may create new matches.
    '           false: loops once and stops, even if the changes made create a new match for the pattern.
    '
    If Not DBG Then On Error GoTo er
    
    If rng.Start = rng.End Then Exit Function ' Empty range
        
    ' Set up criteria
    rng.find.ClearFormatting
    rng.find.Replacement.ClearFormatting
    With rng.find
        .Text = find
        .Replacement.Text = replace
        .Forward = True
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
        .Wrap = wdFindStop ' stop at end of the selection
    End With
        
    If recurse = True Then
        Do
            matchesfound = rng.find.Execute(replace:=wdReplaceAll)
        Loop While matchesfound = True
    Else
        matchesfound = rng.find.Execute(replace:=wdReplaceAll)
    End If
        
    Exit Function

er:
    Select Case Err.number
        Case Else
            MsgBox Err.number & vbCrLf & Err.Source & vbCrLf & Err.Description
            'Err.Raise Err.Number, Err.Source, Err.Description
    End Select
End Function



Sub FixOCR_Ligatures_int(r As Range)
'Fix ligatures

    '"fl ",
    FindReplace "^13fl ", "^13fl", r 'start of line
    FindReplace " fl ", " fl", r 'middle of line
    'can't do middle of word, too many false positives
    
    '"fi ",
    FindReplace "^13fi ", "^13fi", r  'start of line
    FindReplace " fi ", " fi", r  'middle of line
    'can't do middle of word, too many false positives
    
    '"specifi c", etc
    FindReplace "specifi c", "specific", r
    
    'add others as discovered...

End Sub


Sub FixOCR(r As Range)
' Apply fix-ups for text that has been OCR'd or copied from PDF
    
' replace hyphens with minus
    ReplaceWrongDash r
    
' remove line breaks in the middle of sentences (from where the page wrapped on the original hardcopy)
    FindReplace "^11", "^13", r
    
' remove fake bullets
    FindReplace "• ", "", r
    
' Find: Paragraph mark (^13)
    'NOT preceded by punctuation (.!?),
    'NOT followed by CAPS or bullet characters (too many false positives with caps)
    'NOT followed by List item in the form e.g. a) or 1)
    ' Replace with a space
    FindReplace "([!\.\!\?])^13([!A-Z•][!?\)])", "\1 \2", r

    
End Sub


Function PasteOCR_int()
    
    On Error GoTo er
    Dim r As Range
    
    Set r = Selection.Range
    
    'record selected range start and end as it gets erased by various modifications.
    svStart = r.Start
    svend = r.End
    
    'r.Paste
    r.PasteSpecial , , , , wdPasteText
    
    ' If pasting in a table, range might get set to start of table. Correct that.
    ' Order important! (change r.End before r.Start else r.End moves with r.Start)
    If r.End < svStart Then r.End = svend

    ' reset r.start to where cursor was before paste to reselect pasted text
    r.Start = svStart
    
    r.Select
    
        FixOCR r

    Exit Function

er:
    Select Case Err.number
        Case 5342
        ' 5342 can't paste this datatype as plain text, use default
            r.Paste
            Resume Next
        Case Else
            Err.Raise Err.number, Err.Source, Err.Description
    End Select
End Function


Function ReplaceWrongDash(ByRef r As Range)
' Replace unwanted types of dash/hyphen.
' 8722 unicode minus = GOOD DASH

    ' Unwanted character Unicode values
    ' 45 hyphen
    ' 8210 figure dash
    ' 8211 en dash
    ' 8212 em dash
    ' 8213 horizontal Bar
    
    FindReplace ChrW(45), ChrW(8722), r, False
    FindReplace ChrW(8210), ChrW(8722), r, False
    FindReplace ChrW(8211), ChrW(8722), r, False
    FindReplace ChrW(8212), ChrW(8722), r, False
    FindReplace ChrW(8213), ChrW(8722), r, False

End Function


