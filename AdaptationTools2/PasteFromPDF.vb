Module PasteFromPDF

    'TODO buttons and test PasteOCR, FixOCR and FixOCR_Ligatures

    Sub PasteOCR()
        Dim objUndo As Word.UndoRecord

        objUndo = App.UndoRecord
        objUndo.StartCustomRecord("Paste OCR")

        PasteOCR_int()

        objUndo.EndCustomRecord()
    End Sub

    Private Sub PasteOCR_int()

        On Error GoTo er
        Dim r As Word.Range
        Dim svStart As Integer
        Dim svEnd As Integer

        ' Record selected range start and end as it gets erased by various modifications.
        r = App.ActiveDocument.Selection.Range
        svStart = r.Start
        svEnd = r.End

        'r.Paste
        r.PasteSpecial(, , , , Word.WdPasteDataType.wdPasteText)

        ' If pasting in a table, range might get set to start of table. Correct that.
        ' Order important! (change r.End before r.Start else r.End moves with r.Start)
        If r.End < svStart Then r.End = svEnd

        ' Reset r.start to where cursor was before paste to reselect pasted text
        r.Start = svStart

        r.Select()

        ' Fix-up OCR'd text
        FixOCR(r)

        Exit Sub

er:
        Select Case Err.Number
            Case 5342
                ' 5342 can't paste this datatype as plain text, use default
                r.Paste()
                Resume Next
            Case Else
                Err.Raise(Err.Number, Err.Source, Err.Description)
        End Select
    End Sub

    Private Sub FixOCR(r As Word.Range)
        ' Apply fix-ups for text that has been OCR'd or copied from PDF

        ' replace hyphens with minus
        ReplaceWrongDash(r)

        ' remove line breaks in the middle of sentences (from where the page wrapped on the original hardcopy)
        FindReplace("^11", "^13", r)

        ' remove fake bullets
        FindReplace("• ", "", r)

        ' Find: Paragraph mark (^13)
        'NOT preceded by punctuation (.!?),
        'NOT followed by CAPS or bullet characters (too many false positives with caps)
        'NOT followed by List item in the form e.g. a) or 1)
        ' Replace with a space
        FindReplace("([!\.\!\?])^13([!A-Z•][!?\)])", "\1 \2", r)

    End Sub

    Public Sub FixOCR_Ligatures()
        Dim objUndo As Word.UndoRecord

        objUndo = App.UndoRecord
        objUndo.StartCustomRecord("Fix OCR Ligatures")

        Dim rng As Word.Range

        rng = SelectionToRange()

        FixOCR_Ligatures_int(rng)

        rng.Select()

        objUndo.EndCustomRecord()
    End Sub

    Private Sub FixOCR_Ligatures_int(r As Word.Range)
        'Fix ligatures

        '"fl ",
        FindReplace("^13fl ", "^13fl", r) 'start of line
        FindReplace(" fl ", " fl", r) 'middle of line
        'can't do middle of word, too many false positives

        '"fi ",
        FindReplace("^13fi ", "^13fi", r) 'start of line
        FindReplace(" fi ", " fi", r) 'middle of line
        'can't do middle of word, too many false positives

        '"specifi c", etc
        FindReplace("specifi c", "specific", r)

        'add others as discovered...

    End Sub

End Module
