Module Maths

    Sub PrepBrailleMaths()
        ' Fix up Equation Editor Braille Math ready for converting to MathType for Brailling.
        ' -Sets Inline type, as MathType conversion inserts Tab chars when converting from Display type
        ' -Convert 'Text Maths' back to 'Normal Maths'
        ' -Remove all spaces
        ' -Convert all similar looking dash/hyphens characters into proper minus signs.
        On Error GoTo er

        Dim objUndo As Word.UndoRecord
        Dim r As Word.Range


        objUndo = App.UndoRecord
        objUndo.StartCustomRecord("Convert Math for Brailling")

        r = SelectionToRange()

        PrepBrailleMaths_int(r)

        objUndo.EndCustomRecord()
        Exit Sub

er:

        Select Case Err.Number
            Case Else
                MsgBox("PrepBrailleMaths " & vbCrLf & Err.Number & vbCrLf & Err.Description)
        End Select
    End Sub

    Sub PrepLargePrintMaths()
        ' Converts Equations in document to Text (so font can be changed to Arial)
        ' Match (certain elements of) Normal style.
        ' Enlarge sub/supscripts in normal text and in Equations
        '
        On Error GoTo er

        Dim objUndo As Word.UndoRecord
        Dim r As Word.Range

        r = SelectionToRange()
        If r.Start = r.End Then Exit Sub

        objUndo = App.UndoRecord
        objUndo.StartCustomRecord("Convert Math to Large Print")

        PrepLargePrintMaths_int(r)

        objUndo.EndCustomRecord()
        Exit Sub

er:

        Select Case Err.Number
            Case Else
                MsgBox("PrepLargePrintMaths " & vbCrLf & Err.Number & vbCrLf & Err.Description)
        End Select
    End Sub

    Private Sub PrepBrailleMaths_int(r As Word.Range)

        On Error GoTo er

        Dim shp As Word.Shape

        'replace hyphens with minus
        ReplaceWrongDash(r)

        'process Math in main text
        OMaths_to_Braille(r.OMaths)

        'process Math in Shapes
        For Each shp In r.ShapeRange

            'causes err 5917 if shape does not support attached text
            OMaths_to_Braille(shp.TextFrame.TextRange.OMaths)

        Next shp

        Exit Sub

er:
        Select Case Err.Number
            Case 5917
                ' shape does not support attached text
                Resume Next
            Case Else
                Err.Raise(Err.Number, Err.Source, Err.Description)
        End Select
    End Sub

    Private Sub PrepLargePrintMaths_int(r As Word.Range)

        On Error GoTo er

        App.System.Cursor = Word.WdCursorType.wdCursorWait


        ' process normal text
        OMaths_to_TextMath(r.OMaths)

        'replace hyphens with minus (do after converting to textmath)
        ReplaceWrongDash(r)

        IncreaseSuperscripts(r)

        ' Increase Math in Shapes
        For i = 1 To r.ShapeRange.Count
            IncreaseMathInShapes(r.ShapeRange(i))
        Next i


        ' clean up
        App.System.Cursor = Word.WdCursorType.wdCursorNormal

        Exit Sub
er:
        Select Case Err.Number
            Case Else
                App.System.Cursor = Word.WdCursorType.wdCursorNormal
                Err.Raise(Err.Number, Err.Source, Err.Description)
        End Select
    End Sub

    Private Sub IncreaseOMathSupercripts(o As Word.OMath)
        'further enlarge fractions and sub/supscripts in Word.OMath objects

        On Error GoTo er

        Dim f As Word.OMathFunction
        Dim a As Word.OMath
        Dim g As Word.OMathFunction
        Dim j As Word.OMath
        Dim c As Word.Range
        Dim i As Integer
        Dim supNormal, supAdd, supsize As Integer


        'size + 6 + (2 for every 18pt)
        supNormal = App.ActiveDocument.Styles("Normal").Font.Size
        supAdd = Int(supNormal / 18) * 2 '+2 for every 18pt
        supsize = supNormal + 6 + supAdd


        For i = 1 To o.Functions.Count

            f = o.Functions(i)

            'Fractions in Functions
            ' Can't increase Fraction size independent of rest of Function,
            ' so have to just increase Num/Den which makes fraction line spacing visually untidy
            For Each c In f.Range.Characters
                c.OMaths(1).Functions(1).Frac.Den.Range.Font.Size = supsize
                c.OMaths(1).Functions(1).Frac.Num.Range.Font.Size = supsize
            Next c

            f.ScrSup.Sup.Range.Font.Size = supsize 'Superscripts
            f.Frac.Parent.Range.Font.Size = supsize 'Fractions (that stand alone?)
            f.Rad.Deg.Range.Font.Size = supsize 'Radicals with degree e.g. cube root

        Next i
        Exit Sub

er:
        Select Case Err.Number
            Case -2147467259 'element does not exist
                Resume Next
            Case Else
                Err.Raise(Err.Number, Err.Source, Err.Description)
        End Select
    End Sub

    Private Sub IncreaseSuperscripts(r As Word.Range)
        ' further increase sub/supscripts
        ' normal size + 6 + (2 for every 18pt)
        ' this function is very slow when called repeatedly. Therefore it is not used within Shapes code.

        Dim supsize As Integer
        Dim s As Integer
        Dim a As Integer

        s = App.ActiveDocument.Styles("Normal").Font.Size
        a = Int(s / 18) * 2 '+2 for every 18pt
        supsize = s + 6 + a

        'set up selection find defaults
        r.Find.Text = ""
        r.Find.ClearFormatting()
        r.Find.Replacement.Text = ""
        r.Find.Replacement.ClearFormatting()

        With r.Find
            .Forward = True
            .Format = False
            .MatchCase = False
            .MatchWholeWord = False
            .MatchAllWordForms = False
            .MatchSoundsLike = False
            .MatchWildcards = False
            .Wrap = Word.WdFindWrap.wdFindStop
        End With




        'make superscripts bigger
        r.Find.Font.Superscript = True
        r.Find.Replacement.Font.Size = supsize
        r.Find.Execute(Replace:=Word.WdReplace.wdReplaceAll)

        'same again but for subscript.
        r.Find.Font.Subscript = True
        r.Find.Replacement.Font.Size = supsize
        r.Find.Execute(Replace:=Word.WdReplace.wdReplaceAll)

        'superscript2 symbol make normal text with superscript
        r.Find.Font.Subscript = False
        r.Find.Text = ChrW(&HB2)
        r.Find.Replacement.Text = "2"
        r.Find.Replacement.Font.Superscript = True
        r.Find.Execute(Replace:=Word.WdReplace.wdReplaceAll)

        'superscript3 symbol
        r.Find.Text = ChrW(&HB3)
        r.Find.Replacement.Text = "3"
        r.Find.Replacement.Font.Superscript = True
        r.Find.Execute(Replace:=Word.WdReplace.wdReplaceAll)

        'superscript0 symbol
        r.Find.Text = ChrW(&H2070)
        r.Find.Replacement.Text = 0
        r.Find.Replacement.Font.Superscript = True
        r.Find.Execute(Replace:=Word.WdReplace.wdReplaceAll)

        'superscript4-9 symbol
        For i = 4 To 9
            r.Find.Text = ChrW(&H2070 + i)
            r.Find.Replacement.Text = i
            r.Find.Replacement.Font.Superscript = True
            r.Find.Execute(Replace:=Word.WdReplace.wdReplaceAll)
        Next i

    End Sub

    Private Sub IncreaseMathInShapes(shp As Word.Shape)

        On Error GoTo er

        If shp.Type = Office.MsoShapeType.msoGroup Then
            ' recurse into groups
            For i = 1 To shp.GroupItems.Count
                IncreaseMathInShapes(shp.GroupItems(i))
            Next i

        Else
            ' increase text in range
            If shp.TextFrame.HasText Then
                ' Only do Word.OMath. Because IncreaseSuperscripts is very slow on many shapes
                OMaths_to_TextMath(shp.TextFrame.TextRange.OMaths)
                IncreaseSuperscripts(shp.TextFrame.TextRange)
            End If
        End If

        Exit Sub

er:
        Select Case Err.Number
            Case Else
                Err.Raise(Err.Number, Err.Source, Err.Description)
        End Select
    End Sub

    Private Sub IncreaseOMathElements(o As Word.OMath)
        'enlarge certain smaller Word.OMath objects


        Dim f As Word.OMathFunction
        Dim supNormal, supAdd, supsize As Integer

        'size + 6 + (2 for every 18pt)
        supNormal = App.ActiveDocument.Styles("Normal").Font.Size
        supAdd = Int(supNormal / 18) * 2 '+2 for every 18pt
        supsize = supNormal + 6 + supAdd

        On Error Resume Next

        For Each f In o.Functions

            Select Case f.Type
                Case Word.WdOMathFunctionType.wdOMathFunctionFrac 'fractions
                    f.Range.Font.Size = supsize
                Case Word.WdOMathFunctionType.wdOMathFunctionDelim 'brackets
                    IncreaseOMathElements(f.Frac)
            End Select

        Next f

    End Sub

    Private Sub OMaths_to_Braille(oms As Word.OMaths)
        ' -Convert 'Text Math' to 'Normal Math'
        ' -Remove spacing
        ' -Make Inline as MathType conversion inserts Tab chars when converting from Display type

        Dim r As Word.Range
        Dim o As Word.OMath

        'create a range to edit the document as o and c are readonly
        r = App.ActiveDocument.Range

        'OMath Objects
        For Each o In oms

            'Convert Text Math To Proper Math
            o.ConvertToMathText()

            ' remove all spaces from maths objects
            FindReplace(ChrW(&H20), "", o.Range, False)

            'Make Inline, as MathType conversion inserts Tab chars when converting from Display type
            o.Type = Word.WdOMathType.wdOMathInline
            o.Range.Bold = False
        Next o

    End Sub

    Private Sub OMaths_to_TextMath(oms As Word.OMaths)
        ' -Convert 'Normal math' to 'Text math' 
        ' -Match (certain elements of) Normal style.

        On Error GoTo er

        Dim o As Word.OMath
        Dim c As Word.Range
        Dim i As Integer

        For Each o In oms


            o.ConvertToNormalText()

            o.Range.Font.Name = App.ActiveDocument.Styles("Normal").Font.Name
            o.Range.Font.Bold = App.ActiveDocument.Styles("Normal").Font.Bold
            o.Range.Font.Italic = App.ActiveDocument.Styles("Normal").Font.Italic
            o.Range.Font.Size = App.ActiveDocument.Styles("Normal").Font.Size


            IncreaseOMathSupercripts(o)
            'IncreaseOMathElements o


            ' Fix spacing
            i = 1
            Do

                c = o.Range.Characters(i)

                Select Case AscW(c.Text)

                    ' special case minus sign
                    ' hyphen 45, minus 8722
                    Case 45, 8722

                        Select Case c.Previous.Text
                            Case 0 To 9, "a" To "z", "A" To "Z"

                                'space minus sign after a letter or digit (treat as minus sign)
                                If AscW(c.Next.Text) <> AscW(" ") Then
                                    c.InsertAfter(" ")
                                    i = i + 1
                                End If

                            Case Else
                                'don't space minus sign after anything else (treat as negative sign)
                        End Select

                        ' add space before
                        If AscW(c.Previous.Text) <> AscW(" ") Then
                            c.InsertBefore(" ")
                            i = i + 1
                        End If

                    ' All other symbols:
                    ' -surround with spaces on both sides
                    ' -plus 43, times 215, divide 247,
                    ' -< 60, equals 61, >, 62, <= 8804, >= 8805, << 8810, >> 8811, 'proportional to' 8733
                    Case 43, 215, 247, 60, 61, 62, 8733, 8804, 8805, 8810, 8811


                        If AscW(c.Next.Text) <> AscW(" ") Then
                            c.InsertAfter(" ")
                            i = i + 1
                        End If

                        If AscW(c.Previous.Text) <> AscW(" ") Then
                            c.InsertBefore(" ")
                            i = i + 1
                        End If


                    Case Else
                        'do nothing
                End Select

                i = i + 1
            Loop Until i = o.Range.Characters.Count + 1

        Next o
        Exit Sub

er:
        Select Case Err.Number
            Case 91 'no object - first letter has no previous letter object
                Resume Next 'insert space anyway (simplest fix, but inserts an unnecessary space!)
            Case Else
                Err.Raise(Err.Number, Err.Source, Err.Description)
        End Select
    End Sub

End Module
