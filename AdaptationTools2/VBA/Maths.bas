Attribute VB_Name = "Maths"
Option Private Module


Sub PrepBrailleMaths_int(r As Range)
' Fix up Equation Editor Braille Math ready for converting to MathType and Braille.
' -Inline as MathType conversion inserts Tab chars when converting from Display type
' -Convert TextMath back to Math
' -Remove spacing
' -Convert all similar looking dash characters into hyphens.
On Error GoTo er

    Dim shp As Shape

    'replace hyphens with minus
    ReplaceWrongDash r

    'process Math in main text
    OMaths_For_Braille r.OMaths
    
    'process Math in Shapes
    For Each shp In r.ShapeRange
        
        'causes err 5917 if shape does not support attached text
        OMaths_For_Braille shp.TextFrame.TextRange.OMaths
    
    Next shp
    
    Exit Sub

er:
    Select Case Err.number
        Case 5917
        ' shape does not support attached text
            Resume Next
        Case Else
            Err.Raise Err.number, Err.Source, Err.Description
    End Select
End Sub


Sub PrepLargePrintMaths_int(r As Range)
    ' Converts Equations in document to Text (so font can be changed)
    ' Match (certain elements of) Normal style.
    ' Enlarge sub/supscripts in normal text and in Equations
    '
    
    On Error GoTo er
    
    System.Cursor = wdCursorWait
    
    
    ' process normal text
    OMaths_To_TextMath r.OMaths
    
    'replace hyphens with minus (do after converting to textmath)
    ReplaceWrongDash r
    
    IncreaseSuperscripts r
    
    ' Increase Math in Shapes
    For i = 1 To r.ShapeRange.Count
        IncreaseMathInShapes r.ShapeRange(i)
    Next i

    
    ' clean up
    System.Cursor = wdCursorNormal
    
    Exit Sub
er:
    Select Case Err.number
        Case Else
            System.Cursor = wdCursorNormal
            Err.Raise Err.number, Err.Source, Err.Description
    End Select
End Sub


Function IncreaseOMathSupercripts(o As OMath)
        'further enlarge fractions and sub/supscripts in OMath objects
        
        On Error GoTo er
        
        Dim f As OMathFunction
        Dim a As OMath
        Dim g As OMathFunction
        Dim j As OMath
        Dim c As Range
        Dim i As Integer
        Dim supNormal, supAdd, supsize As Integer
        
        
        'size + 6 + (2 for every 18pt)
        supNormal = ActiveDocument.Styles("Normal").Font.Size
        supAdd = Int(supNormal / 18) * 2 '+2 for every 18pt
        supsize = supNormal + 6 + supAdd
     
        
        For i = 1 To o.Functions.Count

            Set f = o.Functions(i)
            
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
Exit Function

er:
Select Case Err.number
    Case -2147467259 'element does not exist
        Resume Next
    Case Else
        Err.Raise Err.number, Err.Source, Err.Description
    End Select
End Function


Function IncreaseSuperscripts(r As Range)
        ' further sub/supscripts
        ' normalsize + 6 + (2 for every 18pt)
        ' this function is very slow when called repeatedly. Therefore it is not used within Shapes code.
        
        Dim supsize As Integer
        Dim s As Integer
        Dim a As Integer
        
        s = ActiveDocument.Styles("Normal").Font.Size
        a = Int(s / 18) * 2 '+2 for every 18pt
        supsize = s + 6 + a
            
        'set up selection find defaults
        r.find.Text = ""
        r.find.ClearFormatting
        r.find.Replacement.Text = ""
        r.find.Replacement.ClearFormatting
        
        With r.find
            .Forward = True
            .Format = False
            .MatchCase = False
            .MatchWholeWord = False
            .MatchAllWordForms = False
            .MatchSoundsLike = False
            .MatchWildcards = False
            .Wrap = wdFindStop
        End With
            


            
        'make superscripts bigger
        r.find.Font.Superscript = True
        r.find.Replacement.Font.Size = supsize
        r.find.Execute replace:=wdReplaceAll
        
        'same again but for subscript.
        r.find.Font.Subscript = True
        r.find.Replacement.Font.Size = supsize
        r.find.Execute replace:=wdReplaceAll
        
        'superscript2 symbol make normal text with superscript
        r.find.Font.Subscript = False
        r.find.Text = ChrW(&HB2)
        r.find.Replacement.Text = "2"
        r.find.Replacement.Font.Superscript = True
        r.find.Execute replace:=wdReplaceAll
        
        'superscript3 symbol
        r.find.Text = ChrW(&HB3)
        r.find.Replacement.Text = "3"
        r.find.Replacement.Font.Superscript = True
        r.find.Execute replace:=wdReplaceAll
        
        'superscript0 symbol
        r.find.Text = ChrW(&H2070)
        r.find.Replacement.Text = 0
        r.find.Replacement.Font.Superscript = True
        r.find.Execute replace:=wdReplaceAll
            
        'superscript4-9 symbol
        For i = 4 To 9
            r.find.Text = ChrW(&H2070 + i)
            r.find.Replacement.Text = i
            r.find.Replacement.Font.Superscript = True
            r.find.Execute replace:=wdReplaceAll
        Next i
        
End Function


Function IncreaseMathInShapes(shp As Shape)
    
    On Error GoTo er

    If shp.Type = msoGroup Then
        ' recurse into groups
        For i = 1 To shp.GroupItems.Count
            IncreaseMathInShapes shp.GroupItems(i)
        Next i
        
    Else
        ' increase text in range
        If shp.TextFrame.HasText Then
            ' Only do OMath. Because IncreaseSuperscripts is very slow on many shapes
            OMaths_To_TextMath shp.TextFrame.TextRange.OMaths
            IncreaseSuperscripts shp.TextFrame.TextRange
        End If
    End If
    
    Exit Function
    
er:
    Select Case Err.number
        Case Else
            Err.Raise Err.number, Err.Source, Err.Description
    End Select
End Function


Function IncreaseOMathElements(o As OMath)
        'enlarge certain smaller OMath objects
        

        Dim f As OMathFunction
        Dim supNormal, supAdd, supsize As Integer
        
        'size + 6 + (2 for every 18pt)
        supNormal = ActiveDocument.Styles("Normal").Font.Size
        supAdd = Int(supNormal / 18) * 2 '+2 for every 18pt
        supsize = supNormal + 6 + supAdd
        
        On Error Resume Next
        
        
        For Each f In o.Functions

            Select Case f.Type
                Case wdOMathFunctionFrac
                    f.Range.Font.Size = supsize
                Case wdOMathFunctionDelim 'brackets
                    IncreaseOMathElements f.Frac
            End Select

        Next f
        
End Function


Function OMaths_For_Braille(oms As OMaths)
' -Convert Text Math To Proper Math
' -Remove spacing
' -Inline as MathType conversion inserts Tab chars when converting from Display type
    
    Dim r As Range
    Dim c As Range
    Dim o As OMath
    'create a range to edit the document as o and c are readonly
    Set r = ActiveDocument.Range
    
    'OMath Objects
    For Each o In oms
        
        'Convert Text Math To Proper Math
        o.ConvertToMathText
        
        ' remove all spaces from maths objects
        FindReplace ChrW(&H20), "", o.Range, False
        
        'Make all Inline as MathType conversion inserts Tab chars when converting from Display type
        o.Type = wdOMathInline
        o.Range.Bold = False
    Next o

End Function


Function OMaths_To_TextMath(oms As OMaths)
'
' Converts OMaths objects to Text and match (certain elements of) Normal style.
'
'
On Error GoTo er

    Dim o As OMath
    Dim c As Range

    For Each o In oms
    
        
        o.ConvertToNormalText
        
        o.Range.Font.name = ActiveDocument.Styles("Normal").Font.name
        o.Range.Font.Bold = ActiveDocument.Styles("Normal").Font.Bold
        o.Range.Font.Italic = ActiveDocument.Styles("Normal").Font.Italic
        o.Range.Font.Size = ActiveDocument.Styles("Normal").Font.Size
        
        
        IncreaseOMathSupercripts o
        'IncreaseOMathElements o
        

        'fix spacing
         i = 1
         Do
            
            Set c = o.Range.Characters(i)
            
            Select Case AscW(c)
            
            ' special case minus sign
            ' hyphen 45, minus 8722
            Case 45, 8722
            
                Select Case c.Previous.Text
                    Case 0 To 9, "a" To "z", "A" To "Z"
                    
                        'space minus sign after a letter or digit (treat as minus sign)
                        If AscW(c.Next.Text) <> AscW(" ") Then
                            c.InsertAfter (" ")
                            i = i + 1
                        End If
                        
                    Case Else
                        'don't space minus sign after anything else (treat as negative sign)
                End Select
                 
                ' add space before
                If AscW(c.Previous.Text) <> AscW(" ") Then
                    c.InsertBefore (" ")
                    i = i + 1
                End If
            
            ' all other symbols surround with spaces on both sides
            ' plus 43, times 215, divide 247,
            ' < 60, equals 61, >, 62, <= 8804, >= 8805, << 8810, >> 8811, proportional to 8733
            Case 43, 215, 247, 60, 61, 62, 8733, 8804, 8805, 8810, 8811
                
                
                If AscW(c.Next.Text) <> AscW(" ") Then
                    c.InsertAfter (" ")
                    i = i + 1
                End If
    
                If AscW(c.Previous.Text) <> AscW(" ") Then
                    c.InsertBefore (" ")
                    i = i + 1
                End If
                
                
            Case Else
                'do nothing
            End Select
        
        i = i + 1
        Loop Until i = o.Range.Characters.Count + 1
        
    Next o
    Exit Function

er:
    Select Case Err.number
        Case 91 'no object - first letter has no previous letter object
            Resume Next 'insert space anyway (simplest fix, but inserts an unnecessary space!)
        Case Else
            Err.Raise Err.number, Err.Source, Err.Description
    End Select
End Function
