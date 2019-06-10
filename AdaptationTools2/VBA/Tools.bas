Attribute VB_Name = "Tools"

' Developed by: Sam Harwood for Wakefield College Adaptations Department
' Contact: s.harwood@wakefield.ac.uk
' Last updated: 29/03/2019

' Macro bindings for toolbars




Sub SymbolCode()
' Helper function
' Select text in document and run macro to print the Character Code(s) to Debug window.

    Dim r As Range
    Set r = SelectionToRange
    
    i = 1
    Do
        Set c = r.Characters(i)
        
        Debug.Print c; AscW(c)
        
        i = i + 1
    Loop Until i = r.Characters.Count + 1
        
End Sub


Sub ReplaceNonBreakingSpaces()
    '
    'Replace the last space in each paragraph with a non-breaking space,
    'so you don't end up with a single hanging word on a new line.
    '
    On Error Resume Next
    Dim r As Range
    Dim objUndo As UndoRecord

    Set objUndo = Application.UndoRecord
    objUndo.StartCustomRecord ("Replace Non-Breaking Spaces")

    Set r = SelectionToRange
    
        ReplaceNonBreakingSpaces_int r
        
    objUndo.EndCustomRecord
End Sub


Sub PrepBrailleMaths()
    ' Fix up Equation Editor Braille Math ready for converting to MathType for Brailling.
    ' -Set Inline type, as MathType conversion inserts Tab chars when converting from Display type
    ' -Convert TextMath back to Math
    ' -Remove spacing
    ' -Convert all similar looking dash characters into hyphens.
    On Error GoTo er
    
        Dim objUndo As UndoRecord
        Dim r As Range
    
    
        Set objUndo = Application.UndoRecord
        objUndo.StartCustomRecord ("Convert Math for Brailling")
        
        Set r = SelectionToRange
            
            PrepBrailleMaths_int r
     
        objUndo.EndCustomRecord
        Exit Sub
    
er:
    
    Select Case Err.number
        Case Else
            MsgBox "PrepBrailleMaths " & vbCrLf & Err.number & vbCrLf & Err.Description
    End Select
End Sub


Sub PrepLargePrintMaths()
    ' Converts Equations in document to Text (so font can be changed)
    ' Match (certain elements of) Normal style.
    ' Enlarge sub/supscripts in normal text and in Equations
    '
    On Error GoTo er
    
        Dim objUndo As UndoRecord
        Dim r As Range
        
        Set r = SelectionToRange
        If r.Start = r.End Then Exit Sub
        
        Set objUndo = Application.UndoRecord
        objUndo.StartCustomRecord ("Convert Math to Large Print")
        
            PrepLargePrintMaths_int r
        
        objUndo.EndCustomRecord
        Exit Sub
    
er:
    
    Select Case Err.number
        Case Else
            MsgBox "PrepLargePrintMaths " & vbCrLf & Err.number & vbCrLf & Err.Description
    End Select
End Sub


Sub FixOCR_Ligatures()
    Dim objUndo As UndoRecord

    Set objUndo = Application.UndoRecord
    objUndo.StartCustomRecord ("Fix OCR Ligatures")
    
    Dim rng As Range
    
    Set rng = Selection.Range
    
        FixOCR_Ligatures_int rng
    
    rng.Select
    
    objUndo.EndCustomRecord
End Sub

Sub GraphMaker()
    GM_UI
End Sub

Sub PasteOCR()
    Dim objUndo As UndoRecord

    Set objUndo = Application.UndoRecord
    objUndo.StartCustomRecord ("Paste OCR")
    
    Dim rng As Range
        
        PasteOCR_int
    
    objUndo.EndCustomRecord
End Sub


Sub Exp_ResizeDocument()
On Error GoTo er
    ' Experimental!
    ' Scales all text in document, and in textboxes and thickens shape outlines, by relative amount.
    ' Input 2 text sizes to calc percentage to enlarge by

    Dim scaleby As Double
    Dim szL As Integer
    Dim szN As Integer
    Dim i As Integer
    Dim rng As Range
    Dim objUndo As UndoRecord

    Set objUndo = Application.UndoRecord
    objUndo.StartCustomRecord ("Resize Images/Textboxes")
    
    szN = InputBox("Input current normal text size: ", "Normal", ActiveDocument.Styles("Normal").Font.Size)
    szL = InputBox("Input text size to enlarge to: ", "Enlarge")
    scaleby = szL / szN
    Set rng = SelectionToRange
    
    ResizeShapeTextAndLines rng, scaleby, szL
    'scale shapes
    rng.ShapeRange.ScaleHeight scaleby, False
    rng.ShapeRange.ScaleWidth scaleby, False
er:
    objUndo.EndCustomRecord
End Sub


Sub ToggleTextBoundaries()
    On Error Resume Next
    ActiveWindow.View.ShowTextBoundaries = _
    Not ActiveWindow.View.ShowTextBoundaries

End Sub

Sub ThickenLines()
        
    Dim objUndo As UndoRecord
    Dim r As Range
    
    Set r = SelectionToRange
    If r.Start = r.End Then Exit Sub
    
    Set objUndo = Application.UndoRecord
    objUndo.StartCustomRecord ("Thicken Lines")
    
        ThickenLines_int r
    
    r.Select
    
    objUndo.EndCustomRecord
    
End Sub


Sub TripleParaToSinglePara()
    '
    'doing it this way retains formatting
    '
    FindReplace "(^13)^13^13", "\1", SelectionToRange, True
End Sub


Sub DisableMT()
    '
    ' Disable MathType
    '
    On Error Resume Next
    AddIns("C:\Program Files\MathType\Office Support\WordCmds.dot").Installed _
        = False
    AddIns( _
        "C:\Program Files\Microsoft Office\Office12\STARTUP\mathtype commands 6 for word.dotm" _
        ).Installed = False
End Sub

Sub EnableMT()
    '
    ' Enable MathType
    '
    On Error Resume Next
        AddIns.Add "C:\Program Files\Microsoft Office\Office12\STARTUP\MathType Commands 6 For Word.dotm"
        AddIns.Add "C:\Program Files\MathType\Office Support\WordCmds.dot"
End Sub

Sub ToggleMT()
    '
    ' Toggle MathType
    '
    On Error Resume Next
        If AddIns("C:\Program Files\Microsoft Office\Office12\STARTUP\MathType Commands 6 For Word.dotm").Installed _
                = False _
            Or AddIns("C:\Program Files\MathType\Office Support\WordCmds.dot").Installed _
                = False _
        Then
            EnableMT
        Else
            DisableMT
        End If
        Exit Sub
er:
        MsgBox "The MathType template could not be found. MathType may not be installed properly on this computer."
End Sub







