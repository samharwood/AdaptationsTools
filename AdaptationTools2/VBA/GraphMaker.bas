Attribute VB_Name = "GraphMaker"
' TODO

' .Axis Labels
' .Plan different profiles

Option Private Module


'500 ~= width of A4 document
Const GM_GUTTER As Integer = 100
Const GM_AXISHEAD As Integer = 30
Const GM_TICK_LENGTH As Integer = 15

Dim graphLL(1) As Single ' Lower Left corner
Dim graphOrigin(1) As Single
Const x As Integer = 0
Const Y As Integer = 1


Dim textHeight As Single
Dim majorGrid As Single
Dim xMinorGrid As Single
Dim yMinorGrid As Single
Dim graphHeight As Single
Dim graphWidth As Single

Dim xTotalLines As Integer
Dim xTotalDivs As Integer
Dim xTotalMajors As Integer

Dim yTotalLines As Integer
Dim yTotalDivs As Integer
Dim yTotalMajors As Integer

Dim xDigitGap As Single
Dim yDigitGap As Single



Sub GM_UI()
    Load GraphMakerUI
    GraphMakerUI.Show vbModeless
    ' If Modal, code execution resumes here after calling Hide on form.

End Sub

Sub GM()
    If Not DBG Then On Error GoTo er

    Dim objUndo As UndoRecord

    Set objUndo = Application.UndoRecord
    objUndo.StartCustomRecord ("Graph Maker")
    
    Word.System.Cursor = wdCursorWait


    ' === Calculate units
    
    textHeight = ActiveDocument.Styles("Normal").Font.Size * 2

    xTotalMajors = Int((Math.Abs(GraphMakerUI.xFrom - GraphMakerUI.xTo) / GraphMakerUI.xNumEvery)) + 1
    ' Combine all elements and roundup to next whole number. It works but don't quite know how!
    xTotalDivs = RoundUp(Math.Abs(GraphMakerUI.xFrom - GraphMakerUI.xTo) / GraphMakerUI.xNumEvery * GraphMakerUI.xDivs)
    xTotalLines = xTotalMajors + xTotalDivs


    yTotalMajors = Int((Math.Abs(GraphMakerUI.yFrom - GraphMakerUI.yTo) / GraphMakerUI.yNumEvery)) + 1
    ' Combine all elements and roundup to next whole number. It works but don't quite know how!
    yTotalDivs = RoundUp(Math.Abs(GraphMakerUI.yFrom - GraphMakerUI.yTo) / GraphMakerUI.yNumEvery * GraphMakerUI.yDivs)
    yTotalLines = yTotalMajors + yTotalDivs


    ' Size graph to best fit the page.
    pgWidth = ActiveDocument.PageSetup.PageWidth - (GM_GUTTER * 2)
    pgHeight = ActiveDocument.PageSetup.PageHeight - (GM_GUTTER * 2)

    ' Major grid is SQUARE. Use shortest XY dimension to make sure it fits on the page.
    If pgWidth / xTotalMajors < pgHeight / yTotalMajors Then
        majorGrid = pgWidth / xTotalMajors
    Else
        majorGrid = pgHeight / yTotalMajors
    End If

    ' Minor grids could be different sizes
    xMinorGrid = majorGrid / (GraphMakerUI.xDivs + 1)
    yMinorGrid = majorGrid / (GraphMakerUI.yDivs + 1)

    graphWidth = xMinorGrid * xTotalLines - xMinorGrid
    graphHeight = yMinorGrid * yTotalLines - yMinorGrid

    graphLL(x) = GM_GUTTER
    graphLL(Y) = GM_GUTTER + graphHeight

    ' Digit gap used to calculate Origin (0,0)
    If graphWidth = 0 Then xDigitGap = 0 _
    Else xDigitGap = graphWidth / Math.Abs(GraphMakerUI.xFrom - GraphMakerUI.xTo)

    If graphHeight = 0 Then yDigitGap = 0 _
    Else yDigitGap = graphHeight / Math.Abs(GraphMakerUI.yFrom - GraphMakerUI.yTo)


    ' === Calculate Origin (0,0)

    g = xDigitGap * GraphMakerUI.xFrom
    If CSng(GraphMakerUI.xFrom) > CSng(GraphMakerUI.xTo) Then g = -g 'if descending sequence
    graphOrigin(x) = graphLL(x) - g

    g = yDigitGap * GraphMakerUI.yFrom
    If CSng(GraphMakerUI.yFrom) > CSng(GraphMakerUI.yTo) Then g = -g 'if descending sequence
    graphOrigin(Y) = graphLL(Y) + g

    ' If Origin is not on the graph, set to LowerLeft instead
    If graphOrigin(x) > graphLL(x) + graphWidth _
    Or graphOrigin(x) < graphLL(x) _
    Then graphOrigin(x) = graphLL(x)
        
    If graphOrigin(Y) < graphLL(Y) - graphHeight _
    Or graphOrigin(Y) > graphLL(Y) _
    Then graphOrigin(Y) = graphLL(Y)

    ' === Plot Graph

    If GraphMakerUI.PlotAsChart Then PlotAsChart Else: PlotAsShapes
        
    objUndo.EndCustomRecord
    Word.System.Cursor = wdCursorNormal
    Application.ScreenUpdating = True
    Exit Sub
    
er:
    objUndo.EndCustomRecord
    Word.System.Cursor = wdCursorNormal
    Application.ScreenUpdating = True
    Select Case Err.number
        Case Else
            MsgBox "Error in GraphMaker" & vbCrLf & Err.number & vbCrLf & Err.Source & vbCrLf & Err.Description
    End Select
End Sub


Sub PlotAsChart()
    
    If Not DBG Then On Error GoTo er
    
    Dim s As Shape
    Dim c As Chart
    Dim a As Axis
    
    
    Set s = ActiveDocument.Shapes.AddChart2(-1, xlXYScatter)
    Set c = s.Chart
    
    c.PlotArea.InsideHeight = graphHeight
    c.PlotArea.InsideWidth = graphWidth

       
    ' Format Chart Elements
    c.ChartData.Workbook.Worksheets("Sheet1").Range("A2:B4") = "" 'clear example values
    c.ChartData.Workbook.Close
    c.HasTitle = False
    c.HasLegend = False

    c.SetElement msoElementPrimaryCategoryGridLinesMinorMajor 'show all grid lines
    c.SetElement msoElementPrimaryValueGridLinesMinorMajor  'show all grid lines
    c.ChartArea.Format.Line.Visible = msoFalse 'chart border line
    c.ChartArea.Format.Fill.Visible = msoFalse 'transparent
    c.PlotArea.Format.Fill.Visible = msoFalse 'transparent
    c.ChartArea.Format.TextFrame2.TextRange.Font.Size = ActiveDocument.Styles("Normal").Font.Size
    c.ChartArea.Format.TextFrame2.TextRange.Font.name = ActiveDocument.Styles("Normal").Font.name
    c.ChartArea.Format.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = wdColorBlack

        
    ' Format Grid Lines
    FormatChartLines c, xlCategory 'X lines
    FormatChartLines c, xlValue 'Y lines
    
    ' Axes Values and Format
    With c.Axes(xlCategory)
        .MinimumScale = GraphMakerUI.xFrom
        .MaximumScale = GraphMakerUI.xTo
        .MajorUnit = GraphMakerUI.xNumEvery
        .MinorUnit = (xMinorGrid / majorGrid) * GraphMakerUI.xNumEvery
        .TickLabels.Format.Fill.BackColor.RGB = wdColorWhite
        .Format.Fill.BackColor.RGB = wdColorWhite
        .Format.Fill.Visible = GraphMakerUI.Numbering
        If GraphMakerUI.Ticks Then .MajorTickMark = xlTickMarkOutside Else .MajorTickMark = xlTickMarkNone
    End With
    
    With c.Axes(xlValue)
        .MinimumScale = GraphMakerUI.yFrom
        .MaximumScale = GraphMakerUI.yTo
        .MajorUnit = GraphMakerUI.yNumEvery
        .MinorUnit = (yMinorGrid / majorGrid) * GraphMakerUI.yNumEvery
        .TickLabels.Format.Fill.BackColor.RGB = wdColorWhite
        .Format.Fill.BackColor.RGB = wdColorWhite
        .Format.Fill.Visible = GraphMakerUI.Numbering
        If GraphMakerUI.Ticks Then .MajorTickMark = xlTickMarkOutside Else .MajorTickMark = xlTickMarkNone
    End With
    
    c.HasAxis(xlValue, xlPrimary) = GraphMakerUI.Axes
    c.HasAxis(xlCategory, xlPrimary) = GraphMakerUI.Axes
    c.ChartArea.Format.TextFrame2.TextRange.Font.Fill.Visible = GraphMakerUI.Numbering
    If GraphMakerUI.AxisLabels Then
        c.SetElement msoElementPrimaryCategoryAxisTitleAdjacentToAxis
        c.SetElement msoElementPrimaryValueAxisTitleHorizontal
    End If
    Exit Sub
    
    
er:
    Select Case Err.number
        Case Else
            MsgBox "Error in GraphMaker.PlotAsChart" & vbCrLf & Err.number & vbCrLf & Err.Source & vbCrLf & Err.Description
    End Select
End Sub


Sub FormatChartLines(c As Chart, v As XlAxisType)

    'Dim a As Axis 'for Debugging

    With c.Axes(v)
        'Axis
        If c.HasAxis(v, xlPrimary) = True Then 'check Axis not disabled
        
            If GraphMakerUI.majorWeight <= 0 Then .Format.Line.Visible = msoFalse Else _
                .Format.Line.Weight = GraphMakerUI.majorWeight * 2
            
            If GraphMakerUI.majorColour.BackColor < 0 Then .Format.Line.Visible = msoFalse Else _
                .Format.Line.ForeColor = GraphMakerUI.majorColour.BackColor
                
            .Format.Line.DashStyle = wdLineStyleSingle
            
        End If
        
        'Major
        If GraphMakerUI.majorWeight <= 0 Then .MajorGridlines.Format.Line.Visible = False Else _
            .MajorGridlines.Format.Line.Weight = GraphMakerUI.majorWeight
        
        If GraphMakerUI.majorColour.BackColor < 0 Then .MajorGridlines.Format.Line.Visible = msoFalse Else _
            .MajorGridlines.Format.Line.ForeColor = GraphMakerUI.majorColour.BackColor
            
        If GraphMakerUI.majorDash <> "Mixed" Then _
            .MajorGridlines.Format.Line.DashStyle = LineDashStyleID(GraphMakerUI.majorDash)


        'Minor
        If GraphMakerUI.minorWeight <= 0 Then .MinorGridlines.Format.Line.Visible = msoFalse Else _
            .MinorGridlines.Format.Line.Weight = GraphMakerUI.minorWeight
        
        If GraphMakerUI.minorColour.BackColor < 0 Then .MinorGridlines.Format.Line.Visible = msoFalse Else _
            .MinorGridlines.Format.Line.ForeColor = GraphMakerUI.minorColour.BackColor
            
        If GraphMakerUI.minorDash <> "Mixed" Then _
            .MinorGridlines.Format.Line.DashStyle = LineDashStyleID(GraphMakerUI.minorDash)
        
    End With
    
End Sub


Sub PlotAsShapes()
    
    Dim GroupAll As New Collection
    Dim i As Integer
    Dim a As Variant
    Dim shpGroup As Shape
   
    If GraphMakerUI.majorColour.BackColor >= 0 And GraphMakerUI.majorWeight > 0 Then 'if visible
        PlotX_Major
        PlotY_Major
        GroupAll.Add "X Major"
        GroupAll.Add "Y Major"
        DoEvents
    End If
    
    If GraphMakerUI.minorColour.BackColor >= 0 And GraphMakerUI.minorWeight > 0 Then 'if visible
        If GraphMakerUI.xDivs > 0 Then
            PlotX_Minor
            GroupAll.Add "X Minor"
        End If
        If GraphMakerUI.yDivs > 0 Then
            PlotY_Minor
            GroupAll.Add "Y Minor"
        End If
        DoEvents
    End If
    
    If GraphMakerUI.Axes Then
        PlotAxes
        GroupAll.Add "Axes"
    End If

    If GraphMakerUI.AxisLabels Then
        PlotAxisLabels
        GroupAll.Add "X Axis Label"
        GroupAll.Add "Y Axis Label"
    End If
   
    If GraphMakerUI.Numbering Then
        PlotX_Numbers
        PlotY_Numbers
        GroupAll.Add "X Labels"
        GroupAll.Add "Y Labels"
        DoEvents
    End If
    
    ' Group all
    ReDim a(GroupAll.Count - 1)
    For i = 0 To GroupAll.Count - 1
        a(i) = GroupAll(i + 1)
    Next i
    
    Set shpGroup = ActiveDocument.Shapes.Range(a).Group
    shpGroup.name = "Graph X" & GraphMakerUI.xFrom & "," & GraphMakerUI.xTo _
                        & " Y" & GraphMakerUI.yFrom & "," & GraphMakerUI.yTo
    

End Sub

Sub PlotX_Major()

 If Not DBG Then On Error GoTo er

 Dim shpLine As Shape
 Dim shpGroup As Shape
 Dim doc As Document
 Set doc = ActiveDocument
 Dim a As Variant
 Dim gap As Single
 
    gap = majorGrid
    
    ReDim a(xTotalMajors - 1)
     
    For i = 0 To xTotalMajors - 1
        
        ln = (gap * i) + graphLL(x)
        
        TopLeft = ln
        TopDown = graphLL(Y)
        BottomLeft = ln
        BottomDown = graphLL(Y) - graphHeight
        
        If GraphMakerUI.Ticks Then TopDown = graphLL(Y) + GM_TICK_LENGTH
        
        Set shpLine = doc.Shapes.AddLine(TopLeft, TopDown, BottomLeft, BottomDown)
        
        shpLine.name = "X " & i
        
        a(i) = shpLine.name
    
    Next i
    
    ' if just 1 shape
    ' set shpGroup to just that 1 shape
    If UBound(a) = 0 Then Set shpGroup = doc.Shapes(a(0)) _
        Else: Set shpGroup = doc.Shapes.Range(a).Group
        
    shpGroup.name = "X Major"
    FormatMajorStyle shpGroup

Exit Sub

er:
    Select Case Err.number
        Case Else
            Err.Raise Err.number
    End Select
End Sub

Sub PlotY_Major()
If Not DBG Then On Error GoTo er

 Dim shpLine As Shape
 Dim shpGroup As Shape
 Dim doc As Document
 Set doc = ActiveDocument
 Dim a As Variant
 Dim gap As Single
 
    gap = majorGrid
    
    ReDim a(yTotalMajors - 1)
     
    For i = 0 To yTotalMajors - 1
        
        ln = (gap * i)
        
        ' Construct bottom up to start major on zero axis
        LeftAlong = graphLL(x)
        LeftDown = graphLL(Y) - ln
        RightAlong = graphLL(x) + graphWidth
        RightDown = graphLL(Y) - ln
        
        If GraphMakerUI.Ticks Then
            LeftAlong = graphLL(x) - GM_TICK_LENGTH
        End If
        
        Set shpLine = doc.Shapes.AddLine(LeftAlong, LeftDown, RightAlong, RightDown)
        
        shpLine.name = "Y " & i
        
        a(i) = shpLine.name
    
    Next i

    ' if just 1 shape
    ' set shpGroup to just that 1 shape
    If UBound(a) = 0 Then Set shpGroup = doc.Shapes(a(0)) _
        Else: Set shpGroup = doc.Shapes.Range(a).Group

    shpGroup.name = "Y Major"
    FormatMajorStyle shpGroup

Exit Sub

er:
    Select Case Err.number
        Case Else
            Err.Raise Err.number
    End Select
End Sub

Sub PlotX_Minor()
If Not DBG Then On Error GoTo er

 Dim shpLine As Shape
 Dim shpGroup As Shape
 Dim doc As Document
 Set doc = ActiveDocument
 Dim a As Variant
 Dim gap As Single
 
    gap = xMinorGrid
    
    ReDim a(xTotalLines - 1)
     
    For i = 0 To xTotalLines - 1
        
        ln = (gap * i) + graphLL(x)
        
        'TopLeft, TopDown, BottomLeft, BottomDown
        Set shpLine = doc.Shapes.AddLine(ln, GM_GUTTER, ln, graphLL(Y))
        shpLine.name = "X " & i
        
        a(i) = shpLine.name
    
    Next i

    ' if just 1 shape
    ' set shpGroup to just that 1 shape
    If UBound(a) = 0 Then Set shpGroup = doc.Shapes(a(0)) _
        Else: Set shpGroup = doc.Shapes.Range(a).Group
        
    shpGroup.name = "X Minor"
    FormatMinorStyle shpGroup

Exit Sub

er:
    Select Case Err.number
        Case Else
            Err.Raise Err.number
    End Select
End Sub

Sub PlotY_Minor()
If Not DBG Then On Error GoTo er

 Dim shpLine As Shape
 Dim shpGroup As Shape
 Dim doc As Document
 Set doc = ActiveDocument
 Dim a As Variant
 Dim gap As Single
 
    gap = yMinorGrid
    
    ReDim a(yTotalLines - 1)
     
    For i = 0 To yTotalLines - 1
        
        ln = (gap * i)
        
        ' Construct bottom up to start major on zero axis
        'LeftAlong, LeftDown, RightAlong, RightDown
        Set shpLine = doc.Shapes.AddLine(graphLL(x), graphLL(Y) - ln, graphLL(x) + graphWidth, graphLL(Y) - ln)
        shpLine.name = "Y " & i
        
        a(i) = shpLine.name
    
    Next i

    ' if just 1 shape
    ' set shpGroup to just that 1 shape
    If UBound(a) = 0 Then Set shpGroup = doc.Shapes(a(0)) _
        Else: Set shpGroup = doc.Shapes.Range(a).Group
        
    shpGroup.name = "Y Minor"
    FormatMinorStyle shpGroup

Exit Sub

er:
    Select Case Err.number
        Case Else
            Err.Raise Err.number
    End Select
End Sub


Sub PlotX_Numbers()

 If Not DBG Then On Error GoTo er

 Dim shp As Shape
 Dim shpGroup As Shape
 Dim doc As Document
 Set doc = ActiveDocument
 Dim a As Variant
 Dim n As Single

    ReDim a(xTotalMajors - 1)

    gap = majorGrid
 
    For i = 0 To xTotalMajors - 1
        ln = gap * i
        'TopLeft, TopDown, LabelWidth, LabelHeight
        Set shp = doc.Shapes.AddTextbox(msoTextOrientationHorizontal, GM_GUTTER + ln, graphOrigin(Y), gap, textHeight * 1.5)
        
        n = i
        If CSng(GraphMakerUI.xFrom) > CSng(GraphMakerUI.xTo) Then n = -i 'descending numbering
        n = GraphMakerUI.xFrom + (n * GraphMakerUI.xNumEvery)
        
        shp.name = "X Label " & n
        shp.TextFrame.TextRange.Text = FormatNumbering(n)
        shp.TextFrame.TextRange.ParagraphFormat.Alignment = wdAlignParagraphCenter
        shp.TextFrame.TextRange.Font.Shading.BackgroundPatternColor = wdColorWhite
        a(i) = shp.name
    Next i

    ' if just 1 shape
    ' set shpGroup to just that 1 shape
    If UBound(a) = 0 Then Set shpGroup = doc.Shapes(a(0)) _
        Else: Set shpGroup = doc.Shapes.Range(a).Group
        
    shpGroup.name = "X Labels"
    shpGroup.Line.Visible = msoFalse
    shpGroup.Fill.Visible = msoFalse
    shpGroup.Left = shpGroup.Left - (gap / 2) 'minus half of label width
    If GraphMakerUI.Ticks Then shpGroup.Top = shpGroup.Top + GM_TICK_LENGTH 'space for ticks

Exit Sub

er:
    Select Case Err.number
        Case Else
            Err.Raise Err.number
    End Select
End Sub



Sub PlotY_Numbers()
If Not DBG Then On Error GoTo er

 Dim shp As Shape
 Dim shpGroup As Shape
 Dim doc As Document
 Set doc = ActiveDocument
 Dim a As Variant
 Dim n As Single
  
 ' Zero based
 ReDim a(yTotalMajors - 1)


    gap = majorGrid
 
    For i = 0 To yTotalMajors - 1
        ln = gap * i
        'TopLeft, TopDown, LabelWidth, LabelHeight
        Set shp = doc.Shapes.AddTextbox(msoTextOrientationHorizontal, graphOrigin(x), graphLL(Y) - ln, textHeight * 1.5, gap)
        
        n = i
        If CSng(GraphMakerUI.yFrom) > CSng(GraphMakerUI.yTo) Then n = -i 'descending numbering
        n = GraphMakerUI.yFrom + (n * GraphMakerUI.yNumEvery)
        
        shp.name = "Y Label " & n
        shp.TextFrame.TextRange.Text = FormatNumbering(n)
        shp.TextFrame.TextRange.ParagraphFormat.Alignment = wdAlignParagraphRight
        shp.TextFrame.TextRange.ParagraphFormat.LineSpacingRule = wdLineSpaceSingle 'for correct alignment with gridline
        shp.TextFrame.VerticalAnchor = msoAnchorMiddle 'for correct alignment with gridline
        shp.TextFrame.TextRange.Font.Shading.BackgroundPatternColor = wdColorWhite
        a(i) = shp.name
    Next i

    ' if just 1 shape
    ' set shpGroup to just that 1 shape
    If UBound(a) = 0 Then Set shpGroup = doc.Shapes(a(0)) _
        Else: Set shpGroup = doc.Shapes.Range(a).Group
        
    shpGroup.name = "Y Labels"
    shpGroup.Line.Visible = msoFalse
    shpGroup.Fill.Visible = msoFalse
    
        
    'reposition
    shpGroup.Top = shpGroup.Top - (gap / 2) 'half label height
    shpGroup.Left = shpGroup.Left - (textHeight * 1.5) 'label width
    If GraphMakerUI.Ticks Then shpGroup.Left = shpGroup.Left - GM_TICK_LENGTH 'space for ticks

Exit Sub

er:
    Select Case Err.number
        Case Else
            Err.Raise Err.number
    End Select
End Sub

Sub PlotAxes()

    Dim shp As Shape
    Dim shpGroup As Shape
    Dim a() As Variant
    
    ReDim a(1)
    'Y
    lineLength = graphLL(Y) - (graphHeight + GM_AXISHEAD)
        
    Set shp = ActiveDocument.Shapes.AddLine(graphOrigin(x), graphLL(Y), graphOrigin(x), lineLength)
    shp.name = "Y Axis"
    a(0) = shp.name
    
    'X
    lineLength = graphLL(x) + graphWidth + GM_AXISHEAD
            
    Set shp = ActiveDocument.Shapes.AddLine(graphLL(x), graphOrigin(Y), lineLength, graphOrigin(Y))
    shp.name = "X Axis"
    a(1) = shp.name
    
    
    'Grouping
    Set shpGroup = ActiveDocument.Shapes.Range(a).Group
    shpGroup.name = "Axes"
    
    If GraphMakerUI.majorWeight <= 0 Then shpGroup.Line.Weight = "6" _
        Else shpGroup.Line.Weight = GraphMakerUI.majorWeight * 2        'might be invisible

    shpGroup.Line.BeginArrowheadStyle = msoArrowheadNone
    shpGroup.Line.EndArrowheadStyle = msoArrowheadTriangle
    shpGroup.Line.ForeColor = 0 'always black. complicated to copy major value because of transparency

    
End Sub

Sub PlotAxisLabels()
    
    Dim shpGroup As Shape
    Dim shp As Shape

    'X
    TopLeft = graphLL(x) + graphWidth + GM_AXISHEAD
    TopDown = graphOrigin(Y) - (textHeight / 2)
    
    'TopLeft, TopDown, LabelWidth, LabelHeight
    Set shp = ActiveDocument.Shapes.AddTextbox(msoTextOrientationHorizontal, TopLeft, TopDown, textHeight, textHeight)
    
    shp.name = "X Axis Label"
    shp.TextFrame.TextRange.Text = "X"
    shp.TextFrame.TextRange.ParagraphFormat.Alignment = wdAlignParagraphCenter
    shp.TextFrame.TextRange.Font.Shading.BackgroundPatternColor = wdColorWhite
    shp.Fill.Visible = msoFalse
    shp.Line.Visible = msoFalse
    
    'Y
    TopLeft = graphOrigin(x) - (textHeight / 2)
    TopDown = graphLL(Y) - (graphHeight + GM_AXISHEAD + textHeight)
    
    'TopLeft, TopDown, LabelWidth, LabelHeight
    Set shp = ActiveDocument.Shapes.AddTextbox(msoTextOrientationHorizontal, TopLeft, TopDown, textHeight, textHeight)
    
    shp.name = "Y Axis Label"
    shp.TextFrame.TextRange.Text = "Y"
    shp.TextFrame.TextRange.ParagraphFormat.Alignment = wdAlignParagraphCenter
    shp.TextFrame.TextRange.Font.Shading.BackgroundPatternColor = wdColorWhite
    shp.Fill.Visible = msoFalse
    shp.Line.Visible = msoFalse


End Sub

Sub FormatMajorStyle(shp As Shape)

    shp.Line.Weight = GraphMakerUI.majorWeight
    shp.Line.ForeColor = GraphMakerUI.majorColour.BackColor
    shp.Line.DashStyle = LineDashStyleID(GraphMakerUI.majorDash)
    shp.Line.BeginArrowheadStyle = msoArrowheadNone
    shp.Line.EndArrowheadStyle = msoArrowheadNone
    shp.ZOrder msoBringToFront
End Sub


Sub FormatMinorStyle(shp As Shape)

    shp.Line.Weight = GraphMakerUI.minorWeight
    shp.Line.ForeColor = GraphMakerUI.minorColour.BackColor
    shp.Line.DashStyle = LineDashStyleID(GraphMakerUI.minorDash)
    shp.Line.BeginArrowheadStyle = msoArrowheadNone
    shp.Line.EndArrowheadStyle = msoArrowheadNone
    shp.ZOrder msoSendToBack
End Sub


Function FormatNumbering(number As Single) As String
' Return UEB Braille coded string for the number or just return the number as is.
    
    If GraphMakerUI.UEBBraille = False Then
        FormatNumbering = number
        Exit Function
    End If
    
    'negatives
    If number < 0 Then
        s = Chr(34) & "-" 'start with quote("), dash(-)
        's = Chr(34) & "-#" 'start with quote("), dash(-) number sign(#)
    Else
        s = "" 'start with nothing
        's = "#" 'start with number sign(#)
    End If
        
    For i = 1 To Len(CStr(number))
    
        n = Mid(CStr(number), i, 1)
        
        Select Case n
            Case Chr(45) 'minus
                'ignore. already converted above
            Case Chr(46)   'decimal point
                s = s & "4"
            Case 0
                s = s & "j"
            Case 1 To 9
                s = s & ChrW(Asc(n) + 48)
            Case "E"
                s = s & ",e" 'exponent(letter e)
            Case "+"
                s = s & Chr(34) & "6" & "#" 'plus sign("6) [and number sign(#) in braille] that comes after exponent
            Case Else
                Err.Raise Err.number
        End Select

    Next i
        
    FormatNumbering = s
    
End Function






