Imports System.ComponentModel


Public Class GraphMaker

    ' TODO

    ' .Redo Ticks
    ' .Plan different profiles

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

    Dim pgWidth As Single
    Dim pgHeight As Single
    Dim xRange As Single
    Dim yRange As Single

    Dim xTotalLines As Integer
    Dim xTotalDivs As Integer
    Dim xTotalMajors As Integer

    Dim yTotalLines As Integer
    Dim yTotalDivs As Integer
    Dim yTotalMajors As Integer

    Dim xDigitGap As Single
    Dim yDigitGap As Single



    ' == Events

    Private Sub GraphMakerUI_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
        e.Cancel = True
        SavedCtls.Default.Save()
        Me.Hide()
    End Sub

    Private Sub GraphMakerUI_Load(sender As Object, e As EventArgs) Handles Me.Load

        PopulateLists()

        xFrom.Text = SavedCtls.Default.xFrom
        yFrom.Text = SavedCtls.Default.yFrom
        xTo.Text = SavedCtls.Default.xTo
        yTo.Text = SavedCtls.Default.yTo

        MajorWeight.Text = SavedCtls.Default.MajorWeight
        MajorColour.BackColor = SavedCtls.Default.MajorColour
        MajorLineStyle.SelectedIndex = SavedCtls.Default.MajorLineStyle

        MinorWeight.Text = SavedCtls.Default.MinorWeight
        MinorColour.BackColor = SavedCtls.Default.MinorColour
        MinorLineStyle.SelectedIndex = SavedCtls.Default.MinorLineStyle

        xNumEvery.Text = SavedCtls.Default.xNumEvery
        yNumEvery.Text = SavedCtls.Default.yNumEvery
        xDivs.Text = SavedCtls.Default.xDivs
        yDivs.Text = SavedCtls.Default.yDivs

        Axes.Checked = SavedCtls.Default.Axes
        AxisLabels.Checked = SavedCtls.Default.AxisLabels
        Ticks.Checked = SavedCtls.Default.Ticks

        Dim r As Windows.Forms.RadioButton
        r = GrpPlotAs.Controls.Item(SavedCtls.Default.PlotAs)
        r.Checked = True

        r = GrpNumbering.Controls.Item(SavedCtls.Default.Numbering)
        r.Checked = True

    End Sub



    ' === Validation

    Private Function ErrNotANumber(ByRef ctl As Windows.Forms.TextBox) As Boolean
        If IsNumeric(ctl.Text) Then
            SavedCtls.Default.PropertyValues.Item(ctl.Name).PropertyValue = ctl.Text
            ErrNotANumber = False
        Else
            ErrNotANumber = True
        End If
    End Function

    Private Sub ErrNotPositive(ByRef ctl As Windows.Forms.TextBox)
        If IsNumeric(ctl.Text) And ctl.Text > 0 Then
            SavedCtls.Default.PropertyValues.Item(ctl.Name).PropertyValue = ctl.Text
        Else
            ctl.Undo()
        End If
    End Sub

    Private Sub Num_Validating(sender As Object, e As CancelEventArgs) Handles xFrom.Validating, xTo.Validating, yFrom.Validating, yTo.Validating, MajorWeight.Validating, MinorWeight.Validating, yNumEvery.Validating, yDivs.Validating, xNumEvery.Validating, xDivs.Validating
        e.Cancel = ErrNotANumber(sender)
    End Sub

    Private Sub Text_Validated(sender As Object, e As EventArgs) Handles xFrom.Validating, xTo.Validated, yFrom.Validated, yTo.Validated, MajorWeight.Validated, MinorWeight.Validated, yNumEvery.Validated, yDivs.Validated, xNumEvery.Validated, xDivs.Validated
        SavedCtls.Default.PropertyValues.Item(sender.Name).PropertyValue = sender.Text
    End Sub

    Private Sub PlotAs_Validated(sender As Object, e As EventArgs) Handles PlotAsChart.Validated, PlotAsShapes.Validated
        SavedCtls.Default.PropertyValues.Item("PlotAs").PropertyValue = sender.Name
    End Sub

    Private Sub Numbering_Validated(sender As Object, e As EventArgs) Handles NumUEB.Validated, NumStandard.Validated, NumNone.Validated
        SavedCtls.Default.PropertyValues.Item("Numbering").PropertyValue = sender.Name
    End Sub

    Private Sub CheckBox_Validated(sender As Object, e As EventArgs) Handles Ticks.Validated, Axes.Validated, AxisLabels.Validated
        SavedCtls.Default.PropertyValues.Item(sender.Name).PropertyValue = sender.Checked
    End Sub

    Private Sub MajorLineStyle_SelectedIndexChanged(sender As Object, e As EventArgs) Handles MajorLineStyle.SelectedIndexChanged
        SavedCtls.Default.MajorLineStyle = MajorLineStyle.SelectedIndex
    End Sub

    Private Sub MinorLineStyle_SelectedIndexChanged(sender As Object, e As EventArgs) Handles MinorLineStyle.SelectedIndexChanged
        SavedCtls.Default.MinorLineStyle = MinorLineStyle.SelectedIndex
    End Sub

    ' == Misc

    Sub PopulateLists()

        ' Copied from Word.WdLineStyle Enum
        ' because can't enumerate Enum to use it directly
        Dim WordLineStyles() As String = {
            "None",
            "Single",
            "Dot",
            "DashSmallGap",
            "DashLargeGap",
            "DashDot",
            "DashDotDot",
            "Double",
            "Triple",
            "ThinThickSmallGap",
            "ThickThinSmallGap",
            "ThinThickThinSmallGap",
            "ThinThickMedGap",
            "ThickThinMedGap",
            "ThinThickThinMedGap",
            "ThinThickLargeGap",
            "ThickThinLargeGap",
            "ThinThickThinLargeGap",
            "SingleWavy",
            "DoubleWavy",
            "DashDotStroked",
            "Emboss3D",
            "Engrave3D",
            "Outset",
            "Inset"
        }

        Dim ChartLineStyles() As String = {
            "Solid",
            "Square Dot",
            "Round Dot",
            "Long Dash",
            "Long Dash Dot",
            "Long Dash Dot Dot",
            "Sys Dash",
            "Sys Dot",
            "Sys Dash Dot"
        }

        Dim DashLineStyles() As String = {
            "Solid",
            "Square Dot",
            "Round Dot",
            "Dash",
            "Dash Dot",
            "Dash Dot Dot",
            "Long Dash",
            "Long Dash Dot",
            "Long Dash Dot Dot",
            "Sys Dash",
            "Sys Dot",
            "Sys Dash Dot"
        }

        MajorLineStyle.Items.AddRange(DashLineStyles)
        MinorLineStyle.Items.AddRange(DashLineStyles)
    End Sub

    ' == UI Actions

    Private Sub btnCreate_Click(sender As Object, e As EventArgs) Handles btnCreate.Click
        CreatGraph()
    End Sub

    Private Sub majorColour_Click(sender As Object, e As EventArgs) Handles MajorColour.Click

        If ColorDialog1.ShowDialog() = DialogResult.OK Then
            MajorColour.BackColor = ColorDialog1.Color
            SavedCtls.Default.MajorColour = ColorDialog1.Color
        End If

    End Sub

    Private Sub minorColour_Click(sender As Object, e As EventArgs) Handles MinorColour.Click

        If ColorDialog1.ShowDialog() = DialogResult.OK Then
            MinorColour.BackColor = ColorDialog1.Color
            SavedCtls.Default.MinorColour = ColorDialog1.Color
        End If

    End Sub

    Private Sub PlotAsShapes_CheckedChanged(sender As Object, e As EventArgs) Handles PlotAsShapes.CheckedChanged
        NumUEB.Enabled = True
    End Sub

    Private Sub PlotAsChart_CheckedChanged(sender As Object, e As EventArgs) Handles PlotAsChart.CheckedChanged
        NumUEB.Enabled = False
        If NumUEB.Checked Then
            NumStandard.Checked = True
            Numbering_Validated(NumStandard, e) 'save the change
        End If

    End Sub

    ' == Create Graph

    Friend Sub CreatGraph()
        If Not DBG Then On Error GoTo er

        Dim objUndo As Word.UndoRecord
        objUndo = App.UndoRecord
        objUndo.StartCustomRecord("Graph Maker")

        App.System.Cursor = Word.WdCursorType.wdCursorWait

        CalculateUnits()

        If PlotAsChart.Checked Then PlotChart() Else PlotShapes()

        objUndo.EndCustomRecord()
        App.System.Cursor = Word.WdCursorType.wdCursorNormal
        Exit Sub

er:
        objUndo.EndCustomRecord()
        App.System.Cursor = Word.WdCursorType.wdCursorNormal
        Select Case Err.Number
            Case Else
                MsgBox("Error in GraphMaker" & vbCrLf & Err.Number & vbCrLf & Err.Source & vbCrLf & Err.Description)
        End Select
    End Sub

    Private Sub CalculateUnits()
        ' === Calculate all units used in module

        xRange = Math.Abs(xFrom.Text - xTo.Text)
        yRange = Math.Abs(yFrom.Text - yTo.Text)

        textHeight = ActiveDocument.Styles("Normal").Font.Size * 2

        xTotalMajors = Int(xRange / xNumEvery.Text) + 1
        ' Combine all elements and roundup to next whole number. It works but don't quite know how!
        xTotalDivs = RoundUp(xRange / xNumEvery.Text * xDivs.Text)
        xTotalLines = xTotalMajors + xTotalDivs


        yTotalMajors = Int(yRange / yNumEvery.Text) + 1
        ' Combine all elements and roundup to next whole number. It works but don't quite know how!
        yTotalDivs = RoundUp(yRange / yNumEvery.Text * yDivs.Text)
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
        xMinorGrid = majorGrid / (xDivs.Text + 1)
        yMinorGrid = majorGrid / (yDivs.Text + 1)

        graphWidth = xMinorGrid * xTotalLines - xMinorGrid
        graphHeight = yMinorGrid * yTotalLines - yMinorGrid

        graphLL(x) = GM_GUTTER
        graphLL(Y) = GM_GUTTER + graphHeight

        ' === Calculate Origin (0,0)

        If graphWidth = 0 Then xDigitGap = 0 _
             Else xDigitGap = graphWidth / xRange

        If graphHeight = 0 Then yDigitGap = 0 _
             Else yDigitGap = graphHeight / yRange

        Dim g As Single

        g = xDigitGap * xFrom.Text
        If CSng(xFrom.Text) > CSng(xTo.Text) Then g = -g 'if descending sequence
        graphOrigin(x) = graphLL(x) - g

        g = yDigitGap * yFrom.Text
        If CSng(yFrom.Text) > CSng(yTo.Text) Then g = -g 'if descending sequence
        graphOrigin(Y) = graphLL(Y) + g

        ' If Origin is not on the graph, set to LowerLeft instead
        If graphOrigin(x) > graphLL(x) + graphWidth _
            Or graphOrigin(x) < graphLL(x) _
            Then graphOrigin(x) = graphLL(x)

        If graphOrigin(Y) < graphLL(Y) - graphHeight _
            Or graphOrigin(Y) > graphLL(Y) _
            Then graphOrigin(Y) = graphLL(Y)
    End Sub

    ' == Plot as Chart

    Sub PlotChart()

        If Not DBG Then On Error GoTo er

        Dim s As Word.Shape
        Dim c As Word.Chart
        Dim a As Word.Axis


        s = ActiveDocument.Shapes.AddChart2(-1, Office.XlChartType.xlXYScatter)
        c = s.Chart

        c.PlotArea.InsideHeight = graphHeight
        c.PlotArea.InsideWidth = graphWidth


        ' Format Chart Elements
        c.ChartData.Workbook.Worksheets("Sheet1").Range("A2:B4") = "" 'clear example values
        c.ChartData.Workbook.Close
        c.HasTitle = False
        c.HasLegend = False


        c.SetElement(Office.MsoChartElementType.msoElementPrimaryCategoryGridLinesMinorMajor) 'show all grid lines
        c.SetElement(Office.MsoChartElementType.msoElementPrimaryValueGridLinesMinorMajor)  'show all grid lines
        c.ChartArea.Format.Line.Visible = Office.MsoTriState.msoFalse 'chart border line
        c.ChartArea.Format.Fill.Visible = Office.MsoTriState.msoFalse 'transparent
        c.PlotArea.Format.Fill.Visible = Office.MsoTriState.msoFalse 'transparent
        c.ChartArea.Format.TextFrame2.TextRange.Font.Size = ActiveDocument.Styles("Normal").Font.Size
        c.ChartArea.Format.TextFrame2.TextRange.Font.Name = ActiveDocument.Styles("Normal").Font.Name
        c.ChartArea.Format.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = Word.WdColor.wdColorBlack


        ' Format Grid Lines
        FormatChartLines(c, Excel.XlAxisType.xlCategory) 'X lines
        FormatChartLines(c, Excel.XlAxisType.xlValue) 'Y lines

        ' Axes Values and Format
        With c.Axes(Excel.XlAxisType.xlCategory)
            .MinimumScale = xFrom.Text
            .MaximumScale = xTo.Text
            .MajorUnit = xNumEvery.Text
            .MinorUnit = (xMinorGrid / majorGrid) * xNumEvery.Text
            .TickLabels.Format.Fill.BackColor.RGB = Word.WdColor.wdColorWhite
            .Format.Fill.BackColor.RGB = Word.WdColor.wdColorWhite
            .Format.Fill.Visible = Not NumNone.Checked
            If Ticks.Checked Then .MajorTickMark = Excel.XlTickMark.xlTickMarkOutside Else .MajorTickMark = Excel.XlTickMark.xlTickMarkNone
        End With

        With c.Axes(Excel.XlAxisType.xlValue)
            .MinimumScale = yFrom.Text
            .MaximumScale = yTo.Text
            .MajorUnit = yNumEvery.Text
            .MinorUnit = (yMinorGrid / majorGrid) * yNumEvery.Text
            .TickLabels.Format.Fill.BackColor.RGB = Word.WdColor.wdColorWhite
            .Format.Fill.BackColor.RGB = Word.WdColor.wdColorWhite
            .Format.Fill.Visible = Not NumNone.Checked
            If Ticks.Checked Then .MajorTickMark = Excel.XlTickMark.xlTickMarkOutside Else .MajorTickMark = Excel.XlTickMark.xlTickMarkNone
        End With


        c.HasAxis(Excel.XlAxisType.xlValue, Excel.XlAxisGroup.xlPrimary) = Axes.Checked
        c.HasAxis(Excel.XlAxisType.xlCategory, Excel.XlAxisGroup.xlPrimary) = Axes.Checked
        c.ChartArea.Format.TextFrame2.TextRange.Font.Fill.Visible = Not NumNone.Checked
        If AxisLabels.Checked Then
            c.SetElement(Office.MsoChartElementType.msoElementPrimaryCategoryAxisTitleAdjacentToAxis)
            c.SetElement(Office.MsoChartElementType.msoElementPrimaryValueAxisTitleHorizontal)
        End If
        Exit Sub

er:
        Select Case Err.Number
            Case Else
                MsgBox("Error in GraphMaker.PlotAsChart" & vbCrLf & Err.Number & vbCrLf & Err.Source & vbCrLf & Err.Description)
        End Select
    End Sub

    Sub FormatChartLines(c As Word.Chart, v As Excel.XlAxisType)

        'Dim a As Axis 'for Debugging

        With c.Axes(v)
            'Axis
            If c.HasAxis(v, Excel.XlAxisGroup.xlPrimary) = True Then 'check Axis not disabled

                If MajorWeight.Text <= 0 Then .Format.Line.Visible = Office.MsoTriState.msoFalse Else _
                    .Format.Line.Weight = MajorWeight.Text * 2
                .Format.Line.ForeColor = MajorColour.BackColor
                .Format.Line.DashStyle = Word.WdLineStyle.wdLineStyleSingle

            End If

            'Major
            If MajorWeight.Text <= 0 Then .MajorGridlines.Format.Line.Visible = False Else _
                .MajorGridlines.Format.Line.Weight = MajorWeight.Text
            .MajorGridlines.Format.Line.ForeColor = MajorColour.BackColor
            .MajorGridlines.Format.Line.DashStyle = MajorLineStyle.SelectedIndex + 1

            'Minor
            If MinorWeight.Text <= 0 Then .MinorGridlines.Format.Line.Visible = Office.MsoTriState.msoFalse Else _
                .MinorGridlines.Format.Line.Weight = MinorWeight.Text
            .MinorGridlines.Format.Line.ForeColor = MinorColour.BackColor
            .MinorGridlines.Format.Line.DashStyle = MinorLineStyle.SelectedIndex + 1

        End With

    End Sub


    ' == Plot as Shapes

    Sub PlotShapes()

        Dim GroupAll As New Collection
        Dim i As Integer
        Dim shpGroup As Word.Shape

        If MajorWeight.Text > 0 Then 'if visible
            PlotX_Major()
            PlotY_Major()
            GroupAll.Add("X Major")
            GroupAll.Add("Y Major")
            'DoEvents
        End If

        If MinorWeight.Text > 0 Then 'if visible
            If xDivs.Text > 0 Then
                PlotX_Minor()
                GroupAll.Add("X Minor")
            End If
            If yDivs.Text > 0 Then
                PlotY_Minor()
                GroupAll.Add("Y Minor")
            End If
            'DoEvents
        End If

        If Axes.Checked Then
            PlotAxes()
            GroupAll.Add("Axes")
        End If

        If AxisLabels.Checked Then
            PlotAxisLabels()
            GroupAll.Add("X Axis Label")
            GroupAll.Add("Y Axis Label")
        End If

        If Not NumNone.Checked Then
            PlotX_Numbers()
            PlotY_Numbers()
            GroupAll.Add("X Labels")
            GroupAll.Add("Y Labels")
            'DoEvents
        End If

        ' Group all
        Dim a() As Object
        ReDim a(GroupAll.Count - 1)
        For i = 0 To GroupAll.Count - 1
            a(i) = GroupAll(i + 1)
        Next i

        If i = 0 Then Exit Sub 'No shapes to group!
        shpGroup = ActiveDocument.Shapes.Range(a).Group
        shpGroup.Name = "Graph X" & xFrom.Text & "," & xTo.Text _
                        & " Y" & yFrom.Text & "," & yTo.Text


    End Sub

    Sub PlotX_Major()

        If Not DBG Then On Error GoTo er

        Dim shpLine As Word.Shape
        Dim shpGroup As Word.Shape

        Dim a() As Object
        Dim i As Integer
        Dim gap As Single
        Dim ln As Single
        Dim TopLeft As Single
        Dim TopDown As Single
        Dim BottomLeft As Single
        Dim BottomDown As Single

        gap = majorGrid

        ReDim a(xTotalMajors - 1)

        For i = 0 To xTotalMajors - 1

            ln = (gap * i) + graphLL(x)

            TopLeft = ln
            TopDown = graphLL(Y)
            BottomLeft = ln
            BottomDown = graphLL(Y) - graphHeight

            If Ticks.Checked Then TopDown = graphLL(Y) + GM_TICK_LENGTH

            shpLine = ActiveDocument.Shapes.AddLine(TopLeft, TopDown, BottomLeft, BottomDown)

            shpLine.Name = "X " & i

            a(i) = shpLine.Name

        Next i

        ' if just 1 shape
        ' set shpGroup to just that 1 shape
        If UBound(a) = 0 Then shpGroup = ActiveDocument.Shapes(a(0)) _
            Else shpGroup = ActiveDocument.Shapes.Range(a).Group

        shpGroup.Name = "X Major"
        FormatMajorStyle(shpGroup)
        Exit Sub

er:
        Select Case Err.Number
            Case Else
                Err.Raise(Err.Number)
        End Select
    End Sub

    Sub PlotY_Major()
        If Not DBG Then On Error GoTo er

        Dim shpLine As Word.Shape
        Dim shpGroup As Word.Shape
        Dim doc As Document

        Dim a() As Object
        Dim gap As Single
        Dim ln As Single
        Dim LeftAlong As Single
        Dim LeftDown As Single
        Dim RightAlong As Single
        Dim RightDown As Single

        gap = majorGrid

        ReDim a(yTotalMajors - 1)

        For i = 0 To yTotalMajors - 1

            ln = (gap * i)

            ' Construct bottom up to start major on zero axis
            LeftAlong = graphLL(x)
            LeftDown = graphLL(Y) - ln
            RightAlong = graphLL(x) + graphWidth
            RightDown = graphLL(Y) - ln

            If Ticks.Checked Then
                LeftAlong = graphLL(x) - GM_TICK_LENGTH
            End If

            shpLine = ActiveDocument.Shapes.AddLine(LeftAlong, LeftDown, RightAlong, RightDown)

            shpLine.Name = "Y " & i

            a(i) = shpLine.Name

        Next i

        ' if just 1 shape
        ' set shpGroup to just that 1 shape
        If UBound(a) = 0 Then shpGroup = ActiveDocument.Shapes(a(0)) _
            Else shpGroup = ActiveDocument.Shapes.Range(a).Group

        shpGroup.Name = "Y Major"
        FormatMajorStyle(shpGroup)

        Exit Sub

er:
        Select Case Err.Number
            Case Else
                Err.Raise(Err.Number)
        End Select
    End Sub

    Sub PlotX_Minor()
        If Not DBG Then On Error GoTo er

        Dim shpLine As Word.Shape
        Dim shpGroup As Word.Shape
        Dim a() As Object
        Dim gap As Single
        Dim ln As Single

        gap = xMinorGrid

        ReDim a(xTotalLines - 1)

        For i = 0 To xTotalLines - 1

            ln = (gap * i) + graphLL(x)

            'TopLeft, TopDown, BottomLeft, BottomDown
            shpLine = ActiveDocument.Shapes.AddLine(ln, GM_GUTTER, ln, graphLL(Y))
            shpLine.Name = "X " & i

            a(i) = shpLine.Name

        Next i

        ' if just 1 shape
        ' set shpGroup to just that 1 shape
        If UBound(a) = 0 Then shpGroup = ActiveDocument.Shapes(a(0)) _
            Else shpGroup = ActiveDocument.Shapes.Range(a).Group

        shpGroup.Name = "X Minor"
        FormatMinorStyle(shpGroup)

        Exit Sub

er:
        Select Case Err.Number
            Case Else
                Err.Raise(Err.Number)
        End Select
    End Sub

    Sub PlotY_Minor()
        If Not DBG Then On Error GoTo er

        Dim shpLine As Word.Shape
        Dim shpGroup As Word.Shape
        Dim doc As Document

        Dim a() As Object
        Dim gap As Single
        Dim ln As Single

        gap = yMinorGrid

        ReDim a(yTotalLines - 1)

        For i = 0 To yTotalLines - 1

            ln = (gap * i)

            ' Construct bottom up to start major on zero axis
            'LeftAlong, LeftDown, RightAlong, RightDown
            shpLine = ActiveDocument.Shapes.AddLine(graphLL(x), graphLL(Y) - ln, graphLL(x) + graphWidth, graphLL(Y) - ln)
            shpLine.Name = "Y " & i

            a(i) = shpLine.Name

        Next i

        ' if just 1 shape
        ' set shpGroup to just that 1 shape
        If UBound(a) = 0 Then shpGroup = ActiveDocument.Shapes(a(0)) _
            Else shpGroup = ActiveDocument.Shapes.Range(a).Group

        shpGroup.Name = "Y Minor"
        FormatMinorStyle(shpGroup)

        Exit Sub

er:
        Select Case Err.Number
            Case Else
                Err.Raise(Err.Number)
        End Select
    End Sub

    Sub PlotX_Numbers()

        If Not DBG Then On Error GoTo er

        Dim shp As Word.Shape
        Dim shpGroup As Word.Shape
        Dim doc As Document

        Dim a() As Object
        Dim n As Single
        Dim gap As Single
        Dim ln As Single

        ReDim a(xTotalMajors - 1)

        gap = majorGrid

        For i = 0 To xTotalMajors - 1
            ln = gap * i
            'TopLeft, TopDown, LabelWidth, LabelHeight
            shp = ActiveDocument.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal, GM_GUTTER + ln, graphOrigin(Y), gap, textHeight * 1.5)

            n = i
            If CSng(xFrom.Text) > CSng(xTo.Text) Then n = -i 'descending numbering
            n = xFrom.Text + (n * xNumEvery.Text)

            shp.Name = "X Label " & n
            shp.TextFrame.TextRange.Text = FormatNumbering(n)
            shp.TextFrame.TextRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
            shp.TextFrame.TextRange.Font.Shading.BackgroundPatternColor = Word.WdColor.wdColorWhite
            a(i) = shp.Name
        Next i

        ' if just 1 shape
        ' set shpGroup to just that 1 shape
        If UBound(a) = 0 Then shpGroup = ActiveDocument.Shapes(a(0)) _
            Else shpGroup = ActiveDocument.Shapes.Range(a).Group

        shpGroup.Name = "X Labels"
        shpGroup.Line.Visible = Office.MsoTriState.msoFalse
        shpGroup.Fill.Visible = Office.MsoTriState.msoFalse
        shpGroup.Left = shpGroup.Left - (gap / 2) 'minus half of label width
        If Ticks.Checked Then shpGroup.Top = shpGroup.Top + GM_TICK_LENGTH 'space for ticks

        Exit Sub

er:
        Select Case Err.Number
            Case Else
                Err.Raise(Err.Number)
        End Select
    End Sub

    Sub PlotY_Numbers()
        If Not DBG Then On Error GoTo er

        Dim shp As Word.Shape
        Dim shpGroup As Word.Shape
        Dim doc As Document

        Dim a() As Object
        Dim ln As Single
        Dim n As Single
        Dim gap As Single

        ' Zero based
        ReDim a(yTotalMajors - 1)


        gap = majorGrid

        For i = 0 To yTotalMajors - 1
            ln = gap * i
            'TopLeft, TopDown, LabelWidth, LabelHeight
            shp = ActiveDocument.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal, graphOrigin(x), graphLL(Y) - ln, textHeight * 1.5, gap)

            n = i
            If CSng(yFrom.Text) > CSng(yTo.Text) Then n = -i 'descending numbering
            n = yFrom.Text + (n * yNumEvery.Text)

            shp.Name = "Y Label " & n
            shp.TextFrame.TextRange.Text = FormatNumbering(n)
            shp.TextFrame.TextRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight
            shp.TextFrame.TextRange.ParagraphFormat.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceSingle 'for correct alignment with gridline
            shp.TextFrame.VerticalAnchor = Office.MsoVerticalAnchor.msoAnchorMiddle 'for correct alignment with gridline
            shp.TextFrame.TextRange.Font.Shading.BackgroundPatternColor = Word.WdColor.wdColorWhite
            a(i) = shp.Name
        Next i

        ' if just 1 shape
        ' set shpGroup to just that 1 shape
        If UBound(a) = 0 Then shpGroup = ActiveDocument.Shapes(a(0)) _
            Else shpGroup = ActiveDocument.Shapes.Range(a).Group

        shpGroup.Name = "Y Labels"
        shpGroup.Line.Visible = Office.MsoTriState.msoFalse
        shpGroup.Fill.Visible = Office.MsoTriState.msoFalse


        'reposition
        shpGroup.Top = shpGroup.Top - (gap / 2) 'half label height
        shpGroup.Left = shpGroup.Left - (textHeight * 1.5) 'label width
        If Ticks.Checked Then shpGroup.Left = shpGroup.Left - GM_TICK_LENGTH 'space for ticks

        Exit Sub

er:
        Select Case Err.Number
            Case Else
                Err.Raise(Err.Number)
        End Select
    End Sub

    Sub PlotAxes()

        Dim shp As Word.Shape
        Dim shpGroup As Word.Shape
        Dim a() As Object
        Dim lineLength As Single


        ReDim a(1)
        'Y
        lineLength = graphLL(Y) - (graphHeight + GM_AXISHEAD)

        shp = ActiveDocument.Shapes.AddLine(graphOrigin(x), graphLL(Y), graphOrigin(x), lineLength)
        shp.Name = "Y Axis"
        a(0) = shp.Name

        'X
        lineLength = graphLL(x) + graphWidth + GM_AXISHEAD

        shp = ActiveDocument.Shapes.AddLine(graphLL(x), graphOrigin(Y), lineLength, graphOrigin(Y))
        shp.Name = "X Axis"
        a(1) = shp.Name


        'Grouping
        shpGroup = ActiveDocument.Shapes.Range(a).Group
        shpGroup.Name = "Axes"

        If MajorWeight.Text <= 0 Then shpGroup.Line.Weight = "6" _
        Else shpGroup.Line.Weight = MajorWeight.Text * 2        'might be invisible

        shpGroup.Line.BeginArrowheadStyle = Office.MsoArrowheadStyle.msoArrowheadNone
        shpGroup.Line.EndArrowheadStyle = Office.MsoArrowheadStyle.msoArrowheadTriangle
        shpGroup.Line.ForeColor.RGB = RGB(0, 0, 0) 'always black.


    End Sub

    Sub PlotAxisLabels()

        Dim shpGroup As Word.Shape
        Dim shp As Word.Shape
        Dim TopLeft As Single
        Dim TopDown As Single

        'X
        TopLeft = graphLL(x) + graphWidth + GM_AXISHEAD
        TopDown = graphOrigin(Y) - (textHeight / 2)

        'TopLeft, TopDown, LabelWidth, LabelHeight
        shp = ActiveDocument.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal, TopLeft, TopDown, textHeight, textHeight)

        shp.Name = "X Axis Label"
        shp.TextFrame.TextRange.Text = "X"
        shp.TextFrame.TextRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
        shp.TextFrame.TextRange.Font.Shading.BackgroundPatternColor = Word.WdColor.wdColorWhite
        shp.Fill.Visible = Office.MsoTriState.msoFalse
        shp.Line.Visible = Office.MsoTriState.msoFalse

        'Y
        TopLeft = graphOrigin(x) - (textHeight / 2)
        TopDown = graphLL(Y) - (graphHeight + GM_AXISHEAD + textHeight)

        'TopLeft, TopDown, LabelWidth, LabelHeight
        shp = ActiveDocument.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal, TopLeft, TopDown, textHeight, textHeight)

        shp.Name = "Y Axis Label"
        shp.TextFrame.TextRange.Text = "Y"
        shp.TextFrame.TextRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
        shp.TextFrame.TextRange.Font.Shading.BackgroundPatternColor = Word.WdColor.wdColorWhite
        shp.Fill.Visible = Office.MsoTriState.msoFalse
        shp.Line.Visible = Office.MsoTriState.msoFalse


    End Sub

    Sub FormatMajorStyle(shp As Word.Shape)

        shp.Line.Weight = MajorWeight.Text
        shp.Line.ForeColor.RGB = RGB(MajorColour.BackColor.R, MajorColour.BackColor.G, MajorColour.BackColor.B)
        shp.Line.DashStyle = MajorLineStyle.SelectedIndex + 1
        shp.Line.BeginArrowheadStyle = Office.MsoArrowheadStyle.msoArrowheadNone
        shp.Line.EndArrowheadStyle = Office.MsoArrowheadStyle.msoArrowheadNone
        shp.ZOrder(Office.MsoZOrderCmd.msoBringToFront)
    End Sub

    Sub FormatMinorStyle(shp As Word.Shape)

        shp.Line.Weight = MinorWeight.Text
        shp.Line.ForeColor.RGB = RGB(MinorColour.BackColor.R, MinorColour.BackColor.G, MinorColour.BackColor.B)
        shp.Line.DashStyle = MinorLineStyle.SelectedIndex + 1
        shp.Line.BeginArrowheadStyle = Office.MsoArrowheadStyle.msoArrowheadNone
        shp.Line.EndArrowheadStyle = Office.MsoArrowheadStyle.msoArrowheadNone
        shp.ZOrder(Office.MsoZOrderCmd.msoSendToBack)
    End Sub

    Function FormatNumbering(number As Single) As String
        ' Return UEB Braille coded string for the number or just return the number as is.

        If NumStandard.Checked Then
            FormatNumbering = number
            Exit Function
        End If

        Dim s As String
        Dim n As String

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
                    Err.Raise(Err.Number)
            End Select

        Next i

        FormatNumbering = s

    End Function


End Class