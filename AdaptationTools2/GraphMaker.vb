Imports System.ComponentModel


Public Class GraphMaker

    ' TODO

    ' .Axis Labels
    ' .Redo Ticks
    ' .Plan different profiles
    Const DBG As Boolean = True

    '500 ~= width of A4 document
    Const GM_GUTTER As Integer = 100
    Const GM_AXISHEAD As Integer = 30
    Const GM_TICK_LENGTH As Integer = 15

    Dim graphLL(1) As Single ' Lower Left corner
    Dim graphOrigin(1) As Single
    Const x As Integer = 0
    Const Y As Integer = 1

    Dim ActiveDocument As Word.Document

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

        'TODO How to save Radio groups?
        'Dim r As Windows.Forms.RadioButton
        'r = Controls.Item(Controls.IndexOfKey(SavedCtls.Default.PlotAs.ToString))
        'r.Checked = True



    End Sub


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

    Private Sub majorColour_Click(sender As Object, e As EventArgs) Handles MajorColour.Click
        ColorDialog1.ShowDialog()

        MajorColour.BackColor = ColorDialog1.Color
        SavedCtls.Default.MajorColour = ColorDialog1.Color

    End Sub


    Private Sub Num_Validating(sender As Object, e As CancelEventArgs) Handles xFrom.Validating, xTo.Validating, yFrom.Validating, yTo.Validating, MajorWeight.Validating, MinorWeight.Validating, yNumEvery.Validating, yDivs.Validating, xNumEvery.Validating, xDivs.Validating
        e.Cancel = ErrNotANumber(sender)
    End Sub

    Private Sub Val_Validated(sender As Object, e As EventArgs) Handles xFrom.Validating, xTo.Validated, yFrom.Validated, yTo.Validated, MajorWeight.Validated, MinorWeight.Validated, yNumEvery.Validated, yDivs.Validated, xNumEvery.Validated, xDivs.Validated
        SavedCtls.Default.PropertyValues.Item(sender.Name).PropertyValue = sender.Text
    End Sub


    Private Sub MajorLineStyle_SelectedIndexChanged(sender As Object, e As EventArgs) Handles MajorLineStyle.SelectedIndexChanged
        SavedCtls.Default.MajorLineStyle = MajorLineStyle.SelectedIndex
    End Sub



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

        MajorLineStyle.Items.AddRange(WordLineStyles)
        MinorLineStyle.Items.AddRange(WordLineStyles)
    End Sub


    Public Function LineDashStyleID(name As String) As Integer
        Dim d As Integer
        Select Case name
            Case "Solid" : d = 1
            Case "Mixed" : d = -2
            Case "Square Dot" : d = 2
            Case "Round Dot" : d = 3
            Case "Dash" : d = 4
            Case "Dash Dot" : d = 5
            Case "Dash Dot Dot" : d = 6
            Case "Long Dash" : d = 7
            Case "Long Dash Dot" : d = 8
            Case "Long Dash Dot Dot" : d = 9
            Case "Sys Dash" : d = 10
            Case "Sys Dot" : d = 11
            Case "Sys Dash Dot" : d = 12
        End Select
        LineDashStyleID = d
    End Function

    Public Function LineDashStyleName(id As Integer) As String
        Dim s As Integer
        Select Case id
            Case 1 : s = "Solid"
            Case -2 : s = "Mixed"
            Case 2 : s = "Square Dot"
            Case 3 : s = "Round Dot"
            Case 4 : s = "Dash"
            Case 5 : s = "Dash Dot"
            Case 6 : s = "Dash Dot Dot"
            Case 7 : s = "Long Dash"
            Case 8 : s = "Long Dash Dot"
            Case 9 : s = "Long Dash Dot Dot"
            Case 10 : s = "Sys Dash"
            Case 11 : s = "Sys Dot"
            Case 12 : s = "Sys Dash Dot"
        End Select
        LineDashStyleName = s
    End Function

    Private Sub btnCreate_Click(sender As Object, e As EventArgs) Handles btnCreate.Click
        PlotGraph()
    End Sub


    Friend Sub PlotGraph()
        If Not DBG Then On Error GoTo er

        Dim objUndo As Word.UndoRecord
        objUndo = App.UndoRecord
        objUndo.StartCustomRecord("Graph Maker")

        App.System.Cursor = Word.WdCursorType.wdCursorWait

        ActiveDocument = App.ActiveDocument

        CalculateUnits()

        ' === Plot Graph

        If PlotAsChart.Checked Then PlotChart() Else MsgBox("PlotShapes()")
        'If PlotAsChart.Checked Then MsgBox("PlotAsChart()") Else MsgBox("PlotAsShapes()")

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

                'If MajorColour.BackColor < 0 Then .Format.Line.Visible = Office.MsoTriState.msoFalse Else _
                .Format.Line.ForeColor = MajorColour.BackColor

                .Format.Line.DashStyle = Word.WdLineStyle.wdLineStyleSingle

            End If

            'Major
            If MajorWeight.Text <= 0 Then .MajorGridlines.Format.Line.Visible = False Else _
                .MajorGridlines.Format.Line.Weight = MajorWeight.Text

            'If MajorColour.BackColor < 0 Then .MajorGridlines.Format.Line.Visible = Office.MsoTriState.msoFalse Else _
            .MajorGridlines.Format.Line.ForeColor = MajorColour.BackColor

            'If majorDash <> "Mixed" Then .MajorGridlines.Format.Line.DashStyle = LineDashStyleID(majorDash)
            .MajorGridlines.Format.Line.DashStyle = MajorLineStyle.SelectedIndex


            'Minor
            If MinorWeight.Text <= 0 Then .MinorGridlines.Format.Line.Visible = Office.MsoTriState.msoFalse Else _
                .MinorGridlines.Format.Line.Weight = MinorWeight.Text

            'If minorColour.BackColor < 0 Then .MinorGridlines.Format.Line.Visible = Office.MsoTriState.msoFalse Else _
            .MinorGridlines.Format.Line.ForeColor = MinorColour.BackColor

            'If minorDash <> "Mixed" Then .MinorGridlines.Format.Line.DashStyle = LineDashStyleID(minorDash)
            .MinorGridlines.Format.Line.DashStyle = MajorLineStyle.SelectedIndex

        End With

    End Sub

    Private Sub PlotAs_CheckedChanged(sender As Object, e As EventArgs) Handles PlotAsChart.Validated, PlotAsShapes.Validated
        'TODO
        SavedCtls.Default.PropertyValues.Item("PlotAs").PropertyValue = sender.Name

    End Sub
End Class