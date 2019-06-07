Imports System.ComponentModel


Public Class GraphMaker

    'Protected Friend app As Word.Application

    ' TODO

    ' .Axis Labels
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
        MajorColour.BackColor = SavedCtls.Default.majorColour
        MajorLineStyle.SelectedIndex = SavedCtls.Default.MajorLineStyle

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
        SavedCtls.Default.majorColour = ColorDialog1.Color

    End Sub


    Private Sub Num_Validating(sender As Object, e As CancelEventArgs) _
        Handles xFrom.Validating, xTo.Validating,
                yFrom.Validating, yTo.Validating,
                MajorWeight.Validating
        e.Cancel = ErrNotANumber(sender)
    End Sub

    Private Sub Val_Validated(sender As Object, e As EventArgs) _
        Handles xFrom.Validating, xTo.Validated,
                yFrom.Validated, yTo.Validated,
                MajorWeight.Validated
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

        MajorLineStyle.Items.AddRange(WordLineStyles)
    End Sub

    Private Sub btnCreate_Click(sender As Object, e As EventArgs) Handles btnCreate.Click
        PlotGraph()
    End Sub


    Friend Sub PlotGraph()
        'On Error GoTo er

        Dim objUndo As Word.UndoRecord
        objUndo = App.UndoRecord
        objUndo.StartCustomRecord("Graph Maker")


        App.System.Cursor = Word.WdCursorType.wdCursorWait

        Dim ActiveDocument As Word.Document
        ActiveDocument = App.ActiveDocument

        ' === Calculate units

        Dim pgWidth As Single
        Dim pgHeight As Single
        Dim xRange As Single
        Dim yRange As Single

        xRange = Math.Abs(xFrom.Text - xTo.Text)
        yRange = Math.Abs(yFrom.Text - yTo.Text)

        textHeight = ActiveDocument.Styles("Normal").Font.Size * 2

        xTotalMajors = Int((xRange / xNumEvery.Text)) + 1
        ' Combine all elements and roundup to next whole number. It works but don't quite know how!
        xTotalDivs = RoundUp(xRange / xNumEvery.Text * xDivs.Text)
        xTotalLines = xTotalMajors + xTotalDivs


        yTotalMajors = Int((yRange / yNumEvery.Text)) + 1
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

        ' Digit gap used to calculate Origin (0,0)
        If graphWidth = 0 Then xDigitGap = 0 _
             Else xDigitGap = graphWidth / xRange

        If graphHeight = 0 Then yDigitGap = 0 _
             Else yDigitGap = graphHeight / yRange


        ' === Calculate Origin (0,0)
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

        ' === Plot Graph

        ' If PlotAsChart Then PlotAsChart() Else : PlotAsShapes()
        If PlotAsChart.Checked Then MsgBox("PlotAsChart()") Else : MsgBox("PlotAsShapes()")

        objUndo.EndCustomRecord()
        app.System.Cursor = Word.WdCursorType.wdCursorNormal
        Exit Sub

er:
        objUndo.EndCustomRecord()
        app.System.Cursor = Word.WdCursorType.wdCursorNormal
        Select Case Err.Number
            Case Else
                MsgBox("Error in GraphMaker" & vbCrLf & Err.Number & vbCrLf & Err.Source & vbCrLf & Err.Description)
        End Select
    End Sub

End Class