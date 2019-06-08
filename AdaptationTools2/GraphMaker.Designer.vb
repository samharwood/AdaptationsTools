<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class GraphMaker
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(GraphMaker))
        Me.ColorDialog1 = New System.Windows.Forms.ColorDialog()
        Me.MajorColour = New System.Windows.Forms.Label()
        Me.xTo = New System.Windows.Forms.TextBox()
        Me.yFrom = New System.Windows.Forms.TextBox()
        Me.yTo = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.MajorWeight = New System.Windows.Forms.TextBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Axes = New System.Windows.Forms.CheckBox()
        Me.AxisLabels = New System.Windows.Forms.CheckBox()
        Me.NumStandard = New System.Windows.Forms.RadioButton()
        Me.GrpNumbering = New System.Windows.Forms.GroupBox()
        Me.NumNone = New System.Windows.Forms.RadioButton()
        Me.NumUEB = New System.Windows.Forms.RadioButton()
        Me.MajorLineStyle = New System.Windows.Forms.ComboBox()
        Me.xFrom = New System.Windows.Forms.TextBox()
        Me.CopyMajorStyle = New System.Windows.Forms.Button()
        Me.SaveToDocChk = New System.Windows.Forms.CheckBox()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.btnCreate = New System.Windows.Forms.Button()
        Me.Ticks = New System.Windows.Forms.CheckBox()
        Me.GrpPlotAs = New System.Windows.Forms.GroupBox()
        Me.PlotAsChart = New System.Windows.Forms.RadioButton()
        Me.PlotAsShapes = New System.Windows.Forms.RadioButton()
        Me.yDivs = New System.Windows.Forms.TextBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.xDivs = New System.Windows.Forms.TextBox()
        Me.xNumEvery = New System.Windows.Forms.TextBox()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.yNumEvery = New System.Windows.Forms.TextBox()
        Me.CopyMinorStyle = New System.Windows.Forms.Button()
        Me.MinorLineStyle = New System.Windows.Forms.ComboBox()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.MinorWeight = New System.Windows.Forms.TextBox()
        Me.MinorColour = New System.Windows.Forms.Label()
        Me.GrpNumbering.SuspendLayout()
        Me.GrpPlotAs.SuspendLayout()
        Me.SuspendLayout()
        '
        'ColorDialog1
        '
        Me.ColorDialog1.FullOpen = True
        '
        'MajorColour
        '
        Me.MajorColour.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.MajorColour.Cursor = System.Windows.Forms.Cursors.Hand
        Me.MajorColour.Location = New System.Drawing.Point(349, 315)
        Me.MajorColour.Name = "MajorColour"
        Me.MajorColour.Size = New System.Drawing.Size(30, 20)
        Me.MajorColour.TabIndex = 0
        '
        'xTo
        '
        Me.xTo.Location = New System.Drawing.Point(257, 73)
        Me.xTo.Name = "xTo"
        Me.xTo.Size = New System.Drawing.Size(34, 20)
        Me.xTo.TabIndex = 1
        '
        'yFrom
        '
        Me.yFrom.Location = New System.Drawing.Point(222, 99)
        Me.yFrom.Name = "yFrom"
        Me.yFrom.Size = New System.Drawing.Size(34, 20)
        Me.yFrom.TabIndex = 2
        '
        'yTo
        '
        Me.yTo.Location = New System.Drawing.Point(222, 47)
        Me.yTo.Name = "yTo"
        Me.yTo.Size = New System.Drawing.Size(34, 20)
        Me.yTo.TabIndex = 3
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(144, 76)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(37, 13)
        Me.Label1.TabIndex = 6
        Me.Label1.Text = "X from"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(231, 76)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(16, 13)
        Me.Label2.TabIndex = 7
        Me.Label2.Text = "to"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(179, 102)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(37, 13)
        Me.Label3.TabIndex = 8
        Me.Label3.Text = "Y from"
        '
        'MajorWeight
        '
        Me.MajorWeight.Location = New System.Drawing.Point(182, 316)
        Me.MajorWeight.Name = "MajorWeight"
        Me.MajorWeight.Size = New System.Drawing.Size(34, 20)
        Me.MajorWeight.TabIndex = 9
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(13, 319)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(59, 13)
        Me.Label4.TabIndex = 10
        Me.Label4.Text = "Major Style"
        '
        'Axes
        '
        Me.Axes.AutoSize = True
        Me.Axes.Location = New System.Drawing.Point(432, 65)
        Me.Axes.Name = "Axes"
        Me.Axes.Size = New System.Drawing.Size(49, 17)
        Me.Axes.TabIndex = 11
        Me.Axes.Text = "Axes"
        Me.Axes.UseVisualStyleBackColor = True
        '
        'AxisLabels
        '
        Me.AxisLabels.AutoSize = True
        Me.AxisLabels.Location = New System.Drawing.Point(432, 88)
        Me.AxisLabels.Name = "AxisLabels"
        Me.AxisLabels.Size = New System.Drawing.Size(79, 17)
        Me.AxisLabels.TabIndex = 12
        Me.AxisLabels.Text = "Axis Labels"
        Me.AxisLabels.UseVisualStyleBackColor = True
        '
        'NumStandard
        '
        Me.NumStandard.AutoSize = True
        Me.NumStandard.Checked = True
        Me.NumStandard.Location = New System.Drawing.Point(6, 19)
        Me.NumStandard.Name = "NumStandard"
        Me.NumStandard.Size = New System.Drawing.Size(68, 17)
        Me.NumStandard.TabIndex = 15
        Me.NumStandard.TabStop = True
        Me.NumStandard.Text = "Standard"
        Me.NumStandard.UseVisualStyleBackColor = True
        '
        'GrpNumbering
        '
        Me.GrpNumbering.Controls.Add(Me.NumNone)
        Me.GrpNumbering.Controls.Add(Me.NumUEB)
        Me.GrpNumbering.Controls.Add(Me.NumStandard)
        Me.GrpNumbering.Location = New System.Drawing.Point(536, 65)
        Me.GrpNumbering.Name = "GrpNumbering"
        Me.GrpNumbering.Size = New System.Drawing.Size(89, 92)
        Me.GrpNumbering.TabIndex = 16
        Me.GrpNumbering.TabStop = False
        Me.GrpNumbering.Text = "Numbering"
        '
        'NumNone
        '
        Me.NumNone.AutoSize = True
        Me.NumNone.Location = New System.Drawing.Point(6, 65)
        Me.NumNone.Name = "NumNone"
        Me.NumNone.Size = New System.Drawing.Size(51, 17)
        Me.NumNone.TabIndex = 17
        Me.NumNone.Text = "None"
        Me.NumNone.UseVisualStyleBackColor = True
        '
        'NumUEB
        '
        Me.NumUEB.AutoSize = True
        Me.NumUEB.Location = New System.Drawing.Point(6, 42)
        Me.NumUEB.Name = "NumUEB"
        Me.NumUEB.Size = New System.Drawing.Size(78, 17)
        Me.NumUEB.TabIndex = 16
        Me.NumUEB.Text = "UEB Braille"
        Me.NumUEB.UseVisualStyleBackColor = True
        '
        'MajorLineStyle
        '
        Me.MajorLineStyle.DropDownWidth = 200
        Me.MajorLineStyle.Location = New System.Drawing.Point(222, 315)
        Me.MajorLineStyle.Name = "MajorLineStyle"
        Me.MajorLineStyle.Size = New System.Drawing.Size(121, 21)
        Me.MajorLineStyle.TabIndex = 17
        '
        'xFrom
        '
        Me.xFrom.Location = New System.Drawing.Point(187, 73)
        Me.xFrom.Name = "xFrom"
        Me.xFrom.Size = New System.Drawing.Size(34, 20)
        Me.xFrom.TabIndex = 0
        '
        'CopyMajorStyle
        '
        Me.CopyMajorStyle.Location = New System.Drawing.Point(78, 309)
        Me.CopyMajorStyle.Name = "CopyMajorStyle"
        Me.CopyMajorStyle.Size = New System.Drawing.Size(98, 37)
        Me.CopyMajorStyle.TabIndex = 18
        Me.CopyMajorStyle.Text = "Copy from selected shape"
        Me.CopyMajorStyle.UseVisualStyleBackColor = True
        '
        'SaveToDocChk
        '
        Me.SaveToDocChk.AccessibleDescription = ""
        Me.SaveToDocChk.AutoSize = True
        Me.SaveToDocChk.Location = New System.Drawing.Point(12, 421)
        Me.SaveToDocChk.Name = "SaveToDocChk"
        Me.SaveToDocChk.Size = New System.Drawing.Size(187, 17)
        Me.SaveToDocChk.TabIndex = 19
        Me.SaveToDocChk.Text = "Save/Load settings per document"
        Me.ToolTip1.SetToolTip(Me.SaveToDocChk, resources.GetString("SaveToDocChk.ToolTip"))
        '
        'ToolTip1
        '
        Me.ToolTip1.AutoPopDelay = 20000
        Me.ToolTip1.InitialDelay = 500
        Me.ToolTip1.ReshowDelay = 100
        Me.ToolTip1.ShowAlways = True
        '
        'btnCreate
        '
        Me.btnCreate.Location = New System.Drawing.Point(690, 401)
        Me.btnCreate.Name = "btnCreate"
        Me.btnCreate.Size = New System.Drawing.Size(98, 37)
        Me.btnCreate.TabIndex = 20
        Me.btnCreate.Text = "Create"
        Me.btnCreate.UseVisualStyleBackColor = True
        '
        'Ticks
        '
        Me.Ticks.AutoSize = True
        Me.Ticks.Location = New System.Drawing.Point(432, 111)
        Me.Ticks.Name = "Ticks"
        Me.Ticks.Size = New System.Drawing.Size(52, 17)
        Me.Ticks.TabIndex = 21
        Me.Ticks.Text = "Ticks"
        Me.Ticks.UseVisualStyleBackColor = True
        '
        'GrpPlotAs
        '
        Me.GrpPlotAs.Controls.Add(Me.PlotAsChart)
        Me.GrpPlotAs.Controls.Add(Me.PlotAsShapes)
        Me.GrpPlotAs.Location = New System.Drawing.Point(690, 315)
        Me.GrpPlotAs.Name = "GrpPlotAs"
        Me.GrpPlotAs.Size = New System.Drawing.Size(89, 69)
        Me.GrpPlotAs.TabIndex = 18
        Me.GrpPlotAs.TabStop = False
        Me.GrpPlotAs.Text = "Plot as"
        '
        'PlotAsChart
        '
        Me.PlotAsChart.AutoSize = True
        Me.PlotAsChart.Location = New System.Drawing.Point(6, 42)
        Me.PlotAsChart.Name = "PlotAsChart"
        Me.PlotAsChart.Size = New System.Drawing.Size(50, 17)
        Me.PlotAsChart.TabIndex = 16
        Me.PlotAsChart.Text = "Chart"
        Me.PlotAsChart.UseVisualStyleBackColor = True
        '
        'PlotAsShapes
        '
        Me.PlotAsShapes.AutoSize = True
        Me.PlotAsShapes.Checked = True
        Me.PlotAsShapes.Location = New System.Drawing.Point(6, 19)
        Me.PlotAsShapes.Name = "PlotAsShapes"
        Me.PlotAsShapes.Size = New System.Drawing.Size(61, 17)
        Me.PlotAsShapes.TabIndex = 15
        Me.PlotAsShapes.TabStop = True
        Me.PlotAsShapes.Text = "Shapes"
        Me.PlotAsShapes.UseVisualStyleBackColor = True
        '
        'yDivs
        '
        Me.yDivs.Location = New System.Drawing.Point(262, 199)
        Me.yDivs.Name = "yDivs"
        Me.yDivs.Size = New System.Drawing.Size(34, 20)
        Me.yDivs.TabIndex = 22
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(168, 202)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(88, 13)
        Me.Label5.TabIndex = 23
        Me.Label5.Text = "Minor Y Divisions"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(168, 176)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(88, 13)
        Me.Label6.TabIndex = 24
        Me.Label6.Text = "Minor X Divisions"
        '
        'xDivs
        '
        Me.xDivs.Location = New System.Drawing.Point(262, 173)
        Me.xDivs.Name = "xDivs"
        Me.xDivs.Size = New System.Drawing.Size(34, 20)
        Me.xDivs.TabIndex = 25
        '
        'xNumEvery
        '
        Me.xNumEvery.Location = New System.Drawing.Point(115, 173)
        Me.xNumEvery.Name = "xNumEvery"
        Me.xNumEvery.Size = New System.Drawing.Size(34, 20)
        Me.xNumEvery.TabIndex = 29
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(25, 176)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(84, 13)
        Me.Label7.TabIndex = 28
        Me.Label7.Text = "Number X Every"
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(25, 202)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(84, 13)
        Me.Label8.TabIndex = 27
        Me.Label8.Text = "Number Y Every"
        '
        'yNumEvery
        '
        Me.yNumEvery.Location = New System.Drawing.Point(115, 199)
        Me.yNumEvery.Name = "yNumEvery"
        Me.yNumEvery.Size = New System.Drawing.Size(34, 20)
        Me.yNumEvery.TabIndex = 26
        '
        'CopyMinorStyle
        '
        Me.CopyMinorStyle.Location = New System.Drawing.Point(78, 357)
        Me.CopyMinorStyle.Name = "CopyMinorStyle"
        Me.CopyMinorStyle.Size = New System.Drawing.Size(98, 37)
        Me.CopyMinorStyle.TabIndex = 34
        Me.CopyMinorStyle.Text = "Copy from selected shape"
        Me.CopyMinorStyle.UseVisualStyleBackColor = True
        '
        'MinorLineStyle
        '
        Me.MinorLineStyle.DropDownWidth = 200
        Me.MinorLineStyle.Location = New System.Drawing.Point(222, 363)
        Me.MinorLineStyle.Name = "MinorLineStyle"
        Me.MinorLineStyle.Size = New System.Drawing.Size(121, 21)
        Me.MinorLineStyle.TabIndex = 33
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(13, 367)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(59, 13)
        Me.Label9.TabIndex = 32
        Me.Label9.Text = "Minor Style"
        '
        'MinorWeight
        '
        Me.MinorWeight.Location = New System.Drawing.Point(182, 364)
        Me.MinorWeight.Name = "MinorWeight"
        Me.MinorWeight.Size = New System.Drawing.Size(34, 20)
        Me.MinorWeight.TabIndex = 31
        '
        'MinorColour
        '
        Me.MinorColour.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.MinorColour.Cursor = System.Windows.Forms.Cursors.Hand
        Me.MinorColour.Location = New System.Drawing.Point(349, 363)
        Me.MinorColour.Name = "MinorColour"
        Me.MinorColour.Size = New System.Drawing.Size(30, 20)
        Me.MinorColour.TabIndex = 30
        '
        'GraphMaker
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(800, 450)
        Me.Controls.Add(Me.CopyMinorStyle)
        Me.Controls.Add(Me.MinorLineStyle)
        Me.Controls.Add(Me.Label9)
        Me.Controls.Add(Me.MinorWeight)
        Me.Controls.Add(Me.MinorColour)
        Me.Controls.Add(Me.xNumEvery)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.Label8)
        Me.Controls.Add(Me.yNumEvery)
        Me.Controls.Add(Me.xDivs)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.yDivs)
        Me.Controls.Add(Me.GrpPlotAs)
        Me.Controls.Add(Me.Ticks)
        Me.Controls.Add(Me.btnCreate)
        Me.Controls.Add(Me.SaveToDocChk)
        Me.Controls.Add(Me.CopyMajorStyle)
        Me.Controls.Add(Me.MajorLineStyle)
        Me.Controls.Add(Me.GrpNumbering)
        Me.Controls.Add(Me.AxisLabels)
        Me.Controls.Add(Me.Axes)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.MajorWeight)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.yTo)
        Me.Controls.Add(Me.yFrom)
        Me.Controls.Add(Me.xTo)
        Me.Controls.Add(Me.xFrom)
        Me.Controls.Add(Me.MajorColour)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "GraphMaker"
        Me.ShowIcon = False
        Me.Text = "Graph Maker"
        Me.GrpNumbering.ResumeLayout(False)
        Me.GrpNumbering.PerformLayout()
        Me.GrpPlotAs.ResumeLayout(False)
        Me.GrpPlotAs.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents ColorDialog1 As Windows.Forms.ColorDialog
    Friend WithEvents MajorColour As Windows.Forms.Label
    Friend WithEvents xFrom As Windows.Forms.TextBox
    Friend WithEvents xTo As Windows.Forms.TextBox
    Friend WithEvents yFrom As Windows.Forms.TextBox
    Friend WithEvents yTo As Windows.Forms.TextBox
    Friend WithEvents Label1 As Windows.Forms.Label
    Friend WithEvents Label2 As Windows.Forms.Label
    Friend WithEvents Label3 As Windows.Forms.Label
    Friend WithEvents MajorWeight As Windows.Forms.TextBox
    Friend WithEvents Label4 As Windows.Forms.Label
    Friend WithEvents Axes As Windows.Forms.CheckBox
    Friend WithEvents AxisLabels As Windows.Forms.CheckBox
    Friend WithEvents NumStandard As Windows.Forms.RadioButton
    Friend WithEvents GrpNumbering As Windows.Forms.GroupBox
    Friend WithEvents NumNone As Windows.Forms.RadioButton
    Friend WithEvents NumUEB As Windows.Forms.RadioButton
    Friend WithEvents MajorLineStyle As Windows.Forms.ComboBox
    Friend WithEvents CopyMajorStyle As Windows.Forms.Button
    Friend WithEvents SaveToDocChk As Windows.Forms.CheckBox
    Friend WithEvents ToolTip1 As Windows.Forms.ToolTip
    Friend WithEvents btnCreate As Windows.Forms.Button
    Friend WithEvents Ticks As Windows.Forms.CheckBox
    Friend WithEvents GrpPlotAs As Windows.Forms.GroupBox
    Friend WithEvents PlotAsChart As Windows.Forms.RadioButton
    Friend WithEvents PlotAsShapes As Windows.Forms.RadioButton
    Friend WithEvents yDivs As Windows.Forms.TextBox
    Friend WithEvents Label5 As Windows.Forms.Label
    Friend WithEvents Label6 As Windows.Forms.Label
    Friend WithEvents xDivs As Windows.Forms.TextBox
    Friend WithEvents xNumEvery As Windows.Forms.TextBox
    Friend WithEvents Label7 As Windows.Forms.Label
    Friend WithEvents Label8 As Windows.Forms.Label
    Friend WithEvents yNumEvery As Windows.Forms.TextBox
    Friend WithEvents CopyMinorStyle As Windows.Forms.Button
    Friend WithEvents MinorLineStyle As Windows.Forms.ComboBox
    Friend WithEvents Label9 As Windows.Forms.Label
    Friend WithEvents MinorWeight As Windows.Forms.TextBox
    Friend WithEvents MinorColour As Windows.Forms.Label
End Class
