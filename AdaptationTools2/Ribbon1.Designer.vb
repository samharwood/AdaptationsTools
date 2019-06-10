Partial Class Ribbon1
    Inherits Microsoft.Office.Tools.Ribbon.RibbonBase

    <System.Diagnostics.DebuggerNonUserCode()>
    Public Sub New(ByVal container As System.ComponentModel.IContainer)
        MyClass.New()

        'Required for Windows.Forms Class Composition Designer support
        If (container IsNot Nothing) Then
            container.Add(Me)
        End If

    End Sub

    <Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2214:DoNotCallOverridableMethodsInConstructors")>
    <System.Diagnostics.DebuggerNonUserCode()>
    Public Sub New()
        MyBase.New(Globals.Factory.GetRibbonFactory())

        'This call is required by the Component Designer.
        InitializeComponent()

    End Sub

    'Component overrides dispose to clean up the component list.
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

    'Required by the Component Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Component Designer
    'It can be modified using the Component Designer.
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.Group2 = Me.Factory.CreateRibbonGroup
        Me.TextBoundariesChk = Me.Factory.CreateRibbonCheckBox
        Me.Group4 = Me.Factory.CreateRibbonGroup
        Me.Group3 = Me.Factory.CreateRibbonGroup
        Me.AdTab1 = Me.Factory.CreateRibbonTab
        Me.Group1 = Me.Factory.CreateRibbonGroup
        Me.GraphMakerBtn = Me.Factory.CreateRibbonButton
        Me.ThickenLinesBtn = Me.Factory.CreateRibbonButton
        Me.MathsToTextBtn = Me.Factory.CreateRibbonButton
        Me.MathsToBrailleBtn = Me.Factory.CreateRibbonButton
        Me.AboutBtn = Me.Factory.CreateRibbonButton
        Me.HelpBtn = Me.Factory.CreateRibbonButton
        Me.Button1 = Me.Factory.CreateRibbonButton
        Me.ImageList1 = New System.Windows.Forms.ImageList(Me.components)
        Me.PasteFromPDF_Btn = Me.Factory.CreateRibbonButton
        Me.Group2.SuspendLayout()
        Me.Group4.SuspendLayout()
        Me.Group3.SuspendLayout()
        Me.AdTab1.SuspendLayout()
        Me.Group1.SuspendLayout()
        Me.SuspendLayout()
        '
        'Group2
        '
        Me.Group2.Items.Add(Me.AboutBtn)
        Me.Group2.Items.Add(Me.HelpBtn)
        Me.Group2.Items.Add(Me.Button1)
        Me.Group2.Label = "Info"
        Me.Group2.Name = "Group2"
        '
        'TextBoundariesChk
        '
        Me.TextBoundariesChk.Label = "Text Boundaries"
        Me.TextBoundariesChk.Name = "TextBoundariesChk"
        '
        'Group4
        '
        Me.Group4.Items.Add(Me.PasteFromPDF_Btn)
        Me.Group4.Items.Add(Me.MathsToTextBtn)
        Me.Group4.Items.Add(Me.MathsToBrailleBtn)
        Me.Group4.Label = "Text Tools"
        Me.Group4.Name = "Group4"
        '
        'Group3
        '
        Me.Group3.Items.Add(Me.GraphMakerBtn)
        Me.Group3.Items.Add(Me.ThickenLinesBtn)
        Me.Group3.Label = "Drawing Tools"
        Me.Group3.Name = "Group3"
        '
        'AdTab1
        '
        Me.AdTab1.Groups.Add(Me.Group3)
        Me.AdTab1.Groups.Add(Me.Group4)
        Me.AdTab1.Groups.Add(Me.Group1)
        Me.AdTab1.Groups.Add(Me.Group2)
        Me.AdTab1.Label = "Adaptations"
        Me.AdTab1.Name = "AdTab1"
        '
        'Group1
        '
        Me.Group1.Items.Add(Me.TextBoundariesChk)
        Me.Group1.Label = "Display"
        Me.Group1.Name = "Group1"
        '
        'GraphMakerBtn
        '
        Me.GraphMakerBtn.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.GraphMakerBtn.ImageName = "MenuPublish"
        Me.GraphMakerBtn.Label = "Graph Maker"
        Me.GraphMakerBtn.Name = "GraphMakerBtn"
        Me.GraphMakerBtn.OfficeImageId = "GridSettings"
        Me.GraphMakerBtn.ShowImage = True
        '
        'ThickenLinesBtn
        '
        Me.ThickenLinesBtn.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.ThickenLinesBtn.ImageName = "MenuPublish"
        Me.ThickenLinesBtn.Label = "Thicken Lines"
        Me.ThickenLinesBtn.Name = "ThickenLinesBtn"
        Me.ThickenLinesBtn.OfficeImageId = "LineThickness"
        Me.ThickenLinesBtn.ShowImage = True
        '
        'MathsToTextBtn
        '
        Me.MathsToTextBtn.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.MathsToTextBtn.ImageName = "MenuPublish"
        Me.MathsToTextBtn.KeyTip = "MT"
        Me.MathsToTextBtn.Label = "Large Print Maths"
        Me.MathsToTextBtn.Name = "MathsToTextBtn"
        Me.MathsToTextBtn.OfficeImageId = "SelectionToMathConvert"
        Me.MathsToTextBtn.ShowImage = True
        '
        'MathsToBrailleBtn
        '
        Me.MathsToBrailleBtn.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.MathsToBrailleBtn.KeyTip = "MB"
        Me.MathsToBrailleBtn.Label = "Braille Maths Fixes"
        Me.MathsToBrailleBtn.Name = "MathsToBrailleBtn"
        Me.MathsToBrailleBtn.OfficeImageId = "EquationEdit"
        Me.MathsToBrailleBtn.ShowImage = True
        '
        'AboutBtn
        '
        Me.AboutBtn.Label = "About"
        Me.AboutBtn.Name = "AboutBtn"
        '
        'HelpBtn
        '
        Me.HelpBtn.Label = "Help"
        Me.HelpBtn.Name = "HelpBtn"
        '
        'Button1
        '
        Me.Button1.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.Button1.Label = "TEST"
        Me.Button1.Name = "Button1"
        Me.Button1.ShowImage = True
        '
        'ImageList1
        '
        Me.ImageList1.ColorDepth = System.Windows.Forms.ColorDepth.Depth8Bit
        Me.ImageList1.ImageSize = New System.Drawing.Size(16, 16)
        Me.ImageList1.TransparentColor = System.Drawing.Color.Transparent
        '
        'PasteFromPDF_Btn
        '
        Me.PasteFromPDF_Btn.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.PasteFromPDF_Btn.KeyTip = "MB"
        Me.PasteFromPDF_Btn.Label = "Paste from PDF"
        Me.PasteFromPDF_Btn.Name = "PasteFromPDF_Btn"
        Me.PasteFromPDF_Btn.ShowImage = True
        '
        'Ribbon1
        '
        Me.Name = "Ribbon1"
        Me.RibbonType = "Microsoft.Word.Document"
        Me.Tabs.Add(Me.AdTab1)
        Me.Group2.ResumeLayout(False)
        Me.Group2.PerformLayout()
        Me.Group4.ResumeLayout(False)
        Me.Group4.PerformLayout()
        Me.Group3.ResumeLayout(False)
        Me.Group3.PerformLayout()
        Me.AdTab1.ResumeLayout(False)
        Me.AdTab1.PerformLayout()
        Me.Group1.ResumeLayout(False)
        Me.Group1.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents Group2 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents AboutBtn As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents HelpBtn As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Group4 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents Group3 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents GraphMakerBtn As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents AdTab1 As Microsoft.Office.Tools.Ribbon.RibbonTab
    Friend WithEvents ThickenLinesBtn As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents TextBoundariesChk As Microsoft.Office.Tools.Ribbon.RibbonCheckBox
    Friend WithEvents Button1 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Group1 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents MathsToTextBtn As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents MathsToBrailleBtn As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents ImageList1 As Windows.Forms.ImageList
    Friend WithEvents PasteFromPDF_Btn As Microsoft.Office.Tools.Ribbon.RibbonButton
End Class

Partial Class ThisRibbonCollection

    <System.Diagnostics.DebuggerNonUserCode()>
    Friend ReadOnly Property Ribbon1() As Ribbon1
        Get
            Return Me.GetRibbon(Of Ribbon1)()
        End Get
    End Property
End Class
