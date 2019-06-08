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
        Me.Group2 = Me.Factory.CreateRibbonGroup
        Me.AboutBtn = Me.Factory.CreateRibbonButton
        Me.HelpBtn = Me.Factory.CreateRibbonButton
        Me.Group4 = Me.Factory.CreateRibbonGroup
        Me.Group3 = Me.Factory.CreateRibbonGroup
        Me.GraphMakerBtn = Me.Factory.CreateRibbonButton
        Me.AdTab1 = Me.Factory.CreateRibbonTab
        Me.ThickenLinesBtn = Me.Factory.CreateRibbonButton
        Me.Group2.SuspendLayout()
        Me.Group3.SuspendLayout()
        Me.AdTab1.SuspendLayout()
        Me.SuspendLayout()
        '
        'Group2
        '
        Me.Group2.Items.Add(Me.AboutBtn)
        Me.Group2.Items.Add(Me.HelpBtn)
        Me.Group2.Label = "Info"
        Me.Group2.Name = "Group2"
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
        'Group4
        '
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
        'GraphMakerBtn
        '
        Me.GraphMakerBtn.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.GraphMakerBtn.Label = "Graph Maker"
        Me.GraphMakerBtn.Name = "GraphMakerBtn"
        Me.GraphMakerBtn.ShowImage = True
        '
        'AdTab1
        '
        Me.AdTab1.Groups.Add(Me.Group3)
        Me.AdTab1.Groups.Add(Me.Group4)
        Me.AdTab1.Groups.Add(Me.Group2)
        Me.AdTab1.Label = "Adaptations"
        Me.AdTab1.Name = "AdTab1"
        '
        'ThickenLinesBtn
        '
        Me.ThickenLinesBtn.Label = "Thicken Lines"
        Me.ThickenLinesBtn.Name = "ThickenLinesBtn"
        Me.ThickenLinesBtn.ShowImage = True
        '
        'Ribbon1
        '
        Me.Name = "Ribbon1"
        Me.RibbonType = "Microsoft.Word.Document"
        Me.Tabs.Add(Me.AdTab1)
        Me.Group2.ResumeLayout(False)
        Me.Group2.PerformLayout()
        Me.Group3.ResumeLayout(False)
        Me.Group3.PerformLayout()
        Me.AdTab1.ResumeLayout(False)
        Me.AdTab1.PerformLayout()
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
End Class

Partial Class ThisRibbonCollection

    <System.Diagnostics.DebuggerNonUserCode()>
    Friend ReadOnly Property Ribbon1() As Ribbon1
        Get
            Return Me.GetRibbon(Of Ribbon1)()
        End Get
    End Property
End Class
