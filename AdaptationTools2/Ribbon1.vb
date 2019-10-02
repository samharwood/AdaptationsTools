Imports Microsoft.Office.Tools.Ribbon

Public Class Ribbon1

    Private Sub Ribbon1_Load(sender As Object, e As RibbonUIEventArgs) Handles Me.Load
        SetTips()

        ' Set UI to match Application state
        ShowTabsChk.Checked = Globals.ThisAddIn.Application.ActiveWindow.View.ShowTabs
        ShowTbChk.Checked = Globals.ThisAddIn.Application.ActiveWindow.View.ShowTextBoundaries
    End Sub



    Sub SetTips()
        'TODO 
        ' Tips for all controls
        ' 

        MathsToTextBtn.ScreenTip =
"Prepare Maths for Large Print" & Strings.StrDup(30, " ")

        MathsToTextBtn.SuperTip =
"Makes currently selected Equations/Maths match Normal font style and size

- Converts Word Equations to 'Text' to allow font face to be changed.
- Matches Normal font style and size.
- Further enlarges sub/supscripts for visibility."


        MathsToBrailleBtn.ScreenTip =
"Prepare Maths for MathType+Duxbury Brailling" & Strings.StrDup(30, " ")

        MathsToBrailleBtn.SuperTip =
"Fix-up problems in currently selected Word Equations/Maths that can cause issues when Brailling in Duxbury using MathType
Run this before converting Word Equations to MathType Equations

- Sets Word Equations to Inline type.
- Converts 'Text' Word Equations back to 'Normal' Word Equations.
- Removes spaces in Word Equations.
- Converts all dash/hyphen/etc characters into minus signs."


        ThickenLinesBtn.ScreenTip =
"Thicken outlines of selected shapes and table borders" & Strings.StrDup(30, " ")

        ThickenLinesBtn.SuperTip =
"Either to specific line weight or in proportion to size of the Normal font style"


        PasteFromPDF_Btn.ScreenTip =
"Paste as plain text and fix-up common errors in text copied from a PDF file." & Strings.StrDup(30, " ")

        PasteFromPDF_Btn.SuperTip =
"- Remove line breaks from text wrapping at page edge.
- Replace hyphen character with minus character
- Remove 'fake' bullet points"


        GraphMakerBtn.ScreenTip =
"Quickly create accurate grids/graph layouts for Large Print/Raised diagrams" & Strings.StrDup(30, " ")

        GraphMakerBtn.SuperTip =
"- Specifiy basic configuration all in one place.
- Option to generate numbering in UEB Braille.
- Option to auto set line weight based on Normal style font size.
- Option to save different settings per document"
    End Sub

    Private Sub GraphMakerBtn_Click(sender As Object, e As RibbonControlEventArgs) Handles GraphMakerBtn.Click

        If GMUI Is Nothing Then GMUI = New GraphMaker
        GMUI.Show()
        GMUI.Activate()
    End Sub

    Private Sub AboutBtn_Click(sender As Object, e As RibbonControlEventArgs) Handles AboutBtn.Click

        If AB Is Nothing Then AB = New AboutBox
        AB.Show()
        AB.Activate()
    End Sub

    Private Sub ThickenLinesBtn_Click(sender As Object, e As RibbonControlEventArgs) Handles ThickenLinesBtn.Click
        ThickenLines()
    End Sub



    Private Sub Button1_Click(sender As Object, e As RibbonControlEventArgs)
        ReplaceWrongDash(SelectionToRange)
    End Sub

    Private Sub MathsToBrailleBtn_Click(sender As Object, e As RibbonControlEventArgs) Handles MathsToBrailleBtn.Click
        PrepBrailleMaths()
    End Sub

    Private Sub MathsToTextBtn_Click(sender As Object, e As RibbonControlEventArgs) Handles MathsToTextBtn.Click
        PrepLargePrintMaths()
    End Sub

    Private Sub PasteFromPDF_Btn_Click(sender As Object, e As RibbonControlEventArgs) Handles PasteFromPDF_Btn.Click
        PasteOCR()
    End Sub

    ' Display
    Private Sub ShowTabsChk_Click(sender As Object, e As RibbonControlEventArgs) Handles ShowTabsChk.Click
        ToggleProperty(App.ActiveWindow.View.ShowTabs, ShowTabsChk)
    End Sub

    Private Sub ShowTbChk_Click(sender As Object, e As RibbonControlEventArgs) Handles ShowTbChk.Click
        ToggleProperty(App.ActiveWindow.View.ShowTextBoundaries, ShowTbChk)
    End Sub

    Sub ToggleProperty(ByRef p As Word.WdBuiltInProperty, ByRef chk As RibbonCheckBox)
        If p = True Then
            p = False
            chk.Checked = False
        Else
            p = True
            chk.Checked = True
        End If
    End Sub

End Class
