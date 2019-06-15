Imports Microsoft.Office.Tools.Ribbon

Public Class Ribbon1

    Private Sub Ribbon1_Load(sender As Object, e As RibbonUIEventArgs) Handles Me.Load
        SetTips()


    End Sub

    Sub SetTips()
        'TODO 
        ' Tips for all controls

        MathsToTextBtn.ScreenTip =
"Prepare Maths for Large Print" & Strings.StrDup(30, " ")

        MathsToTextBtn.SuperTip =
"Makes selected Equations/Maths match Normal font style And size

- Converts Equations to 'Text Equations'
- Matches Normal font style And size.
- Further enlarge sub/supscripts for visibility"


        MathsToBrailleBtn.ScreenTip =
"Prepare Maths for MathType+Duxbury Brailling" & Strings.StrDup(30, " ")

        MathsToBrailleBtn.SuperTip =
"Fix-up problems in selected Equations/Maths that can cause issues when Brailling in Duxbury using MathType
Run this before converting Word Equations to MathType Equations

- Sets Equations to Inline type
- Converts 'Text Equations' back to 'Normal Equations'
- Removes spaces in Equations
- Converts all dash/hyphen/etc characters into minus signs"



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

    Private Sub TextBoundariesChk_Click(sender As Object, e As RibbonControlEventArgs) Handles TextBoundariesChk.Click
        ToggleTextBoundaries()
    End Sub

    Private Sub Button1_Click(sender As Object, e As RibbonControlEventArgs) Handles Button1.Click
        ReplaceWrongDash(SelectionToRange)
    End Sub

    Private Sub MathsToBrailleBtn_Click(sender As Object, e As RibbonControlEventArgs) Handles MathsToBrailleBtn.Click
        PrepBrailleMaths()
    End Sub

    Private Sub MathsToTextBtn_Click(sender As Object, e As RibbonControlEventArgs) Handles MathsToTextBtn.Click
        PrepLargePrintMaths()
    End Sub

End Class
