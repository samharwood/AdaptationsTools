Imports Microsoft.Office.Tools.Ribbon

Public Class Ribbon1


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
End Class
