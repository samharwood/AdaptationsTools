Module Display

    Sub ToggleTextBoundaries()

        On Error Resume Next

        If App.ActiveDocument.ActiveWindow.View.ShowTextBoundaries = True Then
            App.ActiveDocument.ActiveWindow.View.ShowTextBoundaries = False
        Else
            App.ActiveDocument.ActiveWindow.View.ShowTextBoundaries = True
        End If

    End Sub

End Module
