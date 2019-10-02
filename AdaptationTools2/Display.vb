Module Display

    Sub ToggleShowTextBoundaries()
        On Error Resume Next
        If App.ActiveDocument.ActiveWindow.View.ShowTextBoundaries = True Then
            App.ActiveDocument.ActiveWindow.View.ShowTextBoundaries = False
        Else
            App.ActiveDocument.ActiveWindow.View.ShowTextBoundaries = True
        End If
    End Sub

    Sub ToggleShowTabs()
        On Error Resume Next
        Dim p As Word.WdBuiltInProperty
        p = App.ActiveDocument.ActiveWindow.View.ShowTabs
        If p Then
            p = False

        Else
            p = True
        End If
    End Sub

End Module
