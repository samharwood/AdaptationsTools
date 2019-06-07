Public Class ThisAddIn

    Private Sub ThisAddIn_Startup() Handles Me.Startup

        'Set Public variable App to access this instance of Application 
        App = Application

        'for testing
        'Doc = Application.ActiveDocument
        'Doc.Range(0, 0).Text = "sam"
        'Doc = New Word.Document
        'Doc.Activate()

        If GMUI Is Nothing Then GMUI = New GraphMaker
        GMUI.Show()
        GMUI.Activate()


    End Sub

    Private Sub ThisAddIn_Shutdown() Handles Me.Shutdown

    End Sub

End Class
