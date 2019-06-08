Module PublicDeclartions
    'Project-wide declarations

    Public Const DBG As Boolean = True

    Public AB As AboutBox
    Public GMUI As GraphMaker
    Public App As Word.Application
    Public ActiveDocument As Word.Document

End Module

Public Class ThisAddIn

    Private Sub ThisAddIn_Startup() Handles Me.Startup

        'Set Public variable App to access this instance of Application 
        App = Application
        ActiveDocument = App.ActiveDocument

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
