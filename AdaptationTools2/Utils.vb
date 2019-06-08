Module Utils

    'Function RoundTo(x, multiple)
    '    RoundTo = Round(x / multiple) * multiple
    'End Function

    Function RoundUp(val As Object) As Integer
        ' To always round upwards towards the next highest number,
        ' take advantage of the way Int() rounds negative numbers downwards
        RoundUp = -Int(-val)
    End Function


    Public Function LineDashStyleID(name As String) As Integer
        Dim d As Integer
        Select Case name
            Case "Solid" : d = 1
            Case "Square Dot" : d = 2
            Case "Round Dot" : d = 3
            Case "Dash" : d = 4
            Case "Dash Dot" : d = 5
            Case "Dash Dot Dot" : d = 6
            Case "Long Dash" : d = 7
            Case "Long Dash Dot" : d = 8
            Case "Long Dash Dot Dot" : d = 9
            Case "Sys Dash" : d = 10
            Case "Sys Dot" : d = 11
            Case "Sys Dash Dot" : d = 12
            Case "Mixed" : d = -2
        End Select
        LineDashStyleID = d
    End Function

    Public Function LineDashStyleName(id As Integer) As String
        Dim s As String
        Select Case id
            Case 1 : s = "Solid"
            Case 2 : s = "Square Dot"
            Case 3 : s = "Round Dot"
            Case 4 : s = "Dash"
            Case 5 : s = "Dash Dot"
            Case 6 : s = "Dash Dot Dot"
            Case 7 : s = "Long Dash"
            Case 8 : s = "Long Dash Dot"
            Case 9 : s = "Long Dash Dot Dot"
            Case 10 : s = "Sys Dash"
            Case 11 : s = "Sys Dot"
            Case 12 : s = "Sys Dash Dot"
            Case -2 : s = "Mixed"
        End Select
        LineDashStyleName = s
    End Function

End Module