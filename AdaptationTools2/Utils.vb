Module Utils



    Function RoundUp(val As Object) As Integer
        ' To always round upwards towards the next highest number,
        ' take advantage of the way Int() rounds negative numbers downwards
        RoundUp = -Int(-val)
    End Function

End Module