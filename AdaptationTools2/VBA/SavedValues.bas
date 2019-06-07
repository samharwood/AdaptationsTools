Attribute VB_Name = "SavedValues"
Option Private Module

Public Enum GMPropsList
    MajorLineColour = 1
    MinorLineColour = 3
End Enum

Public GMctls() As Variant
Public GMprops() As Variant
Public GMP As New Collection

Sub SetupProps()
    ' Retrieves saved values from CustomDocumentProperties
    ' If GMctls/GMProps declarations are extended, old values are still loaded OK
    
    If Not DBG Then On Error GoTo er
    
    Dim GMprops_LastSave() As Variant
    
    
    'Declare PropertyName, Defaults
    GMprops() = Array( _
        "MajorLineColour", wdColorBlack, _
        "MinorLineColour", wdColorGray60)
    
    'Retrieve saved values
    GMprops_LastSave() = Deserialise(ActiveDocument.CustomDocumentProperties.Item("GraphMakerProps"))
    
    'Use collection to match up PropertyNames in case changed since last save.
    For i = 0 To UBound(GMprops) Step 2
        GMP.Add GMprops(i + 1), GMprops(i)
    Next i
            
    'Overwrite matching collection values with any saved values
    On Error Resume Next 'ignore any saved data that has changed/is old, etc
    For i = 0 To UBound(GMprops_LastSave) Step 2
        GMP(GMprops_LastSave(i)) = GMprops_LastSave(i + 1)
    Next i
    
    'Write collection values back to array
    For i = 1 To GMP.Count
         GMprops(i * 2 - 1) = GMP(i)
    Next i
    
    
    Exit Sub
er:
    Select Case Err.number
        Case 5 '5 =property doesn't exist
            'create it
            StorePropertyString "GraphMakerProps", ""
            Resume
        Case Else
            Err.Raise Err.number
    End Select
End Sub



Sub SetupControls()
    ' Retrieves saved values from CustomDocumentProperties
    ' If GMctls/GMProps declarations are extended, old values are still loaded OK
    
    If Not DBG Then On Error GoTo er
    
    Dim GMctls_LastSave() As Variant
    
    'Declare ControlNames, Defaults
    GMctls() = Array( _
        "xFrom", "0", _
        "yFrom", "0", _
        "xTo", "6", _
        "yTo", "6", _
        "xNumEvery", "1", _
        "yNumEvery", "1", _
        "xDivs", "1", _
        "yDivs", "1", _
        "Axes", "True", _
        "AxisLabels", "True", _
        "Numbering", "True", _
        "Ticks", "True", _
        "majorWeight", "3", _
        "majorDash", "Solid", _
        "minorWeight", "2", _
        "minorDash", "Sys Dash", _
        "PlotAsChart", "True", _
        "PlotAsShapes", "False", _
        "UEBBraille", "False")
        
    'Write defaults to UI
    For i = 0 To UBound(GMctls) Step 2
        GraphMakerUI.Controls(GMctls(i)).Value = GMctls(i + 1)
    Next i
    
    'Overwrite UI defaults with any saved values
    GMctls_LastSave() = Deserialise(ActiveDocument.CustomDocumentProperties.Item("GraphMakerCtls"))
    
    On Error Resume Next 'ignore any saved data that has changed/is old, etc
    For i = 0 To UBound(GMctls_LastSave) Step 2
        GraphMakerUI.Controls(GMctls_LastSave(i)).Value = GMctls_LastSave(i + 1)
    Next i
    
    
    Exit Sub
er:
    Select Case Err.number
        Case 5 '5 =property doesn't exist
            'create it
            StorePropertyString "GraphMakerCtls", ""
            Resume
        Case Else
            Err.Raise Err.number
    End Select
End Sub

Sub clearprops()
    On Error Resume Next
    ActiveDocument.CustomDocumentProperties("GraphMakerCtls").Delete
    ActiveDocument.CustomDocumentProperties("GraphMakerProps").Delete
End Sub

Sub InitSavedValues()

    'If Not DBG Then On Error GoTo er
       
    SetupControls
    SetupProps
    
    ' Setup other UI Properties
    GraphMakerUI.majorColour.BackColor = GMprops(MajorLineColour)
    GraphMakerUI.minorColour.BackColor = GMprops(MinorLineColour)
    GraphMakerUI.Repaint
    
    Exit Sub
'er:
'    Select Case Err.number
'        Case 5, 438, -2147024809 '5 =property doesn't exist,'438 = wrong type
'            'Use Defaults instead
'            SetupDefaults
'            StorePropertyString "GraphMakerCtls", Serialise(GMctls)
'            StorePropertyString "GraphMakerProps", Serialise(GMprops)
'            Resume try_again
'        Case -2147024809 'control not found
'            MsgBox "DEBUG: Control Name not found. Update Controls array!"
'        Case Else
'            Err.Raise Err.number
'    End Select
End Sub


Sub SaveValues()
    ' Update GMctls array with UI values
    ' Serialise and Save
                
    For i = 0 To UBound(GMctls) Step 2
        GMctls(i + 1) = GraphMakerUI.Controls(GMctls(i))
    Next i

    StorePropertyString "GraphMakerCtls", Serialise(GMctls)
    StorePropertyString "GraphMakerProps", Serialise(GMprops)
    
End Sub



Function Serialise(a As Variant) As String
    ' Expects Array of String values
    ' Returns string of Comma seperated values
    Dim s As String
    
    For i = LBound(a) To UBound(a)
        s = s + CStr(a(i)) & ","
    Next i

    Serialise = s

End Function


Function Deserialise(strSerial As String) As Variant
    ' Expects string of Comma seperated values
    ' Returns Array of String values
    Dim a() As Variant
    
    
    j = 0
    i = 1
    prev = 1

    Do
        i = InStr(i + 1, strSerial, ",")
        If i = 0 Then i = Len(strSerial) + 1 'Last element
    
        ReDim Preserve a(j)
        a(j) = Mid(strSerial, prev, i - prev)
        j = j + 1
        prev = i + 1
    Loop Until i >= Len(strSerial) 'Last element calc makes i > Len
    
    Deserialise = a()

End Function




Sub StorePropertyString(name As String, val As String)
    If Not DBG Then On Error GoTo er
    
    ActiveDocument.CustomDocumentProperties(name).Delete
    ActiveDocument.CustomDocumentProperties.Add name, False, msoPropertyTypeString, val

    
    Exit Sub
er:
    Select Case Err.number
        Case 5 'already deleted
            Resume Next
        Case Else
            Err.Raise Err.number
    End Select
End Sub



Public Function LineDashStyleID(ByVal name As Variant) As Integer
    Select Case name
        Case "Solid": d = 1
        Case "Mixed": d = -2
        Case "Square Dot": d = 2
        Case "Round Dot": d = 3
        Case "Dash": d = 4
        Case "Dash Dot": d = 5
        Case "Dash Dot Dot": d = 6
        Case "Long Dash": d = 7
        Case "Long Dash Dot": d = 8
        Case "Long Dash Dot Dot": d = 9
        Case "Sys Dash": d = 10
        Case "Sys Dot": d = 11
        Case "Sys Dash Dot": d = 12
    End Select
    LineDashStyleID = d
End Function

Public Function LineDashStyleName(ByVal id As Variant) As String
    Select Case id
        Case 1: s = "Solid"
        Case -2: s = "Mixed"
        Case 2: s = "Square Dot"
        Case 3: s = "Round Dot"
        Case 4: s = "Dash"
        Case 5: s = "Dash Dot"
        Case 6: s = "Dash Dot Dot"
        Case 7: s = "Long Dash"
        Case 8: s = "Long Dash Dot"
        Case 9: s = "Long Dash Dot Dot"
        Case 10: s = "Sys Dash"
        Case 11: s = "Sys Dot"
        Case 12: s = "Sys Dash Dot"
    End Select
    LineDashStyleName = s
End Function


