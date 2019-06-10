Attribute VB_Name = "Utils"
Option Private Module

' ******** COLOUR DIALOG ********
'This code was originally written by Terry Kreft,
'and modified by Stephen Lebans
'It is not to be altered or distributed,
'except as part of an application.
'You are free to use it in any application,
'provided the copyright notice is left unchanged.
'
' Contact Stephen@lebans.com
'
Private Type COLORSTRUC
  lStructSize As Long
  hwnd As Long
  hInstance As Long
  rgbResult As Long
  lpCustColors As String
  Flags As Long
  lCustData As Long
  lpfnHook As Long
  lpTemplateName As String
End Type

Private Const CC_SOLIDCOLOR = &H80
Private Const CC_RGBINIT = &H1
Private Const CC_FULLOPEN = &H2

Private Declare Function ChooseColor _
    Lib "comdlg32.dll" Alias "ChooseColorA" _
    (pChoosecolor As COLORSTRUC) As Long

Public Function aDialogColor(original As Long) As Long

  Dim x As Long, CS As COLORSTRUC, CustColor(16) As Long
  
  CS.lStructSize = Len(CS)
  CS.hwnd = Empty
  CS.Flags = CC_SOLIDCOLOR + CC_RGBINIT + CC_FULLOPEN
  CS.lpCustColors = String$(16 * 4, 0)
  CS.rgbResult = original
  
  GraphMakerUI.Enabled = False
  x = ChooseColor(CS)
  GraphMakerUI.Enabled = True
  
  If x = 0 Then
    ' ERROR - use Default
    aDialogColor = original
    Exit Function
  Else
    ' Normal processing
     aDialogColor = CS.rgbResult
  End If
  
End Function

' ******** COLOUR DIALOG END ********


Function RoundUp(val As Variant) As Integer
    ' To always round upwards towards the next highest number,
    ' take advantage of the way Int() rounds negative numbers downwards
    RoundUp = -Int(-val)
End Function



