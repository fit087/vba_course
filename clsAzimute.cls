VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsAzimute"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' aAzimute contain the value of the Azimut in a Double type
Private aAzimute As Double

'Get/Let Methods

' Show the value saved in aAzimute
Public Property Get show_valor() As Double
show_valor = aAzimute
End Property

' Set a value in aAzimute
Public Property Let set_valor(ByVal value As Double)

'Dim saida As VbMsgBoxResult
'saida = MsgBox("value = " & value, vbOKOnly, "MsgBox Title")

If value >= 0 Then

    aAzimute = value - (value \ 360) * 360
'saida = MsgBox("conta = " & value - (value \ 360) * 360, vbOKOnly, "MsgBox Title")

ElseIf value < 0 Then

    aAzimute = value + ((-value) \ 360 + 1) * 360

End If
'saida = MsgBox("conta = " & aAzimute, vbOKOnly, "MsgBox Title")
End Property

'General Methods
'Function EmployeeFullInfo() As String
'EmployeeFullInfo = FirstName & " " & LastName & " is a " & Title
'End Function

