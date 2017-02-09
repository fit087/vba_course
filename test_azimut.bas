Attribute VB_Name = "Module1"
Option Explicit

Private Sub test_azimut()
Dim Aproamento As clsAzimute
Set Aproamento = New clsAzimute
Dim graus As Double
Dim saida As VbMsgBoxResult
graus = 727
'saida = MsgBox("graus = " & graus, vbOKOnly, "MsgBox Title")
Aproamento.set_valor = graus

'Debug.Print String(65535, vbCr)
'Print in the immediate window
Debug.Print "Aproamento = "; Aproamento.show_valor

saida = MsgBox("Aproamento = " & Aproamento.show_valor, vbOKOnly)

'Destruindo
Set Aproamento = Nothing

End Sub
