Attribute VB_Name = "Module1"
Option Explicit
' ---------------------------
' Testing the clsAzimute
' ---------------------------
Private Sub test_azimut()
    Dim Aproamento As clsAzimute
    Set Aproamento = New clsAzimute
    Dim graus As Double
    Dim saida As VbMsgBoxResult
    graus = -727
    'saida = MsgBox("graus = " & graus, vbOKOnly, "MsgBox Title")
    Aproamento.set_valor = graus

    'Debug.Print String(65535, vbCr)
    'Print in the immediate window
    Debug.Print "Aproamento = "; Aproamento.show_valor

    saida = MsgBox("Aproamento = " & Aproamento.show_valor, vbOKOnly)

    'Destruindo
    Set Aproamento = Nothing

End Sub


' ---------------------------
' Testing the clsDegrees
' ---------------------------
Private Sub test_degrees()

    Dim angulo1 As clsDegrees
    Set angulo1 = New clsDegrees
    
    Dim testando As Double
    'testando = 36.5656
    testando = 36.56
    
    'angulo1.angle_dec (36.56)
    angulo1.angle_dec = testando 'Muita atencao nao eh usado como uma funcao
    'Debug.Print "Minutes = "; angulo1.min
    Debug.Print "Graus = "; angulo1.degrees()
    Debug.Print "Minutes = "; angulo1.minutes()
    Debug.Print "Segundos = "; angulo1.seconds()
    MsgBox "Segundos = " & angulo1.seconds
    
    'Destruindo
    Set angulo1 = Nothing
    
End Sub



