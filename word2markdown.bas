Attribute VB_Name = "M�dulo1"
Sub pandoc()
    Call Shell("cmd.exe /S /K" & "pandoc -f docx -t markdown foo.docx -o foo.markdown", vbNormalFocus)
    'Shell ("pandoc -f docx -t markdown foo.docx -o foo.markdown")

End Sub
Sub pandoc1()
    Dim wsh As Object
    Set wsh = VBA.CreateObject("WScript.Shell")
    Dim waitOnReturn As Boolean: waitOnReturn = True
    Dim windowStyle As Integer: windowStyle = 1
    wsh.Run "cmd.exe /S /C pandoc -f docx -t markdown foo.docx -o foo.markdown", windowStyle, waitOnReturn
    ' /K deixa o cmd aberto e n�o para a execu��o da rotina
    ' /C fecha o cmd e permite terminar a rotina quando o comando do pandoc
    ' foi concluido.
End Sub


