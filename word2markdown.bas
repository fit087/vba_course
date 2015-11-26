Attribute VB_Name = "Módulo1"
Sub pandoc()
    Call Shell("cmd.exe /S /K" & "pandoc -f docx -t markdown foo.docx -o foo.markdown", vbNormalFocus)
    'Shell ("pandoc -f docx -t markdown foo.docx -o foo.markdown")

End Sub
Sub pandoc1()
    Dim wsh As Object
    Set wsh = VBA.CreateObject("WScript.Shell")
    Dim waitOnReturn As Boolean: waitOnReturn = True
    Dim windowStyle As Integer: windowStyle = 1
    wsh.Run "cmd.exe /S /K pandoc -f docx -t markdown foo.docx -o foo.markdown", windowStyle, waitOnReturn
End Sub


