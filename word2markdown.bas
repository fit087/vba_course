Attribute VB_Name = "Módulo1"
Sub pandoc()
    Call Shell("cmd.exe /S /K" & "pandoc -f docx -t markdown foo.docx -o foo.markdown", vbNormalFocus)
    'Shell ("pandoc -f docx -t markdown foo.docx -o foo.markdown")

End Sub
Sub pandoc_v2()
    ' This functioin call the programm pandoc which is installed on Enviroment Variables of the system.
    ' The pandoc become the word document in a markdown file. And this file can be readed by any version control program.
    
    Dim wsh As Object
    Set wsh = VBA.CreateObject("WScript.Shell")
    Dim waitOnReturn As Boolean: waitOnReturn = True
    Dim windowStyle As Integer: windowStyle = 1
'    wsh.Run "cmd.exe /S /C pandoc -f docx -t markdown foo.docx -o foo.markdown", windowStyle, waitOnReturn
    
    ' /K deixa o cmd aberto e não para a execução da rotina
    ' /C fecha o cmd e permite terminar a rotina quando o comando do pandoc
    ' foi concluido.
    wsh.Run "cmd.exe /S /C pandoc -f docx -t markdown " & get_file_name() & " -o " & get_file_name() & ".markdown", windowStyle, waitOnReturn
    
    
End Sub

Sub test()
'    Dim strPath As String
'    'strPath = ThisDocument.FullName
'    strPath = ActiveDocument.FullName
'    MsgBox (strPath)
    
    MsgBox (get_file_name)
    
End Sub

Private Function get_file_name() As String
    get_file_name = ActiveDocument.Name
End Function

