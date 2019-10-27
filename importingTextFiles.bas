Attribute VB_Name = "Módulo1"
Option Explicit
Sub importar2()
'
' importar Macro
'
' Atalho do teclado: Ctrl+y
'

Dim sFileName  As String

'sFileName = Application.GetOpenFilename("MS Excel (*.xlsx), *.xls")
sFileName = Application.GetOpenFilename("DATA(*.dat), *.dat")

'MsgBox (sFileName)
sFN = "TEXT;" & sFileName

sheetname = Dir(sFileName)
'MsgBox (sFN)


Sheets.Add After:=ActiveSheet
ActiveSheet.Name = sheetname

formato sFN
'
End Sub

Sub importar2_dev()
'
' importar Macro
' Atalho do teclado: Ctrl+u
'

'Dim sFileName  As String
Dim sFileName  As Variant
Dim i As Integer
Dim sheetname, sFN As String

'sFileName = Application.GetOpenFilename("MS Excel (*.xlsx), *.xls")
' Get the path to multiples files
sFileName = Application.GetOpenFilename("DATA(*.dat), *.dat", MultiSelect:=True)

' Are multiple files or is only one
If IsArray(sFileName) Then  '<~~ If user selects multiple file
        ' Loop over the multiple files
        For i = LBound(sFileName) To UBound(sFileName)
            'MsgBox sFileName(i)
            sheetname = Dir(sFileName(i))
            Sheets.Add After:=ActiveSheet
            ActiveSheet.Name = sheetname
            sFN = "TEXT;" & sFileName(i)
            formato sFN
        Next i
    Else '<~~ If user selects single file
        'MsgBox sFileName
        sheetname = Dir(sFileName)
        Sheets.Add After:=ActiveSheet
        ActiveSheet.Name = sheetname
        sFN = "TEXT;" & sFileName
        formato sFN
    End If

'formato sFN
'
End Sub

Sub formato(ByVal sFN As String)
With ActiveSheet.QueryTables.Add(Connection:= _
        sFN _
        , Destination:=Range("$A$1"))
   '     .CommandType = 0
   '     .Name = "Pexppre"
        .FieldNames = True
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .TextFilePromptOnRefresh = False
        .TextFilePlatform = 850
        .TextFileStartRow = 1
        .TextFileParseType = xlDelimited
        .TextFileTextQualifier = xlTextQualifierDoubleQuote
        .TextFileConsecutiveDelimiter = True
        .TextFileTabDelimiter = True
        .TextFileSemicolonDelimiter = False
        .TextFileCommaDelimiter = False
        .TextFileSpaceDelimiter = True
        .TextFileColumnDataTypes = Array(1, 1, 1)
        .TextFileDecimalSeparator = "."
        .TextFileThousandsSeparator = ","
        .TextFileTrailingMinusNumbers = True
        .Refresh BackgroundQuery:=False
    End With
End Sub

Sub importar()
Attribute importar.VB_ProcData.VB_Invoke_Func = "i\n14"
'
' importar Macro
'
' Atalho do teclado: Ctrl+i
'

Dim sFileName  As String

sFileName = Application.GetOpenFilename("MS Excel (*.xlsx), *.xls")
Application.CommandBars.ExecuteMso (sFileName)
sFN = "TEXT;" & sFileName
    With ActiveSheet.QueryTables.Add(Connection:= _
        "TEXT;C:\Users\beth\Documents\fluent-models\Lista2\Report\Experim\Pexppre.dat" _
        , Destination:=Range("$A$1"))
        .CommandType = 0
        .Name = "Pexppre"
        .FieldNames = True
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .TextFilePromptOnRefresh = False
        .TextFilePlatform = 850
        .TextFileStartRow = 1
        .TextFileParseType = xlDelimited
        .TextFileTextQualifier = xlTextQualifierDoubleQuote
        .TextFileConsecutiveDelimiter = True
        .TextFileTabDelimiter = True
        .TextFileSemicolonDelimiter = False
        .TextFileCommaDelimiter = False
        .TextFileSpaceDelimiter = True
        .TextFileColumnDataTypes = Array(1, 1, 1)
        .TextFileDecimalSeparator = "."
        .TextFileThousandsSeparator = ","
        .TextFileTrailingMinusNumbers = True
        .Refresh BackgroundQuery:=False
    End With
End Sub
Sub novaAbaRename()
Attribute novaAbaRename.VB_ProcData.VB_Invoke_Func = " \n14"
'
' novaAbaRename Macro
'

'
    'new_sheet
    Sheets.Add After:=ActiveSheet
    'Rename
    Sheets("Planilha5").Select
    Sheets("Planilha5").Name = "novaAva1"
    'Save WorkBook
    ActiveWorkbook.Save
End Sub
Sub Macro6()


Dim sFileName  As String

sFileName = Application.GetOpenFilename("DATA(*.dat), *.dat")

MsgBox (sFileName)
sFN = "TEXT;" & sFileName
MsgBox (sFN)
End Sub

Sub Macro7()
sFileName = Application.GetOpenFilename("DATA(*.dat), *.dat", MultiSelect:=True)

If IsArray(sFileName) Then  '<~~ If user selects multiple file
        For i = LBound(sFileName) To UBound(sFileName)
            MsgBox sFileName(i)
        Next i
    Else '<~~ If user selects single file
        MsgBox sFileName
    End If
End Sub


