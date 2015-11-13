Attribute VB_Name = "Módulo1"
Sub ctrl_right_abs()
Attribute ctrl_right_abs.VB_Description = "ctrl_right_abs"
Attribute ctrl_right_abs.VB_ProcData.VB_Invoke_Func = " \n14"
'
' ctrl_right_abs Macro
' ctrl_right_abs
'

'
    Selection.End(xlToRight).Select                     'Control+Right
    Selection.End(xlToLeft).Select                      'Control+Left
    Selection.End(xlToRight).Select                     'Control+Right relative
    Selection.End(xlToLeft).Select                      'Control+Left relative
    Range(Selection, Selection.End(xlToRight)).Select   'Control+Shift+Right
    Range("A1").Select                                  'Select A1
    Range(Selection, Selection.End(xlToRight)).Select   '
    Selection.End(xlToLeft).Select
End Sub
Sub movimentos()
    Range("B3").Select
    Call aspetta
    Range("B5").Select
    Call aspetta
    'ActiveCell.Offset(-2, 1).Range("A1").Select 'C3
    'ActiveCell.Offset(-2, 1).Range("A1:B2").Select 'C3:D4
    ActiveCell.Item(-1, 2).Select   'C3
    
    'ActiveCell.Item(1, -2).Range("A1").Select
    'Range.Cells(2, 3)
    'Range("C4").Item(2, 2).Select
    
    Call aspetta
    Range("B6").Offset(-2, 1).Select
    

End Sub

Sub select_relative()
Attribute select_relative.VB_ProcData.VB_Invoke_Func = " \n14"
'
' select_relative Macro
'

'
    Range("B5").Select
    ActiveCell.Offset(-2, 1).Range("A1").Select
End Sub

Sub aspetta()
Application.Wait (Now + TimeValue("0:00:03"))
End Sub
Sub somatorio()
Attribute somatorio.VB_ProcData.VB_Invoke_Func = " \n14"
'
' somatorio Macro
'

'
    Range("C7").Select
    ActiveCell.FormulaR1C1 = "=MAX(R[-6]C:R[-1]C)"
    Range("C8").Select
    ActiveCell.Offset(-1, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=MAX(R[-6]C:R[-1]C)"
    ActiveCell.Offset(0, 1).Range("A1").Select
End Sub

Sub suma()
    Range("A1").Offset(0, 4).Select
    'Selection.End(xlToRight).Select
    'ActiveCell.Offset(1, 0).FormulaR1C1 = "=MAX(RC[-4]:RC[-1])"
    'ActiveCell.Offset(0, 1).Select
    'ActiveCell.FormulaR1C1 = "=MAX(RC[-4]:RC[-1])"
    ActiveCell.Formula = "=MAX(A1:D1)"
    'Call aspetta
    Dim vector_linha
    vector_linha = Range("A2:D2").Value
    'Dim soma As Single
    Dim soma
    soma = 0
    
    'For i = 1 To 4
    '   soma = soma + vector_linha(1, i)
    'Next i
    
    'For Ndx = LBound(InputArray) To (UBound(InputArray)
    
    For Each element In vector_linha
        soma = soma + element
    Next element
    
    
    Cells(2, 5).Value = soma
    'Cells(2, 5).Value = vetor_linha(1, 1)
    'Range(2, 5).Value = soma
    
    Dim vector_linha2
    vector_linha2 = Range("A3:D3").Value
    Cells(3, 5).Value = WorksheetFunction.Sum(vector_linha2)
    'WorksheetFunction.
    



End Sub
Sub ctrl_plus()
'
' ctrl_plus Macro
'

'
    ActiveCell.Offset(0, 6).Columns("A:A").EntireColumn.Select
    Selection.Insert Shift:=xlToRight
    ActiveCell.Range("A1:A6").Select
    Selection.Insert Shift:=xlToRight
    Selection.Delete Shift:=xlToLeft
    ActiveCell.Offset(0, 1).Columns("A:A").EntireColumn.Select
    Selection.Delete Shift:=xlToLeft
End Sub

Sub soma()
    'Cells(1, 10)
    Range("J1").Select
    'Selection.End(xlToDown).Select
    Range(Selection, Selection.End(xlDown)).Select   'Control+Shift+Down
    
'    Range("J1", Selection.End(xlDown)).Select   'Control+Shift+Down
    
    'ActiveCell.Offset(0, 1).Select
    'Selection.Offset(0, 1).Select
     
    'Selection.Offset(0, 1).Insert shift:=xlToRight
    
    Selection.Offset(0, 1).Select
    
    Selection.Insert Shift:=xlToRight
    
    
 
    'Range("J1", Selection.End(xlDown)).Offset(0, 1).Insert (xlToRight)
    
    Selection.FormulaR1C1 = "=MAX(RC[-5]:RC[-1])"
    
    Range("D8").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Offset(0, 1).Select
    With Selection                              'Estatico
    '    .Offset(0, 1).Select
        .Insert Shift:=xlToRight
        .FormulaR1C1 = "=MAX(RC[-4]:RC[-1])"
    End With
    
    
End Sub

Sub fin()
'
' fin Macro
'

'
    ActiveCell.SpecialCells(xlLastCell).Select  'ctrl+End
    Range("A1").Select
    Selection.End(xlDown).Select
    Selection.End(xlUp).Select
    ActiveCell.SpecialCells(xlLastCell).Select
    Range("A1").Select
    ActiveWindow.LargeScroll Down:=1
    ActiveCell.Offset(23, 0).Range("A1").Select
    ActiveWindow.LargeScroll Down:=-1
    ActiveCell.Offset(-23, 0).Range("A1").Select
    ActiveWindow.LargeScroll Down:=1
    Range("A24").Select
    ActiveWindow.LargeScroll Down:=-1
    Range("A1").Select
    ActiveWindow.ScrollColumn = 3
    ActiveWindow.ScrollColumn = 4
    Range("K1:K6").Select
    Range("K6").Activate
    Selection.Delete Shift:=xlToLeft
    Range("C8:C13").Select
    Selection.AutoFill Destination:=Range("C8:D13"), Type:=xlFillDefault
    Range("C8:D13").Select
    Range("E8").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Delete Shift:=xlToLeft
    Range("G5").Select
    ActiveWindow.ScrollColumn = 2
    ActiveWindow.ScrollColumn = 3
    Range("K6").Select
    Range(Selection, Selection.End(xlUp)).Select
    Selection.Delete Shift:=xlToLeft
    Range("E6").Select
    ActiveWindow.ScrollColumn = 2
    ActiveWindow.ScrollColumn = 1
    Range("C8:C13").Select
    Selection.AutoFill Destination:=Range("C8:D13"), Type:=xlFillDefault
    Range("C8:D13").Select
    Range("E8:E13").Select
    Selection.Delete Shift:=xlToLeft
    Range("K1").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Delete Shift:=xlToLeft
    Range("K6").Select
    Range(Selection, Selection.End(xlUp)).Select
    Selection.Delete Shift:=xlToLeft
    Range("H5").Select
    ActiveWindow.ScrollColumn = 3
    ActiveWindow.ScrollColumn = 2
    ActiveWindow.ScrollColumn = 1
    Range("E8").Select
    ActiveWindow.ScrollColumn = 2
    ActiveWindow.ScrollColumn = 3
    ActiveWindow.ScrollColumn = 4
    ActiveWindow.ScrollColumn = 5
    ActiveWindow.ScrollColumn = 6
    ActiveWindow.ScrollColumn = 5
    ActiveWindow.ScrollColumn = 4
    ActiveWindow.ScrollColumn = 5
    Range("F8").Select
    ActiveWindow.ScrollColumn = 4
    ActiveWindow.ScrollColumn = 3
    ActiveWindow.ScrollColumn = 2
    ActiveWindow.ScrollColumn = 1
End Sub

Sub ctrl_shift_end()
'
' ctrl_shift_end Macro
'

'
    Range(Selection, ActiveCell.SpecialCells(xlLastCell)).Select
    Range("A1").Select
    Range(Selection, ActiveCell.SpecialCells(xlLastCell)).Select
    Range("A1").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToLeft)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToLeft)).Select
    Range(Selection, Selection.End(xlUp)).Select
    Range(Selection, Selection.End(xlUp)).Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("D8").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.AutoFill Destination:=Range("D8:E13"), Type:=xlFillDefault
    Range("D8:E13").Select
    Range("E8:E13").Select
    Selection.ClearContents
End Sub
