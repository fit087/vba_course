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
    
    'For Ndx = LBound(InputArray) To UBound(InputArray)
    
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
