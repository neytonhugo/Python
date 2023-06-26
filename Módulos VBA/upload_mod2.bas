Attribute VB_Name = "Módulo2"
Sub Limpar_Inventory()
Attribute Limpar_Inventory.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Limpar_Inventory Macro

    Dim resposta As VbMsgBoxResult

    resposta = MsgBox("Excluir TODOS os dados?", vbYesNo)
     
    If resposta = vbNo Then
        Exit Sub
    Else
        Application.ScreenUpdating = False
    ActiveSheet.Range("$A$1:$L$50000").AutoFilter Field:=1
    ActiveSheet.Range("$A$1:$L$50000").AutoFilter Field:=2
    ActiveSheet.Range("$A$1:$L$50000").AutoFilter Field:=3
    ActiveSheet.Range("$A$1:$L$50000").AutoFilter Field:=4
    ActiveSheet.Range("$A$1:$L$50000").AutoFilter Field:=5
    ActiveSheet.Range("$A$1:$L$50000").AutoFilter Field:=6
    Range("A2:F2000").Select
    Selection.ClearContents
    Range("A2").Select
    End If
End Sub

Sub Limpar_Part_List()
Attribute Limpar_Part_List.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Limpar_Part_List Macro
    Dim resposta As VbMsgBoxResult

    resposta = MsgBox("Excluir TODOS os dados?", vbYesNo)
     
    If resposta = vbNo Then
        Exit Sub
    Else
        Application.ScreenUpdating = False
 
    ActiveSheet.Range("$A$1:$L$50000").AutoFilter Field:=1
    ActiveSheet.Range("$A$1:$L$50000").AutoFilter Field:=2
    ActiveSheet.Range("$A$1:$L$50000").AutoFilter Field:=3
    ActiveSheet.Range("$A$1:$L$50000").AutoFilter Field:=4
    ActiveSheet.Range("$A$1:$L$50000").AutoFilter Field:=5
    ActiveSheet.Range("$A$1:$L$50000").AutoFilter Field:=6
    ActiveSheet.Range("$A$1:$L$50000").AutoFilter Field:=7
    ActiveSheet.Range("$A$1:$L$50000").AutoFilter Field:=8
    ActiveSheet.Range("$A$1:$L$50000").AutoFilter Field:=9
    ActiveSheet.Range("$A$1:$L$50000").AutoFilter Field:=10
    ActiveSheet.Range("$A$1:$L$50000").AutoFilter Field:=11
    Range("A2:K20000").Select
    'Range(Selection, Selection.End(xlDown).End(xlDown)).Select
    Selection.ClearContents
    Range("A2").Select
    End If
End Sub
Sub Limpar_Lay_Out_Consumo()
Attribute Limpar_Lay_Out_Consumo.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Limpar_Lay_Out_Consumo Macro
'
Application.ScreenUpdating = False

    Dim resposta As VbMsgBoxResult

    resposta = MsgBox("Excluir TODOS os dados?", vbYesNo)
     
    If resposta = vbNo Then
        Exit Sub
    Else
        Application.ScreenUpdating = False
        
    ActiveSheet.Outline.ShowLevels RowLevels:=0, ColumnLevels:=1
    Range("A2:BG2").Select
    Range(Selection, Selection.End(xlDown).End(xlDown)).Select
    Selection.SpecialCells(xlCellTypeVisible).Select
    Selection.ClearContents
    Range("A2").Select
    End If
    
    
End Sub
Sub Limpar_layout_horas()

Application.ScreenUpdating = False

    Dim resposta As VbMsgBoxResult

    resposta = MsgBox("Excluir TODOS os dados?", vbYesNo)
     
    If resposta = vbNo Then
        Exit Sub
    Else
        Application.ScreenUpdating = False

    ActiveSheet.Range("$A$1:$L$50000").AutoFilter Field:=15
    ActiveSheet.Range("$A$1:$L$50000").AutoFilter Field:=24
    ActiveSheet.Range("$A$1:$L$50000").AutoFilter Field:=38

ActiveSheet.Outline.ShowLevels RowLevels:=0, ColumnLevels:=1
Range("O2:AL10000").Select
Selection.SpecialCells(xlCellTypeVisible).Select
Selection.ClearContents
Range("O2").Select
End If
End Sub

Sub WO_Anterior()
'
' Ultima_wo Macro mostra a WO anterior
'
    Sheets("LAY_OUT_CONSUMO").Select
    Range("N1").Select
    Selection.End(xlDown).Select
    Selection.Copy
    Sheets("FORM").Select
    Range("T1").PasteSpecial (xlPasteValues)
    Range("D1").PasteSpecial (xlPasteValues)
    Application.CutCopyMode = False
End Sub
Sub Exportar()
'
' Exportar Macro
'
    Sheets("LAY_OUT_CONSUMO").Select
    ActiveSheet.Outline.ShowLevels RowLevels:=0, ColumnLevels:=2
    Range("A2:BG2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Application.CutCopyMode = False
    Selection.Copy
End Sub
Sub Deletar_linha_atual()
Attribute Deletar_linha_atual.VB_ProcData.VB_Invoke_Func = "d\n14"

' Atalho do teclado: Ctrl+d
    linhadeletada = ActiveCell.Row
    If linhadeletada <= 9 Then
    
    aviso = MsgBox("Você não pode excluir a linha selecionada")
    
    Else
    
    Range("A" & ActiveCell.Row & ":AC" & ActiveCell.Row).Select
    Selection.Delete shift:=xlUp
    Range("F" & ActiveCell.Row).Select
    End If
End Sub
