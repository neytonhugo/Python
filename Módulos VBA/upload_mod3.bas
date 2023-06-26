Attribute VB_Name = "Módulo3"
Sub Upload_automatico_consumo()
Attribute Upload_automatico_consumo.VB_ProcData.VB_Invoke_Func = "q\n14"
' Atalho do teclado: Ctrl+Q
Sheets("LAY_OUT_CONSUMO").Select
ultimalinha = Range("A1").End(xlDown).Row

Range("BG" & ultimalinha & ":A2").Copy
    ActiveSheet.Outline.ShowLevels RowLevels:=0, ColumnLevels:=2

End Sub
Sub Upload_automatico_horas()
Attribute Upload_automatico_horas.VB_ProcData.VB_Invoke_Func = "h\n14"
' Atalho do teclado: Ctrl+H

Sheets("LAY_OUT_HORAS").Select
ultimalinha = Range("O1").End(xlDown).Row

Range("AL" & ultimalinha & ":A2").Copy
    ActiveSheet.Outline.ShowLevels RowLevels:=0, ColumnLevels:=2

End Sub
Sub Upload_automatico_cop_lote_horas()
Attribute Upload_automatico_cop_lote_horas.VB_ProcData.VB_Invoke_Func = "H\n14"
' Atalho do teclado: Ctrl+shift+H

Sheets("LAY_OUT_HORAS").Select

ActiveSheet.Outline.ShowLevels RowLevels:=0, ColumnLevels:=2

ultimalinha2 = Range("O1").End(xlDown).Row

Range("B" & ultimalinha2 & ":B2").Copy

End Sub

Sub TransporValores()
Attribute TransporValores.VB_ProcData.VB_Invoke_Func = "V\n14"

    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=True
End Sub


