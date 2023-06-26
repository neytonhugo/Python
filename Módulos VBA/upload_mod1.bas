Attribute VB_Name = "Módulo1"
Sub inserir_itens_bom()
Attribute inserir_itens_bom.VB_Description = "Insere os itens da lista técnica na WO."
Attribute inserir_itens_bom.VB_ProcData.VB_Invoke_Func = " \n14"
Sheets("FORM").Select

Application.CutCopyMode = False
Application.ScreenUpdating = False

'Limpando antes de colar

Range("C10:C50").ClearContents
'Range("F10:F50").ClearContents
Range("G10:G50").ClearContents

Dim WO As Double

WO = Range("$C$5").Value

    'Selecionando a aba FORM e copiando # da WO.
    Sheets("FORM").Select
    
    'Selecionando aba WO_PART_LIST, inserindo auto-filtro com o # da WO e copiando os códigos dos itens filhos.
    Sheets("WO_PART_LIST").Select
    ActiveSheet.Range("A1:D100000").AutoFilter Field:=1, Criteria1:=WO
    Filtro = Range("D1").End(xlDown).Row
    Range("D2" & ":D" & Filtro).Copy
     
    'Selecionando aba FORM e colando infomações de itens filhos.
    Sheets("FORM").Select
    ActiveSheet.Range("c9").Select
    Range("C10").PasteSpecial Paste:=xlPasteValues
    'copiando quant stardard
    'Sheets("WO_PART_LIST").Select
    'Range("E2" & ":E" & Filtro).Copy
    'Sheets("FORM").Select
    'Range("F10").PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
    Application.ScreenUpdating = False
    
End Sub

Sub inserir_itens_bom2()
Sheets("FORM").Select

Application.CutCopyMode = False
Application.ScreenUpdating = False

'Pergunta de confirmação antes de exucutar a SUB
Dim resposta As VbMsgBoxResult

    resposta = MsgBox("Deseja Sobrescrever os dados?", vbYesNo)
     
    If resposta = vbNo Then
        Exit Sub
    Else
WO = Range("$C$5").Value

    'Selecionando a aba FORM e copiando # da WO.
    Sheets("FORM").Select
    
    'Selecionando aba WO_PART_LIST, inserindo auto-filtro com o # da WO e copiando os códigos dos itens filhos.
    Sheets("WO_PART_LIST").Select
    ActiveSheet.Range("A1:D100000").AutoFilter Field:=1, Criteria1:=WO
    Filtro = Range("D1").End(xlDown).Row
    Range("D2" & ":D" & Filtro).Copy
     
    'Selecionando aba FORM e colando infomações de itens filhos.
    Sheets("FORM").Select
    ActiveSheet.Range("c9").Select
    Range("C10").PasteSpecial Paste:=xlPasteValues
    'copiando quant stardard
    Sheets("WO_PART_LIST").Select
    Range("E2" & ":E" & Filtro).Copy
    Sheets("FORM").Select
    Range("F10").PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
    Application.ScreenUpdating = False
    End If
End Sub
Sub inserir_lote()
Attribute inserir_lote.VB_ProcData.VB_Invoke_Func = "i\n14"

Dim inserir_lote As Integer
Dim inserir_lote2 As Integer
Dim aviso As VbMsgBoxResult

rangelinha = ActiveCell.Row
If rangelinha >= 10 Then

        Application.ScreenUpdating = False
        Dim resposta As VbMsgBoxResult

    resposta = MsgBox("Igualar Saldo e Qtd Real - (Calculando Sobra)?", vbYesNo)
     
    If resposta = vbNo Then
    
     'Selecionando aba FORM, verificando qual o último item que teve o # lote digitado.
    Sheets("FORM").Select
    LinhaLote = ActiveCell.Row
    LinhaLote2 = ActiveCell.Row + 1
    
    'Inserindo linha para novo # de lote, copiando dados da linha superior e colando na linha nova.
    Range("A" & LinhaLote & ":AA" & LinhaLote).Select
    Selection.Insert shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A" & LinhaLote2 & ":AA" & LinhaLote2).Copy
    ActiveSheet.Paste
    Range("C" & ActiveCell.Row + 1 & ":J" & ActiveCell.Row + 1).Select
       With Selection.Interior
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent5
        .TintAndShade = 0.799981688894314
        .PatternTintAndShade = 0
    End With
    Application.CutCopyMode = False
    Range("F" & ActiveCell.Row - 1).Select
       Exit Sub
    Else
     'Selecionando aba FORM, verificando qual o último item que teve o # lote digitado.
    Sheets("FORM").Select
    LinhaLote = ActiveCell.Row
    LinhaLote2 = ActiveCell.Row + 1
    
    'Inserindo linha para novo # de lote, copiando dados da linha superior e colando na linha nova.
    Range("A" & LinhaLote & ":AA" & LinhaLote).Select
    Selection.Insert shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A" & LinhaLote2 & ":AA" & LinhaLote2).Copy
    ActiveSheet.Paste
    Range("C" & ActiveCell.Row + 1 & ":J" & ActiveCell.Row + 1).Select
       With Selection.Interior
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent5
        .TintAndShade = 0.799981688894314
        .PatternTintAndShade = 0
    End With
    Application.CutCopyMode = False
    Range("F" & ActiveCell.Row - 1).Select
    
    Saldo = ActiveCell.Row + 1
    Saldo1 = ActiveCell.Row
    Saldo2 = Range("J" & Saldo)
    Range("J" & Saldo).Copy
    Range("F" & Saldo1).PasteSpecial (xlPasteValues)
    Range("F" & Saldo1 + 1).Select
    Call calcular_sobra_manual
    Range("G" & ActiveCell.Row).ClearContents
    Range("G" & ActiveCell.Row).Select
    Application.CutCopyMode = False
    End If
Else
aviso = MsgBox("Selecione uma linha válida")
End If
End Sub
Sub calcular_sobra_manual()
Attribute calcular_sobra_manual.VB_ProcData.VB_Invoke_Func = "m\n14"

Range("F" & ActiveCell.Row).Copy

Range("F" & ActiveCell.Row).PasteSpecial (xlPasteValues)

Range("F" & ActiveCell.Row) = Range("F" & ActiveCell.Row) - Range("F" & ActiveCell.Row - 1)

End Sub

Sub salvar_dados()
     
    If Range("Q1") = FALSO Then
  
    Application.ScreenUpdating = False
      
    'Comando Colar Código
    Sheets("FORM").Select
    codigo = Range("c9").End(xlDown).Row
    Range("c10" & ":c" & codigo).Copy
    Sheets("LAY_OUT_CONSUMO").Select
    cola_codigo = Range("AO1").End(xlDown).End(xlDown).End(xlUp).Row + 1
    Range("AO" & cola_codigo).PasteSpecial (xlPasteValues)
    
    'Comando Colar U.M
    Sheets("FORM").Select
    Range("e10" & ":e" & codigo).Copy
    Sheets("LAY_OUT_CONSUMO").Select
    Range("BG" & cola_codigo).PasteSpecial (xlPasteValues)

    'Comando Colar Quantidade Real.
    Sheets("FORM").Select
    Range("f10" & ":f" & codigo).Copy
    Sheets("LAY_OUT_CONSUMO").Select
    Range("BB" & cola_codigo).PasteSpecial (xlPasteValues)
  
    'Comando Colar Lote.
    Sheets("FORM").Select
    Range("G10" & ":G" & codigo).Copy
    Sheets("LAY_OUT_CONSUMO").Select
    Range("AU" & cola_codigo).PasteSpecial (xlPasteValues)
  
    'Comando Colar Local Estoque.
    Sheets("FORM").Select
    Range("H10" & ":H" & codigo).Copy
    Sheets("LAY_OUT_CONSUMO").Select
    Range("AT" & cola_codigo).PasteSpecial (xlPasteValues)
    
    'Comando Colar Sequencia Operação.
    Sheets("FORM").Select
    Range("I10" & ":I" & codigo).Copy
    Sheets("LAY_OUT_CONSUMO").Select
    Range("AE" & cola_codigo).PasteSpecial (xlPasteValues)
   
    'Copiando e colando # da WO na aba LAY_OUT_CONSUMO.
    Sheets("FORM").Select
    Range("c5").Copy
    Sheets("LAY_OUT_CONSUMO").Select
    'cola_wo = Range("AO1").End(xlDown).End(xlDown).End(xlUp).Row
    cola_wo2 = Range("N1").End(xlDown).End(xlDown).End(xlUp).Row + 1
    cola_wo3 = Range("AO1").End(xlDown).End(xlDown).End(xlUp).Row
    Range("N" & cola_wo2 & ":N" & cola_wo3).PasteSpecial (xlPasteValues)
  
    'Copiando # do usuário da aba FORM e colando-o na aba LAY_OUT_CONSUMO.
    Sheets("FORM").Select
    Range("j1").Copy
    Sheets("LAY_OUT_CONSUMO").Select
    Range_usuario_ini = Range("a1").End(xlDown).End(xlDown).End(xlUp).Row + 1
    Range_usuario_fim = Range("n1").End(xlDown).End(xlDown).End(xlUp).Row
    Range("A" & Range_usuario_ini & ":A" & Range_usuario_fim).PasteSpecial (xlPasteValues)
    
'Atualização de estoque após salvar template
     
ActiveWorkbook.RefreshAll
  Call limpa_planilha
  Call WO_Anterior
  Range("C5").Select
  Application.CutCopyMode = False

Else
    Application.ScreenUpdating = False
           'Comando Colar Código
    Sheets("FORM").Select
    codigo = Range("c9").End(xlDown).Row
    Range("c10" & ":c" & codigo).Copy
    Sheets("LAY_OUT_CONSUMO").Select
    cola_codigo = Range("AO1").End(xlDown).End(xlDown).End(xlUp).Row + 1
    Range("AO" & cola_codigo).PasteSpecial (xlPasteValues)
    
    'Comando Colar U.M
    Sheets("FORM").Select
    Range("e10" & ":e" & codigo).Copy
    Sheets("LAY_OUT_CONSUMO").Select
    Range("BG" & cola_codigo).PasteSpecial (xlPasteValues)
 
    'Comando Colar Quantidade Real.
    Sheets("FORM").Select
    Range("f10" & ":f" & codigo).Copy
    Sheets("LAY_OUT_CONSUMO").Select
    Range("BB" & cola_codigo).PasteSpecial (xlPasteValues)
    Linha_quant = Range("AO" & cola_codigo).Row
    Linha_quant2 = Range("AO1").End(xlDown).End(xlDown).End(xlUp).Row
    Range("CX" & Linha_quant & ":CX" & Linha_quant2).Copy
    Range("BB" & cola_codigo).PasteSpecial (xlPasteValues)
    
    'Comando Colar Lote.
    Sheets("FORM").Select
    Range("G10" & ":G" & codigo).Copy
    Sheets("LAY_OUT_CONSUMO").Select
    Range("AU" & cola_codigo).PasteSpecial (xlPasteValues)
  
    'Comando Colar Local Estoque.
    Sheets("FORM").Select
    Range("H10" & ":H" & codigo).Copy
    Sheets("LAY_OUT_CONSUMO").Select
    Range("AT" & cola_codigo).PasteSpecial (xlPasteValues)
    
    'Comando Colar Sequencia Operação.
    Sheets("FORM").Select
    Range("I10" & ":I" & codigo).Copy
    Sheets("LAY_OUT_CONSUMO").Select
    Range("AE" & cola_codigo).PasteSpecial (xlPasteValues)
    
    'Copiando e colando # da WO na aba LAY_OUT_CONSUMO.
    Sheets("FORM").Select
    Range("c5").Copy
    Sheets("LAY_OUT_CONSUMO").Select
    'cola_wo = Range("AO1").End(xlDown).End(xlDown).End(xlUp).Row
    cola_wo2 = Range("N1").End(xlDown).End(xlDown).End(xlUp).Row + 1
    cola_wo3 = Range("AO1").End(xlDown).End(xlDown).End(xlUp).Row
    Range("N" & cola_wo2 & ":N" & cola_wo3).PasteSpecial (xlPasteValues)
           
    'Copiando # do usuário da aba FORM e colando-o na aba LAY_OUT_CONSUMO.
    Sheets("FORM").Select
    Range("j1").Copy
    Sheets("LAY_OUT_CONSUMO").Select
    Range_usuario_ini = Range("a1").End(xlDown).End(xlDown).End(xlUp).Row + 1
    Range_usuario_fim = Range("n1").End(xlDown).End(xlDown).End(xlUp).Row
    Range("A" & Range_usuario_ini & ":A" & Range_usuario_fim).PasteSpecial (xlPasteValues)

   '-------------------------------------------
   'HORAS Colando WO e Seq Op
        
    'HORAS Colando WO
        Sheets("FORM").Select
        Range("C5").Copy
        Sheets("LAY_OUT_HORAS").Select
        Cola_H = Range("O1").End(xlDown).End(xlDown).End(xlUp).Row + 1
        Range("O" & Cola_H).PasteSpecial xlPasteValues
    'HORAS Colando Seq Op
        Sheets("FORM").Select
        Range("E7").Copy
        Sheets("LAY_OUT_HORAS").Select
        Range("X" & Cola_H).PasteSpecial xlPasteValues
    'COLANDO HORAS
        Sheets("FORM").Select
        Range("E8").Copy
        Sheets("LAY_OUT_HORAS").Select
        Range("AL" & Cola_H).PasteSpecial xlPasteValues
    'HORAS Colando WO
        Sheets("FORM").Select
        Range("C5").Copy
        Sheets("LAY_OUT_HORAS").Select
        Cola_H = Range("O1").End(xlDown).End(xlDown).End(xlUp).Row + 1
        Range("O" & Cola_H).PasteSpecial xlPasteValues
    'HORAS Colando Seq Op
        Sheets("FORM").Select
        Range("F7").Copy
        Sheets("LAY_OUT_HORAS").Select
        Range("X" & Cola_H).PasteSpecial xlPasteValues
    'COLANDO HORAS
        Sheets("FORM").Select
        Range("F8").Copy
        Sheets("LAY_OUT_HORAS").Select
        Range("AL" & Cola_H).PasteSpecial xlPasteValues
    'HORAS Colando WO
        Sheets("FORM").Select
        Range("C5").Copy
        Sheets("LAY_OUT_HORAS").Select
        Cola_H = Range("O1").End(xlDown).End(xlDown).End(xlUp).Row + 1
        Range("O" & Cola_H).PasteSpecial xlPasteValues
    'HORAS Colando Seq Op
        Sheets("FORM").Select
        Range("G7").Copy
        Sheets("LAY_OUT_HORAS").Select
        Range("X" & Cola_H).PasteSpecial xlPasteValues
    'COLANDO HORAS
        Sheets("FORM").Select
        Range("G8").Copy
        Sheets("LAY_OUT_HORAS").Select
        Range("AL" & Cola_H).PasteSpecial xlPasteValues
    'HORAS Colando WO
        Sheets("FORM").Select
        Range("C5").Copy
        Sheets("LAY_OUT_HORAS").Select
        Cola_H = Range("O1").End(xlDown).End(xlDown).End(xlUp).Row + 1
        Range("O" & Cola_H).PasteSpecial xlPasteValues
    'HORAS Colando Seq Op
        Sheets("FORM").Select
        Range("H7").Copy
        Sheets("LAY_OUT_HORAS").Select
        Range("X" & Cola_H).PasteSpecial xlPasteValues
    'COLANDO HORAS
        Sheets("FORM").Select
        Range("H8").Copy
        Sheets("LAY_OUT_HORAS").Select
        Range("AL" & Cola_H).PasteSpecial xlPasteValues
    'HORAS Colando WO
        Sheets("FORM").Select
        Range("C5").Copy
        Sheets("LAY_OUT_HORAS").Select
        Cola_H = Range("O1").End(xlDown).End(xlDown).End(xlUp).Row + 1
        Range("O" & Cola_H).PasteSpecial xlPasteValues
    'HORAS Colando Seq Op
        Sheets("FORM").Select
        Range("I7").Copy
        Sheets("LAY_OUT_HORAS").Select
        Range("X" & Cola_H).PasteSpecial xlPasteValues
    'COLANDO HORAS
        Sheets("FORM").Select
        Range("I8").Copy
        Sheets("LAY_OUT_HORAS").Select
        Range("AL" & Cola_H).PasteSpecial xlPasteValues
    'HORAS Colando WO
        Sheets("FORM").Select
        Range("C5").Copy
        Sheets("LAY_OUT_HORAS").Select
        Cola_H = Range("O1").End(xlDown).End(xlDown).End(xlUp).Row + 1
        Range("O" & Cola_H).PasteSpecial xlPasteValues
    'HORAS Colando Seq Op
        Sheets("FORM").Select
        Range("J7").Copy
        Sheets("LAY_OUT_HORAS").Select
        Range("X" & Cola_H).PasteSpecial xlPasteValues
    'COLANDO HORAS
        Sheets("FORM").Select
        Range("J8").Copy
        Sheets("LAY_OUT_HORAS").Select
        Range("AL" & Cola_H).PasteSpecial xlPasteValues
    
    'FILTRO E REMOVE "-" VAZIAS
          
    Sheets("LAY_OUT_HORAS").Select
    ActiveSheet.Range("$A$1:$BL$10000").AutoFilter Field:=38, Criteria1:="=-", _
        Operator:=xlOr, Criteria2:="="
    Range("O2:AL5000").Select
    Selection.SpecialCells(xlCellTypeVisible).Select
    Selection.ClearContents
    ActiveSheet.ShowAllData
    Range("O1").End(xlDown).Select
    'Retornando a aba FORM, calular #WO.
    Sheets("FORM").Select
    Application.ScreenUpdating = True
    'Atualização de estoque após salvar template
      ActiveWorkbook.RefreshAll
 Call limpa_planilha
 Call WO_Anterior
    Range("C5").Select
    Application.ScreenUpdating = False
End If
End Sub

Sub limpa_planilha()

    Application.ScreenUpdating = False
    'USUARIO
    Sheets("FORM").Select
    Range("J1").Copy
    Range("Z1").PasteSpecial
    
    'Apaga tudo e cola a cópia por cima
    Sheets("FORM").Select
    Range("A1:K70").ClearContents
    Range("A1:K70").Clear
    Range("T1").ClearContents
    Range("Q1:AB51").Copy
    Range("A1").PasteSpecial
    Range("Q9:Q200").Copy
    Range("A9").PasteSpecial
    
    'Atualizando formulas
    Range("R10:AA10").Select
    Selection.AutoFill Destination:=Range("R10:AA50"), Type:=xlFillDefault
    Range("B10:E10").Select
    Selection.AutoFill Destination:=Range("B10:E41"), Type:=xlFillValues
    Range("H10:K10").Select
    Selection.AutoFill Destination:=Range("H10:K41"), Type:=xlFillValues
    Range("c5").Select
End Sub
Sub abre_form()
Attribute abre_form.VB_ProcData.VB_Invoke_Func = "a\n14"

' abre_form Macro
' Abre formulário consulta lote.
' Atalho do teclado: Ctrl+a
frm_consulta_lote.Show

End Sub
