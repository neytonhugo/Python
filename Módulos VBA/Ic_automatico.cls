VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Planilha1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub IC_automatico()

'Seleciona e apaga linhas vazias
Range("P:P").Select

cont = 1
While cont < 50

Call apagar_linhas_vazias
cont = cont + 1
Wend

'Tirando letra Q antes de começar novamente
Range("AB2").Copy
Range("Z2:Z500").PasteSpecial

Range("Z2:Z500").Copy
Range("P2:P500").PasteSpecial xlPasteValues

'Seleciona formatos da primeira coluna para as demais
Range("A2:P2").Copy
Range("A3:P500").PasteSpecial xlPasteFormats

'Seleciona formulas da primeira coluna e copia para as demais
Range("W2:X2").Copy

Range("T2:U2").PasteSpecial (xlPasteAll)

Range("T2:U2").Copy

Range("T3:U500").PasteSpecial (xlPasteFormulas)

Call Tracos

ActiveWorkbook.Save

Set nome = CreateObject("Wscript.network")

Call Shell("C:\Users\" & nome.UserName & "\OneDrive - Baxter\MFG\Carga Componentes\Arquivos IC automatico\IC_automatico-SIH.exe", vbMinimizedNoFocus)

ThisWorkbook.Close savechanges:=True

'Application.Quit

End Sub

Sub apagar_linhas_vazias()
Attribute apagar_linhas_vazias.VB_ProcData.VB_Invoke_Func = "r\n14"
    
    Linha = Range("P1").End(xlDown).Row + 1
    Range("A" & Linha & ":P" & Linha).Select
    Application.CutCopyMode = False
    Selection.Delete Shift:=xlUp
    
End Sub

Sub Limpar_planilha()

Range("A2:P3000").ClearContents
Range("A2").Select
End Sub

Sub Tracos()

Range("U2:U500").Copy

Range("P2").PasteSpecial (xlPasteValues)

End Sub
