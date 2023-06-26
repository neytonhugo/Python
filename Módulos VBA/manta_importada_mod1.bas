Attribute VB_Name = "Módulo1"
Sub Salvando()
'
Set nome = CreateObject("Wscript.network")

'Testando campo de data
If Range("C2").Value = "" Then

MsgBox ("DIGITE A DATA DE PRODUÇÃO")
Range("C2").Select
Exit Sub
End If

'Salvando novo arquivo com novo nome "Data"
DataDia = Range("A2").Value
DataMesAno = Range("A3").Value

    ActiveWorkbook.SaveAs Filename:= _
        "C:\Users\" & nome.UserName & "\OneDrive - Baxter\Teste\Manta Nacional " & DataDia & DataMesAno & ".xlsm", FileFormat:= _
        xlOpenXMLWorkbookMacroEnabled, CreateBackup:=False

'Limpando planilha

Application.ScreenUpdating = False

        Range("A4:A100").ClearContents
        Range("O4:O100").ClearContents
        Range("N4:N100").ClearContents
        Range("P4:P100").ClearContents
        Range("Q4:Q100").ClearContents
        Range("C2").ClearContents
        Range("C35:J50").ClearContents

ThisWorkbook.Close savechanges:=True

End Sub

