Attribute VB_Name = "Módulo3"
Function Acento(Caract As String)
Dim A As String
Dim B As String
Dim i As Integer
Const AccChars = "ŠšŸÀÁÂÃÄÅÇÈÉÊËÌÍÎÏĞÑÒÓÔÕÖÙÚÛÜİàáâãäåçèéêëìíîïğñòóôõöùúûüıÿ"
Const RegChars = "SZszYAAAAAACEEEEIIIIDNOOOOOUUUUYaaaaaaceeeeiiiidnooooouuuuyy"
For i = 1 To Len(AccChars)
A = Mid(AccChars, i, 1)
B = Mid(RegChars, i, 1)
Caract = Replace(Caract, A, B)
Next
Acento = Caract
End Function

