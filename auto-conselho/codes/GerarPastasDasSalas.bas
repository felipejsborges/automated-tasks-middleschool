Attribute VB_Name = "Módulo1"
Sub GerarPastasDasSalas()
    Dim foldersName As String
    foldersName = ActiveWorkbook.Path
    Call Shell(foldersName & "\gerador_salas.exe")
    Application.SendKeys ("{F7}")
End Sub
