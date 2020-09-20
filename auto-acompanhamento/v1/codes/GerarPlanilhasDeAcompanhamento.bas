Attribute VB_Name = "Módulo1"
Sub GerarPlanilhasDeAcompanhamento()
    'Criando planilhas para todas as salas
    Dim foldersName As String
    foldersName = ActiveWorkbook.Path
    Call Shell(foldersName & "\gerador.exe")
    Application.SendKeys ("{F7}")
End Sub
