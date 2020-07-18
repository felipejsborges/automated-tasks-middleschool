Attribute VB_Name = "Módulo5"
Sub ImprimirBoletins()
'ESSE MÓDULO SELECIONA A PLANILHA DE BOLETINS DE UMA SALA E ABRE AS OPÇÕES DE IMPRESSÃO
    MsgBox "Este programa navega até os boletins de um sala e abre a caixa de impressão"
    
    MsgBox "Primeiramente, escolha o arquivo da sala da qual se deseja imprimir os boletins"
    'Selecionar planilha para impressão
    Dim ArquivoPlaniha As Variant
    ArquivoPlaniha = Application.GetOpenFilename(FileFilter:="Arquivos de Excel (*.xlsm),*.xlsm", Title:="Selecione a sala")
    If ArquivoPlaniha = False Then
        MsgBox "Você deve selecionar um arquivo de excel para executar o programa. Tente novamente"
        Exit Sub
    End If
    
    Set srcwrkbk = ThisWorkbook
    'Abrindo planilha
    Set wrkbk = Workbooks.Open(ArquivoPlaniha)
    Worksheets("Boletins").Activate
        
    'Abrir caixa de impressão
    Application.Dialogs(xlDialogPrint).Show
    
    'Fechar a planilha sem salvar
    wrkbk.Close SaveChanges:=False
    srcwrkbk.Activate
    
    'Mensagem de sucesso
    MsgBox "Processo finalizado com sucesso."
End Sub
