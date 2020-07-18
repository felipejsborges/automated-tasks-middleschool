Attribute VB_Name = "M�dulo5"
Sub ImprimirBoletins()
'ESSE M�DULO SELECIONA A PLANILHA DE BOLETINS DE UMA SALA E ABRE AS OP��ES DE IMPRESS�O
    MsgBox "Este programa navega at� os boletins de um sala e abre a caixa de impress�o"
    
    MsgBox "Primeiramente, escolha o arquivo da sala da qual se deseja imprimir os boletins"
    'Selecionar planilha para impress�o
    Dim ArquivoPlaniha As Variant
    ArquivoPlaniha = Application.GetOpenFilename(FileFilter:="Arquivos de Excel (*.xlsm),*.xlsm", Title:="Selecione a sala")
    If ArquivoPlaniha = False Then
        MsgBox "Voc� deve selecionar um arquivo de excel para executar o programa. Tente novamente"
        Exit Sub
    End If
    
    Set srcwrkbk = ThisWorkbook
    'Abrindo planilha
    Set wrkbk = Workbooks.Open(ArquivoPlaniha)
    Worksheets("Boletins").Activate
        
    'Abrir caixa de impress�o
    Application.Dialogs(xlDialogPrint).Show
    
    'Fechar a planilha sem salvar
    wrkbk.Close SaveChanges:=False
    srcwrkbk.Activate
    
    'Mensagem de sucesso
    MsgBox "Processo finalizado com sucesso."
End Sub
