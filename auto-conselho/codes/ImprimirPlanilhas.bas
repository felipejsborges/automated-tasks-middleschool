Attribute VB_Name = "Módulo3"
Sub ImprimirPlanilhas()
'ESSE MÓDULO FORMATA E ORGANIZA OS DADOS DA PLANILHA PARA REALIZAR A IMPRESSÃO
    MsgBox "Este programa organiza a planilha de uma sala específica e prepara para impressão"
    
    MsgBox "Primeiramente, escolha o arquivo da sala da qual se deseja imprimir a planilha"
    'Selecionar planilha para impressão
    Dim ArquivoPlaniha As Variant
    ArquivoPlaniha = Application.GetOpenFilename(FileFilter:="Arquivos de Excel (*.xlsm),*.xlsm", Title:="Selecione a planilha que deseja imprimir")
    If ArquivoPlaniha = False Then
        MsgBox "Você deve selecionar uma planilha para executar o programa. Tente novamente"
        Exit Sub
    End If
    
    MsgBox "Agora, selecione a imagem contendo os campos de assinatura"
    'Selecionar imagem com os campos de assinatura
    Dim ArquivoAssinaturas As Variant
    Dim img As Picture
    ArquivoAssinaturas = Application.GetOpenFilename(FileFilter:="Arquivos de Imagem (*.png),*.png", Title:="Selecione a imagem com os campos de assinatura")
    If ArquivoAssinaturas = False Then
        MsgBox "Você deve selecionar uma imagem com os campos de assinatura. Tente novamente"
        Exit Sub
    End If
    
    Set srcwrkbk = ThisWorkbook
    'Abrindo planilha
    Set wrkbk = Workbooks.Open(ArquivoPlaniha)
    Worksheets("Acompanhamento").Activate
    
    'Desprotegendo planilha para poder editar
    ActiveSheet.Unprotect Password:="sme"
    
    'Escondendo columas EFETI
    Columns("BB:BC").Hidden = True
    
    'Organizando tamanho da fonte
    Range("B16:BD65").Font.Size = 10
    
    'Ativando quebra de linha para que todo o conteúdo seja exibido
    Range("BD16:BD65").WrapText = True
    
    'Ajustando tamanho das linhas
    Rows("16:65").AutoFit
    
    'Ajustando largura da coluna com a turma
    With Worksheets("Acompanhamento").Columns("AO:AP")
        .ColumnWidth = 5
    End With
    
    'Contando alunos
    alunos = 0
    Do While Sheets("Acompanhamento").Cells(16 + alunos, 2) <> ""
        alunos = alunos + 1
    Loop
    
    'Escondendo linhas em excesso
    Rows((alunos + 16) & ":68").Hidden = True
    
    'Inserindo linhas para adicionar campo de assinaturas
    Rows("70:70").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
        
    Set img = ActiveSheet.Pictures.Insert(ArquivoAssinaturas)
    With img
       .Left = ActiveSheet.Range("A71").Left
       .Top = ActiveSheet.Range("A71").Top
       .Height = ActiveSheet.Range("A71:B72").Height
       .Width = ActiveSheet.Range("A71:BD71").Width
       .Placement = 1
       .PrintObject = True
    End With

    ActiveSheet.Protect Password:="sme", DrawingObjects:=True, Contents:=True, Scenarios:=True
    
    MsgBox "Pronto! Agora, realize a impressão e feche este arquivo SEM SALVAR."
End Sub
