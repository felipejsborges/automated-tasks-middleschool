Attribute VB_Name = "Módulo2"
Sub NomParaAcompanhamento()
'COPIA DADOS DA LISTA NOMINAL PARA AS PLANILHAS DE ACOMPANHAMENTO
    'Desativando alertas
    With Application
        .DisplayAlerts = False
        .AlertBeforeOverwriting = False
        .ScreenUpdating = False
    End With
    
    'Iniciando e setando variáveis
    Dim Turma As Variant, Turmas As Variant
    Turmas = Array("1º ANO A", "1º ANO B", "1º ANO C", "2º ANO A", "2º ANO B", "2º ANO C", "3º ANO A", "3º ANO B", "3º ANO C", "4º ANO A", "4º ANO B", "4º ANO C", "5º ANO A", "5º ANO B", "5º ANO C", "6º ANO A", "6º ANO B", "6º ANO C", "7º ANO A", "7º ANO B", "7º ANO C", "8º ANO A", "8º ANO B", "8º ANO C", "9º ANO A", "9º ANO B", "9º ANO C")
            
    'Selecionando lista nominal
    MsgBox "A seguir, selecione o arquivo que contém a lista nominal de todas as salas"
    'Abrindo pasta de trabalho da lista nominal
    Dim ArquivoListaNominal As Variant
    ArquivoListaNominal = Application.GetOpenFilename(FileFilter:="Arquivos de Excel (*.xlsx),*.xlsx", Title:="Selecione a planilha da lista nominal de todos os alunos")
    If ArquivoListaNominal = False Then
        MsgBox "Você deve selecionar uma planilha para executar o programa. Tente novamente"
        Exit Sub
    End If
    Set wrkbkListaNominal = Workbooks.Open(ArquivoListaNominal)
    
    'Indicando caminho da pasta onde estão as planilhas em branco
    Dim CaminhoPastaPlanilhaAcompanhamento As Variant
    CaminhoPastaPlanilhaAcompanhamento = Application.InputBox("Insira o caminho da pasta onde estão as planilhas de acompanhamento conforme exemplo abaixo", Title:="Diretório das planilhas do conselho", Default:="C:\Users\User\Pasta\", Type:=2)
    If CaminhoPastaPlanilhaAcompanhamento = False Then
        MsgBox "Você deve inserir um caminho para executar o programa. Tente novamente"
        Exit Sub
    End If
     
    'Repetindo processo para cada turma
    For Each Turma In Turmas
        'Abrindo planilha da turma na lista nominal
        Set wrkshtListaNominal = wrkbkListaNominal.Worksheets(Turma)
        
        'Abrindo pasta de trabalho da turma e planilha do bimestre selecionado
        Set wrkbkAcompanhamento = Workbooks.Open(CaminhoPastaPlanilhaAcompanhamento & Turma & ".xlsx")
        
        'Arrumando nome da turma
        wrkbkAcompanhamento.Worksheets(1).Range("A4").Copy
        wrkbkAcompanhamento.Worksheets(1).Range("A4").PasteSpecial xlPasteValues
        
        'Preenchendo nome
        wrkshtListaNominal.Range("B9:B48").Copy
        wrkbkAcompanhamento.Worksheets(1).Range("B9").PasteSpecial xlPasteValues
        
        'Preenchendo situação
        wrkshtListaNominal.Range("M9:M48").Copy
        wrkbkAcompanhamento.Worksheets(1).Range("C9").PasteSpecial xlPasteValues
        
        Range("A1").Select
        
        'Fechando planilha do conselho salvando os dados
        wrkbkAcompanhamento.Close SaveChanges:=True
    Next Turma
    
    'Fechando lista nominal sem salvar alterações
    wrkbkListaNominal.Close SaveChanges:=False
    
    'Reativando alertas
    With Application
        .DisplayAlerts = True
        .AlertBeforeOverwriting = True
        .ScreenUpdating = True
    End With
    
    'Mensagem de sucesso
    MsgBox "Processo finalizado com sucesso."
End Sub

