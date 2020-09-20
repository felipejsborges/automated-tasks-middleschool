Attribute VB_Name = "M�dulo2"
Sub NomParaAcompanhamento()
'COPIA DADOS DA LISTA NOMINAL PARA AS PLANILHAS DE ACOMPANHAMENTO
    'Desativando alertas
    With Application
        .DisplayAlerts = False
        .AlertBeforeOverwriting = False
        .ScreenUpdating = False
    End With
    
    'Iniciando e setando vari�veis
    Dim Turma As Variant, Turmas As Variant
    Turmas = Array("1� ANO A", "1� ANO B", "1� ANO C", "2� ANO A", "2� ANO B", "2� ANO C", "3� ANO A", "3� ANO B", "3� ANO C", "4� ANO A", "4� ANO B", "4� ANO C", "5� ANO A", "5� ANO B", "5� ANO C", "6� ANO A", "6� ANO B", "6� ANO C", "7� ANO A", "7� ANO B", "7� ANO C", "8� ANO A", "8� ANO B", "8� ANO C", "9� ANO A", "9� ANO B", "9� ANO C")
            
    'Selecionando lista nominal
    MsgBox "A seguir, selecione o arquivo que cont�m a lista nominal de todas as salas"
    'Abrindo pasta de trabalho da lista nominal
    Dim ArquivoListaNominal As Variant
    ArquivoListaNominal = Application.GetOpenFilename(FileFilter:="Arquivos de Excel (*.xlsx),*.xlsx", Title:="Selecione a planilha da lista nominal de todos os alunos")
    If ArquivoListaNominal = False Then
        MsgBox "Voc� deve selecionar uma planilha para executar o programa. Tente novamente"
        Exit Sub
    End If
    Set wrkbkListaNominal = Workbooks.Open(ArquivoListaNominal)
    
    'Indicando caminho da pasta onde est�o as planilhas em branco
    Dim CaminhoPastaPlanilhaAcompanhamento As Variant
    CaminhoPastaPlanilhaAcompanhamento = Application.InputBox("Insira o caminho da pasta onde est�o as planilhas de acompanhamento conforme exemplo abaixo", Title:="Diret�rio das planilhas do conselho", Default:="C:\Users\User\Pasta\", Type:=2)
    If CaminhoPastaPlanilhaAcompanhamento = False Then
        MsgBox "Voc� deve inserir um caminho para executar o programa. Tente novamente"
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
        
        'Preenchendo situa��o
        wrkshtListaNominal.Range("M9:M48").Copy
        wrkbkAcompanhamento.Worksheets(1).Range("C9").PasteSpecial xlPasteValues
        
        Range("A1").Select
        
        'Fechando planilha do conselho salvando os dados
        wrkbkAcompanhamento.Close SaveChanges:=True
    Next Turma
    
    'Fechando lista nominal sem salvar altera��es
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

