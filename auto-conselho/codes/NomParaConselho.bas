Attribute VB_Name = "Módulo2"
Sub NomParaConselho()
'COPIA DADOS DA LISTA NOMINAL PARA AS PLANILHAS DE CONSELHO
    MsgBox "Este programa copia os dados necessários dos alunos para compor a planilha de conselho de classe. As planilhas em branco das turmas devem estar na mesma pasta e com o seguinte padrão de nome: 1º ANO A, 6º ANO C..."

    'Desativando alertas
    With Application
        .DisplayAlerts = False
        .AlertBeforeOverwriting = False
        .ScreenUpdating = False
    End With
    
    'Iniciando e setando variáveis
    Dim Turmas As Variant, Turma As Variant
    Turmas = Array("1º ANO A", "1º ANO B", "1º ANO C", "2º ANO A", "2º ANO B", "2º ANO C", "3º ANO A", "3º ANO B", "3º ANO C", "4º ANO A", "4º ANO B", "4º ANO C", "5º ANO A", "5º ANO B", "5º ANO C", "6º ANO A", "6º ANO B", "6º ANO C", "7º ANO A", "7º ANO B", "7º ANO C", "8º ANO A", "8º ANO B", "8º ANO C", "9º ANO A", "9º ANO B", "9º ANO C")
    
    Dim CaminhoPastaPlanilhaConselho As Variant
    CaminhoPastaPlanilhaConselho = Application.InputBox("Insira o caminho da pasta onde estão as planilhas do conselho conforme exemplo abaixo", Title:="Diretório das planilhas do conselho", Default:="C:\Users\Usuario\Pasta\", Type:=2)
    If CaminhoPastaPlanilhaConselho = False Then
        MsgBox "Você deve inserir um caminho para executar o programa. Tente novamente"
        Exit Sub
    End If
    
    Dim AnoVigente As Variant
    AnoVigente = Application.InputBox("Insira o ano vigente conforme exemplo abaixo", Title:="Ano vigente", Default:="2020", Type:=2)
    If AnoVigente = False Then
        MsgBox "Você deve inserir um ano para executar o programa. Tente novamente"
        Exit Sub
    End If
    
    MsgBox "A seguir, selecione o arquivo que contém a lista nominal de todas as salas"
    'Abrindo pasta de trabalho da lista nominal
    Dim ArquivoListaNominal As Variant
    ArquivoListaNominal = Application.GetOpenFilename(FileFilter:="Arquivos de Excel (*.xlsx),*.xlsx", Title:="Selecione a planilha da lista nominal de todos os alunos")
    If ArquivoListaNominal = False Then
        MsgBox "Você deve selecionar uma planilha para executar o programa. Tente novamente"
        Exit Sub
    End If
    Set wrkbkListaNominal = Workbooks.Open(ArquivoListaNominal)
    
    'Repetindo processo para cada turma
    For Each Turma In Turmas
        'Abrindo planilha da turma na lista nominal
        Set wrkshtListaNominal = wrkbkListaNominal.Worksheets(Turma)
        
        'Abrindo pasta de trabalho da turma e planilha do bimestre selecionado
        Set wrkbkConselho = Workbooks.Open(CaminhoPastaPlanilhaConselho & Turma & ".xlsm")
        Set wrkshtConselho = Worksheets("Acompanhamento")
        
        'Desprotegendo planilha para poder editar
        ActiveSheet.Unprotect Password:="sme"
    
        'Preenchendo nome da turma
        Range("AO1").Select
        ActiveCell.FormulaR1C1 = Turma
    
        'Preenchendo ano vigente
        Range("AY1").Select
        ActiveCell.FormulaR1C1 = AnoVigente
    
        'Preenchendo nome e data de nascimento
        wrkshtListaNominal.Range("B9:C43").Copy
        wrkshtConselho.Range("B16").PasteSpecial xlPasteValues
    
        'Preenchendo nome da escola
        wrkshtConselho.Range("D1").Select
        Selection.UnMerge
        wrkshtListaNominal.Range("A3").Copy
        wrkshtConselho.Range("D1").PasteSpecial xlPasteValues
        wrkshtConselho.Range("D1:AI1").Select
        Selection.Merge
        
        'Preenchendo professor responsável
        wrkshtConselho.Range("A3").Select
        Selection.UnMerge
        wrkshtListaNominal.Range("A6").Copy
        wrkshtConselho.Range("A3").PasteSpecial xlPasteValues
        wrkshtConselho.Range("A3:F3").Select
        Selection.Merge
        
        'Protegendo planilha para fechar
        ActiveSheet.Protect Password:="sme", DrawingObjects:=True, Contents:=True, Scenarios:=True
        
        'Fechando planilha do conselho salvando os dados
        wrkbkConselho.Close SaveChanges:=True
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
