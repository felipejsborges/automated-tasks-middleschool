Attribute VB_Name = "M�dulo2"
Sub NomParaConselho()
'COPIA DADOS DA LISTA NOMINAL PARA AS PLANILHAS DE CONSELHO
    MsgBox "Este programa copia os dados necess�rios dos alunos para compor a planilha de conselho de classe. As planilhas em branco das turmas devem estar na mesma pasta e com o seguinte padr�o de nome: 1� ANO A, 6� ANO C..."

    'Desativando alertas
    With Application
        .DisplayAlerts = False
        .AlertBeforeOverwriting = False
        .ScreenUpdating = False
    End With
    
    'Iniciando e setando vari�veis
    Dim Turmas As Variant, Turma As Variant
    Turmas = Array("1� ANO A", "1� ANO B", "1� ANO C", "2� ANO A", "2� ANO B", "2� ANO C", "3� ANO A", "3� ANO B", "3� ANO C", "4� ANO A", "4� ANO B", "4� ANO C", "5� ANO A", "5� ANO B", "5� ANO C", "6� ANO A", "6� ANO B", "6� ANO C", "7� ANO A", "7� ANO B", "7� ANO C", "8� ANO A", "8� ANO B", "8� ANO C", "9� ANO A", "9� ANO B", "9� ANO C")
    
    Dim CaminhoPastaPlanilhaConselho As Variant
    CaminhoPastaPlanilhaConselho = Application.InputBox("Insira o caminho da pasta onde est�o as planilhas do conselho conforme exemplo abaixo", Title:="Diret�rio das planilhas do conselho", Default:="C:\Users\Usuario\Pasta\", Type:=2)
    If CaminhoPastaPlanilhaConselho = False Then
        MsgBox "Voc� deve inserir um caminho para executar o programa. Tente novamente"
        Exit Sub
    End If
    
    Dim AnoVigente As Variant
    AnoVigente = Application.InputBox("Insira o ano vigente conforme exemplo abaixo", Title:="Ano vigente", Default:="2020", Type:=2)
    If AnoVigente = False Then
        MsgBox "Voc� deve inserir um ano para executar o programa. Tente novamente"
        Exit Sub
    End If
    
    MsgBox "A seguir, selecione o arquivo que cont�m a lista nominal de todas as salas"
    'Abrindo pasta de trabalho da lista nominal
    Dim ArquivoListaNominal As Variant
    ArquivoListaNominal = Application.GetOpenFilename(FileFilter:="Arquivos de Excel (*.xlsx),*.xlsx", Title:="Selecione a planilha da lista nominal de todos os alunos")
    If ArquivoListaNominal = False Then
        MsgBox "Voc� deve selecionar uma planilha para executar o programa. Tente novamente"
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
        
        'Preenchendo professor respons�vel
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
