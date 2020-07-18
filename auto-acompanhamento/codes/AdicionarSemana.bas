Attribute VB_Name = "Módulo3"
Sub AdicionarSemana()
    'INPUT PEDINDO O DIA QUE INICIA E DIA Q FINALIZA A SEMANA
    MsgBox "A seguir, serão solicitados os dias que a semana inicia e se finaliza. Use hífen (-) no lugar das barras(/). Exemplo: 01-01"
    
    'Indicando o dia q inicia a semana
    Dim InicioSemana As Variant
    InicioSemana = Application.InputBox("Insira o dia que inicia a semana conforme exemplo abaixo", Title:="Dia que inicia a semana", Default:="01-01", Type:=2)
    If InicioSemana = False Then
        MsgBox "Você deve inserir um dia para executar o programa. Tente novamente"
        Exit Sub
    End If
    
    'Indicando o dia q finda a semana
    Dim FimSemana As Variant
    FimSemana = Application.InputBox("Insira o último dia da semana conforme exemplo abaixo", Title:="Dia que acaba a semana", Default:="05-01", Type:=2)
    If FimSemana = False Then
        MsgBox "Você deve inserir um dia para executar o programa. Tente novamente"
        Exit Sub
    End If
    
    'Insirindo a pasta onde estão todas as turmas
    Dim PastaPlanilhas As Variant
    PastaPlanilhas = Application.InputBox("Insira a pasta que contém as planilhas das turmas conforme exemplo abaixo", Title:="Pasta das planilhas", Default:="C:\Users\User\Pasta\", Type:=2)
    If PastaPlanilhas = False Then
        MsgBox "Você deve inserir o caminho da pasta que contém as planilhas para executar o programa. Tente novamente"
        Exit Sub
    End If
    
    Dim Turma As Variant, Turmas As Variant
    Turmas = Array("1º ANO A", "1º ANO B", "1º ANO C", "2º ANO A", "2º ANO B", "2º ANO C", "3º ANO A", "3º ANO B", "3º ANO C", "4º ANO A", "4º ANO B", "4º ANO C", "5º ANO A", "5º ANO B", "5º ANO C", "6º ANO A", "6º ANO B", "6º ANO C", "7º ANO A", "7º ANO B", "7º ANO C", "8º ANO A", "8º ANO B", "8º ANO C", "9º ANO A", "9º ANO B", "9º ANO C")
    
    'Repetir para todas as turmas
    For Each Turma In Turmas
        'Abrindo planilha
        Set wrkbkAcompanhamento = Workbooks.Open(PastaPlanilhas & Turma & ".xlsx")
        Worksheets("Modelo").Visible = True
        
        'Adicionar aba com o nome da semana
        Sheets(1).Copy After:=Sheets(1)
        Sheets(2).Name = InicioSemana & " a " & FimSemana
        
        'Alterando semana no cabeçalho
        Range("A5").Select
        ActiveCell.FormulaR1C1 = "Semana de " & InicioSemana & " a " & FimSemana
        
        'Fechando planilha do conselho salvando os dados
        Worksheets("Modelo").Visible = False
        wrkbkAcompanhamento.Close SaveChanges:=True
    Next Turma
    
    'Mensagem de sucesso
    MsgBox "Processo finalizado com sucesso."
End Sub
