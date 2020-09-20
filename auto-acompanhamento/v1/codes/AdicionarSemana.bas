Attribute VB_Name = "M�dulo3"
Sub AdicionarSemana()
    'INPUT PEDINDO O DIA QUE INICIA E DIA Q FINALIZA A SEMANA
    MsgBox "A seguir, ser�o solicitados os dias que a semana inicia e se finaliza. Use h�fen (-) no lugar das barras(/). Exemplo: 01-01"
    
    'Indicando o dia q inicia a semana
    Dim InicioSemana As Variant
    InicioSemana = Application.InputBox("Insira o dia que inicia a semana conforme exemplo abaixo", Title:="Dia que inicia a semana", Default:="01-01", Type:=2)
    If InicioSemana = False Then
        MsgBox "Voc� deve inserir um dia para executar o programa. Tente novamente"
        Exit Sub
    End If
    
    'Indicando o dia q finda a semana
    Dim FimSemana As Variant
    FimSemana = Application.InputBox("Insira o �ltimo dia da semana conforme exemplo abaixo", Title:="Dia que acaba a semana", Default:="05-01", Type:=2)
    If FimSemana = False Then
        MsgBox "Voc� deve inserir um dia para executar o programa. Tente novamente"
        Exit Sub
    End If
    
    'Insirindo a pasta onde est�o todas as turmas
    Dim PastaPlanilhas As Variant
    PastaPlanilhas = Application.InputBox("Insira a pasta que cont�m as planilhas das turmas conforme exemplo abaixo", Title:="Pasta das planilhas", Default:="C:\Users\User\Pasta\", Type:=2)
    If PastaPlanilhas = False Then
        MsgBox "Voc� deve inserir o caminho da pasta que cont�m as planilhas para executar o programa. Tente novamente"
        Exit Sub
    End If
    
    Dim Turma As Variant, Turmas As Variant
    Turmas = Array("1� ANO A", "1� ANO B", "1� ANO C", "2� ANO A", "2� ANO B", "2� ANO C", "3� ANO A", "3� ANO B", "3� ANO C", "4� ANO A", "4� ANO B", "4� ANO C", "5� ANO A", "5� ANO B", "5� ANO C", "6� ANO A", "6� ANO B", "6� ANO C", "7� ANO A", "7� ANO B", "7� ANO C", "8� ANO A", "8� ANO B", "8� ANO C", "9� ANO A", "9� ANO B", "9� ANO C")
    
    'Repetir para todas as turmas
    For Each Turma In Turmas
        'Abrindo planilha
        Set wrkbkAcompanhamento = Workbooks.Open(PastaPlanilhas & Turma & ".xlsx")
        Worksheets("Modelo").Visible = True
        
        'Adicionar aba com o nome da semana
        Sheets(1).Copy After:=Sheets(1)
        Sheets(2).Name = InicioSemana & " a " & FimSemana
        
        'Alterando semana no cabe�alho
        Range("A5").Select
        ActiveCell.FormulaR1C1 = "Semana de " & InicioSemana & " a " & FimSemana
        
        'Fechando planilha do conselho salvando os dados
        Worksheets("Modelo").Visible = False
        wrkbkAcompanhamento.Close SaveChanges:=True
    Next Turma
    
    'Mensagem de sucesso
    MsgBox "Processo finalizado com sucesso."
End Sub
