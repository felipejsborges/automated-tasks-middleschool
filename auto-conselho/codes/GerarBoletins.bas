Attribute VB_Name = "Módulo4"
Sub GerarBoletins()
'GERADOR DE BOLETINS DE TODAS AS SALAS
    MsgBox "Este programa gera automaticamente os boletins de todas as salas."
    
    'Desativando alertas do Excel
    With Application
        .DisplayAlerts = False
        .AlertBeforeOverwriting = False
        .ScreenUpdating = False
    End With
    
    'Copiando o caminho para a pasta com as planilhas
    Dim CaminhoPastaPlanilhaConselho As Variant
    CaminhoPastaPlanilhaConselho = Application.InputBox("Insira o caminho da pasta onde estão as planilhas do conselho conforme exemplo abaixo", Title:="Diretório das planilhas do conselho", Default:="C:\Users\Usuario\Pasta\", Type:=2)
    If CaminhoPastaPlanilhaConselho = False Then
        MsgBox "Você deve inserir um caminho para executar o programa. Tente novamente"
        Exit Sub
    End If
    
    'Criando lista de turmas
    Dim Turmas As Variant, Turma As Variant
    Turmas = Array("1º ANO A", "1º ANO B", "1º ANO C", "2º ANO A", "2º ANO B", "2º ANO C", "3º ANO A", "3º ANO B", "3º ANO C", "4º ANO A", "4º ANO B", "4º ANO C", "5º ANO A", "5º ANO B", "5º ANO C", "6º ANO A", "6º ANO B", "6º ANO C", "7º ANO A", "7º ANO B", "7º ANO C", "8º ANO A", "8º ANO B", "8º ANO C", "9º ANO A", "9º ANO B", "9º ANO C")
        
    'Repetindo para cada turma
    For Each Turma In Turmas
        'Abrindo planilha
        Set wrkbk = Workbooks.Open(CaminhoPastaPlanilhaConselho + Turma)
        
        'Contando quantidade de alunos da sala
        Sheets("Acompanhamento").Select
        alunos = 0
        Do While Sheets("Acompanhamento").Cells(16 + alunos, 2) <> ""
            alunos = alunos + 1
        Loop
        
        'Apagando dados existentes, caso existam
        Sheets("Boletins").Select
        Columns("A:Z").Select
        Selection.Delete Shift:=xlToLeft
        
        'Copiando dados para os boletins
        For x = 0 To alunos - 1
            'Copiando modelo de boletim
            Sheets("Ficha Modelo").Visible = True
            Sheets("Ficha Modelo").Select
            Range("A1:O47").Select
            Selection.Copy
            Sheets("Boletins").Select
            Range("A" & (x * 47) + 1).Select
            ActiveSheet.Paste
        
            'Copiando dados Basicos
            Sheets("Boletins").Cells((x * 47) + 5, 1) = "ANO LETIVO " + Str(Sheets("Acompanhamento").Cells(1, 51))
            Sheets("Boletins").Cells((x * 47) + 6, 2) = Sheets("Acompanhamento").Cells(1, 4)
            Sheets("Boletins").Cells((x * 47) + 7, 2) = Sheets("Acompanhamento").Cells(x + 16, 2)
            Sheets("Boletins").Cells((x * 47) + 7, 14) = Sheets("Acompanhamento").Cells(x + 16, 3)
            Sheets("Boletins").Cells((x * 47) + 8, 2) = Sheets("Acompanhamento").Cells(1, 41)
            Sheets("Boletins").Cells((x * 47) + 8, 4) = Sheets("Acompanhamento").Cells(1, 51)
            Sheets("Boletins").Cells((x * 47) + 8, 8) = Sheets("Acompanhamento").Cells(3, 1)
        
            'Copiando conceitos
            Sheets("Boletins").Cells((x * 47) + 11, 8) = Sheets("Acompanhamento").Cells(x + 16, 10)
            Sheets("Boletins").Cells((x * 47) + 12, 8) = Sheets("Acompanhamento").Cells(x + 16, 12)
            Sheets("Boletins").Cells((x * 47) + 13, 8) = Sheets("Acompanhamento").Cells(x + 16, 14)
            Sheets("Boletins").Cells((x * 47) + 14, 8) = Sheets("Acompanhamento").Cells(x + 16, 16)
            Sheets("Boletins").Cells((x * 47) + 15, 8) = Sheets("Acompanhamento").Cells(x + 16, 18)
            Sheets("Boletins").Cells((x * 47) + 16, 8) = Sheets("Acompanhamento").Cells(x + 16, 20)
            Sheets("Boletins").Cells((x * 47) + 17, 8) = Sheets("Acompanhamento").Cells(x + 16, 22)
            Sheets("Boletins").Cells((x * 47) + 18, 8) = Sheets("Acompanhamento").Cells(x + 16, 24)
            Sheets("Boletins").Cells((x * 47) + 19, 8) = Sheets("Acompanhamento").Cells(x + 16, 26)
            Sheets("Boletins").Cells((x * 47) + 19, 1) = Sheets("Acompanhamento").Cells(5, 26)
        
            'Quebras de página na horizontal
            Range("A1:N" & (alunos * 47)).Select
            ActiveSheet.PageSetup.PrintArea = "$A$1:$N$" & (alunos * 47)
            ActiveSheet.HPageBreaks.Add Before:=Cells((35 + (47 * x)), 1)
            ActiveSheet.HPageBreaks.Add Before:=Cells((48 + (47 * x)), 1)
        Next x
        
        'Quebra de página vertical
        Sheets("Boletins").Select
        ActiveWindow.View = xlPageBreakPreview
        ActiveSheet.VPageBreaks(1).DragOff Direction:=xlToRight, RegionIndex:=1
        
        'Ocultando modelo de boletim
        Sheets("Ficha Modelo").Visible = False
        
        'Fechando planilha e salvando
        wrkbk.Close SaveChanges:=True
    Next Turma
    
    'Reativando alertas
    With Application
        .DisplayAlerts = True
        .AlertBeforeOverwriting = True
        .ScreenUpdating = True
    End With
    
    'Mensagem de sucesso
    MsgBox "Processo finalizado com sucesso."
End Sub
