F7::
Macro1:
classesIniciais := ["1º ANO A", "1º ANO B", "1º ANO C", "2º ANO A", "2º ANO B", "2º ANO C", "3º ANO A", "3º ANO B", "3º ANO C", "4º ANO A", "4º ANO B", "4º ANO C", "5º ANO A", "5º ANO B", "5º ANO C"]
classesFinais := ["6º ANO A", "6º ANO B", "6º ANO C", "7º ANO A", "7º ANO B", "7º ANO C", "8º ANO A", "8º ANO B", "8º ANO C", "9º ANO A", "9º ANO B", "9º ANO C"]
MsgBox, 0, , A seguir`, selecione a planilha MODELO - ANOS INICIAIS
FileSelectFile, templateFileIniciais, , , Selecione o arquivo de modelo dos anos iniciais.
MsgBox, 0, , A seguir`, selecione a planilha MODELO - ANOS FINAIS
FileSelectFile, templateFileFinais, , , Selecione o arquivo de modelo dos anos finais.
MsgBox, 0, , A seguir`, selecione a pasta destino para criar as planilhas de cada turma
FileSelectFolder, outputFolder, D:\Users\felipe.jose\Desktop\macros, , Selecione a pasta destino para as planilhas das turmas.
For index, class in classesIniciais
{
    FileCopy, %templateFileIniciais%, %outputFolder%\%class%.xlsx
}
For index, class in classesFinais
{
    FileCopy, %templateFileFinais%, %outputFolder%\%class%.xlsx
}
MsgBox, 0, , Processo finalizado! Navegue até a pasta destino para conferir.
Return
