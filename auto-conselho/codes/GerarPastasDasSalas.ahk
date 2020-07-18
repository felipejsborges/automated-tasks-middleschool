F7::
Macro1:
MsgBox, 0, , Este programa gerará`, automaticamente`, as planilhas das turmas para o conselho.
classes := ["1º ANO A", "1º ANO B", "1º ANO C", "2º ANO A", "2º ANO B", "2º ANO C", "3º ANO A", "3º ANO B", "3º ANO C", "4º ANO A", "4º ANO B", "4º ANO C", "5º ANO A", "5º ANO B", "5º ANO C", "6º ANO A", "6º ANO B", "6º ANO C", "7º ANO A", "7º ANO B", "7º ANO C", "8º ANO A", "8º ANO B", "8º ANO C", "9º ANO A", "9º ANO B", "9º ANO C"]
FileSelectFile, templateFile, , , Selecione o arquivo de modelo (template).
FileSelectFolder, outputFolder, D:\Users\felipe.jose\Desktop\macros, , Selecione a pasta destino para as planilhas das turmas.
For index, class in classes
{
    FileCopy, %templateFile%, %outputFolder%\%class%.xlsm
}
MsgBox, 0, , Processo finalizado! Abra a pasta destino para ver o resultado.
Return
