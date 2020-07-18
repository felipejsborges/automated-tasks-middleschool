import json
import random

nomes = json.load(open('nomes.json'))

with open('nomesFinais.txt', "w", encoding="UTF-8") as file:
    for i in range(215):
        nome = random.choice(nomes)
        sobrenome = random.choice(nomes)

        while(nome == sobrenome):
            sobrenome = random.choice(nomes)

        nomeFinal = nome + " " + sobrenome + "\n"

        # removing special chars
        AccChars = "ŠŽšžŸÀÁÂÃÄÅÇÈÉÊËÌÍÎÏÐÑÒÓÔÕÖÙÚÛÜÝàáâãäåçèéêëìíîïðñòóôõöùúûüýÿ"
        RegChars = "SZszYAAAAAACEEEEIIIIDNOOOOOUUUUYaaaaaaceeeeiiiidnooooouuuuyy"
        for j in nomeFinal:
            oldCharIndex = nomeFinal.find(j)
            newCharIndex = AccChars.find(nomeFinal[oldCharIndex])

            if newCharIndex > -1:
                nomeFinal = nomeFinal.replace(nomeFinal[oldCharIndex], RegChars[newCharIndex])

        # writing final value
        file.write(nomeFinal)
