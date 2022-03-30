import os

path = r"C:\001090\SPED EFD Contribuições\SPED EFD Contribuicoes '001090' (2022-02).txt"

arquivo = open(path,'r+')

novoarquivo = []

for registro in arquivo:


    novoarquivo.append(registro)
    #novoarquivo.append('\n')


if '|0120|' in novoarquivo[4]:
    print('Registro 120 ativo')

else:
    print('Não ativo o 120')