primeiro_valor = input("Digite o primeiro valor: ")
segundo_valor = input("Digite o segundo valor: ")

if primeiro_valor > segundo_valor:
    print("O primeiro valor é maior que o segundo valor.")
elif primeiro_valor < segundo_valor:
    print("O primeiro valor é menor que o segundo valor.")
elif primeiro_valor == segundo_valor:
    print("O primeiro valor é igual ao segundo valor.")
else:
    print("Digite um número válido.")