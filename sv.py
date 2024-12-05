
saldo = 1900

nome = input("digite useu nome: \n")

senha = (input("digite uma senha: \n"))

deposito1 = input("Você deseja fazer um deposito?(s/n)")

if deposito1 == 's':
    depositar = float(input("\n Digite o valor do deposito: \n"))
    senha1 = (input("\n Digite sua senha para fazer o deposito: \n"))
    
    while senha1 != senha:
            tentativas = 5
            tentativas = tentativas - 1
            senha1 = (input("\n Digite sua senha para fazer o deposito: \n"))
            if tentativas <= 0:
                print("O deposito não foi efetuador!")
                break

    
    if senha1 == senha:
        saldo = saldo + depositar
        print(f"Deposito efetuado com sucesso, valor da sua conta bancaria {saldo}")

if deposito1 == 'n':
     saque1 = input("Você deseja realizar um saque?(s/n)")

     if saque1 == 's':
          saque2 = float(input("\nDigite o valor do saque que deseja realizar: \n"))
          senha1 = (input("\n Digite sua senha para fazer o deposito: \n"))
    
     while senha1 != senha:
            tentativas = 5
            tentativas = tentativas - 1
            senha1 = (input("\nDigite sua senha para fazer o deposito: \n"))
            if tentativas <= 0:
                print("O deposito não foi efetuador!")
                break

     if senha1 == senha:
        if saldo < saque2:
            print("saldo insuficiente!")
            

        elif saldo < saque2:
            saldo = saldo - saque2
            print(f"Saque efetuado com sucesso, valor da sua conta bancaria {saldo}")
