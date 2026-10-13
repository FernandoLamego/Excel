import openpyxl

#descrição: esse projeto é um controle financeiro,
#que tem como objetivo controlar as despesas e receitas de uma pessoa,
#podendo ser utilizado para controle pessoal ou familiar.

oppy = openpyxl.Workbook()
oppy.create_sheet("Transações")

aba_ativa = oppy.active

def entrada_dados():
    receitas = []
    despesas = []
    while True:
        
        data = input("insira a data: ")
        if data.lower() == 'sair':
            break
        
        produto_servico = input(str("insira o produto ou serviço: "))
        categoria = input(str("insira a categoria: "))
        valor = float(input("insira o valor: "))
        if categoria.lower() == "receita":
            receitas.append(valor)
        elif categoria.lower() == "despesa":
            despesas.append(valor)
        else:
            print("Categoria inválida. Por favor, insira 'Receita' ou 'Despesa'.")
            continue

        aba_ativa.append([data, produto_servico, categoria, valor])

    
    print(f"Receita: R${sum(receitas):.2f}")
    print(f"Despesa: R${sum(despesas):.2f}")
    print(f"Saldo total: R${sum(receitas) - sum(despesas):.2f}")

    oppy.save("controle_financeiro.xlsx")

# aba_ativa.append(["Data", "Produto/Serviço", "Categoria", "Valor"])
# aba_ativa.append(["2024-06-01", "Salário", "Receita", 5000])
# aba_ativa.append(["2024-06-02", "Aluguel", "Despesa", -1500])
# aba_ativa.append(["2024-06-03", "Supermercado", "Despesa", -300])
# aba_ativa.append(["2024-06-04", "Freelance", "Receita", 800])
# aba_ativa.append(["2024-06-05", "Transporte", "Despesa", -200])
# aba_ativa.append(["2024-06-06", "Investimento", "Receita", 1000])
# aba_ativa.append(["2024-06-07", "Lazer", "Despesa", -150])

oppy.active.append(["Data", "Produto/Serviço", "Categoria", "Valor"])

entrada_dados()
