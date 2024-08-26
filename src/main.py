from calculo.comissao import *
from validacao.validacao import *


def main():
    df = carregar_dados("data/vendas.xlsx", "Vendas")
    resultado = aplicar_calculo_comissoes(df)
    salvar_resultado(df, resultado, "data/Comissões.xlsx")  # type: ignore
    print("Planilha criada 'Comissões.xlsx'")

    df_vendas, df_comissao = carregar_planilhas(
        "data/vendas.xlsx", "data/Comissões.xlsx"
    )
    df_juntar = juntar_dados(df_vendas, df_comissao)
    resultado = formatar_resultado(df_juntar)
    salvar_resultado_(resultado, "data/Comissões corretas.xlsx")
    print("Planilha criada 'Comissões corretas'")


if __name__ == "__main__":
    main()
