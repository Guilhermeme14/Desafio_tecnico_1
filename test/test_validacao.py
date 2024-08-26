import unittest
from unittest.mock import MagicMock, patch

import numpy as np
import pandas as pd

from src.validacao.validacao import carregar_planilhas, formatar_resultado, juntar_dados


class TestComissaoProcessamento(unittest.TestCase):
    # Testa a função de carregamento de planilhas
    @patch("src.validacao.validacao.pd.read_excel")
    def test_carregar_planilhas(self, mock_read_excel):
        mock_read_excel.side_effect = [
            pd.DataFrame(
                {"Nome do Vendedor": ["João", "Maria"], "Comissão_y": [100, 200]}
            ),
            pd.DataFrame(
                {"Nome do Vendedor": ["João", "Maria"], "Comissão Final": [100, 200]}
            ),
        ]

        df_vendas, df_comissao = carregar_planilhas(
            "vendas.xlsx", "resultado_comissoes.xlsx"
        )

        self.assertEqual(len(df_vendas), 2)
        self.assertEqual(len(df_comissao), 2)

        mock_read_excel.assert_called()

    # Testa a função de junção de dados das planilhas
    def test_juntar_dados(self):
        df_vendas = pd.DataFrame(
            {"Nome do Vendedor": ["João", "Maria"], "Comissão_y": [100, 200]}
        )
        df_comissao = pd.DataFrame(
            {"Nome do Vendedor": ["João", "Maria"], "Comissão Final": [100, 250]}
        )

        df_juntar = juntar_dados(df_vendas, df_comissao)

        self.assertEqual(df_juntar.shape[0], 2)
        self.assertIn("Valor correto?", df_juntar.columns)
        self.assertEqual(df_juntar.loc[0, "Valor correto?"], "Correto")
        self.assertEqual(df_juntar.loc[1, "Valor correto?"], "Incorreto")

    # Testa a função de formatação do resultado final
    def test_formatar_resultado(self):
        df_juntar = pd.DataFrame(
            {
                "Nome do Vendedor": ["João", "Maria"],
                "Comissão_y": [100, 200],
                "Valor correto?": ["Correto", "Incorreto"],
                "Comissão Final": [100, 250],
            }
        )

        resultado = formatar_resultado(df_juntar)

        # Verifica se as colunas esperadas estão no resultado final
        self.assertIn("Valor Pago", resultado.columns)
        self.assertIn("Valor Correto", resultado.columns)
        self.assertEqual(resultado.loc[0, "Valor Pago"], 100)
        self.assertEqual(resultado.loc[1, "Valor Correto"], 250)
