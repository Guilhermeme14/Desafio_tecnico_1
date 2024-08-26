import unittest
from unittest.mock import MagicMock, patch

import pandas as pd

from src.calculo.comissao import (
    aplicar_calculo_comissoes,
    calcular_comissao,
    carregar_dados,
)


class TestComissaoFunctions(unittest.TestCase):

    @patch("src.calculo.comissao.pd.read_excel")
    def test_carregar_dados_sucesso(self, mock_read_excel):
        # Simula o retorno de um DataFrame
        mock_df = pd.DataFrame({"A": [1, 2], "B": [3, 4]})
        mock_read_excel.return_value = mock_df

        resultado = carregar_dados("teste.xlsx", "Teste")
        mock_read_excel.assert_called_once_with("teste.xlsx", sheet_name="Teste")
        pd.testing.assert_frame_equal(resultado, mock_df)

    def test_calcular_comissao_venda_online(self):
        row = {"Valor da Venda": 2000, "Canal de Venda": "Online"}
        comissao, comissao_marketing, comissao_gerente, comissao_final = (
            calcular_comissao(row)
        )

        self.assertEqual(comissao, 200.0)
        self.assertEqual(comissao_marketing, 40.0)
        self.assertEqual(comissao_gerente, 0.0)
        self.assertEqual(comissao_final, 160.0)

    def test_calcular_comissao_venda_offline(self):
        row = {"Valor da Venda": 1000, "Canal de Venda": "Offline"}
        comissao, comissao_marketing, comissao_gerente, comissao_final = (
            calcular_comissao(row)
        )

        self.assertEqual(comissao, 100.0)
        self.assertEqual(comissao_marketing, 0.0)
        self.assertEqual(comissao_gerente, 0.0)
        self.assertEqual(comissao_final, 100.0)

    def test_aplicar_calculo_comissoes(self):
        df = pd.DataFrame(
            {
                "Nome do Vendedor": ["Vendedor 1", "Vendedor 2"],
                "Valor da Venda": [2000, 1000],
                "Canal de Venda": ["Online", "Offline"],
            }
        )
        resultado = aplicar_calculo_comissoes(df)

        df_esperado = pd.DataFrame(
            {
                "Nome do Vendedor": ["Vendedor 1", "Vendedor 2"],
                "Comissão": [200.0, 100.0],
                "Comissão Marketing": [40.0, 0.0],
                "Comissão Gerente": [0.0, 0.0],
                "Comissão Final": [160.0, 100.0],
            }
        )

        pd.testing.assert_frame_equal(resultado, df_esperado)
