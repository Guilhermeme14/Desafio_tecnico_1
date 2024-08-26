# Calculadora de Comissões e Validação de Pagamentos

Este projeto foi desenvolvido para automatizar o processo de cálculo de comissões de vendas e a validação dos valores de comissão pagos, garantindo precisão e eficiência no processo. A solução lê dados de vendas, calcula as comissões devidas com base em regras predefinidas, valida os pagamentos realizados e gera relatórios formatados em planilhas Excel.

## Dependências

- Python 3.11.2

O projeto utiliza as seguintes bibliotecas Python:

- `pandas`
- `openpyxl`
- `numpy`
- `unittest`

Para instalar as dependências, execute:

```bash
pip install pandas openpyxl numpy
```
## Funcionalidades

### Cálculo de Comissões (`src/calculo/comissao.py`)

Este módulo é responsável por calcular as comissões devidas com base nos dados de vendas. Ele realiza as seguintes operações:

- **`carregar_dados(caminho_arquivo, nome_aba):`** Carrega os dados de uma planilha Excel específica.
- **`calcular_comissao(row):`** Calcula a comissão para cada venda considerando regras para o vendedor, marketing e gerente.
- **`aplicar_calculo_comissoes(df):`** Aplica o cálculo de comissões a todas as vendas e agrupa os resultados por vendedor.
- **`salvar_resultado(df, resultado, caminho_arquivo):`** Salva os resultados do cálculo em uma planilha Excel formatada.

### Validação de Pagamentos (`src/validacao/validacao.py`)

Este módulo valida os valores pagos de comissão comparando-os com os valores calculados. As principais operações incluem:

- **`carregar_planilhas(caminho_vendas, caminho_comissao):`** Carrega as planilhas de vendas e comissões.
- **`juntar_dados(df_vendas, df_comissao):`** Mescla as tabelas de vendas e comissões, comparando os valores pagos e os calculados.
- **`formatar_resultado(df_juntar):`** Formata o resultado para visualização, destacando discrepâncias.
- **`salvar_resultado_(resultado, caminho_arquivo):`** Salva o resultado da validação em uma nova planilha Excel.
- 

### Executar `(src/main.py)`

Este script principal coordena a execução dos módulos de cálculo e validação.

- Carrega dados de vendas.
- Calcula comissões e salva o resultado em uma planilha.
- Valida as comissões calculadas e salva o resultado em uma nova planilha.

### Testes

Os testes são realizados utilizando `unittest` e estão divididos em dois arquivos:

* **`test_calculo.py`** : Contém testes para o módulo de cálculo de comissões.
* **`test_validacao.py`** : Contém testes para o módulo de validação de comissões.

Para executar os testes, use o seguinte comando:

```bash
python -m unittest discover -s test
```

## Execução

Para executar o projeto, siga estas etapas:

1. Prepare seus arquivos de dados no diretório `/data/` (por exemplo, `vendas.xlsx`).
2. Execute o script principal:
```bash
python src/main.py
```
Isso irá gerar os seguintes arquivos de saída dentro do diretório `data/`:

`Comissões.xlsx`: Contém as comissões calculadas para cada vendedor.
`Comissões corretas.xlsx`: Contém a validação dos valores pagos aos vendedores, indicando se estão corretos ou não.
