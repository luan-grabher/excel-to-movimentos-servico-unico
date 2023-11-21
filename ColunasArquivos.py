class ColunasArquivos (object):
    def __init__(self):
        self.colunas_arquivo_recibos = [
            # Nome Razão Social	CPF	CNPJ	Logradouro	Bairro	Cidade	UF	CEP	Data Emissão	CFOP	Numero	Valor NF	Situação
            'Nome Razão Social',
            'CPF',
            'CNPJ',
            'Logradouro',
            'Bairro',
            'Cidade',
            'UF',
            'CEP',
            'Data Emissão',
            'CFOP',
            'Numero',
            'Valor NF',
            'Situação'
        ]
        self.colunas_arquivo_notas = [
            # Nome Razão Social	CPF	CNPJ	Logradouro	Bairro	Cidade	UF	CEP	Data Emissão	CFOP	Numero	Valor Base Calculo	Alíquota ICMS	Valor ICMS	Valor NF	Situação
            'Nome Razão Social',
            'CPF',
            'CNPJ',
            'Logradouro',
            'Bairro',
            'Cidade',
            'UF',
            'CEP',
            'Data Emissão',
            'CFOP',
            'Numero',
            'Valor Base Calculo',
            'Alíquota ICMS',
            'Valor ICMS',
            'Valor NF',
            'Situação'
        ]

        self.colunas_arquivo_movimentos_unico = [
            '#documento',
            'data_aaaammdd',
            'valor_contabil',
            'valor_base_calculo_iss',
            'aliquota_icms',
            'valor_icms',
            'estado'
            # Caso precise ser na ordem que esta no pdf, adiciona as colunas aqui na ordem que esta no pdf e ajusta o metodo get_arquivo_row_as_movimento no arquivo main.py
        ]
