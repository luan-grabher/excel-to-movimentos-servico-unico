import os
os.system('python.exe -m pip install --upgrade pip')
os.system('pip install -r requirements.txt')
os.system('cls')

from tkinter import messagebox
from easygui import fileopenbox
import pandas as pd
from ColunasArquivos import ColunasArquivos
from tqdm import tqdm

def main():

    try:
        messagebox.showinfo('Escolha', 'Escolha o arquivo de Recibos ou Notas')
        arquivo = fileopenbox(filetypes=['*.xlsx'], title='Escolha o arquivo de Recibos ou Notas', default='*.xlsx')

        if arquivo == None:
            raise Exception('Nenhum arquivo selecionado')
        
        if not arquivo.endswith('.xlsx'):
            raise Exception('Arquivo não é um arquivo excel')
        

        print('Lendo arquivo...')
        df = pd.read_excel(arquivo)

        print('Ajustando arquivo...')
        df = df.fillna('')
        df = df.drop_duplicates()
        df = df.reset_index(drop=True)
        df = df.drop(df[(df['Data Emissão'] == '') & (df['Valor NF'] == '') & (df['CNPJ'] == '') & (df['CPF'] == '')].index)

        colunas_arquivos = ColunasArquivos()
        colunas_arquivo_recibos = colunas_arquivos.colunas_arquivo_recibos
        colunas_arquivo_notas = colunas_arquivos.colunas_arquivo_notas
        colunas_arquivo_movimentos_unico = colunas_arquivos.colunas_arquivo_movimentos_unico


        is_arquivo_recibo = has_this_cols_on_df(df, colunas_arquivo_recibos)
        is_arquivo_nota = has_this_cols_on_df(df, colunas_arquivo_notas)

        if not is_arquivo_recibo and not is_arquivo_nota:
            raise Exception('Arquivo não é um arquivo de Recibo ou Nota')
        
        movimentos_unico = []

        for i, row in tqdm(df.iterrows(), total=len(df), desc='Convertendo arquivo para movimentos unico'):
            if i == 0:
                continue

            movimento_unico = get_arquivo_row_as_movimento(row)
            if movimento_unico != None:
                movimentos_unico.append(movimento_unico)

        
        print('Convertendo movimentos para txt...')
        df_movimentos_unico = pd.DataFrame(movimentos_unico, columns=colunas_arquivo_movimentos_unico)
        df_movimentos_unico = df_movimentos_unico.fillna('')
        df_movimentos_unico = df_movimentos_unico.drop_duplicates()
        df_movimentos_unico = df_movimentos_unico.reset_index(drop=True)


        #save on desktop as csv
        desktop = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop')
        nome_arquivo_csv = 'movimentos_unico_adyl_' + ('nota' if is_arquivo_nota else 'recibo') + '.csv'
        path_arquivo_csv = os.path.join(desktop, nome_arquivo_csv)
        df_movimentos_unico.to_csv(path_arquivo_csv, sep=';', encoding='utf-8-sig', index=False)

        messagebox.showinfo('Sucesso', 'Arquivo salvo em: {}'.format(path_arquivo_csv))

    except Exception as e:
        messagebox.showerror('Erro', '{}'.format(e))

def has_this_cols_on_df(df, cols):
    for col in cols:
        if col not in df.columns:
            return False
    return True

def get_arquivo_row_as_movimento(dt_row):
    documento = str(dt_row['CNPJ'] if dt_row['CNPJ'] != '' else dt_row['CPF'])
    documento_only_numbers = ''.join(filter(str.isdigit, documento))
    if len(documento_only_numbers) == 0:
        documento = ''

    data_emissao_str = str(dt_row['Data Emissão'])
    if data_emissao_str == '':
        return None

    data_emissao_str = data_emissao_str.split('/')
    mesMM = data_emissao_str[1]
    if len(mesMM) == 1:
        mesMM = '0' + mesMM

    diaDD = data_emissao_str[0]
    if len(diaDD) == 1:
        diaDD = '0' + diaDD

    data_aaaammdd = data_emissao_str[2] + mesMM + diaDD
    
    valor_contabil = dt_row['Valor NF']    
    valor_contabil = str(valor_contabil).replace('.', ',') if valor_contabil != '' else ''

    valor_base_calculo_iss = dt_row['Valor Base Calculo'] if 'Valor Base Calculo' in dt_row else ''
    valor_base_calculo_iss = str(valor_base_calculo_iss).replace('.', ',') if valor_base_calculo_iss != '' else ''

    aliquota_icms = dt_row['Alíquota ICMS'] if 'Alíquota ICMS' in dt_row else ''
    aliquota_icms = str(aliquota_icms).replace('.', ',') if aliquota_icms != '' else ''

    valor_icms = dt_row['Valor ICMS'] if 'Valor ICMS' in dt_row else ''
    valor_icms = str(valor_icms).replace('.', ',') if valor_icms != '' else ''

    estado = dt_row['UF']
    
    return [
        documento,
        data_aaaammdd,
        valor_contabil,
        valor_base_calculo_iss,
        aliquota_icms,
        valor_icms,    
        estado
        # Caso precise ser na ordem que esta no pdf, adiciona as colunas aqui na ordem que esta no pdf
    ]


if __name__ == '__main__':
    main()