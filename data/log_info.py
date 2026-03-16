import pandas as pd


def insert_line(df : pd.DataFrame, date:str) -> None:

    """
    Método que insere uma linha no DataFrame com a data atual.
    :param df: DataFrame a ser modificado.
    :param date: data a ser inserida.
    :return: None
    """

    df_log = pd.read_excel(r'data/log.xlsx')
    year = date.split('/')[2]
    month = date.split('/')[1]
    actual_sheet = f'{month}_{year}'
    if actual_sheet not in df_log.sheet_names:
        with pd.ExcelWriter(r'data/log.xlsx', engine='openpyxl', mode='a') as writer:
            df.to_excel(writer, sheet_name=actual_sheet, index=False)
    else:
        with pd.ExcelWriter(r'data/log.xlsx', engine='openpyxl', mode='a') as writer:
            df.to_excel(writer, sheet_name=actual_sheet, index=False, startrow=len(df_log[actual_sheet]) + 1, header=False)


def get_last_client():

    """
    Método que retorna o último cliente baixado.
    :return: string com o último cliente baixado.
    """
    try:
        # Ler o arquivo Excel
        df = pd.read_excel(r'data/log.xlsx')
        # Verificar se a coluna 'Clientes' existe
        if 'Clientes' not in df.columns:
            raise ValueError("A coluna 'Clientes' não foi encontrada no arquivo.")
        
        ultimo_cliente = df.iloc[-1]['Clientes']
        return ultimo_cliente
    
    except FileNotFoundError:
        print("O arquivo 'log.xlsx' não foi encontrado.")
        return None
    except Exception as e:
        print(f"Ocorreu um erro: {e}")
        return None
    