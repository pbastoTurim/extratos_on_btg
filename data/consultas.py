from classes.BancoDados import BancoDados
import pandas as pd


def get_accounts()->pd.DataFrame:
    """
    Método que retorna todas as contas do banco BTG.
    :return: dataset
    """
    banco = BancoDados()    
    connection = banco.get_connection()

    query = """
    SELECT 
        c.numero_cliente,
        b.nome_banco,
        cc.agencia,
        cc.conta_corrente 
    FROM 
        global.clientes c
	inner join onshore.clientes_contas cc on c.id_cliente = cc.id_cliente
	inner join onshore.bancos b on b.id_banco = cc.id_banco
    where cc.id_status = 1
	and c.id_status = 1
	and nome_banco like '%BTG%'
    """
    result = banco.get_multiple_result(connection, query)
    df = pd.read_excel(result)
    return df

