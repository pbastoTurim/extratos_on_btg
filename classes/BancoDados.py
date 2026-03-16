import pyodbc
import logging
import datetime

driver = '{SQL Server};'
server = 'servidorbdsp;'
database = 'Offshore_web;'

class BancodeDados:
    

    @staticmethod
    def db_connection(driver=driver, server=server, database=database):
        """
        Establish a connection to a database using pyodbc lib
        :param driver: str
            database driver name
        :param server: str
            database server
        :param database: str
            database name
        :return: pyodbc connection object
        """
        try:
            connection = pyodbc.connect(f'Driver={driver};Server={server};Database={database};Trusted_Connection=yes;')
            if not connection:  # CASO NÃO HAJA CONEXÃO COM O BANCO DE DADOS
                print('Erro ao se conectar ao banco de dados')
        except Exception as error:
            print(f'Erro ao se conectar ao banco de dados: {error}')
            raise Exception(f'Erro ao se conectar ao banco de dados. \n \n {error}')
        return connection

    @staticmethod
    def execute_query(connection, query, params=None, many=False, persistence=False):
        cursor = connection.cursor()
        try:
            if many is False:
                if params is None:
                    cursor.execute(query)
                else:
                    cursor.execute(query, params)
            else:
                cursor.fast_executemany = True
                cursor.executemany(query, params)

            if persistence:
                cursor.commit()
            
                if cursor.rowcount > 1:
                    logging.info(f'Affected rows: {cursor.rowcount}')
        except pyodbc.IntegrityError as ie:
            logging.error('An integrity error has occurred.', exc_info=True)
            raise ie
        except Exception as e:
            logging.error(f"Error while running a query execution: {e}")
            raise Exception(f"Error while running a query execution: {e}")
        finally:
            cursor.close()
            connection.close()

    @staticmethod
    def get_single_result(connection, query, *args):
        """
        Método que retorna um resultado único do banco.
        :param connection: pyodbc connection obj
        :param query: str
        :param args: tuple
        :return: dataset
        """
        try:
            cursor = connection.cursor()
            cursor.execute(query, *args)
            dados = cursor.fetchone()
            result = dados[0]
            return result
        except Exception as error:
            raise Exception(f'Erro ao executar a query "{query}" \n \n ERRO: {error}')
        finally:
            cursor.close()
            connection.close()

    @staticmethod
    def get_multiple_result(connection, query, *args):
        """
        Método que retorna múltiplos resultados do banco
        :param connection: pyodbc connection obj
        :param query: str
        :param args: tuple
        :return: dataset
        """
        cursor = connection.cursor()
        results = []  # Lista para armazenamento dos resultados junto os nomes de colunas
        try:
            cursor.execute(query, *args)
            columns = [column[0] for column in cursor.description]  # Capturando as colunas do dataset
            for row in cursor.fetchall():  # Para cada linha do resultado do banco, mesclar o resultado do banco junto ao nome da coluna.
                results.append(dict(zip(columns, row)))
            return results
        except Exception as error:
            raise Exception(f'Erro ao executar a query "{query}" \n \n ERRO: {error}')
        finally:
            cursor.close()
            connection.close()

    @staticmethod
    def execute_procedure(connection, query):
        no_count = 'SET NOCOUNT ON;'
        proc = no_count + query
        # result = get_multiple_result(connection, proc)
        # return result

    @staticmethod
    def close_connection(connection):
        """
        Close the database connection
        :param connection: pyodbc connection object
        """
        try:
            if connection is not None and connection.connected:
                connection.close()
        except Exception as e:
            logging.error(f"Error while closing the connection: {e}")
            
            
