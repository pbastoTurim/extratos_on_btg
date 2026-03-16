from datetime import datetime


def get_actual_date() -> str:
    """
    Método que retorna a data atual no formato 'dd/mm/yyyy'.
    :return: string com a data atual.
    """
    return datetime.now().strftime('%d/%m/%Y')