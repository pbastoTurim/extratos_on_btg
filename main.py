import os
from time import sleep
from dotenv import load_dotenv
from bot.conciliacao_btg import baixar_posicoes
from utils.dates import get_reference_date
from modules.boxes_input import app_option

# Load environment variables

load_dotenv()

cpf = os.getenv('LOGIN')
senha = os.getenv('SENHA')

opcao = app_option()
if not opcao:
    print("Nenhuma opção selecionada. Encerrando")
    raise SystemExit(0)

days_ago = 0
actual_date = get_reference_date(days_ago)
month_year = actual_date.strftime("%m/%Y")

# Run the bot with retries
max_attempts = 3
attempt = 0
success = False

while attempt < max_attempts and not success:
    if opcao == "extratos":
        try:
            print("Opção escolhida: Extratos")
            baixar_posicoes(cpf, senha, month_year)
            success = True
        except Exception as e:
            attempt += 1
            print(f"\nTentativa {attempt} de {max_attempts} falhou: {e}")
            if attempt < max_attempts:
                print("Aguardando 10 segundos antes da próxima tentativa...")
                sleep(10)
            else:
                print("Todas as tentativas falharam. Por favor, verifique as credenciais e tente novamente.")
    elif opcao == "informe":
        try:
            print("Opção escolhida: Informes")
            # FUNÇOES AQUI
            success = True
        except Exception as e:
            attempt += 1
            print(f"\nTentativa {attempt} de {max_attempts} falhou: {e}")
            if attempt < max_attempts:
                print("Aguardando 10 segundos antes da próxima tentativa...")
                sleep(10)
            else:
                print("Todas as tentativas falharam. Por favor, verifique as credenciais e tente novamente.")