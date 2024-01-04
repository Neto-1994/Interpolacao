import mysql.connector
from decouple import config
try:
    # Função para criar ou reutilizar uma conexão
    def obter_conexao():
        # Se a conexão não existir, crie uma nova conexão
        if 'conexao' not in globals():
            host = config('HOST')
            port = config('PORT_DB')
            database = config('DATABASE')
            user = config('USER')
            password = config('PASSWORD')
            globals()['conexao'] = mysql.connector.connect(
                host=host, port=port, database=database, user=user, password=password)

            '''db_info = globals()['conexao'].get_server_info()
            print("Conectado ao servidor MySQL versão: ", db_info)
            cursor = globals()['conexao'].cursor()
            cursor.execute("select database();")
            banco = cursor.fetchone()
            print("Conectado ao banco de dados: ", banco)'''

        elif not globals()['conexao'].is_connected():
            # Se a conexão existente não estiver mais ativa, reconecte
            globals()['conexao'] = mysql.connector.connect(
                host=host, port=port, database=database, user=user, password=password)

        return globals()['conexao']

except OSError as e:
    print("Erro ao acessar banco MySQL: ", e)
