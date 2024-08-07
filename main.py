import pandas as pd
import asyncio
from unidecode import unidecode
from gestaoinfo import Gestao

# Passo 1: Carregar o arquivo Excel para listar as abas
xls = pd.ExcelFile('cadastros.xlsx')
abas = xls.sheet_names
print("Abas disponíveis:", abas)

# Função para extrair primeiro e último nome do motorista


def gerar_nome_usuario(nome_completo):
    nomes = unidecode(nome_completo).split()
    primeiro_nome = nomes[0]
    ultimo_nome = nomes[-1]
    return f"{primeiro_nome.lower()}.{ultimo_nome.lower()}"

# Função para gerar a senha


def gerar_senha(ultimo_nome, cpf):
    ultimo_nome_sem_acento = unidecode(ultimo_nome)
    return f"{ultimo_nome_sem_acento.lower()}{cpf[:3]}"

# Função principal para processar todas as abas


async def main():
    print("RODOU O MAIN")
    gestao = Gestao()
    await gestao.auth()

    # DataFrames para armazenar os dados de usuários criados e não criados
    all_data = {}
    failed_users = []

    for aba in abas:
        df = pd.read_excel(xls, sheet_name=aba)

        # Verificar se a coluna 'NOME COMPLETO' existe
        if 'NOME COMPLETO' not in df.columns:
            print(
                f"Coluna 'NOME COMPLETO' não encontrada na aba '{aba}'. Pulando aba.")
            continue

        df['usuario'] = df['NOME COMPLETO'].apply(
            lambda x: gerar_nome_usuario(x) if pd.notna(x) else '')
        df['senha'] = df.apply(lambda x: gerar_senha(x['NOME COMPLETO'].split()[-1], str(x['CPF']).replace(
            ".", "").replace("-", "")) if pd.notna(x['NOME COMPLETO']) and pd.notna(x['CPF']) else '', axis=1)

        # DataFrame para armazenar dados que não foram criados
        failed_users_df = df.copy()

        # Lista para armazenar usuários que foram criados com sucesso
        created_users = []

        # Iterar sobre as linhas do DataFrame e enviar cada usuário para a API
        for index, row in df.iterrows():
            user_data = {
                "Cpf": row['CPF'],
                "DisplayName": row['NOME COMPLETO'],
                "Email": "",
                "Empresas": [],
                "Password": row['senha'],
                "Source": "site",
                "Telefone": "",
                "TenantId": "4",
                "UserImage": "",
                "Username": row['usuario'],
                "Signature": "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAASwAAACWCAYAAABkW7XSAAAAAXNSR0IArs4c6QAABXNJREFUeF7t2sFq21AARNGnL0/75U1UUOkiIcTc2R1tgg2+NgdnkG1df875dRwEeoGPt9a5zjn33+e4b9/Hc9+rt8vG8xpeaT6P+Unjs8dUnZ+6vvK83z3HK45fNe/W/++h67oH6+Meo9X/wyoSIBAK/N0qgxWKShEgMBMwWDNaYQIEaoFnsN6uc37XcT0CBAiUAgar1NQiQGAqYLCmvOIECJQCvsMqNbUIEJgKGKwprzgBAqWAwSo1tQgQmAoYrCmvOAECpYDBKjW1CBCYCviVcMorToBAKWCwSk0tAgSmAj4STnnFCRAoBQxWqalFgMBUwGBNecUJECgFDFapqUWAwFTAYE15xQkQKAUMVqmpRYDAVMCvhFNecQIESgGDVWpqESAwFTBYU15xAgRKAd9hlZpaBAhMBQzWlFecAIFSwGCVmloECEwFDNaUV5wAgVLAYJWaWgQITAUM1pRXnACBUsBglZpaBAhMBQzWlFecAIFSwGCVmloECEwFDNaUV5wAgVLAYJWaWgQITAUM1pRXnACBUsBglZpaBAhMBQzWlFecAIFSwGCVmloECEwFDNaUV5wAgVLAYJWaWgQITAUM1pRXnACBUsBglZpaBAhMBQzWlFecAIFSwGCVmloECEwFDNaUV5wAgVLAYJWaWgQITAUM1pRXnACBUsBglZpaBAhMBQzWlFecAIFSwGCVmloECEwFDNaUV5wAgVLAYJWaWgQITAUM1pRXnACBUsBglZpaBAhMBQzWlFecAIFSwGCVmloECEwFDNaUV5wAgVLAYJWaWgQITAUM1pRXnACBUsBglZpaBAhMBQzWlFecAIFSwGCVmloECEwFDNaUV5wAgVLAYJWaWgQITAUM1pRXnACBUsBglZpaBAhMBQzWlFecAIFSwGCVmloECEwFDNaUV5wAgVLAYJWaWgQITAUM1pRXnACBUsBglZpaBAhMBQzWlFecAIFSwGCVmloECEwFDNaUV5wAgVLAYJWaWgQITAUM1pRXnACBUsBglZpaBAhMBQzWlFecAIFSwGCVmloECEwFDNaUV5wAgVLAYJWaWgQITAUM1pRXnACBUuDfYJVRLQIECKwE3gGdaC0fiZPcHwAAAABJRU5ErkJggg==",
            }
            try:
                resultado = await gestao.createUsers(user_data)
                if resultado.get('Error'):
                    failed_users.append(
                        {'Nome Completo': row['NOME COMPLETO'], 'CPF': row['CPF']})
                else:
                    created_users.append(
                        {'Nome Completo': row['NOME COMPLETO'], 'Usuario': row['usuario'], 'Senha': row['senha']})
            except Exception as e:
                print(f"Erro ao criar usuário: {e}")
                failed_users.append(
                    {'Nome Completo': row['NOME COMPLETO'], 'CPF': row['CPF']})

        # Adicionar os dados ao arquivo de sucesso
        all_data[aba] = pd.DataFrame(created_users)

    # Salvar todos os usuários criados em um novo arquivo Excel
    with pd.ExcelWriter('usuarios_criados.xlsx') as writer:
        for aba, data in all_data.items():
            data.to_excel(writer, sheet_name=aba, index=False)

    # Salvar os usuários não criados em um novo arquivo Excel
    if failed_users:
        failed_users_df = pd.DataFrame(failed_users)
        failed_users_df.to_excel('usuarios_falhados.xlsx', index=False)

    print("Arquivos gerados com sucesso!")

asyncio.run(main())
