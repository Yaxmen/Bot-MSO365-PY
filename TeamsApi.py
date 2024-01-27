import requests
import json
import msal

# Configurações do aplicativo no Azure AD
client_id = "5b44f698-ad3f-4403-b569-3e71fa2a0613"
tenant_id = "6af8f826-d4c2-47de-9f6d-c04908aa4e88"
client_secret = "cf577455-5f66-4a08-9c96-70c4498adb41"

# ID da equipe e do canal
team_id = "6952c254-8b4d-4b63-9e5c-20ef74697638"
channel_id = "19%3a5221297621be4c8496179fcfd02a2bb3%40thread.tacv2/02.%2520Torre%2520Automa%25C3%25A7%25C3%25A3o"

# URL para obter um token de acesso
token_url = f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token"

# Configuração do escopo e API do Microsoft Graph
scope = ["https://graph.microsoft.com/.default"]
graph_api_url = f"https://graph.microsoft.com/v1.0/teams/{team_id}/channels/{channel_id}/messages"

# Dados da mensagem a ser enviada
message_data = {
    "body": {
        "content": "Sua mensagem de notificação aqui."
    }
}

# Autenticação usando MSAL
app = msal.ConfidentialClientApplication(
    client_id,
    authority=f"https://login.microsoftonline.com/{tenant_id}",
    client_credential=client_secret
)

# Solicitar um token de acesso
token_response = app.acquire_token_for_client(scope)

# Verificar se o token foi obtido com sucesso
if "access_token" in token_response:
    access_token = token_response["access_token"]

    # Enviar a mensagem para o canal no Teams
    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/json"
    }

    response = requests.post(graph_api_url, headers=headers, data=json.dumps(message_data))

    if response.status_code == 201:
        print("Mensagem enviada com sucesso para o Teams!")
    else:
        print(f"Erro ao enviar mensagem para o Teams: {response.status_code} - {response.text}")
else:
    print(f"Erro ao obter token de acesso: {token_response.get('error')} - {token_response.get('error_description')}")