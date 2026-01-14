# 📧 Extrator de Leads de Email via IMAP

Este é um script em Python projetado para automatizar a captura de leads recebidos por email. Ele se conecta a uma conta de email (atualmente configurado para Gmail), busca por mensagens não lidas com um assunto específico, extrai informações do corpo do email usando expressões regulares (Regex) e salva os dados de forma estruturada em um arquivo de texto.

## 🚀 Funcionalidades

- **Conexão Segura:** Utiliza `IMAP4_SSL` para uma conexão segura com o servidor de email.
- **Busca Inteligente:** Procura por emails **não lidos** que correspondam a um assunto específico, garantindo que apenas novos leads sejam processados.
- **Extração de Dados:** Extrai as seguintes informações do corpo do email:
  - Nome (`Name`)
  - Data de Entrada (`Data Entrada`)
  - Hora de Entrada (`Hora Entrada`)
  - Email (`Email`)
  - Origem da Mídia (`Publico`, via `utm_medium`)
  - Campanha/Anúncio (`Anuncio`, via `utm_content`)
- **Armazenamento:** Salva os leads capturados no arquivo `leads_extraidos.txt`, adicionando novos leads ao final do arquivo.
- **Gerenciamento de Segredos:** Utiliza um arquivo `.env` para carregar a senha do email, evitando que informações sensíveis sejam expostas no código.

## 🛠️ Como Usar

Siga os passos abaixo para configurar e executar o script.

### 1. Pré-requisitos

- Python 3.x instalado.
- Uma conta de email para a qual você tenha acesso via IMAP (o exemplo usa Gmail).

**Atenção (Para usuários Gmail):**
Para que o script funcione com o Gmail, talvez seja necessário gerar uma **"Senha de App"**. O Google não permite mais o uso da senha principal em aplicações de terceiros por padrão.
- Acesse as configurações da sua Conta Google.
- Vá para "Segurança".
- Ative a "Verificação em duas etapas".
- Após ativar, a opção "Senhas de app" aparecerá. Gere uma nova senha para este aplicativo e use-a no lugar da sua senha normal.

### 2. Clone o Repositório

```bash
git clone <URL-DO-SEU-REPOSITORIO-NO-GITHUB>
cd <NOME-DA-PASTA>
```

### 3. Instale as Dependências

Este projeto depende da biblioteca `python-dotenv`. Instale-a usando o `pip`:

```bash
pip install -r requirements.txt
```

### 4. Configure as Variáveis de Ambiente

Crie um arquivo chamado `.env` na raiz do projeto. Este arquivo guardará sua senha de email.

```
EMAIL_PASS="SUA_SENHA_DE_APP_AQUI"
```
Substitua `SUA_SENHA_DE_APP_AQUI` pela senha de app que você gerou (ou sua senha normal, se o provedor permitir).

### 5. Configure o Script

Abra o arquivo `extrairEmail.py` e, se necessário, ajuste as seguintes variáveis no topo do arquivo:

- `EMAIL_USER`: Seu endereço de email.
- `IMAP_SERVER`: O servidor IMAP do seu provedor (ex: `imap.gmail.com`).
- A linha `mail.search(None,'UNSEEN SUBJECT' , '"Lead - Cobertura Concept"')`: Altere `"Lead - Cobertura Concept"` para o assunto exato dos emails de lead que você deseja processar.

### 6. Execute o Script

Finalmente, execute o script a partir do seu terminal:

```bash
python extrairEmail.py
```

O script irá se conectar, buscar por novos leads e, se encontrar algum, salvará os dados no arquivo `excel` na mesma pasta.

```
