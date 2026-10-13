# Score Prestador Odontológico

Sistema que recebe um Excel com CNPJs, consulta dados da Receita Federal, busca o CNES no MongoDB, enriquece com Google Places (horários, avaliações, transporte) e analisa comentários/fotos com Claude.

## Arquivos

- `app.py` — aplicação Flask
- `requirements.txt` — dependências
- `.gitignore`

## Rodar localmente

```bash
pip install -r requirements.txt
python app.py
```

Acesse http://127.0.0.1:5000

## Deploy no Render (recomendado, gratuito)

1. Faça push deste código para o seu repositório no GitHub.
2. Entre em https://render.com e crie uma conta (pode usar a conta do GitHub).
3. Clique em **New** → **Web Service**.
4. Conecte o repositório.
5. Configure:
   - **Runtime**: Python
   - **Build Command**: `pip install -r requirements.txt`
   - **Start Command**: `gunicorn app:app`
6. Em **Environment** adicione as variáveis (recomendado, mesmo que o código tenha fallback):
   - `MONGO_URI` = sua connection string do Mongo
   - `GOOGLE_API_KEY` = sua chave do Google
   - `CLAUDE_API_KEY` = sua chave da Anthropic
7. Clique em **Create Web Service**.

Depois de alguns minutos o site estará no ar com uma URL do tipo `https://seu-app.onrender.com`.

## Observações

- O processamento de muitos CNPJs pode demorar e consumir cota das APIs (Google e Claude).
- No plano gratuito do Render o serviço “dorme” depois de inatividade e demora alguns segundos para acordar.
- Nunca deixe as chaves expostas em repositório público se não quiser. Use as variáveis de ambiente.
