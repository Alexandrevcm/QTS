# Guia de publicação: Streamlit Cloud + Supabase

## 1. Criar banco no Supabase

1. Acesse sua conta no Supabase.
2. Crie um novo projeto.
3. Guarde a senha do banco.
4. Vá em **Project Settings > Database**.
5. Copie a connection string PostgreSQL.
6. Prefira a string do pooler quando disponível.

Exemplo do formato:

```text
postgresql://postgres.xxxxx:SUA_SENHA@aws-0-sa-east-1.pooler.supabase.com:6543/postgres
```

## 2. Preparar GitHub

1. Crie um repositório no GitHub.
2. Envie estes arquivos para o repositório:
   - `app.py`
   - `requirements.txt`
   - `runtime.txt`
   - `.streamlit/config.toml`
   - `README.md`
3. Não envie `.streamlit/secrets.toml` com dados reais.

## 3. Publicar no Streamlit Cloud

1. Acesse o Streamlit Community Cloud.
2. Clique em **Create app**.
3. Escolha o repositório, branch e arquivo `app.py`.
4. Antes ou depois do primeiro deploy, abra **App settings > Secrets**.
5. Cole os secrets:

```toml
ADMIN_PASSWORD = "troque-esta-senha"
BASE_URL = "https://seu-app.streamlit.app"
DATABASE_MODE = "postgres"
DATABASE_URL = "postgresql://postgres.xxxxx:SUA_SENHA@aws-0-sa-east-1.pooler.supabase.com:6543/postgres"
POSTGRES_SSLMODE = "require"
```

## 4. Primeiro acesso

1. Abra o link público do Streamlit.
2. Entre com a senha administrativa.
3. Acesse **7. Ambiente/Nuvem**.
4. Confirme se aparece **PostgreSQL/Supabase compartilhado**.
5. Crie uma semana.
6. Gere a grade padrão.
7. Copie o link público na aba **6. Link público**.

## 5. Observações importantes

- O SQLite local serve apenas para testes.
- O PostgreSQL/Supabase é o banco compartilhado para todos os instrutores.
- Se `BASE_URL` estiver errado, o link público gerado ficará errado.
- Troque `ADMIN_PASSWORD` antes de usar em produção.
