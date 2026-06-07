# Quadro Semanal de Instrutores v1.3

Sistema web para criação de quadro semanal de instrutores.

## O que existe nesta versão

- Cadastro de semanas.
- Cadastro de instrutores.
- Geração automática da grade padrão de segunda a sexta.
- Horários vagos sem exigência de disciplina.
- Link público para instrutores escolherem horários.
- Bloqueio de vagas preenchidas.
- Bloqueio de conflito de horário para o mesmo instrutor.
- Limite de h/a por instrutor na semana.
- Inclusão e remoção manual pelo administrador.
- Quadro final com exportação Excel, PDF e texto para WhatsApp.
- Banco local SQLite para testes.
- Banco compartilhado PostgreSQL/Supabase para publicação online.
- Aba administrativa “Ambiente/Nuvem” para conferir o modo de banco.

## Rodar localmente

1. Instale as dependências:

```bash
pip install -r requirements.txt
```

2. Rode o app:

```bash
streamlit run app.py
```

Ou clique em `RODAR_APP.bat` no Windows.

Senha inicial do administrador:

```text
admin123
```

## Publicação online recomendada

A publicação foi preparada para o caminho:

```text
Streamlit Community Cloud + Supabase/PostgreSQL
```

O app continua funcionando localmente com SQLite, mas para enviar link aos instrutores pela internet é necessário usar banco compartilhado.

## Secrets para o Streamlit Cloud

No Streamlit Cloud, configure em **App settings > Secrets**:

```toml
ADMIN_PASSWORD = "troque-esta-senha"
BASE_URL = "https://seu-app.streamlit.app"
DATABASE_MODE = "postgres"
DATABASE_URL = "postgresql://postgres.xxxxx:SUA_SENHA@aws-0-sa-east-1.pooler.supabase.com:6543/postgres"
POSTGRES_SSLMODE = "require"
```

Não coloque sua senha real do banco dentro do código.

## Banco de dados

O app cria automaticamente as tabelas necessárias no primeiro acesso.

Tabelas criadas:

- `instrutores`
- `semanas`
- `horarios`
- `escolhas`

## Link para instrutores

Depois de publicar e configurar `BASE_URL`, entre na aba:

```text
6. Link público
```

Copie o link gerado e envie aos instrutores.

## Migração de dados locais para a nuvem

Se você já tiver dados no SQLite local e quiser enviar para o PostgreSQL/Supabase, use:

```bash
python migrar_sqlite_para_postgres.py --sqlite data/quadro_instrutores.db --database-url "SUA_DATABASE_URL"
```

Faça isso somente após criar o banco no Supabase.
