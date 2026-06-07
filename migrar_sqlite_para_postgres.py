# -*- coding: utf-8 -*-
"""
Migra dados do SQLite local para PostgreSQL/Supabase.

Uso:
    python migrar_sqlite_para_postgres.py --sqlite data/quadro_instrutores.db --database-url "postgresql://..."

Atenção:
- Use em um banco PostgreSQL vazio ou em ambiente de teste.
- O script preserva os IDs originais para manter os vínculos entre semanas, horários e escolhas.
"""

from __future__ import annotations

import argparse
import sqlite3
from pathlib import Path

import psycopg2
from psycopg2.extras import RealDictCursor


SCHEMA = """
CREATE TABLE IF NOT EXISTS instrutores (
    id SERIAL PRIMARY KEY,
    nome TEXT NOT NULL,
    posto_grad TEXT,
    matricula TEXT UNIQUE,
    telefone TEXT,
    areas TEXT,
    codigo_acesso TEXT,
    ativo INTEGER NOT NULL DEFAULT 1,
    criado_em TEXT NOT NULL DEFAULT to_char(now(), 'YYYY-MM-DD"T"HH24:MI:SS')
);

CREATE TABLE IF NOT EXISTS semanas (
    id SERIAL PRIMARY KEY,
    token TEXT NOT NULL UNIQUE,
    titulo TEXT NOT NULL,
    data_inicio TEXT NOT NULL,
    data_fim TEXT NOT NULL,
    prazo_escolha TEXT,
    status TEXT NOT NULL DEFAULT 'Aberta',
    observacoes TEXT,
    limite_horas_instrutor INTEGER NOT NULL DEFAULT 10,
    encerrada_em TEXT,
    publicada_em TEXT,
    criado_em TEXT NOT NULL DEFAULT to_char(now(), 'YYYY-MM-DD"T"HH24:MI:SS')
);

CREATE TABLE IF NOT EXISTS horarios (
    id SERIAL PRIMARY KEY,
    semana_id INTEGER NOT NULL REFERENCES semanas(id) ON DELETE CASCADE,
    data_aula TEXT NOT NULL,
    hora_inicio TEXT NOT NULL,
    hora_fim TEXT NOT NULL,
    disciplina TEXT NOT NULL DEFAULT '',
    local TEXT,
    vagas INTEGER NOT NULL DEFAULT 1,
    habilitacao TEXT,
    carga_horaria INTEGER NOT NULL DEFAULT 0,
    bloqueado INTEGER NOT NULL DEFAULT 0,
    observacoes TEXT,
    criado_em TEXT NOT NULL DEFAULT to_char(now(), 'YYYY-MM-DD"T"HH24:MI:SS')
);

CREATE TABLE IF NOT EXISTS escolhas (
    id SERIAL PRIMARY KEY,
    horario_id INTEGER NOT NULL REFERENCES horarios(id) ON DELETE CASCADE,
    instrutor_id INTEGER NOT NULL REFERENCES instrutores(id) ON DELETE CASCADE,
    status TEXT NOT NULL DEFAULT 'Confirmada',
    criado_em TEXT NOT NULL DEFAULT to_char(now(), 'YYYY-MM-DD"T"HH24:MI:SS'),
    UNIQUE (horario_id, instrutor_id)
);
"""


def rows(conn: sqlite3.Connection, tabela: str) -> list[sqlite3.Row]:
    return conn.execute(f"SELECT * FROM {tabela} ORDER BY id").fetchall()


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--sqlite", default="data/quadro_instrutores.db", help="Caminho do banco SQLite local")
    parser.add_argument("--database-url", required=True, help="URL do PostgreSQL/Supabase")
    parser.add_argument("--sslmode", default="require")
    args = parser.parse_args()

    sqlite_path = Path(args.sqlite)
    if not sqlite_path.exists():
        raise FileNotFoundError(f"SQLite não encontrado: {sqlite_path}")

    sqlite_conn = sqlite3.connect(sqlite_path)
    sqlite_conn.row_factory = sqlite3.Row

    pg = psycopg2.connect(args.database_url, cursor_factory=RealDictCursor, sslmode=args.sslmode)
    try:
        cur = pg.cursor()
        cur.execute(SCHEMA)

        # Limpa destino respeitando FKs.
        cur.execute("TRUNCATE TABLE escolhas, horarios, semanas, instrutores RESTART IDENTITY CASCADE")

        for r in rows(sqlite_conn, "instrutores"):
            cur.execute(
                """
                INSERT INTO instrutores (id, nome, posto_grad, matricula, telefone, areas, codigo_acesso, ativo, criado_em)
                VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s)
                """,
                (r["id"], r["nome"], r["posto_grad"], r["matricula"], r["telefone"], r["areas"], r["codigo_acesso"], r["ativo"], r["criado_em"]),
            )

        for r in rows(sqlite_conn, "semanas"):
            cur.execute(
                """
                INSERT INTO semanas (id, token, titulo, data_inicio, data_fim, prazo_escolha, status, observacoes, limite_horas_instrutor, encerrada_em, publicada_em, criado_em)
                VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)
                """,
                (r["id"], r["token"], r["titulo"], r["data_inicio"], r["data_fim"], r["prazo_escolha"], r["status"], r["observacoes"], r["limite_horas_instrutor"], r["encerrada_em"], r["publicada_em"], r["criado_em"]),
            )

        for r in rows(sqlite_conn, "horarios"):
            cur.execute(
                """
                INSERT INTO horarios (id, semana_id, data_aula, hora_inicio, hora_fim, disciplina, local, vagas, habilitacao, carga_horaria, bloqueado, observacoes, criado_em)
                VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)
                """,
                (r["id"], r["semana_id"], r["data_aula"], r["hora_inicio"], r["hora_fim"], r["disciplina"], r["local"], r["vagas"], r["habilitacao"], r["carga_horaria"], r["bloqueado"], r["observacoes"], r["criado_em"]),
            )

        for r in rows(sqlite_conn, "escolhas"):
            cur.execute(
                """
                INSERT INTO escolhas (id, horario_id, instrutor_id, status, criado_em)
                VALUES (%s,%s,%s,%s,%s)
                """,
                (r["id"], r["horario_id"], r["instrutor_id"], r["status"], r["criado_em"]),
            )

        for tabela in ("instrutores", "semanas", "horarios", "escolhas"):
            cur.execute(
                f"SELECT setval(pg_get_serial_sequence('{tabela}', 'id'), COALESCE((SELECT MAX(id) FROM {tabela}), 1), true)"
            )

        pg.commit()
        print("Migração concluída com sucesso.")
    except Exception:
        pg.rollback()
        raise
    finally:
        sqlite_conn.close()
        pg.close()


if __name__ == "__main__":
    main()
