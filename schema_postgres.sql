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

CREATE INDEX IF NOT EXISTS idx_instrutores_matricula ON instrutores(matricula);
CREATE INDEX IF NOT EXISTS idx_semanas_token ON semanas(token);
CREATE INDEX IF NOT EXISTS idx_horarios_semana ON horarios(semana_id);
CREATE INDEX IF NOT EXISTS idx_escolhas_horario ON escolhas(horario_id);
CREATE INDEX IF NOT EXISTS idx_escolhas_instrutor ON escolhas(instrutor_id);
