-- ══════════════════════════════════════════════════════════════
-- VELAS DASHBOARD — Supabase Setup
-- Corre este SQL no Supabase SQL Editor (https://supabase.com)
--
-- Ordem de criação (respeita dependências FK):
--   inventario → receitas → gastos → lotes → stock → vendas
-- ══════════════════════════════════════════════════════════════

-- ── 1. Inventário de Matérias-Primas ──────────────────────────
CREATE TABLE IF NOT EXISTS inventario (
  id          text PRIMARY KEY,
  name        text NOT NULL,
  type        text DEFAULT 'outro',
  unit        text DEFAULT 'un',
  qty         numeric DEFAULT 0,
  price       numeric DEFAULT 0,
  min         numeric DEFAULT 0,
  supplier    text DEFAULT '',
  notes       text DEFAULT '',
  created_at  timestamptz DEFAULT now()
);

-- ── 2. Receitas de Velas ───────────────────────────────────────
CREATE TABLE IF NOT EXISTS receitas (
  id          text PRIMARY KEY,
  name        text NOT NULL,
  wax         text DEFAULT '',
  weight      numeric DEFAULT 0,
  fragpct     numeric DEFAULT 0,
  temp        numeric DEFAULT 0,
  wick        text DEFAULT '',
  container   text DEFAULT '',
  cure        numeric DEFAULT 0,
  frags       text DEFAULT '',
  color       text DEFAULT '',
  rating      integer DEFAULT 0,
  notes       text DEFAULT '',
  date        date,
  created_at  timestamptz DEFAULT now()
);

-- ── 3. Gastos & Despesas ───────────────────────────────────────
CREATE TABLE IF NOT EXISTS gastos (
  id          text PRIMARY KEY,
  date        date,
  descricao   text DEFAULT '',
  cat         text DEFAULT 'outro',
  amount      numeric DEFAULT 0,
  vendor      text DEFAULT '',
  notes       text DEFAULT '',
  created_at  timestamptz DEFAULT now()
);

-- ── 4. Lotes da Calculadora ────────────────────────────────────
-- Um lote pode estar ligado a uma receita e a materiais do inventário
CREATE TABLE IF NOT EXISTS lotes (
  id          text PRIMARY KEY,
  name        text,
  date        date,
  qty         numeric,
  weight      numeric,
  fp          numeric,   -- % essência
  wp          numeric,   -- €/kg cera
  frp         numeric,   -- €/kg essência
  cont        numeric,   -- €/un recipiente
  wick        numeric,   -- €/un pavio
  lbl         numeric,   -- €/un etiqueta
  other       numeric,
  margin      numeric,
  wax_kg      numeric,
  frag_g      numeric,
  wax_cost    numeric,
  frag_cost   numeric,
  cont_cost   numeric,
  wick_cost   numeric,
  lbl_cost    numeric,
  total_cost  numeric,
  cpc         numeric,
  pvp         numeric,
  -- FK para a receita usada neste lote
  receita_id  text REFERENCES receitas(id) ON DELETE SET NULL,
  -- FK para o material de cera usado (opcional)
  cera_id     text REFERENCES inventario(id) ON DELETE SET NULL,
  -- FK para o material de essência usado (opcional)
  essencia_id text REFERENCES inventario(id) ON DELETE SET NULL,
  created_at  timestamptz DEFAULT now()
);

-- ── 5. Stock de Velas Acabadas ─────────────────────────────────
-- Uma vela em stock vem de um lote e segue uma receita
CREATE TABLE IF NOT EXISTS stock (
  id          text PRIMARY KEY,
  name        text NOT NULL,
  frag        text DEFAULT '',
  size        text DEFAULT '',
  qty         numeric DEFAULT 0,
  cost        numeric DEFAULT 0,
  pvp         numeric DEFAULT 0,
  notes       text DEFAULT '',
  -- FK para a receita desta vela
  receita_id  text REFERENCES receitas(id) ON DELETE SET NULL,
  -- FK para o lote de produção
  lote_id     text REFERENCES lotes(id) ON DELETE SET NULL,
  created_at  timestamptz DEFAULT now()
);

-- ── 6. Vendas & Encomendas ─────────────────────────────────────
-- Uma venda referencia uma vela do stock
CREATE TABLE IF NOT EXISTS vendas (
  id          text PRIMARY KEY,
  date        date,
  client      text DEFAULT '',
  product     text DEFAULT '',
  qty         numeric DEFAULT 1,
  price       numeric DEFAULT 0,
  status      text DEFAULT 'pendente',
  pay         text DEFAULT '',
  notes       text DEFAULT '',
  -- FK para a vela vendida (opcional — product fica como texto livre também)
  stock_id    text REFERENCES stock(id) ON DELETE SET NULL,
  created_at  timestamptz DEFAULT now()
);

-- ══════════════════════════════════════════════════════════════
-- Row Level Security
-- ══════════════════════════════════════════════════════════════
ALTER TABLE inventario ENABLE ROW LEVEL SECURITY;
ALTER TABLE receitas   ENABLE ROW LEVEL SECURITY;
ALTER TABLE gastos     ENABLE ROW LEVEL SECURITY;
ALTER TABLE lotes      ENABLE ROW LEVEL SECURITY;
ALTER TABLE stock      ENABLE ROW LEVEL SECURITY;
ALTER TABLE vendas     ENABLE ROW LEVEL SECURITY;

CREATE POLICY "anon_all" ON inventario FOR ALL TO anon USING (true) WITH CHECK (true);
CREATE POLICY "anon_all" ON receitas   FOR ALL TO anon USING (true) WITH CHECK (true);
CREATE POLICY "anon_all" ON gastos     FOR ALL TO anon USING (true) WITH CHECK (true);
CREATE POLICY "anon_all" ON lotes      FOR ALL TO anon USING (true) WITH CHECK (true);
CREATE POLICY "anon_all" ON stock      FOR ALL TO anon USING (true) WITH CHECK (true);
CREATE POLICY "anon_all" ON vendas     FOR ALL TO anon USING (true) WITH CHECK (true);

-- ══════════════════════════════════════════════════════════════
-- Índices para performance em queries frequentes
-- ══════════════════════════════════════════════════════════════
CREATE INDEX IF NOT EXISTS idx_lotes_receita    ON lotes(receita_id);
CREATE INDEX IF NOT EXISTS idx_stock_receita    ON stock(receita_id);
CREATE INDEX IF NOT EXISTS idx_stock_lote       ON stock(lote_id);
CREATE INDEX IF NOT EXISTS idx_vendas_stock     ON vendas(stock_id);
CREATE INDEX IF NOT EXISTS idx_vendas_status    ON vendas(status);
CREATE INDEX IF NOT EXISTS idx_vendas_date      ON vendas(date);
CREATE INDEX IF NOT EXISTS idx_gastos_date      ON gastos(date);
CREATE INDEX IF NOT EXISTS idx_gastos_cat       ON gastos(cat);
CREATE INDEX IF NOT EXISTS idx_inventario_type  ON inventario(type);

-- ══════════════════════════════════════════════════════════════
-- Para adicionar mais coisas no futuro, exemplos:
--
-- Adicionar coluna:
--   ALTER TABLE vendas ADD COLUMN tracking text DEFAULT '';
--
-- Nova tabela de clientes (ligada a vendas):
--   CREATE TABLE clientes (...);
--   ALTER TABLE vendas ADD COLUMN cliente_id text REFERENCES clientes(id);
--
-- Nova tabela de fornecedores (ligada a inventario):
--   CREATE TABLE fornecedores (...);
--   ALTER TABLE inventario ADD COLUMN fornecedor_id text REFERENCES fornecedores(id);
-- ══════════════════════════════════════════════════════════════
