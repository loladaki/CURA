-- ══════════════════════════════════════════════════════════════
-- VELAS DASHBOARD — Migração v2
-- Corre este SQL no Supabase SQL Editor para adicionar:
--   • Campos específicos por tipo no inventário (melt_temp, flash_point, volume, density)
--   • Estado "a caminho" / "em stock"
--   • Custo e PVP nas receitas
--   • FKs receita → materiais do inventário
-- ══════════════════════════════════════════════════════════════

-- INVENTÁRIO: campos por tipo + estado
ALTER TABLE inventario
  ADD COLUMN IF NOT EXISTS melt_temp    numeric DEFAULT 0,
  ADD COLUMN IF NOT EXISTS flash_point  numeric DEFAULT 0,
  ADD COLUMN IF NOT EXISTS volume       numeric DEFAULT 0,
  ADD COLUMN IF NOT EXISTS density      numeric DEFAULT 0,
  ADD COLUMN IF NOT EXISTS status       text    DEFAULT 'em_stock';

-- RECEITAS: custo, PVP, e FKs para materiais
ALTER TABLE receitas
  ADD COLUMN IF NOT EXISTS cost          numeric DEFAULT 0,
  ADD COLUMN IF NOT EXISTS pvp           numeric DEFAULT 0,
  ADD COLUMN IF NOT EXISTS cera_id       text REFERENCES inventario(id) ON DELETE SET NULL,
  ADD COLUMN IF NOT EXISTS essencia_id   text REFERENCES inventario(id) ON DELETE SET NULL,
  ADD COLUMN IF NOT EXISTS recipiente_id text REFERENCES inventario(id) ON DELETE SET NULL,
  ADD COLUMN IF NOT EXISTS pavio_id      text REFERENCES inventario(id) ON DELETE SET NULL;

-- Índices para os novos FKs (queries mais rápidas)
CREATE INDEX IF NOT EXISTS idx_receitas_cera       ON receitas(cera_id);
CREATE INDEX IF NOT EXISTS idx_receitas_essencia   ON receitas(essencia_id);
CREATE INDEX IF NOT EXISTS idx_receitas_recipiente ON receitas(recipiente_id);
CREATE INDEX IF NOT EXISTS idx_receitas_pavio      ON receitas(pavio_id);
CREATE INDEX IF NOT EXISTS idx_inventario_status   ON inventario(status);
