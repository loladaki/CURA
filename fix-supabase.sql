-- ══════════════════════════════════════════════════════════════
-- VELAS DASHBOARD — Fix de Permissões
-- Corre este SQL no Supabase SQL Editor se vires o erro:
-- "❌ Sem acesso escrita" ou "permission denied for table"
-- ══════════════════════════════════════════════════════════════

-- Dá permissão total de leitura e escrita ao role anon
GRANT USAGE ON SCHEMA public TO anon, authenticated;

GRANT SELECT, INSERT, UPDATE, DELETE ON inventario TO anon, authenticated;
GRANT SELECT, INSERT, UPDATE, DELETE ON vendas     TO anon, authenticated;
GRANT SELECT, INSERT, UPDATE, DELETE ON stock      TO anon, authenticated;
GRANT SELECT, INSERT, UPDATE, DELETE ON gastos     TO anon, authenticated;
GRANT SELECT, INSERT, UPDATE, DELETE ON lotes      TO anon, authenticated;
GRANT SELECT, INSERT, UPDATE, DELETE ON receitas   TO anon, authenticated;

-- Para tabelas criadas no futuro também terem permissão automática
ALTER DEFAULT PRIVILEGES IN SCHEMA public
  GRANT SELECT, INSERT, UPDATE, DELETE ON TABLES TO anon, authenticated;
