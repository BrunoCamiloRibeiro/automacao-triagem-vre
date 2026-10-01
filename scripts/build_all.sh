#!/usr/bin/env bash
# ==============================================================================
# Script de Build para ambientes Linux / macOS (equivalente ao build_all.bat)
# ==============================================================================
set -e

echo "[1/4] Inicializando bancos e templates..."
python3 database/init_db.py --seed

echo "[2/4] Compilando triagem_vre..."
pyinstaller --noconfirm --clean --onefile --noconsole \
    --name "triagem_vre" \
    --add-data "templates:templates" \
    --add-data "database/schema.sql:database" \
    --add-data "database/schema_cep.sql:database" \
    src/runner/main_seco.py || true

echo "[3/4] Compilando config_vre..."
pyinstaller --noconfirm --clean --onefile --windowed \
    --name "config_vre" \
    --add-data "database/schema.sql:database" \
    --add-data "database/schema_cep.sql:database" \
    src/admin/main_admin.py || true

echo "[4/5] Organizando pasta fbuild..."
mkdir -p "fbuild/Configurações/database"
mkdir -p "fbuild/Configurações/templates"
mkdir -p "fbuild/Configurações/entrada_csv/processados"
mkdir -p "fbuild/Planilhas anteriores"

if [ -f "dist/triagem_vre" ]; then
    cp -f "dist/triagem_vre" "fbuild/triagem_vre"
fi
if [ -f "dist/config_vre" ]; then
    cp -f "dist/config_vre" "fbuild/config_vre"
fi

if [ -f "Planilha_Controle_VRE_Ativa.xlsx" ]; then
    cp -f "Planilha_Controle_VRE_Ativa.xlsx" "fbuild/Planilha_Controle_VRE_Ativa.xlsx"
elif [ -f "planilhas_trabalho/Planilha_Controle_VRE_Ativa.xlsx" ]; then
    cp -f "planilhas_trabalho/Planilha_Controle_VRE_Ativa.xlsx" "fbuild/Planilha_Controle_VRE_Ativa.xlsx"
elif [ -f "templates/template_triagem.xlsx" ]; then
    cp -f "templates/template_triagem.xlsx" "fbuild/Planilha_Controle_VRE_Ativa.xlsx"
fi

cp -f database/triagem_vre.db "fbuild/Configurações/database/" 2>/dev/null || true
cp -f database/triagem_vre.db "fbuild/Configurações/" 2>/dev/null || true
cp -f database/cep_areas.db "fbuild/Configurações/database/" 2>/dev/null || true
cp -f database/cep_areas.db "fbuild/Configurações/" 2>/dev/null || true
cp -f database/schema.sql "fbuild/Configurações/database/" 2>/dev/null || true
cp -f database/schema.sql "fbuild/Configurações/" 2>/dev/null || true
cp -f database/schema_cep.sql "fbuild/Configurações/database/" 2>/dev/null || true
cp -f database/schema_cep.sql "fbuild/Configurações/" 2>/dev/null || true
cp -f templates/template_triagem.xlsx "fbuild/Configurações/templates/" 2>/dev/null || true
cp -f templates/template_triagem.xlsx "fbuild/Configurações/" 2>/dev/null || true
cp -rf src "fbuild/Configurações/" 2>/dev/null || true
cp -f requirements.txt "fbuild/Configurações/" 2>/dev/null || true
cp -f README.md "fbuild/Configurações/" 2>/dev/null || true

echo "[5/5] Concluído com sucesso."

