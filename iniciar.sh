#!/bin/bash
cd "$(dirname "$0")"
echo ""
echo "══════════════════════════════════════════════════"
echo "  GENERADOR DE INFORMES — DISTRITO 17D08"
echo "══════════════════════════════════════════════════"
echo ""
pip3 install -r requirements.txt --quiet 2>/dev/null || pip install -r requirements.txt --quiet 2>/dev/null
echo "  Abriendo navegador..."
python3 app.py
