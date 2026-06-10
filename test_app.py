"""
Tests del Generador de Informes de Nómina.

Verifican que tras externalizar PREV a novedades_anteriores.json:
  1. La app importa y las rutas siguen registradas.
  2. /novedades_plantilla sigue devolviendo la estructura (sin datos personales).
  3. /previous_data/<id> devuelve los datos del JSON (atajo conectado).
  4. _load_prev() carga las 27 personas del JSON real.
  5. app.py NO contiene cédulas hardcodeadas.
  6. Editar el JSON se refleja en /previous_data (recarga en caliente).
  7. Sin el JSON, /previous_data degrada a {'rows':[]} sin romperse.
"""
import json, re, shutil
from pathlib import Path
import app as appmod

HERE = Path(__file__).parent
PREV = HERE / 'novedades_anteriores.json'
CED = re.compile(r"['\"][0-9]{10}['\"]")

res = []
def check(nombre, cond, det=""):
    res.append((cond, f"[{'OK  ' if cond else 'FALLA'}] {nombre}" + (f" — {det}" if det and not cond else "")))

client = appmod.app.test_client()

# 1) rutas registradas
rutas = {r.rule for r in appmod.app.url_map.iter_rules()}
for ruta in ['/', '/upload', '/analyze', '/novedades_plantilla', '/previous_data/<nid>', '/generate']:
    check(f"ruta {ruta} registrada", ruta in rutas)

# 2) /novedades_plantilla = estructura, sin cédulas
r = client.get('/novedades_plantilla')
check("/novedades_plantilla responde 200", r.status_code == 200, f"status={r.status_code}")
nov = r.get_json()
check("/novedades_plantilla es lista de plantillas", isinstance(nov, list) and len(nov) > 0)
check("estructura sin cédulas", len(CED.findall(json.dumps(nov))) == 0)

# 3) _load_prev carga el JSON real (27 personas)
prev = appmod._load_prev()
total = sum(len(v.get('rows', [])) for v in prev.values())
check("_load_prev carga 27 personas del JSON", total == 27, f"got {total}")

# 4) /previous_data/<id> devuelve datos del JSON
r2 = client.get('/previous_data/encargos')
check("/previous_data/encargos responde 200", r2.status_code == 200)
data = r2.get_json()
check("encargos trae 4 filas", len(data.get('rows', [])) == 4, f"got {len(data.get('rows',[]))}")
r3 = client.get('/previous_data/jubilados_pension')
check("jubilados_pension trae 14 filas", len(r3.get_json().get('rows', [])) == 14)

# 5) app.py sin cédulas hardcodeadas
appsrc = (HERE / 'app.py').read_text(encoding='utf-8')
check("app.py sin cédulas hardcodeadas", len(CED.findall(appsrc)) == 0,
      f"encontradas {len(CED.findall(appsrc))}")

# 6) edición del JSON se refleja
backup = PREV.read_text(encoding='utf-8')
try:
    PREV.write_text(json.dumps({"encargos": {"rows": [["9", "9999999999", "PRUEBA TEST", "X", "Y"]]}},
                               ensure_ascii=False), encoding='utf-8')
    d = client.get('/previous_data/encargos').get_json()
    check("nueva fila editada aparece", d['rows'][0][2] == 'PRUEBA TEST')
    check("dato viejo ya no aparece", 'GOMEZ PRADO' not in json.dumps(d))
finally:
    PREV.write_text(backup, encoding='utf-8')

# 7) sin el JSON, degrada sin romperse
tmp = HERE / 'novedades_anteriores.json.bak'
shutil.move(str(PREV), str(tmp))
try:
    r4 = client.get('/previous_data/encargos')
    check("/previous_data responde 200 SIN el JSON", r4.status_code == 200)
    check("degrada a rows vacío sin el JSON", r4.get_json() == {'rows': []})
finally:
    shutil.move(str(tmp), str(PREV))

print("\n" + "=" * 60 + "\n  RESULTADOS\n" + "=" * 60)
for _, l in res: print(" ", l)
fallos = [l for ok, l in res if not ok]
print("=" * 60 + f"\n  {len(res)-len(fallos)}/{len(res)} OK\n" + "=" * 60)
raise SystemExit(1 if fallos else 0)
