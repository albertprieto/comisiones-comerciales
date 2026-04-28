#!/usr/bin/env python3
"""Lee la hoja Drive 'Pagos Comisiones - Industrial Shields' (publicada como CSV)
y la guarda como pagos_registrados.csv en formato esperado por build_html.py.

Variable de entorno requerida:
  PAGOS_CSV_URL  URL pública de la hoja en formato CSV.
                 Generada vía Archivo → Compartir → Publicar en la web → CSV.
"""
import os, csv, urllib.request, io

URL = os.environ.get('PAGOS_CSV_URL', '').strip()
HERE = os.path.dirname(__file__)
OUT  = os.path.join(HERE, 'pagos_registrados.csv')

if not URL:
    print("[fetch_pagos] PAGOS_CSV_URL no configurada — uso CSV local si existe.")
    if os.path.exists(OUT):
        print(f"[fetch_pagos] CSV local presente: {OUT}")
    else:
        # Crea CSV vacío con cabecera para que build_html.py no falle
        with open(OUT, 'w', encoding='utf-8-sig', newline='') as f:
            f.write('periodo_cobro;comercial;importe_pagado_eur;fecha_pago;notas\n')
        print(f"[fetch_pagos] Creado CSV vacío: {OUT}")
    raise SystemExit(0)

print(f"[fetch_pagos] Descargando {URL}")
req = urllib.request.Request(URL, headers={'User-Agent': 'Mozilla/5.0'})
with urllib.request.urlopen(req, timeout=30) as resp:
    raw = resp.read().decode('utf-8-sig')

# Detectar separador automaticamente
first_line = raw.split('\n', 1)[0]
sep = ',' if first_line.count(',') >= first_line.count(';') else ';'

reader = csv.DictReader(io.StringIO(raw), delimiter=sep)
rows = list(reader)
print(f"[fetch_pagos] {len(rows)} filas leídas (separador: '{sep}')")

# Reescribir con separador ';' que es el formato esperado por build_html.py
fields = ['periodo_cobro', 'comercial', 'importe_pagado_eur', 'fecha_pago', 'notas']
with open(OUT, 'w', encoding='utf-8-sig', newline='') as f:
    w = csv.DictWriter(f, fieldnames=fields, delimiter=';')
    w.writeheader()
    for r in rows:
        out = {k: (r.get(k) or '').strip() for k in fields}
        # Normalizar importe (puede venir con coma decimal)
        try:
            imp = float((out['importe_pagado_eur'] or '0').replace(',', '.'))
            out['importe_pagado_eur'] = f"{imp:.2f}"
        except Exception:
            pass
        w.writerow(out)

print(f"[fetch_pagos] Guardado: {OUT}")
