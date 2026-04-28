#!/usr/bin/env python3
"""Genera Comisiones_Comerciales_Industrial_Shields.html con dataset_base.json embebido.

El payload se incrusta gzip+base64 (pako.js en navegador) para soportar >20k filas
sin hinchar el HTML. Tabla con limite de renderizado configurable.

Dashboard auto-contenido con:
  - Vista Tabla (filtros por cabecera, dropdowns, numericos, export CSV)
  - Vista Dashboard (agrupaciones por customer_type/cliente/familia/producto/
                     pais/comercial, KPIs y graficos bar+donut con Chart.js)
  - Vista Condiciones (reglas de comision por año, solo lectura)
  - Cabecera con 'Mi vista' para filtrar todo por un comercial concreto.

Todos los importes se muestran en EUR (tras aplicar FX historico).
"""
import json, os, gzip, base64, csv, sys, re

HERE = os.path.dirname(os.path.abspath(__file__))

# Modo por-comercial: genera un HTML por cada comercial con SOLO sus datos.
PER_COMMERCIAL = '--per-commercial' in sys.argv

# =============================================================================
# Google OAuth — Sign-In con cuenta @industrialshields.com
# =============================================================================
# Para activar: pega aquí el Client ID generado en Google Cloud Console
# (APIs & Services → Credentials → OAuth 2.0 Client ID).
# Si está vacío, los HTMLs por-comercial se generan SIN gate OAuth (igual que ahora).
GOOGLE_OAUTH_CLIENT_ID = "910332428446-60brh8gdts65on3hbr5otntkcr9lhdbq.apps.googleusercontent.com"

# Mapeo Comercial → email autorizado (sacado de res.users en Odoo)
COMMERCIAL_EMAILS = {
    "Albert Prieto":             "apm@industrialshields.com",
    "Eloi Davila Lopez":         "edl@industrialshields.com",
    "Garima Arora":              "gar@industrialshields.com",
    "Gerard Montero Martínez":   "gmm@industrialshields.com",
    "Jordi Hernandez":           "jhs@industrialshields.com",
    "Josep Massó":               "jmp@industrialshields.com",
    "Ramon Boncompte":           "rbt@industrialshields.com",
}

# Comerciales con vista de GERENTE (acceso a TODOS los datos sin filtro)
MANAGER_COMMERCIALS = {"Albert Prieto"}

# =============================================================================
# GOOGLE DRIVE — Hoja "Pagos Comisiones - Industrial Shields" (apm@)
# =============================================================================
# Esta hoja es la fuente de verdad. Albert la edita en Drive (importes y fechas
# de pago). build_html.py lee el CSV local; el ID/URL se exponen en el HTML
# como link para abrirla con un click.
#
# Permisos: solo el owner (apm@industrialshields.com) y los writers ya
# concedidos en Drive. Para que SOLO Albert pueda editarla, hay que ajustar
# permisos manualmente desde Drive (no se modifica desde aqui por seguridad).
DRIVE_SHEET_ID  = "1urDlTjSZaxcapOiXWz4GUHz9Y4MJiBFT4yP6JRJy3sI"
DRIVE_SHEET_URL = f"https://docs.google.com/spreadsheets/d/{DRIVE_SHEET_ID}/edit"

with open(os.path.join(HERE, 'dataset_base.json'), encoding='utf-8') as f:
    data = json.load(f)

# Compactar: JSON minificado -> gzip -> base64 (se descomprime en JS con pako)
_json_bytes = json.dumps(data, ensure_ascii=False, separators=(',', ':')).encode('utf-8')
_gz_bytes = gzip.compress(_json_bytes, compresslevel=9)
_payload_b64 = base64.b64encode(_gz_bytes).decode('ascii')

# Tarifas historicas deducidas por (producto, año)
_tariffs_b64 = ''
_tp = os.path.join(HERE, 'tariff_history.json')
if os.path.exists(_tp):
    with open(_tp, encoding='utf-8') as f:
        _tariffs_json = f.read()
    _tariffs_b64 = base64.b64encode(
        gzip.compress(_tariffs_json.encode('utf-8'), compresslevel=9)
    ).decode('ascii')

# =============================================================================
# REGISTRO DE PAGOS DE COMISIONES
# =============================================================================
# Fuente de verdad: Google Sheet "Pagos Comisiones - Industrial Shields"
# (id arriba). Albert la edita en Drive; el HTML enlaza a la URL.
#
# Cache local: pagos_registrados.csv (mismo esquema que la hoja).
# Se puede refrescar con `python3 build_html.py --sync-drive` (lee desde
# Drive y sobreescribe el CSV local antes de generar el HTML).
#
# Esquema (5 columnas):
#   periodo_cobro;comercial;importe_pagado_eur;fecha_pago;notas
#
# El script:
#   - Lee el CSV existente (no toca filas previas).
#   - Añade nuevas filas para (periodo_cobro, comercial) que no estaban
#     y tengan importe COBRADO > 0.
#   - Embebe los datos en el HTML para mostrar Pagado/Pendiente.
PAGOS_CSV = os.path.join(HERE, 'pagos_registrados.csv')
PAGOS_FIELDS = ['periodo_cobro', 'comercial', 'importe_pagado_eur', 'fecha_pago', 'notas']

# Si se pasa --sync-drive, simplemente avisamos: la sincronizacion real
# (descarga del Sheet -> CSV) se hace via el MCP de Drive del agente.
if '--sync-drive' in sys.argv:
    print("[sync-drive] Para sincronizar pagos_registrados.csv con la hoja de Drive,")
    print(f"   pide al agente: 'sincroniza pagos desde Drive' (id={DRIVE_SHEET_ID}).")
    print("   El agente leera la hoja, sobreescribira el CSV local y volvera a")
    print("   ejecutar este script. Sin --sync-drive el build usa el CSV local.")

_EXCL_SP = {"Industrial Shields - Website","ADMIN","Alba Sánchez Honrado",
            "Sònia Gabarró","Albert Macià","Abel Codina","Luis Nunes",
            "Francesc Duarri","Susana Guerra","Joan F. Aubets - Industrial Shields"}
_SP_TYPE = {"Jordi Hernandez":1,"Garima Arora":1,"Eloi Davila Lopez":1,
            "Gerard Montero Martínez":1,"Josep Massó":2,"Ramon Boncompte":2,
            "Albert Prieto":2,"SalesPerson":2}

def _period_for_payment(date_str):
    if not date_str: return None
    try:
        yr = int(date_str[:4])
        mo = int(date_str[5:7] or '1')
    except Exception:
        return None
    if yr < 2025: return None
    if yr == 2025:
        if mo < 7: return None
        return f"2025-Q{3 if mo<=9 else 4}"
    return f"{yr}-{mo:02d}"

def _commission_line(r):
    if r.get('is_section') or r.get('state') != 'sale': return 0.0
    sp = r.get('salesperson')
    if not sp or sp in _EXCL_SP: return 0.0
    cat = r.get('product_category') or ''
    nm  = r.get('product_name') or ''
    code = r.get('product_code') or ''
    if 'Shipping' in cat or 'Shipping' in nm: return 0.0
    if 'Controllino' in cat: return 0.0
    sub = r.get('price_subtotal_eur') or 0.0
    if code.startswith('PHP-') or 'Projects' in cat:
        return sub * 0.03
    t = _SP_TYPE.get(sp)
    if not t: return 0.0
    d = max(0, min(r.get('discount_pct') or 0, 30))
    rate = max(0, (3.6 if t==1 else 3.1) - d*0.1)
    return sub * rate / 100.0

# Calcula cobrado por (periodo_cobro, comercial) — usado para detectar nuevos pares
_cobrado_by = {}
for r in data:
    com = _commission_line(r)
    if com <= 0: continue
    if r.get('payment_state_agg') != 'paid': continue
    period = _period_for_payment(r.get('last_invoice_date'))
    if not period: continue
    sp = r['salesperson']
    _cobrado_by[(period, sp)] = _cobrado_by.get((period, sp), 0) + com

# Lee CSV existente sin modificar filas
_pagos_rows = []
_pagos_keys = set()
if os.path.exists(PAGOS_CSV):
    try:
        with open(PAGOS_CSV, encoding='utf-8-sig', newline='') as fh:
            for row in csv.DictReader(fh, delimiter=';'):
                # Normalizar campos faltantes
                out = {k: (row.get(k) or '').strip() for k in PAGOS_FIELDS}
                _pagos_rows.append(out)
                _pagos_keys.add((out['periodo_cobro'], out['comercial']))
    except Exception as e:
        print(f"  AVISO: no se pudo leer pagos_registrados.csv: {e}")

# Añade nuevas filas (period, sp) que tienen cobrado > 0 y no están en el CSV
_new_count = 0
for (period, sp), com in sorted(_cobrado_by.items()):
    if (period, sp) in _pagos_keys: continue
    _pagos_rows.append({
        'periodo_cobro': period,
        'comercial': sp,
        'importe_pagado_eur': '0.00',
        'fecha_pago': '',
        'notas': '',
    })
    _new_count += 1

# Ordenar y guardar (preserva filas existentes intactas, solo añade)
_pagos_rows.sort(key=lambda r: (r['periodo_cobro'], r['comercial']))
with open(PAGOS_CSV, 'w', encoding='utf-8-sig', newline='') as fh:
    w = csv.DictWriter(fh, fieldnames=PAGOS_FIELDS, delimiter=';')
    w.writeheader()
    for row in _pagos_rows:
        w.writerow(row)

# Embebe en HTML como JSON gzip+b64
def _to_float(s):
    if not s: return 0.0
    s = (str(s).strip().replace(',', '.'))
    try: return float(s)
    except: return 0.0

_pagos_payload = {
    'drive_url': DRIVE_SHEET_URL,
    'drive_id':  DRIVE_SHEET_ID,
    'rows': [
        {
            'period': r['periodo_cobro'],
            'sp': r['comercial'],
            'pagado': _to_float(r.get('importe_pagado_eur')),
            'fecha': r.get('fecha_pago', ''),
            'notas': r.get('notas', ''),
        }
        for r in _pagos_rows
    ],
}
_pagos_b64 = base64.b64encode(
    gzip.compress(json.dumps(_pagos_payload, ensure_ascii=False).encode('utf-8'), compresslevel=9)
).decode('ascii')

HTML = r"""<!doctype html>
<html lang="es">
<head>
<meta charset="utf-8">
<title>Comisiones Comerciales · Industrial Shields</title>
<script src="https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.1/chart.umd.min.js"></script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/pako/2.1.0/pako.min.js"></script>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap" rel="stylesheet">
<style>
  /* ==== Paleta Industrial Shields (fondo claro, acento rojo corporativo) ==== */
  :root{
    --bg:#f5f7fa;          /* fondo general claro */
    --bg-alt:#eef1f5;      /* variación */
    --panel:#ffffff;       /* superficies (cards, tablas) */
    --ink:#1a2433;         /* texto principal */
    --muted:#6b7a8c;       /* texto secundario */
    --line:#e1e6ed;        /* bordes */
    --line-strong:#c9d1dc;
    --navy:#1a2f5c;        /* navy corporativo */
    --navy-dark:#0f1c3a;
    --accent:#e30613;      /* rojo industrial shields */
    --accent-dark:#b80510;
    --accent-soft:#fde8ea;
    --ok:#2e7d32;
    --ok-soft:#e6f4e7;
    --warn:#f57c00;
    --warn-soft:#fff4e5;
    --err:#d32f2f;
    --row-alt:#f8fafc;
  }
  *{box-sizing:border-box}
  html,body{height:100%}
  body{margin:0;
       font:14px/1.5 "Inter",-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,
       Helvetica,Arial,sans-serif;
       background:var(--bg);color:var(--ink);
       -webkit-font-smoothing:antialiased;-moz-osx-font-smoothing:grayscale;
       letter-spacing:-0.01em;}

  header{
    display:flex;align-items:center;gap:14px;padding:12px 20px;
    background:#ffffff;
    border-bottom:3px solid var(--accent);
    box-shadow:0 1px 3px rgba(26,36,51,0.04);
    position:sticky;top:0;z-index:20;
    flex-wrap:wrap; /* evita que items se salgan del viewport */
  }
  header h1{
    margin:0;font-size:15px;font-weight:700;color:var(--navy);
    letter-spacing:-0.02em;white-space:nowrap;display:flex;align-items:center;gap:10px;
  }
  header h1::before{
    content:"";display:inline-block;width:3px;height:20px;background:var(--accent);
    border-radius:2px;flex-shrink:0;
  }
  header .tag{background:var(--bg-alt);color:var(--muted);padding:3px 10px;
       border-radius:4px;font-size:11px;font-weight:500;border:1px solid var(--line);
       letter-spacing:0;white-space:nowrap;}

  .view-switch{display:flex;gap:0;border:1px solid var(--line);
               border-radius:6px;overflow:hidden;background:#fff;flex-shrink:0}
  .view-switch button{
    background:#ffffff;color:var(--muted);border:0;border-right:1px solid var(--line);
    padding:7px 14px;cursor:pointer;font-size:13px;font-weight:500;
    font-family:inherit;transition:background 0.15s,color 0.15s;white-space:nowrap;
  }
  .view-switch button:last-child{border-right:0}
  .view-switch button:hover{background:var(--bg-alt);color:var(--navy)}
  .view-switch button.active{
    background:var(--accent);color:#fff;font-weight:600;
  }

  .my-view-wrap{
    display:flex;align-items:center;gap:6px;color:var(--muted);font-size:12px;
    font-weight:500;white-space:nowrap;
  }
  .my-view-wrap select, .my-view-wrap input[type="checkbox"]{
    background:#fff;color:var(--ink);border:1px solid var(--line);
    padding:5px 8px;border-radius:4px;font-size:12px;min-width:130px;
    font-family:inherit;
  }
  .my-view-wrap input[type="checkbox"]{min-width:0;width:14px;height:14px}
  .my-view-wrap.active select{border-color:var(--accent);box-shadow:0 0 0 2px var(--accent-soft)}
  .my-view-wrap input[type="date"], .my-view-wrap input[type="month"]{
    background:#fff;color:var(--ink);border:1px solid var(--line);
    padding:4px 7px;border-radius:4px;font-size:12px;font-family:inherit;
    min-width:120px;
  }
  .my-view-wrap input[type="search"]{
    background:#fff;color:var(--ink);border:1px solid var(--line);
    padding:5px 8px;border-radius:4px;font-size:12px;font-family:inherit;
  }
  .my-view-wrap input[type="search"]:focus{outline:none;border-color:var(--accent);box-shadow:0 0 0 2px var(--accent-soft)}
  .my-view-wrap.active input[type="search"]{border-color:var(--accent);box-shadow:0 0 0 2px var(--accent-soft);background:var(--accent-soft)}
  /* Cells clickable que fijan filtro global */
  [data-click-filter]{cursor:pointer;border-bottom:1px dotted transparent;transition:all 0.1s}
  [data-click-filter]:hover{border-bottom-color:var(--accent);background:var(--accent-soft)}
  .my-view-wrap input[type="date"]:focus{outline:none;border-color:var(--accent);box-shadow:0 0 0 2px var(--accent-soft)}
  .range-wrap select#global-preset{min-width:155px}
  .range-wrap.active select, .range-wrap.active input[type="date"]{
    border-color:var(--accent);box-shadow:0 0 0 1px var(--accent-soft);
  }

  .stats{display:flex;gap:16px;color:var(--muted);font-size:12px;font-weight:500;
         flex-wrap:wrap;align-items:center}
  .stats b{color:var(--navy);font-weight:700}
  .header-sep{flex:1;min-width:0} /* empuja los siguientes items al final */
  .header-divider{width:1px;height:20px;background:var(--line);margin:0 2px}

  /* ===== Toolbar tabla ===== */
  .toolbar{
    display:flex;align-items:center;gap:10px;padding:12px 28px;
    background:#fff;border-bottom:1px solid var(--line);flex-wrap:wrap;
  }
  .toolbar input[type="search"]{
    background:#fff;color:var(--ink);border:1px solid var(--line);
    padding:7px 12px;border-radius:4px;width:280px;font-size:13px;
    font-family:inherit;transition:border-color 0.15s,box-shadow 0.15s;
  }
  .toolbar input[type="search"]:focus{outline:none;border-color:var(--accent);box-shadow:0 0 0 3px var(--accent-soft)}
  .toolbar select, .toolbar input[type="number"]{
    background:#fff;color:var(--ink);border:1px solid var(--line);
    padding:6px 10px;border-radius:4px;font-size:13px;font-family:inherit;
  }
  .toolbar button, .panel-btn{
    background:#fff;color:var(--navy);border:1px solid var(--line);
    padding:7px 14px;border-radius:4px;cursor:pointer;font-size:13px;font-weight:500;
    font-family:inherit;transition:all 0.15s;
  }
  .toolbar button:hover, .panel-btn:hover{background:var(--bg-alt);border-color:var(--line-strong)}
  .toolbar button.primary, .panel-btn.primary{
    background:var(--accent);color:#fff;border-color:var(--accent);
  }
  .toolbar button.primary:hover{background:var(--accent-dark);border-color:var(--accent-dark)}
  .toolbar .spacer{flex:1}

  /* ===== Botón Drive (verde Google Sheets) ===== */
  a.panel-btn.drive-btn{
    background:#0f9d58;color:#fff;border-color:#0b7a44;text-decoration:none;
    display:inline-flex;align-items:center;gap:6px;font-weight:600;
  }
  a.panel-btn.drive-btn:hover{background:#0b7a44;border-color:#0b7a44}
  .drive-ico{font-size:14px;line-height:1}
  button.panel-btn.drive-btn{
    background:#0f9d58;color:#fff;border-color:#0b7a44;
    display:inline-flex;align-items:center;gap:6px;font-weight:600;cursor:pointer;
  }
  button.panel-btn.drive-btn:hover{background:#0b7a44;border-color:#0b7a44}
  button.panel-btn.drive-btn.drive-sync{background:#1976d2;border-color:#1565c0}
  button.panel-btn.drive-btn.drive-sync:hover{background:#1565c0;border-color:#0d47a1}
  button.panel-btn.drive-btn.is-loading{opacity:0.7;cursor:wait}
  button.panel-btn.drive-btn.is-loading .drive-ico{animation:spin 1s linear infinite;display:inline-block}
  @keyframes spin{from{transform:rotate(0)}to{transform:rotate(360deg)}}
  /* Toast de notificación de sincronización */
  .pay-toast{
    position:fixed;bottom:24px;right:24px;z-index:9999;
    background:#1a2f5c;color:#fff;padding:14px 18px;border-radius:6px;
    box-shadow:0 8px 24px rgba(0,0,0,0.18);font-size:13px;max-width:420px;
    display:flex;align-items:flex-start;gap:10px;
    animation:toastIn 0.25s ease-out;
  }
  .pay-toast.ok    {background:#2e7d32}
  .pay-toast.warn  {background:#ed6c02}
  .pay-toast.err   {background:#c62828}
  .pay-toast .pt-msg{flex:1;line-height:1.4}
  .pay-toast .pt-x{cursor:pointer;background:transparent;border:0;color:#fff;
                   font-size:18px;line-height:1;padding:0 4px;opacity:0.7}
  .pay-toast .pt-x:hover{opacity:1}
  .pay-toast .pt-actions{margin-top:8px;display:flex;gap:8px;flex-wrap:wrap}
  .pay-toast .pt-btn{
    background:rgba(255,255,255,0.15);color:#fff;border:1px solid rgba(255,255,255,0.3);
    padding:5px 10px;border-radius:4px;cursor:pointer;font-size:12px;font-family:inherit;
  }
  .pay-toast .pt-btn:hover{background:rgba(255,255,255,0.25)}
  @keyframes toastIn{from{opacity:0;transform:translateY(8px)}to{opacity:1;transform:translateY(0)}}
  a.drive-link{color:#0f9d58;text-decoration:none;font-weight:600}
  a.drive-link:hover{text-decoration:underline}

  /* ===== Tabla ===== */
  .wrap{overflow:auto;max-height:calc(100vh - 110px);background:#fff}
  table.data{border-collapse:separate;border-spacing:0;width:100%;min-width:2700px;background:#fff}
  table.data thead th{
    position:sticky;top:0;background:#f8fafc;color:var(--navy);
    text-align:left;padding:10px 10px 6px 10px;border-bottom:1px solid var(--line);
    font-size:12px;font-weight:600;white-space:nowrap;letter-spacing:0;
  }
  table.data thead th.sub{top:36px;background:#f8fafc;padding:2px 8px 10px;border-bottom:2px solid var(--line-strong)}
  /* Fila de totales acumulados — sobria, mismo tamaño que las demás celdas */
  table.data thead tr.totals-row td{
    position:sticky;top:80px;z-index:3;
    background:#eef3f8;
    border-top:1px solid var(--line-strong);border-bottom:1px solid var(--line-strong);
    padding:7px 10px;font-size:13px;color:var(--navy);
    font-weight:700;
  }
  table.data thead tr.totals-row td.totals-label{
    background:#dde6f0;text-align:left;line-height:1.2;font-weight:700;
  }
  table.data thead tr.totals-row td.totals-val.num{
    text-align:right;font-variant-numeric:tabular-nums;font-feature-settings:"tnum";
  }
  table.data thead tr.totals-row td b{color:var(--navy);font-weight:700}
  table.data thead tr.totals-row td .muted{font-weight:400}
  table.data thead input, table.data thead select{
    width:100%;background:#fff;color:var(--ink);border:1px solid var(--line);
    padding:4px 8px;border-radius:4px;font-size:11px;font-family:inherit;
  }
  table.data thead input:focus, table.data thead select:focus{outline:none;border-color:var(--accent);box-shadow:0 0 0 2px var(--accent-soft)}
  table.data tbody td{
    padding:7px 10px;border-bottom:1px solid #eef1f5;white-space:nowrap;
    max-width:280px;overflow:hidden;text-overflow:ellipsis;font-size:13px;
  }
  table.data tbody tr:nth-child(even) td{background:var(--row-alt)}
  table.data tbody tr:hover td{background:#f0f4fa}
  td.num{text-align:right;font-variant-numeric:tabular-nums;font-feature-settings:"tnum"}
  a.odoo-link{color:var(--accent);text-decoration:none;font-weight:500}
  a.odoo-link:hover{text-decoration:underline}

  .pill{display:inline-block;padding:2px 10px;border-radius:4px;font-size:11px;
        border:1px solid transparent;font-weight:500}
  .pill-draft{background:#f1f3f7;color:#5b6b7c;border-color:#d8dee6}
  .pill-sent{background:#e3f2fd;color:#0d47a1;border-color:#bbdefb}
  .pill-sale{background:#e8f5e9;color:#1b5e20;border-color:#c8e6c9}
  .pill-done{background:#f5f5f5;color:#424242;border-color:#e0e0e0}
  .pill-cancel{background:#ffebee;color:#b71c1c;border-color:#ffcdd2}
  .ct-dist{background:#f3e5f5;color:#4a148c;border-color:#e1bee7}
  .ct-int {background:#e3f2fd;color:#0d47a1;border-color:#bbdefb}
  .ct-end {background:#e8f5e9;color:#1b5e20;border-color:#c8e6c9}
  .ct-none{background:#f5f5f5;color:#616161;border-color:#e0e0e0}
  .fx-pill{background:#fff3e0;color:#e65100;border-color:#ffe0b2}
  .err{color:var(--err);font-weight:600}
  .muted{color:var(--muted)}
  th.sortable{cursor:pointer;user-select:none}
  th.sortable:hover{color:var(--accent)}
  th .arrow{opacity:.5;font-size:10px;margin-left:4px}
  .nohit{text-align:center;padding:40px;color:var(--muted)}

  /* ===== Dashboard ===== */
  .dash{padding:20px 28px 40px;max-width:100%;display:none}
  .dash-toolbar{
    display:flex;align-items:center;gap:10px;padding:12px 16px;flex-wrap:wrap;
    background:#fff;border:1px solid var(--line);border-radius:6px;margin-bottom:16px;
  }
  .dash-toolbar label{color:var(--muted);font-size:12px;display:flex;align-items:center;gap:6px;font-weight:500}
  .dash-toolbar select, .dash-toolbar input[type="search"], .dash-toolbar input[type="number"]{
    background:#fff;color:var(--ink);border:1px solid var(--line);
    padding:6px 10px;border-radius:4px;font-size:13px;min-width:170px;font-family:inherit;
  }
  .dash-toolbar select:focus, .dash-toolbar input:focus{outline:none;border-color:var(--accent);box-shadow:0 0 0 2px var(--accent-soft)}
  .dash-toolbar .spacer{flex:1}

  .kpis{display:grid;grid-template-columns:repeat(6,minmax(160px,1fr));gap:12px;
        margin:0 0 20px}
  .kpi{background:var(--panel);border:1px solid var(--line);border-radius:6px;
       padding:16px 18px;position:relative;overflow:hidden;}
  .kpi::before{content:"";position:absolute;left:0;top:0;bottom:0;width:3px;background:var(--navy)}
  .kpi.dist::before{background:#8e24aa}
  .kpi.int::before {background:#1976d2}
  .kpi.end::before {background:#2e7d32}
  .kpi .lab{color:var(--muted);font-size:11px;text-transform:uppercase;
            letter-spacing:.6px;margin-bottom:6px;font-weight:600}
  .kpi .val{font-size:22px;font-weight:700;color:var(--navy);letter-spacing:-0.02em}
  .kpi .sub{color:var(--muted);font-size:11px;margin-top:3px}
  .kpi.dist .val{color:#6a1b9a}
  .kpi.int  .val{color:#0d47a1}
  .kpi.end  .val{color:#1b5e20}
  .kpi.clickable{cursor:pointer;transition:all .15s}
  .kpi.clickable:hover{border-color:var(--accent);box-shadow:0 2px 8px rgba(227,6,19,0.08);transform:translateY(-1px)}
  .kpi.clickable.active{border-color:var(--accent);box-shadow:0 0 0 2px var(--accent-soft)}
  .clickable{cursor:pointer}
  .clickable:hover{background:#f0f4fa !important}
  .active-filter{
    background:var(--accent-soft);color:var(--accent-dark);padding:3px 10px;border-radius:4px;
    border:1px solid #fbbec2;font-size:11px;font-weight:500;
  }
  .filter-chips{display:flex;gap:6px;flex-wrap:wrap;padding:0 0 12px}
  .chip{
    display:inline-flex;align-items:center;gap:4px;
    background:var(--accent-soft);color:var(--accent-dark);padding:4px 10px;border-radius:4px;
    border:1px solid #fbbec2;font-size:12px;font-weight:500;cursor:pointer;
  }
  .chip:hover{background:#fbd4d7}
  .chip .x{font-weight:700;opacity:.6}

  .charts{display:grid;grid-template-columns:2fr 1fr;gap:16px;margin-top:0}
  .card{background:var(--panel);border:1px solid var(--line);border-radius:6px;
        padding:18px 20px;}
  .card h3{margin:0 0 14px 0;font-size:14px;color:var(--navy);font-weight:600;
           display:flex;align-items:center;gap:8px;letter-spacing:-0.01em}
  .card h3 .muted{font-weight:400;font-size:11px}
  .chart-area{position:relative;height:360px}
  .chart-area.tall{height:480px}

  .tbl-sum{width:100%;border-collapse:collapse;margin-top:10px;font-size:13px;background:#fff}
  .tbl-sum th,.tbl-sum td{padding:8px 10px;border-bottom:1px solid var(--line);
                          text-align:left}
  .tbl-sum th{color:var(--navy);font-weight:600;background:#f8fafc;position:sticky;top:0;
              border-bottom:2px solid var(--line-strong);font-size:12px}
  .tbl-sum td.num{text-align:right;font-variant-numeric:tabular-nums;font-feature-settings:"tnum"}
  .tbl-sum tbody tr:hover td{background:#f0f4fa}
  .tbl-sum tfoot td{background:var(--bg-alt);font-weight:600;border-top:2px solid var(--navy)}
  /* th sortable común a tablas-resumen y detalles */
  th.sortable{cursor:pointer;user-select:none;transition:color 0.1s, background 0.1s}
  th.sortable:hover{color:var(--accent);background:#eef3fb}
  th.sortable.active{color:var(--accent-dark);background:#fde8ea}
  th.sortable .sort-arrow{
    display:inline-block;margin-left:4px;font-size:10px;color:#bbb;
    transition:color 0.1s;font-family:Arial,sans-serif;
  }
  th.sortable.active .sort-arrow{color:var(--accent)}
  th.sortable:hover .sort-arrow{color:var(--accent)}
  .tbl-sum tbody tr.pay-ok td{background:#e8f5e9}
  .tbl-sum tbody tr.pay-ok:hover td{background:#c8e6c9}
  .tbl-sum tbody tr.pay-pend td{background:#fff3e0}
  .tbl-sum tbody tr.pay-pend:hover td{background:#ffe0b2}

  /* ===== Tabla pivote de Tarifas ===== */
  .pv-wrap{overflow:auto;max-height:calc(100vh - 280px);background:#fff;border:1px solid var(--line);border-radius:6px}
  .tbl-pivot{border-collapse:separate;border-spacing:0;width:auto;min-width:100%;font-size:12px}
  .tbl-pivot th, .tbl-pivot td{padding:6px 9px;border-bottom:1px solid var(--line);border-right:1px solid #f0f2f5;white-space:nowrap}
  .tbl-pivot thead th{background:#f8fafc;color:var(--navy);font-weight:600;position:sticky;top:0;z-index:5}
  .tbl-pivot th.pv-period{font-variant-numeric:tabular-nums;text-align:right;font-size:11px}
  .tbl-pivot td.pv-cell{text-align:right;font-variant-numeric:tabular-nums;font-feature-settings:"tnum"}
  .tbl-pivot td.pv-cell.pv-ok{background:#fff}
  .tbl-pivot td.pv-cell.pv-low{background:#fff5f5;color:#9a0000}
  .tbl-pivot td.pv-cell.pv-empty{background:#fafafa;color:#bbb}
  .tbl-pivot td.pv-cell.pv-merged{background:#eaf5ff;border-right-color:#bcd9f0}
  .tbl-pivot td.pv-cell.pv-merged.pv-low{background:#ffe9e9;border-right-color:#f0bcbc}
  .tbl-pivot tbody tr:hover td.pv-cell{filter:brightness(0.96)}
  /* Columnas sticky a la izquierda */
  .tbl-pivot .pv-sticky{position:sticky;background:#fff;z-index:3}
  .tbl-pivot thead .pv-sticky{z-index:7;background:#f8fafc}
  .tbl-pivot .pv-code{left:0;min-width:130px;border-right:1px solid var(--line)}
  .tbl-pivot .pv-name{left:130px;min-width:240px;max-width:240px;overflow:hidden;text-overflow:ellipsis;border-right:1px solid var(--line)}
  .tbl-pivot .pv-fam{left:370px;min-width:170px;max-width:170px;overflow:hidden;text-overflow:ellipsis;font-size:11px;color:var(--navy);background:#f0f4fa !important;border-right:1px solid var(--line)}
  .tbl-pivot .pv-list{left:540px;min-width:90px;text-align:right;border-right:2px solid var(--line-strong);background:#f8fafc !important}
  .tbl-pivot tbody tr:hover .pv-sticky{background:#f5f9ff}
  /* Cabecera de grupo (familia) */
  .tbl-pivot tbody tr.pv-group-hdr td{
    background:linear-gradient(90deg,#fde8ea,#fff);
    border-top:2px solid var(--accent);border-bottom:1px solid var(--line);
    padding:8px 12px;position:sticky;left:0;z-index:2;
  }
  .tbl-pivot tbody tr.pv-group-hdr .pv-group-name{
    color:var(--accent-dark);font-weight:700;font-size:12px;letter-spacing:0.3px;text-transform:uppercase;
  }
  .pv-legend{display:flex;align-items:center;gap:10px;font-size:12px;color:var(--muted);margin-bottom:10px;flex-wrap:wrap}
  .pv-legend .pv-chip{display:inline-block;width:14px;height:14px;border-radius:3px;border:1px solid var(--line);margin-right:4px;vertical-align:middle}
  .pv-legend .pv-chip.pv-ok{background:#fff}
  .pv-legend .pv-chip.pv-low{background:#fff5f5;border-color:#f0bcbc}
  .pv-legend .pv-chip.pv-empty{background:#fafafa}

  .two-col{display:grid;grid-template-columns:repeat(2,1fr);gap:16px;margin-top:16px}
  @media (max-width: 1100px){
    .charts{grid-template-columns:1fr}
    .two-col{grid-template-columns:1fr}
    .kpis{grid-template-columns:repeat(3,1fr)}
  }

  /* ===== Condiciones ===== */
  .rules{padding:20px 28px 40px;display:none}
  .rules-tabs{display:flex;gap:0;border-bottom:1px solid var(--line);margin-bottom:20px;flex-wrap:wrap;
              background:#fff;padding:0 4px;border-radius:6px 6px 0 0}
  .rules-tabs button{
    background:transparent;color:var(--muted);border:0;padding:12px 20px;
    cursor:pointer;border-bottom:3px solid transparent;font-size:13px;font-weight:500;
    font-family:inherit;transition:all 0.15s;
  }
  .rules-tabs button:hover{color:var(--navy)}
  .rules-tabs button.active{color:var(--accent);border-bottom-color:var(--accent);font-weight:600}
  .rules-grid{display:grid;grid-template-columns:repeat(2,1fr);gap:16px}
  @media (max-width:1000px){.rules-grid{grid-template-columns:1fr}}
  .rules .card h3{margin-top:0}
  .rules .card h4{margin:0 0 6px 0;font-size:12px;color:var(--muted);text-transform:uppercase;letter-spacing:.5px}
  .type-box{padding:12px 14px;border-radius:4px;border:1px solid var(--line);margin-bottom:8px;
            border-left-width:3px}
  .type-box.t1{background:#faf5ff;border-color:#e9d5ff;border-left-color:#8e24aa}
  .type-box.t2{background:#f0f7ff;border-color:#bbdefb;border-left-color:#1976d2}
  .type-box .nm{font-weight:600;font-size:13px;color:var(--navy)}
  .type-box .ds{color:var(--muted);font-size:12px;margin-top:2px}
  .rules-note{background:var(--bg-alt);border:1px solid var(--line);border-left:3px solid var(--navy);
              border-radius:4px;padding:10px 14px;color:var(--muted);font-size:12px;margin-top:10px}
  .chip-type{display:inline-block;padding:2px 10px;border-radius:4px;font-size:11px;
             border:1px solid transparent;margin-right:4px;font-weight:500}
  .chip-type.t1{background:#f3e5f5;color:#4a148c;border-color:#e1bee7}
  .chip-type.t2{background:#e3f2fd;color:#0d47a1;border-color:#bbdefb}
  .chip-type.tnone{background:#f5f5f5;color:#616161;border-color:#e0e0e0}
  .formula{font-family:"SF Mono",Menlo,monospace;background:var(--bg-alt);padding:10px 14px;
           border-radius:4px;color:var(--navy);font-size:12px;border:1px solid var(--line);font-weight:500}

  /* ===== Icono info con tooltip al clic ===== */
  .info{
    display:inline-block;width:14px;height:14px;line-height:13px;text-align:center;
    border-radius:50%;background:#eef1f5;color:var(--muted);font-size:10px;font-weight:700;
    margin-left:4px;cursor:help;user-select:none;font-style:normal;
    border:1px solid var(--line);vertical-align:middle;
  }
  .info:hover{background:var(--accent);color:#fff;border-color:var(--accent)}
  .info-tip{
    position:absolute;background:#fff;color:var(--ink);
    border:1px solid var(--line-strong);border-left:3px solid var(--accent);border-radius:4px;padding:10px 12px;
    font-size:12px;max-width:320px;z-index:100;
    box-shadow:0 4px 16px rgba(26,36,51,0.12);line-height:1.5;
    pointer-events:none;
  }
</style>
</head>
<body>

<header>
  <h1>Comisiones Comerciales · Industrial Shields</h1>
  <span class="tag">Base Odoo</span>
  <div class="view-switch">
    <button id="btn-view-table" class="active">Tabla</button>
    <button id="btn-view-dash">Dashboard</button>
    <button id="btn-view-comm">Comisiones</button>
    <button id="btn-view-pay">PAGO</button>
    <button id="btn-view-tariffs">Tarifas</button>
    <button id="btn-view-rules">Condiciones</button>
  </div>
  <span class="header-divider"></span>
  <label class="my-view-wrap">
    Comercial:
    <select id="my-view">
      <option value="">(todos)</option>
    </select>
    <span id="my-view-tipo" class="muted" style="font-size:11px"></span>
  </label>
  <label class="my-view-wrap range-wrap">
    Periodo
    <select id="global-preset" title="Atajos comunes; al cambiar las fechas manualmente, vuelve a 'Personalizado'.">
      <option value="">Todos</option>
      <option value="2024">Año 2024</option>
      <option value="2025">Año 2025</option>
      <option value="2026">Año 2026</option>
      <option value="ytd">Año en curso (YTD)</option>
      <option value="last30">Últimos 30 días</option>
      <option value="last90">Últimos 90 días</option>
      <option value="last365">Últimos 365 días</option>
      <option value="q3_2025_plus">Desde Q3-2025</option>
      <option value="custom">Personalizado…</option>
    </select>
    <input type="date" id="global-from" title="Desde (incluido)">
    <span style="color:var(--muted);font-size:11px">→</span>
    <input type="date" id="global-to" title="Hasta (incluido)">
    <span class="info" data-tip="Filtro de tramo de fechas GLOBAL. Aplica a Tabla, Dashboard, Comisiones, PAGO y Tarifas (filtra por date_order de las líneas; en Tarifas filtra los periodos que aparecen en pivote).">i</span>
  </label>
  <label class="my-view-wrap">
    Producto
    <input type="search" id="filter-prod" placeholder="código o nombre…" style="min-width:140px;padding:5px 8px;border:1px solid var(--line);border-radius:4px;font-size:12px;font-family:inherit">
    <span class="info" data-tip="Filtra GLOBAL por código o nombre de producto (coincidencia parcial). Se aplica a todas las vistas. También puedes hacer click en un código de producto en cualquier tabla para fijar este filtro.">i</span>
  </label>
  <label class="my-view-wrap">
    Cliente
    <input type="search" id="filter-client" placeholder="nombre cliente…" style="min-width:140px;padding:5px 8px;border:1px solid var(--line);border-radius:4px;font-size:12px;font-family:inherit">
    <span class="info" data-tip="Filtra GLOBAL por nombre de cliente. También puedes hacer click en un cliente en cualquier tabla para fijar este filtro.">i</span>
  </label>
  <label class="my-view-wrap">
    <input type="checkbox" id="hide-excluded" checked> Ocultar excluidos
    <span class="info" data-tip="Oculta en todas las vistas a los usuarios que no devengan comisión (Website, ADMIN, Alba, Sònia, Abel, Macià, Luis Nunes, Francesc Duarri, Susana Guerra, Joan F. Aubets).">i</span>
  </label>
  <button id="btn-reset-all" class="panel-btn" title="Limpia TODOS los filtros de todas las vistas">Limpiar todo</button>
  <div class="header-sep"></div>
  <div class="stats" id="stats-bar">
    <!-- Se rellena segun la vista activa -->
  </div>
</header>

<!-- Loading overlay (con failsafe: se auto-oculta a los 8s pase lo que pase) -->
<div id="loading" style="position:fixed;inset:0;background:#ffffffee;display:flex;align-items:center;justify-content:center;z-index:9999;flex-direction:column;gap:14px">
  <div style="width:48px;height:48px;border:3px solid #eef1f5;border-top-color:var(--accent);border-radius:50%;animation:spin 0.9s linear infinite"></div>
  <div style="color:var(--muted);font-size:13px;font-weight:500">Descomprimiendo datos…</div>
</div>
<style>@keyframes spin{to{transform:rotate(360deg)}}</style>
<script>
  // Failsafe: quitar el overlay pase lo que pase después de 8s (por si hay un error JS)
  setTimeout(function(){
    var el=document.getElementById("loading");
    if(el) el.style.display="none";
  }, 8000);
</script>

<!-- ============================================================= VISTA TABLA -->
<div id="view-table">
  <div class="toolbar">
    <input id="q" type="search" placeholder="Buscar en toda la fila…" />
    <label style="display:flex;align-items:center;gap:6px;color:var(--muted)">
      <input type="checkbox" id="hideSections" checked> Ocultar secciones
    </label>
    <label style="display:flex;align-items:center;gap:6px;color:var(--muted)">
      <input type="checkbox" id="onlyCostMissing"> Solo coste=0
    </label>
    <label style="display:flex;align-items:center;gap:6px;color:var(--muted)">
      <input type="checkbox" id="onlyFX"> Solo FX aplicado
    </label>
    <label style="display:flex;align-items:center;gap:6px;color:var(--muted)">
      Mostrar
      <select id="tableLimit">
        <option value="200">200</option>
        <option value="500" selected>500</option>
        <option value="1000">1.000</option>
        <option value="5000">5.000</option>
        <option value="999999">Todas</option>
      </select>
    </label>
    <div class="spacer"></div>
    <button id="btn-clear">Limpiar filtros</button>
    <button id="btn-export">Exportar CSV</button>
  </div>
  <div class="wrap">
    <table id="tbl" class="data">
      <thead>
        <tr id="cols-row"></tr>
        <tr id="filters-row"></tr>
        <tr id="totals-row" class="totals-row"></tr>
      </thead>
      <tbody id="tbody"></tbody>
    </table>
    <div id="nohit" class="nohit" hidden>Sin resultados con los filtros actuales.</div>
  </div>
</div>

<!-- ============================================================= VISTA DASHBOARD -->
<div id="view-dash" class="dash">

  <div class="dash-toolbar">
    <label>Agrupar por <span class="info" data-tip="Dimensión por la que se agrega el dashboard (barras, donut y tabla resumen). 'Año' agrupa por YYYY; 'Mes' por YYYY-MM.">i</span>
      <select id="group-by">
        <option value="customer_type">Cliente tipo</option>
        <option value="commercial_entity_name">Cliente (entidad comercial)</option>
        <option value="product_category">Familia producto</option>
        <option value="product_name">Producto</option>
        <option value="country">País</option>
        <option value="salesperson">Comercial (SO)</option>
        <option value="partner_salesperson">Comercial (CE)</option>
        <option value="salesteam">Sales Team</option>
        <option value="pricelist">Pricelist (SO)</option>
        <option value="state_label">Status</option>
        <option value="year">Año</option>
        <option value="month">Mes (YYYY-MM)</option>
      </select>
    </label>
    <!-- El filtro Año ahora es GLOBAL en la cabecera -->

    <label>Métrica <span class="info" data-tip="Magnitud que se usa para ordenar y pintar las barras: Subtotal € (ventas sin IVA en EUR), Margen € (Subtotal − coste ficha × qty), Cantidad (uds vendidas), Nº líneas o Nº ofertas (pedidos únicos).">i</span>
      <select id="metric">
        <option value="price_subtotal_eur">Subtotal €</option>
        <option value="margin_eur">Margen €</option>
        <option value="qty">Cantidad</option>
        <option value="count">Nº líneas</option>
        <option value="orders">Nº ofertas</option>
      </select>
    </label>
    <label>Top N <span class="info" data-tip="Cuántas categorías mostrar en el gráfico principal. El resto quedan agrupadas si usas 'Todos'.">i</span>
      <select id="topn">
        <option value="10">10</option>
        <option value="15" selected>15</option>
        <option value="25">25</option>
        <option value="50">50</option>
        <option value="9999">Todos</option>
      </select>
    </label>
    <label>Cliente tipo <span class="info" data-tip="Clasificación del cliente: Distributor, Integrador o Cliente final. Se deriva primero de los tags del partner en Odoo y, si no hay, de la pricelist aplicada.">i</span>
      <select id="f-ct"><option value="">(todos)</option></select>
    </label>
    <label>País
      <select id="f-country"><option value="">(todos)</option></select>
    </label>
    <label>Estado <span class="info" data-tip="Estado del pedido en Odoo: Quotation (borrador), Quotation Sent (enviada), Sales Order (confirmada), Locked, Cancelled. Solo las Sales Order devengan comisión.">i</span>
      <select id="f-state"><option value="">(todos)</option></select>
    </label>
    <label>Comercial (SO) <span class="info" data-tip="Comercial asignado al pedido (sale.order.user_id). Es el que se usa para calcular la comisión.">i</span>
      <select id="f-sp"><option value="">(todos)</option></select>
    </label>
    <label>Comercial (CE) <span class="info" data-tip="Comercial asignado a la ficha del cliente / Commercial Entity (partner.user_id). Puede diferir del comercial del pedido.">i</span>
      <select id="f-sp-ce"><option value="">(todos)</option></select>
    </label>
    <label>Cliente
      <select id="f-client"><option value="">(todos)</option></select>
    </label>
    <label style="display:flex;align-items:center;gap:6px">
      <input type="checkbox" id="f-hideSec" checked> Ocultar secciones
    </label>
    <div class="spacer"></div>
    <button class="panel-btn" id="btn-dash-reset">Limpiar filtros</button>
    <button class="panel-btn" id="btn-dash-csv">Exportar agrupado CSV</button>
  </div>

  <div class="filter-chips" id="active-filters"></div>

  <div class="kpis" id="kpis"></div>

  <div class="charts">
    <div class="card">
      <h3>Subtotal y margen por <span id="dim-label" class="muted">grupo</span></h3>
      <div class="chart-area tall"><canvas id="chart-bar"></canvas></div>
    </div>
    <div class="card">
      <h3>Reparto <span class="muted">(% subtotal)</span></h3>
      <div class="chart-area"><canvas id="chart-donut"></canvas></div>
    </div>
  </div>

  <div class="two-col">
    <div class="card">
      <h3>Evolución por mes <span class="muted">(subtotal €)</span></h3>
      <div class="chart-area"><canvas id="chart-time"></canvas></div>
    </div>
    <div class="card">
      <h3>Mix por Cliente tipo <span class="muted">(subtotal €)</span></h3>
      <div class="chart-area"><canvas id="chart-ct"></canvas></div>
    </div>
  </div>

  <div class="card" style="margin-top:16px">
    <h3>Tabla resumen por <span id="dim-label2" class="muted">grupo</span></h3>
    <div style="max-height:420px;overflow:auto">
      <table class="tbl-sum" id="tbl-sum">
        <thead><tr>
          <th>Grupo</th>
          <th class="num">Nº líneas</th>
          <th class="num">Nº ofertas</th>
          <th class="num">Cantidad</th>
          <th class="num">Subtotal €</th>
          <th class="num">Margen €</th>
          <th class="num">Margen %</th>
          <th class="num">% s/total</th>
        </tr></thead>
        <tbody></tbody>
      </table>
    </div>
  </div>

  <div class="two-col">
    <div class="card">
      <h3>Por país <span class="muted">(subtotal €)</span></h3>
      <div class="chart-area"><canvas id="chart-country"></canvas></div>
    </div>
    <div class="card">
      <h3>Tabla por país</h3>
      <div style="max-height:340px;overflow:auto">
        <table class="tbl-sum" id="tbl-country">
          <thead><tr>
            <th>País</th>
            <th class="num">Ofertas</th>
            <th class="num">Líneas</th>
            <th class="num">Subtotal €</th>
            <th class="num">Margen €</th>
            <th class="num">Margen %</th>
            <th class="num">% s/total</th>
          </tr></thead>
          <tbody></tbody>
        </table>
      </div>
    </div>
  </div>

  <div class="two-col">
    <div class="card">
      <h3>Por Comercial (CE) <span class="muted">(subtotal €)</span></h3>
      <div class="chart-area"><canvas id="chart-sp"></canvas></div>
    </div>
    <div class="card">
      <h3>Tabla por Comercial (CE) <span class="muted">· con Tipo de comisión</span></h3>
      <div style="max-height:340px;overflow:auto">
        <table class="tbl-sum" id="tbl-sp">
          <thead><tr>
            <th>Comercial (CE)</th>
            <th>Tipo</th>
            <th class="num">Base %</th>
            <th class="num">Clientes</th>
            <th class="num">Ofertas</th>
            <th class="num">Líneas</th>
            <th class="num">Subtotal €</th>
            <th class="num">Margen €</th>
            <th class="num">Margen %</th>
            <th class="num">% s/total</th>
          </tr></thead>
          <tbody></tbody>
        </table>
      </div>
    </div>
  </div>

  <div class="card" style="margin-top:16px">
    <h3>Resumen por Cliente <span class="muted">(entidades comerciales, ordenadas por subtotal)</span></h3>
    <div style="max-height:560px;overflow:auto">
      <table class="tbl-sum" id="tbl-clients">
        <thead><tr>
          <th>Cliente</th>
          <th>Tipo</th>
          <th>País</th>
          <th>Pricelist</th>
          <th>Comercial</th>
          <th class="num">Ofertas</th>
          <th class="num">Líneas</th>
          <th class="num">Subtotal €</th>
          <th class="num">Margen €</th>
          <th class="num">Margen %</th>
          <th class="num">Ticket medio €</th>
          <th>Última</th>
        </tr></thead>
        <tbody></tbody>
      </table>
    </div>
  </div>

</div>

<!-- ============================================================= VISTA COMISIONES -->
<div id="view-comm" class="rules">
  <div class="dash-toolbar" style="margin-bottom:8px">
    <label>Vista <span class="info" data-tip="Resumen: una fila por comercial sumando todo el rango. Mensual: pivote comercial × mes con totales. YoY: comparativa año a año del acumulado mensual.">i</span>
      <select id="comm-view">
        <option value="summary" selected>Resumen por comercial</option>
        <option value="monthly">Detalle mensual</option>
        <option value="yoy">Acumulado anual (YoY)</option>
      </select>
    </label>
    <label>Métrica <span class="info" data-tip="Qué importe muestran las celdas. Devengada = la generada por SO confirmados. PAGABLE = solo las cobradas y facturadas en el periodo (lo que se liquida realmente).">i</span>
      <select id="comm-metric">
        <option value="devengada" selected>Comisión devengada</option>
        <option value="pagable">Comisión PAGABLE</option>
        <option value="ventas">Ventas comisionables</option>
      </select>
    </label>
    <!-- Comercial: oculto, sincronizado desde 'Mi vista' del header global -->
    <select id="comm-sp" style="display:none"><option value="">(todos)</option></select>
    <label style="display:flex;align-items:center;gap:6px">
      <input type="checkbox" id="comm-detail" checked> Ver detalle por pedido
    </label>
    <label>Buscar SO <span class="info" data-tip="Busca por número de pedido. Coincidencia parcial (ej: '36034' lo encuentra dentro de SO36034).">i</span>
      <input type="search" id="comm-q" placeholder="SO36034…" style="min-width:140px">
    </label>
    <label>Comisión ≥ <span class="info" data-tip="Filtra pedidos cuya comisión TOTAL devengada sea ≥ este valor (€). Útil para ocultar ruido.">i</span>
      <input type="number" id="comm-min" placeholder="0" step="10" style="width:100px">
    </label>
    <label>Comisión ≤
      <input type="number" id="comm-max" placeholder="∞" step="10" style="width:100px">
    </label>
    <label>Cobro <span class="info" data-tip="Estado de cobro agregado del pedido. Cobrado = todas las facturas posted y pagadas. Parcial = alguna residual. No cobrado = facturada pero sin pagar. A facturar / Sin factura para pedidos aún sin facturar.">i</span>
      <select id="comm-pay">
        <option value="">(todos)</option>
        <option value="paid">Cobrado</option>
        <option value="partial">Parcialmente</option>
        <option value="not_paid">No cobrado (factura emitida)</option>
        <option value="to_invoice">Pendiente de facturar</option>
        <option value="none">Sin factura</option>
      </select>
    </label>
    <div class="spacer"></div>
    <button class="panel-btn" id="btn-comm-reset">Limpiar</button>
    <button class="panel-btn" id="btn-comm-csv">Exportar CSV</button>
  </div>
  <div class="rules-tabs" id="comm-tabs"></div>
  <div id="comm-body"></div>
</div>

<!-- ============================================================= VISTA PAGO COMISIONES -->
<div id="view-pay" class="rules">
  <div class="dash-toolbar" style="margin-bottom:8px">
    <label>Vista <span class="info" data-tip="Por periodo: una fila por (periodo, comercial). Acumulado: suma todos los periodos seleccionados en una sola fila por comercial — útil para liquidar varios meses juntos.">i</span>
      <select id="pay-view">
        <option value="byperiod" selected>Por periodo</option>
        <option value="cum">Acumulado en rango</option>
      </select>
    </label>
    <label>Desde mes <span class="info" data-tip="Mes inicial del acumulado (incluido). Solo se aplica en vista 'Acumulado'.">i</span>
      <input type="month" id="pay-from" style="min-width:135px">
    </label>
    <label>Hasta mes <span class="info" data-tip="Mes final del acumulado (incluido). Solo se aplica en vista 'Acumulado'.">i</span>
      <input type="month" id="pay-to" style="min-width:135px">
    </label>
    <label>Periodo <span class="info" data-tip="Solo en vista 'Por periodo': filtra a un solo periodo concreto (Q3-2025, 2026-01, etc.).">i</span>
      <select id="pay-period"><option value="">(todos desde Q3-2025)</option></select>
    </label>
    <label style="display:flex;align-items:center;gap:6px">
      <input type="checkbox" id="pay-detail"> Ver detalle por línea
      <span class="info" data-tip="Activa para listar cada línea de producto con su comisión, estado factura, estado cobro y periodo asignado.">i</span>
    </label>
    <!-- Comercial: oculto, sincronizado desde 'Mi vista' del header global -->
    <select id="pay-sp" style="display:none"><option value="">(todos)</option></select>
    <div class="spacer"></div>
    <button class="panel-btn drive-btn drive-sync" id="btn-pay-sync"
       title="Lee la hoja de Drive ahora mismo y refresca las columnas Pagado/Pendiente sin tocar nada más.">
       <span class="drive-ico">🔄</span> Sincronizar con Drive
    </button>
    <a id="btn-pay-drive" class="panel-btn drive-btn"
       href="__DRIVE_SHEET_URL__" target="_blank" rel="noopener"
       title="Abrir el registro de pagos en Google Drive (solo apm@industrialshields.com)">
       <span class="drive-ico">📊</span> Hoja Drive
    </a>
    <button class="panel-btn" id="btn-pay-csv">Exportar CSV</button>
  </div>
  <input type="file" id="pay-sync-file" accept=".csv,text/csv" style="display:none">
  <div id="pay-toast" class="pay-toast" style="display:none"></div>
  <div class="rules-note" style="margin:0 0 12px 0">
    <b>Registro de pagos en Google Drive.</b>
    El registro vive en la hoja
    <a href="__DRIVE_SHEET_URL__" target="_blank" rel="noopener" class="drive-link">
      <b>Pagos Comisiones - Industrial Shields</b>
    </a>
    (Drive). Edita ahí los importes ya pagados (<code>importe_pagado_eur</code>),
    fecha y notas. Pulsa <b>🔄 Sincronizar con Drive</b> y las columnas
    <b>Pagado / Pendiente</b> se refrescarán al instante.
    <br><br>
    <b>Lógica de pago:</b>
    <b>2025-Q3 y 2025-Q4</b>: comisiones se liquidan trimestralmente sobre las líneas <i>facturadas y cobradas</i>.
    <b>Desde 2026</b>: liquidación <i>mensual</i> de las líneas facturadas y cobradas en el mes.
    Las columnas <b>Generado / Facturado / Cobrado</b> agrupan cada línea por
    <code>date_order</code> (Generado) y <code>last_invoice_date</code> (Facturado y Cobrado).
    Solo aplica desde Q3-2025 — periodos anteriores no se incluyen aquí.
  </div>
  <div id="pay-body"></div>
</div>

<!-- ============================================================= VISTA TARIFAS -->
<div id="view-tariffs" class="rules">
  <div class="dash-toolbar" style="margin-bottom:8px">
    <label>Buscar producto <span class="info" data-tip="Busca por código o nombre del producto. Coincidencia parcial.">i</span>
      <input type="search" id="tariffs-q" placeholder="código o nombre..." style="min-width:240px">
    </label>
    <label>Vista <span class="info" data-tip="Pivote: productos en filas, periodos en columnas (recomendado). Lista: una fila por (producto, periodo).">i</span>
      <select id="tariffs-mode">
        <option value="pivot" selected>Pivote (productos × periodos)</option>
        <option value="list">Lista detallada</option>
      </select>
    </label>
    <label>Granularidad <span class="info" data-tip="Auto: trimestral para 2024-2025 + mensual para 2026. Anual: solo agregado YYYY. Trimestral: solo Q1..Q4. Mensual: solo MM. El tramo de fechas global de la cabecera filtra qué periodos se muestran.">i</span>
      <select id="tariffs-gran">
        <option value="auto" selected>Auto (trim + mensual 2026)</option>
        <option value="annual">Anual</option>
        <option value="quarterly">Trimestral</option>
        <option value="monthly">Mensual</option>
      </select>
    </label>
    <label>Familia <span class="info" data-tip="Filtra por familia de productos (Controllers/Ethernet PLC, Panel PC, IOs Module, etc). 'todas' muestra todo. Activa 'Agrupar' para ver separadores por familia.">i</span>
      <select id="tariffs-fam"><option value="">(todas)</option></select>
    </label>
    <label style="display:flex;align-items:center;gap:6px">
      <input type="checkbox" id="tariffs-group" checked> Agrupar por familia
      <span class="info" data-tip="Inserta una fila separadora por cada familia, con el número de productos.">i</span>
    </label>
    <label class="lbl-only-list" style="display:none">Confianza <span class="info" data-tip="Filtra por confianza mínima del cálculo. Productos con baja confianza (precios muy variables) NO se usan para la comisión.">i</span>
      <select id="tariffs-conf">
        <option value="">(todos)</option>
        <option value="0.55">≥ 0,55 (los que SÍ se usan)</option>
        <option value="-0.55">&lt; 0,55 (descartados)</option>
      </select>
    </label>
    <div class="spacer"></div>
    <button class="panel-btn" id="btn-tariffs-csv">Exportar CSV</button>
  </div>
  <div class="rules-note" style="margin:0 0 12px 0">
    <b>Cómo se calcula:</b> moda de <code>price_unit_eur / (1 − dto%)</code> sobre las ventas/cotizaciones de cada producto en cada periodo.
    Granularidad: <b>trimestral 2024-2025</b>, <b>mensual 2026</b> + agregados anuales como fallback.
    Solo se usa para detectar descuentos efectivos cuando <i>confidence</i> ≥ 0,55 y la línea va en pricelist pública.
    <br>
    <b>Pivote:</b> celdas consecutivas con la misma tarifa se fusionan automáticamente (— —) para que sea fácil ver cuándo hubo cambio de precio.
    Pasa el ratón por una celda para ver unidades, importe y confianza.
  </div>
  <div id="tariffs-body"></div>
</div>

<!-- ============================================================= VISTA CONDICIONES -->
<div id="view-rules" class="rules">
  <div class="rules-tabs" id="rules-tabs"></div>
  <div id="rules-body"></div>
</div>

<script id="payload" type="application/octet-stream">__PAYLOAD__</script>
<script id="tariffs" type="application/octet-stream">__TARIFFS__</script>
<script id="pagos" type="application/octet-stream">__PAGOS__</script>

<script>
// =============================================================================
// Columnas del DATA_SO (vista Tabla)
// =============================================================================
const COLUMNS = [
  { key:"order_name",              label:"SO",                  type:"link",  w:100 },
  { key:"external_id",             label:"External ID",         type:"text",  w:180 },
  { key:"state_label",             label:"Status",              type:"pill",  w:120 },
  { key:"customer_type",           label:"Cliente tipo",        type:"ctype", w:130 },
  { key:"date_order",              label:"Order date",          type:"text",  w:100 },
  { key:"create_date",             label:"Creation date",       type:"text",  w:100 },
  { key:"contact_name",            label:"Contact",             type:"text",  w:200 },
  { key:"commercial_entity_name",  label:"Commercial entity",   type:"text",  w:200 },
  { key:"country",                 label:"Country",             type:"text",  w:110 },
  { key:"salesperson",             label:"Salesperson (SO)",    type:"text",  w:140 },
  { key:"sp_tipo",                 label:"Tipo",                type:"sptipo",w:70  },
  { key:"salesteam",               label:"Sales Team (SO)",     type:"text",  w:100 },
  { key:"partner_salesperson",     label:"Salesperson (CE)",    type:"text",  w:140 },
  { key:"partner_salesteam",       label:"Sales Team (CE)",     type:"text",  w:100 },
  { key:"partner_pricelist",       label:"Pricelist (CE)",      type:"text",  w:180 },
  { key:"pricelist",               label:"Pricelist (SO)",      type:"text",  w:180 },
  { key:"currency",                label:"Cur",                 type:"text",  w:60  },
  { key:"fx_rate",                 label:"FX  (foreign/€)",     type:"fx",    w:120 },
  { key:"product_code",            label:"Ref",                 type:"text",  w:140 },
  { key:"product_name",            label:"Product",             type:"text",  w:240 },
  { key:"product_category",        label:"Category",            type:"text",  w:220 },
  { key:"qty",                     label:"Qty",                 type:"num",   w:70  },
  { key:"price_unit",              label:"Unit (cur)",          type:"money", w:90  },
  { key:"list_price",              label:"List €",              type:"money", w:90  },
  { key:"standard_price",          label:"Cost €",              type:"money", w:90  },
  { key:"discount_pct",            label:"Dto %",               type:"pct",   w:70  },
  { key:"price_subtotal",          label:"Subtotal (cur)",      type:"money", w:120 },
  { key:"price_subtotal_eur",      label:"Subtotal €",          type:"money", w:110 },
  { key:"margin_eur",              label:"Margin €",            type:"money", w:100 },
  { key:"margin_pct",              label:"Margin %",            type:"pct",   w:90  },
];

// --- Descomprimir payload (gzip+base64) usando pako ---
function _decodeB64Gzip(elId){
  const el = document.getElementById(elId);
  if (!el) return null;
  const b64 = el.textContent.trim();
  if (!b64) return null;
  const raw = atob(b64);
  const bytes = new Uint8Array(raw.length);
  for (let i=0; i<raw.length; i++) bytes[i] = raw.charCodeAt(i);
  return JSON.parse(pako.ungzip(bytes, { to: 'string' }));
}
let DATA = _decodeB64Gzip("payload");
let TARIFFS = _decodeB64Gzip("tariffs") || { products:{}, by_product_year:{} };

// Pagos registrados (importes ya liquidados al comercial). Fuente de verdad:
// Google Sheet "Pagos Comisiones - Industrial Shields". Cache local CSV.
let PAGOS_REG = _decodeB64Gzip("pagos") || { rows:[], drive_url:"" };
const PAGOS_BY = (() => {
  const m = new Map();
  for (const r of (PAGOS_REG.rows || [])){
    m.set(r.period + "|" + r.sp, { pagado: +r.pagado || 0, fecha: r.fecha || "", notas: r.notas || "" });
  }
  return m;
})();
function pagosFor(period, sp){
  return PAGOS_BY.get(period + "|" + sp) || { pagado:0, fecha:"", notas:"" };
}

// =============================================================================
// SINCRONIZACIÓN AUTOMÁTICA con la hoja de Drive
// =============================================================================
const DRIVE_SHEET_URL_JS = (PAGOS_REG && PAGOS_REG.drive_url) || "";
const DRIVE_SHEET_ID_JS  = (PAGOS_REG && PAGOS_REG.drive_id)  || "";
function _driveExportCsvUrl(){
  if (!DRIVE_SHEET_ID_JS) return null;
  return `https://docs.google.com/spreadsheets/d/${DRIVE_SHEET_ID_JS}/export?format=csv`;
}

// Toast simple — top-right
function _showToast(msg, kind, opts){
  opts = opts || {};
  const el = document.getElementById('pay-toast');
  if (!el) return;
  const actionsHtml = (opts.actions || []).map((a,i) =>
    `<button class="pt-btn" data-i="${i}">${escapeHtml(a.label)}</button>`).join("");
  el.className = 'pay-toast ' + (kind || '');
  el.innerHTML = `
    <div class="pt-msg">${msg}${actionsHtml ? `<div class="pt-actions">${actionsHtml}</div>` : ""}</div>
    <button class="pt-x" title="Cerrar">×</button>`;
  el.style.display = 'flex';
  el.querySelector('.pt-x').onclick = () => { el.style.display='none'; };
  (opts.actions || []).forEach((a, i) => {
    const btn = el.querySelector(`.pt-btn[data-i="${i}"]`);
    if (btn) btn.onclick = () => { try { a.fn(); } finally { el.style.display='none'; } };
  });
  if (opts.timeout !== 0) setTimeout(() => { el.style.display='none'; }, opts.timeout || 6000);
}

// Parser CSV simple — soporta ; y , y comillas
function _parseCsv(text){
  // Quitar BOM
  if (text.charCodeAt(0) === 0xFEFF) text = text.slice(1);
  // Detectar separador: el más frecuente en la primera línea
  const firstLine = text.split(/\r?\n/, 1)[0] || "";
  const nC = (firstLine.match(/,/g) || []).length;
  const nS = (firstLine.match(/;/g) || []).length;
  const sep = nS > nC ? ';' : ',';
  // Parser
  const rows = [];
  let i = 0, field = "", row = [], inQ = false;
  while (i < text.length){
    const c = text[i];
    if (inQ){
      if (c === '"'){
        if (text[i+1] === '"'){ field += '"'; i += 2; continue; }
        inQ = false; i++; continue;
      }
      field += c; i++;
    } else {
      if (c === '"' && field === ""){ inQ = true; i++; continue; }
      if (c === sep){ row.push(field); field = ""; i++; continue; }
      if (c === '\n' || c === '\r'){
        row.push(field); field = "";
        if (row.length > 1 || row[0] !== "") rows.push(row);
        row = [];
        if (c === '\r' && text[i+1] === '\n') i += 2; else i++;
        continue;
      }
      field += c; i++;
    }
  }
  if (field !== "" || row.length){ row.push(field); rows.push(row); }
  if (!rows.length) return [];
  // Primera fila = headers
  const head = rows[0].map(h => h.trim());
  return rows.slice(1).map(r => {
    const o = {};
    for (let j=0; j<head.length; j++) o[head[j]] = (r[j] || "").trim();
    return o;
  });
}

// Aplica los datos de pagos (rows del CSV) a PAGOS_BY y re-renderiza
function _applyPagosCsv(text){
  const rows = _parseCsv(text);
  if (!rows.length) throw new Error('CSV vacío o no se pudo parsear');
  // Validar columnas
  const need = ['periodo_cobro', 'comercial', 'importe_pagado_eur'];
  if (!need.every(k => k in rows[0])) {
    throw new Error('La hoja no tiene las columnas esperadas (periodo_cobro, comercial, importe_pagado_eur, fecha_pago, notas).');
  }
  // Re-llenar PAGOS_BY
  PAGOS_BY.clear();
  let n = 0;
  for (const r of rows){
    const period = (r.periodo_cobro || "").trim();
    const sp     = (r.comercial || "").trim();
    if (!period || !sp) continue;
    const pagado = parseFloat((r.importe_pagado_eur || '0').replace(',', '.')) || 0;
    PAGOS_BY.set(period + "|" + sp, {
      pagado, fecha: (r.fecha_pago || "").trim(), notas: (r.notas || "").trim(),
    });
    n++;
  }
  // Refrescar la vista
  if (document.getElementById('view-pay').style.display !== 'none') renderPayments();
  return n;
}

// Botón "Sincronizar con Drive" — intenta fetch, si falla cae a manual
async function syncPagosFromDrive(){
  const btn = document.getElementById('btn-pay-sync');
  const url = _driveExportCsvUrl();
  if (!url){
    _showToast('No hay URL de hoja Drive configurada.', 'err'); return;
  }
  // Detectar file:// — sabemos que va a fallar por CORS, mejor explicar directamente
  if (location.protocol === 'file:'){
    _showToast(
      `<b>Para sincronización automática</b> hay que abrir el HTML desde un servidor local
      (Chrome bloquea las peticiones a Drive cuando el HTML se abre como <code>file://</code>).<br>
      <b>Solución 1-clic:</b> doble-click en <code>Abrir_Comisiones.command</code>
      (en la misma carpeta que el HTML). Arranca un mini servidor y abre el HTML
      en <code>http://localhost</code>. Desde ahí el botón funciona perfecto.<br>
      <b>Mientras tanto:</b> puedes cargar manualmente el CSV descargado de Drive:`,
      'warn',
      {
        timeout: 0,
        actions: [
          { label: '📂 Cargar CSV descargado', fn: () => document.getElementById('pay-sync-file').click() },
          { label: '📊 Abrir hoja Drive',       fn: () => window.open(DRIVE_SHEET_URL_JS, '_blank') },
        ]
      }
    );
    return;
  }
  if (btn) btn.classList.add('is-loading');
  try {
    // Fetch con cookies de sesión Google del usuario
    const resp = await fetch(url, { credentials: 'include', mode: 'cors' });
    if (!resp.ok) throw new Error('HTTP ' + resp.status);
    const text = await resp.text();
    // Google a veces devuelve HTML de login con 200 OK
    if (/<html/i.test(text.slice(0, 200))) throw new Error('Drive devolvió HTML (no autenticado o sin permiso). Abre la hoja en otra pestaña para iniciar sesión y vuelve a probar.');
    const n = _applyPagosCsv(text);
    if (btn) btn.classList.remove('is-loading');
    _showToast(`✓ Sincronizado <b>${n}</b> filas desde Drive · ${new Date().toLocaleTimeString('es-ES',{hour:'2-digit',minute:'2-digit'})}`, 'ok');
  } catch (err){
    if (btn) btn.classList.remove('is-loading');
    _showToast(
      `<b>No se pudo leer Drive automáticamente</b> (${escapeHtml(err.message||err)})<br>
       Asegúrate de tener iniciada la sesión de Google con <b>apm@industrialshields.com</b>
       en este navegador, y vuelve a probar.<br>
       Mientras tanto puedes cargar el CSV manualmente:`,
      'warn',
      {
        timeout: 0,
        actions: [
          { label: '📂 Cargar CSV descargado', fn: () => document.getElementById('pay-sync-file').click() },
          { label: '📊 Abrir hoja Drive',       fn: () => window.open(DRIVE_SHEET_URL_JS, '_blank') },
        ]
      }
    );
  }
}

// Cargar CSV descargado manualmente
function _onPaySyncFile(ev){
  const f = ev.target.files && ev.target.files[0];
  if (!f) return;
  const rd = new FileReader();
  rd.onload = e => {
    try {
      const n = _applyPagosCsv(e.target.result);
      _showToast(`✓ Cargado <b>${n}</b> filas desde <code>${escapeHtml(f.name)}</code>`, 'ok');
    } catch (err){
      _showToast(`Error al leer el CSV: ${escapeHtml(err.message||err)}`, 'err', {timeout:0});
    }
    ev.target.value = "";
  };
  rd.readAsText(f, 'utf-8');
}

// Filtramos por confianza minima — productos con precios muy variables (Pre-payment,
// Custom...) tienen confidence baja y no deberian usarse para penalizar comision.
const MIN_TARIFF_CONFIDENCE = 0.55;

// Calcula la clave de periodo: 'YYYY-Q1..Q4' si yr<=2025, 'YYYY-MM' si yr>=2026
function periodKey(date_order){
  const d = (date_order||"").slice(0,10);
  const yr = d.slice(0,4);
  if (!/^\d{4}$/.test(yr)) return null;
  const yrN = parseInt(yr, 10);
  const mo = (d.length >= 7) ? Math.max(1, parseInt(d.slice(5,7) || "1", 10)) : 1;
  if (yrN <= 2025){
    const q = Math.floor((mo-1)/3) + 1;
    return `${yr}-Q${q}`;
  }
  return `${yr}-${String(mo).padStart(2,"0")}`;
}

// Devuelve la tarifa deducida para (producto, fecha). Cascada:
// 1) trimestre exacto (si conf >= MIN), 2) año completo (si conf >= MIN),
// 3) periodo más reciente anterior con tarifa fiable.
function getDeducedTariff(code, date_order){
  if (!code || !TARIFFS) return null;
  const p = TARIFFS.by_product_year[code];
  if (!p) return null;
  const acceptable = (d) => d && d.tariff > 0 && (d.confidence == null || d.confidence >= MIN_TARIFF_CONFIDENCE);
  const yr = (date_order||"").slice(0,4);
  // 1) trimestre o año exacto
  const pk = periodKey(date_order);
  if (acceptable(p[pk])) return p[pk].tariff;
  // 2) si el periodo era trimestre, probar año completo como fallback
  if (pk && pk !== yr && acceptable(p[yr])) return p[yr].tariff;
  // 3) periodo más reciente anterior con tarifa fiable
  const keys = Object.keys(p).sort();
  let cand = null;
  for (const k of keys){ if (k <= pk && acceptable(p[k])) cand = p[k]; }
  if (cand) return cand.tariff;
  return null;
}

// Solo aplicamos el "descuento efectivo" cuando la pricelist es PUBLICA
// (Public Pricelist, Website Standard...). En distribuidor/integrador/Mouser
// el precio_unit lo fija el contrato, no el comercial — no debemos penalizarle.
function _isPublicPricelist(pl){
  if (!pl) return false;
  const n = pl.toLowerCase().trim();
  return n.startsWith("public") || n.startsWith("website");
}

// Calcula el descuento EFECTIVO de una linea contra la tarifa deducida.
// Devuelve {declared, effective, deduced_tariff, suspect, use, applies_eff}
function effectiveDiscount(r){
  const declared = r.discount_pct || 0;
  const applies_eff = _isPublicPricelist(r.pricelist);
  if (!applies_eff){
    return { declared, effective: declared, deduced_tariff: null,
             use: declared, suspect: false, applies_eff: false };
  }
  const code = r.product_code || "";
  const tariff = getDeducedTariff(code, r.date_order);
  let effective = declared;
  if (tariff && r.price_unit_eur > 0){
    effective = Math.max(0, Math.min(100, (1 - r.price_unit_eur / tariff) * 100));
  }
  const use = Math.max(declared, effective);
  const suspect = (effective - declared) >= 5 && tariff != null;
  return { declared, effective, deduced_tariff: tariff, use, suspect, applies_eff: true };
}

// --- Helper icono info con tooltip al clic ---
function infoIcon(txt){
  return `<span class="info" data-tip="${(""+txt).replace(/"/g,"&quot;")}">i</span>`;
}
// Listener global: al clic en .info mostramos burbuja; al clicar fuera, la quitamos
document.addEventListener("click", (e) => {
  document.querySelectorAll(".info-tip").forEach(t => t.remove());
  const el = e.target.closest(".info");
  if (!el) return;
  const tip = document.createElement("div");
  tip.className = "info-tip";
  tip.textContent = el.dataset.tip || "";
  document.body.appendChild(tip);
  const r = el.getBoundingClientRect();
  let left = r.left + window.scrollX;
  let top  = r.bottom + window.scrollY + 6;
  // Evitar desborde por la derecha
  const maxLeft = window.innerWidth - 340;
  if (left > maxLeft) left = maxLeft;
  tip.style.left = left + "px";
  tip.style.top  = top  + "px";
  e.stopPropagation();
});

let SORT = { key: "date_order", dir: -1 };
let TABLE_LIMIT = 500;  // filas a renderizar en tabla (modificable por toolbar)

// =============================================================================
// ORDEN GENERICO DE TABLAS (click en cabecera asc/desc)
// =============================================================================
// Estado de orden por tabla.  -1 = descendente, 1 = ascendente.
const TBL_SORT = {
  comm_sum:   { key: 'com_total',  dir: -1 },
  comm_det:   { key: 'com_total',  dir: -1 },
  pay_sum:    { key: '_periodOrd', dir:  1 },
  pay_det:    { key: 'commission', dir: -1 },
  tarif_list: { key: 'code',       dir:  1 },
};
function _setTblSort(table, key, defaultDir){
  const s = TBL_SORT[table];
  if (!s) return;
  if (s.key === key) s.dir = -s.dir;
  else { s.key = key; s.dir = (defaultDir != null) ? defaultDir : -1; }
}
function _sortBy(arr, key, dir){
  const sign = dir < 0 ? -1 : 1;
  arr.sort((a, b) => {
    let va = a[key], vb = b[key];
    if (va == null && vb == null) return 0;
    if (va == null) return 1;
    if (vb == null) return -1;
    if (typeof va === 'number' && typeof vb === 'number') return sign * (va - vb);
    // strings con localeCompare español
    return sign * String(va).localeCompare(String(vb), 'es', {numeric:true});
  });
  return arr;
}
// Genera el HTML de un th sortable. opts: {tip, num, defaultDir}
function _sortHead(table, key, label, opts){
  opts = opts || {};
  const s = TBL_SORT[table] || {};
  const active = s.key === key;
  const arrow = active ? (s.dir > 0 ? '▲' : '▼') : '⇅';
  const help = opts.tip ? ' ' + infoIcon(opts.tip) : '';
  const cls = 'sortable' + (active ? ' active' : '') + (opts.num ? ' num' : '');
  return `<th class="${cls}" data-tbl="${table}" data-sk="${escapeHtml(key)}" data-dd="${opts.defaultDir != null ? opts.defaultDir : -1}">`+
         `<span class="sortable-lbl">${label}</span>${help} `+
         `<span class="sort-arrow">${arrow}</span></th>`;
}
// Click delegation para th.sortable
document.addEventListener('click', (e) => {
  const th = e.target.closest('th.sortable');
  if (!th || !th.dataset.tbl) return;
  const table = th.dataset.tbl;
  const key = th.dataset.sk;
  const defDir = parseInt(th.dataset.dd, 10);
  _setTblSort(table, key, defDir);
  if (table === 'comm_sum' || table === 'comm_det') renderCommissions();
  else if (table === 'pay_sum' || table === 'pay_det') renderPayments();
  else if (table === 'tarif_list') renderTariffs();
});

// Sincroniza el selector local de comercial (de una vista) con "Mi vista" global.
// Se llama al inicio de cada render para garantizar consistencia, incluso si
// el dropdown se acaba de poblar.
function _syncMyViewToLocal(localId){
  const my = document.getElementById('my-view');
  const local = document.getElementById(localId);
  if (!my || !local) return;
  const v = my.value || "";
  // Solo aplicar si la opcion existe en el dropdown local
  if (v && !local.querySelector(`option[value="${v.replace(/"/g,'\\"')}"]`)) return;
  local.value = v;
}

// Filtro tramo de fechas GLOBAL (cabecera). Devuelve {from, to} en YYYY-MM-DD o null.
function globalDateRange(){
  const f = document.getElementById("global-from");
  const t = document.getElementById("global-to");
  const from = (f && f.value) ? f.value : null;
  const to   = (t && t.value) ? t.value : null;
  if (!from && !to) return null;
  return { from, to };
}
// Compatibilidad legacy (algun callsite antiguo). Devuelve año si el rango cae todo dentro.
function globalYear(){
  const r = globalDateRange();
  if (!r) return null;
  const yF = r.from && r.from.slice(0,4);
  const yT = r.to   && r.to.slice(0,4);
  if (yF && yT && yF === yT) return yF;
  return null;
}

// Devuelve DATA filtrando lineas de usuarios excluidos, rango global de fechas,
// "Mi vista" (comercial), filtro de producto y filtro de cliente — TODOS los
// filtros globales del header se aplican aquí.
function visibleData(){
  let src = DATA;
  const hide = document.getElementById("hide-excluded");
  if (hide && hide.checked){
    src = src.filter(r => !EXCLUDED_SP_SET.has(r.salesperson));
  }
  const r = globalDateRange();
  if (r){
    if (r.from) src = src.filter(x => (x.date_order||"") >= r.from);
    if (r.to)   src = src.filter(x => (x.date_order||"") <= (r.to + 'T23:59:59'));
  }
  // Mi vista (comercial)
  const my = document.getElementById('my-view');
  if (my && my.value){
    src = src.filter(x => x.salesperson === my.value);
  }
  // Filtro Producto (código o nombre, partial)
  const fp = document.getElementById('filter-prod');
  if (fp && fp.value){
    const q = fp.value.toLowerCase().trim();
    src = src.filter(x => ((x.product_code||'').toLowerCase().includes(q) ||
                           (x.product_name||'').toLowerCase().includes(q)));
  }
  // Filtro Cliente (nombre, partial)
  const fc = document.getElementById('filter-client');
  if (fc && fc.value){
    const q = fc.value.toLowerCase().trim();
    src = src.filter(x => ((x.commercial_entity_name||'').toLowerCase().includes(q) ||
                           (x.partner_name||'').toLowerCase().includes(q)));
  }
  return src;
}

// Comprueba si un periodo de tarifas (YYYY | YYYY-Qn | YYYY-MM) intersecta con el rango global
function _periodBounds(p){
  if (/^\d{4}$/.test(p))           return [p+'-01-01', p+'-12-31'];
  if (/^\d{4}-Q\d$/.test(p)){
    const [y,q] = p.split('-Q');
    const m1 = (parseInt(q)-1)*3+1; const m2 = m1+2;
    return [`${y}-${String(m1).padStart(2,'0')}-01`, `${y}-${String(m2).padStart(2,'0')}-31`];
  }
  if (/^\d{4}-\d{2}$/.test(p))     return [p+'-01', p+'-31'];
  return ['',''];
}
function periodInGlobalRange(p){
  const r = globalDateRange();
  if (!r) return true;
  const [pStart, pEnd] = _periodBounds(p);
  if (r.from && pEnd   < r.from) return false;
  if (r.to   && pStart > r.to)   return false;
  return true;
}
// Filtros por defecto (se aplican al cargar y al pulsar "Limpiar filtros"):
//   state_label = "Sales Order"  -> muestra SOs confirmadas por defecto.
// Se puede cambiar desde el dropdown de la cabecera de la columna Status.
const DEFAULT_FILTERS = { state_label: "Sales Order" };
const FILTERS = { ...DEFAULT_FILTERS };
const dropdownCols = new Set([
  "state_label","customer_type","country","salesperson","salesteam",
  "partner_salesperson","partner_salesteam","currency","pricelist",
  "partner_pricelist","product_category"
]);

// =============================================================================
// Helpers formato
// =============================================================================
const fmtMoney = n => (n==null||isNaN(n)) ? "" :
  n.toLocaleString("es-ES",{minimumFractionDigits:2,maximumFractionDigits:2});
const fmtMoneyShort = n => (n==null||isNaN(n)) ? "" : (function(){
  const a=Math.abs(n);
  if (a>=1e6) return (n/1e6).toLocaleString("es-ES",{maximumFractionDigits:2})+"M";
  if (a>=1e3) return (n/1e3).toLocaleString("es-ES",{maximumFractionDigits:1})+"k";
  return n.toLocaleString("es-ES",{maximumFractionDigits:0});
})();
const fmtPct = n => (n==null||isNaN(n)) ? "" :
  (n*((Math.abs(n)<=1)?100:1)).toLocaleString("es-ES",
    {minimumFractionDigits:1,maximumFractionDigits:1}) + "%";
const fmtNum = n => (n==null||isNaN(n)) ? "" :
  n.toLocaleString("es-ES",{maximumFractionDigits:2});

function pill(state){
  const map={Quotation:"draft","Quotation Sent":"sent","Sales Order":"sale",
             Locked:"done",Cancelled:"cancel"};
  const cls = map[state] || "draft";
  return `<span class="pill pill-${cls}">${state||""}</span>`;
}
function ctypePill(v){
  const map={"Distributor":"dist","Integrador":"int","Cliente final":"end","Sin clasificar":"none"};
  const cls = map[v] || "none";
  return `<span class="pill ct-${cls}">${v||""}</span>`;
}
function fxCell(r){
  if (!r.fx_applied) return `<span class="muted">1,0000</span>`;
  const r4 = (r.fx_rate||0).toLocaleString("es-ES",{minimumFractionDigits:4,maximumFractionDigits:4});
  return `<span class="pill fx-pill" title="Tasa Odoo @ ${r.fx_rate_date||''}">${r.currency} ${r4}</span>`;
}
function escapeHtml(s){
  return (s==null?"":(""+s)).replace(/[&<>"]/g, c=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;"}[c]));
}

// =============================================================================
// TABLA
// =============================================================================
function renderHead(){
  const colRow = document.getElementById("cols-row");
  const filRow = document.getElementById("filters-row");
  colRow.innerHTML = ""; filRow.innerHTML = "";
  for (const c of COLUMNS){
    const th = document.createElement("th");
    th.className = "sortable"; th.dataset.key = c.key;
    th.style.minWidth = c.w+"px";
    th.innerHTML = `${c.label}<span class="arrow" data-arrow></span>`;
    th.addEventListener("click", () => {
      if (SORT.key === c.key) SORT.dir = -SORT.dir;
      else { SORT.key = c.key; SORT.dir = 1; }
      renderTable();
    });
    colRow.appendChild(th);

    const thf = document.createElement("th");
    thf.className = "sub";
    if (c.type === "sptipo"){
      // columna de Tipo derivado de COMMISSION_CONFIG — filtro dropdown
      const sel = document.createElement("select");
      sel.dataset.key = "sp_tipo";
      sel.innerHTML = '<option value="">(todos)</option><option value="1">Tipo 1</option><option value="2">Tipo 2</option><option value="none">(sin tipo)</option>';
      if (FILTERS["sp_tipo"]) sel.value = FILTERS["sp_tipo"];
      sel.addEventListener("change", e => { FILTERS["sp_tipo"] = e.target.value; renderTable(); });
      thf.appendChild(sel);
    } else if (dropdownCols.has(c.key)){
      const sel = document.createElement("select");
      sel.dataset.key = c.key;
      sel.innerHTML = '<option value="">(todos)</option>';
      const opts = [...new Set(DATA.map(r => r[c.key]).filter(v => v!=null && v!==""))]
                  .sort((a,b)=>(""+a).localeCompare(""+b,"es"));
      for (const v of opts) sel.innerHTML += `<option value="${escapeHtml(v)}">${escapeHtml(v)}</option>`;
      // Reflejar filtro actual (incluye defaults como state_label="Sales Order")
      if (FILTERS[c.key]) sel.value = FILTERS[c.key];
      sel.addEventListener("change", e => { FILTERS[c.key] = e.target.value; renderTable(); });
      thf.appendChild(sel);
    } else {
      const inp = document.createElement("input");
      inp.placeholder = (c.type==="num"||c.type==="money"||c.type==="pct"||c.type==="fx")
                        ? "≥ valor ó =" : "filtrar…";
      inp.dataset.key = c.key;
      if (FILTERS[c.key]) inp.value = FILTERS[c.key];
      inp.addEventListener("input", e => { FILTERS[c.key] = e.target.value; renderTable(); });
      thf.appendChild(inp);
    }
    filRow.appendChild(thf);
  }
}

function applyFilters(rows){
  const q = (document.getElementById("q").value||"").toLowerCase().trim();
  const hideSec = document.getElementById("hideSections").checked;
  const onlyCost0 = document.getElementById("onlyCostMissing").checked;
  const onlyFX = document.getElementById("onlyFX").checked;

  return rows.filter(r => {
    if (hideSec && r.is_section) return false;
    if (onlyCost0 && !r.cost_missing) return false;
    if (onlyFX && !r.fx_applied) return false;

    for (const [k,raw] of Object.entries(FILTERS)){
      if (raw === "" || raw == null) continue;
      // Filtro sintetico por Tipo de comision del SO salesperson
      if (k === "sp_tipo"){
        const t = salespersonType(r.salesperson);
        if (raw === "none" && t != null) return false;
        if (raw !== "none" && String(t) !== raw) return false;
        continue;
      }
      const v = r[k];
      const col = COLUMNS.find(c => c.key===k);
      if (!col) continue;

      if (dropdownCols.has(k)){
        if ((v??"") !== raw) return false;
        continue;
      }
      if (col.type==="num"||col.type==="money"||col.type==="pct"||col.type==="fx"){
        const m = raw.match(/^\s*(>=|<=|>|<|=)?\s*(-?[\d.,]+)\s*$/);
        if (!m){
          if (!(""+v).toLowerCase().includes(raw.toLowerCase())) return false;
          continue;
        }
        const op = m[1]||"=";
        const num = parseFloat(m[2].replace(",", "."));
        const nv = (v==null) ? NaN : parseFloat(v);
        if (isNaN(nv)) return false;
        const nvCmp = col.type==="pct" && Math.abs(nv)<=1 ? nv*100 : nv;
        if (op===">=" && !(nvCmp>=num)) return false;
        if (op==="<=" && !(nvCmp<=num)) return false;
        if (op===">"  && !(nvCmp> num)) return false;
        if (op==="<"  && !(nvCmp< num)) return false;
        if (op==="="  && !(Math.abs(nvCmp-num)<1e-6)) return false;
      } else {
        if (!(""+(v??"")).toLowerCase().includes(raw.toLowerCase())) return false;
      }
    }
    if (q){
      const hay = COLUMNS.map(c => r[c.key]).join(" | ").toLowerCase();
      if (!hay.includes(q)) return false;
    }
    return true;
  });
}

function sortRows(rows){
  const { key, dir } = SORT;
  const col = COLUMNS.find(c => c.key===key);
  const numeric = col && (col.type==="num"||col.type==="money"||col.type==="pct"||col.type==="fx");
  return [...rows].sort((a,b) => {
    let va = a[key], vb = b[key];
    if (va==null) va = numeric ? -Infinity : "";
    if (vb==null) vb = numeric ? -Infinity : "";
    if (numeric){ return (parseFloat(va)-parseFloat(vb)) * dir; }
    return (""+va).localeCompare(""+vb,"es",{numeric:true}) * dir;
  });
}

function renderRow(r){
  const tr = document.createElement("tr");
  for (const c of COLUMNS){
    const td = document.createElement("td");
    const v = r[c.key];
    if (c.type==="pill") td.innerHTML = pill(v);
    else if (c.type==="ctype") td.innerHTML = ctypePill(v);
    else if (c.type==="sptipo"){
      const t = salespersonType(r.salesperson);
      if (t) td.innerHTML = `<span class="chip-type t${t}">T${t}</span>`;
      else if (isExcludedSalesperson(r.salesperson))
        td.innerHTML = `<span class="chip-type tnone" title="Excluido — sin comisión">exc.</span>`;
      else td.innerHTML = `<span class="muted">—</span>`;
    }
    else if (c.type==="fx"){ td.className="num"; td.innerHTML = fxCell(r); }
    else if (c.type==="link"){
      td.innerHTML = r.odoo_url
        ? `<a class="odoo-link" href="${r.odoo_url}" target="_blank" rel="noopener">${escapeHtml(v||"")}</a>`
        : escapeHtml(v||"");
    }
    else if (c.type==="money"){ td.className="num"; td.textContent = fmtMoney(v); }
    else if (c.type==="pct"){ td.className="num"; td.textContent = fmtPct(v); }
    else if (c.type==="num"){ td.className="num"; td.textContent = fmtNum(v); }
    else td.textContent = v ?? "";
    // Click-to-filter para producto/cliente
    if (c.key === "product_code" && v){
      td.dataset.clickFilter = "prod"; td.dataset.fv = v;
      td.title = "Click: filtrar por este producto";
    } else if (c.key === "product_name" && v){
      td.dataset.clickFilter = "prod"; td.dataset.fv = r.product_code || v;
      td.title = "Click: filtrar por este producto";
    } else if ((c.key === "commercial_entity_name" || c.key === "partner_name") && v){
      td.dataset.clickFilter = "client"; td.dataset.fv = v;
      td.title = "Click: filtrar por este cliente";
    }
    if (c.key==="standard_price" && r.cost_missing) td.className = "num err";
    td.title = v==null ? "" : (""+v);
    tr.appendChild(td);
  }
  return tr;
}

function updateArrows(){
  for (const th of document.querySelectorAll("#cols-row th")){
    const key = th.dataset.key;
    const arrow = th.querySelector("[data-arrow]");
    if (SORT.key===key) arrow.textContent = SORT.dir>0 ? "▲" : "▼";
    else arrow.textContent = "";
  }
}

// Stats bar contextual: pinta metricas apropiadas segun la vista activa.
function renderStats(){
  const bar = document.getElementById("stats-bar");
  if (!bar) return;
  const view =
    (vTable && vTable.style.display !== "none") ? "table" :
    (vDash  && vDash.style.display  !== "none") ? "dash"  :
    (vComm  && vComm.style.display  !== "none") ? "comm"  :
    null;
  if (view === "table"){
    // Stats de la tabla (las que ya tenia), usando estado actual
    const sorted = sortRows(applyFilters(visibleData()));
    const limit = parseInt(document.getElementById("tableLimit").value, 10) || 500;
    const rendered = Math.min(limit, sorted.length);
    const rowsLbl = rendered < sorted.length ? `${rendered} / ${sorted.length}` : `${sorted.length}`;
    const sum = sorted.reduce((s,r) => s + (r.price_subtotal_eur||0), 0);
    const mar = sorted.reduce((s,r) => s + (r.margin_eur||0), 0);
    bar.innerHTML = `
      <span>Filas <b>${rowsLbl}</b> / <b>${DATA.length}</b></span>
      <span>Ofertas <b>${new Set(sorted.map(r => r.order_name)).size}</b></span>
      <span>Total <b>${fmtMoney(sum)} €</b></span>
      <span>Margen <b>${fmtMoney(mar)} €</b></span>`;
  } else if (view === "dash"){
    const rows = dashFiltered();
    const sum = rows.reduce((s,r)=>s+(r.price_subtotal_eur||0),0);
    const mar = rows.reduce((s,r)=>s+(r.margin_eur||0),0);
    const marPct = sum ? (mar/sum) : null;
    const orders = new Set(rows.map(r=>r.order_name)).size;
    bar.innerHTML = `
      <span>Líneas <b>${rows.filter(r=>!r.is_section).length}</b></span>
      <span>Ofertas <b>${orders}</b></span>
      <span>Total <b>${fmtMoney(sum)} €</b></span>
      <span>Margen <b>${fmtMoney(mar)} €</b>${marPct!=null?` (${fmtPct(marPct)})`:""}</span>`;
  } else if (view === "comm"){
    // Calcular totales agregados de comisiones visibles
    const { rows, excluded } = computeCommissionsByYear();
    const devengada = rows.reduce((s,r)=>s+r.com_total,0);
    const pagable   = rows.reduce((s,r)=>s+(r.com_paid||0),0);
    bar.innerHTML = `
      <span>Com. devengada <b>${fmtMoney(devengada)} €</b></span>
      <span>Com. PAGABLE <b style="color:#2e7d32">${fmtMoney(pagable)} €</b></span>
      <span>% cobrado <b>${devengada?fmtPct(pagable/devengada):"—"}</b></span>
      <span class="muted">Shipping excl. ${fmtMoney(excluded.shipping)} € · Controllino ${fmtMoney(excluded.controllino)} €</span>`;
  } else {
    bar.innerHTML = "";
  }
}
function updateStats(sorted, rendered){ /* legacy shim */ renderStats(); }

function renderTable(){
  // Sincronizar Mi vista global con el filtro de columna salesperson
  const my = document.getElementById('my-view');
  if (my){
    FILTERS["salesperson"] = my.value || "";
    const headerSel = document.querySelector('#filters-row select[data-key="salesperson"]');
    if (headerSel) headerSel.value = my.value || "";
  }
  const tbody = document.getElementById("tbody");
  const filtered = applyFilters(visibleData());
  const sorted = sortRows(filtered);
  const limit = parseInt(document.getElementById("tableLimit").value, 10) || 500;
  const toRender = sorted.slice(0, limit);
  tbody.innerHTML = "";
  const frag = document.createDocumentFragment();
  for (const r of toRender) frag.appendChild(renderRow(r));
  tbody.appendChild(frag);
  renderTotalsRow(sorted);
  document.getElementById("nohit").hidden = sorted.length>0;
  updateArrows();
  updateStats(sorted, toRender.length);
}

// Fila de totales acumulados — fija arriba, suma sobre TODAS las filas filtradas
// (no solo las renderizadas) para reflejar el verdadero total del filtro actual.
function renderTotalsRow(rows){
  const tr = document.getElementById("totals-row");
  if (!tr) return;
  tr.innerHTML = "";
  // Para cada columna decide qué mostrar
  for (let i=0; i<COLUMNS.length; i++){
    const c = COLUMNS[i];
    const td = document.createElement("td");
    if (i === 0){
      td.innerHTML = `<b style="color:var(--accent)">Σ TOTAL</b><br><span class="muted" style="font-size:10px;font-weight:400">${rows.length} líneas</span>`;
      td.className = 'totals-label';
      tr.appendChild(td);
      continue;
    }
    if (c.type === "money"){
      const sum = rows.reduce((s,r) => s + (parseFloat(r[c.key])||0), 0);
      td.className = "num totals-val";
      td.innerHTML = sum ? `<b>${fmtMoney(sum)}</b>` : '<span class="muted">·</span>';
    } else if (c.type === "num"){
      const sum = rows.reduce((s,r) => s + (parseFloat(r[c.key])||0), 0);
      td.className = "num totals-val";
      td.innerHTML = sum ? `<b>${fmtNum(sum)}</b>` : '<span class="muted">·</span>';
    } else if (c.type === "pct"){
      // Promedio ponderado por subtotal si existe; si no, simple media
      const vals = rows.map(r => parseFloat(r[c.key])).filter(v => !isNaN(v));
      td.className = "num totals-val";
      if (vals.length){
        const avg = vals.reduce((a,b)=>a+b,0) / vals.length;
        td.innerHTML = `<span class="muted" title="Promedio simple sobre ${vals.length} líneas">avg ${fmtPct(avg)}</span>`;
      } else { td.innerHTML = '<span class="muted">·</span>'; }
    } else {
      td.className = 'totals-val';
      td.innerHTML = '<span class="muted">·</span>';
    }
    tr.appendChild(td);
  }
}

function clearFilters(){
  // Restaurar defaults (no todo a vacio): asi "Limpiar" vuelve a
  // state_label="Sales Order" y no se pierde el default cada vez.
  for (const k of Object.keys(FILTERS)) delete FILTERS[k];
  Object.assign(FILTERS, DEFAULT_FILTERS);
  document.querySelectorAll("#filters-row input, #filters-row select")
    .forEach(el => { el.value = FILTERS[el.dataset.key] || ""; });
  document.getElementById("q").value = "";
  document.getElementById("hideSections").checked = true;  // default
  document.getElementById("onlyCostMissing").checked = false;
  document.getElementById("onlyFX").checked = false;
  renderTable();
}

function exportCSV(){
  const rows = sortRows(applyFilters(DATA));
  const cols = COLUMNS.map(c => c.key);
  const head = COLUMNS.map(c => c.label).join(";");
  const body = rows.map(r => cols.map(k => {
    const v = r[k];
    if (v==null) return "";
    const s = (""+v).replace(/"/g,'""');
    return /[;"\n]/.test(s) ? `"${s}"` : s;
  }).join(";")).join("\n");
  const csv = "\uFEFF" + head + "\n" + body;
  const blob = new Blob([csv], { type:"text/csv;charset=utf-8" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url; a.download = `DATA_SO_${new Date().toISOString().slice(0,10)}.csv`;
  a.click(); URL.revokeObjectURL(url);
}

document.getElementById("q").addEventListener("input", renderTable);
document.getElementById("hideSections").addEventListener("change", renderTable);
document.getElementById("onlyCostMissing").addEventListener("change", renderTable);
document.getElementById("onlyFX").addEventListener("change", renderTable);
document.getElementById("tableLimit").addEventListener("change", renderTable);
document.getElementById("btn-clear").addEventListener("click", clearFilters);
document.getElementById("btn-export").addEventListener("click", exportCSV);

// =============================================================================
// DASHBOARD
// =============================================================================
const CT_COLORS = {
  "Distributor":  "#8e24aa",
  "Integrador":   "#1976d2",
  "Cliente final":"#2e7d32",
  "Sin clasificar":"#90a4ae",
};
// Paleta Industrial: rojo corporativo + navy + secundarios profesionales
const PALETTE = [
  "#e30613","#1a2f5c","#1976d2","#2e7d32","#f57c00","#8e24aa",
  "#00838f","#c62828","#3949ab","#6d4c41","#ad1457","#558b2f",
  "#ef6c00","#0277bd","#7b1fa2","#00695c","#5d4037","#d84315",
];

let chartBar=null, chartDonut=null, chartTime=null, chartCT=null,
    chartCountry=null, chartSP=null;

// =============================================================================
// CONDICIONES DE COMISIONES (configuracion embebida, editable aqui mismo)
// =============================================================================
// - "type" por comercial: 1 (base 3.6%) o 2 (base 3.1%).
// - Formula por linea: rate = max(0, base − discount_pct × 0.1)  saturado en 30% dto.
// - Factor anual (solo 2026+): escalado segun ventas YTD, prorrateado mes a mes.
// Para añadir/quitar comerciales o tipos: editar COMMISSION_CONFIG.salespersons.
const COMMISSION_CONFIG = {
  // === Comerciales con tipo asignado (Fuente: Plantilla calculo comisiones - Drive) ===
  salespersons: {
    "Jordi Hernandez":          { type: 1 },
    "Garima Arora":             { type: 1 },
    "Eloi Davila Lopez":        { type: 1 },
    "Gerard Montero Martínez":  { type: 1 },
    "Josep Massó":              { type: 2 },
    "Ramon Boncompte":          { type: 2 },
    "Albert Prieto":            { type: 2 },
    "SalesPerson":              { type: 2 },  // usuario generico
  },
  // === Usuarios que NO devengan comision (Drive: "No Salesperson") ===
  excludedSalespersons: [
    "Industrial Shields - Website",
    "ADMIN",
    "Alba Sánchez Honrado",
    "Sònia Gabarró",
    "Albert Macià",
    "Abel Codina",
    "Luis Nunes",
    "Francesc Duarri",
    "Susana Guerra",
    "Joan F. Aubets - Industrial Shields",
  ],
  // === Formula estandar catalogo (Controllers/PLC/IOs/Panel PC/...) ===
  formula: {
    1: { base: 3.6, stepPerPct: 0.1, maxDiscountPct: 30 },
    2: { base: 3.1, stepPerPct: 0.1, maxDiscountPct: 30 },
  },
  // === Regla PROJECTS / PHP (codigos PHP-*, categoria Projects) ===
  //     Comision PLANA del 3% sobre subtotal de linea,
  //     independiente del descuento aplicado.
  projectRule: {
    flatRate: 3,   // %
  },
  projectMatch: {
    productCodePrefix: "PHP-",
    productCategoryContains: "Projects",
  },
  // Reglas especificas por año
  byYear: {
    2024: { factorTiers: null, note: "Sin factor por tramos. Comisión = rate × base de línea." },
    2025: { factorTiers: null, note: "Sin factor por tramos. Comisión = rate × base de línea." },
    2026: {
      factorTiers: [
        { upToAnnual: 100000,  factor: 0.5  },
        { upToAnnual: 200000,  factor: 0.65 },
        { upToAnnual: 300000,  factor: 0.75 },
        { upToAnnual: 400000,  factor: 0.85 },
        { upToAnnual: 500000,  factor: 0.9  },
        { upToAnnual: 600000,  factor: 0.95 },
        { upToAnnual: 700000,  factor: 1.0  },
        { upToAnnual: 800000,  factor: 1.05 },
        { upToAnnual: 900000,  factor: 1.1  },
        { upToAnnual: 1000000, factor: 1.15 },
        { upToAnnual: 1100000, factor: 1.2  },
        { upToAnnual: Infinity,factor: 1.25 },
      ],
      note: "Factor anual sobre la comisión base: umbrales anuales se prorratean mensualmente (umbral_mes = umbral_anual × M/12). Solo se recalculan al cerrar mes; las comisiones del mes M se liquidan con el factor vigente."
    },
  },
};

function commissionRate(type, discountPct){
  const f = COMMISSION_CONFIG.formula[type];
  if (!f) return 0;
  const d = Math.min(Math.max(discountPct ?? 0, 0), f.maxDiscountPct);
  return Math.max(0, f.base - d * f.stepPerPct);
}
function salespersonType(name){
  const e = COMMISSION_CONFIG.salespersons[name];
  return e ? e.type : null;
}
// ¿Esta el comercial explicitamente excluido (No Salesperson)?
const EXCLUDED_SP_SET = new Set(COMMISSION_CONFIG.excludedSalespersons||[]);
function isExcludedSalesperson(name){ return EXCLUDED_SP_SET.has(name); }

// ¿Es una linea de Projects/PHP? (por codigo o categoria)
function isProjectLine(row){
  const pm = COMMISSION_CONFIG.projectMatch;
  if (!pm) return false;
  const code = row.product_code || "";
  if (pm.productCodePrefix && code.startsWith(pm.productCodePrefix)) return true;
  const cat = row.product_category || "";
  if (pm.productCategoryContains && cat.includes(pm.productCategoryContains)) return true;
  return false;
}
// Comision para lineas Projects/PHP — plana, independiente del descuento
function projectCommissionRate(discountPct){
  const pr = COMMISSION_CONFIG.projectRule;
  if (!pr) return null;
  if (pr.flatRate != null) return pr.flatRate;
  if (pr.tiers){
    const d = Math.max(0, discountPct ?? 0);
    for (const t of pr.tiers){
      if (d <= t.maxDiscount) return t.commission;
    }
  }
  return 0;
}

// Mapeo campo-dataset -> id del <select> que lo filtra en dashboard
const DASH_FILTER_FIELDS = {
  customer_type:           "f-ct",
  country:                 "f-country",
  state_label:             "f-state",
  salesperson:             "f-sp",
  partner_salesperson:     "f-sp-ce",
  commercial_entity_name:  "f-client",
};

function dashFiltered(){
  const hideSec = document.getElementById("f-hideSec").checked;
  const f = {};
  for (const [field, sel] of Object.entries(DASH_FILTER_FIELDS)){
    f[field] = document.getElementById(sel).value;
  }
  // visibleData() ya aplica año global y excluidos
  return visibleData().filter(r => {
    if (hideSec && r.is_section) return false;
    for (const [field, val] of Object.entries(f)){
      if (val && r[field] !== val) return false;
    }
    return true;
  });
}

function initDashFilters(){
  // Poblar cada select con los valores unicos del dataset
  for (const [field, sel] of Object.entries(DASH_FILTER_FIELDS)){
    const el = document.getElementById(sel);
    const opts = [...new Set(DATA.map(r => r[field]).filter(Boolean))]
                .sort((a,b)=>(""+a).localeCompare(""+b,"es"));
    opts.forEach(v => el.insertAdjacentHTML("beforeend",
      `<option value="${escapeHtml(v)}">${escapeHtml(v)}</option>`));
  }
}

// ===== Click-to-filter: setear filtro de dashboard por campo+valor =====
function applyQuickFilter(field, value){
  const selId = DASH_FILTER_FIELDS[field];
  if (!selId) return;
  const el = document.getElementById(selId);
  // toggle: si ya estaba ese valor, lo quitamos
  el.value = (el.value === value) ? "" : value;
  renderDash();
}

function clearDashFilters(){
  for (const sel of Object.values(DASH_FILTER_FIELDS)){
    document.getElementById(sel).value = "";
  }
  renderDash();
}

function renderActiveFilterChips(){
  const box = document.getElementById("active-filters");
  box.innerHTML = "";
  const labels = {
    customer_type:"Cliente tipo", country:"País", state_label:"Estado",
    salesperson:"Comercial (SO)", partner_salesperson:"Comercial (CE)",
    commercial_entity_name:"Cliente",
  };
  // Chip del año global (si está filtrado)
  const yv = globalYear();
  if (yv){
    const chip = document.createElement("span");
    chip.className = "chip";
    chip.innerHTML = `Año (global): <b>${escapeHtml(yv)}</b> <span class="x">×</span>`;
    chip.addEventListener("click", () => {
      document.getElementById("global-year").value = "";
      renderAll();
    });
    box.appendChild(chip);
  }
  for (const [field, sel] of Object.entries(DASH_FILTER_FIELDS)){
    const v = document.getElementById(sel).value;
    if (!v) continue;
    const chip = document.createElement("span");
    chip.className = "chip";
    chip.innerHTML = `${labels[field]}: <b>${escapeHtml(v)}</b> <span class="x">×</span>`;
    chip.addEventListener("click", () => applyQuickFilter(field, v));
    box.appendChild(chip);
  }
}

function metricVal(r, metric){
  if (metric==="count")  return 1;
  if (metric==="orders") return r.order_name;  // se deduplica abajo
  return r[metric] || 0;
}

// Obtiene el valor de la dimension para una fila (campo real o derivado).
function dimValue(r, dim){
  if (dim === 'year')  return (r.date_order || '').slice(0,4) || '(sin fecha)';
  if (dim === 'month') return (r.date_order || '').slice(0,7) || '(sin fecha)';
  return r[dim];
}

function aggregate(rows, dim, metric){
  const bag = new Map();
  for (const r of rows){
    const k = dimValue(r, dim) || "(sin dato)";
    if (!bag.has(k)) bag.set(k, {
      group:k, n:0, qty:0, subtotal_eur:0, margin_eur:0,
      orders:new Set(), metric:0,
    });
    const b = bag.get(k);
    b.n += 1;
    b.qty += (r.qty||0);
    b.subtotal_eur += (r.price_subtotal_eur||0);
    b.margin_eur   += (r.margin_eur||0);
    if (r.order_name) b.orders.add(r.order_name);
    if (metric==="count") b.metric += 1;
    else if (metric==="orders"){ /* set size al final */ }
    else b.metric += (r[metric]||0);
  }
  const out = [...bag.values()].map(b => ({
    group: b.group,
    n: b.n,
    qty: b.qty,
    subtotal_eur: b.subtotal_eur,
    margin_eur: b.margin_eur,
    margin_pct: b.subtotal_eur ? (b.margin_eur/b.subtotal_eur) : null,
    orders: b.orders.size,
    metric: metric==="orders" ? b.orders.size : b.metric,
  }));
  out.sort((a,b) => (b.metric||0) - (a.metric||0));
  return out;
}

function renderKPIs(rows){
  const totalSub  = rows.reduce((s,r)=>s+(r.price_subtotal_eur||0),0);
  const totalMar  = rows.reduce((s,r)=>s+(r.margin_eur||0),0);
  const totalQty  = rows.reduce((s,r)=>s+(r.qty||0),0);
  const nLines    = rows.filter(r=>!r.is_section).length;
  const nOrders   = new Set(rows.map(r=>r.order_name)).size;
  const marPct    = totalSub ? (totalMar/totalSub) : null;

  // Split por customer_type
  const byCT = {};
  for (const r of rows){
    if (r.is_section) continue;
    const k = r.customer_type || "Sin clasificar";
    byCT[k] = (byCT[k]||0) + (r.price_subtotal_eur||0);
  }

  const curCT = document.getElementById("f-ct").value;
  const kpiCT = (cls, ct) => {
    const active = curCT === ct ? " active" : "";
    const val = fmtMoneyShort(byCT[ct]||0);
    const pct = totalSub ? fmtPct((byCT[ct]||0)/totalSub) : "";
    return `<div class="kpi clickable ${cls}${active}" data-ct="${ct}">
              <div class="lab">${ct}</div>
              <div class="val">${val} €</div>
              <div class="sub">${pct}</div>
            </div>`;
  };
  const kpiHtml = [
    `<div class="kpi"><div class="lab">Subtotal €</div><div class="val">${fmtMoney(totalSub)} €</div><div class="sub">Margen ${fmtPct(marPct)}</div></div>`,
    `<div class="kpi"><div class="lab">Margen €</div><div class="val">${fmtMoney(totalMar)} €</div><div class="sub">${marPct!=null?fmtPct(marPct):""} del subtotal</div></div>`,
    `<div class="kpi"><div class="lab">Ofertas</div><div class="val">${nOrders}</div><div class="sub">${nLines} líneas producto</div></div>`,
    kpiCT("dist","Distributor"),
    kpiCT("int","Integrador"),
    kpiCT("end","Cliente final"),
  ].join("");
  const kpiBox = document.getElementById("kpis");
  kpiBox.innerHTML = kpiHtml;
  // click en KPI -> filtrar por customer_type
  kpiBox.querySelectorAll(".kpi[data-ct]").forEach(el => {
    el.addEventListener("click", () => applyQuickFilter("customer_type", el.dataset.ct));
  });
}

function dimLabel(dim){
  const map = {
    customer_type:"Cliente tipo",
    commercial_entity_name:"Cliente",
    product_category:"Familia producto",
    product_name:"Producto",
    country:"País",
    salesperson:"Comercial (SO)",
    partner_salesperson:"Comercial (CE)",
    salesteam:"Sales Team",
    pricelist:"Pricelist",
    state_label:"Status",
    year:"Año",
    month:"Mes",
  };
  return map[dim] || dim;
}

function metricLabel(m){
  return {price_subtotal_eur:"Subtotal €",margin_eur:"Margen €",qty:"Cantidad",
          count:"Nº líneas",orders:"Nº ofertas"}[m] || m;
}

function makeChart(ctx, cfg){ return new Chart(ctx, cfg); }

// Handler de click en grafico -> buscar el label clicado y filtrar si es un campo filtrable
function chartClickHandler(field){
  return function(evt, elements, chart){
    if (!elements || !elements.length) return;
    const idx = elements[0].index;
    const label = chart.data.labels[idx];
    if (!label) return;
    // Si el campo es filtrable, aplicamos filtro rapido
    if (DASH_FILTER_FIELDS[field]) applyQuickFilter(field, label);
  };
}

function renderDash(){
  _syncMyViewToLocal('f-sp');
  const rows = dashFiltered();
  const dim = document.getElementById("group-by").value;
  const metric = document.getElementById("metric").value;
  const topN = parseInt(document.getElementById("topn").value, 10);

  document.getElementById("dim-label").textContent = dimLabel(dim);
  document.getElementById("dim-label2").textContent = dimLabel(dim);

  renderKPIs(rows);

  const agg = aggregate(rows, dim, metric);
  const totalMetric = agg.reduce((s,g) => s + (g.metric||0), 0);
  const totalSub    = agg.reduce((s,g) => s + (g.subtotal_eur||0), 0);
  const top = agg.slice(0, topN);

  const labels = top.map(g => g.group);
  const bar_sub = top.map(g => Math.round(g.subtotal_eur));
  const bar_mar = top.map(g => Math.round(g.margin_eur));

  // --- Chart 1: Bar subtotal + margin (horizontal) ---
  const colors = dim==="customer_type"
      ? labels.map(l => CT_COLORS[l] || "#64748b")
      : labels.map((_,i) => PALETTE[i % PALETTE.length]);
  if (chartBar) chartBar.destroy();
  chartBar = makeChart(document.getElementById("chart-bar"), {
    type:"bar",
    data:{
      labels,
      datasets:[
        { label:"Subtotal €", data:bar_sub, backgroundColor:colors, borderRadius:4 },
        { label:"Margen €",   data:bar_mar, backgroundColor:colors.map(c=>c+"80"),
          borderColor:colors, borderWidth:1, borderRadius:4 },
      ]
    },
    options:{
      indexAxis:"y",
      responsive:true, maintainAspectRatio:false,
      onClick: chartClickHandler(dim),
      plugins:{
        legend:{labels:{color:"#1a2433",font:{size:11}}},
        tooltip:{callbacks:{
          label: c => `${c.dataset.label}: ${fmtMoney(c.parsed.x)} €`
        }}
      },
      scales:{
        x:{ticks:{color:"#6b7a8c",callback:v=>fmtMoneyShort(v)},grid:{color:"#eef1f5"}},
        y:{ticks:{color:"#1a2433",autoSkip:false},grid:{color:"#eef1f5"}}
      }
    }
  });

  // --- Chart 2: Donut subtotal ---
  if (chartDonut) chartDonut.destroy();
  chartDonut = makeChart(document.getElementById("chart-donut"), {
    type:"doughnut",
    data:{labels, datasets:[{data: top.map(g=>g.subtotal_eur), backgroundColor:colors}]},
    options:{
      responsive:true, maintainAspectRatio:false,
      onClick: chartClickHandler(dim),
      plugins:{
        legend:{position:"right", labels:{color:"#1a2433",boxWidth:12,font:{size:11}}},
        tooltip:{callbacks:{
          label: c => {
            const v=c.parsed, t=totalSub||1;
            return `${c.label}: ${fmtMoney(v)} €  (${fmtPct(v/t)})`;
          }
        }}
      }
    }
  });

  // --- Chart 3: Evolucion mensual subtotal ---
  const byMonth = new Map();
  for (const r of rows){
    const d = (r.date_order||"").slice(0,7);
    if (!d) continue;
    byMonth.set(d, (byMonth.get(d)||0) + (r.price_subtotal_eur||0));
  }
  const monthLabels = [...byMonth.keys()].sort();
  const monthVals = monthLabels.map(k => Math.round(byMonth.get(k)));
  if (chartTime) chartTime.destroy();
  chartTime = makeChart(document.getElementById("chart-time"), {
    type:"bar",
    data:{labels:monthLabels, datasets:[{
      label:"Subtotal €", data:monthVals,
      backgroundColor:"#1a2f5c", borderColor:"#0f1c3a", borderWidth:1, borderRadius:3,
    }]},
    options:{
      responsive:true, maintainAspectRatio:false,
      plugins:{
        legend:{display:false},
        tooltip:{callbacks:{label:c=>`${fmtMoney(c.parsed.y)} €`}}
      },
      scales:{
        x:{ticks:{color:"#6b7a8c"},grid:{color:"#eef1f5"}},
        y:{ticks:{color:"#6b7a8c",callback:v=>fmtMoneyShort(v)},grid:{color:"#eef1f5"}}
      }
    }
  });

  // --- Chart 4: Mix customer_type ---
  const mixCT = new Map();
  for (const r of rows){
    if (r.is_section) continue;
    const k = r.customer_type || "Sin clasificar";
    mixCT.set(k, (mixCT.get(k)||0) + (r.price_subtotal_eur||0));
  }
  const ctLabels = [...mixCT.keys()];
  const ctVals = ctLabels.map(k => Math.round(mixCT.get(k)));
  if (chartCT) chartCT.destroy();
  chartCT = makeChart(document.getElementById("chart-ct"), {
    type:"pie",
    data:{labels:ctLabels, datasets:[{
      data:ctVals,
      backgroundColor:ctLabels.map(l=>CT_COLORS[l]||"#64748b"),
    }]},
    options:{
      responsive:true, maintainAspectRatio:false,
      onClick: chartClickHandler("customer_type"),
      plugins:{
        legend:{position:"right", labels:{color:"#1a2433",boxWidth:12}},
        tooltip:{callbacks:{label:c=>{
          const sum=ctVals.reduce((a,b)=>a+b,0)||1;
          return `${c.label}: ${fmtMoney(c.parsed)} €  (${fmtPct(c.parsed/sum)})`;
        }}}
      }
    }
  });

  // --- Tabla resumen por dimension (fila clickable si es filtrable) ---
  const tbody = document.querySelector("#tbl-sum tbody");
  tbody.innerHTML = "";
  const frag = document.createDocumentFragment();
  const dimIsFilterable = !!DASH_FILTER_FIELDS[dim];
  for (const g of agg){
    const pct = totalSub ? g.subtotal_eur/totalSub : 0;
    const tr = document.createElement("tr");
    if (dimIsFilterable) tr.classList.add("clickable");
    tr.innerHTML = `
      <td>${dim==="customer_type" ? ctypePill(g.group) : escapeHtml(g.group)}</td>
      <td class="num">${g.n}</td>
      <td class="num">${g.orders}</td>
      <td class="num">${fmtNum(g.qty)}</td>
      <td class="num">${fmtMoney(g.subtotal_eur)} €</td>
      <td class="num">${fmtMoney(g.margin_eur)} €</td>
      <td class="num">${g.margin_pct!=null?fmtPct(g.margin_pct):""}</td>
      <td class="num">${fmtPct(pct)}</td>`;
    if (dimIsFilterable){
      tr.addEventListener("click", () => applyQuickFilter(dim, g.group));
    }
    frag.appendChild(tr);
  }
  tbody.appendChild(frag);

  // --- Seccion PAÍS (siempre, independientemente de la dim elegida) ---
  renderCountrySection(rows);

  // --- Seccion COMERCIAL (CE) ---
  renderSalespersonSection(rows);

  // --- Resumen por CLIENTE al final ---
  renderClientsSummary(rows);

  // chips de filtros activos
  renderActiveFilterChips();

  updateStats(rows.filter(r => !r.is_section));
}

// -----------------------------------------------------------------------------
// Seccion PAÍS: grafico + tabla
// -----------------------------------------------------------------------------
function renderCountrySection(rows){
  // Agregar por country (excluyendo secciones)
  const bag = new Map();
  for (const r of rows){
    if (r.is_section) continue;
    const k = r.country || "(sin país)";
    if (!bag.has(k)) bag.set(k, {n:0, orders:new Set(), sub:0, mar:0});
    const b = bag.get(k);
    b.n += 1;
    b.orders.add(r.order_name);
    b.sub += r.price_subtotal_eur || 0;
    b.mar += r.margin_eur || 0;
  }
  const list = [...bag.entries()].map(([country, b]) => ({
    country, n:b.n, orders:b.orders.size,
    sub:b.sub, mar:b.mar,
    marPct: b.sub ? b.mar/b.sub : null,
  })).sort((a,b) => b.sub - a.sub);

  const totalSub = list.reduce((s,g)=>s+g.sub,0);
  const labels = list.map(g => g.country);
  const vals   = list.map(g => Math.round(g.sub));
  const colors = labels.map((_,i) => PALETTE[i % PALETTE.length]);

  if (chartCountry) chartCountry.destroy();
  chartCountry = makeChart(document.getElementById("chart-country"), {
    type:"bar",
    data:{labels, datasets:[{
      label:"Subtotal €", data:vals,
      backgroundColor:colors, borderRadius:4,
    }]},
    options:{
      indexAxis:"y",
      responsive:true, maintainAspectRatio:false,
      onClick: chartClickHandler("country"),
      plugins:{
        legend:{display:false},
        tooltip:{callbacks:{label:c=>{
          const t=totalSub||1;
          return `${fmtMoney(c.parsed.x)} €  (${fmtPct(c.parsed.x/t)})`;
        }}}
      },
      scales:{
        x:{ticks:{color:"#6b7a8c",callback:v=>fmtMoneyShort(v)},grid:{color:"#eef1f5"}},
        y:{ticks:{color:"#1a2433",autoSkip:false},grid:{color:"#eef1f5"}}
      }
    }
  });

  const tbody = document.querySelector("#tbl-country tbody");
  tbody.innerHTML = "";
  const frag = document.createDocumentFragment();
  for (const g of list){
    const tr = document.createElement("tr");
    tr.classList.add("clickable");
    const pct = totalSub ? g.sub/totalSub : 0;
    tr.innerHTML = `
      <td>${escapeHtml(g.country)}</td>
      <td class="num">${g.orders}</td>
      <td class="num">${g.n}</td>
      <td class="num">${fmtMoney(g.sub)} €</td>
      <td class="num">${fmtMoney(g.mar)} €</td>
      <td class="num">${g.marPct!=null?fmtPct(g.marPct):""}</td>
      <td class="num">${fmtPct(pct)}</td>`;
    tr.addEventListener("click", () => applyQuickFilter("country", g.country));
    frag.appendChild(tr);
  }
  tbody.appendChild(frag);
}

// -----------------------------------------------------------------------------
// Seccion COMERCIAL (CE): grafico + tabla (usa partner_salesperson, el
// salesperson de la entidad comercial, no el del SO)
// -----------------------------------------------------------------------------
function renderSalespersonSection(rows){
  const bag = new Map();
  for (const r of rows){
    if (r.is_section) continue;
    const k = r.partner_salesperson || "(sin asignar)";
    if (!bag.has(k)) bag.set(k, {
      sp:k, n:0, orders:new Set(), clients:new Set(),
      sub:0, mar:0,
    });
    const b = bag.get(k);
    b.n += 1;
    if (r.order_name) b.orders.add(r.order_name);
    if (r.commercial_entity_id != null) b.clients.add(r.commercial_entity_id);
    b.sub += r.price_subtotal_eur || 0;
    b.mar += r.margin_eur || 0;
  }
  const list = [...bag.values()].map(b => ({
    sp: b.sp,
    n: b.n,
    orders: b.orders.size,
    clients: b.clients.size,
    sub: b.sub, mar: b.mar,
    marPct: b.sub ? b.mar/b.sub : null,
  })).sort((a,b) => b.sub - a.sub);

  const totalSub = list.reduce((s,g)=>s+g.sub,0);
  const labels = list.map(g => g.sp);
  const vals   = list.map(g => Math.round(g.sub));
  const colors = labels.map((_,i) => PALETTE[i % PALETTE.length]);

  if (chartSP) chartSP.destroy();
  chartSP = makeChart(document.getElementById("chart-sp"), {
    type:"bar",
    data:{labels, datasets:[{
      label:"Subtotal €", data:vals,
      backgroundColor:colors, borderRadius:4,
    }]},
    options:{
      indexAxis:"y",
      responsive:true, maintainAspectRatio:false,
      onClick: chartClickHandler("partner_salesperson"),
      plugins:{
        legend:{display:false},
        tooltip:{callbacks:{label:c=>{
          const t=totalSub||1;
          return `${fmtMoney(c.parsed.x)} €  (${fmtPct(c.parsed.x/t)})`;
        }}}
      },
      scales:{
        x:{ticks:{color:"#6b7a8c",callback:v=>fmtMoneyShort(v)},grid:{color:"#eef1f5"}},
        y:{ticks:{color:"#1a2433",autoSkip:false},grid:{color:"#eef1f5"}}
      }
    }
  });

  const tbody = document.querySelector("#tbl-sp tbody");
  tbody.innerHTML = "";
  const frag = document.createDocumentFragment();
  for (const g of list){
    const tr = document.createElement("tr");
    tr.classList.add("clickable");
    const pct = totalSub ? g.sub/totalSub : 0;
    // Tipo de comision (si el comercial aparece en COMMISSION_CONFIG)
    const t = salespersonType(g.sp);
    let tipoCell, baseCell;
    if (t){
      tipoCell = `<span class="chip-type t${t}">Tipo ${t}</span>`;
      baseCell = `${COMMISSION_CONFIG.formula[t].base}%`;
    } else if (isExcludedSalesperson(g.sp)){
      tipoCell = `<span class="chip-type tnone" title="Excluido — sin comisión">Sin com.</span>`;
      baseCell = `<span class="muted">—</span>`;
    } else {
      tipoCell = `<span class="chip-type tnone" title="Sin tipo definido">—</span>`;
      baseCell = `<span class="muted">—</span>`;
    }
    tr.innerHTML = `
      <td>${escapeHtml(g.sp)}</td>
      <td>${tipoCell}</td>
      <td class="num">${baseCell}</td>
      <td class="num">${g.clients}</td>
      <td class="num">${g.orders}</td>
      <td class="num">${g.n}</td>
      <td class="num">${fmtMoney(g.sub)} €</td>
      <td class="num">${fmtMoney(g.mar)} €</td>
      <td class="num">${g.marPct!=null?fmtPct(g.marPct):""}</td>
      <td class="num">${fmtPct(pct)}</td>`;
    tr.addEventListener("click", () => applyQuickFilter("partner_salesperson", g.sp));
    frag.appendChild(tr);
  }
  tbody.appendChild(frag);
}

// -----------------------------------------------------------------------------
// Resumen por CLIENTE (entidad comercial)
// -----------------------------------------------------------------------------
function renderClientsSummary(rows){
  const bag = new Map();
  for (const r of rows){
    if (r.is_section) continue;
    const cid = r.commercial_entity_id;
    if (cid == null) continue;
    if (!bag.has(cid)) bag.set(cid, {
      id: cid,
      name: r.commercial_entity_name,
      ctype: r.customer_type,
      country: r.country,
      pricelist: r.partner_pricelist || r.pricelist,
      salesperson: r.partner_salesperson || r.salesperson,
      orders: new Set(),
      n: 0,
      sub: 0,
      mar: 0,
      lastDate: "",
    });
    const b = bag.get(cid);
    b.n += 1;
    b.orders.add(r.order_name);
    b.sub += r.price_subtotal_eur || 0;
    b.mar += r.margin_eur || 0;
    if ((r.date_order||"") > b.lastDate) b.lastDate = r.date_order || "";
    // Si hay mezcla, preferir el ctype del partner con mas peso: nos quedamos con el que aparece, simple
  }
  const list = [...bag.values()].map(b => ({
    ...b,
    ordersN: b.orders.size,
    marPct: b.sub ? b.mar/b.sub : null,
    avg: b.orders.size ? b.sub/b.orders.size : 0,
  })).sort((a,b) => b.sub - a.sub);

  const tbody = document.querySelector("#tbl-clients tbody");
  tbody.innerHTML = "";
  const frag = document.createDocumentFragment();
  for (const c of list){
    const tr = document.createElement("tr");
    tr.classList.add("clickable");
    tr.innerHTML = `
      <td>${escapeHtml(c.name||"")}</td>
      <td>${ctypePill(c.ctype)}</td>
      <td>${escapeHtml(c.country||"")}</td>
      <td class="muted" title="${escapeHtml(c.pricelist||"")}">${escapeHtml((c.pricelist||"").slice(0,30))}</td>
      <td>${escapeHtml(c.salesperson||"")}</td>
      <td class="num">${c.ordersN}</td>
      <td class="num">${c.n}</td>
      <td class="num">${fmtMoney(c.sub)} €</td>
      <td class="num">${fmtMoney(c.mar)} €</td>
      <td class="num">${c.marPct!=null?fmtPct(c.marPct):""}</td>
      <td class="num">${fmtMoney(c.avg)} €</td>
      <td>${escapeHtml(c.lastDate||"")}</td>`;
    tr.addEventListener("click", () => applyQuickFilter("commercial_entity_name", c.name));
    frag.appendChild(tr);
  }
  tbody.appendChild(frag);
}

function exportDashCSV(){
  const rows = dashFiltered();
  const dim = document.getElementById("group-by").value;
  const metric = document.getElementById("metric").value;
  const agg = aggregate(rows, dim, metric);
  const totalSub = agg.reduce((s,g)=>s+(g.subtotal_eur||0),0);
  const head = ["Grupo","N_lineas","N_ofertas","Cantidad","Subtotal_EUR","Margen_EUR","Margen_pct","Pct_s_total"].join(";");
  const body = agg.map(g => [
    g.group, g.n, g.orders, g.qty,
    (g.subtotal_eur||0).toFixed(2),
    (g.margin_eur||0).toFixed(2),
    g.margin_pct!=null?(g.margin_pct*100).toFixed(2):"",
    totalSub?((g.subtotal_eur/totalSub)*100).toFixed(2):"",
  ].map(v=>{
    const s=(""+v).replace(/"/g,'""');
    return /[;"\n]/.test(s)?`"${s}"`:s;
  }).join(";")).join("\n");
  const csv="\uFEFF"+head+"\n"+body;
  const blob=new Blob([csv],{type:"text/csv;charset=utf-8"});
  const url=URL.createObjectURL(blob);
  const a=document.createElement("a");
  a.href=url; a.download=`DATA_SO_grouped_${dim}_${new Date().toISOString().slice(0,10)}.csv`;
  a.click(); URL.revokeObjectURL(url);
}

// Wire up (f-year removido — ahora es global en la cabecera)
["group-by","metric","topn","f-ct","f-country","f-state","f-sp","f-sp-ce","f-client"].forEach(id =>
  document.getElementById(id).addEventListener("change", renderDash));
document.getElementById("f-hideSec").addEventListener("change", renderDash);
document.getElementById("btn-dash-csv").addEventListener("click", exportDashCSV);
document.getElementById("btn-dash-reset").addEventListener("click", clearDashFilters);

// =============================================================================
// VISTA COMISIONES (calculo efectivo por comercial y año)
// =============================================================================
// Reglas aplicadas:
//  - Solo SO confirmadas (state=sale) — los drafts/sent 2026 no devengan.
//  - Shipping (category/name con "Shipping") y Controllino (cat con "Controllino") NO pagan.
//  - Usuarios excluidos (Website, ADMIN, Alba, Sònia, Abel, Macià, etc.) no pagan.
//  - PHP/Projects (codigo PHP-* o cat "Projects"): 3% plano.
//  - Resto catalogo: Tipo 1 (3.6%) o Tipo 2 (3.1%) menos 0.1 por cada 1% dto.
//  - No se aplica el factor anual 2026 (pendiente).
function classifyLineForCommission(r){
  const cat = r.product_category || "";
  const name = r.product_name || "";
  const code = r.product_code || "";
  if (cat.includes("Shipping") || name.includes("Shipping")) return "shipping";
  if (cat.includes("Controllino")) return "controllino";
  if (code.startsWith("PHP-")) return "php";
  if (cat.includes("Projects")) return "php";
  return "catalog";
}

let COMM_YEAR = null;   // estado: año activo en la pestaña

function computeCommissionsByYear(){
  // Por (year, sp) acumulamos com devengada total Y pagable (solo si SO cobrada).
  // Para no doble-contar (varias lineas de la misma SO), primero agrupamos por SO.
  const bag = new Map();  // key: year|sp
  const excluded = { shipping:0, controllino:0, excluded_sp:0, excluded_orders:new Set() };

  const src = visibleData();
  for (const r of src){
    if (r.is_section) continue;
    if (r.state !== 'sale') continue;  // solo SO confirmadas
    const d = r.date_order || "";
    if (!d || !/^\d{4}/.test(d)) continue;
    const year = parseInt(d.slice(0,4), 10);
    const sp = r.salesperson || "(sin asignar)";
    const sub = r.price_subtotal_eur || 0;
    const cls = classifyLineForCommission(r);

    // Agregados globales de excluidos (para mostrar resumen)
    if (cls === "shipping")    excluded.shipping    += sub;
    if (cls === "controllino") excluded.controllino += sub;
    if (isExcludedSalesperson(sp)) {
      excluded.excluded_sp += sub;
      if (r.order_name) excluded.excluded_orders.add(r.order_name);
    }

    const k = year+"|"+sp;
    if (!bag.has(k)) bag.set(k, {
      year, sp, type:null, status:"",
      orders:new Set(), orders_paid:new Set(), n_lines:0,
      sub_catalog:0, sub_php:0, sub_controllino:0, sub_shipping:0,
      com_catalog:0, com_php:0,
      com_catalog_paid:0, com_php_paid:0,  // solo cobradas
    });
    const b = bag.get(k);
    b.n_lines += 1;
    if (r.order_name) b.orders.add(r.order_name);
    const isPaid = (r.payment_state_agg === 'paid');
    if (isPaid && r.order_name) b.orders_paid.add(r.order_name);

    // subtotales por clase (visibilidad)
    if (cls === "shipping")    b.sub_shipping    += sub;
    else if (cls === "controllino") b.sub_controllino += sub;
    else if (cls === "php")    b.sub_php         += sub;
    else                       b.sub_catalog     += sub;

    // Comision: excluir SP excluidos
    if (isExcludedSalesperson(sp)) { b.status = "Excluido"; continue; }
    const t = salespersonType(sp);
    b.type = t;
    if (cls === "shipping" || cls === "controllino") continue;  // no paga
    if (cls === "php") {
      const c = sub * (COMMISSION_CONFIG.projectRule.flatRate / 100);
      b.com_php += c;
      if (isPaid) b.com_php_paid += c;
    } else {  // catalog
      if (!t) { b.status = "Sin tipo"; continue; }
      // Descuento EFECTIVO contra tarifa deducida (max(declarado, deducido))
      const eff = effectiveDiscount(r);
      const rate = commissionRate(t, eff.use);
      const c = sub * (rate / 100);
      b.com_catalog += c;
      if (isPaid) b.com_catalog_paid += c;
    }
  }

  const rows = [...bag.values()].map(b => ({
    ...b,
    orders: b.orders.size,
    orders_paid: b.orders_paid.size,
    com_total: b.com_catalog + b.com_php,
    com_paid:  b.com_catalog_paid + b.com_php_paid,
  }));
  const years = [...new Set(rows.map(r=>r.year))].sort();
  excluded.excluded_orders = excluded.excluded_orders.size;
  return { years, rows, excluded };
}

function renderCommissionsTabs(years){
  if (COMM_YEAR == null) COMM_YEAR = "aggregate";  // por defecto, agregado
  const tabs = document.getElementById("comm-tabs");
  // Si hay año global seleccionado, las pestañas son redundantes -> ocultas
  if (globalYear()){
    tabs.style.display = "none";
    tabs.innerHTML = "";
    return;
  }
  tabs.style.display = "";
  tabs.innerHTML = "";
  const mk = (v, label) => {
    const b = document.createElement("button");
    b.textContent = label;
    if (v === COMM_YEAR) b.classList.add("active");
    b.addEventListener("click", () => { COMM_YEAR = v; renderCommissions(); });
    return b;
  };
  tabs.appendChild(mk("aggregate","Agregado (todos los años)"));
  for (const y of years) tabs.appendChild(mk(String(y), String(y)));
}

function renderCommissionsTable(rows, excluded){
  // Añadir métricas derivadas necesarias para sort (ventas comisionables = catalog + php)
  for (const r of rows){ r._comisionable = (r.sub_catalog||0) + (r.sub_php||0); }
  // Ordenar segun estado de orden de la tabla
  const so = TBL_SORT.comm_sum;
  _sortBy(rows, so.key, so.dir);
  const totals = rows.reduce((t,r) => ({
    sub_catalog:t.sub_catalog+r.sub_catalog, sub_php:t.sub_php+r.sub_php,
    sub_controllino:t.sub_controllino+r.sub_controllino, sub_shipping:t.sub_shipping+r.sub_shipping,
    com_catalog:t.com_catalog+r.com_catalog, com_php:t.com_php+r.com_php, com_total:t.com_total+r.com_total,
  }), {sub_catalog:0, sub_php:0, sub_controllino:0, sub_shipping:0, com_catalog:0, com_php:0, com_total:0});

  const typeChip = r => {
    if (r.status === "Excluido") return `<span class="chip-type tnone">exc.</span>`;
    if (r.status === "Sin tipo") return `<span class="chip-type tnone">—</span>`;
    if (r.type)  return `<span class="chip-type t${r.type}">T${r.type}</span>`;
    return `<span class="chip-type tnone">—</span>`;
  };

  const totCompaid = rows.reduce((s,r)=>s+(r.com_paid||0), 0);
  const totOrdersPaid = rows.reduce((s,r)=>s+(r.orders_paid||0), 0);
  const tbl = `
    <table class="tbl-sum" style="width:100%">
      <thead><tr>
        ${_sortHead('comm_sum','sp','Comercial',{defaultDir:1})}
        ${_sortHead('comm_sum','type','Tipo',{tip:"Tipo 1 (3,6%) o Tipo 2 (3,1%). −0,1 puntos por cada 1% dto, saturado a 30%.",defaultDir:1})}
        ${_sortHead('comm_sum','orders','Ofertas',{num:true,tip:"SOs únicas confirmadas en el periodo."})}
        ${_sortHead('comm_sum','orders_paid','Cobradas',{num:true,tip:"SOs con todas las facturas pagadas. Liquidables."})}
        ${_sortHead('comm_sum','sub_catalog','Ventas catálogo €',{num:true,tip:"Subtotal de líneas de catálogo estándar (Controllers, IOs, Panel PC…). Excluye Shipping, Controllino y PHP."})}
        ${_sortHead('comm_sum','sub_php','Ventas PHP €',{num:true,tip:"Subtotal de líneas PHP-* / Projects (3% plano)."})}
        ${_sortHead('comm_sum','_comisionable','Comisionable €',{num:true,tip:"Catálogo + PHP. Base sobre la que aplica comisión."})}
        ${_sortHead('comm_sum','com_catalog','Com. catálogo €',{num:true})}
        ${_sortHead('comm_sum','com_php','Com. PHP €',{num:true})}
        ${_sortHead('comm_sum','com_total','Com. devengada €',{num:true,tip:"Comisión total generada (independiente de cobro)."})}
        ${_sortHead('comm_sum','com_paid','Com. PAGABLE €',{num:true,tip:"Comisión cuya SO está totalmente cobrada. Liquidable hoy."})}
        ${_sortHead('comm_sum','status','Estado',{defaultDir:1})}
      </tr></thead>
      <tbody>
        ${rows.map(r => {
          const comBase = r.sub_catalog + r.sub_php;
          return `
          <tr>
            <td>${escapeHtml(r.sp)}</td>
            <td>${typeChip(r)}</td>
            <td class="num">${r.orders}</td>
            <td class="num">${r.orders_paid||0}</td>
            <td class="num">${fmtMoney(r.sub_catalog)}</td>
            <td class="num">${fmtMoney(r.sub_php)}</td>
            <td class="num"><b>${fmtMoney(comBase)}</b></td>
            <td class="num">${fmtMoney(r.com_catalog)}</td>
            <td class="num">${fmtMoney(r.com_php)}</td>
            <td class="num">${fmtMoney(r.com_total)}</td>
            <td class="num"><b style="color:#2e7d32">${fmtMoney(r.com_paid||0)}</b></td>
            <td>${r.status ? `<span class="muted" style="font-size:11px">${r.status}</span>` : ""}</td>
          </tr>`;}).join("")}
      </tbody>
      <tfoot>
        <tr style="border-top:2px solid var(--accent);font-weight:600">
          <td colspan="2">TOTAL</td>
          <td class="num">${rows.reduce((s,r)=>s+r.orders,0)}</td>
          <td class="num">${totOrdersPaid}</td>
          <td class="num">${fmtMoney(totals.sub_catalog)}</td>
          <td class="num">${fmtMoney(totals.sub_php)}</td>
          <td class="num"><b>${fmtMoney(totals.sub_catalog + totals.sub_php)}</b></td>
          <td class="num">${fmtMoney(totals.com_catalog)}</td>
          <td class="num">${fmtMoney(totals.com_php)}</td>
          <td class="num">${fmtMoney(totals.com_total)}</td>
          <td class="num"><b style="color:#2e7d32">${fmtMoney(totCompaid)}</b></td>
          <td></td>
        </tr>
      </tfoot>
    </table>`;

  const exclBlock = excluded ? `
    <div class="card" style="margin-top:16px">
      <h3>Volumen excluido <span class="muted">(sobre el cual NO hay comisión)</span></h3>
      <div style="display:grid;grid-template-columns:repeat(auto-fit,minmax(220px,1fr));gap:12px">
        <div class="kpi"><div class="lab">Shipping</div><div class="val">${fmtMoney(excluded.shipping)} €</div></div>
        <div class="kpi"><div class="lab">Controllino</div><div class="val">${fmtMoney(excluded.controllino)} €</div><div class="sub">pendiente regla aparte</div></div>
        <div class="kpi"><div class="lab">Usuarios excluidos</div><div class="val">${fmtMoney(excluded.excluded_sp)} €</div><div class="sub">${excluded.excluded_orders} ofertas</div></div>
      </div>
    </div>` : "";
  return `<div class="card">${tbl}</div>${exclBlock}`;
}

// -----------------------------------------------------------------------------
// Calculo por pedido (para la vista detalle)
// -----------------------------------------------------------------------------
function computeCommissionsByOrder(filterSp, filterYear){
  // Devuelve lista de pedidos con subtotales clasificados, comisiones y % efectivos.
  const map = new Map();  // key: order_name
  for (const r of visibleData()){
    if (r.is_section) continue;
    if (r.state !== 'sale') continue;
    const d = r.date_order || "";
    if (!d) continue;
    const year = parseInt(d.slice(0,4), 10);
    if (filterYear && year !== filterYear) continue;
    const sp = r.salesperson || "(sin asignar)";
    if (filterSp && sp !== filterSp) continue;

    const oid = r.order_name;
    if (!map.has(oid)) map.set(oid, {
      order_name: oid,
      odoo_url: r.odoo_url,
      date_order: r.date_order,
      year,
      sp,
      type: salespersonType(sp),
      status: isExcludedSalesperson(sp) ? "Excluido" : (salespersonType(sp) ? "" : "Sin tipo"),
      customer: r.commercial_entity_name,
      country: r.country,
      total: 0,
      total_native: 0,             // suma en moneda original (sin convertir)
      currency:      r.currency || "EUR",
      fx_rate:       r.fx_rate || null,
      fx_rate_date:  r.fx_rate_date || null,
      fx_applied:    !!r.fx_applied,
      sub_catalog: 0, sub_php: 0, sub_controllino: 0, sub_shipping: 0,
      com_catalog: 0, com_php: 0,
      // Para descuento medio ponderado (sobre lineas comisionables)
      _disc_w_num: 0,  // Σ (discount_pct × subtotal)
      _disc_w_den: 0,  // Σ (subtotal)
      n_lines: 0,
      // Facturacion/cobro (agregado SO, viene propagado en todas las lineas)
      invoice_status:    r.invoice_status,
      payment_state_agg: r.payment_state_agg,
      invoiced_amount:   r.invoiced_amount,
      residual_amount:   r.residual_amount,
      refunded_amount:   r.refunded_amount,
      n_invoices:        r.n_invoices,
      invoice_names:     r.invoice_names,
      last_invoice_date: r.last_invoice_date,
    });
    const o = map.get(oid);
    o.n_lines += 1;
    const sub = r.price_subtotal_eur || 0;
    o.total += sub;
    o.total_native += (r.price_subtotal || 0);
    const cls = classifyLineForCommission(r);
    if (cls === "shipping")    o.sub_shipping    += sub;
    else if (cls === "controllino") o.sub_controllino += sub;
    else if (cls === "php")    o.sub_php         += sub;
    else                       o.sub_catalog     += sub;

    // Descuento medio ponderado: usa el EFECTIVO (max(declarado, deducido))
    // Solo para lineas catalog (PHP no tiene tarifa deducida fiable y va a 3% plano)
    if ((cls === "catalog" || cls === "php") && sub > 0){
      const eff = effectiveDiscount(r);
      o._disc_w_num += eff.use * sub;
      o._disc_w_den += sub;
    }

    // Comision
    if (isExcludedSalesperson(sp)) continue;
    if (cls === "shipping" || cls === "controllino") continue;
    if (cls === "php"){
      o.com_php += sub * (COMMISSION_CONFIG.projectRule.flatRate / 100);
    } else {
      const t = salespersonType(sp);
      if (!t) continue;
      // Descuento EFECTIVO contra tarifa deducida (max(declarado, deducido))
      const eff = effectiveDiscount(r);
      const rate = commissionRate(t, eff.use);
      o.com_catalog += sub * (rate / 100);
      // Si hay sospecha (precio_unit < tarifa pero declaró 0 dto), marcamos el pedido
      if (eff.suspect) o.has_suspect = true;
    }
  }
  const out = [...map.values()].map(o => ({
    ...o,
    commissionable: o.sub_catalog + o.sub_php,
    com_total: o.com_catalog + o.com_php,
    avg_discount: o._disc_w_den ? (o._disc_w_num / o._disc_w_den) : 0,
    effective_rate: (o.sub_catalog + o.sub_php) ? ((o.com_catalog + o.com_php) / (o.sub_catalog + o.sub_php) * 100) : 0,
  }));
  // Limpieza
  out.forEach(o => { delete o._disc_w_num; delete o._disc_w_den; });
  out.sort((a,b) => (b.date_order||"").localeCompare(a.date_order||"") || a.order_name.localeCompare(b.order_name));
  return out;
}

function renderOrdersTable(orders){
  // Aplicar orden segun TBL_SORT.comm_det
  const so = TBL_SORT.comm_det;
  // Campo derivado para sort por estado de cobro y "Com. pagable"
  for (const o of orders){
    o._payable = (o.payment_state_agg === 'paid') ? o.com_total : 0;
    o._pay_state = o.payment_state_agg || 'none';
  }
  _sortBy(orders, so.key, so.dir);
  const totals = orders.reduce((t,o)=>({
    total:t.total+o.total, sub_catalog:t.sub_catalog+o.sub_catalog, sub_php:t.sub_php+o.sub_php,
    sub_controllino:t.sub_controllino+o.sub_controllino, sub_shipping:t.sub_shipping+o.sub_shipping,
    commissionable:t.commissionable+o.commissionable,
    com_catalog:t.com_catalog+o.com_catalog, com_php:t.com_php+o.com_php, com_total:t.com_total+o.com_total,
  }), {total:0, sub_catalog:0, sub_php:0, sub_controllino:0, sub_shipping:0, commissionable:0, com_catalog:0, com_php:0, com_total:0});
  const totAvgDisc = totals.commissionable ? (orders.reduce((s,o)=>s+o.avg_discount*o.commissionable,0) / totals.commissionable) : 0;
  const totEffRate = totals.commissionable ? (totals.com_total / totals.commissionable * 100) : 0;

  // Chips de cobro
  const payChip = ps => {
    const map = {
      paid:        ['Cobrado','t1'],
      partial:     ['Parcial','tnone'],
      not_paid:    ['No cobrado','tnone'],
      to_invoice:  ['A facturar','tnone'],
      none:        ['Sin factura','tnone'],
    };
    const [lbl, cls] = map[ps||'none'] || ['—','tnone'];
    return `<span class="chip-type ${cls}">${lbl}</span>`;
  };
  // Comision pagable: solo si cobrada
  const payableCom = o => (o.payment_state_agg === 'paid') ? o.com_total : 0;
  const totPayable = orders.reduce((s,o)=>s+payableCom(o), 0);

  return `
    <div class="card" style="margin-top:16px">
      <h3>Detalle por pedido <span class="muted">(${orders.length} pedidos)</span>
          <span style="margin-left:auto;font-weight:400;font-size:11px;color:var(--muted)">
            · Pagable (solo cobradas): <b style="color:#2e7d32">${fmtMoney(totPayable)} €</b>
          </span>
      </h3>
      <div style="max-height:600px;overflow:auto">
        <table class="tbl-sum" style="width:100%;min-width:1800px">
          <thead><tr>
            ${_sortHead('comm_det','order_name','SO',{defaultDir:1,tip:"Número de pedido. Clic abre el pedido en Odoo."})}
            ${_sortHead('comm_det','date_order','Fecha',{defaultDir:-1})}
            ${_sortHead('comm_det','sp','Comercial',{defaultDir:1})}
            ${_sortHead('comm_det','customer','Cliente',{defaultDir:1})}
            ${_sortHead('comm_det','currency','Cur',{defaultDir:1,tip:"Moneda original. EUR = sin conversión."})}
            ${_sortHead('comm_det','total_native','Total (orig)',{num:true,tip:"Subtotal en moneda original."})}
            ${_sortHead('comm_det','fx_rate','FX rate',{num:true,tip:"Tasa Odoo as-of-date. EUR = orig / rate."})}
            ${_sortHead('comm_det','fx_rate_date','Fecha FX',{defaultDir:-1})}
            ${_sortHead('comm_det','total','Total €',{num:true})}
            ${_sortHead('comm_det','commissionable','Con comisión €',{num:true,tip:"Catálogo + PHP. Excluye shipping y Controllino."})}
            ${_sortHead('comm_det','avg_discount','Dto efectivo',{num:true,tip:"Descuento medio efectivo ponderado por subtotal."})}
            ${_sortHead('comm_det','effective_rate','% comisión',{num:true,tip:"Tasa efectiva = Com. TOTAL / Con comisión."})}
            ${_sortHead('comm_det','sub_shipping','Shipping €',{num:true})}
            ${_sortHead('comm_det','sub_controllino','Controllino €',{num:true})}
            ${_sortHead('comm_det','com_catalog','Com. catálogo €',{num:true})}
            ${_sortHead('comm_det','com_php','Com. PHP €',{num:true})}
            ${_sortHead('comm_det','com_total','Com. TOTAL €',{num:true,tip:"Com. catálogo + Com. PHP."})}
            ${_sortHead('comm_det','invoiced_amount','Facturado',{num:true,tip:"Importe ya facturado (€)."})}
            ${_sortHead('comm_det','_pay_state','Cobrado',{defaultDir:1,tip:"Estado de cobro: paid / partial / not_paid / to_invoice / none."})}
            ${_sortHead('comm_det','_payable','Com. pagable €',{num:true,tip:"Liquidable hoy = Com. TOTAL si totalmente cobrado, si no 0."})}
          </tr></thead>
          <tbody>
            ${orders.map(o => {
              const invLbl = (o.n_invoices || 0) === 0
                ? (o.invoice_status === 'to invoice'
                    ? `<span class="chip-type tnone">A facturar</span>`
                    : `<span class="muted">—</span>`)
                : `<span class="muted" title="${escapeHtml((o.invoice_names||[]).join(', '))}">${o.n_invoices} fac · ${fmtMoney(o.invoiced_amount||0)} €</span>`;
              const pagable = payableCom(o);
              const fxLabel = o.fx_applied
                ? `<span class="pill fx-pill" title="Tasa res.currency.rate @ ${escapeHtml(o.fx_rate_date||"")}">${escapeHtml(o.currency)} ${o.fx_rate ? o.fx_rate.toLocaleString("es-ES",{minimumFractionDigits:4,maximumFractionDigits:4}) : ""}</span>`
                : `<span class="muted">—</span>`;
              const suspectIcon = o.has_suspect
                  ? ` <span title="Hay líneas con precio bajo la tarifa deducida y descuento declarado bajo: el cálculo de comisión usa el descuento EFECTIVO." style="color:#f57c00;font-weight:700">⚠</span>`
                  : "";
              return `
              <tr${o.has_suspect?' style="background:#fff8e1"':''}>
                <td>${o.odoo_url
                    ? `<a class="odoo-link" href="${o.odoo_url}" target="_blank" rel="noopener">${escapeHtml(o.order_name)}</a>${suspectIcon}`
                    : escapeHtml(o.order_name) + suspectIcon}</td>
                <td>${escapeHtml((o.date_order||"").slice(0,10))}</td>
                <td>${escapeHtml(o.sp)}</td>
                <td data-click-filter="client" data-fv="${escapeHtml(o.customer||"")}" title="Click: filtrar por este cliente">${escapeHtml(o.customer||"")}</td>
                <td>${o.currency==="EUR" ? `<span class="muted">EUR</span>` : `<span class="pill fx-pill">${escapeHtml(o.currency)}</span>`}</td>
                <td class="num">${o.currency==="EUR" ? `<span class="muted">—</span>` : fmtMoney(o.total_native)}</td>
                <td class="num">${fxLabel}</td>
                <td>${o.fx_applied ? `<span class="muted">${escapeHtml(o.fx_rate_date||"")}</span>` : `<span class="muted">—</span>`}</td>
                <td class="num">${fmtMoney(o.total)}</td>
                <td class="num"><b>${fmtMoney(o.commissionable)}</b></td>
                <td class="num">${o.commissionable>0 ? fmtPct(o.avg_discount/100) : ""}</td>
                <td class="num">${o.commissionable>0 ? fmtPct(o.effective_rate/100) : ""}</td>
                <td class="num muted">${o.sub_shipping    ? fmtMoney(o.sub_shipping)    : ""}</td>
                <td class="num muted">${o.sub_controllino ? fmtMoney(o.sub_controllino) : ""}</td>
                <td class="num">${o.com_catalog ? fmtMoney(o.com_catalog) : ""}</td>
                <td class="num">${o.com_php     ? fmtMoney(o.com_php)     : ""}</td>
                <td class="num"><b>${fmtMoney(o.com_total)}</b></td>
                <td>${invLbl}</td>
                <td>${payChip(o.payment_state_agg)}</td>
                <td class="num"><b style="color:${pagable>0?'#86efac':'var(--muted)'}">${pagable>0?fmtMoney(pagable):'—'}</b></td>
              </tr>`;}).join("")}
          </tbody>
          <tfoot>
            <tr style="border-top:2px solid var(--accent);font-weight:600">
              <td colspan="4">TOTAL (${orders.length})</td>
              <td colspan="4" class="muted" style="font-size:11px">— FX columns —</td>
              <td class="num">${fmtMoney(totals.total)}</td>
              <td class="num"><b>${fmtMoney(totals.commissionable)}</b></td>
              <td class="num">${totals.commissionable ? fmtPct(totAvgDisc/100) : ""}</td>
              <td class="num">${totals.commissionable ? fmtPct(totEffRate/100) : ""}</td>
              <td class="num muted">${fmtMoney(totals.sub_shipping)}</td>
              <td class="num muted">${fmtMoney(totals.sub_controllino)}</td>
              <td class="num">${fmtMoney(totals.com_catalog)}</td>
              <td class="num">${fmtMoney(totals.com_php)}</td>
              <td class="num"><b>${fmtMoney(totals.com_total)}</b></td>
              <td></td>
              <td></td>
              <td class="num"><b style="color:#2e7d32">${fmtMoney(totPayable)}</b></td>
            </tr>
          </tfoot>
        </table>
      </div>
    </div>`;
}

function populateCommSpSelect(){
  const sel = document.getElementById("comm-sp");
  if (!sel || sel.dataset.populated) return;
  sel.dataset.populated = "1";
  const sps = [...new Set(DATA.map(r => r.salesperson).filter(Boolean))]
                .filter(sp => !EXCLUDED_SP_SET.has(sp))
                .sort((a,b)=>a.localeCompare(b,"es"));
  for (const sp of sps){
    sel.insertAdjacentHTML("beforeend",
      `<option value="${escapeHtml(sp)}">${escapeHtml(sp)}</option>`);
  }
  sel.addEventListener("change", renderCommissions);
  document.getElementById("comm-view").addEventListener("change", renderCommissions);
  document.getElementById("comm-metric").addEventListener("change", renderCommissions);
  document.getElementById("comm-detail").addEventListener("change", renderCommissions);
  document.getElementById("comm-q").addEventListener("input", renderCommissions);
  document.getElementById("comm-min").addEventListener("input", renderCommissions);
  document.getElementById("comm-max").addEventListener("input", renderCommissions);
  document.getElementById("comm-pay").addEventListener("change", renderCommissions);
  document.getElementById("btn-comm-csv").addEventListener("click", exportCommissionsCSV);
  document.getElementById("btn-comm-reset").addEventListener("click", () => {
    document.getElementById("comm-sp").value = "";
    document.getElementById("comm-detail").checked = true;
    document.getElementById("comm-q").value = "";
    document.getElementById("comm-min").value = "";
    document.getElementById("comm-max").value = "";
    document.getElementById("comm-pay").value = "";
    renderCommissions();
  });
}

// Aplica los filtros de busqueda y rango de comision a un array de pedidos
function applyOrderFilters(orders){
  const q   = (document.getElementById("comm-q").value   || "").trim().toLowerCase();
  const min = document.getElementById("comm-min").value; const minN = min==="" ? null : parseFloat(min);
  const max = document.getElementById("comm-max").value; const maxN = max==="" ? null : parseFloat(max);
  const pay = document.getElementById("comm-pay").value;
  return orders.filter(o => {
    if (q && !o.order_name.toLowerCase().includes(q)) return false;
    if (minN != null && !(o.com_total >= minN)) return false;
    if (maxN != null && !(o.com_total <= maxN)) return false;
    if (pay && (o.payment_state_agg || 'none') !== pay) return false;
    return true;
  });
}

function exportCommissionsCSV(){
  const sp = document.getElementById("comm-sp").value;
  const y  = (COMM_YEAR==="aggregate") ? null : parseInt(COMM_YEAR, 10);
  const orders = applyOrderFilters(computeCommissionsByOrder(sp || null, y));
  const head = ["SO","Fecha","Comercial","Cliente","Cur","Total_orig","FX_rate","Fecha_FX","Total_EUR","Con_comision_EUR","Dto_medio_pct","Pct_comision_efectivo","Shipping_EUR","Controllino_EUR","Com_catalogo_EUR","Com_PHP_EUR","Com_TOTAL_EUR","N_facturas","Facturado_EUR","Pagado_estado","Ultima_factura","Com_pagable_EUR"].join(";");
  const body = orders.map(o => [
    o.order_name, (o.date_order||"").slice(0,10), o.sp, o.customer||"",
    o.currency || 'EUR',
    (o.total_native||0).toFixed(2),
    o.fx_rate ? o.fx_rate.toFixed(4) : '',
    o.fx_rate_date || '',
    o.total.toFixed(2), o.commissionable.toFixed(2),
    o.avg_discount.toFixed(2), o.effective_rate.toFixed(2),
    o.sub_shipping.toFixed(2), o.sub_controllino.toFixed(2),
    o.com_catalog.toFixed(2), o.com_php.toFixed(2), o.com_total.toFixed(2),
    o.n_invoices || 0,
    (o.invoiced_amount||0).toFixed(2),
    o.payment_state_agg || 'none',
    o.last_invoice_date || '',
    (o.payment_state_agg==='paid' ? o.com_total : 0).toFixed(2),
  ].map(v => {
    const s = (""+v).replace(/"/g,'""');
    return /[;"\n]/.test(s) ? `"${s}"` : s;
  }).join(";")).join("\n");
  const csv = "\uFEFF" + head + "\n" + body;
  const blob = new Blob([csv], { type:"text/csv;charset=utf-8" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = `comisiones_${sp||'todos'}_${COMM_YEAR}_${new Date().toISOString().slice(0,10)}.csv`;
  a.click();
  URL.revokeObjectURL(url);
}

// =============================================================================
// VISTAS COMISIONES: detalle MENSUAL + acumulado YOY
// =============================================================================
// Devuelve métrica (devengada / pagable / ventas) por linea
function _commMetricForLine(r, metric){
  if (metric === 'ventas'){
    if (r.is_section || r.state !== 'sale') return 0;
    const sp = r.salesperson; if (!sp || EXCLUDED_SP_SET.has(sp)) return 0;
    const cat = r.product_category || '', nm = r.product_name || '';
    if (cat.includes('Shipping') || nm.includes('Shipping')) return 0;
    if (cat.includes('Controllino')) return 0;
    return r.price_subtotal_eur || 0;
  }
  const com = _commissionForLine(r);
  if (com <= 0) return 0;
  if (metric === 'pagable'){
    return (r.payment_state_agg === 'paid') ? com : 0;
  }
  return com; // devengada
}

// Recolecta datos por (mes 'YYYY-MM', sp) usando date_order de la línea
function _commByMonthSp(metric){
  const m = new Map();   // key 'YYYY-MM|sp' -> num
  const months = new Set();
  const sps = new Set();
  for (const r of visibleData()){
    const v = _commMetricForLine(r, metric);
    if (v <= 0) continue;
    const d = (r.date_order || '').slice(0,10);
    if (d.length < 7) continue;
    const month = d.slice(0,7);
    const sp = r.salesperson;
    months.add(month); sps.add(sp);
    const key = month + '|' + sp;
    m.set(key, (m.get(key) || 0) + v);
  }
  return { byKey: m,
    months: [...months].sort(),
    sps: [...sps].sort((a,b)=>a.localeCompare(b,'es')),
  };
}

// Vista MENSUAL: pivote (filas=comerciales, columnas=meses)
function _renderCommMonthly(){
  const metric = document.getElementById('comm-metric').value || 'devengada';
  const spFilter = document.getElementById('comm-sp').value || '';
  const { byKey, months, sps } = _commByMonthSp(metric);
  const rowSps = spFilter ? sps.filter(s => s === spFilter) : sps;

  if (months.length === 0 || rowSps.length === 0){
    document.getElementById('comm-body').innerHTML =
      `<div class='card'><p class='muted'>Sin datos en el rango actual. Prueba a ampliar el filtro de Periodo.</p></div>`;
    return;
  }

  const monthLabel = (m) => {
    const [y, mm] = m.split('-');
    const names = ['Ene','Feb','Mar','Abr','May','Jun','Jul','Ago','Sep','Oct','Nov','Dic'];
    return `${names[parseInt(mm)-1]}'${y.slice(2)}`;
  };
  const monthTitle = (m) => `Mes ${m} (${monthLabel(m)})`;

  // Totales por mes (col) y por comercial (fila)
  const colTotals = months.map(mo => rowSps.reduce((s,sp) => s + (byKey.get(mo+'|'+sp) || 0), 0));
  const rowTotals = rowSps.map(sp => months.reduce((s,mo) => s + (byKey.get(mo+'|'+sp) || 0), 0));
  const grand = rowTotals.reduce((a,b)=>a+b, 0);
  const maxCell = Math.max(...months.flatMap(mo => rowSps.map(sp => byKey.get(mo+'|'+sp) || 0)));

  const cellHtml = (v) => {
    if (!v) return `<td class="num muted">·</td>`;
    const intensity = maxCell ? Math.min(1, v / maxCell) : 0;
    const bg = `rgba(46,125,50,${0.07 + intensity * 0.30})`;
    return `<td class="num" style="background:${bg}"><b>${fmtMoney(v)}</b></td>`;
  };

  const metricLabel = metric==='pagable' ? 'PAGABLE' : metric==='ventas' ? 'Ventas comisionables' : 'Devengada';

  const html = `
    <div class="card">
      <h3>Detalle mensual · <span style="color:var(--accent)">${escapeHtml(metricLabel)}</span>
        <span class="muted">${months.length} meses · ${rowSps.length} comercial${rowSps.length===1?'':'es'} · TOTAL ${fmtMoney(grand)} €</span>
      </h3>
      <div style="overflow:auto;max-height:calc(100vh - 300px)">
        <table class="tbl-sum tbl-monthly">
          <thead><tr>
            <th class="pv-sticky pv-code" style="left:0;min-width:200px">Comercial</th>
            ${months.map(mo => `<th class="num" title="${escapeHtml(monthTitle(mo))}">${escapeHtml(monthLabel(mo))}</th>`).join("")}
            <th class="num" style="border-left:2px solid var(--line-strong);background:#f8fafc">TOTAL</th>
          </tr></thead>
          <tbody>
            ${rowSps.map((sp,i) => `
              <tr>
                <td class="pv-sticky pv-code" style="left:0">${escapeHtml(sp)}</td>
                ${months.map(mo => cellHtml(byKey.get(mo+'|'+sp) || 0)).join("")}
                <td class="num" style="border-left:2px solid var(--line-strong);background:#f8fafc"><b>${fmtMoney(rowTotals[i])} €</b></td>
              </tr>`).join("")}
          </tbody>
          <tfoot>
            <tr style="border-top:2px solid var(--accent);font-weight:600">
              <td class="pv-sticky pv-code" style="left:0;background:var(--bg-alt)">TOTAL mes</td>
              ${colTotals.map(v => `<td class="num"><b>${fmtMoney(v)}</b></td>`).join("")}
              <td class="num" style="border-left:2px solid var(--line-strong);background:var(--bg-alt)"><b>${fmtMoney(grand)} €</b></td>
            </tr>
          </tfoot>
        </table>
      </div>
    </div>`;
  document.getElementById('comm-body').innerHTML = html;
}

// Vista YOY: filas=meses (Ene..Dic), columnas=años; valores acumulados desde Ene
let _commYoYChart = null;
function _renderCommYoY(){
  const metric = document.getElementById('comm-metric').value || 'devengada';
  const spFilter = document.getElementById('comm-sp').value || '';

  // Datos por (año, mes) limitados al spFilter si lo hay
  const yearMap = new Map(); // year -> [12 months]
  const yearsSet = new Set();
  for (const r of visibleData()){
    const sp = r.salesperson;
    if (spFilter && sp !== spFilter) continue;
    const v = _commMetricForLine(r, metric);
    if (v <= 0) continue;
    const d = (r.date_order || '').slice(0,10);
    if (d.length < 7) continue;
    const yr = d.slice(0,4);
    const mo = parseInt(d.slice(5,7), 10) - 1;
    yearsSet.add(yr);
    if (!yearMap.has(yr)) yearMap.set(yr, new Array(12).fill(0));
    yearMap.get(yr)[mo] += v;
  }
  const years = [...yearsSet].sort();
  if (years.length === 0){
    document.getElementById('comm-body').innerHTML =
      `<div class='card'><p class='muted'>Sin datos en el rango actual.</p></div>`;
    return;
  }
  // Acumulados: [m0, m0+m1, m0+m1+m2, ...]
  const cumByYear = {};
  for (const y of years){
    const arr = yearMap.get(y) || new Array(12).fill(0);
    const cum = []; let s = 0;
    for (let i=0; i<12; i++){ s += arr[i]; cum.push(s); }
    cumByYear[y] = { monthly: arr, cum };
  }

  const monthNames = ['Ene','Feb','Mar','Abr','May','Jun','Jul','Ago','Sep','Oct','Nov','Dic'];
  const metricLabel = metric==='pagable' ? 'PAGABLE' : metric==='ventas' ? 'Ventas comisionables' : 'Devengada';

  // Tabla: rows=mes, columns=year, valor=acumulado, + Δ% YoY (comparando con año previo)
  const yoyDelta = (y, mIdx) => {
    const idx = years.indexOf(y);
    if (idx <= 0) return null;
    const prev = years[idx-1];
    const prevV = cumByYear[prev].cum[mIdx];
    const curV  = cumByYear[y].cum[mIdx];
    if (!prevV) return null;
    return (curV - prevV) / prevV;
  };

  let rows = "";
  for (let m=0; m<12; m++){
    rows += `<tr>
      <td><b>${monthNames[m]}</b></td>
      ${years.map((y,i) => {
        const c = cumByYear[y].cum[m];
        const mo = cumByYear[y].monthly[m];
        const dlt = yoyDelta(y, m);
        const dltTxt = (dlt == null) ? '<span class="muted">—</span>'
          : `<span style="color:${dlt>0?'#2e7d32':dlt<0?'#c62828':'#5b6b7c'};font-size:11px">${dlt>0?'+':''}${(dlt*100).toFixed(1)}%</span>`;
        return `<td class="num" title="Mes: ${fmtMoney(mo)} €">
          <b>${fmtMoney(c)} €</b>
          ${i>0 ? `<br>${dltTxt}` : ''}
        </td>`;
      }).join("")}
    </tr>`;
  }
  // Total año
  rows += `<tr style="border-top:2px solid var(--accent);font-weight:600;background:var(--bg-alt)">
    <td>Total año</td>
    ${years.map((y,i) => {
      const total = cumByYear[y].cum[11];
      const dlt = yoyDelta(y, 11);
      const dltTxt = (dlt == null) ? ''
        : ` <span style="color:${dlt>0?'#2e7d32':dlt<0?'#c62828':'#5b6b7c'};font-size:11px">(${dlt>0?'+':''}${(dlt*100).toFixed(1)}%)</span>`;
      return `<td class="num"><b>${fmtMoney(total)} €</b>${i>0 ? dltTxt : ''}</td>`;
    }).join("")}
  </tr>`;

  const html = `
    <div class="card">
      <h3>Acumulado anual (YoY) · <span style="color:var(--accent)">${escapeHtml(metricLabel)}</span>
        ${spFilter ? `<span class="muted">· ${escapeHtml(spFilter)}</span>` : ''}
      </h3>
      <div style="display:grid;grid-template-columns:1fr;gap:16px">
        <div style="height:260px;position:relative">
          <canvas id="comm-yoy-chart"></canvas>
        </div>
        <div style="overflow:auto">
          <table class="tbl-sum">
            <thead><tr>
              <th>Mes</th>
              ${years.map(y => `<th class="num">${y}<br><span class="muted" style="font-size:10px;font-weight:400">acumulado · YoY%</span></th>`).join("")}
            </tr></thead>
            <tbody>${rows}</tbody>
          </table>
        </div>
      </div>
    </div>`;
  document.getElementById('comm-body').innerHTML = html;

  // Render chart con Chart.js
  try {
    if (_commYoYChart) { _commYoYChart.destroy(); _commYoYChart = null; }
    const palette = ['#1976d2','#e30613','#2e7d32','#ed6c02','#7b1fa2','#0097a7'];
    const ctx = document.getElementById('comm-yoy-chart').getContext('2d');
    _commYoYChart = new Chart(ctx, {
      type: 'line',
      data: {
        labels: monthNames,
        datasets: years.map((y, i) => ({
          label: y,
          data: cumByYear[y].cum,
          borderColor: palette[i % palette.length],
          backgroundColor: palette[i % palette.length] + '20',
          fill: false,
          tension: 0.25,
          borderWidth: 2,
          pointRadius: 3,
        })),
      },
      options: {
        maintainAspectRatio: false,
        plugins: {
          legend: { position: 'top', labels:{font:{size:11}} },
          tooltip: { callbacks: { label: (ctx) => `${ctx.dataset.label}: ${fmtMoney(ctx.parsed.y)} €` } },
        },
        scales: {
          y: { ticks: { callback: (v) => fmtMoneyShort(v)+' €', font:{size:11} } },
          x: { ticks: { font:{size:11} } },
        },
      },
    });
  } catch(e){ console.warn('YoY chart error', e); }
}

function renderCommissions(){
  populateCommSpSelect();
  _syncMyViewToLocal('comm-sp');
  const view = document.getElementById('comm-view').value || 'summary';
  if (view === 'monthly') return _renderCommMonthly();
  if (view === 'yoy')     return _renderCommYoY();
  const { years, rows, excluded } = computeCommissionsByYear();
  // Si hay año global, forzamos la vista a "aggregate" (los datos ya vienen filtrados)
  if (globalYear()) COMM_YEAR = "aggregate";
  renderCommissionsTabs(years);

  const body = document.getElementById("comm-body");
  const headerNote = `
    <div class="rules-note" style="margin-bottom:12px">
      <b>Reglas aplicadas:</b> solo SO confirmadas (state=sale). Shipping y Controllino NO pagan comisión.
      Usuarios excluidos marcados con <span class="chip-type tnone">exc.</span>. Catálogo: Tipo 1/Tipo 2 (rate − 0.1 por % dto, saturado a 30%).
      PHP/Projects: 3% plano. <b>No</b> se aplica factor anual 2026.
    </div>`;

  const spFilter = document.getElementById("comm-sp").value || "";
  const showDetail = document.getElementById("comm-detail").checked || !!spFilter;

  let summary = "";
  if (COMM_YEAR === "aggregate"){
    const byS = new Map();
    for (const r of rows){
      if (spFilter && r.sp !== spFilter) continue;
      if (!byS.has(r.sp)) byS.set(r.sp, {
        year:"all", sp:r.sp, type:r.type, status:r.status,
        orders:0, orders_paid:0, n_lines:0,
        sub_catalog:0, sub_php:0, sub_controllino:0, sub_shipping:0,
        com_catalog:0, com_php:0, com_total:0, com_paid:0,
      });
      const a = byS.get(r.sp);
      a.orders += r.orders; a.orders_paid += (r.orders_paid||0); a.n_lines += r.n_lines;
      a.sub_catalog += r.sub_catalog; a.sub_php += r.sub_php;
      a.sub_controllino += r.sub_controllino; a.sub_shipping += r.sub_shipping;
      a.com_catalog += r.com_catalog; a.com_php += r.com_php; a.com_total += r.com_total;
      a.com_paid += (r.com_paid||0);
      if (r.type) a.type = r.type;
      if (r.status) a.status = r.status;
    }
    summary = renderCommissionsTable([...byS.values()], spFilter ? null : excluded);
  } else {
    const y = parseInt(COMM_YEAR, 10);
    const yr = rows.filter(r => r.year === y && (!spFilter || r.sp === spFilter));
    summary = renderCommissionsTable(yr, spFilter ? null : excluded);
  }

  // Si hay busqueda o rango de comision, forzamos el detalle
  const hasOrderFilters = !!document.getElementById("comm-q").value
                       || document.getElementById("comm-min").value !== ""
                       || document.getElementById("comm-max").value !== "";

  let detail = "";
  if (showDetail || hasOrderFilters){
    const y = (COMM_YEAR === "aggregate") ? null : parseInt(COMM_YEAR, 10);
    const orders = applyOrderFilters(computeCommissionsByOrder(spFilter || null, y));
    detail = renderOrdersTable(orders);
  }

  body.innerHTML = headerNote + summary + detail;
}

// =============================================================================
// VISTA PAGO COMISIONES
// =============================================================================
// Reglas:
//  - Solo desde 2025-Q3 (julio 2025).
//  - 2025-Q3 y 2025-Q4 -> trimestrales.
//  - 2026+ -> mensuales (YYYY-MM).
//  - Cada linea de producto contribuye a:
//     · Generado en el periodo de date_order (cuando se confirma la SO)
//     · Facturado en el periodo de last_invoice_date (si invoice_status invoiced/parcial)
//     · Cobrado en el periodo de last_invoice_date (si payment_state_agg = paid)
//  - Solo SO confirmadas (state=sale). Excluye Shipping, Controllino y SP excluidos.
function _periodForPayment(date_str){
  if (!date_str) return null;
  const yr = parseInt(date_str.slice(0,4), 10);
  const mo = parseInt(date_str.slice(5,7) || "1", 10);
  if (isNaN(yr)) return null;
  if (yr < 2025) return null;
  if (yr === 2025){
    if (mo < 7) return null;
    const q = mo <= 9 ? 3 : 4;
    return "2025-Q" + q;
  }
  return yr + "-" + String(mo).padStart(2,"0");
}

function _commissionForLine(r){
  // Devuelve comision EUR para esta linea (catalog/php), o 0 si no aplica.
  if (r.is_section || r.state !== 'sale') return 0;
  const sp = r.salesperson;
  if (!sp || isExcludedSalesperson(sp)) return 0;
  const cls = classifyLineForCommission(r);
  if (cls === 'shipping' || cls === 'controllino') return 0;
  const sub = r.price_subtotal_eur || 0;
  if (cls === 'php') return sub * (COMMISSION_CONFIG.projectRule.flatRate / 100);
  // catalog
  const t = salespersonType(sp);
  if (!t) return 0;
  const eff = effectiveDiscount(r);
  const rate = commissionRate(t, eff.use);
  return sub * (rate / 100);
}

function computePayments(){
  // Devuelve {byPair: Map(period|sp -> {generated,invoiced,collected,...}),
  //           lines: [{period,sp,...}], periods: Set, sps: Set}
  const byPair = new Map();
  const lines = [];
  const periods = new Set();
  const sps = new Set();
  const get = (period, sp) => {
    const k = period + "|" + sp;
    if (!byPair.has(k)) byPair.set(k, {
      period, sp, type: salespersonType(sp), generated:0, invoiced:0, collected:0,
      ventas:0, n_gen:0, n_inv:0, n_col:0,
    });
    return byPair.get(k);
  };

  for (const r of visibleData()){
    const com = _commissionForLine(r);
    if (com <= 0) continue;
    const sp = r.salesperson;
    sps.add(sp);
    // Volumen de venta comisionable (sin Shipping ni Controllino — los excluye _commissionForLine)
    const sub = r.price_subtotal_eur || 0;

    const pGen = _periodForPayment(r.date_order);
    if (pGen){
      const b = get(pGen, sp);
      b.generated += com;
      b.ventas += sub;
      b.n_gen += 1;
      periods.add(pGen);
    }
    const pInv = _periodForPayment(r.last_invoice_date);
    if (pInv){
      const isInvoiced = ['invoiced','to invoice','upselling'].includes(r.invoice_status) || (r.n_invoices||0)>0;
      const isPaid = (r.payment_state_agg === 'paid');
      if (isInvoiced){
        const b = get(pInv, sp);
        b.invoiced += com;
        b.n_inv += 1;
        periods.add(pInv);
      }
      if (isPaid){
        const b = get(pInv, sp);
        b.collected += com;
        b.n_col += 1;
        periods.add(pInv);
        // Para detalle por linea
        lines.push({
          period: pInv, sp, type: salespersonType(sp),
          order_name: r.order_name, odoo_url: r.odoo_url,
          date_order: r.date_order, last_invoice_date: r.last_invoice_date,
          customer: r.commercial_entity_name,
          product_code: r.product_code, product_name: r.product_name,
          qty: r.qty, sub: r.price_subtotal_eur || 0,
          discount_pct: r.discount_pct || 0,
          payment_state_agg: r.payment_state_agg,
          commission: com,
        });
      }
    }
  }
  return { byPair, lines, periods, sps };
}

function _periodOrder(p){
  // Q3-2025 < Q4-2025 < 2026-01 < 2026-02 ...
  if (/^\d{4}-Q\d$/.test(p)){
    const [y,q] = p.split('-Q');
    return parseInt(y)*100 + (parseInt(q)*3-2);  // Q3 -> 7, Q4 -> 10
  }
  if (/^\d{4}-\d{2}$/.test(p)){
    return parseInt(p.slice(0,4))*100 + parseInt(p.slice(5,7));
  }
  return 0;
}

function _populatePaySelectors(){
  if (document.getElementById("pay-sp").dataset.populated) return;
  document.getElementById("pay-sp").dataset.populated = "1";
  const { sps, periods } = computePayments();
  const spArr = [...sps].sort((a,b)=>a.localeCompare(b,"es"));
  const sel = document.getElementById("pay-sp");
  spArr.forEach(s => sel.insertAdjacentHTML("beforeend",
    `<option value="${escapeHtml(s)}">${escapeHtml(s)}</option>`));
  const psel = document.getElementById("pay-period");
  const pArr = [...periods].sort((a,b)=>_periodOrder(a)-_periodOrder(b));
  pArr.forEach(p => psel.insertAdjacentHTML("beforeend",
    `<option value="${escapeHtml(p)}">${escapeHtml(p)}</option>`));
  document.getElementById("pay-sp").addEventListener("change", renderPayments);
  document.getElementById("pay-period").addEventListener("change", renderPayments);
  document.getElementById("pay-detail").addEventListener("change", renderPayments);
  document.getElementById("pay-view").addEventListener("change", renderPayments);
  document.getElementById("pay-from").addEventListener("change", renderPayments);
  document.getElementById("pay-to").addEventListener("change", renderPayments);
  document.getElementById("btn-pay-csv").addEventListener("click", exportPaymentsCSV);
  const bSync = document.getElementById("btn-pay-sync");
  if (bSync) bSync.addEventListener("click", syncPagosFromDrive);
  const fSync = document.getElementById("pay-sync-file");
  if (fSync) fSync.addEventListener("change", _onPaySyncFile);
}

// Convierte un periodo de PAGO ('YYYY-Qn' | 'YYYY-MM') al primer mes 'YYYY-MM'
function _payPeriodToMonth(p){
  if (/^\d{4}-Q\d$/.test(p)){
    const [y,q] = p.split('-Q');
    const m = (parseInt(q)-1)*3 + 1;
    return `${y}-${String(m).padStart(2,'0')}`;
  }
  return p; // ya es YYYY-MM
}
function _payPeriodInMonthRange(p, from, to){
  // Permite que un trimestre Q se incluya si ALGUNO de sus 3 meses cae dentro [from, to]
  if (!from && !to) return true;
  let pStart, pEnd;
  if (/^\d{4}-Q\d$/.test(p)){
    const [y,q] = p.split('-Q');
    const m1 = (parseInt(q)-1)*3 + 1; const m2 = m1+2;
    pStart = `${y}-${String(m1).padStart(2,'0')}`;
    pEnd   = `${y}-${String(m2).padStart(2,'0')}`;
  } else {
    pStart = p; pEnd = p;
  }
  if (from && pEnd   < from) return false;
  if (to   && pStart > to)   return false;
  return true;
}

function renderPayments(){
  _populatePaySelectors();
  _syncMyViewToLocal('pay-sp');
  const fSp = document.getElementById("pay-sp").value;
  const fPeriod = document.getElementById("pay-period").value;
  const view = document.getElementById("pay-view").value || 'byperiod';
  const showDetail = document.getElementById("pay-detail").checked || !!fSp;
  const monthFrom = document.getElementById("pay-from").value || '';
  const monthTo   = document.getElementById("pay-to").value   || '';

  const { byPair, lines } = computePayments();
  let rows = [...byPair.values()];
  if (fSp) rows = rows.filter(r => r.sp === fSp);
  if (view === 'byperiod' && fPeriod) rows = rows.filter(r => r.period === fPeriod);
  // En vista acumulada, aplica el rango de meses pay-from/pay-to
  if (view === 'cum' && (monthFrom || monthTo)){
    rows = rows.filter(r => _payPeriodInMonthRange(r.period, monthFrom, monthTo));
  }
  rows.sort((a,b) => _periodOrder(a.period) - _periodOrder(b.period) || a.sp.localeCompare(b.sp));

  // Para cada fila, mira el registro Drive de pagos y calcula Pendiente
  for (const r of rows){
    const pg = pagosFor(r.period, r.sp);
    r.pagado    = pg.pagado;
    r.fecha_pago = pg.fecha;
    r.notas     = pg.notas;
    r.pendiente = Math.max(0, +(r.collected - r.pagado).toFixed(2));
    r.saldado   = (r.collected > 0) && (r.pagado + 0.005 >= r.collected);
  }

  // Si vista=cum, agregar por sp
  let displayRows = rows;
  if (view === 'cum'){
    const byS = new Map();
    for (const r of rows){
      if (!byS.has(r.sp)) byS.set(r.sp, {
        sp: r.sp, type: r.type, period: '', periodList: new Set(),
        generated:0, invoiced:0, collected:0, ventas:0,
        pagado:0, pendiente:0, n_gen:0, n_inv:0, n_col:0,
        fecha_pago:'', notas:'',
      });
      const a = byS.get(r.sp);
      a.generated += r.generated; a.invoiced += r.invoiced; a.collected += r.collected;
      a.ventas    += r.ventas    || 0;
      a.pagado    += r.pagado    || 0;
      a.n_gen     += r.n_gen;    a.n_inv += r.n_inv;       a.n_col += r.n_col;
      a.periodList.add(r.period);
    }
    displayRows = [...byS.values()];
    for (const a of displayRows){
      a.pendiente = Math.max(0, +(a.collected - a.pagado).toFixed(2));
      a.saldado   = (a.collected > 0) && (a.pagado + 0.005 >= a.collected);
      a.period    = `${[...a.periodList].sort((x,y)=>_periodOrder(x)-_periodOrder(y)).join(', ')}`;
    }
    displayRows.sort((a,b) => a.sp.localeCompare(b.sp));
  }

  // Aplicar orden de la tabla resumen
  const ssum = TBL_SORT.pay_sum;
  for (const r of displayRows){ r._periodOrd = _periodOrder(r.period.split(',')[0].trim()); }
  _sortBy(displayRows, ssum.key, ssum.dir);

  const tot = displayRows.reduce((t,r)=>({
    v: t.v+(r.ventas||0),
    g: t.g+r.generated, i: t.i+r.invoiced, c: t.c+r.collected,
    p: t.p+(r.pagado||0), pend: t.pend+(r.pendiente||0),
    ng: t.ng+r.n_gen, ni: t.ni+r.n_inv, nc: t.nc+r.n_col,
  }),{v:0,g:0,i:0,c:0,p:0,pend:0,ng:0,ni:0,nc:0});

  const typeChip = (t) => t ? `<span class="chip-type t${t}">T${t}</span>` : `<span class="chip-type tnone">—</span>`;
  const periodColLabel = view==='cum' ? 'Periodos incluidos' : 'Periodo';

  const summary = `
    <div class="card">
      <h3>${view==='cum' ? 'Liquidación acumulada' : 'Liquidación por periodo y comercial'}
        <span class="muted">(${displayRows.length} filas${view==='cum' && (monthFrom||monthTo) ? ` · meses ${monthFrom||'…'} → ${monthTo||'…'}`:''})</span>
      </h3>
      <div style="max-height:600px;overflow:auto">
        <table class="tbl-sum" style="width:100%;min-width:1500px">
          <thead><tr>
            ${_sortHead('pay_sum','_periodOrd',periodColLabel,{defaultDir:1,tip:"2025-Q3 / 2025-Q4 trimestrales · 2026-MM mensuales."})}
            ${_sortHead('pay_sum','sp','Comercial',{defaultDir:1})}
            ${_sortHead('pay_sum','type','Tipo',{defaultDir:1})}
            ${_sortHead('pay_sum','ventas','Ventas €',{num:true,tip:"Volumen de venta comisionable. Excluye Shipping y Controllino."})}
            ${_sortHead('pay_sum','generated','Generado €',{num:true,tip:"Comisión generada por SOs confirmadas en este periodo."})}
            ${_sortHead('pay_sum','n_gen','Lns gen',{num:true})}
            ${_sortHead('pay_sum','invoiced','Facturado €',{num:true,tip:"Comisión de SOs facturados en este periodo."})}
            ${_sortHead('pay_sum','n_inv','Lns fact',{num:true})}
            ${_sortHead('pay_sum','collected','COBRADO €',{num:true,tip:"Comisión que CORRESPONDE PAGAR (SO totalmente cobrado)."})}
            ${_sortHead('pay_sum','n_col','Lns cob',{num:true})}
            ${_sortHead('pay_sum','pagado','Pagado €',{num:true,tip:"Importe ya liquidado (lectura del Sheet en Drive)."})}
            ${_sortHead('pay_sum','pendiente','Pendiente €',{num:true,tip:"COBRADO − Pagado."})}
            ${_sortHead('pay_sum','fecha_pago','Fecha pago',{defaultDir:-1})}
            <th>Notas</th>
          </tr></thead>
          <tbody>
            ${displayRows.map(r => `
              <tr class="${r.saldado ? 'pay-ok' : (r.pendiente>0 ? 'pay-pend' : '')}">
                <td><b>${escapeHtml(r.period)}</b></td>
                <td>${escapeHtml(r.sp)}</td>
                <td>${typeChip(r.type)}</td>
                <td class="num">${fmtMoney(r.ventas||0)} €</td>
                <td class="num">${fmtMoney(r.generated)} €</td>
                <td class="num muted">${r.n_gen}</td>
                <td class="num">${fmtMoney(r.invoiced)} €</td>
                <td class="num muted">${r.n_inv}</td>
                <td class="num"><b style="color:#2e7d32">${fmtMoney(r.collected)} €</b></td>
                <td class="num muted">${r.n_col}</td>
                <td class="num">${fmtMoney(r.pagado||0)} €</td>
                <td class="num"><b style="color:${r.pendiente>0?'#c62828':'#2e7d32'}">${fmtMoney(r.pendiente||0)} €</b></td>
                <td>${escapeHtml(r.fecha_pago||"")}</td>
                <td class="muted" style="max-width:200px;font-size:11px">${escapeHtml(r.notas||"")}</td>
              </tr>`).join("")}
          </tbody>
          <tfoot>
            <tr style="border-top:2px solid var(--accent);font-weight:600">
              <td colspan="3">TOTAL</td>
              <td class="num"><b>${fmtMoney(tot.v)} €</b></td>
              <td class="num">${fmtMoney(tot.g)} €</td>
              <td class="num muted">${tot.ng}</td>
              <td class="num">${fmtMoney(tot.i)} €</td>
              <td class="num muted">${tot.ni}</td>
              <td class="num"><b style="color:#2e7d32">${fmtMoney(tot.c)} €</b></td>
              <td class="num muted">${tot.nc}</td>
              <td class="num">${fmtMoney(tot.p)} €</td>
              <td class="num"><b style="color:${tot.pend>0?'#c62828':'#2e7d32'}">${fmtMoney(tot.pend)} €</b></td>
              <td colspan="2"></td>
            </tr>
          </tfoot>
        </table>
      </div>
    </div>`;

  let detail = "";
  if (showDetail){
    let det = lines.slice();
    if (fSp) det = det.filter(l => l.sp === fSp);
    if (fPeriod) det = det.filter(l => l.period === fPeriod);
    // Aplicar orden de la tabla detalle
    for (const l of det){ l._periodOrd = _periodOrder(l.period); }
    const sdet = TBL_SORT.pay_det;
    _sortBy(det, sdet.key, sdet.dir);
    const totDet = det.reduce((s,l)=>s+l.commission, 0);
    detail = `
      <div class="card" style="margin-top:16px">
        <h3>Detalle por línea cobrada <span class="muted">(${det.length} líneas)</span>
            <span style="margin-left:auto;font-weight:400;font-size:11px;color:var(--muted)">
              · Total a pagar: <b style="color:#2e7d32">${fmtMoney(totDet)} €</b>
            </span>
        </h3>
        <div style="max-height:560px;overflow:auto">
          <table class="tbl-sum" style="width:100%;min-width:1500px">
            <thead><tr>
              ${_sortHead('pay_det','_periodOrd','Periodo cobro',{defaultDir:1})}
              ${_sortHead('pay_det','sp','Comercial',{defaultDir:1})}
              ${_sortHead('pay_det','order_name','SO',{defaultDir:1})}
              ${_sortHead('pay_det','date_order','Fecha SO',{defaultDir:-1})}
              ${_sortHead('pay_det','last_invoice_date','Fecha factura',{defaultDir:-1})}
              ${_sortHead('pay_det','customer','Cliente',{defaultDir:1})}
              ${_sortHead('pay_det','product_code','Cód. producto',{defaultDir:1})}
              ${_sortHead('pay_det','qty','Qty',{num:true})}
              ${_sortHead('pay_det','sub','Subtotal €',{num:true})}
              ${_sortHead('pay_det','discount_pct','Dto %',{num:true})}
              ${_sortHead('pay_det','commission','Comisión €',{num:true})}
            </tr></thead>
            <tbody>
              ${det.slice(0, 2000).map(l => `
                <tr>
                  <td><b>${escapeHtml(l.period)}</b></td>
                  <td>${escapeHtml(l.sp)}</td>
                  <td>${l.odoo_url
                    ? `<a class="odoo-link" href="${l.odoo_url}" target="_blank" rel="noopener">${escapeHtml(l.order_name||"")}</a>`
                    : escapeHtml(l.order_name||"")}</td>
                  <td>${escapeHtml((l.date_order||"").slice(0,10))}</td>
                  <td>${escapeHtml((l.last_invoice_date||"").slice(0,10))}</td>
                  <td data-click-filter="client" data-fv="${escapeHtml(l.customer||"")}" title="Click: filtrar por este cliente">${escapeHtml(l.customer||"")}</td>
                  <td><b data-click-filter="prod" data-fv="${escapeHtml(l.product_code||"")}" title="Click: filtrar por este producto">${escapeHtml(l.product_code||"")}</b> ${escapeHtml((l.product_name||"")).slice(0,40)}</td>
                  <td class="num">${fmtNum(l.qty)}</td>
                  <td class="num">${fmtMoney(l.sub)}</td>
                  <td class="num">${fmtPct(l.discount_pct/100)}</td>
                  <td class="num"><b style="color:#2e7d32">${fmtMoney(l.commission)}</b></td>
                </tr>`).join("")}
              ${det.length > 2000 ? `<tr><td colspan="11" class="muted" style="text-align:center">${det.length-2000} líneas más — usa el filtro o exporta CSV</td></tr>` : ""}
            </tbody>
          </table>
        </div>
      </div>`;
  }
  document.getElementById("pay-body").innerHTML = summary + detail;
}

function exportPaymentsCSV(){
  const fSp = document.getElementById("pay-sp").value;
  const fPeriod = document.getElementById("pay-period").value;
  const { byPair, lines } = computePayments();
  let det = lines.slice();
  if (fSp) det = det.filter(l => l.sp === fSp);
  if (fPeriod) det = det.filter(l => l.period === fPeriod);
  det.sort((a,b) => _periodOrder(a.period)-_periodOrder(b.period) || a.sp.localeCompare(b.sp));
  const head = ["Periodo_cobro","Comercial","SO","Fecha_SO","Fecha_factura","Cliente","Cod_producto","Producto","Qty","Subtotal_EUR","Dto_pct","Comision_EUR"].join(";");
  const body = det.map(l => [
    l.period, l.sp, l.order_name||"",
    (l.date_order||"").slice(0,10), (l.last_invoice_date||"").slice(0,10),
    l.customer||"", l.product_code||"", (l.product_name||"").replace(/[\n;]/g,' '),
    (l.qty||0).toFixed(2), l.sub.toFixed(2),
    (l.discount_pct||0).toFixed(2), l.commission.toFixed(2),
  ].map(v => {
    const s = (""+v).replace(/"/g,'""');
    return /[;"\n]/.test(s) ? `"${s}"` : s;
  }).join(";")).join("\n");
  const csv = "﻿" + head + "\n" + body;
  const blob = new Blob([csv], { type:"text/csv;charset=utf-8" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url; a.download = `pago_comisiones_${fSp||'todos'}_${fPeriod||'todos'}_${new Date().toISOString().slice(0,10)}.csv`;
  a.click(); URL.revokeObjectURL(url);
}

// =============================================================================
// VISTA TARIFAS (tarifas historicas deducidas)
// =============================================================================
function _flatTariffRows(){
  const rows = [];
  if (!TARIFFS || !TARIFFS.by_product_year) return rows;
  for (const [code, byPeriod] of Object.entries(TARIFFS.by_product_year)){
    const meta = TARIFFS.products[code] || {};
    for (const [period, d] of Object.entries(byPeriod)){
      if (!d) continue;
      rows.push({
        code, name: meta.name || code, current: meta.current_list_price,
        period, tariff: d.tariff, votes: d.votes, total: d.total_lines,
        confidence: d.confidence == null ? null : d.confidence,
        used_public: d.used_public, second: d.second_value, second_votes: d.second_votes,
        qty: d.qty || 0, revenue: d.revenue || 0, n_orders: d.n_orders || 0,
      });
    }
  }
  return rows;
}

// --- familias de producto (deducidas de product_category) ---
function _famForCategory(cat){
  if (!cat) return '(sin categoría)';
  const parts = cat.split(' / ').map(s=>s.trim());
  if ((parts[0]||'').startsWith('0-CATALOGO')){
    if (parts[1] === 'Controllers' && parts[2]) return parts[1] + ' / ' + parts[2];
    if (parts[1] === 'Panel PC' && parts[2])    return parts[1] + ' / ' + parts[2];
    if (parts[1] === 'Solutions' && parts[2])   return parts[1] + ' / ' + parts[2];
    return parts[1] || parts[0];
  }
  // Operaciones, Logistica, Administracion...
  return (parts[0]||'').replace(/^\d+-/, '') + (parts[1] ? ' / ' + parts[1] : '');
}
// Mapa code -> familia, calculado una vez (moda de category por code)
const CODE_TO_FAM = (() => {
  const cnt = new Map(); // code -> Map(fam -> n)
  for (const r of DATA){
    if (r.is_section) continue;
    const code = r.product_code; if (!code) continue;
    const fam = _famForCategory(r.product_category);
    if (!cnt.has(code)) cnt.set(code, new Map());
    const m = cnt.get(code);
    m.set(fam, (m.get(fam) || 0) + 1);
  }
  const out = new Map();
  for (const [code, m] of cnt){
    let bestFam = null, bestN = -1;
    for (const [fam, n] of m) if (n > bestN){ bestN = n; bestFam = fam; }
    out.set(code, bestFam || '(sin categoría)');
  }
  return out;
})();
function famForCode(code){ return CODE_TO_FAM.get(code) || '(sin categoría)'; }

// --- helpers comunes para Tarifas ---
function _tariffPeriodKind(p){
  if (/^\d{4}$/.test(p)) return 'annual';
  if (/^\d{4}-Q\d$/.test(p)) return 'quarterly';
  if (/^\d{4}-\d{2}$/.test(p)) return 'monthly';
  return 'other';
}
function _tariffPeriodOrder(p){
  // anual primero (sirve de "fallback"); luego trim/mensual cronológico
  const kind = _tariffPeriodKind(p);
  if (kind === 'annual') return parseInt(p)*100;
  if (kind === 'quarterly') {
    const [y,q] = p.split('-Q'); return parseInt(y)*100 + (parseInt(q)*3-2);
  }
  if (kind === 'monthly') {
    return parseInt(p.slice(0,4))*100 + parseInt(p.slice(5,7));
  }
  return 0;
}
function _filterPeriods(periods, _year, gran){
  // _year se mantiene en la firma por compatibilidad — ya no se usa: el filtro
  // de fechas global de la cabecera aplica via periodInGlobalRange().
  return periods.filter(p => {
    const kind = _tariffPeriodKind(p);
    // filtro de granularidad
    if (gran === 'annual'    && kind !== 'annual')    return false;
    if (gran === 'quarterly' && kind !== 'quarterly') return false;
    if (gran === 'monthly'   && kind !== 'monthly')   return false;
    if (gran === 'auto'){
      // Auto: trimestral para <=2025, mensual para >=2026, sin agregados anuales
      const y = parseInt(p.slice(0,4));
      if (kind === 'annual') return false;
      if (y <= 2025 && kind !== 'quarterly') return false;
      if (y >= 2026 && kind !== 'monthly')   return false;
    }
    // filtro de tramo de fechas GLOBAL
    if (!periodInGlobalRange(p)) return false;
    return true;
  });
}

function _bindTariffControls(){
  const root = document.getElementById('view-tariffs');
  if (!root || root.dataset.bound) return;
  root.dataset.bound = '1';
  // Poblar dropdown de familias (solo familias presentes en TARIFFS)
  const fams = new Set();
  if (TARIFFS && TARIFFS.products){
    for (const code of Object.keys(TARIFFS.products)) fams.add(famForCode(code));
  }
  const famSel = document.getElementById('tariffs-fam');
  [...fams].sort((a,b)=>a.localeCompare(b,'es')).forEach(f => {
    famSel.insertAdjacentHTML('beforeend', `<option value="${escapeHtml(f)}">${escapeHtml(f)}</option>`);
  });
  document.getElementById("tariffs-q").addEventListener("input", renderTariffs);
  document.getElementById("tariffs-mode").addEventListener("change", renderTariffs);
  document.getElementById("tariffs-gran").addEventListener("change", renderTariffs);
  document.getElementById("tariffs-fam").addEventListener("change", renderTariffs);
  document.getElementById("tariffs-group").addEventListener("change", renderTariffs);
  document.getElementById("tariffs-conf").addEventListener("change", renderTariffs);
  document.getElementById("btn-tariffs-csv").addEventListener("click", exportTariffsCSV);
}

function renderTariffs(){
  _bindTariffControls();
  const mode = document.getElementById("tariffs-mode").value || 'pivot';
  // Mostrar/ocultar el selector de confianza segun modo
  const lbl = document.querySelector('#view-tariffs .lbl-only-list');
  if (lbl) lbl.style.display = (mode === 'list') ? '' : 'none';
  if (mode === 'list') return _renderTariffsList();
  return _renderTariffsPivot();
}

// =============================================================================
// VISTA PIVOTE: filas = productos, columnas = periodos, celdas fusionadas
// =============================================================================
function _renderTariffsPivot(){
  const q = (document.getElementById("tariffs-q").value || "").toLowerCase().trim();
  const year = ""; // se ignora — sustituido por filtro de tramo de fechas global
  const gran = document.getElementById("tariffs-gran").value || 'auto';
  const fFam = document.getElementById("tariffs-fam").value || '';
  const grouped = document.getElementById("tariffs-group").checked;

  // 1) Recolectar productos y periodos visibles
  const products = new Map();   // code -> { code, name, current, byPeriod: Map(period -> tariffData) }
  const allPeriods = new Set();
  if (TARIFFS && TARIFFS.by_product_year){
    for (const [code, byPeriod] of Object.entries(TARIFFS.by_product_year)){
      const meta = TARIFFS.products[code] || {};
      const name = meta.name || code;
      const fam  = famForCode(code);
      if (q && !(code+" "+name).toLowerCase().includes(q)) continue;
      if (fFam && fam !== fFam) continue;
      const m = new Map();
      for (const [period, d] of Object.entries(byPeriod)){
        if (!d) continue;
        m.set(period, d);
        allPeriods.add(period);
      }
      if (m.size === 0) continue;
      products.set(code, { code, name, fam, current: meta.current_list_price, byPeriod: m });
    }
  }
  // 2) Filtrar periodos por año + granularidad
  const periods = _filterPeriods([...allPeriods], year, gran)
                    .sort((a,b) => _tariffPeriodOrder(a) - _tariffPeriodOrder(b));

  // 3) Filtrar productos: solo los que tienen alguna celda en los periodos visibles
  const products2 = [];
  for (const p of products.values()){
    const has = periods.some(per => p.byPeriod.has(per));
    if (has) products2.push(p);
  }
  // Ordenar: por familia y luego por código (la agrupación lo necesita igual)
  products2.sort((a,b) => a.fam.localeCompare(b.fam,'es') || a.code.localeCompare(b.code));

  const limit = 600;
  const visible = products2.slice(0, limit);

  // Agrupar productos visibles por familia
  const groups = []; // [{fam, prods}]
  let curFam = null, curList = null;
  for (const p of visible){
    if (p.fam !== curFam){
      curFam = p.fam; curList = [];
      groups.push({ fam: curFam, prods: curList });
    }
    curList.push(p);
  }

  // 4) Construir HTML de la tabla con merge de celdas consecutivas iguales
  const fmt = v => (v==null) ? '' : fmtMoney(v);
  const periodTitle = (p) => {
    const kind = _tariffPeriodKind(p);
    if (kind === 'annual') return p + ' (agregado)';
    if (kind === 'quarterly'){
      const [y,q] = p.split('-Q');
      const m1 = (q-1)*3+1; const m2 = m1+2;
      return `${p} · ${String(m1).padStart(2,'0')}-${String(m2).padStart(2,'0')}/${y}`;
    }
    return p;
  };

  const headerRow = `
    <tr>
      <th class="pv-sticky pv-code">Código</th>
      <th class="pv-sticky pv-name">Producto</th>
      <th class="pv-sticky pv-fam">Familia</th>
      <th class="num pv-sticky pv-list">List actual €</th>
      ${periods.map(p => `<th class="num pv-period" title="${escapeHtml(periodTitle(p))}">${escapeHtml(p)}</th>`).join("")}
    </tr>`;

  const renderProdRow = (prod) => {
    const cells = [];
    let i = 0;
    while (i < periods.length){
      const per = periods[i];
      const d = prod.byPeriod.get(per);
      const t = d ? d.tariff : null;
      let j = i+1;
      while (j < periods.length){
        const dn = prod.byPeriod.get(periods[j]);
        const tn = dn ? dn.tariff : null;
        if (t === tn) j++;
        else break;
      }
      const span = j - i;
      const fiable = d && d.confidence != null && d.confidence >= MIN_TARIFF_CONFIDENCE;
      const cur = prod.current;
      const delta = (cur && t!=null) ? (t - cur) : null;
      const cls = (t == null) ? 'pv-empty' : (fiable ? 'pv-ok' : 'pv-low');
      const merged = span > 1 ? ' pv-merged' : '';
      const tooltip = d
        ? `${per} · ${fmtMoney(t)} € · conf=${d.confidence!=null?(d.confidence*100).toFixed(0)+'%':'?'} (${d.votes}/${d.total_lines}) · qty=${fmtNum(d.qty||0)} · ${fmtMoney(d.revenue||0)} €${delta!=null?` · Δ vs list: ${(delta>0?'+':'')}${delta.toFixed(2)} €`:""}`
        : `${per} · sin datos`;
      cells.push(`<td colspan="${span}" class="num pv-cell ${cls}${merged}" title="${escapeHtml(tooltip)}">${
        t == null ? '<span class="muted">—</span>' : `<b>${fmt(t)}</b>${span>1?` <span class="muted" style="font-size:10px">×${span}</span>`:""}`
      }</td>`);
      i = j;
    }
    return `
      <tr>
        <td class="pv-sticky pv-code"><b>${escapeHtml(prod.code)}</b></td>
        <td class="pv-sticky pv-name" title="${escapeHtml(prod.name)}">${escapeHtml(prod.name).slice(0,60)}</td>
        <td class="pv-sticky pv-fam" title="${escapeHtml(prod.fam)}">${escapeHtml(prod.fam).slice(0,28)}</td>
        <td class="num pv-sticky pv-list">${prod.current!=null?fmt(prod.current):'<span class="muted">—</span>'}</td>
        ${cells.join("")}
      </tr>`;
  };

  const totalCols = 4 + periods.length;
  let bodyRows = '';
  if (grouped){
    for (const g of groups){
      bodyRows += `
        <tr class="pv-group-hdr">
          <td colspan="${totalCols}">
            <span class="pv-group-name">${escapeHtml(g.fam)}</span>
            <span class="muted" style="margin-left:8px;font-weight:400">· ${g.prods.length} producto${g.prods.length===1?'':'s'}</span>
          </td>
        </tr>`;
      bodyRows += g.prods.map(renderProdRow).join("");
    }
  } else {
    // Sin agrupación: orden por código
    const flat = visible.slice().sort((a,b)=>a.code.localeCompare(b.code));
    bodyRows = flat.map(renderProdRow).join("");
  }

  const html = `
    <div class="card">
      <h3>Tarifas (pivote) <span class="muted">${products2.length} productos · ${groups.length} familia${groups.length===1?'':'s'} · ${periods.length} periodos${products2.length>limit?` · mostrando primeros ${limit}`:""}</span></h3>
      <div class="pv-legend">
        <span class="pv-chip pv-ok"></span>tarifa fiable (conf ≥ 55%)
        <span class="pv-chip pv-low"></span>baja confianza (no se usa para comisión)
        <span class="pv-chip pv-empty"></span>sin datos
        <span class="muted" style="margin-left:18px">Las celdas con la misma tarifa en periodos consecutivos se fusionan (×N).</span>
      </div>
      <div class="pv-wrap">
        <table class="tbl-sum tbl-pivot">
          <thead>${headerRow}</thead>
          <tbody>${bodyRows || `<tr><td colspan="${totalCols}" class="muted" style="text-align:center;padding:24px">Sin tarifas para los filtros actuales.</td></tr>`}</tbody>
        </table>
      </div>
    </div>`;
  document.getElementById("tariffs-body").innerHTML = html;
}

// =============================================================================
// VISTA LISTA (la antigua, mantenida como secundaria)
// =============================================================================
function _renderTariffsList(){
  const q = (document.getElementById("tariffs-q").value || "").toLowerCase().trim();
  const year = ""; // se ignora — sustituido por filtro de tramo de fechas global
  const gran = document.getElementById("tariffs-gran").value || 'auto';
  const fFam = document.getElementById("tariffs-fam").value || '';
  const cf = document.getElementById("tariffs-conf").value;

  let rows = _flatTariffRows().map(r => ({...r, fam: famForCode(r.code)}));
  if (q) rows = rows.filter(r => (r.code+" "+r.name).toLowerCase().includes(q));
  if (fFam) rows = rows.filter(r => r.fam === fFam);
  // Filtrar periodos por año+granularidad
  const periodsAll = [...new Set(rows.map(r=>r.period))];
  const periodsOk = new Set(_filterPeriods(periodsAll, year, gran));
  rows = rows.filter(r => periodsOk.has(r.period));
  if (cf === "0.55") rows = rows.filter(r => (r.confidence||0) >= MIN_TARIFF_CONFIDENCE);
  if (cf === "-0.55") rows = rows.filter(r => (r.confidence||0) < MIN_TARIFF_CONFIDENCE);

  // Métricas derivadas para sort
  for (const r of rows){
    r._eurUnit = r.qty > 0 ? r.revenue / r.qty : 0;
    r._delta   = (r.current && r.tariff) ? (r.tariff - r.current) : 0;
    r._periodOrd = _tariffPeriodOrder(r.period);
  }
  // Aplicar orden de la tabla
  const slist = TBL_SORT.tarif_list;
  _sortBy(rows, slist.key, slist.dir);
  const limit = 1000;
  const visible = rows.slice(0, limit);

  const tot = rows.reduce((t,r) => ({
    qty: t.qty + r.qty, revenue: t.revenue + r.revenue, n_orders: t.n_orders + r.n_orders,
  }), {qty:0, revenue:0, n_orders:0});

  const html = `
    <div class="card">
      <h3>Tarifas deducidas (lista) <span class="muted">(${rows.length} filas${rows.length>limit?`, mostrando primeras ${limit}`:""})</span></h3>
      <div style="max-height:620px;overflow:auto">
        <table class="tbl-sum" style="width:100%;min-width:1500px">
          <thead><tr>
            ${_sortHead('tarif_list','code','Código',{defaultDir:1})}
            ${_sortHead('tarif_list','name','Producto',{defaultDir:1})}
            ${_sortHead('tarif_list','fam','Familia',{defaultDir:1})}
            ${_sortHead('tarif_list','_periodOrd','Periodo',{defaultDir:1})}
            ${_sortHead('tarif_list','tariff','Tarifa €',{num:true})}
            ${_sortHead('tarif_list','qty','Unidades',{num:true,tip:"Total de unidades vendidas en el periodo."})}
            ${_sortHead('tarif_list','n_orders','Pedidos',{num:true,tip:"SOs únicas que contienen este producto."})}
            ${_sortHead('tarif_list','revenue','Importe €',{num:true,tip:"Subtotal EUR de todas las líneas del producto en el periodo."})}
            ${_sortHead('tarif_list','_eurUnit','€/unidad',{num:true,tip:"Precio medio efectivo = Importe / Unidades."})}
            ${_sortHead('tarif_list','confidence','Confianza',{num:true})}
            ${_sortHead('tarif_list','current','List actual €',{num:true})}
            ${_sortHead('tarif_list','_delta','Δ vs actual',{num:true})}
            <th>Uso</th>
          </tr></thead>
          <tbody>
            ${visible.map(r => {
              const conf = r.confidence;
              const fiable = conf != null && conf >= MIN_TARIFF_CONFIDENCE;
              const useChip = fiable
                ? `<span class="chip-type t1" style="font-size:10px">EN USO</span>`
                : `<span class="chip-type tnone" style="font-size:10px">descartada</span>`;
              const delta = (r.current && r.tariff) ? (r.tariff - r.current) : null;
              const deltaTxt = (delta == null)
                ? `<span class="muted">—</span>`
                : `<span style="color:${delta>0?'#2e7d32':delta<0?'#d32f2f':'#5b6b7c'}">${(delta>0?'+':'')}${delta.toFixed(2)} €</span>`;
              const eurPerUnit = r.qty > 0 ? r.revenue / r.qty : null;
              const dtoMedio = (eurPerUnit && r.tariff) ? (1 - eurPerUnit/r.tariff) : null;
              const eurUnitTxt = eurPerUnit != null
                ? `${fmtMoney(eurPerUnit)} €${dtoMedio>0.01?` <span class="muted" style="font-size:11px">(-${(dtoMedio*100).toFixed(0)}%)</span>`:""}`
                : '<span class="muted">—</span>';
              return `<tr>
                <td><b>${escapeHtml(r.code)}</b></td>
                <td title="${escapeHtml(r.name)}">${escapeHtml(r.name).slice(0,55)}</td>
                <td><span class="muted" style="font-size:11px">${escapeHtml(r.fam)}</span></td>
                <td>${escapeHtml(r.period)}</td>
                <td class="num"><b>${fmtMoney(r.tariff)} €</b></td>
                <td class="num">${fmtNum(r.qty)}</td>
                <td class="num">${r.n_orders}</td>
                <td class="num"><b>${fmtMoney(r.revenue)} €</b></td>
                <td class="num">${eurUnitTxt}</td>
                <td class="num"><span style="color:${fiable?'#2e7d32':'#d32f2f'}">${conf!=null?fmtPct(conf):"—"}</span> <span class="muted" style="font-size:10px">(${r.votes}/${r.total})</span></td>
                <td class="num">${r.current!=null?fmtMoney(r.current)+" €":""}</td>
                <td class="num">${deltaTxt}</td>
                <td>${useChip}</td>
              </tr>`;
            }).join("")}
          </tbody>
          <tfoot>
            <tr style="border-top:2px solid var(--accent);font-weight:600">
              <td colspan="5">TOTAL filtrado (${rows.length} filas)</td>
              <td class="num">${fmtNum(tot.qty)}</td>
              <td class="num">${tot.n_orders}</td>
              <td class="num"><b>${fmtMoney(tot.revenue)} €</b></td>
              <td colspan="5"></td>
            </tr>
          </tfoot>
        </table>
      </div>
    </div>`;
  document.getElementById("tariffs-body").innerHTML = html;
}

function exportTariffsCSV(){
  const q = (document.getElementById("tariffs-q").value || "").toLowerCase().trim();
  const year = ""; // se ignora — sustituido por filtro de tramo de fechas global
  const gran = document.getElementById("tariffs-gran").value || 'auto';
  const fFam = document.getElementById("tariffs-fam").value || '';
  const cf = document.getElementById("tariffs-conf").value;
  let rows = _flatTariffRows().map(r => ({...r, fam: famForCode(r.code)}));
  if (q) rows = rows.filter(r => (r.code+" "+r.name).toLowerCase().includes(q));
  if (fFam) rows = rows.filter(r => r.fam === fFam);
  const periodsAll = [...new Set(rows.map(r=>r.period))];
  const periodsOk = new Set(_filterPeriods(periodsAll, year, gran));
  rows = rows.filter(r => periodsOk.has(r.period));
  if (cf === "0.55") rows = rows.filter(r => (r.confidence||0) >= MIN_TARIFF_CONFIDENCE);
  if (cf === "-0.55") rows = rows.filter(r => (r.confidence||0) < MIN_TARIFF_CONFIDENCE);
  rows.sort((a,b) => a.fam.localeCompare(b.fam,'es') || a.code.localeCompare(b.code) || _tariffPeriodOrder(a.period)-_tariffPeriodOrder(b.period));
  const head = ["Codigo","Producto","Familia","Periodo","Tarifa_EUR","Unidades","Pedidos","Importe_EUR","EUR_por_unidad","Dto_medio_pct","Votos","Total","Confidence","Usada","Segundo_valor","Segundo_votos","List_actual_EUR"].join(";");
  const body = rows.map(r => {
    const eu = r.qty>0 ? r.revenue/r.qty : 0;
    const dto = (eu && r.tariff) ? (1 - eu/r.tariff)*100 : 0;
    return [
      r.code, (r.name||"").replace(/[\n;]/g,' '), r.fam, r.period,
      (r.tariff||0).toFixed(2), (r.qty||0).toFixed(2), r.n_orders||0,
      (r.revenue||0).toFixed(2), eu.toFixed(2), dto.toFixed(2),
      r.votes, r.total,
      r.confidence!=null?r.confidence.toFixed(3):"",
      (r.confidence!=null && r.confidence>=MIN_TARIFF_CONFIDENCE)?"Y":"N",
      r.second!=null?r.second:"", r.second_votes||"",
      r.current!=null?r.current:"",
    ];
  }).map(arr => arr.map(v => {
    const s = (""+v).replace(/"/g,'""');
    return /[;"\n]/.test(s) ? `"${s}"` : s;
  }).join(";")).join("\n");
  const csv = "﻿" + head + "\n" + body;
  const blob = new Blob([csv], { type:"text/csv;charset=utf-8" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url; a.download = `tarifas_${new Date().toISOString().slice(0,10)}.csv`;
  a.click(); URL.revokeObjectURL(url);
}

// =============================================================================
// VISTA CONDICIONES (reglas de comision por año)
// =============================================================================
let RULES_YEAR = null;

function renderRulesTabs(){
  const years = Object.keys(COMMISSION_CONFIG.byYear).map(Number).sort();
  if (RULES_YEAR == null) RULES_YEAR = years[years.length-1]; // ultimo por defecto
  const tabs = document.getElementById("rules-tabs");
  tabs.innerHTML = "";
  for (const y of years){
    const b = document.createElement("button");
    b.textContent = y;
    if (y === RULES_YEAR) b.classList.add("active");
    b.addEventListener("click", () => { RULES_YEAR = y; renderRulesTabs(); renderRulesBody(); });
    tabs.appendChild(b);
  }
}

function renderRulesBody(){
  const y = RULES_YEAR;
  const year = COMMISSION_CONFIG.byYear[y];
  const body = document.getElementById("rules-body");
  if (!year){ body.innerHTML = "<div class='card'><p class='muted'>Sin reglas configuradas para este año.</p></div>"; return; }

  // Tabla descuento -> rate Tipo 1/Tipo 2  (paso 1%, de 0 a 30%)
  const discountTable = [];
  for (let d=0; d<=30; d+=1){
    discountTable.push({d, r1: commissionRate(1, d), r2: commissionRate(2, d)});
  }

  // Comerciales por tipo
  const byType = {1:[], 2:[]};
  for (const [name, info] of Object.entries(COMMISSION_CONFIG.salespersons)){
    if (!byType[info.type]) byType[info.type] = [];
    byType[info.type].push(name);
  }
  byType[1].sort((a,b)=>a.localeCompare(b,"es"));
  byType[2].sort((a,b)=>a.localeCompare(b,"es"));

  const f1 = COMMISSION_CONFIG.formula[1];
  const f2 = COMMISSION_CONFIG.formula[2];

  let html = `
    <div class="rules-grid">

      <div class="card">
        <h3>1 · Clasificación del comercial <span class="muted">(año ${y})</span></h3>
        <p class="muted" style="font-size:12px;margin:0 0 10px 0">
          Cada comercial tiene un tipo (1 o 2). El tipo determina el <b>porcentaje base</b> antes de aplicar el descuento de línea.
        </p>
        <div class="type-box t1">
          <div class="nm">Tipo 1 <span class="muted">· base ${f1.base}%</span></div>
          <div class="ds">${byType[1].length ? byType[1].join(" · ") : "(sin comerciales)"}</div>
        </div>
        <div class="type-box t2">
          <div class="nm">Tipo 2 <span class="muted">· base ${f2.base}%</span></div>
          <div class="ds">${byType[2].length ? byType[2].join(" · ") : "(sin comerciales)"}</div>
        </div>
      </div>

      <div class="card">
        <h3>2 · Comisión según descuento de línea</h3>
        <p class="muted" style="font-size:12px;margin:0 0 10px 0">
          Por cada <b>1%</b> de descuento aplicado en la línea se resta <b>${f1.stepPerPct}%</b> al porcentaje base. Se satura al <b>${f1.maxDiscountPct}%</b> de descuento.
        </p>
        <div class="formula">rate = max(0, base − discount_pct × ${f1.stepPerPct})</div>
        <div style="max-height:320px;overflow:auto;margin-top:10px;border:1px solid var(--line);border-radius:6px">
          <table class="tbl-sum" style="margin:0">
            <thead><tr>
              <th>Descuento</th>
              <th class="num"><span class="chip-type t1">Tipo 1</span></th>
              <th class="num"><span class="chip-type t2">Tipo 2</span></th>
            </tr></thead>
            <tbody>
              ${discountTable.map(row => `
                <tr>
                  <td>${row.d}%</td>
                  <td class="num">${row.r1.toFixed(2)}%</td>
                  <td class="num">${row.r2.toFixed(2)}%</td>
                </tr>`).join("")}
            </tbody>
          </table>
        </div>
      </div>

      <div class="card" style="grid-column:1 / -1">
        <h3>3 · Factor anual por ventas ${year.factorTiers ? "<span class='chip-type t1'>activo</span>" : "<span class='chip-type tnone'>no aplica</span>"}</h3>
        <p class="muted" style="font-size:12px;margin:0 0 10px 0">${escapeHtml(year.note||"")}</p>
        ${year.factorTiers ? `
        <div class="formula" style="margin-bottom:10px">comisión_final = rate × base_línea × factor(ventas_YTD)</div>
        <table class="tbl-sum">
          <thead><tr>
            <th>Tramo (anual)</th>
            <th class="num">Umbral anual</th>
            <th class="num">Factor</th>
          </tr></thead>
          <tbody>
            ${year.factorTiers.map((t,i) => {
              const from = i===0 ? 0 : year.factorTiers[i-1].upToAnnual;
              const toLbl = t.upToAnnual === Infinity ? "∞" : fmtMoney(t.upToAnnual) + " €";
              return `<tr>
                <td>${fmtMoney(from)} € → ${toLbl}</td>
                <td class="num">${toLbl}</td>
                <td class="num"><b>×${t.factor}</b></td>
              </tr>`;
            }).join("")}
          </tbody>
        </table>
        <div class="rules-note">
          <b>Prorrateo mensual:</b> al cerrar el mes M, los umbrales se calculan como <code>umbral_anual × M/12</code> y se aplica el factor correspondiente a las ventas YTD a fin de mes. Las comisiones ya liquidadas en meses anteriores <b>no se recalculan</b>.
        </div>
        ` : ""}
      </div>

      <div class="card" style="grid-column:1 / -1">
        <h3>4 · Regla especial para Projects / PHP <span class="chip-type t1">plana 3%</span></h3>
        <p class="muted" style="font-size:12px;margin:0 0 10px 0">
          Productos con código <code>PHP-*</code> ("Pack desarrollo proyectos software Programador")
          o categoría <code>0-CATALOGO / Projects / ...</code> cobran comisión
          <b>plana del ${COMMISSION_CONFIG.projectRule.flatRate}%</b> sobre el subtotal de la línea,
          independientemente del descuento aplicado y del Tipo (1 ó 2) del comercial.
        </p>
        <div class="formula">
          comisión_PHP = subtotal_línea × ${COMMISSION_CONFIG.projectRule.flatRate}%
        </div>
        <div class="rules-note">
          <b>Ejemplo:</b> línea PHP-200 con subtotal 12.000 € → comisión = 12.000 × 3% = <b>360 €</b>,
          aunque el descuento aplicado al cliente haya sido 0%, 30% o cualquier otro valor.
        </div>
        <div class="rules-note" style="margin-top:8px">
          <b>Controllino</b> también tiene regla propia (fee de marca), pero eso se aborda más adelante — aquí no se aplica aún.
        </div>
      </div>

      <div class="card" style="grid-column:1 / -1;border-left:4px solid #c62828">
        <h3>4-bis · Cuándo NO se aplica comisión <span class="chip-type tnone">exclusiones</span></h3>
        <p class="muted" style="font-size:12px;margin:0 0 12px 0">
          Una línea NO genera comisión si cumple <b>cualquiera</b> de los siguientes casos
          (mismo orden de evaluación que el código real <code>_commissionForLine</code>):
        </p>
        <ol style="font-size:13px;line-height:1.7;margin:0;padding-left:22px">
          <li>
            <b>Sección o nota</b> del pedido (<code>is_section = true</code>) — son separadores, no productos.
          </li>
          <li>
            <b>Estado del pedido ≠ Sale Order</b> (<code>state ≠ "sale"</code>) →
            quotations (draft / sent), cancelados, locked y borradores no devengan comisión.
            <span class="muted">En 2026 las quotations sí se contabilizan en la base de datos para previsiones, pero solo las confirmadas pasan a comisión.</span>
          </li>
          <li>
            <b>Salesperson vacío</b> o en la lista de <b>usuarios excluidos</b> (ver tarjeta 5).
          </li>
          <li>
            <b>Categoría o nombre del producto contiene "Shipping"</b> →
            portes, transporte y gastos de envío no entran en la base.
            <span class="muted">Filtro: <code>product_category</code> o <code>product_name</code> incluye literalmente "Shipping".</span>
          </li>
          <li>
            <b>Categoría del producto contiene "Controllino"</b> →
            la marca Controllino tiene fee aparte (regla independiente, todavía no liquidada aquí).
            <span class="muted">Filtro: <code>product_category</code> incluye "Controllino".</span>
          </li>
          <li>
            <b>Salesperson sin Tipo definido</b> (no aparece en T1/T2 ni en excluidos) →
            quedan en estado "Sin tipo"; solo cobran comisión PHP/Projects (3%).
          </li>
          <li>
            <b>Descuento ≥ 100%</b> (línea regalada / 0 € efectivo) → la línea no aporta nada al subtotal,
            la comisión es 0 € por construcción.
          </li>
        </ol>
        <div class="rules-note" style="margin-top:14px">
          <b>Resumen visual en otras vistas:</b>
          en <i>Comisiones</i>, los importes excluidos se muestran al pie como
          <code>Shipping</code> y <code>Controllino</code>. Los usuarios excluidos llevan el chip
          <span class="chip-type tnone" style="font-size:11px">exc.</span> y, si activas
          <i>Ocultar excluidos</i> en la cabecera, desaparecen de todas las tablas.
        </div>
      </div>

      <div class="card" style="grid-column:1 / -1">
        <h3>5 · Usuarios sin comisión <span class="chip-type tnone">excluidos</span></h3>
        <p class="muted" style="font-size:12px;margin:0 0 10px 0">
          Los siguientes usuarios aparecen como "Salesperson" en los pedidos pero NO devengan comisión
          (web, admin, back-office, ingeniería interna).  Fuente: Drive · "Plantilla calculo comisiones" → columna <i>No Salesperson</i>.
        </p>
        <div style="display:flex;flex-wrap:wrap;gap:6px">
          ${COMMISSION_CONFIG.excludedSalespersons.map(n =>
            `<span class="chip-type tnone">${escapeHtml(n)}</span>`).join("")}
        </div>
      </div>

      <div class="card" style="grid-column:1 / -1">
        <h3>6 · Ejemplo de cálculo estándar</h3>
        <p class="muted" style="font-size:12px;margin:0 0 10px 0">Línea de catálogo 10.000 € con 10% de descuento, comercial Tipo 1:</p>
        <div class="formula">
          rate = 3.6 − 10 × 0.1 = 2.6%<br>
          comisión_base = 10.000 € × 2.6% = 260 €${year.factorTiers ? "<br>comisión_final = 260 € × factor(YTD)" : ""}
        </div>
      </div>

    </div>`;
  body.innerHTML = html;
}

function renderRules(){
  renderRulesTabs();
  renderRulesBody();
}

// =============================================================================
// MI VISTA: filtra globalmente por un salesperson (SO)
// =============================================================================
function initMyView(){
  const sel = document.getElementById("my-view");
  // Poblar con los salespersons que aparezcan en los datos, EXCLUYENDO los que no cobran comisión
  const sps = [...new Set(DATA.map(r => r.salesperson).filter(Boolean))]
                .filter(sp => !EXCLUDED_SP_SET.has(sp))
                .sort((a,b)=>a.localeCompare(b,"es"));
  for (const sp of sps){
    sel.insertAdjacentHTML("beforeend",
      `<option value="${escapeHtml(sp)}">${escapeHtml(sp)}</option>`);
  }
  sel.addEventListener("change", applyMyView);
  // Re-render cuando cambia el toggle de excluidos (usa renderAll para incluir stats)
  document.getElementById("hide-excluded").addEventListener("change", renderAll);
}
function applyMyView(){
  const sel = document.getElementById("my-view");
  const v = sel.value;
  sel.parentElement.classList.toggle("active", !!v);

  // Mostrar Tipo del comercial seleccionado (o excluido / sin definir)
  const badge = document.getElementById("my-view-tipo");
  if (v){
    const t = salespersonType(v);
    if (t){
      const base = COMMISSION_CONFIG.formula[t].base;
      badge.innerHTML = `<span class="chip-type t${t}" style="margin-left:4px">Tipo ${t} · base ${base}%</span>`;
    } else if (isExcludedSalesperson(v)) {
      badge.innerHTML = `<span class="chip-type tnone" style="margin-left:4px" title="Excluido — sin comisión">Sin comisión</span>`;
    } else {
      badge.innerHTML = `<span class="chip-type tnone" title="Tipo sin definir" style="margin-left:4px">Tipo —</span>`;
    }
  } else {
    badge.innerHTML = "";
  }

  // Tabla: set filter on header select if exists
  FILTERS["salesperson"] = v;
  const headerSel = document.querySelector('#filters-row select[data-key="salesperson"]');
  if (headerSel) headerSel.value = v;

  // Dashboard: set f-sp select
  const fsp = document.getElementById("f-sp");
  if (fsp) fsp.value = v;

  // Comisiones: set comm-sp select
  const commSp = document.getElementById("comm-sp");
  if (commSp) commSp.value = v;

  // PAGO: set pay-sp select
  const paySp = document.getElementById("pay-sp");
  if (paySp) paySp.value = v;

  // Rerender la vista activa (sea cual sea)
  renderAll();
}

// =============================================================================
// VIEW SWITCH
// =============================================================================
const vTable   = document.getElementById("view-table");
const vDash    = document.getElementById("view-dash");
const vComm    = document.getElementById("view-comm");
const vPay     = document.getElementById("view-pay");
const vTariffs = document.getElementById("view-tariffs");
const vRules   = document.getElementById("view-rules");
const bT  = document.getElementById("btn-view-table");
const bD  = document.getElementById("btn-view-dash");
const bC  = document.getElementById("btn-view-comm");
const bP  = document.getElementById("btn-view-pay");
const bTa = document.getElementById("btn-view-tariffs");
const bR  = document.getElementById("btn-view-rules");

function clearActive(){
  [bT,bD,bC,bP,bTa,bR].forEach(b => b.classList.remove("active"));
}
function hideAll(){
  vTable.style.display="none"; vDash.style.display="none";
  vComm.style.display="none";  vPay.style.display="none";
  vTariffs.style.display="none"; vRules.style.display="none";
}
function showTable()  { hideAll(); vTable.style.display="";        clearActive(); bT.classList.add("active");  renderTable(); renderStats(); }
function showDash()   { hideAll(); vDash.style.display="block";    clearActive(); bD.classList.add("active");  renderDash();  renderStats(); }
function showComm()   { hideAll(); vComm.style.display="block";    clearActive(); bC.classList.add("active");  renderCommissions(); renderStats(); }
function showPay()    { hideAll(); vPay.style.display="block";     clearActive(); bP.classList.add("active");  renderPayments();  const b=document.getElementById("stats-bar"); if(b) b.innerHTML=""; }
function showTariffs(){ hideAll(); vTariffs.style.display="block"; clearActive(); bTa.classList.add("active"); renderTariffs();   const b=document.getElementById("stats-bar"); if(b) b.innerHTML=""; }
function showRules()  { hideAll(); vRules.style.display="block";   clearActive(); bR.classList.add("active");  renderRules();     const b=document.getElementById("stats-bar"); if(b) b.innerHTML=""; }
bT.addEventListener("click",  showTable);
bD.addEventListener("click",  showDash);
bC.addEventListener("click",  showComm);
bP.addEventListener("click",  showPay);
bTa.addEventListener("click", showTariffs);
bR.addEventListener("click",  showRules);

// Re-render la vista activa (sea cual sea) — usado por filtros globales
function renderAll(){
  if (vTable.style.display   !== "none") renderTable();
  else if (vDash.style.display  !== "none") renderDash();
  else if (vComm.style.display  !== "none") renderCommissions();
  else if (vPay.style.display   !== "none") renderPayments();
  else if (vTariffs.style.display !== "none") renderTariffs();
  renderStats();
}

// =============================================================================
// FILTRO GLOBAL: PERIODO (tramo de fechas)
// =============================================================================
function _todayIso(){ return new Date().toISOString().slice(0,10); }
function _addDays(iso, days){
  const d = new Date(iso + 'T00:00:00');
  d.setDate(d.getDate() + days);
  return d.toISOString().slice(0,10);
}
function _applyPreset(p){
  const f = document.getElementById("global-from");
  const t = document.getElementById("global-to");
  const today = _todayIso();
  const yearNow = today.slice(0,4);
  let from = "", to = "";
  if (p === "")             { from = ""; to = ""; }
  else if (p === "2024")    { from = "2024-01-01"; to = "2024-12-31"; }
  else if (p === "2025")    { from = "2025-01-01"; to = "2025-12-31"; }
  else if (p === "2026")    { from = "2026-01-01"; to = "2026-12-31"; }
  else if (p === "ytd")     { from = yearNow + "-01-01"; to = today; }
  else if (p === "last30")  { from = _addDays(today, -30); to = today; }
  else if (p === "last90")  { from = _addDays(today, -90); to = today; }
  else if (p === "last365") { from = _addDays(today, -365); to = today; }
  else if (p === "q3_2025_plus") { from = "2025-07-01"; to = ""; }
  else if (p === "custom")  { /* dejar como están */ return; }
  f.value = from; t.value = to;
}
function _refreshRangeUiActive(){
  const wrap = document.querySelector('.range-wrap');
  if (!wrap) return;
  const r = globalDateRange();
  wrap.classList.toggle("active", !!r);
}
(function initGlobalRange(){
  const preset = document.getElementById("global-preset");
  const f = document.getElementById("global-from");
  const t = document.getElementById("global-to");
  preset.addEventListener("change", () => {
    _applyPreset(preset.value);
    _refreshRangeUiActive();
    renderAll();
  });
  // Cuando el usuario edita las fechas a mano, marcar el preset como custom
  const onManualChange = () => {
    preset.value = "custom";
    _refreshRangeUiActive();
    renderAll();
  };
  f.addEventListener("change", onManualChange);
  t.addEventListener("change", onManualChange);
})();

// =============================================================================
// FILTROS GLOBALES: Producto + Cliente (con click-to-filter desde tablas)
// =============================================================================
function _refreshGlobalFilterChips(){
  // Marca como activo el wrap si tiene valor (visualmente)
  ['filter-prod','filter-client'].forEach(id => {
    const el = document.getElementById(id);
    if (!el) return;
    el.parentElement.classList.toggle('active', !!el.value);
  });
}
(function initGlobalProdClient(){
  const fp = document.getElementById('filter-prod');
  const fc = document.getElementById('filter-client');
  let dt;
  const onIn = () => {
    clearTimeout(dt);
    dt = setTimeout(() => { _refreshGlobalFilterChips(); renderAll(); }, 250);
  };
  fp.addEventListener('input', onIn);
  fc.addEventListener('input', onIn);
})();

// Helpers públicos: setProductFilter / setClientFilter — los usan los onclick
// de las celdas para fijar el filtro y re-renderizar todo.
function setProductFilter(code){
  const el = document.getElementById('filter-prod');
  if (!el) return;
  el.value = (code || '').trim();
  _refreshGlobalFilterChips();
  renderAll();
  // Pequeño feedback visual
  el.focus(); el.select();
}
function setClientFilter(name){
  const el = document.getElementById('filter-client');
  if (!el) return;
  el.value = (name || '').trim();
  _refreshGlobalFilterChips();
  renderAll();
  el.focus(); el.select();
}
// Delegación global: cualquier elemento con data-click-filter="prod|client"
// y data-fv="..." se convierte en un chip clickable que fija el filtro.
document.addEventListener('click', (e) => {
  const el = e.target.closest('[data-click-filter]');
  if (!el) return;
  const kind = el.dataset.clickFilter;
  const v = el.dataset.fv || el.textContent.trim();
  if (kind === 'prod')   setProductFilter(v);
  if (kind === 'client') setClientFilter(v);
});

// Boton Limpiar todo: resetea TODOS los filtros globales y de vistas
document.getElementById("btn-reset-all").addEventListener("click", () => {
  // Rango global
  document.getElementById("global-preset").value = "";
  document.getElementById("global-from").value = "";
  document.getElementById("global-to").value = "";
  _refreshRangeUiActive();
  // Mi vista
  document.getElementById("my-view").value = "";
  document.getElementById("my-view-tipo").innerHTML = "";
  // Filtros producto/cliente
  const fp = document.getElementById("filter-prod"); if (fp) fp.value = "";
  const fc = document.getElementById("filter-client"); if (fc) fc.value = "";
  document.querySelectorAll(".my-view-wrap").forEach(w => w.classList.remove("active"));
  // Tabla
  clearFilters();
  // Dashboard
  clearDashFilters();
  // Comisiones
  const cs = document.getElementById("comm-sp"); if (cs) cs.value = "";
  const cq = document.getElementById("comm-q"); if (cq) cq.value = "";
  const cmin = document.getElementById("comm-min"); if (cmin) cmin.value = "";
  const cmax = document.getElementById("comm-max"); if (cmax) cmax.value = "";
  const cp = document.getElementById("comm-pay"); if (cp) cp.value = "";
  const cd = document.getElementById("comm-detail"); if (cd) cd.checked = true;
  COMM_YEAR = "aggregate";
  renderAll();
});

// init con manejo de errores para no quedarse con overlay colgado
try {
  initDashFilters();
  initMyView();
  renderHead();
  renderTable();
  renderStats();
} catch (e) {
  console.error("Init error:", e);
  alert("Error al inicializar la vista: " + e.message + "\n\nMira la consola (F12) para más detalles.");
}
// Ocultar overlay de carga SIEMPRE al terminar (haya o no error)
const _ld = document.getElementById("loading");
if (_ld) _ld.style.display = "none";
</script>
</body>
</html>
"""

html = (HTML.replace('__PAYLOAD__',         _payload_b64)
            .replace('__TARIFFS__',         _tariffs_b64)
            .replace('__PAGOS__',           _pagos_b64)
            .replace('__DRIVE_SHEET_URL__', DRIVE_SHEET_URL))
payload = _payload_b64  # para el print de abajo
print(f"  pagos_registrados.csv: {len(_pagos_rows)} filas (+{_new_count} añadidas)")

out = os.path.join(HERE, 'Comisiones_Comerciales_Industrial_Shields.html')
with open(out, 'w', encoding='utf-8') as f:
    f.write(html)

# Limpiar artefacto antiguo DATA_SO.html si existia
legacy = os.path.join(HERE, 'DATA_SO.html')
if os.path.exists(legacy):
    os.remove(legacy)
    print(f"  (eliminado {legacy})")

print(f"{os.path.basename(out)} escrito: {len(html):,} bytes  |  {len(data)} filas  |  {len(payload):,} bytes payload")

# =============================================================================
# Generar HTMLs por comercial (--per-commercial)
# =============================================================================
# Cada comercial recibe su propio fichero con SOLO sus líneas + sus pagos.
# Aislamiento real (los datos del resto NO viajan en el fichero).
if PER_COMMERCIAL:
    print("\n[per-commercial] Generando HTML personalizado por comercial...")
    # Comerciales activos (T1/T2) que TIENEN al menos una línea en el dataset
    sp_with_data = {r.get('salesperson') for r in data if r.get('salesperson')}
    sp_list = sorted([sp for sp, t in _SP_TYPE.items()
                      if t and sp in sp_with_data and sp != 'SalesPerson'])
    out_dir = os.path.join(HERE, 'por_comercial')
    os.makedirs(out_dir, exist_ok=True)

    # Helper para nombre de fichero seguro preservando caracteres legibles
    def _slug(s):
        # Mapeo de acentos comunes → ASCII (preserva la sílaba final ó/í etc.)
        repl = str.maketrans({
            'á':'a','é':'e','í':'i','ó':'o','ú':'u','ñ':'n',
            'Á':'A','É':'E','Í':'I','Ó':'O','Ú':'U','Ñ':'N',
            'à':'a','è':'e','ì':'i','ò':'o','ù':'u',
            'À':'A','È':'E','Ì':'I','Ò':'O','Ù':'U',
            'ç':'c','Ç':'C',
        })
        s = s.translate(repl)
        return re.sub(r'[^A-Za-z0-9_-]+', '_', s).strip('_')

    for sp in sp_list:
        # Vista GERENTE: dataset completo sin filtros (Albert ve TODO)
        is_manager = sp in MANAGER_COMMERCIALS

        # 1) Filtrar lineas a las del comercial (gerente ve todas)
        if is_manager:
            sub_data = data
        else:
            sub_data = [r for r in data if r.get('salesperson') == sp]
        sub_payload_bytes = json.dumps(sub_data, ensure_ascii=False, separators=(',', ':')).encode('utf-8')
        sub_payload_b64 = base64.b64encode(gzip.compress(sub_payload_bytes, compresslevel=9)).decode('ascii')

        # 2) Filtrar pagos a los del comercial (gerente ve todos)
        if is_manager:
            sub_pagos_rows = _pagos_rows
        else:
            sub_pagos_rows = [r for r in _pagos_rows if r.get('comercial') == sp]
        sub_pagos_payload = {
            'drive_url': DRIVE_SHEET_URL,
            'drive_id':  DRIVE_SHEET_ID,
            'rows': [
                {
                    'period': r['periodo_cobro'], 'sp': r['comercial'],
                    'pagado': _to_float(r.get('importe_pagado_eur')),
                    'fecha': r.get('fecha_pago', ''), 'notas': r.get('notas', ''),
                }
                for r in sub_pagos_rows
            ],
        }
        sub_pagos_b64 = base64.b64encode(
            gzip.compress(json.dumps(sub_pagos_payload, ensure_ascii=False).encode('utf-8'), compresslevel=9)
        ).decode('ascii')

        # 3) Construir HTML personalizado
        sp_html = (HTML.replace('__PAYLOAD__',         sub_payload_b64)
                       .replace('__TARIFFS__',         _tariffs_b64)
                       .replace('__PAGOS__',           sub_pagos_b64)
                       .replace('__DRIVE_SHEET_URL__', DRIVE_SHEET_URL))

        # 4) Insertar banner identificativo + bloquear "Mi vista" al cargar.
        if is_manager:
            BANNER = (
                f'<div style="background:linear-gradient(90deg,#1a2f5c,#2c4373);'
                f'border-bottom:2px solid #1a2f5c;padding:8px 18px;font-size:12px;'
                f'color:#fff;display:flex;align-items:center;gap:10px;'
                f'font-weight:500">'
                f'<span style="background:#fff;color:#1a2f5c;padding:3px 9px;'
                f'border-radius:3px;font-weight:700;font-size:11px;letter-spacing:0.3px">VISTA GERENTE</span>'
                f'<span>Hola <b>{sp}</b> — acceso completo a todos los comerciales y datos. '
                f'Usa el filtro "Comercial:" para analizar a cada uno individualmente.</span>'
                f'</div>'
            )
            # Gerente: NO bloquear dropdown, valor por defecto vacío
            LOCK_JS = ""
        else:
            BANNER = (
                f'<div style="background:linear-gradient(90deg,#fde8ea,#fff5f5);'
                f'border-bottom:2px solid #e30613;padding:8px 18px;font-size:12px;'
                f'color:#1a2f5c;display:flex;align-items:center;gap:10px;'
                f'font-weight:500">'
                f'<span style="background:#e30613;color:#fff;padding:3px 9px;'
                f'border-radius:3px;font-weight:700;font-size:11px;letter-spacing:0.3px">VISTA PERSONAL</span>'
                f'<span>Hola <b>{sp}</b> — este documento contiene únicamente '
                f'<b>tus</b> pedidos y pagos. Los datos del resto de comerciales no están en este fichero.</span>'
                f'</div>'
            )
            # Bloquear el dropdown "Comercial:" después del init y poner su valor
            LOCK_JS = (
                '<script>'
                'document.addEventListener("DOMContentLoaded", function(){ setTimeout(function(){'
                f' var s=document.getElementById("my-view"); if(s){{ s.value={json.dumps(sp)};'
                f' s.dispatchEvent(new Event("change")); s.disabled=true; s.title="Bloqueado en vista personal"; }}'
                '}, 200); });'
                '</script>'
            )
        sp_html = sp_html.replace('<header', BANNER + '\n<header', 1)
        if LOCK_JS:
            sp_html = sp_html.replace('</body>', LOCK_JS + '\n</body>', 1)

        # 4-bis) Google OAuth gate (solo si hay CLIENT_ID configurado)
        if GOOGLE_OAUTH_CLIENT_ID:
            allowed_email = COMMERCIAL_EMAILS.get(sp, '')
            GATE_HTML = f'''
<div id="oauth-gate" style="position:fixed;inset:0;background:#f5f7fa;z-index:99999;display:flex;flex-direction:column;align-items:center;justify-content:center;padding:24px;text-align:center;font-family:Inter,-apple-system,sans-serif">
  <div style="max-width:480px;background:#fff;border-radius:8px;padding:32px;box-shadow:0 4px 16px rgba(0,0,0,0.08);border-top:4px solid #e30613">
    <h2 style="margin:0 0 8px 0;color:#1a2f5c;font-size:22px">Comisiones · Industrial Shields</h2>
    <p style="margin:0 0 6px 0;color:#5b6b7c;font-size:13px">Vista personal de <b style="color:#e30613">{sp}</b></p>
    <p style="margin:0 0 24px 0;color:#5b6b7c;font-size:12px">Inicia sesión con tu cuenta corporativa<br><b>{allowed_email}</b></p>
    <div id="g_id_signin" data-type="standard" data-theme="filled_blue" data-size="large" data-text="signin_with" data-shape="rectangular" data-locale="es"></div>
    <div id="oauth-err" style="margin-top:18px;color:#c62828;font-size:12px;display:none;background:#ffebee;padding:10px;border-radius:4px;border-left:3px solid #c62828"></div>
    <div style="margin-top:24px;font-size:11px;color:#a0aab8">Si tienes problemas para acceder, contacta con <a href="mailto:apm@industrialshields.com" style="color:#1976d2;text-decoration:none">apm@industrialshields.com</a></div>
  </div>
</div>
<script src="https://accounts.google.com/gsi/client" async defer></script>
<script>
  (function(){{
    var ALLOWED_EMAIL = {json.dumps(allowed_email)};
    var CLIENT_ID = {json.dumps(GOOGLE_OAUTH_CLIENT_ID)};
    function showErr(msg){{
      var el = document.getElementById('oauth-err');
      el.textContent = msg; el.style.display = 'block';
    }}
    // Esconder body hasta autenticación válida
    var s = document.createElement('style');
    s.textContent = 'body > *:not(#oauth-gate){{display:none !important}}';
    document.head.appendChild(s);
    window.handleGoogleCredential = function(resp){{
      if (!resp || !resp.credential) return showErr('Sin credenciales recibidas.');
      // Decodificar JWT (sin verificar firma — Google ya la validó del lado del cliente)
      try {{
        var parts = resp.credential.split('.');
        var payload = JSON.parse(atob(parts[1].replace(/-/g,'+').replace(/_/g,'/')));
        if (payload.hd && payload.hd !== 'industrialshields.com'){{
          return showErr('Solo cuentas @industrialshields.com pueden acceder. Has entrado con: ' + payload.email);
        }}
        if (payload.email !== ALLOWED_EMAIL){{
          return showErr('Esta cuenta no coincide. Este fichero está asignado a ' + ALLOWED_EMAIL + '. Has accedido con: ' + payload.email);
        }}
        // OK — quitar el gate y revelar la app
        s.remove();
        document.getElementById('oauth-gate').remove();
      }} catch(e){{
        showErr('Error al validar credenciales: ' + e.message);
      }}
    }};
    // Init Google Identity Services cuando cargue
    window.onload = function(){{
      if (typeof google === 'undefined' || !google.accounts){{
        setTimeout(window.onload, 200); return;
      }}
      google.accounts.id.initialize({{
        client_id: CLIENT_ID,
        callback: window.handleGoogleCredential,
        hd: 'industrialshields.com',
        ux_mode: 'popup',
        auto_select: false,
      }});
      google.accounts.id.renderButton(
        document.getElementById('g_id_signin'),
        {{type:'standard', theme:'filled_blue', size:'large', text:'signin_with', shape:'rectangular', locale:'es'}}
      );
    }};
  }})();
</script>
'''
            # Insertar antes del </body>
            sp_html = sp_html.replace('</body>', GATE_HTML + '\n</body>', 1)

        # 5) Cambiar título de la pestaña
        sp_html = sp_html.replace(
            '<title>Comisiones Comerciales · Industrial Shields</title>',
            f'<title>Comisiones · {sp} — Industrial Shields</title>'
        )

        # 6) Nombre de fichero seguro (con acentos transliterados)
        safe_name = _slug(sp)
        sp_path = os.path.join(out_dir, f'Comisiones_{safe_name}.html')
        with open(sp_path, 'w', encoding='utf-8') as f:
            f.write(sp_html)
        print(f"  · {safe_name:<28} {len(sub_data):>5} líneas  ·  {len(sp_html):>9,} bytes  ·  {sp_path}")

    print(f"[per-commercial] {len(sp_list)} ficheros en {out_dir}/")
