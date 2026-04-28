#!/usr/bin/env python3
"""Deduce la tarifa historica por (producto, año) desde el dataset.

Idea: por cada linea con precio en EUR, el precio implicito de tarifa es
    list_price_implicito = price_unit_eur / (1 - discount_pct/100)
La MODA de ese valor sobre todas las ventas de un producto en un año es,
con muchas observaciones, la tarifa real que habia en aquel momento.

Casos por pricelist:
- Lineas con pricelist publica ("Public Pricelist *", "Website Standard *")
  reflejan directamente el list_price publico.
- Lineas a distribuidor / integrador / mouser / etc. tienen un precio
  pre-descontado por la pricelist; su implicit tariff sale por debajo.

Estrategia:
1) MODA preferente: solo lineas con pricelist PUBLICA.
2) Si no hay (o muy pocas), MODA sobre TODAS y luego corregimos por
   factor pricelist conocido (heuristica). Para un primer pase, usamos
   la moda directa con todas y dejamos en metadata si fue PUBLIC vs ALL.
"""
import json, os, math
from collections import Counter, defaultdict

HERE = os.path.dirname(os.path.abspath(__file__))

with open(os.path.join(HERE, 'dataset_base.json'), encoding='utf-8') as f:
    DATA = json.load(f)

# ---- Pricelists "publicas" (reflejan list_price oficial) ----
PUBLIC_PL_PREFIXES = (
    'public pricelist', 'public_pricelist',
    'website standard price', 'website',
)
def is_public_pricelist(pl):
    if not pl: return False
    n = pl.lower().strip()
    return any(n.startswith(p) for p in PUBLIC_PL_PREFIXES)

# ---- Recoger candidatos por (product_id, year) ----
# candidates[pid][year] = lista de (implicit_tariff, is_public, sub_eur, source_so)
candidates = defaultdict(lambda: defaultdict(list))
prod_meta  = {}  # pid -> {name, code, current_list_price}

for r in DATA:
    if r.get('is_section'): continue
    pid = r.get('product_id') if 'product_id' in r else None
    # En el dataset compactado actual no tenemos product_id directo,
    # solo product_code y product_name. Usamos product_code como clave.
    pcode = r.get('product_code') or ''
    pname = r.get('product_name') or ''
    key = pcode or pname
    if not key: continue

    pu  = r.get('price_unit_eur') or 0
    dto = r.get('discount_pct') or 0
    qty = r.get('qty') or 0
    sub = r.get('price_subtotal_eur') or 0
    if pu <= 0: continue
    if dto >= 100: continue
    impl = pu / (1 - dto/100)
    if impl <= 0 or not math.isfinite(impl): continue

    d = (r.get('date_order') or '')[:10]
    yr = d[:4]
    if not yr.isdigit(): continue
    # Granularidad: trimestral 2024-2025, MENSUAL desde 2026
    yr_n = int(yr)
    if len(d) >= 7 and d[5:7].isdigit():
        mo = int(d[5:7])
    else:
        mo = 1
    if yr_n <= 2025:
        q = (mo - 1) // 3 + 1
        period = f"{yr}-Q{q}"
    else:
        period = f"{yr}-{mo:02d}"

    is_pub = is_public_pricelist(r.get('pricelist'))
    sample = (round(impl, 2), is_pub, sub, r.get('order_name'), qty)
    candidates[key][period].append(sample)
    # Para años trimestrales/mensuales, AÑADIMOS también un agregado anual
    # como fallback cuando el subperiodo no llegue a la confianza mínima.
    if period != yr:
        candidates[key][yr].append(sample)

    if key not in prod_meta:
        prod_meta[key] = {
            'name': pname,
            'code': pcode,
            'current_list_price': r.get('list_price'),
        }

# ---- Calcular moda con preferencia por pricelist publica ----
def deduce_tariff(samples):
    """samples: [(impl, is_pub, sub_eur, so, qty), ...]"""
    if not samples: return None
    pubs = [s for s in samples if s[1]]
    pool = pubs if len(pubs) >= 3 else samples
    cnt = Counter(s[0] for s in pool)
    most = cnt.most_common()
    if not most: return None
    top_val, top_n = most[0]
    n_total = len(pool)
    confidence = top_n / n_total if n_total else 0
    # Agregados de volumen / facturación sobre TODAS las muestras (no solo pool)
    qty_total = sum(s[4] for s in samples)
    revenue   = sum(s[2] for s in samples)
    n_orders  = len(set(s[3] for s in samples))
    return {
        'tariff':       top_val,
        'votes':        top_n,
        'total_lines':  n_total,
        'confidence':   round(confidence, 3),
        'used_public':  bool(pubs and len(pubs) >= 3),
        'second_value': most[1][0] if len(most) > 1 else None,
        'second_votes': most[1][1] if len(most) > 1 else 0,
        'qty':          round(qty_total, 2),
        'revenue':      round(revenue,  2),
        'n_orders':     n_orders,
    }

# ---- Construir tabla final ----
tariffs = {}  # key -> { year: deduced }
for key, by_year in candidates.items():
    tariffs[key] = {}
    for yr, samples in by_year.items():
        tariffs[key][yr] = deduce_tariff(samples)

out = {
    '_note': ("Tarifa deducida por moda de price_unit_eur / (1 - dto%) sobre todas las "
              "ventas/cotizaciones de cada producto. Granularidad: TRIMESTRAL para "
              "años <=2025 (claves 'YYYY-Q1'..'YYYY-Q4'), MENSUAL para >=2026 ('YYYY-MM'). "
              "Ademas se incluye el agregado anual 'YYYY' como fallback. "
              "Preferimos lineas con pricelist publica si hay >= 3, si no usamos toda la población. "
              "El campo 'confidence' = votos_moda / total_lineas; >=0.55 se considera fiable."),
    'products': prod_meta,
    'by_product_year': tariffs,  # {product_code: {periodo: {tariff,...}}}
}

with open(os.path.join(HERE, 'tariff_history.json'), 'w', encoding='utf-8') as f:
    json.dump(out, f, ensure_ascii=False, separators=(',', ':'))

# ---- Estadisticas de calidad ----
n_prod = len(prod_meta)
n_cells = sum(len(v) for v in tariffs.values())
n_with = sum(1 for v in tariffs.values() for d in v.values() if d)
n_public = sum(1 for v in tariffs.values() for d in v.values() if d and d['used_public'])
print(f"Productos: {n_prod}")
print(f"Celdas (producto, año): {n_cells}")
print(f"  con tarifa deducida: {n_with}")
print(f"  con base PUBLIC (>= 3 obs publicas): {n_public}")

# Top 10 productos por nº de pedidos
prod_orders = {k: sum(len(v) for v in by.values()) for k, by in candidates.items()}
top = sorted(prod_orders.items(), key=lambda kv: -kv[1])[:15]
print("\nTop productos por nº de líneas:")
print(f"  {'codigo':<15} {'nombre':<55} {'lineas':>8}")
for k, n in top:
    nm = (prod_meta.get(k) or {}).get('name','')[:55]
    print(f"  {k:<15} {nm:<55} {n:>8}")

# Mostrar la tarifa deducida para los top
print("\nTarifas deducidas (top 5 productos):")
for k, _n in top[:5]:
    nm = (prod_meta.get(k) or {}).get('name','')[:50]
    cur = (prod_meta.get(k) or {}).get('current_list_price')
    print(f"\n  [{k}] {nm}  (list_price actual: {cur})")
    for yr in sorted(tariffs[k].keys()):
        d = tariffs[k][yr]
        if not d: continue
        flag = '★pub' if d['used_public'] else 'all '
        sec = f", 2º={d['second_value']}€×{d['second_votes']}" if d['second_value'] else ''
        print(f"    {yr}: tariff={d['tariff']:>10.2f}€  votos={d['votes']:>3}/{d['total_lines']:<3} ({flag}){sec}")

print(f"\ntariff_history.json escrito: {os.path.getsize(os.path.join(HERE,'tariff_history.json')):,} bytes")
