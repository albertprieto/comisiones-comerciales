#!/usr/bin/env python3
"""Une raw_orders + raw_lines + raw_partners + raw_products + raw_partner_tags +
raw_fx_rates en un dataset plano (una fila por linea de oferta) listo para dashboards.

Transformaciones aplicadas (sobre la base 'original'):
  1) Se excluyen ofertas cuyo equipo de ventas sea 'Box Shields'.
  2) Se clasifica cada fila en customer_type ∈ {Distributor, Integrador, Cliente final,
     Sin clasificar}, con regla tags-de-partner primero y pricelist como fallback.
  3) Para monedas != EUR se aplica tipo de cambio histórico (res.currency.rate)
     a la fecha del pedido y se añaden columnas *_eur.
"""
import json, os
from datetime import datetime
from bisect import bisect_right

HERE = os.path.dirname(os.path.abspath(__file__))

# ---------------------------------------------------------------- utils
def load(name):
    with open(os.path.join(HERE, name), encoding='utf-8') as f:
        return json.load(f)

def m2o(val):
    """Devuelve (id, name) o (None, None) si el campo m2o esta vacio."""
    if isinstance(val, list) and len(val) == 2:
        return val[0], val[1]
    return None, None

# ---------------------------------------------------------------- datos
orders        = load('raw_orders.json')
lines         = load('raw_lines.json')
partners      = load('raw_partners.json')
products      = load('raw_products.json')
partner_tags  = load('raw_partner_tags.json')
fx_raw        = load('raw_fx_rates.json')
order_meta    = load('raw_order_meta.json')
# NUEVO: estado de facturacion por SO + lista de facturas
try:
    so_inv   = load('raw_so_invoice.json')
    invoices = load('raw_invoices.json')
except FileNotFoundError:
    so_inv, invoices = [], []

by_order    = {o['id']: o for o in orders}
by_partner  = {p['id']: p for p in partners}
by_product  = {p['id']: p for p in products}
by_ometa    = {m['id']: m for m in order_meta['items']}
ODOO_BASE   = order_meta['odoo_base_url']

# --- Indices de facturacion ---
inv_by_id    = {i['id']: i for i in invoices}
so_inv_by_id = {s['id']: s for s in so_inv}

def so_invoice_agg(order_id):
    """Resumen de facturacion para una SO.
    Devuelve:
      invoice_status    (Odoo: 'invoiced'|'no'|'to invoice'|'upselling')
      n_invoices         (solo out_invoice, state=posted)
      invoiced_amount    (suma amount_total_signed de out_invoice posted)
      residual_amount    (suma amount_residual_signed)
      refunded_amount    (suma amount_total_signed de out_refund posted, abs)
      payment_state_agg  ('paid'|'partial'|'not_paid'|'none'|'to_invoice')
      invoice_names      lista de nombres (ordenado por fecha)
      last_invoice_date  YYYY-MM-DD de la ultima factura posted
    """
    meta = so_inv_by_id.get(order_id)
    if not meta:
        return None
    inv_ids = meta.get('invoice_ids') or []
    invs = [inv_by_id[i] for i in inv_ids if i in inv_by_id]
    posted_out = [i for i in invs if i.get('state')=='posted' and i.get('move_type')=='out_invoice']
    posted_refund = [i for i in invs if i.get('state')=='posted' and i.get('move_type')=='out_refund']
    inv_status = meta.get('invoice_status') or 'no'

    if not posted_out:
        ps = 'to_invoice' if inv_status == 'to invoice' else 'none'
    else:
        residual_sum = sum((i.get('amount_residual_signed') or 0) for i in posted_out)
        total_sum    = sum((i.get('amount_total_signed')    or 0) for i in posted_out)
        if abs(residual_sum) < 0.01:
            ps = 'paid'
        elif abs(residual_sum) < abs(total_sum):
            ps = 'partial'
        else:
            ps = 'not_paid'

    invs_sorted = sorted(posted_out + posted_refund, key=lambda i: i.get('invoice_date') or '')
    names = [i.get('name') for i in invs_sorted if i.get('name')]
    last_date = invs_sorted[-1]['invoice_date'] if invs_sorted else None

    return {
        'invoice_status':    inv_status,
        'n_invoices':        len(posted_out),
        'invoiced_amount':   round(sum((i.get('amount_total_signed') or 0) for i in posted_out), 2),
        'residual_amount':   round(sum((i.get('amount_residual_signed') or 0) for i in posted_out), 2),
        'refunded_amount':   round(abs(sum((i.get('amount_total_signed') or 0) for i in posted_refund)), 2),
        'payment_state_agg': ps,
        'invoice_names':     names,
        'last_invoice_date': last_date,
    }

tag_catalog = {int(k): v for k, v in partner_tags['tag_catalog'].items()}
ptags_by_id = {int(k): list(v) for k, v in partner_tags['partner_tags'].items()}

STATE_LABEL = {
    'draft': 'Quotation',
    'sent': 'Quotation Sent',
    'sale': 'Sales Order',
    'done': 'Locked',
    'cancel': 'Cancelled',
}

EXCLUDED_TEAMS = {'Box Shields'}

# ---------------------------------------------------------------- FX
# rate en Odoo: cuantas unidades de foreign por 1 EUR (empresa en EUR)
#   ->  importe_eur = importe_foreign / rate
# Buscamos la tasa más reciente con fecha <= date_order
fx_by_cur = {}
for r in fx_raw['rates']:
    fx_by_cur.setdefault(r['currency'], []).append((r['date'], r['rate']))
# ordenar por fecha ascendente y preparar listas paralelas
fx_index = {}
for cur, rows in fx_by_cur.items():
    rows.sort(key=lambda x: x[0])
    fx_index[cur] = (
        [r[0] for r in rows],   # fechas
        [r[1] for r in rows],   # tasas
    )

def fx_rate(currency, date):
    """Devuelve (rate, rate_date_used) o (None, None) si no aplica.
    EUR siempre devuelve (1.0, None)."""
    if not currency or currency == 'EUR':
        return 1.0, None
    idx = fx_index.get(currency)
    if not idx or not date:
        return None, None
    dates, rates = idx
    pos = bisect_right(dates, date) - 1
    if pos < 0:
        # si no hay rate previo, coger el primero disponible
        return rates[0], dates[0]
    return rates[pos], dates[pos]

# ---------------------------------------------------------------- clasificacion
DIST_TAG_IDS = {110, 112, 2, 97, 98, 99}        # '0-Dist*', 'distributor*'
INT_TAG_IDS  = {113, 114, 115, 81}              # '1-Int*', 'integrador'
END_TAG_IDS  = {116, 117, 118}                  # '2-*'

# Pricelists cuya descripcion indica tipo de cliente (fallback por tarifa)
def classify_by_pricelist(pl_name):
    if not pl_name:
        return None
    n = pl_name.lower()
    # distribucion
    if (n.startswith('distributor_') or n.startswith('distribuidor_') or
        n.startswith('mouser') or n.startswith('digi-key') or
        n.startswith('farnell_') or n.startswith('rs_pricelist') or
        n.startswith('robotshop_') or n.startswith('plexus')):
        return 'Distributor'
    # cliente final
    if n.startswith('public pricelist') or n.startswith('public_pricelist'):
        return 'Cliente final'
    if n.startswith('website standard price') or n.startswith('website'):
        return 'Cliente final'
    # integrador
    if (n.startswith('rt_pricelist') or n.startswith('sunmobility') or
        n.startswith('aitecsa') or n.startswith('arduino_') or
        n.startswith('weidmuller') or n.startswith('montronic') or
        n.startswith('beijing') or n.startswith('benjamin')):
        return 'Integrador'
    # discounts genericos
    if n.startswith('10 % discount') or n.startswith('15% dropshipping') or n.startswith('25% nous'):
        return 'Integrador'
    if n.startswith('formación') or n.startswith('formacion'):
        return 'Cliente final'
    return None

def classify_by_tags(tag_ids):
    tag_ids = set(tag_ids or [])
    if tag_ids & DIST_TAG_IDS:
        return 'Distributor'
    if tag_ids & INT_TAG_IDS:
        return 'Integrador'
    if tag_ids & END_TAG_IDS:
        return 'Cliente final'
    return None

def classify(partner_id, pricelist_name):
    tags = ptags_by_id.get(partner_id, [])
    via = classify_by_tags(tags)
    if via:
        return via, 'tag'
    via = classify_by_pricelist(pricelist_name)
    if via:
        return via, 'pricelist'
    return 'Sin clasificar', 'default'

# ---------------------------------------------------------------- build rows
rows = []
excluded_boxshields = 0

for ln in lines:
    oid, _ = m2o(ln.get('order_id'))
    o = by_order.get(oid)
    if not o:
        continue

    partner_id, contact_name = m2o(o.get('partner_id'))
    comm_id, comm_name       = m2o(o.get('commercial_partner_id'))
    sp_id,   sp_name         = m2o(o.get('user_id'))
    tm_id,   tm_name         = m2o(o.get('team_id'))
    pl_id,   pl_name         = m2o(o.get('pricelist_id'))
    cur_id,  cur_name        = m2o(o.get('currency_id'))

    # >>> Filtro Box Shields <<<
    if tm_name in EXCLUDED_TEAMS:
        excluded_boxshields += 1
        continue

    # Partner commercial entity
    cp = by_partner.get(comm_id) or {}
    country_id, country_name = m2o(cp.get('country_id'))
    cp_sp_id, cp_sp_name     = m2o(cp.get('user_id'))
    cp_tm_id, cp_tm_name     = m2o(cp.get('team_id'))
    cp_pl_id, cp_pl_name     = m2o(cp.get('property_product_pricelist'))
    cp_tag_ids               = ptags_by_id.get(comm_id, [])
    cp_tag_names             = [tag_catalog.get(t, str(t)) for t in cp_tag_ids]

    # Clasificacion
    customer_type, via_source = classify(comm_id, pl_name)

    # Producto
    prod_id, prod_name = m2o(ln.get('product_id'))
    tmpl_id, tmpl_name = m2o(ln.get('product_template_id'))
    product = by_product.get(prod_id) or {}
    prod_code = product.get('default_code')
    # Limpia el "-" inicial que añade Odoo en variantes sin codigo propio
    if prod_code:
        prod_code = prod_code.lstrip('-').strip() or None
    # Y limpia tambien el "[-XXX]" inicial que se cuela en el display_name
    if prod_name and isinstance(prod_name, str):
        if prod_name.startswith('[-'):
            prod_name = '[' + prod_name[2:]
    list_price = product.get('list_price')
    cost = product.get('standard_price')
    categ_id, categ_name = m2o(product.get('categ_id'))

    # Economia (en moneda del pedido)
    qty = ln.get('product_uom_qty') or 0
    price_unit = ln.get('price_unit') or 0
    discount_pct = ln.get('discount') or 0
    subtotal = ln.get('price_subtotal') or 0
    total = ln.get('price_total') or 0

    # Coste de la ficha esta en EUR (company currency). Para comparar con subtotal
    # del pedido (que puede estar en USD), trabajamos el margen siempre en EUR.
    order_date = (o.get('date_order') or '')[:10]
    rate, rate_date = fx_rate(cur_name, order_date)
    fx_applied = (cur_name and cur_name != 'EUR' and rate is not None)

    def to_eur(x):
        if x is None or rate is None:
            return None
        return round(x / rate, 4)

    subtotal_eur = to_eur(subtotal)
    total_eur    = to_eur(total)
    price_unit_eur = to_eur(price_unit)

    # Coste ya esta en EUR (ficha). Margen se calcula en EUR.
    line_cost_eur = (cost or 0) * qty  # EUR
    margin_eur = (subtotal_eur - line_cost_eur) if (subtotal_eur is not None and cost is not None and qty) else None
    margin_pct = (margin_eur / subtotal_eur) if (margin_eur is not None and subtotal_eur) else None

    # External ID + deep-link URL al pedido en Odoo
    meta = by_ometa.get(oid) or {}
    ext_id = meta.get('xmlid') or None
    odoo_url = f"{ODOO_BASE}/web#id={oid}&model=sale.order&view_type=form"

    # Estado de facturacion/cobro agregado por SO
    inv_agg = so_invoice_agg(oid) or {}

    # Redondeo consistente para reducir bytes en el payload embebido
    def r2(x): return None if x is None else round(x, 2)
    def r4(x): return None if x is None else round(x, 4)

    rows.append({
        # === IDs ===
        'order_id':        oid,
        'order_name':      o.get('name'),
        'external_id':     ext_id,
        'odoo_url':        odoo_url,
        'is_section':      prod_id is None,

        # === Fechas y estado ===
        'date_order':      order_date,
        'state':           o.get('state'),
        'state_label':     STATE_LABEL.get(o.get('state'), o.get('state')),

        # === Cliente ===
        'commercial_entity_id':    comm_id,
        'commercial_entity_name':  comm_name,
        'country':                 country_name,

        # === Clasificacion ===
        'customer_type':           customer_type,
        'customer_type_source':    via_source,

        # === Comercial ===
        'salesperson':          sp_name,
        'salesteam':            tm_name,
        'partner_salesperson':  cp_sp_name,

        # === Pricelist / moneda ===
        'pricelist':            pl_name,
        'partner_pricelist':    cp_pl_name,
        'currency':             cur_name,

        # === FX ===
        'fx_rate':          r4(rate),
        'fx_rate_date':     rate_date,
        'fx_applied':       bool(fx_applied),

        # === Producto ===
        'product_name':     prod_name,
        'product_code':     prod_code,
        'product_category': categ_name,

        # === Ficha producto (EUR) ===
        'list_price':       r2(list_price),
        'standard_price':   r2(cost),

        # === Linea (moneda del pedido) ===
        'qty':              qty,
        'price_unit':       r2(price_unit),
        'discount_pct':     r2(discount_pct),
        'price_subtotal':   r2(subtotal),

        # === EUR ===
        'price_unit_eur':     r2(price_unit_eur),
        'price_subtotal_eur': r2(subtotal_eur),

        # === Calculados (EUR) ===
        'margin_eur':       r2(margin_eur),
        'margin_pct':       r4(margin_pct),

        # === Flags ===
        'cost_missing':     (cost is not None and cost <= 0),

        # === Facturacion / Cobro (agregado a nivel SO) ===
        'invoice_status':     inv_agg.get('invoice_status'),      # Odoo: invoiced/no/to invoice/upselling
        'payment_state_agg':  inv_agg.get('payment_state_agg'),   # paid/partial/not_paid/to_invoice/none
        'invoiced_amount':    inv_agg.get('invoiced_amount'),
        'residual_amount':    inv_agg.get('residual_amount'),
        'refunded_amount':    inv_agg.get('refunded_amount'),
        'n_invoices':         inv_agg.get('n_invoices') or 0,
        'invoice_names':      inv_agg.get('invoice_names') or [],
        'last_invoice_date':  inv_agg.get('last_invoice_date'),
    })

# Filtrar outliers evidentes (datos de prueba en drafts: subtotales > 10M€)
OUTLIER_THRESHOLD_EUR = 10_000_000
outlier_orders = {r['order_name'] for r in rows
                  if r.get('price_subtotal_eur') and r['price_subtotal_eur'] > OUTLIER_THRESHOLD_EUR}
if outlier_orders:
    before = len(rows)
    rows = [r for r in rows if r['order_name'] not in outlier_orders]
    print(f"Filtradas {before-len(rows)} lineas de {len(outlier_orders)} ofertas outlier "
          f"(subtotal > {OUTLIER_THRESHOLD_EUR:,}€): {sorted(outlier_orders)[:5]}")

# Ordenar por fecha de pedido
rows.sort(key=lambda r: (r['date_order'] or '', r['order_name'] or ''))

# Sin indent para maximo ahorro de espacio en el HTML embebido
with open(os.path.join(HERE, 'dataset_base.json'), 'w', encoding='utf-8') as f:
    json.dump(rows, f, ensure_ascii=False, separators=(',', ':'))

# ================================================================ TEST
from collections import Counter

total_rows = len(rows)
sections = sum(1 for r in rows if r['is_section'])
product_rows = total_rows - sections

print(f"Dataset base construido: {total_rows} filas")
print(f"  - Lineas producto:       {product_rows}")
print(f"  - Secciones/notas:       {sections}")
print(f"  - Excluidas Box Shields: {excluded_boxshields}")
print()

orders_set = {r['order_name'] for r in rows}
print(f"Ofertas unicas: {len(orders_set)}")

print("\n--- Estados ---")
for s, c in Counter(r['state'] for r in rows if not r['is_section']).most_common():
    print(f"  state={s}: {c} lineas")

print("\n--- customer_type ---")
ct_counts = Counter((r['customer_type'], r['customer_type_source']) for r in rows if not r['is_section'])
for (ct, src), c in ct_counts.most_common():
    print(f"  {ct:<15s} via {src:<10s} : {c} lineas")

print("\n--- Por cliente + tags + tarifa -> customer_type ---")
seen = set()
for r in rows:
    if r['is_section']:
        continue
    key = r['commercial_entity_id']
    if key in seen: continue
    seen.add(key)
    print(f"  [{r['customer_type']:<15s}] {r['commercial_entity_name']}")
    print(f"       pricelist: {r['pricelist']}  -> via {r['customer_type_source']}")

print("\n--- FX aplicado ---")
fx_rows = [r for r in rows if r['fx_applied']]
for r in fx_rows[:5]:
    print(f"  {r['order_name']} {r['date_order']} cur={r['currency']} "
          f"rate={r['fx_rate']} (@ {r['fx_rate_date']})  "
          f"subtotal={r['price_subtotal']} {r['currency']} -> {r['price_subtotal_eur']} EUR")
if fx_rows:
    # Verificacion: suma en USD vs EUR del primer pedido FX
    o0 = fx_rows[0]['order_name']
    usd_sum = sum(r['price_subtotal'] for r in fx_rows if r['order_name']==o0 and not r['is_section'])
    eur_sum = sum(r['price_subtotal_eur'] for r in fx_rows if r['order_name']==o0 and not r['is_section'])
    rate_u  = fx_rows[0]['fx_rate']
    print(f"\n  Total {o0}: {usd_sum:.2f} USD  ->  {eur_sum:.2f} EUR  (rate {rate_u})")
    print(f"  Verificacion: {usd_sum:.2f} / {rate_u} = {usd_sum/rate_u:.2f}")
else:
    print("  (ninguna fila en moneda distinta a EUR)")

print("\n--- Productos con coste=0 ---")
cost_missing = sorted({r['order_name'] + ' | ' + (r['product_name'] or '')
                       for r in rows if r['cost_missing']})
for x in cost_missing:
    print(f"  !! {x}")

print("\n--- Rango fechas ---")
dates = [r['date_order'] for r in rows if r['date_order']]
print(f"  {min(dates)} -> {max(dates)}")
