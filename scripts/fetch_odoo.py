#!/usr/bin/env python3
"""Descarga datos crudos de Odoo via XML-RPC y los guarda como raw_*.json.

Reemplaza al pull manual via MCP. Usa solo stdlib (xmlrpc.client) por simplicidad.

Variables de entorno requeridas:
  ODOO_URL      ej: https://industrialshields-prod.odoo.com
  ODOO_DB       ej: industrialshields-prod
  ODOO_USER     ej: apm@industrialshields.com
  ODOO_PASSWORD password / API key del usuario
"""
import os, json, sys, time
from xmlrpc.client import ServerProxy

# Pedidos y líneas desde 2024-01-01 (mismo recorte que ya usamos en el dashboard)
DATE_FROM = "2024-01-01"

ODOO_URL  = os.environ['ODOO_URL'].rstrip('/')
ODOO_DB   = os.environ['ODOO_DB']
ODOO_USER = os.environ['ODOO_USER']
ODOO_PASS = os.environ['ODOO_PASSWORD']

print(f"[fetch_odoo] Connecting to {ODOO_URL} db={ODOO_DB} as {ODOO_USER}")
common = ServerProxy(f'{ODOO_URL}/xmlrpc/2/common', allow_none=True)
uid = common.authenticate(ODOO_DB, ODOO_USER, ODOO_PASS, {})
if not uid:
    sys.exit("[fetch_odoo] ERROR: authenticate failed")
print(f"[fetch_odoo] uid={uid}")

models = ServerProxy(f'{ODOO_URL}/xmlrpc/2/object', allow_none=True)

def call(model, method, args, kwargs=None):
    return models.execute_kw(ODOO_DB, uid, ODOO_PASS, model, method, args, kwargs or {})

def search_read(model, domain, fields, batch=500, order=None):
    """Pagina en batches para evitar timeout."""
    out = []
    offset = 0
    while True:
        kw = {'fields': fields, 'offset': offset, 'limit': batch}
        if order: kw['order'] = order
        chunk = call(model, 'search_read', [domain], kw)
        if not chunk: break
        out.extend(chunk)
        offset += len(chunk)
        if len(chunk) < batch: break
        print(f"   {model}: {offset}…")
    return out

def save(name, data):
    path = os.path.join(os.path.dirname(__file__), name)
    with open(path, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False)
    print(f"   wrote {name}  ({os.path.getsize(path):,} bytes)")

# =============================================================================
# 1) sale.order  (state in [draft, sent, sale, done, cancel])  desde 2024
# =============================================================================
print("[fetch_odoo] sale.order …")
orders = search_read(
    'sale.order',
    [['date_order', '>=', DATE_FROM], ['team_id.name', '!=', 'Box Shields']],
    ['id', 'name', 'state', 'date_order', 'create_date',
     'partner_id', 'commercial_partner_id', 'user_id',
     'team_id', 'pricelist_id', 'currency_id'],
    order='date_order asc',
)
save('raw_orders.json', orders)

# =============================================================================
# 2) sale.order.line  (de los pedidos anteriores)
# =============================================================================
order_ids = [o['id'] for o in orders]
print(f"[fetch_odoo] sale.order.line  ({len(order_ids)} pedidos)…")
# Vamos por chunks de 500 IDs porque domain con muchos IDs puede dar timeout
lines = []
for i in range(0, len(order_ids), 500):
    chunk = order_ids[i:i+500]
    lines.extend(search_read(
        'sale.order.line',
        [['order_id', 'in', chunk]],
        ['id', 'order_id', 'display_type', 'product_id', 'product_template_id',
         'product_uom_qty', 'price_unit', 'price_subtotal', 'price_total', 'discount'],
    ))
    print(f"   lines so far: {len(lines)}")
save('raw_lines.json', lines)

# =============================================================================
# 3) res.partner  (los que aparecen en los pedidos)
# =============================================================================
partner_ids = set()
for o in orders:
    for k in ('partner_id', 'commercial_partner_id'):
        v = o.get(k)
        if isinstance(v, list) and v: partner_ids.add(v[0])
partner_ids = list(partner_ids)
print(f"[fetch_odoo] res.partner  ({len(partner_ids)})…")
partners = []
for i in range(0, len(partner_ids), 500):
    chunk = partner_ids[i:i+500]
    partners.extend(search_read(
        'res.partner',
        [['id', 'in', chunk]],
        ['id', 'name', 'commercial_partner_id', 'country_id',
         'category_id', 'team_id', 'user_id', 'property_product_pricelist'],
    ))
save('raw_partners.json', partners)

# =============================================================================
# 4) product.product  (los que aparecen en las lineas)
# =============================================================================
product_ids = set()
for l in lines:
    v = l.get('product_id')
    if isinstance(v, list) and v: product_ids.add(v[0])
product_ids = list(product_ids)
print(f"[fetch_odoo] product.product  ({len(product_ids)})…")
products = []
for i in range(0, len(product_ids), 500):
    chunk = product_ids[i:i+500]
    products.extend(search_read(
        'product.product',
        [['id', 'in', chunk]],
        ['id', 'default_code', 'name', 'list_price', 'standard_price',
         'categ_id', 'product_tmpl_id'],
    ))
save('raw_products.json', products)

# =============================================================================
# 5) res.partner.category (tags de partner) + tag_catalog
# =============================================================================
print("[fetch_odoo] res.partner.category …")
all_tags = call('res.partner.category', 'search_read',
                [[]], {'fields': ['id', 'name', 'parent_id']})
# Mapa partner -> tags (ids)
partner_tags = {}
for p in partners:
    tag_list = p.get('category_id') or []
    if tag_list:
        partner_tags[str(p['id'])] = tag_list
save('raw_partner_tags.json', {
    'partner_tags': partner_tags,
    'tag_catalog': {str(t['id']): t['name'] for t in all_tags},
})

# =============================================================================
# 6) res.currency.rate  (FX historico)
# =============================================================================
print("[fetch_odoo] res.currency.rate …")
rates = call('res.currency.rate', 'search_read',
             [[['name', '>=', '2023-01-01']]],
             {'fields': ['id', 'name', 'rate', 'currency_id', 'company_id'],
              'order': 'name asc', 'limit': 100000})
# Compañia EUR
companies = call('res.company', 'search_read', [[]],
                 {'fields': ['id', 'currency_id'], 'limit': 5})
company_currency = 'EUR'
if companies:
    cur = companies[0].get('currency_id')
    if isinstance(cur, list) and len(cur) >= 2: company_currency = cur[1]
save('raw_fx_rates.json', {
    '_note': 'Tasas Odoo res.currency.rate. Convención: rate = unidades extranjera por 1 EUR.',
    'company_currency': company_currency,
    'rates': rates,
})

# =============================================================================
# 7) sale.order (invoice fields)  — invoice_count / invoice_ids / invoice_status
# =============================================================================
print("[fetch_odoo] sale.order (invoice fields) …")
so_inv = []
for i in range(0, len(order_ids), 500):
    chunk = order_ids[i:i+500]
    so_inv.extend(call('sale.order', 'read',
                       [chunk], {'fields': ['id', 'invoice_count', 'invoice_ids', 'invoice_status']}))
save('raw_so_invoice.json', so_inv)

# =============================================================================
# 8) account.move (facturas vinculadas a los SOs)
# =============================================================================
inv_ids = set()
for s in so_inv:
    for x in s.get('invoice_ids') or []:
        inv_ids.add(x)
inv_ids = list(inv_ids)
print(f"[fetch_odoo] account.move  ({len(inv_ids)} facturas)…")
invoices = []
for i in range(0, len(inv_ids), 500):
    chunk = inv_ids[i:i+500]
    invoices.extend(call('account.move', 'read',
                          [chunk], {'fields': ['id', 'name', 'move_type', 'state',
                                                'payment_state', 'invoice_date',
                                                'invoice_origin',
                                                'amount_total_signed',
                                                'amount_residual_signed']}))
save('raw_invoices.json', invoices)

# =============================================================================
# 9) raw_order_meta — datos auxiliares (URL base de Odoo + algun mas)
# =============================================================================
save('raw_order_meta.json', {
    '_note': 'Metadatos auxiliares.',
    'odoo_base_url': ODOO_URL,
    'items': {},
})

print("[fetch_odoo] DONE.")
