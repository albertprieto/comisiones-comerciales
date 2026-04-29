#!/usr/bin/env python3
"""Envía email semanal a los comerciales con su URL personal y un resumen.

Solo envía a: Eloi, Gerard, Jordi, Josep (configurable abajo).
Lee dataset_base.json + pagos_registrados.csv para calcular el resumen.

Variables de entorno:
  GMAIL_USER     Usuario Gmail emisor (ej: apm@industrialshields.com)
  GMAIL_APP_PASS App Password de 16 caracteres (no la password normal).
                 Generar en https://myaccount.google.com/apppasswords
  REPORT_FROM    From: header (defaults a GMAIL_USER)
"""
import os, json, csv, smtplib, ssl
from collections import defaultdict
from datetime import datetime, date
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.utils import formataddr, formatdate

HERE = os.path.dirname(__file__)

# ─── Configuración del envío ──────────────────────────────────────────────────
RECIPIENTS = {
    "Albert Prieto":             "apm@industrialshields.com",
    "Eloi Davila Lopez":         "edl@industrialshields.com",
    "Gerard Montero Martínez":   "gmm@industrialshields.com",
    "Jordi Hernandez":           "jhs@industrialshields.com",
    "Josep Massó":               "jmp@industrialshields.com",
}

# Nombre de fichero seguro (mismo slug que build_html.py)
import re
def _slug(s):
    repl = str.maketrans({
        'á':'a','é':'e','í':'i','ó':'o','ú':'u','ñ':'n',
        'Á':'A','É':'E','Í':'I','Ó':'O','Ú':'U','Ñ':'N',
        'à':'a','è':'e','ì':'i','ò':'o','ù':'u',
        'À':'A','È':'E','Ì':'I','Ò':'O','Ù':'U',
        'ç':'c','Ç':'C',
    })
    return re.sub(r'[^A-Za-z0-9_-]+', '_', s.translate(repl)).strip('_')

PAGES_BASE = "https://albertprieto.github.io/comisiones-comerciales"

# Reglas de comisión (mismo logic que build_html.py)
EXCLUDED_SP = {"Industrial Shields - Website","ADMIN","Alba Sánchez Honrado",
               "Sònia Gabarró","Albert Macià","Abel Codina","Luis Nunes",
               "Francesc Duarri","Susana Guerra","Joan F. Aubets - Industrial Shields"}
SP_TYPE = {"Jordi Hernandez":1,"Garima Arora":1,"Eloi Davila Lopez":1,
           "Gerard Montero Martínez":1,"Josep Massó":2,"Ramon Boncompte":2,
           "Albert Prieto":2}

def _commission_line(r):
    if r.get('is_section') or r.get('state') != 'sale': return 0.0
    sp = r.get('salesperson')
    if not sp or sp in EXCLUDED_SP: return 0.0
    cat = r.get('product_category') or ''
    nm  = r.get('product_name') or ''
    code = r.get('product_code') or ''
    if 'Shipping' in cat or 'Shipping' in nm: return 0.0
    if 'Controllino' in cat: return 0.0
    sub = r.get('price_subtotal_eur') or 0.0
    if code.startswith('PHP-') or 'Projects' in cat:
        return sub * 0.03
    t = SP_TYPE.get(sp)
    if not t: return 0.0
    d = max(0, min(r.get('discount_pct') or 0, 30))
    rate = max(0, (3.6 if t==1 else 3.1) - d*0.1)
    return sub * rate / 100.0

def fmt_eur(x): return f"{x:,.2f} €".replace(',', '·').replace('.', ',').replace('·', '.')

# ─── Cargar dataset y construir resumen por comercial ─────────────────────────
def build_summary():
    dataset = os.path.join(HERE, 'dataset_base.json')
    if not os.path.exists(dataset):
        print(f"[send_emails] WARN: no existe {dataset}; resumen vacío.")
        return {}
    with open(dataset, encoding='utf-8') as f:
        data = json.load(f)

    # Agregados YTD del año en curso por comercial
    yr = str(date.today().year)
    agg = defaultdict(lambda: {'devengada':0.0, 'pagable':0.0, 'ventas':0.0, 'sos':set()})
    for r in data:
        com = _commission_line(r)
        if com <= 0: continue
        sp = r.get('salesperson')
        if (r.get('date_order') or '').startswith(yr):
            a = agg[sp]
            a['devengada'] += com
            a['ventas']    += r.get('price_subtotal_eur') or 0
            a['sos'].add(r.get('order_name'))
            if r.get('payment_state_agg') == 'paid':
                a['pagable']  += com
    for sp in agg: agg[sp]['sos'] = len(agg[sp]['sos'])
    return agg

def build_html_body(name, summary, url):
    s = summary.get(name, {'devengada':0,'pagable':0,'ventas':0,'sos':0})
    yr = date.today().year
    return f"""<!doctype html>
<html><head><meta charset="utf-8"></head>
<body style="font-family:Inter,-apple-system,Helvetica,Arial,sans-serif;background:#f5f7fa;margin:0;padding:24px;color:#1a2f5c">
  <div style="max-width:560px;margin:0 auto;background:#fff;border-radius:8px;overflow:hidden;box-shadow:0 2px 12px rgba(0,0,0,0.06);border-top:4px solid #e30613">
    <div style="padding:24px 28px">
      <h2 style="margin:0 0 4px 0;font-size:20px;color:#1a2f5c">Hola {name},</h2>
      <p style="margin:0 0 18px 0;color:#5b6b7c;font-size:14px">
        Aquí tienes tu actualización semanal de comisiones (datos de Odoo refrescados hoy).
      </p>

      <div style="background:#f8fafc;border-left:3px solid #e30613;padding:14px 16px;border-radius:4px;font-size:13px;line-height:1.6">
        <b style="color:#1a2f5c">Resumen YTD {yr}</b><br>
        Ventas comisionables: <b>{fmt_eur(s['ventas'])}</b><br>
        Comisión devengada: <b>{fmt_eur(s['devengada'])}</b><br>
        Comisión PAGABLE (cobrada): <b style="color:#2e7d32">{fmt_eur(s['pagable'])}</b><br>
        Pedidos: {s['sos']}
      </div>

      <p style="margin:22px 0 14px 0;font-size:14px">Abre tu panel personal con tu cuenta corporativa de Google:</p>
      <p style="text-align:center;margin:0">
        <a href="{url}" style="display:inline-block;background:#e30613;color:#fff;padding:12px 22px;border-radius:6px;text-decoration:none;font-weight:600;font-size:14px">
          Abrir mi panel
        </a>
      </p>

      <p style="margin:24px 0 0 0;font-size:11px;color:#a0aab8;line-height:1.5">
        El acceso está restringido a tu email <b>@industrialshields.com</b>.
        El panel se actualiza automáticamente cada viernes.<br>
        Si tienes alguna duda, contesta este email o escribe a apm@industrialshields.com.
      </p>
    </div>
  </div>
</body></html>"""

def send_email(to_name, to_email, html_body, subject):
    user = os.environ['GMAIL_USER']
    pwd  = os.environ['GMAIL_APP_PASS']
    sender = os.environ.get('REPORT_FROM') or user

    msg = MIMEMultipart('alternative')
    msg['Subject'] = subject
    msg['From']    = formataddr(("Comisiones · Industrial Shields", sender))
    msg['To']      = formataddr((to_name, to_email))
    msg['Date']    = formatdate(localtime=True)
    msg.attach(MIMEText(html_body, 'html', 'utf-8'))

    ctx = ssl.create_default_context()
    with smtplib.SMTP_SSL('smtp.gmail.com', 465, context=ctx) as s:
        s.login(user, pwd)
        s.sendmail(sender, [to_email], msg.as_string())
    print(f"   ✓ enviado a {to_name} <{to_email}>")

if __name__ == '__main__':
    summary = build_summary()
    today = date.today().strftime('%d %b %Y')
    subject = f"📊 Comisiones · actualización del viernes {today}"
    print(f"[send_emails] Enviando a {len(RECIPIENTS)} destinatarios…")
    for name, email in RECIPIENTS.items():
        url = f"{PAGES_BASE}/Comisiones_{_slug(name)}.html"
        body = build_html_body(name, summary, url)
        try:
            send_email(name, email, body, subject)
        except Exception as e:
            print(f"   ✗ ERROR enviando a {name}: {e}")
    print("[send_emails] DONE.")
