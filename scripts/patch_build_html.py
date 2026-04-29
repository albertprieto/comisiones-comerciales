#!/usr/bin/env python3
"""Aplica el refactor de la botonera de vistas (dropdown 'Comisiones') a scripts/build_html.py.

Idempotente: si el parche ya está aplicado, no hace nada.
"""
import os, sys

HERE = os.path.dirname(os.path.abspath(__file__))
TARGET = os.path.join(HERE, 'build_html.py')

with open(TARGET, encoding='utf-8') as f:
    src = f.read()

if 'view-menu-trigger' in src and 'initViewMenu' in src:
    print('[patch_build_html] Ya aplicado, no hago nada.')
    raise SystemExit(0)

# -----------------------------------------------------------------------------
# 1) CSS: reemplazar bloque .view-switch por .view-menu (dropdown)
# -----------------------------------------------------------------------------
OLD_CSS = '''  .view-switch{display:flex;gap:0;border:1px solid var(--line);
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
  }'''

NEW_CSS = '''  .view-switch{display:flex;gap:0;flex-shrink:0}
  /* Dropdown menu "Comisiones" con submenus */
  .view-menu{position:relative;display:inline-block}
  .view-menu-trigger{
    background:#fff;color:var(--navy);border:1px solid var(--line);border-radius:6px;
    padding:7px 14px;cursor:pointer;font-size:13px;font-weight:600;
    font-family:inherit;display:inline-flex;align-items:center;gap:8px;white-space:nowrap;
    transition:background 0.15s,border-color 0.15s,box-shadow 0.15s;
  }
  .view-menu-trigger:hover{background:var(--bg-alt)}
  .view-menu-trigger.open{background:var(--bg-alt);border-color:var(--accent);box-shadow:0 0 0 2px var(--accent-soft)}
  .view-menu-label{color:var(--accent);font-weight:700;letter-spacing:0.2px}
  .view-menu-current{color:var(--muted);font-weight:500;font-size:12px}
  .view-menu-current::before{content:"·";margin-right:6px;color:var(--line)}
  .view-menu-caret{color:var(--muted);font-size:9px;margin-left:2px;transition:transform .15s}
  .view-menu-trigger.open .view-menu-caret{transform:rotate(180deg)}
  .view-menu-panel{
    display:none;position:absolute;top:calc(100% + 4px);left:0;
    background:#fff;border:1px solid var(--line);border-radius:6px;
    box-shadow:0 6px 20px rgba(26,47,92,0.10);
    min-width:200px;overflow:hidden;z-index:60;
  }
  .view-menu-panel.open{display:block}
  .view-menu-panel button{
    background:#fff;color:var(--ink);border:0;border-bottom:1px solid var(--line);
    padding:10px 16px;cursor:pointer;font-size:13px;font-weight:500;
    font-family:inherit;text-align:left;width:100%;display:block;
    transition:background 0.12s,color 0.12s;
  }
  .view-menu-panel button:last-child{border-bottom:0}
  .view-menu-panel button:hover{background:var(--bg-alt);color:var(--navy)}
  .view-menu-panel button.active{background:var(--accent-soft);color:var(--accent);font-weight:600}'''

if OLD_CSS not in src:
    sys.exit('[patch_build_html] OLD_CSS not found, abort')
src = src.replace(OLD_CSS, NEW_CSS, 1)

# -----------------------------------------------------------------------------
# 2) HTML: reemplazar 6 botones por dropdown
# -----------------------------------------------------------------------------
OLD_HTML = '''  <div class="view-switch">
    <button id="btn-view-table" class="active">Tabla</button>
    <button id="btn-view-dash">Dashboard</button>
    <button id="btn-view-comm">Comisiones</button>
    <button id="btn-view-pay">PAGO</button>
    <button id="btn-view-tariffs">Tarifas</button>
    <button id="btn-view-rules">Condiciones</button>
  </div>'''

NEW_HTML = '''  <div class="view-switch">
    <div class="view-menu">
      <button type="button" id="btn-view-toggle" class="view-menu-trigger" aria-haspopup="true" aria-expanded="false">
        <span class="view-menu-label">Comisiones</span>
        <span class="view-menu-current" id="view-menu-current">Tabla</span>
        <span class="view-menu-caret">▼</span>
      </button>
      <div class="view-menu-panel" id="view-menu-panel" role="menu">
        <button id="btn-view-table" class="active" role="menuitem">Tabla</button>
        <button id="btn-view-dash"   role="menuitem">Dashboard</button>
        <button id="btn-view-comm"   role="menuitem">Comisiones</button>
        <button id="btn-view-pay"    role="menuitem">PAGO</button>
        <button id="btn-view-tariffs" role="menuitem">Tarifas</button>
        <button id="btn-view-rules"  role="menuitem">Condiciones</button>
      </div>
    </div>
  </div>'''

if OLD_HTML not in src:
    sys.exit('[patch_build_html] OLD_HTML not found, abort')
src = src.replace(OLD_HTML, NEW_HTML, 1)

# -----------------------------------------------------------------------------
# 3) JS: insertar bloque initViewMenu despues de los addEventListener show*
# -----------------------------------------------------------------------------
OLD_JS_ANCHOR = '''bT.addEventListener("click",  showTable);
bD.addEventListener("click",  showDash);
bC.addEventListener("click",  showComm);
bP.addEventListener("click",  showPay);
bTa.addEventListener("click", showTariffs);
bR.addEventListener("click",  showRules);'''

NEW_JS_BLOCK = OLD_JS_ANCHOR + '''

// =============================================================================
// VIEW MENU (dropdown "Comisiones ▾ <vista actual>")
// =============================================================================
(function initViewMenu(){
  const trigger = document.getElementById("btn-view-toggle");
  const panel   = document.getElementById("view-menu-panel");
  const current = document.getElementById("view-menu-current");
  if (!trigger || !panel) return;

  function setOpen(open){
    panel.classList.toggle("open", open);
    trigger.classList.toggle("open", open);
    trigger.setAttribute("aria-expanded", open ? "true" : "false");
  }
  trigger.addEventListener("click", e => {
    e.stopPropagation();
    setOpen(!panel.classList.contains("open"));
  });
  document.addEventListener("click", e => {
    if (!e.target.closest(".view-menu")) setOpen(false);
  });
  document.addEventListener("keydown", e => {
    if (e.key === "Escape") setOpen(false);
  });
  panel.querySelectorAll("button").forEach(btn => {
    btn.addEventListener("click", () => {
      if (current) current.textContent = btn.textContent.trim();
      setOpen(false);
    });
  });
})();'''

if OLD_JS_ANCHOR not in src:
    sys.exit('[patch_build_html] OLD_JS_ANCHOR not found, abort')
src = src.replace(OLD_JS_ANCHOR, NEW_JS_BLOCK, 1)

# -----------------------------------------------------------------------------
with open(TARGET, 'w', encoding='utf-8') as f:
    f.write(src)
print('[patch_build_html] OK - dropdown applied')
