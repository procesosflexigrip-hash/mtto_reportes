"""
generar_formulario.py
----------------------
Regenera captura_mantenimiento.html a partir del Excel de listas (validacion.xlsx).

USO:
    python3 generar_formulario.py
    python3 generar_formulario.py ruta/a/otro_validacion.xlsx

Cada vez que agregues/quites una máquina, técnico, área, etc. en el Excel
(pestaña "validacion", una columna por lista), solo corre este script de nuevo
y se genera un captura_mantenimiento.html actualizado.

Columnas esperadas en la pestaña "validacion" del Excel:
    AREA, MAQUINA, SOLICITANTE, TECNICO, PIN, PRIORIDAD, FALLA, TIPO, PARO MAQUINA, BLOQUE Y CANDADEO, REFACCION
(el orden no importa, los nombres de columna sí deben coincidir exactamente.
 TECNICO y PIN deben estar en la MISMA fila para cada técnico — es su PIN individual.)

CONEXIÓN A GOOGLE SHEETS:
    Antes de correr este script, edita la línea SHEET_WEBAPP_URL más abajo
    y pega ahí la URL de tu Apps Script publicado (ver instrucciones en el
    archivo apps_script_codigo.gs). Si la dejas sin editar, el formulario
    sigue funcionando pero solo guarda datos mientras la pestaña esté abierta.
"""
import json
import sys
import pandas as pd

EXCEL_PATH = sys.argv[1] if len(sys.argv) > 1 else "validacion.xlsx"
HOJA = "validacion"
SALIDA = "index.html"

COLUMNAS_ESPERADAS = ["AREA", "MAQUINA", "SOLICITANTE", "TECNICO", "PRIORIDAD", "FALLA", "TIPO", "PARO MAQUINA", "BLOQUE Y CANDADEO", "REFACCION"]

print(f"Leyendo listas desde: {EXCEL_PATH} (hoja '{HOJA}')")
df = pd.read_excel(EXCEL_PATH, sheet_name=HOJA)

faltantes = [c for c in COLUMNAS_ESPERADAS if c not in df.columns]
if faltantes:
    print(f"⚠ AVISO: faltan estas columnas en el Excel: {faltantes}")
    print("  El formulario se generará igual, pero esas listas quedarán vacías.")

lists = {}
for col in COLUMNAS_ESPERADAS:
    if col in df.columns:
        vals = [str(v).strip() for v in df[col].dropna().tolist()]
        seen = set()
        out = []
        for v in vals:
            if v not in seen:
                seen.add(v)
                out.append(v)
        lists[col] = out
    else:
        lists[col] = []

for col, vals in lists.items():
    print(f"  {col}: {len(vals)} valores")

# Mantiene utilizable el cierre mientras se llena por primera vez el catálogo.
if not lists["REFACCION"]:
    lists["REFACCION"] = ["SIN REFACCIONES"]
    print("  REFACCION: se usará temporalmente 'SIN REFACCIONES'; agrega las piezas reales al Excel.")

# TECNICO + PIN van alineados por fila (cada técnico con su propio PIN individual)
tecnicos_pin = []
if "TECNICO" in df.columns and "PIN" in df.columns:
    for _, fila in df[["TECNICO", "PIN"]].dropna().iterrows():
        nombre = str(fila["TECNICO"]).strip()
        pin_val = fila["PIN"]
        if isinstance(pin_val, float) and pin_val.is_integer():
            pin_str = str(int(pin_val))
        else:
            pin_str = str(pin_val).strip()
        tecnicos_pin.append({"nombre": nombre, "pin": pin_str})
    print(f"  TECNICO+PIN: {len(tecnicos_pin)} técnicos con PIN individual")
else:
    print("  ⚠ AVISO: falta la columna PIN junto a TECNICO — los técnicos no van a poder entrar.")
lists["TECNICOS_PIN"] = tecnicos_pin

lists_js = json.dumps(lists, ensure_ascii=False)

turnos = ["1°", "2°", "3°"]  # <-- si cambian los turnos, edita esta línea
turnos_js = json.dumps(turnos, ensure_ascii=False)

GERENTE_MANTENIMIENTO = "JORGE ADRIAN VEGA BARON"  # <-- solo este solicitante puede elegir el Tipo; todos los demás quedan como CORRECTIVO

SUPERVISOR_PIN = "2580"  # <-- CAMBIA ESTE PIN. Da acceso a asignar/cerrar servicios y ver todos los registros.
SHEET_WEBAPP_URL = "https://script.google.com/macros/s/AKfycbzqFLePhr_RQo1aZeHO3BLJZDx6RCcmGekwBZSmhpPW0K-7FGX293qtUt1pKKDhGQjV/exec"  # <-- ya configurada

html = """<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1">
<title>REGISTRO DE MANTENIMIENTO PREVENTIVO Y CORRECTIVO (FOR-MAN-01-02-03)</title>
<style>
  :root{
    --bg: #12181f;
    --panel: #1a222c;
    --panel2: #202b37;
    --line: #2c3946;
    --text: #eef2f6;
    --muted: #8ea0b3;
    --accent: #ff8a3d;
    --accent2: #3fb8af;
    --danger: #ef5b5b;
    --ok: #4caf7d;
    --warn: #e0a53c;
    --radius: 10px;
    font-size: 15px;
  }
  *{box-sizing:border-box;}
  html,body{margin:0;padding:0;background:var(--bg);color:var(--text);
    font-family: 'Segoe UI', Roboto, -apple-system, Arial, sans-serif;}
  body{padding:18px 18px 60px;}
  .wrap{max-width:1180px;margin:0 auto;}

  header.top{display:flex;align-items:center;justify-content:space-between;
    gap:16px;margin-bottom:18px;flex-wrap:wrap;border-bottom:1px solid var(--line);padding-bottom:16px;}
  header.top .brand{display:flex;align-items:center;gap:12px;}
  header.top .mark{width:40px;height:40px;border-radius:8px;background:linear-gradient(135deg,var(--accent),#c85e1f);
    display:flex;align-items:center;justify-content:center;font-weight:800;color:#fff;font-size:18px;flex:none;}
  header.top h1{font-size:19px;margin:0;letter-spacing:.2px;}
  header.top p{margin:2px 0 0;color:var(--muted);font-size:12.5px;}
  .clock{font-variant-numeric:tabular-nums;color:var(--muted);font-size:12.5px;text-align:right;}
  .clock b{color:var(--text);font-size:14px;display:block;}

  .tabs{display:flex;gap:8px;margin-bottom:18px;flex-wrap:wrap;}
  .tab-btn{flex:1;min-width:200px;background:var(--panel);border:1px solid var(--line);color:var(--muted);
    padding:14px 16px;border-radius:var(--radius);cursor:pointer;text-align:left;font-family:inherit;}
  .tab-btn .t{display:block;font-size:14px;font-weight:800;color:var(--text);}
  .tab-btn .s{display:block;font-size:11.5px;margin-top:2px;}
  .tab-btn.active{border-color:var(--accent);background:linear-gradient(180deg,rgba(255,138,61,.10),transparent);}
  .tab-btn.active .t{color:var(--accent);}
  .badge{background:var(--warn);color:#241a04;font-size:11px;font-weight:800;
    border-radius:20px;padding:2px 9px;margin-left:6px;}
  .tab-btn .badge{float:right;background:var(--warn);color:#241a04;font-size:11px;font-weight:800;
    border-radius:20px;padding:2px 9px;}

  .view{display:none;}
  .view.active{display:block;}

  .role-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(220px,1fr));gap:14px;margin-top:6px;}
  .role-card{background:var(--panel);border:1px solid var(--line);border-radius:var(--radius);
    padding:28px 20px;cursor:pointer;text-align:center;display:flex;flex-direction:column;align-items:center;
    gap:8px;font-family:inherit;transition:border-color .15s, transform .08s;}
  .role-card:hover{border-color:var(--accent);transform:translateY(-2px);}
  .role-card:active{transform:scale(.98);}
  .role-card .ic{font-size:34px;}
  .role-card .t{font-size:15.5px;font-weight:800;color:var(--text);}
  .role-card .s{font-size:12px;color:var(--muted);}

  .back-link{background:none;border:none;color:var(--muted);cursor:pointer;font-size:13px;
    padding:6px 0;margin-bottom:14px;font-family:inherit;}
  .back-link:hover{color:var(--accent);}

  .kpis{display:grid;grid-template-columns:repeat(auto-fit,minmax(140px,1fr));gap:10px;margin-bottom:18px;}
  .kpi{background:var(--panel);border:1px solid var(--line);border-radius:var(--radius);padding:12px 14px;}
  .kpi .n{font-size:21px;font-weight:800;color:var(--accent);font-variant-numeric:tabular-nums;}
  .kpi .l{font-size:11px;color:var(--muted);text-transform:uppercase;letter-spacing:.4px;margin-top:2px;}
  .kpi.warn{border-color:var(--warn);}
  .kpi.warn .n{color:var(--warn);}
  .ticket.ticket-alerta{border-color:var(--warn);}

  .card{background:var(--panel);border:1px solid var(--line);border-radius:var(--radius);padding:18px;margin-bottom:18px;}
  .card h2{font-size:14px;margin:0 0 14px;text-transform:uppercase;letter-spacing:.6px;color:var(--accent2);
    display:flex;align-items:center;gap:8px;}
  .card h2 .num{width:22px;height:22px;border-radius:50%;background:var(--accent2);color:#0c1116;
    display:flex;align-items:center;justify-content:center;font-size:12px;font-weight:800;}
  .card p.hint{color:var(--muted);font-size:12.5px;margin:-8px 0 16px;}

  .grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(200px,1fr));gap:12px 14px;}
  .field{display:flex;flex-direction:column;gap:5px;position:relative;}
  .field.span2{grid-column:span 2;}
  .field label{font-size:11.5px;color:var(--muted);text-transform:uppercase;letter-spacing:.4px;}
  .field label .req{color:var(--danger);}
  .char-counter{font-size:11px;color:var(--muted);text-align:right;margin-top:-2px;}
  .refacciones-editor{display:flex;flex-direction:column;gap:8px;}
  .refaccion-row{display:grid;grid-template-columns:minmax(180px,1fr) 90px auto;gap:8px;align-items:end;}
  .refaccion-row .field{min-width:0;}
  .refaccion-row button{padding:10px 12px;}
  @media(max-width:600px){
    .refaccion-row{grid-template-columns:minmax(0,1fr) 78px auto;}
  }

  .ac-list{position:absolute;top:100%;left:0;right:0;z-index:60;display:none;margin-top:4px;
    max-height:240px;overflow-y:auto;background:var(--panel2);border:1px solid var(--accent);
    border-radius:8px;box-shadow:0 8px 20px rgba(0,0,0,.45);}
  .ac-list.show{display:block;}
  .ac-item{padding:9px 11px;font-size:14px;cursor:pointer;border-bottom:1px solid var(--line);}
  .ac-item:last-child{border-bottom:none;}
  .ac-item:hover, .ac-item.active{background:var(--panel);color:var(--accent);}
  .ac-hint{padding:7px 11px;font-size:11.5px;color:var(--muted);border-top:1px solid var(--line);}
  input, select, textarea{
    background:var(--panel2);border:1px solid var(--line);color:var(--text);
    border-radius:8px;padding:10px 11px;font-size:14.5px;font-family:inherit;width:100%;
  }
  input:focus, select:focus, textarea:focus{outline:none;border-color:var(--accent);}
  textarea{resize:vertical;min-height:44px;}
  select{-webkit-appearance:none;appearance:none;
    background-image:url("data:image/svg+xml;utf8,<svg xmlns='http://www.w3.org/2000/svg' width='10' height='6'><path d='M0 0l5 6 5-6z' fill='%238ea0b3'/></svg>");
    background-repeat:no-repeat;background-position:right 12px center;padding-right:30px;}

  .divider{border:none;border-top:1px dashed var(--line);margin:16px 0;}

  .autotime{display:flex;align-items:center;gap:10px;background:var(--panel2);border:1px solid var(--line);
    border-radius:8px;padding:12px 16px;margin-bottom:16px;}
  .autotime .ico{font-size:20px;}
  .autotime b{font-variant-numeric:tabular-nums;color:var(--accent);}
  .autotime .sub{color:var(--muted);font-size:11.5px;}

  .codigo-box{background:var(--panel2);border:2px solid var(--accent);border-radius:10px;
    padding:24px;text-align:center;margin:6px 0 4px;}
  .codigo-num{font-size:44px;font-weight:800;letter-spacing:10px;color:var(--accent);
    font-variant-numeric:tabular-nums;}
  .codigo-sub{color:var(--muted);font-size:12.5px;margin-top:6px;}

  .duration-box{display:flex;gap:14px;align-items:center;background:var(--panel2);border:1px solid var(--line);
    border-radius:8px;padding:12px 16px;flex-wrap:wrap;}
  .duration-box .big{font-size:26px;font-weight:800;color:var(--accent);font-variant-numeric:tabular-nums;}
  .duration-box .sub{font-size:11.5px;color:var(--muted);}
  .duration-box .pill{background:#0c1116;border:1px solid var(--line);border-radius:20px;padding:5px 12px;font-size:12px;color:var(--muted);}

  .actions{display:flex;gap:10px;margin-top:18px;flex-wrap:wrap;}
  button{cursor:pointer;border:none;border-radius:8px;padding:12px 20px;font-size:14px;font-weight:700;
    font-family:inherit;transition:transform .08s ease, opacity .15s;}
  button:active{transform:scale(.97);}
  button:disabled{opacity:.4;cursor:not-allowed;}
  .btn-primary{background:var(--accent);color:#1a0f04;}
  .btn-primary:hover{opacity:.92;}
  .btn-ghost{background:transparent;border:1px solid var(--line);color:var(--muted);}
  .btn-ghost:hover{color:var(--text);border-color:var(--muted);}
  .btn-export{background:var(--accent2);color:#06231f;}
  .btn-start{background:var(--accent2);color:#06231f;padding:8px 14px;font-size:12.5px;}
  .btn-finish{background:var(--accent);color:#1a0f04;padding:8px 14px;font-size:12.5px;}

  .msg{font-size:13px;padding:10px 14px;border-radius:8px;margin-top:12px;display:none;}
  .msg.ok{display:block;background:rgba(76,175,125,.12);border:1px solid var(--ok);color:#b6f0cf;}
  .msg.err{display:block;background:rgba(239,91,91,.12);border:1px solid var(--danger);color:#ffc7c7;}
  .aviso-toast{position:fixed;bottom:20px;right:16px;z-index:80;max-width:360px;background:var(--panel2);
    border:1px solid var(--accent2);color:var(--text);padding:12px 16px;border-radius:8px;font-size:13px;
    display:none;}
  .aviso-toast.show{display:block;}

  .pin-gate{max-width:340px;margin:10px auto;text-align:center;padding:26px 20px;}
  .pin-gate .lock{font-size:32px;margin-bottom:6px;}
  .pin-gate input{text-align:center;font-size:22px;letter-spacing:6px;margin:14px 0;}

  .subhead{font-size:12px;color:var(--muted);text-transform:uppercase;letter-spacing:.5px;margin:18px 0 8px;
    display:flex;align-items:center;gap:8px;}
  .subhead .dot{width:8px;height:8px;border-radius:50%;}
  .subhead .dot.wait{background:var(--warn);}
  .subhead .dot.prog{background:var(--accent2);}

  .ticket-list{display:flex;flex-direction:column;gap:8px;margin-bottom:6px;}
  .ticket{background:var(--panel2);border:1px solid var(--line);border-radius:8px;padding:12px 14px;
    display:flex;justify-content:space-between;align-items:center;gap:10px;flex-wrap:wrap;}
  .ticket .left b{font-size:13.5px;}
  .ticket .left .meta{color:var(--muted);font-size:11.5px;margin-top:2px;}
  .ticket .right{font-size:11px;color:var(--muted);text-align:right;display:flex;align-items:center;gap:10px;}
  .ticket .elapsed{font-variant-numeric:tabular-nums;color:var(--accent2);font-weight:700;}

  .table-toolbar{display:flex;justify-content:space-between;align-items:center;gap:10px;flex-wrap:wrap;margin-bottom:12px;}
  .table-toolbar input[type=text]{max-width:260px;}
  .table-scroll{overflow-x:auto;border:1px solid var(--line);border-radius:8px;}
  table{border-collapse:collapse;width:100%;font-size:12.5px;min-width:1000px;}
  thead th{position:sticky;top:0;background:#0e141b;color:var(--muted);text-align:left;
    padding:9px 10px;font-size:10.5px;text-transform:uppercase;letter-spacing:.4px;border-bottom:1px solid var(--line);white-space:nowrap;}
  tbody td{padding:8px 10px;border-bottom:1px solid var(--line);white-space:nowrap;color:var(--text);}
  tbody tr:hover{background:var(--panel2);}
  tbody tr:last-child td{border-bottom:none;}
  .tag{display:inline-block;padding:2px 8px;border-radius:20px;font-size:10.5px;font-weight:700;}
  .tag.Alta{background:rgba(239,91,91,.18);color:#ffb0b0;}
  .tag.Moderada{background:rgba(224,165,60,.18);color:#f3cb87;}
  .tag.Baja{background:rgba(76,175,125,.18);color:#a6e6c4;}
  .tag.Pendiente{background:rgba(224,165,60,.18);color:#f3cb87;}
  .tag.Enreparacion{background:rgba(63,184,175,.18);color:#a6e6e0;}
  .tag.Pausado{background:rgba(142,160,179,.18);color:#c3d0dc;}
  .tag.Asignado{background:rgba(63,184,175,.12);color:#7dd3c9;}
  .tag.Cerrado{background:rgba(76,175,125,.18);color:#a6e6c4;}
  .tag.Liberado{background:rgba(76,175,125,.32);color:#d3f7e4;}
  .empty{padding:24px;text-align:center;color:var(--muted);font-size:13px;}

  .ticket-col{flex-direction:column;align-items:stretch;}
  .tec-panel{margin-top:10px;padding-top:10px;border-top:1px dashed var(--line);}
  .tec-panel .tec-title{font-size:11.5px;color:var(--muted);text-transform:uppercase;letter-spacing:.4px;margin-bottom:8px;}
  .tec-checks{display:flex;flex-wrap:wrap;gap:6px 10px;margin-bottom:12px;}
  .tec-checks label{display:flex;align-items:center;gap:5px;font-size:12.5px;color:var(--text);cursor:pointer;
    background:#161d24;border:1px solid var(--line);border-radius:6px;padding:5px 9px;}
  .tec-checks input{width:auto;}
  .btn-row{display:flex;gap:8px;flex-wrap:wrap;}
  .tec-names{color:var(--accent2);font-size:11.5px;margin-top:2px;display:flex;flex-wrap:wrap;gap:6px;align-items:center;}
  .tec-chip{display:inline-flex;align-items:center;gap:4px;background:#0c1116;border:1px solid var(--line);
    border-radius:20px;padding:2px 4px 2px 10px;}
  .tec-chip button{background:none;border:none;color:var(--muted);cursor:pointer;padding:2px 6px;font-size:11px;border-radius:50%;}
  .tec-chip button:hover{color:var(--danger);}
  .tec-checks label.tec-busy{opacity:.45;cursor:not-allowed;}
  .busy-tag{font-size:9px;color:var(--danger);text-transform:uppercase;margin-left:3px;}
  .btn-add{background:transparent;border:1px dashed var(--line);color:var(--muted);padding:6px 12px;font-size:11.5px;}
  .btn-add:hover{color:var(--text);border-color:var(--accent2);}
  .btn-pause{background:transparent;border:1px solid var(--warn);color:var(--warn);padding:8px 14px;font-size:12.5px;}
  .mini-x{background:none;border:none;color:var(--danger);cursor:pointer;font-size:11px;padding:0 3px;}
  .mini-x:hover{text-decoration:underline;}
  .tec-chk-wrap label.disabled{opacity:.4;cursor:not-allowed;}

  footer.note{margin-top:22px;font-size:12px;color:var(--muted);line-height:1.6;border-top:1px solid var(--line);padding-top:14px;}
  footer.note b{color:var(--text);}

  @media (max-width:640px){
    .field.span2{grid-column:span 1;}
    .duration-box{flex-direction:column;align-items:flex-start;}
  }
</style>
</head>
<body>
<div class="wrap">
<div class="aviso-toast" id="avisoToast"></div>

  <header class="top">
    <div class="brand">
      <div class="mark">FG</div>
      <div>
        <h1>REGISTRO DE MANTENIMIENTO PREVENTIVO Y CORRECTIVO (FOR-MAN-01-02-03)</h1>
      </div>
    </div>
    <div class="clock">
      <b id="clockTime">--:--</b>
      <span id="clockDate">-- - --</span>
    </div>
  </header>

  <!-- ==================== LANDING: SELECCIÓN DE ROL ==================== -->
  <div class="view active" id="viewLanding">
    <div class="role-grid">
      <button class="role-card" onclick="irA('reporte')">
        <span class="ic">🧾</span>
        <span class="t">Generar reporte</span>
        <span class="s">Cualquier persona — operador, supervisor, gerente</span>
      </button>
      <button class="role-card" onclick="irA('supervisor')">
        <span class="ic">👷</span>
        <span class="t">Supervisor mantenimiento</span>
        <span class="s">Asignar técnicos y cerrar servicios — requiere PIN</span>
      </button>
      <button class="role-card" onclick="irA('tecnico')">
        <span class="ic">🔧</span>
        <span class="t">Técnicos</span>
        <span class="s">Consultar tus asignaciones — requiere PIN</span>
      </button>
    </div>
  </div>

  <!-- ==================== VISTA REPORTE (abierta a todos, sin PIN) ==================== -->
  <div class="view" id="viewReporte">
    <button class="back-link" onclick="cerrarCodigoCard(); irA('landing')">← Volver</button>
    <div class="card">
      <h2><span class="num">1</span> Reportar falla / solicitar servicio</h2>

      <div class="autotime">
        <span class="ico">🕒</span>
        <div>
          <div>Se registrará como fecha/hora de solicitud: <b id="previewNow">--:--</b></div>
        </div>
      </div>

      <div class="grid">
        <div class="field">
          <label>Área <span class="req">*</span></label>
          <select id="f_area" required><option value="">Selecciona…</option></select>
        </div>
        <div class="field">
          <label>Máquina <span class="req">*</span></label>
          <select id="f_maquina" required><option value="">Selecciona…</option></select>
        </div>
        <div class="field">
          <label>Solicitante <span class="req">*</span></label>
          <input type="text" id="f_solicitante" placeholder="Escribe el nombre…" required autocomplete="off" oninput="chequearGerente()">
          <div class="ac-list" id="ac_f_solicitante"></div>
        </div>
        <div class="field">
          <label>Prioridad <span class="req">*</span></label>
          <select id="f_prioridad" required><option value="">Selecciona…</option></select>
        </div>
        <div class="field" id="campoTipoGerente" style="display:none;">
          <label>Tipo <span class="req">*</span></label>
          <select id="f_tipo"><option value="">Selecciona…</option></select>
        </div>
        <div class="field span2">
          <label>Problema reportado</label>
          <textarea id="f_problema" maxlength="200" placeholder="Descripción de la falla / motivo de la solicitud" oninput="actualizarContadorProblema()"></textarea>
          <div class="char-counter" id="problemaContador">0/200</div>
        </div>
      </div>
      <div class="actions">
        <button class="btn-primary" id="btnEnviarSolicitud" onclick="crearSolicitud()">Enviar solicitud</button>
        <button class="btn-ghost" onclick="limpiarFormSolicitud()">Limpiar</button>
      </div>
      <div class="msg" id="solMsg"></div>
    </div>

    <div class="card" id="codigoCard" style="display:none;">
      <h2><span class="num">✔</span> Solicitud enviada</h2>
      <p class="hint">Anota este código. Dáselo al técnico cuando llegue a la máquina — es la forma de confirmar que se presentó en sitio.</p>
      <div class="codigo-box">
        <div class="codigo-num" id="codigoNum">----</div>
        <div class="codigo-sub">Código de confirmación — servicio #<span id="codigoTicketId"></span></div>
      </div>
      <div class="actions">
        <button class="btn-primary" onclick="cerrarCodigoCard()">Entendido</button>
      </div>
    </div>

    <div class="card">
      <h2><span class="num">✔</span> Liberar un servicio ya cerrado</h2>
      <p class="hint">Aquí aparecen todos los servicios cerrados. Escribe y selecciona tu nombre antes de liberar el área.</p>
      <div class="grid">
        <div class="field">
          <label>Nombre de quien libera <span class="req">*</span></label>
          <input type="text" id="lib_nombre" placeholder="Escribe y selecciona tu nombre…" autocomplete="off">
          <div class="ac-list" id="ac_lib_nombre"></div>
        </div>
      </div>
      <div class="ticket-list" id="liberarLista" style="margin-top:14px;"></div>
      <div class="empty" id="liberarEmpty" style="display:none;">No hay servicios cerrados pendientes de liberar.</div>
    </div>
  </div>

  <!-- ==================== VISTA SUPERVISOR (con PIN) ==================== -->
  <div class="view" id="viewSupervisor">
    <button class="back-link" onclick="irA('landing')">← Volver</button>

    <div class="card pin-gate" id="pinGateSupervisor">
      <div class="lock">🔒</div>
      <h2 style="justify-content:center;">Acceso supervisor</h2>
      <p class="hint">Ingresa el PIN de supervisor para continuar.</p>
      <input type="password" id="pinInputSupervisor" inputmode="numeric" maxlength="8" placeholder="••••">
      <div class="actions" style="justify-content:center;">
        <button class="btn-primary" onclick="validarPinSupervisor()">Entrar</button>
      </div>
      <div class="msg" id="pinMsgSupervisor"></div>
    </div>

    <div id="supervisorContent" style="display:none;">
      <div class="kpis" id="kpiRow">
        <div class="kpi"><div class="n" id="kpiTotal">0</div><div class="l">Total</div></div>
        <div class="kpi"><div class="n" id="kpiPendientes">0</div><div class="l">Esperando asignar</div></div>
        <div class="kpi"><div class="n" id="kpiAsignados">0</div><div class="l">Por confirmar</div></div>
        <div class="kpi"><div class="n" id="kpiEnCurso">0</div><div class="l">En reparación</div></div>
        <div class="kpi"><div class="n" id="kpiPausados">0</div><div class="l">Pausados</div></div>
        <div class="kpi" id="kpiPausadosViejosBox"><div class="n" id="kpiPausadosViejos">0</div><div class="l">Pausados &gt; 8 h</div></div>
        <div class="kpi"><div class="n" id="kpiEspera">0.0 h</div><div class="l">Espera promedio</div></div>
        <div class="kpi"><div class="n" id="kpiReparacion">0.0 h</div><div class="l">Reparación promedio</div></div>
      </div>

      <div class="card">
        <h2><span class="num">2</span> Servicios</h2>
        <p class="hint">Asigna técnicos cuando un servicio vaya a comenzar, y finalízalo cuando termine. La hora la toma el sistema, no se escribe.</p>

        <div class="subhead"><span class="dot wait"></span> Esperando asignar técnico</div>
        <div class="ticket-list" id="ticketListPend"></div>
        <div class="empty" id="emptyPend" style="display:none;">No hay servicios esperando técnico 🎉</div>

        <div class="subhead"><span class="dot" style="background:var(--accent2);"></span> Asignados — esperando confirmación en sitio</div>
        <div class="ticket-list" id="ticketListAsignados"></div>
        <div class="empty" id="emptyAsignados" style="display:none;">No hay servicios esperando confirmación.</div>

        <div class="subhead"><span class="dot prog"></span> En reparación</div>
        <div class="ticket-list" id="ticketListProg"></div>
        <div class="empty" id="emptyProg" style="display:none;">Nadie está reparando algo en este momento.</div>

        <div class="subhead"><span class="dot" style="background:var(--muted);"></span> Pausados — esperando retomar <span id="pausadosViejosHint" style="text-transform:none;letter-spacing:0;font-weight:400;"></span></div>
        <div class="ticket-list" id="ticketListPaused"></div>
        <div class="empty" id="emptyPaused" style="display:none;">No hay servicios pausados.</div>
      </div>

      <div class="card" id="cierreCard" style="display:none;">
        <h2><span class="num">3</span> Finalizar reparación <span id="cierreTicketId" style="color:var(--muted);font-weight:400;text-transform:none;letter-spacing:0;"></span></h2>

        <div class="autotime">
          <span class="ico">🕒</span>
          <div>
            <div>Se registrará como fecha/hora de fin: <b id="previewNowFin">--:--</b></div>
            <div class="sub">Automático — se toma al dar clic en "Guardar cierre".</div>
          </div>
        </div>

        <div class="grid">
          <div class="field">
            <label>Falla <span class="req">*</span></label>
            <select id="f_falla" required><option value="">Selecciona…</option></select>
          </div>
          <div class="field">
            <label>Turno que atendió <span class="req">*</span></label>
            <select id="f_turno" required><option value="">Selecciona…</option></select>
          </div>
          <div class="field">
            <label>Paro máquina <span class="req">*</span></label>
            <select id="f_paro" required><option value="">Selecciona…</option></select>
          </div>
          <div class="field">
            <label>Bloqueo y candadeo <span class="req">*</span></label>
            <select id="f_bloqueo" required><option value="">Selecciona…</option></select>
          </div>
          <div class="field span2">
            <label>Mantenimiento realizado</label>
            <textarea id="f_mantenimiento" placeholder="Trabajo / reparación efectuada"></textarea>
          </div>
          <div class="field span2">
            <label>Refacciones utilizadas <span class="req">*</span></label>
            <div class="refacciones-editor" id="refaccionesEditor"></div>
            <div class="actions" style="margin-top:8px;">
              <button type="button" class="btn-ghost" onclick="agregarRefaccion()">+ Agregar refacción</button>
            </div>
            <p class="hint" style="margin:4px 0 0;">Selecciona una refacción de la lista e indica la cantidad. Si no se usaron piezas, elige “SIN REFACCIONES”.</p>
          </div>
          <div class="field span2">
            <label>Observaciones</label>
            <textarea id="f_observaciones" placeholder="Notas adicionales"></textarea>
          </div>
        </div>

        <hr class="divider">
        <div class="duration-box">
          <div>
            <div class="big" id="durBig">0.00 h</div>
            <div class="sub">Tiempo de reparación (pausas descontadas)</div>
          </div>
          <div class="pill" id="durEspera">Espera solicitud → sitio: —</div>
          <div class="pill" id="durInicio">Inicio: —</div>
        </div>

        <div class="actions">
          <button class="btn-primary" id="btnGuardarCierre" onclick="cerrarServicio()">Guardar cierre</button>
          <button class="btn-ghost" onclick="cancelarCierre()">Cancelar</button>
        </div>
        <div class="msg" id="cierreMsg"></div>
      </div>

      <div class="card">
        <h2><span class="num">📋</span> Resumen de turno (para WhatsApp)</h2>
        <p class="hint">Arma el reporte del turno con los servicios cerrados o ya liberados. La fecha es el día en que <b>empezó</b> el turno (el 3° va de 22:00 a 07:00 del día siguiente).</p>
        <div class="grid">
          <div class="field">
            <label>Turno</label>
            <select id="r_turno"><option value="">Todos</option></select>
          </div>
          <div class="field">
            <label>Fecha de inicio del turno</label>
            <input type="date" id="r_fecha">
          </div>
        </div>
        <p class="hint" id="resumenTurnoHint" style="margin-top:8px;">El 3° de una fecha incluye lo cerrado esa noche y la madrugada siguiente (hasta las 07:00).</p>
        <div class="actions">
          <button class="btn-primary" onclick="generarResumen()">Generar resumen</button>
        </div>
        <textarea id="resumenTexto" readonly style="display:none;min-height:200px;margin-top:14px;font-family:monospace;font-size:12.5px;white-space:pre-wrap;"></textarea>
        <div class="actions" id="resumenActions" style="display:none;">
          <button class="btn-export" onclick="copiarResumen()">📋 Copiar</button>
        </div>
        <div class="msg" id="resumenMsg"></div>
      </div>

      <div class="card">
        <h2><span class="num">📋</span> Todos los registros</h2>
        <div class="table-toolbar">
          <input type="text" id="searchBox" placeholder="Buscar por máquina, área, solicitante…" oninput="renderTable()">
          <div style="display:flex;gap:8px;">
            <button class="btn-export" onclick="exportarCSV()">⬇ Exportar a Excel (CSV)</button>
            <button class="btn-ghost" onclick="cargarRegistros(true)">↻ Actualizar</button>
          </div>
        </div>
        <div class="table-scroll">
          <table>
            <thead>
              <tr>
                <th>ID</th><th>Estado</th><th>Fecha sol.</th><th>Área</th><th>Máquina</th><th>Solicitante</th>
                <th>Prioridad</th><th>Técnicos vigentes</th><th>Espera sol.→sitio (h)</th><th>Reparación (h)</th><th>Turno</th><th>Paro</th>
              </tr>
            </thead>
            <tbody id="tbody"></tbody>
          </table>
          <div class="empty" id="emptyMsg" style="display:none;">Aún no hay registros.</div>
        </div>
      </div>
    </div>
  </div>

  <!-- ==================== VISTA TÉCNICOS (con PIN, solo consulta) ==================== -->
  <div class="view" id="viewTecnico">
    <button class="back-link" onclick="salirTecnico(); irA('landing')">← Volver</button>

    <div class="card pin-gate" id="pinGateTecnico">
      <div class="lock">🔒</div>
      <h2 style="justify-content:center;">Acceso técnicos</h2>
      <p class="hint">Ingresa tu PIN personal para continuar.</p>
      <input type="password" id="pinInputTecnico" inputmode="numeric" maxlength="8" placeholder="••••">
      <div class="actions" style="justify-content:center;">
        <button class="btn-primary" onclick="validarPinTecnico()">Entrar</button>
      </div>
      <div class="msg" id="pinMsgTecnico"></div>
    </div>

    <div id="tecnicoContent" style="display:none;">
      <div class="card">
        <h2><span class="num">🔧</span> Hola, <span id="tecnicoSaludoNombre"></span></h2>
        <p class="hint">Solo puedes consultar tus propias asignaciones — esta vista no permite modificar nada.</p>
      </div>

      <div class="card" id="tecnicoPorConfirmar" style="display:none;">
        <h2><span class="num">📍</span> Por confirmar en sitio</h2>
        <p class="hint">Pídele el código al solicitante cuando llegues a la máquina.</p>
        <div class="ticket-list" id="tecnicoListaPorConfirmar"></div>
        <div class="empty" id="tecnicoEmptyPorConfirmar" style="display:none;">No tienes servicios pendientes por confirmar.</div>
      </div>

      <div class="card" id="tecnicoAsignados" style="display:none;">
        <h2><span class="num">📌</span> Asignado actualmente</h2>
        <div class="ticket-list" id="tecnicoListaActivos"></div>
        <div class="empty" id="tecnicoEmptyActivos" style="display:none;">No tienes servicios asignados en este momento.</div>
      </div>

      <div class="card" id="tecnicoCerrados" style="display:none;">
        <h2><span class="num">✔</span> Cerrados recientes</h2>
        <div class="ticket-list" id="tecnicoListaCerrados"></div>
      </div>
    </div>
  </div>

</div>

<script>
const LISTS = __LISTS_JSON__;
const TURNOS = __TURNOS_JSON__;
const STORAGE_KEY = 'mant-registros-flexigrip';
const SUPERVISOR_PIN = "__SUPERVISOR_PIN__"; // <-- CAMBIA ESTE PIN
const GERENTE_MANTENIMIENTO = "__GERENTE_MANTENIMIENTO__"; // <-- solo él puede elegir Tipo distinto de CORRECTIVO
const SHEET_WEBAPP_URL = "__SHEET_WEBAPP_URL__"; // <-- pega aquí la URL de tu Apps Script (termina en /exec)

const DIAS = ['Domingo','Lunes','Martes','Miércoles','Jueves','Viernes','Sábado'];
const MESES = ['Enero','Febrero','Marzo','Abril','Mayo','Junio','Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre'];

let registros = [];
let supervisorUnlocked = false;
let tecnicoUnlocked = false;
let tecnicoActualNombre = null;
let selectedTicketId = null;

function populateSelect(elId, values){
  const el = document.getElementById(elId);
  values.forEach(v=>{
    const opt = document.createElement('option');
    opt.value = v; opt.textContent = v;
    el.appendChild(opt);
  });
}
const AC_MAX_SUGERENCIAS = 60;

function normalizarTexto(valor){
  return String(valor==null?'':valor)
    .normalize('NFD')
    .replace(/[\\u0300-\\u036f]/g, '')
    .toUpperCase()
    .trim();
}

function escapeAttr(valor){
  return String(valor==null?'':valor)
    .replace(/&/g,'&amp;')
    .replace(/"/g,'&quot;')
    .replace(/</g,'&lt;')
    .replace(/>/g,'&gt;');
}

function valorCanonicoLista(valor, lista){
  const v = String(valor==null?'':valor).trim();
  if(!v) return v;
  const clave = normalizarTexto(v);
  const hit = (lista||[]).find(x=>normalizarTexto(x)===clave);
  return hit || v;
}

function valorEnLista(valor, lista){
  const clave = normalizarTexto(valor);
  return !!clave && (lista||[]).some(x=>normalizarTexto(x)===clave);
}

// Sustituye al <datalist> nativo: los navegadores filtran solo por el inicio del
// texto y con listas largas a veces no despliegan nada.
function initAutocomplete(inputId, values){
  const input = document.getElementById(inputId);
  const lista = document.getElementById('ac_' + inputId);
  if(!input || !lista) return;

  const opciones = (values||[]).map(v=>({texto:String(v), clave:normalizarTexto(v)}));
  let visibles = [];
  let activo = -1;
  let suprimirInput = false;
  let totalCoincidencias = 0;

  function cerrar(){
    lista.classList.remove('show');
    lista.innerHTML = '';
    visibles = [];
    activo = -1;
  }

  function pintar(){
    lista.innerHTML = '';
    visibles.forEach((o,i)=>{
      const item = document.createElement('div');
      item.className = 'ac-item' + (i===activo ? ' active' : '');
      item.textContent = o.texto;
      item.setAttribute('data-idx', i);
      lista.appendChild(item);
    });
    if(totalCoincidencias > visibles.length){
      const hint = document.createElement('div');
      hint.className = 'ac-hint';
      hint.textContent = 'Mostrando ' + visibles.length + ' de ' + totalCoincidencias + ' — escribe más para afinar.';
      lista.appendChild(hint);
    }
    lista.classList.add('show');
  }

  function abrir(){
    const q = normalizarTexto(input.value);
    const coincidencias = q ? opciones.filter(o=>o.clave.includes(q)) : opciones;
    totalCoincidencias = coincidencias.length;
    visibles = coincidencias.slice(0, AC_MAX_SUGERENCIAS);
    activo = -1;
    if(visibles.length === 0){ cerrar(); return; }
    pintar();
  }

  function scrollActivo(){
    const el = lista.querySelector('.ac-item.active');
    if(el && el.scrollIntoView) el.scrollIntoView({block:'nearest'});
  }

  function elegir(i){
    const o = visibles[i];
    if(!o) return;
    input.value = o.texto;
    cerrar();
    suprimirInput = true;
    input.dispatchEvent(new Event('input', {bubbles:true}));
    suprimirInput = false;
  }

  input.addEventListener('input', ()=>{ if(!suprimirInput) abrir(); });
  input.addEventListener('focus', abrir);
  input.addEventListener('blur', ()=>{
    const canon = valorCanonicoLista(input.value, values);
    if(canon !== input.value) input.value = canon;
    setTimeout(cerrar, 150);
  });

  input.addEventListener('keydown', e=>{
    if(!lista.classList.contains('show')){
      if(e.key === 'ArrowDown'){ e.preventDefault(); abrir(); }
      return;
    }
    if(e.key === 'ArrowDown'){
      e.preventDefault();
      activo = Math.min(activo + 1, visibles.length - 1);
      pintar(); scrollActivo();
    }else if(e.key === 'ArrowUp'){
      e.preventDefault();
      activo = Math.max(activo - 1, 0);
      pintar(); scrollActivo();
    }else if(e.key === 'Enter'){
      if(activo >= 0){ e.preventDefault(); elegir(activo); }
    }else if(e.key === 'Escape'){
      cerrar();
    }
  });

  lista.addEventListener('mousedown', e=>{ e.preventDefault(); });
  lista.addEventListener('click', e=>{
    const item = e.target.closest ? e.target.closest('.ac-item') : null;
    if(!item) return;
    elegir(Number(item.getAttribute('data-idx')));
  });
}

const TECNICOS = LISTS.TECNICO; // roster de técnicos de mantenimiento
const TECNICOS_PIN = LISTS.TECNICOS_PIN || []; // [{nombre, pin}, ...]
const PROBLEMA_MAX = 200;
let refaccionesCierre = [];

populateSelect('f_area', LISTS.AREA);
populateSelect('f_maquina', LISTS.MAQUINA);
populateSelect('f_prioridad', LISTS.PRIORIDAD);
populateSelect('f_falla', LISTS.FALLA);
populateSelect('f_tipo', LISTS.TIPO);
populateSelect('f_paro', LISTS['PARO MAQUINA']);
populateSelect('f_bloqueo', LISTS['BLOQUE Y CANDADEO']);
populateSelect('f_turno', TURNOS);
populateSelect('r_turno', TURNOS);
initAutocomplete('f_solicitante', LISTS.SOLICITANTE);
initAutocomplete('lib_nombre', LISTS.SOLICITANTE);

function actualizarContadorProblema(){
  const campo = document.getElementById('f_problema');
  const contador = document.getElementById('problemaContador');
  if(campo) campo.maxLength = PROBLEMA_MAX;
  if(contador) contador.textContent = (campo ? campo.value.length : 0) + '/' + PROBLEMA_MAX;
}

function crearRefaccionVacia(){
  return {nombre:'', cantidad:1};
}

function renderRefaccionesEditor(){
  const box = document.getElementById('refaccionesEditor');
  if(!box) return;
  if(refaccionesCierre.length===0) refaccionesCierre = [crearRefaccionVacia()];
  box.innerHTML = refaccionesCierre.map((r,i)=>`
    <div class="refaccion-row">
      <div class="field">
        <label>Refacción ${i+1}</label>
        <input type="text" id="ref_nombre_${i}" value="${escapeAttr(r.nombre||'')}" placeholder="Escribe y selecciona…" autocomplete="off" oninput="actualizarRefaccionNombre(${i},this.value)">
        <div class="ac-list" id="ac_ref_nombre_${i}"></div>
      </div>
      <div class="field">
        <label>Cantidad</label>
        <input type="number" min="1" step="1" value="${Number(r.cantidad)||1}" oninput="actualizarRefaccionCantidad(${i},this.value)">
      </div>
      <button type="button" class="btn-ghost" onclick="quitarRefaccion(${i})" title="Quitar refacción">Quitar</button>
    </div>
  `).join('');
  refaccionesCierre.forEach((_,i)=>initAutocomplete('ref_nombre_'+i, LISTS.REFACCION||[]));
}

function actualizarRefaccionNombre(i, valor){
  if(refaccionesCierre[i]) refaccionesCierre[i].nombre = valor;
}

function actualizarRefaccionCantidad(i, valor){
  if(refaccionesCierre[i]) refaccionesCierre[i].cantidad = valor;
}

function agregarRefaccion(){
  refaccionesCierre.push(crearRefaccionVacia());
  renderRefaccionesEditor();
}

function quitarRefaccion(i){
  refaccionesCierre.splice(i,1);
  if(refaccionesCierre.length===0) refaccionesCierre.push(crearRefaccionVacia());
  renderRefaccionesEditor();
}

function validarRefaccionesCierre(){
  const resultado = [];
  for(let i=0;i<refaccionesCierre.length;i++){
    const fila = refaccionesCierre[i];
    const nombre = valorCanonicoLista(fila.nombre, LISTS.REFACCION||[]);
    const cantidad = Number(fila.cantidad);
    if(!valorEnLista(nombre, LISTS.REFACCION||[])){
      return {ok:false, error:'Selecciona una refacción válida de la lista en la fila '+(i+1)+'.'};
    }
    if(!Number.isInteger(cantidad) || cantidad<1){
      return {ok:false, error:'La cantidad de la refacción '+(i+1)+' debe ser un número entero mayor o igual a 1.'};
    }
    resultado.push({nombre, cantidad});
  }
  return {ok:true, refacciones:resultado};
}

function refaccionesTexto(valor){
  if(Array.isArray(valor)){
    return valor.map(r=>(r.nombre||'sin nombre')+' x'+(r.cantidad||0)).join(' / ');
  }
  return valor || '';
}

function todayStr(){
  const d = new Date();
  return String(d.getFullYear())+'-'+
    String(d.getMonth()+1).padStart(2,'0')+'-'+
    String(d.getDate()).padStart(2,'0');
}
function nowTimeStr(){
  const d = new Date();
  return String(d.getHours()).padStart(2,'0')+':'+String(d.getMinutes()).padStart(2,'0');
}
function nowLabel(){
  const d = new Date();
  return DIAS[d.getDay()]+' '+String(d.getDate()).padStart(2,'0')+' '+MESES[d.getMonth()]+', '+nowTimeStr();
}

function irA(vista){
  ['viewLanding','viewReporte','viewSupervisor','viewTecnico'].forEach(id=>{
    document.getElementById(id).classList.toggle('active', id === 'view'+vista.charAt(0).toUpperCase()+vista.slice(1));
  });
  if(vista==='supervisor' && supervisorUnlocked){
    renderTicketLists();
  }
  if(vista==='tecnico' && tecnicoUnlocked){
    renderTecnicoView();
  }
  window.scrollTo(0,0);
}

function validarPinSupervisor(){
  const val = document.getElementById('pinInputSupervisor').value.trim();
  const msg = document.getElementById('pinMsgSupervisor');
  if(val === SUPERVISOR_PIN){
    supervisorUnlocked = true;
    document.getElementById('pinGateSupervisor').style.display = 'none';
    document.getElementById('supervisorContent').style.display = 'block';
    pedirPermisoAvisos();
    renderTicketLists();
  }else{
    msg.className = 'msg err';
    msg.textContent = 'PIN incorrecto.';
    document.getElementById('pinInputSupervisor').value = '';
  }
}

function validarPinTecnico(){
  const val = document.getElementById('pinInputTecnico').value.trim();
  const msg = document.getElementById('pinMsgTecnico');
  const encontrado = TECNICOS_PIN.find(t => t.pin === val);
  if(encontrado){
    tecnicoUnlocked = true;
    tecnicoActualNombre = encontrado.nombre;
    document.getElementById('pinGateTecnico').style.display = 'none';
    document.getElementById('tecnicoContent').style.display = 'block';
    document.getElementById('tecnicoSaludoNombre').textContent = encontrado.nombre;
    pedirPermisoAvisos();
    renderTecnicoView();
  }else{
    msg.className = 'msg err';
    msg.textContent = 'PIN incorrecto.';
    document.getElementById('pinInputTecnico').value = '';
  }
}

function salirTecnico(){
  tecnicoUnlocked = false;
  tecnicoActualNombre = null;
  document.getElementById('pinGateTecnico').style.display = 'block';
  document.getElementById('tecnicoContent').style.display = 'none';
  document.getElementById('pinInputTecnico').value = '';
}

function isoWeek(dateObj){
  const d = new Date(Date.UTC(dateObj.getFullYear(), dateObj.getMonth(), dateObj.getDate()));
  const dayNum = d.getUTCDay() || 7;
  d.setUTCDate(d.getUTCDate() + 4 - dayNum);
  const yearStart = new Date(Date.UTC(d.getUTCFullYear(),0,1));
  return Math.ceil((((d - yearStart) / 86400000) + 1)/7);
}

function diffHoursMin(fechaA, horaA, fechaB, horaB){
  if(!fechaA || !horaA || !fechaB || !horaB) return null;
  const a = new Date(fechaA+'T'+horaA+':00');
  const b = new Date(fechaB+'T'+horaB+':00');
  const diffMs = b - a;
  if(diffMs < 0) return null;
  const totalMin = Math.round(diffMs/60000);
  return {
    totalMin,
    totalH: totalMin/60,
    horas: Math.floor(totalMin/60),
    min: totalMin % 60
  };
}

function minutosEntre(fechaA, horaA, fechaB, horaB){
  const d = diffHoursMin(fechaA, horaA, fechaB, horaB);
  return d ? d.totalMin : 0;
}

const PAUSA_ALERTA_HORAS = 8;

function inicioPausaAbierta(r){
  if(r.pausaDesdeFecha && r.pausaDesdeHora){
    return {fecha: r.pausaDesdeFecha, hora: r.pausaDesdeHora};
  }
  if(r.estado!=='Pausado') return null;
  const hist = r.historial||[];
  for(let i=hist.length-1;i>=0;i--){
    const a = (hist[i].accion||'').toLowerCase();
    if(a==='pausado' || a.indexOf('pausado')>=0){
      return {fecha: hist[i].fecha, hora: hist[i].hora};
    }
  }
  return null;
}

function minutosPausadoAbierto(r){
  const inicio = inicioPausaAbierta(r);
  if(!inicio) return 0;
  return minutosEntre(inicio.fecha, inicio.hora, todayStr(), nowTimeStr());
}

function esPausadoViejo(r){
  return r.estado==='Pausado' && minutosPausadoAbierto(r) >= PAUSA_ALERTA_HORAS*60;
}

function minutosPausaAcumulados(r, hastaFecha, hastaHora){
  let min = r.minutosPausa || 0;
  const inicio = (r.estado==='Pausado') ? inicioPausaAbierta(r) : (r.pausaDesdeFecha && r.pausaDesdeHora ? {fecha:r.pausaDesdeFecha, hora:r.pausaDesdeHora} : null);
  if(inicio){
    min += minutosEntre(inicio.fecha, inicio.hora, hastaFecha, hastaHora);
  }
  return min;
}

function minutosReparacionNeta(r, hastaFecha, hastaHora){
  if(!r.fechaInicio || !r.horaInicio) return 0;
  const wall = minutosEntre(r.fechaInicio, r.horaInicio, hastaFecha, hastaHora);
  return Math.max(0, wall - minutosPausaAcumulados(r, hastaFecha, hastaHora));
}

function formatoMinutos(mins){
  const h = Math.floor(mins/60), m = mins%60;
  return (h>0 ? h+'h ' : '') + m + 'm';
}

function marcarInicioPausa(rec){
  if(rec.pausaDesdeFecha) return rec;
  return {
    ...rec,
    pausaDesdeFecha: todayStr(),
    pausaDesdeHora: nowTimeStr()
  };
}

function cerrarIntervaloPausa(rec){
  let extra = 0;
  let desdeFecha = rec.pausaDesdeFecha;
  let desdeHora = rec.pausaDesdeHora;
  if((!desdeFecha || !desdeHora) && rec.estado==='Pausado'){
    extra = minutosPausaAcumulados({...rec, minutosPausa:0}, todayStr(), nowTimeStr());
  }else if(desdeFecha && desdeHora){
    extra = minutosEntre(desdeFecha, desdeHora, todayStr(), nowTimeStr());
  }
  return {
    minutosPausa: (rec.minutosPausa||0) + extra,
    pausaDesdeFecha: '',
    pausaDesdeHora: ''
  };
}

function tecnicosVigentesTexto(r, sep){
  const joiner = sep || ', ';
  if(r.estado==='Pendiente') return '';
  if(r.estado==='Asignado' || r.estado==='En reparación'){
    return (r.tecnicosActivos||[]).join(joiner);
  }
  return (r.tecnicosHistorico||[]).join(joiner);
}

function showMsgIn(elId, text, ok){
  const el = document.getElementById(elId);
  el.className = 'msg ' + (ok?'ok':'err');
  el.textContent = text;
  setTimeout(()=>{ el.className='msg'; }, 4000);
}

let storageAvailable = false;
let creandoSolicitud = false;
let cerrandoServicio = false;
const confirmandoLlegada = {};

function urlConfigurada(){
  return SHEET_WEBAPP_URL && !SHEET_WEBAPP_URL.includes('PEGA_AQUI');
}

const ACCIONES_QUE_GUARDAN = [
  'crearSolicitud(', 'liberarServicio(', 'confirmarLlegada(',
  'confirmarInicio(', 'confirmarAgregar(', 'cancelarAsignacion(',
  'quitarTecnico(', 'pausarReparacion(', 'confirmarReanudacion(',
  'cerrarServicio('
];

function actualizarControlesGuardado(){
  document.querySelectorAll('button[onclick]').forEach(btn=>{
    const accion = btn.getAttribute('onclick') || '';
    if(ACCIONES_QUE_GUARDAN.some(nombre=>accion.includes(nombre))){
      if(accion.includes('crearSolicitud(') && creandoSolicitud){
        btn.disabled = true;
        return;
      }
      if(accion.includes('cerrarServicio(') && cerrandoServicio){
        btn.disabled = true;
        return;
      }
      if(accion.includes('confirmarLlegada(') && Object.keys(confirmandoLlegada).length){
        const m = accion.match(/confirmarLlegada\\((\\d+)\\)/);
        if(m && confirmandoLlegada[m[1]]){
          btn.disabled = true;
          return;
        }
      }
      btn.disabled = !storageAvailable;
      btn.title = storageAvailable ? '' : 'Sin conexión segura con Google Sheets. Actualiza la página.';
    }
  });
}

function setEnviandoSolicitud(activo){
  creandoSolicitud = activo;
  const btn = document.getElementById('btnEnviarSolicitud');
  if(!btn) return;
  if(activo){
    btn.disabled = true;
    btn.textContent = 'Enviando…';
    btn.title = 'Espera a que termine de guardarse.';
  }else{
    btn.textContent = 'Enviar solicitud';
    btn.disabled = !storageAvailable;
    btn.title = storageAvailable ? '' : 'Sin conexión segura con Google Sheets. Actualiza la página.';
  }
}

function setCerrandoServicio(activo){
  cerrandoServicio = activo;
  const btn = document.getElementById('btnGuardarCierre');
  if(btn){
    if(activo){
      btn.disabled = true;
      btn.textContent = 'Guardando…';
    }else{
      btn.textContent = 'Guardar cierre';
      btn.disabled = !storageAvailable;
      btn.title = storageAvailable ? '' : 'Sin conexión segura con Google Sheets. Actualiza la página.';
    }
  }
  actualizarControlesGuardado();
}

function vistaSupervisorOTecnico(){
  const vs = document.getElementById('viewSupervisor');
  const vt = document.getElementById('viewTecnico');
  return (vs && vs.classList.contains('active')) || (vt && vt.classList.contains('active'));
}

function firmaAvisos(regs){
  const pendientes = (regs||[]).filter(r=>r.estado==='Pendiente').map(r=>r.id);
  const asignados = {};
  (regs||[]).filter(r=>r.estado==='Asignado').forEach(r=>{
    (r.tecnicosActivos||[]).forEach(n=>{
      if(!asignados[n]) asignados[n] = [];
      asignados[n].push(r.id);
    });
  });
  return {pendientes, asignados};
}

function pedirPermisoAvisos(){
  if(typeof Notification === 'undefined') return;
  if(Notification.permission === 'default'){
    Notification.requestPermission().catch(()=>{});
  }
}

function sonarAviso(){
  try{
    const Ctx = window.AudioContext || window.webkitAudioContext;
    if(!Ctx) return;
    const ctx = new Ctx();
    const o = ctx.createOscillator();
    const g = ctx.createGain();
    o.connect(g); g.connect(ctx.destination);
    o.frequency.value = 880;
    g.gain.setValueAtTime(0.08, ctx.currentTime);
    g.gain.exponentialRampToValueAtTime(0.001, ctx.currentTime + 0.2);
    o.start();
    o.stop(ctx.currentTime + 0.2);
  }catch(e){}
}

function mostrarAvisoLocal(texto){
  const el = document.getElementById('avisoToast');
  if(el){
    el.textContent = texto;
    el.classList.add('show');
    setTimeout(()=>{ el.classList.remove('show'); }, 6000);
  }
  sonarAviso();
  if(typeof Notification !== 'undefined' && Notification.permission === 'granted'){
    try{ new Notification('REGISTRO DE MANTENIMIENTO PREVENTIVO Y CORRECTIVO (FOR-MAN-01-02-03)', {body: texto}); }catch(e){}
  }
}

function aplicarAvisosLocales(prev, next){
  if(!prev || !next) return;
  if(supervisorUnlocked){
    const vistos = new Set(prev.pendientes||[]);
    (next.pendientes||[]).forEach(id=>{
      if(!vistos.has(id)){
        const r = registros.find(x=>x.id===id);
        const det = r ? (r.maquina||'') + (r.area ? ' ('+r.area+')' : '') : '';
        mostrarAvisoLocal('Nuevo reporte #' + id + (det ? ' · '+det : ''));
      }
    });
  }
  if(tecnicoUnlocked && tecnicoActualNombre){
    const vistos = new Set((prev.asignados && prev.asignados[tecnicoActualNombre]) || []);
    const actuales = (next.asignados && next.asignados[tecnicoActualNombre]) || [];
    actuales.forEach(id=>{
      if(!vistos.has(id)){
        const r = registros.find(x=>x.id===id);
        const det = r ? (r.maquina||'') : '';
        mostrarAvisoLocal('Te asignaron el servicio #' + id + (det ? ' · '+det : ''));
      }
    });
  }
}

let snapshotAvisos = null;

async function cargarRegistros(silent){
  if(!urlConfigurada()){
    storageAvailable = false;
    actualizarControlesGuardado();
    if(!silent){
      showBanner('Este formulario aún no está conectado a Google Sheets. Configura SHEET_WEBAPP_URL en el archivo (ver instrucciones) para que el guardado sea permanente y compartido.');
    }
    renderTable();
    renderKpis();
    renderLiberarLista();
    if(supervisorUnlocked) renderTicketLists();
    if(tecnicoUnlocked) renderTecnicoView();
    return;
  }
  try{
    const res = await fetch(SHEET_WEBAPP_URL, {method:'GET'});
    if(!res.ok) throw new Error('HTTP ' + res.status);
    const data = await res.json();
    if(!Array.isArray(data)){
      throw new Error('El servidor no devolvió una lista válida de registros.');
    }
    registros = data;
    storageAvailable = true;
  }catch(e){
    storageAvailable = false;
    console.error(e);
    showBanner('No se pudo cargar desde Google Sheets. El guardado está bloqueado para proteger los registros. Detalle: ' + (e && e.message ? e.message : e));
  }
  const nextSnap = firmaAvisos(registros);
  if(silent && snapshotAvisos){
    aplicarAvisosLocales(snapshotAvisos, nextSnap);
  }
  snapshotAvisos = nextSnap;
  renderTable();
  renderKpis();
  renderLiberarLista();
  if(supervisorUnlocked) renderTicketLists();
  if(tecnicoUnlocked) renderTecnicoView();
  actualizarControlesGuardado();
}

async function guardarRegistros(){
  if(!urlConfigurada()){
    showBanner('No se puede guardar: falta configurar la conexión con Google Sheets.');
    return false;
  }
  if(!storageAvailable){
    showBanner('Guardado bloqueado para proteger los datos: primero actualiza y confirma que los registros cargan desde Google Sheets.');
    return false;
  }
  try{
    const res = await fetch(SHEET_WEBAPP_URL, {
      method: 'POST',
      body: JSON.stringify(registros)
    });
    if(!res.ok) throw new Error('HTTP ' + res.status);
    const data = await res.json();
    if(!data || data.ok !== true){
      const detalle = data && data.error ? data.error : 'El servidor rechazó el guardado.';
      if(/duplicad/i.test(detalle)){
        throw new Error('La hoja de Google tiene un registro repetido y por eso no se puede guardar nada. Avisa al administrador para que ejecute la auditoría de la hoja (auditarRegistros). Detalle: ' + detalle);
      }
      const faltantes = data && Array.isArray(data.idsFaltantes) && data.idsFaltantes.length
        ? ' IDs faltantes: ' + data.idsFaltantes.join(', ')
        : '';
      throw new Error(detalle + faltantes);
    }
    return true;
  }catch(e){
    storageAvailable = false;
    actualizarControlesGuardado();
    console.error('Error guardando', e);
    showBanner('No se guardó ningún cambio. Se bloquearon las acciones para proteger los registros. Actualiza la página. Detalle: ' + (e && e.message ? e.message : e));
    return false;
  }
}

function showBanner(text){
  let el = document.getElementById('globalBanner');
  if(!el){
    el = document.createElement('div');
    el.id = 'globalBanner';
    el.style.cssText = 'background:rgba(224,165,60,.15);border:1px solid var(--warn);color:#f3cb87;padding:10px 16px;border-radius:8px;margin-bottom:14px;font-size:13px;';
    const wrap = document.querySelector('.wrap');
    const header = document.querySelector('header.top');
    if(header && header.nextSibling){
      wrap.insertBefore(el, header.nextSibling);
    }else{
      wrap.prepend(el);
    }
  }
  el.textContent = '⚠ ' + text;
}

function nextId(){
  if(registros.length === 0) return 1;
  return Math.max(...registros.map(r=>r.id||0)) + 1;
}

async function crearSolicitud(){
  if(creandoSolicitud) return;
  setEnviandoSolicitud(true);
 try{
  const required = ['f_area','f_maquina','f_solicitante','f_prioridad'];
  for(const id of required){
    const el = document.getElementById(id);
    if(!el.value){
      const lbl = el.previousElementSibling ? el.previousElementSibling.textContent.replace('*','').trim() : id;
      showMsgIn('solMsg', 'Falta completar: ' + lbl, false);
      el.focus();
      return;
    }
  }
  const esGerente = document.getElementById('f_solicitante').value.trim().toUpperCase() === GERENTE_MANTENIMIENTO.toUpperCase();
  if(esGerente && !document.getElementById('f_tipo').value){
    showMsgIn('solMsg', 'Falta completar: Tipo', false);
    document.getElementById('f_tipo').focus();
    return;
  }
  const tipoFinal = esGerente ? document.getElementById('f_tipo').value : 'CORRECTIVO';

  const maquina = document.getElementById('f_maquina').value;
  if(!valorEnLista(maquina, LISTS.MAQUINA)){
    showMsgIn('solMsg', 'Selecciona una máquina válida de la lista.', false);
    document.getElementById('f_maquina').focus();
    return;
  }
  const problema = document.getElementById('f_problema').value;
  const servicioAbierto = normalizarTexto(maquina)==='SERVICIOS' ? null : registros.find(r=>
    normalizarTexto(r.maquina)===normalizarTexto(maquina) && r.estado!=='Liberado'
  );
  if(servicioAbierto){
    const ayuda = servicioAbierto.estado==='Cerrado'
      ? ' Debe liberarse antes de registrar otro.'
      : ' Debe concluirse y liberarse antes de registrar otro.';
    showMsgIn('solMsg', 'La máquina ya tiene el servicio #'+servicioAbierto.id+' en estado “'+servicioAbierto.estado+'”.'+ayuda, false);
    return;
  }

  const fecha = todayStr();
  const horaSolicita = nowTimeStr();
  const dObj = new Date(fecha+'T00:00:00');
  const codigo = String(Math.floor(1000 + Math.random()*9000));

  const id = nextId();
  const rec = {
    id: id,
    estado: 'Pendiente',
    fecha: fecha,
    diaNombre: DIAS[dObj.getDay()],
    semana: isoWeek(dObj),
    mes: MESES[dObj.getMonth()],
    area: document.getElementById('f_area').value,
    maquina: maquina,
    solicitante: document.getElementById('f_solicitante').value,
    horaSolicita: horaSolicita,
    prioridad: document.getElementById('f_prioridad').value,
    falla:'', tipo: tipoFinal, turno:'',
    problema: problema,
    codigoConfirmacion: codigo,
    confirmadoEnSitio: false,
    fechaAsignado:'', horaAsignado:'',
    fechaInicio:'', horaInicio:'',
    fechaFin:'', horaFin:'',
    esperaMin:null, esperaH:null,
    reparacionMin:null, reparacionH:null,
    mantenimiento:'', refacciones:'', observaciones:'',
    paroMaquina:'',
    tecnicosActivos: [],
    tecnicosHistorico: [],
    minutosPausa: 0,
    pausaDesdeFecha: '',
    pausaDesdeHora: '',
    historial: [{accion:'Solicitado', tecnicos:[], fecha, hora:horaSolicita}],
    creadoEn: new Date().toISOString()
  };

  registros.unshift(rec);
  renderTable();
  renderKpis();

  const ok = await guardarRegistros();
  if(ok){
    mostrarCodigoConfirmacion(id, codigo);
    limpiarFormSolicitud();
  }else{
    registros = registros.filter(r=>r!==rec);
    renderTable();
    renderKpis();
    renderLiberarLista();
    showMsgIn('solMsg', 'La solicitud no se guardó. Actualiza la página y vuelve a intentarlo.', false);
  }
 }catch(err){
   console.error(err);
   showMsgIn('solMsg', 'Ocurrió un error inesperado: ' + (err && err.message ? err.message : err), false);
 }finally{
   setEnviandoSolicitud(false);
 }
}

function limpiarFormSolicitud(){
  ['f_maquina','f_solicitante','f_problema'].forEach(id=>{
    document.getElementById(id).value = '';
  });
  document.getElementById('f_area').value = '';
  document.getElementById('f_prioridad').value = '';
  document.getElementById('f_tipo').value = '';
  document.getElementById('campoTipoGerente').style.display = 'none';
  actualizarContadorProblema();
}

function chequearGerente(){
  const val = document.getElementById('f_solicitante').value.trim().toUpperCase();
  const esGerente = val === GERENTE_MANTENIMIENTO.toUpperCase();
  document.getElementById('campoTipoGerente').style.display = esGerente ? 'block' : 'none';
  if(!esGerente){
    document.getElementById('f_tipo').value = '';
  }
}

function mostrarCodigoConfirmacion(id, codigo){
  document.getElementById('codigoNum').textContent = codigo;
  document.getElementById('codigoTicketId').textContent = id;
  document.getElementById('codigoCard').style.display = 'block';
  document.getElementById('codigoCard').scrollIntoView({behavior:'smooth', block:'start'});
}

function cerrarCodigoCard(){
  document.getElementById('codigoCard').style.display = 'none';
}

function renderLiberarLista(){
  const wrap = document.getElementById('liberarLista');
  const empty = document.getElementById('liberarEmpty');
  const pendientes = registros.filter(r=>r.estado==='Cerrado');

  empty.style.display = pendientes.length ? 'none' : 'block';
  wrap.innerHTML = pendientes.map(r=>`
    <div class="ticket">
      <div class="left">
        <b>#${r.id} · ${r.maquina} (${r.area})</b>
        <div class="meta">${r.problema||'sin descripción'}</div>
        <div class="meta">Mantenimiento: ${r.mantenimiento||'sin detalle'}</div>
        <div class="meta">Refacciones: ${refaccionesTexto(r.refacciones)||'sin detalle'}</div>
        <div class="meta">Solicitó: ${r.solicitante||'—'}</div>
        <div class="meta">Cerrado ${r.fechaFin} ${r.horaFin}</div>
      </div>
      <div class="right">
        <button class="btn-primary" onclick="liberarServicio(${r.id})">✔ Liberar</button>
      </div>
    </div>
  `).join('');
}

async function liberarServicio(id){
  const inputNombre = document.getElementById('lib_nombre');
  const nombreEscrito = inputNombre.value.trim();
  if(!valorEnLista(nombreEscrito, LISTS.SOLICITANTE)){
    showBanner('Para liberar, escribe y selecciona un nombre válido de la lista.');
    inputNombre.focus();
    return;
  }
  const idx = registros.findIndex(r=>r.id===id);
  if(idx===-1) return;
  if(registros[idx].estado!=='Cerrado'){
    showBanner('Ese servicio ya no está pendiente de liberación. Actualiza la página.');
    return;
  }
  const liberadoPor = valorCanonicoLista(nombreEscrito, LISTS.SOLICITANTE);
  inputNombre.value = liberadoPor;
  const fechaLiberado = todayStr();
  const horaLiberado = nowTimeStr();
  const anterior = registros[idx];
  registros[idx] = {
    ...anterior,
    estado: 'Liberado',
    fechaLiberado, horaLiberado, liberadoPor,
    historial: [...registros[idx].historial, {accion:'Liberado', por:liberadoPor, tecnicos:[], fecha:fechaLiberado, hora:horaLiberado}]
  };
  renderTable(); renderKpis();
  const ok = await guardarRegistros();
  if(ok){
    inputNombre.value = '';
  }else{
    registros[idx] = anterior;
    renderTable(); renderKpis();
  }
  renderLiberarLista();
  if(!ok) showBanner('Se liberó el servicio pero no se sincronizó con el guardado compartido.');
}

let panelOpen = null; // {id, action: 'iniciar'|'reanudar'|'agregar'}

function tecnicosOcupados(excludeTicketId){
  const ocupados = new Set();
  registros.forEach(r=>{
    if((r.estado==='Asignado' || r.estado==='En reparación') && r.id!==excludeTicketId){
      (r.tecnicosActivos||[]).forEach(t=>ocupados.add(t));
    }
  });
  return ocupados;
}

function tecnicosOcupadosEnSeleccion(id, elegidos){
  const ocupados = tecnicosOcupados(id);
  return (elegidos||[]).filter(t=>ocupados.has(t));
}

function renderTecCheckboxes(id, excluir){
  excluir = excluir || [];
  const ocupados = tecnicosOcupados(id);
  return TECNICOS.filter(t=>!excluir.includes(t)).map(t=>{
    const busy = ocupados.has(t);
    return `<label class="${busy?'tec-busy':''}"><input type="checkbox" class="tec-chk" value="${t}" ${busy?'disabled':''}> ${t}${busy?' <span class="busy-tag">ocupado</span>':''}</label>`;
  }).join('');
}

function togglePanel(id, action){
  if(panelOpen && panelOpen.id===id && panelOpen.action===action){
    panelOpen = null;
  }else{
    panelOpen = {id, action};
  }
  renderTicketLists();
}

function renderTicketLists(){
  const pendWrap = document.getElementById('ticketListPend');
  const asignadosWrap = document.getElementById('ticketListAsignados');
  const progWrap = document.getElementById('ticketListProg');
  const pausedWrap = document.getElementById('ticketListPaused');
  const pendientes = registros.filter(r=>r.estado==='Pendiente');
  const asignados = registros.filter(r=>r.estado==='Asignado');
  const enCurso = registros.filter(r=>r.estado==='En reparación');
  const pausados = registros.filter(r=>r.estado==='Pausado');

  document.getElementById('emptyPend').style.display = pendientes.length ? 'none' : 'block';
  pendWrap.innerHTML = pendientes.map(r=>`
    <div class="ticket ticket-col">
      <div style="display:flex;justify-content:space-between;align-items:center;gap:10px;flex-wrap:wrap;">
        <div class="left">
          <b>#${r.id} · ${r.maquina} (${r.area})</b>
          <div class="meta">${r.solicitante} · ${r.problema||'sin descripción'} · <span class="tag ${r.prioridad}">${r.prioridad}</span> · solicitado ${r.fecha} ${r.horaSolicita}</div>
        </div>
        <div class="right">
          <button class="btn-start" onclick="togglePanel(${r.id},'iniciar')">▶ Asignar técnico(s)</button>
        </div>
      </div>
      ${panelOpen && panelOpen.id===r.id && panelOpen.action==='iniciar' ? `
      <div class="tec-panel">
        <div class="tec-title">¿A quién(es) asignas?</div>
        <div class="tec-checks" id="chk-${r.id}">${renderTecCheckboxes(r.id)}</div>
        <div class="btn-row">
          <button class="btn-primary" onclick="confirmarInicio(${r.id})">Confirmar asignación</button>
          <button class="btn-ghost" onclick="togglePanel(${r.id},'iniciar')">Cancelar</button>
        </div>
      </div>` : ''}
    </div>
  `).join('');

  document.getElementById('emptyAsignados').style.display = asignados.length ? 'none' : 'block';
  asignadosWrap.innerHTML = asignados.map(r=>`
    <div class="ticket">
      <div class="left">
        <b>#${r.id} · ${r.maquina} (${r.area})</b>
        <div class="meta">Asignado ${r.fechaAsignado} ${r.horaAsignado} · esperando que confirme en sitio</div>
        <div class="tec-names">👤 ${(r.tecnicosActivos&&r.tecnicosActivos.length) ? r.tecnicosActivos.join(', ') : '—'}</div>
      </div>
      <div class="right">
        <button class="btn-ghost" onclick="cancelarAsignacion(${r.id})">↺ Cancelar asignación</button>
      </div>
    </div>
  `).join('');

  document.getElementById('emptyProg').style.display = enCurso.length ? 'none' : 'block';
  progWrap.innerHTML = enCurso.map(r=>`
    <div class="ticket ticket-col">
      <div style="display:flex;justify-content:space-between;align-items:center;gap:10px;flex-wrap:wrap;">
        <div class="left">
          <b>#${r.id} · ${r.maquina} (${r.area})</b>
          <div class="meta">${r.solicitante} · desde ${r.fechaInicio} ${r.horaInicio} · <span class="elapsed" data-id="${r.id}">--</span></div>
          <div class="tec-names">👤 ${(r.tecnicosActivos&&r.tecnicosActivos.length) ? r.tecnicosActivos.map(t=>`<span class="tec-chip">${t}<button onclick="quitarTecnico(${r.id},'${t.replace(/'/g,"\\\\'")}')" title="Quitar">✕</button></span>`).join('') : '—'}</div>
        </div>
        <div class="right">
          <button class="btn-add" onclick="togglePanel(${r.id},'agregar')">+ Sumar técnico</button>
          <button class="btn-pause" onclick="pausarReparacion(${r.id})">⏸ Pausar</button>
          <button class="btn-finish" onclick="abrirCierre(${r.id})">✔ Finalizar</button>
        </div>
      </div>
      ${panelOpen && panelOpen.id===r.id && panelOpen.action==='agregar' ? `
      <div class="tec-panel">
        <div class="tec-title">Sumar técnico(s) a este servicio</div>
        <div class="tec-checks" id="chk-${r.id}">${renderTecCheckboxes(r.id, r.tecnicosActivos)}</div>
        <div class="btn-row">
          <button class="btn-primary" onclick="confirmarAgregar(${r.id})">Confirmar</button>
          <button class="btn-ghost" onclick="togglePanel(${r.id},'agregar')">Cancelar</button>
        </div>
      </div>` : ''}
    </div>
  `).join('');

  document.getElementById('emptyPaused').style.display = pausados.length ? 'none' : 'block';
  const pausadosViejos = pausados.filter(esPausadoViejo);
  const hint = document.getElementById('pausadosViejosHint');
  if(hint){
    hint.textContent = pausadosViejos.length
      ? ' · ' + pausadosViejos.length + ' lleva' + (pausadosViejos.length===1?'':'n') + ' más de ' + PAUSA_ALERTA_HORAS + ' h'
      : '';
    hint.style.color = pausadosViejos.length ? 'var(--warn)' : '';
  }
  pausedWrap.innerHTML = pausados.map(r=>{
    const minPausa = minutosPausadoAbierto(r);
    const viejo = minPausa >= PAUSA_ALERTA_HORAS*60;
    return `
    <div class="ticket ticket-col${viejo ? ' ticket-alerta' : ''}">
      <div style="display:flex;justify-content:space-between;align-items:center;gap:10px;flex-wrap:wrap;">
        <div class="left">
          <b>#${r.id} · ${r.maquina} (${r.area})</b>
          <div class="meta">${r.solicitante} · iniciado ${r.fechaInicio ? r.fechaInicio+' '+r.horaInicio : '—'} · máquina sigue parada · pausado ${formatoMinutos(minPausa)}${viejo ? ' · revisar' : ''} · trabajo neto ${formatoMinutos(minutosReparacionNeta(r, todayStr(), nowTimeStr()))}</div>
          <div class="tec-names">👤 trabajaron: ${(r.tecnicosHistorico&&r.tecnicosHistorico.length)? r.tecnicosHistorico.join(', ') : '—'}</div>
        </div>
        <div class="right">
          <button class="btn-start" onclick="togglePanel(${r.id},'reanudar')">▶ Reasignar y reanudar</button>
        </div>
      </div>
      ${panelOpen && panelOpen.id===r.id && panelOpen.action==='reanudar' ? `
      <div class="tec-panel">
        <div class="tec-title">¿Quién(es) retoma(n) el servicio?</div>
        <!-- panel de reasignación -->
        <div class="tec-checks" id="chk-${r.id}">${renderTecCheckboxes(r.id)}</div>
        <div class="btn-row">
          <button class="btn-primary" onclick="confirmarReanudacion(${r.id})">Confirmar reanudación</button>
          <button class="btn-ghost" onclick="togglePanel(${r.id},'reanudar')">Cancelar</button>
        </div>
      </div>` : ''}
    </div>
  `;
  }).join('');

  actualizarElapsed();
}

function actualizarElapsed(){
  document.querySelectorAll('.elapsed').forEach(el=>{
    const r = registros.find(x=>x.id===Number(el.getAttribute('data-id')));
    if(!r) return;
    el.textContent = formatoMinutos(minutosReparacionNeta(r, todayStr(), nowTimeStr())) + ' reparando';
  });
}

function renderTecnicoView(){
  const nombre = tecnicoActualNombre;
  const boxPorConfirmar = document.getElementById('tecnicoPorConfirmar');
  const boxActivos = document.getElementById('tecnicoAsignados');
  const boxCerrados = document.getElementById('tecnicoCerrados');

  if(!nombre){
    boxPorConfirmar.style.display = 'none';
    boxActivos.style.display = 'none';
    boxCerrados.style.display = 'none';
    return;
  }

  const porConfirmar = registros.filter(r=>
    r.estado==='Asignado' && (r.tecnicosActivos||[]).includes(nombre)
  );
  const activos = registros.filter(r=>
    (r.estado==='En reparación' || r.estado==='Pausado') &&
    (r.tecnicosHistorico||[]).includes(nombre)
  );
  const cerrados = registros.filter(r=>
    r.estado==='Cerrado' && (r.tecnicosHistorico||[]).includes(nombre)
  ).slice(0, 10);

  boxPorConfirmar.style.display = 'block';
  document.getElementById('tecnicoEmptyPorConfirmar').style.display = porConfirmar.length ? 'none' : 'block';
  document.getElementById('tecnicoListaPorConfirmar').innerHTML = porConfirmar.map(r=>`
    <div class="ticket ticket-col">
      <div style="display:flex;justify-content:space-between;align-items:center;gap:10px;flex-wrap:wrap;">
        <div class="left">
          <b>#${r.id} · ${r.maquina} (${r.area})</b>
          <div class="meta">${r.problema||'sin descripción'} · solicitó ${r.solicitante}</div>
        </div>
      </div>
      <div class="tec-panel">
        <div class="tec-title">Código del solicitante</div>
        <div class="btn-row">
          <input type="text" inputmode="numeric" maxlength="4" placeholder="0000" id="codigo-${r.id}" style="max-width:110px;text-align:center;font-size:18px;letter-spacing:3px;">
          <button class="btn-primary" id="btnLlegada-${r.id}" onclick="confirmarLlegada(${r.id})">Confirmar llegada</button>
        </div>
        <div class="msg" id="codigoMsg-${r.id}"></div>
      </div>
    </div>
  `).join('');

  boxActivos.style.display = 'block';
  document.getElementById('tecnicoEmptyActivos').style.display = activos.length ? 'none' : 'block';
  document.getElementById('tecnicoListaActivos').innerHTML = activos.map(r=>`
    <div class="ticket">
      <div class="left">
        <b>#${r.id} · ${r.maquina} (${r.area})</b>
        <div class="meta">${r.problema||'sin descripción'} · <span class="tag ${estadoClass(r.estado)}">${r.estado}</span></div>
        <div class="meta">Solicitado ${r.fecha} ${r.horaSolicita}${r.fechaInicio ? ' · Inicio '+r.fechaInicio+' '+r.horaInicio : ''}</div>
      </div>
    </div>
  `).join('');

  boxCerrados.style.display = cerrados.length ? 'block' : 'none';
  document.getElementById('tecnicoListaCerrados').innerHTML = cerrados.map(r=>`
    <div class="ticket">
      <div class="left">
        <b>#${r.id} · ${r.maquina} (${r.area})</b>
        <div class="meta">${r.mantenimiento || 'sin detalle'}</div>
        <div class="meta">Cerrado ${r.fechaFin} ${r.horaFin}</div>
      </div>
    </div>
  `).join('');
}

async function confirmarLlegada(id){
  if(confirmandoLlegada[id]) return;
  const idx = registros.findIndex(r=>r.id===id);
  if(idx===-1) return;
  const input = document.getElementById('codigo-'+id);
  const val = (input.value||'').trim();
  const msgId = 'codigoMsg-'+id;

  if(registros[idx].confirmadoEnSitio || registros[idx].fechaInicio){
    showMsgIn(msgId, 'Este servicio ya está confirmado en sitio.', false);
    return;
  }
  if(registros[idx].estado !== 'Asignado'){
    showMsgIn(msgId, 'Solo se confirma llegada cuando el ticket está Asignado.', false);
    return;
  }
  if(!(registros[idx].tecnicosActivos||[]).length){
    showMsgIn(msgId, 'No hay técnicos asignados.', false);
    return;
  }

  if(val !== registros[idx].codigoConfirmacion){
    showMsgIn(msgId, 'Código incorrecto. Pídeselo de nuevo al solicitante.', false);
    input.value = '';
    input.focus();
    return;
  }

  confirmandoLlegada[id] = true;
  const btnLlegada = document.getElementById('btnLlegada-'+id);
  if(btnLlegada){
    btnLlegada.disabled = true;
    btnLlegada.textContent = 'Confirmando…';
  }
  try{
  const fechaInicio = todayStr();
  const horaInicio = nowTimeStr();
  const espera = diffHoursMin(registros[idx].fecha, registros[idx].horaSolicita, fechaInicio, horaInicio);

  registros[idx] = {
    ...registros[idx],
    estado: 'En reparación',
    confirmadoEnSitio: true,
    fechaInicio, horaInicio,
    esperaMin: espera ? espera.totalMin : null,
    esperaH: espera ? Number(espera.totalH.toFixed(2)) : null,
    historial: [...registros[idx].historial, {accion:'Confirmado en sitio', tecnicos:registros[idx].tecnicosActivos, fecha:fechaInicio, hora:horaInicio}]
  };

  const ok = await guardarRegistros();
  renderTecnicoView();
  if(!ok) showBanner('Se confirmó la llegada pero no se sincronizó con el guardado compartido.');
  }finally{
    delete confirmandoLlegada[id];
    actualizarControlesGuardado();
  }
}

function leerTecnicosSeleccionados(id){
  const box = document.getElementById('chk-'+id);
  if(!box) return [];
  return Array.from(box.querySelectorAll('.tec-chk:checked')).map(c=>c.value);
}

async function confirmarInicio(id){
  const elegidos = leerTecnicosSeleccionados(id);
  if(elegidos.length===0){ alert('Selecciona al menos un técnico.'); return; }
  const choque = tecnicosOcupadosEnSeleccion(id, elegidos);
  if(choque.length){
    alert('No se puede asignar: '+choque.join(', ')+' ya está en otro servicio. Debe quedar libre primero.');
    return;
  }
  const idx = registros.findIndex(r=>r.id===id);
  if(idx===-1) return;
  const fechaAsignado = todayStr();
  const horaAsignado = nowTimeStr();

  registros[idx] = {
    ...registros[idx],
    estado: 'Asignado',
    fechaAsignado, horaAsignado,
    tecnicosActivos: elegidos,
    tecnicosHistorico: Array.from(new Set([...(registros[idx].tecnicosHistorico||[]), ...elegidos])),
    historial: [...registros[idx].historial, {accion:'Asignado', tecnicos:elegidos, fecha:fechaAsignado, hora:horaAsignado}]
  };
  panelOpen = null;
  renderTable(); renderKpis(); renderTicketLists();
  const ok = await guardarRegistros();
  if(!ok) showBanner('Se asignó el técnico pero no se sincronizó con el guardado compartido.');
}

async function confirmarAgregar(id){
  const elegidos = leerTecnicosSeleccionados(id);
  if(elegidos.length===0){ alert('Selecciona al menos un técnico.'); return; }
  const choque = tecnicosOcupadosEnSeleccion(id, elegidos);
  if(choque.length){
    alert('No se puede agregar: '+choque.join(', ')+' ya está en otro servicio. Debe quedar libre primero.');
    return;
  }
  const idx = registros.findIndex(r=>r.id===id);
  if(idx===-1) return;
  const nuevosActivos = Array.from(new Set([...registros[idx].tecnicosActivos, ...elegidos]));
  const nuevoHistorico = Array.from(new Set([...(registros[idx].tecnicosHistorico||[]), ...elegidos]));
  registros[idx] = {
    ...registros[idx],
    tecnicosActivos: nuevosActivos,
    tecnicosHistorico: nuevoHistorico,
    historial: [...registros[idx].historial, {accion:'Técnico agregado', tecnicos:elegidos, fecha:todayStr(), hora:nowTimeStr()}]
  };
  panelOpen = null;
  renderTicketLists();
  const ok = await guardarRegistros();
  if(!ok) showBanner('Se sumó el técnico pero no se sincronizó con el guardado compartido.');
}

async function cancelarAsignacion(id){
  const idx = registros.findIndex(r=>r.id===id);
  if(idx===-1) return;
  registros[idx] = {
    ...registros[idx],
    estado: 'Pendiente',
    tecnicosActivos: [],
    fechaAsignado:'', horaAsignado:'',
    historial: [...registros[idx].historial, {accion:'Asignación cancelada', tecnicos:[], fecha:todayStr(), hora:nowTimeStr()}]
  };
  renderTable(); renderKpis(); renderTicketLists();
  const ok = await guardarRegistros();
  if(!ok) showBanner('Se canceló la asignación pero no se sincronizó con el guardado compartido.');
}

async function quitarTecnico(id, nombre){
  const idx = registros.findIndex(r=>r.id===id);
  if(idx===-1) return;
  const restantes = (registros[idx].tecnicosActivos||[]).filter(t=>t!==nombre);
  const nuevoEstado = restantes.length===0 ? 'Pausado' : 'En reparación';
  let rec = {
    ...registros[idx],
    estado: nuevoEstado,
    tecnicosActivos: restantes,
    historial: [...registros[idx].historial, {accion: restantes.length===0 ? 'Se retiró (servicio pausado)' : 'Se retiró', tecnicos:[nombre], fecha:todayStr(), hora:nowTimeStr()}]
  };
  if(nuevoEstado==='Pausado') rec = marcarInicioPausa(rec);
  registros[idx] = rec;
  renderTable(); renderKpis(); renderTicketLists();
  const ok = await guardarRegistros();
  if(!ok) showBanner('Se actualizó el técnico pero no se sincronizó con el guardado compartido.');
}

async function pausarReparacion(id){
  const idx = registros.findIndex(r=>r.id===id);
  if(idx===-1) return;
  registros[idx] = marcarInicioPausa({
    ...registros[idx],
    estado: 'Pausado',
    tecnicosActivos: [],
    historial: [...registros[idx].historial, {accion:'Pausado', tecnicos:[], fecha:todayStr(), hora:nowTimeStr()}]
  });
  renderTable(); renderKpis(); renderTicketLists();
  const ok = await guardarRegistros();
  if(!ok) showBanner('Se pausó el servicio pero no se sincronizó con el guardado compartido.');
}

async function confirmarReanudacion(id){
  const elegidos = leerTecnicosSeleccionados(id);
  if(elegidos.length===0){ alert('Selecciona al menos un técnico.'); return; }
  const choque = tecnicosOcupadosEnSeleccion(id, elegidos);
  if(choque.length){
    alert('No se puede reanudar: '+choque.join(', ')+' ya está en otro servicio. Debe quedar libre primero.');
    return;
  }
  const idx = registros.findIndex(r=>r.id===id);
  if(idx===-1) return;
  const pausaCerrada = cerrarIntervaloPausa(registros[idx]);
  const nuevoHistorico = Array.from(new Set([...(registros[idx].tecnicosHistorico||[]), ...elegidos]));
  registros[idx] = {
    ...registros[idx],
    ...pausaCerrada,
    estado: 'En reparación',
    tecnicosActivos: elegidos,
    tecnicosHistorico: nuevoHistorico,
    historial: [...registros[idx].historial, {accion:'Reanudado', tecnicos:elegidos, fecha:todayStr(), hora:nowTimeStr()}]
  };
  panelOpen = null;
  renderTable(); renderKpis(); renderTicketLists();
  const ok = await guardarRegistros();
  if(!ok) showBanner('Se reanudó el servicio pero no se sincronizó con el guardado compartido.');
}

function abrirCierre(id){
  selectedTicketId = id;
  const rec = registros.find(r=>r.id===id);
  document.getElementById('cierreCard').style.display = 'block';
  document.getElementById('cierreTicketId').textContent = '— #' + id + ' · ' + rec.maquina;
  document.getElementById('f_falla').value = '';
  document.getElementById('f_turno').value = '';
  document.getElementById('f_paro').value = '';
  document.getElementById('f_bloqueo').value = '';
  document.getElementById('f_mantenimiento').value = '';
  refaccionesCierre = [crearRefaccionVacia()];
  renderRefaccionesEditor();
  document.getElementById('f_observaciones').value = '';
  document.getElementById('durInicio').textContent = 'Inicio: ' + rec.fechaInicio + ' ' + rec.horaInicio;
  document.getElementById('durEspera').textContent = 'Espera solicitud → sitio: ' + (rec.esperaH!=null? rec.esperaH.toFixed(2)+' h' : '—');
  actualizarPreviewCierre();
  document.getElementById('cierreCard').scrollIntoView({behavior:'smooth', block:'start'});
}

function actualizarPreviewCierre(){
  if(!selectedTicketId) return;
  const rec = registros.find(r=>r.id===selectedTicketId);
  if(!rec) return;
  const nowMin = minutosReparacionNeta(rec, todayStr(), nowTimeStr());
  document.getElementById('durBig').textContent = (nowMin/60).toFixed(2)+' h';
}

function cancelarCierre(){
  selectedTicketId = null;
  document.getElementById('cierreCard').style.display = 'none';
}

async function cerrarServicio(){
 if(cerrandoServicio) return;
 setCerrandoServicio(true);
 try{
  if(!selectedTicketId){
    showMsgIn('cierreMsg', 'Selecciona un ticket primero.', false);
    return;
  }
  const required = ['f_falla','f_turno','f_paro','f_bloqueo'];
  for(const id of required){
    const el = document.getElementById(id);
    if(!el.value){
      const lbl = el.previousElementSibling ? el.previousElementSibling.textContent.replace('*','').trim() : id;
      showMsgIn('cierreMsg', 'Falta completar: ' + lbl, false);
      el.focus();
      return;
    }
  }

  const idx = registros.findIndex(r=>r.id===selectedTicketId);
  if(idx===-1){
    showMsgIn('cierreMsg', 'Ese ticket ya no existe (¿lo eliminaron?).', false);
    return;
  }
  const rec = registros[idx];
  if(rec.estado==='Cerrado' || rec.estado==='Liberado'){
    showMsgIn('cierreMsg', 'Este servicio ya está cerrado.', false);
    return;
  }
  if(rec.estado!=='En reparación'){
    showMsgIn('cierreMsg', 'Solo puedes cerrar un servicio que está en reparación.', false);
    return;
  }
  const refsValidas = validarRefaccionesCierre();
  if(!refsValidas.ok){
    showMsgIn('cierreMsg', refsValidas.error, false);
    return;
  }
  const fechaFin = todayStr();
  const horaFin = nowTimeStr();
  const netaMin = minutosReparacionNeta(rec, fechaFin, horaFin);
  if(!rec.fechaInicio || !rec.horaInicio){
    showMsgIn('cierreMsg', 'No se pudo calcular el tiempo de reparación (revisa el inicio del ticket).', false);
    return;
  }

  registros[idx] = {
    ...rec,
    estado: 'Cerrado',
    fechaFin, horaFin,
    falla: document.getElementById('f_falla').value,
    turno: document.getElementById('f_turno').value,
    paroMaquina: document.getElementById('f_paro').value,
    bloqueoCandadeo: document.getElementById('f_bloqueo').value,
    reparacionMin: netaMin,
    reparacionH: Number((netaMin/60).toFixed(2)),
    minutosPausa: minutosPausaAcumulados(rec, fechaFin, horaFin),
    pausaDesdeFecha: '',
    pausaDesdeHora: '',
    mantenimiento: document.getElementById('f_mantenimiento').value,
    refacciones: refsValidas.refacciones,
    observaciones: document.getElementById('f_observaciones').value,
    historial: [...rec.historial, {accion:'Cerrado', tecnicos:[], fecha:fechaFin, hora:horaFin}],
    cerradoEn: new Date().toISOString()
  };
  panelOpen = null;

  renderTable();
  renderKpis();
  renderTicketLists();

  const ok = await guardarRegistros();
  if(ok){
    showMsgIn('cierreMsg', 'Servicio #' + selectedTicketId + ' cerrado correctamente (' + horaFin + ').', true);
  }else{
    showMsgIn('cierreMsg', 'Cierre guardado en esta sesión, pero no se pudo sincronizar. Expórtalo a CSV para no perderlo.', false);
  }
  document.getElementById('cierreCard').style.display = 'none';
  selectedTicketId = null;
 }catch(err){
   console.error(err);
   showMsgIn('cierreMsg', 'Ocurrió un error inesperado al guardar: ' + (err && err.message ? err.message : err), false);
 }finally{
   setCerrandoServicio(false);
 }
}

function estadoClass(estado){
  return estado.replace(' ','').replace('ó','o');
}

function renderTable(){
  const tbody = document.getElementById('tbody');
  const empty = document.getElementById('emptyMsg');
  const q = (document.getElementById('searchBox').value||'').toLowerCase();
  const filtered = registros.filter(r=>{
    if(!q) return true;
    return (r.maquina+' '+r.area+' '+r.solicitante+' '+(r.problema||'')).toLowerCase().includes(q);
  });

  if(filtered.length===0){
    tbody.innerHTML='';
    empty.style.display='block';
    return;
  }
  empty.style.display='none';

  tbody.innerHTML = filtered.map(r=>`
    <tr>
      <td>#${r.id}</td>
      <td><span class="tag ${estadoClass(r.estado)}">${r.estado}</span></td>
      <td>${r.fecha}</td>
      <td>${r.area}</td>
      <td>${r.maquina}</td>
      <td>${r.solicitante}</td>
      <td><span class="tag ${r.prioridad}">${r.prioridad||'—'}</span></td>
      <td>${tecnicosVigentesTexto(r, ', ') || '—'}</td>
      <td>${r.esperaH!=null? r.esperaH.toFixed(2):'—'}</td>
      <td>${r.reparacionH!=null? r.reparacionH.toFixed(2):'—'}</td>
      <td>${r.turno||'—'}</td>
      <td>${r.paroMaquina||'—'}</td>
    </tr>
  `).join('');
}

function renderKpis(){
  document.getElementById('kpiTotal').textContent = registros.length;
  document.getElementById('kpiPendientes').textContent = registros.filter(r=>r.estado==='Pendiente').length;
  document.getElementById('kpiAsignados').textContent = registros.filter(r=>r.estado==='Asignado').length;
  document.getElementById('kpiEnCurso').textContent = registros.filter(r=>r.estado==='En reparación').length;
  document.getElementById('kpiPausados').textContent = registros.filter(r=>r.estado==='Pausado').length;
  const pausadosViejosN = registros.filter(esPausadoViejo).length;
  const kpiViejos = document.getElementById('kpiPausadosViejos');
  const kpiViejosBox = document.getElementById('kpiPausadosViejosBox');
  if(kpiViejos) kpiViejos.textContent = pausadosViejosN;
  if(kpiViejosBox) kpiViejosBox.classList.toggle('warn', pausadosViejosN>0);

  const conEspera = registros.filter(r=>r.esperaH!=null);
  document.getElementById('kpiEspera').textContent = conEspera.length
    ? (conEspera.reduce((s,r)=>s+r.esperaH,0)/conEspera.length).toFixed(1) + ' h' : '0.0 h';

  const cerrados = registros.filter(r=>r.estado==='Cerrado' && r.reparacionH!=null);
  document.getElementById('kpiReparacion').textContent = cerrados.length
    ? (cerrados.reduce((s,r)=>s+r.reparacionH,0)/cerrados.length).toFixed(1) + ' h' : '0.0 h';
}

function csvEscape(v){
  if(v===undefined||v===null) return '';
  const s = String(v).replace(/"/g,'""');
  return '"' + s + '"';
}

function historialTexto(hist){
  if(!hist) return '';
  return hist.map(h=>`${h.accion}${h.por?' por '+h.por:''}${h.tecnicos&&h.tecnicos.length? ' ('+h.tecnicos.join('/')+')':''} @ ${h.fecha} ${h.hora}`).join(' → ');
}

function exportarCSV(){
  const headers = ['ID','ESTADO','FECHA','NO. DE DIA','NO DE SEMANA','MES','AREA','MAQUINA','SOLICITANTE',
    'HORA SOLICITA','PRIORIDAD','FALLA','TIPO','PARO MAQUINA','TECNICOS','TECNICOS ACTIVOS','ASIGNADO FECHA','ASIGNADO HORA',
    'CONFIRMADO EN SITIO','FECHA ENTREGA','HORA DE ENTREGA',
    'INICIO REPARACION FECHA','INICIO REPARACION HORA','TIEMPO ESPERA (MIN) solicitud→sitio','TIEMPO ESPERA (h) solicitud→sitio',
    'TIEMPO REPARACION (MIN)','TIEMPO REPARACION (h)','PAUSA (MIN)','DURACION TOTAL (h)',
    'BLOQUE Y CANDADEO','FECHA LIBERADO','HORA LIBERADO','LIBERADO POR',
    'PROBLEMA','MANTENIMIENTO','REFACCIONES','TURNO PRODUCCION','OBSERVACIONES','HISTORIAL'];

  const rows = registros.slice().reverse().map(r=>{
    const totalH = (r.esperaH!=null && r.reparacionH!=null) ? Number((r.esperaH + r.reparacionH).toFixed(2)) : '';
    return [
      r.id, r.estado, r.fecha, r.diaNombre, r.semana, r.mes, r.area, r.maquina, r.solicitante,
      r.horaSolicita, r.prioridad, r.falla, r.tipo, r.paroMaquina, tecnicosVigentesTexto(r, ' / '),
      (r.tecnicosActivos||[]).join(' / '),
      r.fechaAsignado, r.horaAsignado, (r.confirmadoEnSitio ? 'Sí' : 'No'), r.fechaFin, r.horaFin,
      r.fechaInicio, r.horaInicio, r.esperaMin, r.esperaH,
      r.reparacionMin, r.reparacionH, r.minutosPausa!=null ? r.minutosPausa : '', totalH,
      r.bloqueoCandadeo||'', r.fechaLiberado||'', r.horaLiberado||'', r.liberadoPor||'',
      r.problema, r.mantenimiento, refaccionesTexto(r.refacciones), r.turno, r.observaciones, historialTexto(r.historial)
    ];
  });

  let csv = headers.map(csvEscape).join(',') + '\\n';
  csv += rows.map(row=>row.map(csvEscape).join(',')).join('\\n');

  const blob = new Blob(['\\ufeff'+csv], {type:'text/csv;charset=utf-8;'});
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = 'MANT_' + todayStr() + '.csv';
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

function minutosDeHora(hora){
  const p = String(hora||'').split(':');
  const h = Number(p[0]);
  const m = Number(p[1]||0);
  if(!Number.isFinite(h) || !Number.isFinite(m)) return null;
  return h*60 + m;
}

function restarUnDia(fechaStr){
  const d = new Date(fechaStr+'T00:00:00');
  if(isNaN(d.getTime())) return '';
  d.setDate(d.getDate()-1);
  return String(d.getFullYear())+'-'+
    String(d.getMonth()+1).padStart(2,'0')+'-'+
    String(d.getDate()).padStart(2,'0');
}

function fechaTurno(fechaFin, horaFin, turno){
  if(!fechaFin || !horaFin) return '';
  const mins = minutosDeHora(horaFin);
  if(mins==null) return '';
  const t = String(turno||'');
  const esNoche = t==='3°' || (!t && (mins>=22*60 || mins<7*60));
  if(esNoche && mins<7*60) return restarUnDia(fechaFin);
  return fechaFin;
}

document.getElementById('r_fecha').value = fechaTurno(todayStr(), nowTimeStr(), '') || todayStr();

function etiquetaTurnoResumen(turno){
  if(turno==='1°') return '1° (07:00–15:00)';
  if(turno==='2°') return '2° (15:00–22:00)';
  if(turno==='3°') return '3° (22:00–07:00)';
  return 'todos los turnos';
}

function generarResumen(){
  const turno = document.getElementById('r_turno').value;
  const fecha = document.getElementById('r_fecha').value;

  const items = registros.filter(r=>{
    if(r.estado!=='Cerrado' && r.estado!=='Liberado') return false;
    if(!r.fechaFin || !r.horaFin) return false;
    if(turno && r.turno !== turno) return false;
    if(fecha){
      const ft = fechaTurno(r.fechaFin, r.horaFin, r.turno);
      if(ft !== fecha) return false;
    }
    return true;
  }).slice().reverse();

  const fechaLabel = fecha || 'todas las fechas de inicio de turno';
  const turnoLabel = etiquetaTurnoResumen(turno);
  let texto = `*Reporte de mantenimiento — Turno ${turnoLabel}*\\n*Fecha (inicio del turno):* ${fechaLabel}\\n*Servicios atendidos:* ${items.length}\\n\\n`;

  if(items.length===0){
    texto += 'No hay servicios cerrados con esos filtros.';
  }else{
    items.forEach((r,i)=>{
      texto += `${i+1}. *${r.maquina}*\\n`;
      texto += `   Problema: ${r.problema || 'sin descripción'}\\n`;
      texto += `   Mantenimiento: ${r.mantenimiento || 'sin detalle'}\\n\\n`;
    });
  }

  const box = document.getElementById('resumenTexto');
  box.value = texto;
  box.style.display = 'block';
  document.getElementById('resumenActions').style.display = 'flex';
}

function copiarResumen(){
  const box = document.getElementById('resumenTexto');
  box.select();
  navigator.clipboard.writeText(box.value).then(()=>{
    showMsgIn('resumenMsg', 'Copiado. Ya lo puedes pegar en WhatsApp.', true);
  }).catch(()=>{
    showMsgIn('resumenMsg', 'No se pudo copiar automáticamente, selecciona el texto y copia manualmente (Ctrl+C).', false);
  });
}


function tickClock(){
  const d = new Date();
  document.getElementById('clockTime').textContent =
    String(d.getHours()).padStart(2,'0')+':'+String(d.getMinutes()).padStart(2,'0');
  document.getElementById('clockDate').textContent =
    DIAS[d.getDay()] + ' ' + d.getDate() + ' ' + MESES[d.getMonth()];
  document.getElementById('previewNow').textContent = nowLabel();
  const pFin = document.getElementById('previewNowFin');
  if(pFin) pFin.textContent = nowLabel();
  actualizarElapsed();
  actualizarPreviewCierre();
}
tickClock();
setInterval(tickClock, 1000*15);

setInterval(function(){
  if(!urlConfigurada()) return;
  if(!supervisorUnlocked && !tecnicoUnlocked) return;
  if(!vistaSupervisorOTecnico()) return;
  cargarRegistros(true);
}, 45000);

actualizarControlesGuardado();
cargarRegistros();
</script>
</body>
</html>
"""

html = html.replace('__LISTS_JSON__', lists_js).replace('__TURNOS_JSON__', turnos_js).replace('__SUPERVISOR_PIN__', SUPERVISOR_PIN).replace('__GERENTE_MANTENIMIENTO__', GERENTE_MANTENIMIENTO).replace('__SHEET_WEBAPP_URL__', SHEET_WEBAPP_URL)

with open(SALIDA, 'w', encoding='utf-8') as f:
    f.write(html)

print(f"\nListo: se genero '{SALIDA}' con las listas actualizadas.")
print(f"   Recuerda revisar el PIN de supervisor ('SUPERVISOR_PIN = \"{SUPERVISOR_PIN}\"') y el Gerente de Mantenimiento ('GERENTE_MANTENIMIENTO = \"{GERENTE_MANTENIMIENTO}\"') en este script si los quieres cambiar.")
print(f"   Los tecnicos ahora entran con su PIN individual (columna PIN del Excel, junto a TECNICO).")
if SHEET_WEBAPP_URL.startswith("PEGA_AQUI"):
    print("   AVISO: SHEET_WEBAPP_URL todavia no esta configurada — el formulario guardara solo en la sesion actual.")
else:
    print(f"   Conectado a Google Sheets via: {SHEET_WEBAPP_URL}")

