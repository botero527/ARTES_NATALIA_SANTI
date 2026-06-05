"""
probar_arte.py — Test rápido del pipeline de arte sobre el DWG activo en AutoCAD.
Ejecutar:  python probar_arte.py
Muestra todo en consola en tiempo real.
"""
import sys
import traceback
import time

# ── Verificar dependencias ────────────────────────────────────────────────────
try:
    import win32com.client
    import pythoncom
except ImportError:
    print("[ERROR] Falta pywin32.  Ejecuta:  pip install pywin32")
    input("Enter para salir...")
    sys.exit(1)

# ── Logger con timestamp ──────────────────────────────────────────────────────
def clog(msg):
    ts = time.strftime("%H:%M:%S")
    print(f"[{ts}] {msg}", flush=True)

# ── Importar pipeline ─────────────────────────────────────────────────────────
clog("Importando pipeline desde crear_arte_acad.py...")
try:
    from crear_arte_acad import pipeline
    clog("  Pipeline importado OK")
except Exception as e:
    clog(f"  ERROR importando pipeline: {e}")
    traceback.print_exc()
    input("Enter para salir...")
    sys.exit(1)

# ── Conectar a AutoCAD ────────────────────────────────────────────────────────
clog("Conectando a AutoCAD...")
pythoncom.CoInitialize()
try:
    acad = win32com.client.GetActiveObject("AutoCAD.Application")
    doc  = acad.ActiveDocument
    clog(f"  Documento activo: {doc.Name}")
except Exception as e:
    clog(f"  ERROR: {e}")
    clog("  Abre AutoCAD con un plano antes de ejecutar esto.")
    input("Enter para salir...")
    pythoncom.CoUninitialize()
    sys.exit(1)

# ── Confirmar ─────────────────────────────────────────────────────────────────
print()
print("=" * 60)
print(f"  Se va a procesar:  {doc.Name}")
print("=" * 60)
resp = input("¿Continuar? (s/n): ").strip().lower()
if resp != "s":
    clog("Cancelado.")
    pythoncom.CoUninitialize()
    sys.exit(0)

# ── Ejecutar pipeline ─────────────────────────────────────────────────────────
print()
clog("=== INICIANDO PIPELINE ===")
t0 = time.time()

try:
    doc.Activate()
    pipeline(doc, log_fn=clog)
    elapsed = time.time() - t0
    print()
    clog(f"=== PIPELINE COMPLETADO en {elapsed:.1f}s ===")
except Exception as e:
    elapsed = time.time() - t0
    print()
    clog(f"=== ERROR FATAL en {elapsed:.1f}s ===")
    clog(f"  {e}")
    print()
    traceback.print_exc()
finally:
    pythoncom.CoUninitialize()

print()
input("Enter para cerrar...")
