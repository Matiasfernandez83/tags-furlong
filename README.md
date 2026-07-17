# PeajeControl Go

Control de peajes y gastos de flota (SaaS multi-empresa). Lee facturas PDF (AUSA/AUSOL/TelePASE),
extrae los pases, cruza cada TAG contra la base de flota y consolida gastos. **Sin IA, sin API keys.**
Flask + pdfplumber. Base de datos: SQLite en local, PostgreSQL (Supabase) en produccion.

## Modulos
- **Tablero**: KPIs con doble-click (abren detalle), graficos por responsable y verificacion de TAGs.
- **Importar**: sube facturas PDF/Excel de peajes; matching contra la base maestra + semaforo de control.
- **Gastos Tarjetas**: sube resumenes bancarios; tarjetas por banco, deuda consolidada, categorizacion.
- **Reportes**: historial de archivos con re-descarga del PDF original y export maestro a Excel.
- **Configuracion**: usuarios, cambio de clave, estado del sistema.

## Correr local
    pip install -r requirements.txt
    python app.py
Abrir http://localhost:5000 — usuario inicial: admin@peajecontrol.com / admin
(cambiar la clave desde Configuracion apenas ingreses).

## Deploy en Render (plan free) + Supabase
El `render.yaml` deja el Web Service listo en plan **free** (sin disco pago). La base
corre en **PostgreSQL gratuito de Supabase**. Pasos detallados en [`docs/DEPLOY.md`](docs/DEPLOY.md).
- Build: pip install -r requirements.txt
- Start: gunicorn app:app --workers 1 --timeout 120
- `SECRET_KEY` se genera sola; hay que cargar **`DATABASE_URL`** (conexion de Supabase)
  a mano en el dashboard de Render.
