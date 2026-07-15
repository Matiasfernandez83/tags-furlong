# TAGs Transporte Furlong

Gestion de TAGs y peajes de flota. Lee facturas PDF (AUSA/AUSOL/TelePASE), extrae los pases,
cruza cada TAG contra la base de flota y consolida gastos. **Sin IA, sin API keys.**
Flask + SQLite + pdfplumber.

## Modulos
- **Tablero**: KPIs con doble-click (abren detalle), graficos por responsable y verificacion de TAGs.
- **Importar**: sube facturas PDF/Excel de peajes; matching contra la base maestra + semaforo de control.
- **Gastos Tarjetas**: sube resumenes bancarios; tarjetas por banco, deuda consolidada, categorizacion.
- **Reportes**: historial de archivos con re-descarga del PDF original y export maestro a Excel.
- **Configuracion**: usuarios, cambio de clave, estado del sistema.

## Correr local
    pip install -r requirements.txt
    python app.py
Abrir http://localhost:5000 — usuario inicial: admin@transportefurlong.com.ar / admin
(cambiar la clave desde Configuracion apenas ingreses).

## Deploy en Render
El archivo `render.yaml` ya deja todo configurado (disco persistente incluido).
- Build: pip install -r requirements.txt
- Start: gunicorn app:app
- SECRET_KEY se genera sola; DB_PATH apunta al disco persistente.
