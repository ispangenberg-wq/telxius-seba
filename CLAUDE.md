# Telxius - Generador Carga Salesforce

## Qué es

Herramienta web interna para Telxius Cable (empresa de infraestructura de telecomunicaciones, subsidiaria de Telefónica). Transforma un Excel de provisiones mensual en el formato de carga que necesita Salesforce.

## Cómo funciona

El usuario sube un Excel de provisiones (.xlsx) por el frontend y descarga otro Excel listo para importar en Salesforce.

### Excel de entrada

Tiene 3 pestañas de datos + pestañas de lookup:

- **Lease** (o **CAP** en Brasil) — circuitos, enlaces, cross connects. Tiene columnas NRC y MRC.
- **O&M** — operaciones y mantenimiento (fibra oscura, housing, power). Solo tiene MRC (no NRC).
- **IP** — tránsito IP, seguridad (escudos DDoS). Tiene NRC y MRC.
- **Informe mes actual Provisiones-** (obligatoria) — tabla de lookup con columnas `EFC Number` → `Elemento a Facturar ID`. Se usa para mapear tanto EP como EP EXTORNO.

### Columnas clave del Excel de entrada

Cada pestaña (Lease/O&M/IP) tiene dos secciones: mes actual y mes anterior.

**Mes actual (provisiones positivas):**
- `Elemento a Provisionar (EP)` — código EFC que se busca en la lookup de provisiones
- `Invoice Period` — determina Inicio/Fin Período de Facturación para provisiones
- `NRC` — cargo no recurrente (solo Lease e IP)
- `MRC` — cargo mensual recurrente

**Mes anterior (extornos):**
- `EP EXTORNO >>> SI APLICA` — código EFC que se busca en la misma lookup de provisiones
- `Invoice Period.1` — determina Inicio/Fin Período de Facturación para extornos
- `NRC.1` / `MRC.1` — montos del mes anterior (cualquier valor != 0, siempre se convierte a negativo con `-abs()`)

### Lógica de procesamiento

1. Para cada fila con EP válido (empieza con "EFC") en Lease/O&M/IP:
   - Buscar EP en "Informe mes actual Provisiones-" → obtener `Elemento a Facturar ID`
   - Si MRC > 0 → generar fila en Provisiones Positivas (Tipo de Cargo = "MRC", o "O&M" si viene de pestaña O&M)
   - Si NRC > 0 → generar fila en Provisiones Positivas (Tipo de Cargo = "NRC")

2. Para cada fila con EP EXTORNO válido:
   - Buscar EP EXTORNO en "Informe mes actual Provisiones-" (misma lookup) → obtener `Elemento a Facturar ID`
   - Tomar montos de columnas "Mes anterior" (MRC.1, NRC.1) — cualquier valor != 0
   - Convertir siempre a negativo (`-abs(valor)`) y generar filas en Extornos

### Excel de salida

Dos pestañas: **Provisiones Positivas** y **Extornos**. Mismas columnas:

| Columna | Valor |
|---------|-------|
| Elemento a facturar ID | De la lookup |
| EFC Number | El EP o EP EXTORNO |
| Pendiente de revision Local | Siempre "Revisado" |
| Estado de factura | Siempre "Provisionado" |
| Año de facturación | Año actual (hoy) |
| Mes de facturación | Mes actual (hoy) |
| Inicio Período de Facturación | Parseado del Invoice Period |
| Fin Período de Facturación | Parseado del Invoice Period |
| Tipo de Cargo | "MRC", "NRC", o "O&M" |
| Importe en Curso | Monto (positivo en provisiones, negativo en extornos) |

### Parseo del Invoice Period

- `"April, 2026"` → 01/04/2026 - 30/04/2026
- `"burst April, 2026"` → 01/04/2026 - 30/04/2026 (ignora prefijos)
- `"January - April, 2026"` → 01/01/2026 - 30/04/2026 (rango con un año)
- `"September, 2021 - December, 2024"` → 01/09/2021 - 31/12/2024 (rango con año en ambos lados)
- `"(September, 2021 - December, 2024) > cuota 11de24"` → 01/09/2021 - 31/12/2024 (ignora texto extra)
- `"January - April"` (sin año) → 01/01/YYYY - 30/04/YYYY (usa año actual)
- Vacío → 01/01/YYYY - 31/12/YYYY (año completo)

### Formato de números

- Columna "Importe en Curso": redondeado a 2 decimales, formato Excel `#,##0.00` (ej: `2,820.00`, `-3,333.33`)

## Stack técnico

- **Backend:** Python + FastAPI
- **Procesamiento Excel:** pandas + openpyxl
- **Frontend:** HTML/CSS/JS vanilla (sin framework)
- **Deploy:** Vercel (serverless Python)
- **Repo:** github.com/ispangenberg-wq/telxius-seba

## Estructura de archivos

```
app.py              → versión local (para correr con uvicorn directo, monta /static para archivos estáticos)
api/index.py        → versión Vercel (serverless, misma lógica)
static/index.html   → frontend con branding Telxius (favicon de telxius.com, fondo imagen cable submarino)
static/bg.jpg       → imagen de fondo (cable submarino de fibra óptica)
requirements.txt    → dependencias Python
vercel.json         → config de deploy Vercel (rutas para API + archivos estáticos)
```

## Cómo correr local

```
pip install -r requirements.txt
uvicorn app:app --host 0.0.0.0 --port 8000
```

Abrir http://localhost:8000

## Contexto de negocio

- Telxius opera en múltiples países: Argentina (provARG), Brasil (provBRA), Chile (provCHI). Los clientes incluyen Telefónica, Google, Amazon, Antel, Starlink, etc.
- Cada país puede tener nombres de columnas ligeramente distintos (ej: "CAP" en lugar de "Lease", "MRC Neto" en lugar de "MRC", "Period" en lugar de "Invoice Period"). El código usa detección dinámica con `find_col()`.
- Este proceso se hace mensualmente. El Excel de entrada siempre viene con la misma estructura base.
- "Provisiones" = facturación estimada del mes actual. "Extornos" = reversiones de provisiones del mes anterior.
- El output se importa en Salesforce para registrar la facturación.
