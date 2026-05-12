# Conexión a la base de datos local (PostgreSQL)

> **Se usa PostgreSQL 14 instalado con Homebrew** (NO Docker).
> El script `dev_backend.sh` lo arranca automáticamente y crea el rol/base si no existen.

---

## Paso a paso para levantar el entorno local

### 1. Iniciar PostgreSQL (Homebrew)

```bash
brew services start postgresql@14
```

Solo hace falta la primera vez o si se reinició la máquina. El script de dev lo hace automáticamente si no está corriendo.

### 2. Activar el entorno virtual y levantar el backend

```bash
source /Users/inti/GitHub/vyp/app360/.venv/bin/activate
cd /Users/inti/GitHub/vyp/app360
bash scripts/dev_backend.sh
```

El script hace todo solo:
- Verifica que PostgreSQL esté corriendo (y lo arranca con Homebrew si no)
- Crea el rol `app360` y la base `app360` si no existen
- Corre las migraciones / setup de tablas
- Levanta el servidor Flask en modo desarrollo

---

## Datos de conexión

| Campo    | Valor       |
|----------|-------------|
| Host     | `localhost` |
| Puerto   | `5432`      |
| Base     | `app360`    |
| Usuario  | `app360`    |
| Password | `password`  |

**URL completa:**
```
postgresql://app360:password@localhost:5432/app360
```

### Conectarse con psql

```bash
psql -h localhost -p 5432 -U app360 -d app360
```

### Conectarse desde un notebook Python

```python
import sqlalchemy as sa

engine = sa.create_engine("postgresql://app360:password@localhost:5432/app360")

with engine.connect() as conn:
    result = conn.execute(sa.text("SELECT COUNT(*) FROM stg_sabiogo_ventas"))
    print(result.scalar())
```

---

## Tablas SabioGo Staging

Las tablas `stg_sabiogo_compras` y `stg_sabiogo_ventas` tienen **la misma estructura** (datos crudos importados desde archivos del ERP SabioGo).

| Columna | Tipo | Descripción |
|---|---|---|
| `id` | BigInteger PK | — |
| `rubro` | Text | — |
| `fecha` | DateTime | Fecha del comprobante |
| `subrubro` | Text | — |
| `codigo` | Text | Código de producto (normalizado) |
| `codigo_raw` | String(80) | Código original del archivo fuente |
| `modelo` | Text | — |
| `producto` | Text | Descripción del producto |
| `marca` | Text | — |
| `comprobante` | Text | Tipo de comprobante |
| `nro_comprob` | Text | Número de comprobante |
| `cliente` | Text | — |
| `condicion_iva` | Text | — |
| `vendedor` | Text | — |
| `deposito` | Text | — |
| `uninegocio` | Text | Unidad de negocio |
| `forma_pago` | Text | — |
| `kilos` | Float | — |
| `cajas` | Float | — |
| `dtos` | Float | Descuentos |
| `neto` | Float | Importe neto |
| `localidad` | Text | — |
| `zona` | Text | — |
| `sucursal` | Text | — |
| `iva` | Float | Porcentaje IVA |
| `moniva` | Float | Monto IVA |
| `tipo` | Text | — |
| `subtot` | Float | Subtotal |
| `numlispre` | Integer | Número de lista de precios |
| `lisprec` | Text | Lista de precios |
| `origen` | String(20) | `CARNES`, `LACTEOS` o `LEGACY` (default) |
| `fuente_archivo` | Text | Nombre del archivo importado |
| `hash_archivo` | Text | Hash del archivo para deduplicación |
| `estado` | String(50) | `cargado` por defecto |
| `fecha_import` | DateTime | Timestamp de importación |
| `created_at` | DateTime | — |
| `updated_at` | DateTime | — |

**Unique constraint:** `(nro_comprob, tipo, fecha)` por tabla.
**Índices:** `fecha`, `codigo`, `vendedor`, `sucursal`, `origen`, `(origen, fecha, codigo)`.