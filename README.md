## Como funciona

1. Lee los CSVs de la carpeta `input/` (uno por marca)
2. Para cada producto, entra a la URL y detecta si hay stock
3. Guarda los resultados en un Excel dentro de `output/`
4. El cache evita re-chequear productos revisados hace menos de 7 dias

## Requisitos

- Python 3.11+
- Las dependencias listadas en `requirements.txt`

## Instalacion

```bash
pip install -r requirements.txt
playwright install chromium
```

## Uso local (Excel)

Poner los CSVs de Findclo en la carpeta `input/` y correr:

```bash
python checker.py
```

El Excel se genera en `output/` con la fecha y hora en el nombre.

## Uso como API

```bash
uvicorn api:app --host 0.0.0.0 --port 8000
```

Requiere la variable de entorno `DATABASE_URL` con la connection string de Neon.

### Endpoints

| Metodo | Ruta | Descripcion |
|--------|------|-------------|
| POST | `/check` | Chequea una URL y devuelve el estado |
| POST | `/check/bulk` | Chequea una lista de URLs en background |
| GET | `/status?url=...` | Devuelve el ultimo resultado guardado |

### Ejemplo

```bash
curl -X POST http://localhost:8000/check \
  -H "Content-Type: application/json" \
  -d '{"url": "https://tienda.com/producto/123", "name": "Remera Boxy"}'
```

```json
{
  "url": "https://tienda.com/producto/123",
  "name": "Remera Boxy",
  "status": "OK",
  "reason": "",
  "checked_at": "2026-06-09T21:32:00"
}
```

### Estados posibles

| Status | Significado |
|--------|-------------|
| `OK` | Tiene stock |
| `REVISAR` | Algunos talles sin stock |
| `SIN_STOCK` | Sin stock en todos los talles |
| `PRODUCTO_ELIMINADO` | La URL no existe mas (404) |
| `ERROR` | No se pudo cargar la pagina |

## Variables de entorno

| Variable | Descripcion |
|----------|-------------|
| `DATABASE_URL` | Connection string de Neon (PostgreSQL) |

## Estructura

```
findclo_checker/
├── checker.py        # logica de scraping
├── api.py            # API FastAPI (usa checker.py)
├── requirements.txt  # dependencias
├── input/            # CSVs de Findclo 
├── output/           # Excels generados 
└── cache.json        # cache local 
```

## Deploy en Railway

1. Conectar el repo en Railway
2. Agregar la variable de entorno `DATABASE_URL`
3. Comando de inicio: `uvicorn api:app --host 0.0.0.0 --port $PORT`
