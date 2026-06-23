"""
Findclo Checker — API
Expone el checker como un servicio HTTP para que Findclo lo consuma.
Toda la logica de scraping esta en checker.py, este archivo solo agrega
la parte de entrada/salida 

Uso:
    uvicorn api:app --host 0.0.0.0 --port 8000
"""

import asyncio
import os
from datetime import datetime, timedelta

import asyncpg
from fastapi import FastAPI, HTTPException, BackgroundTasks
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel

# Importamos solo lo que necesitamos de checker.py
from checker import (
    check_product_browser,
    detect_platform_from_url,
    quick_http_check,
)
from playwright.async_api import async_playwright

app = FastAPI(title="Findclo Checker API", version="1.0.0")

# Permitir requests desde Findclo (Next.js en Vercel)
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],  # en produccion poner el dominio de Findclo
    allow_methods=["*"],
    allow_headers=["*"],
)

# URL de conexion a Neon — se configura como variable de entorno
DATABASE_URL = os.environ.get("DATABASE_URL", "")

CACHE_DAYS = 7  # mismo criterio que el checker (cambiar)


# Modelos de entrada y salida

class CheckRequest(BaseModel):
    url: str
    name: str = ""  # nombre del producto (opcional, para guardar en la BD)

class CheckResult(BaseModel):
    url: str
    name: str
    status: str
    reason: str
    checked_at: str


# Conexion a Neon

async def get_db():
    return await asyncpg.connect(DATABASE_URL)


# Crear la tabla si no existe (correr una vez al iniciar)

async def init_db():
    db = await get_db()
    await db.execute("""
        CREATE TABLE IF NOT EXISTS stock_status (
            url         TEXT PRIMARY KEY,
            name        TEXT,
            status      TEXT,
            reason      TEXT,
            checked_at  TIMESTAMP
        )
    """)
    await db.close()


@app.on_event("startup")
async def startup():
    if DATABASE_URL:
        await init_db()


# Leer resultado guardado en Neon

async def get_cached(url: str) -> dict | None:
    db = await get_db()
    row = await db.fetchrow(
        "SELECT * FROM stock_status WHERE url = $1", url
    )
    await db.close()
    if not row:
        return None
    checked_at = row["checked_at"]
    if datetime.now() - checked_at < timedelta(days=CACHE_DAYS):
        return dict(row)
    return None  # cache vencido


# Guardar resultado en Neon

async def save_result(url: str, name: str, status: str, reason: str):
    db = await get_db()
    await db.execute("""
        INSERT INTO stock_status (url, name, status, reason, checked_at)
        VALUES ($1, $2, $3, $4, $5)
        ON CONFLICT (url) DO UPDATE
            SET name = $2, status = $3, reason = $4, checked_at = $5
    """, url, name, status, reason, datetime.now())
    await db.close()


# Chequear una URL con Playwright (misma logica que checker.py)

async def run_check(url: str, name: str) -> dict:
    # Primero chequeo rapido por HTTP
    quick = quick_http_check(url)
    if quick == "PRODUCTO_ELIMINADO":
        return {"status": "PRODUCTO_ELIMINADO", "reason": "HTTP 404"}

    # Si necesita browser, abrimos Playwright
    async with async_playwright() as p:
        browser = await p.chromium.launch(
            headless=True,
            args=["--disable-blink-features=AutomationControlled"],
        )
        context = await browser.new_context(
            user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
            viewport={"width": 1280, "height": 800},
            extra_http_headers={"Accept-Language": "es-AR,es;q=0.9"},
        )

        # Bloquear imagenes para que cargue mas rapido
        async def _block(route, request):
            if request.resource_type in ("image", "media", "font", "other"):
                await route.abort()
            else:
                await route.continue_()
        await context.route("**/*", _block)

        page = await context.new_page()
        product = {"url": url, "name": name, "description": ""}
        result = await check_product_browser(page, product)
        await browser.close()

    return {"status": result["status"], "reason": result.get("reason", "")}


# Endpoints


@app.get("/")
async def root():
    return {"status": "ok", "service": "Findclo Checker API"}


@app.post("/check", response_model=CheckResult)
async def check(req: CheckRequest):
    """
    Chequea el stock de una URL.
    Si el resultado es reciente (menos de 7 dias), devuelve el cache de Neon.
    Si no, escanea la pagina y guarda el resultado nuevo.
    """
    # Intentar desde cache primero
    if DATABASE_URL:
        cached = await get_cached(req.url)
        if cached:
            return CheckResult(
                url=cached["url"],
                name=cached["name"] or req.name,
                status=cached["status"],
                reason=cached["reason"] or "",
                checked_at=cached["checked_at"].isoformat(),
            )

    # No hay cache valido, escanear
    try:
        result = await run_check(req.url, req.name)
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

    # Guardar en Neon
    if DATABASE_URL:
        await save_result(req.url, req.name, result["status"], result["reason"])

    return CheckResult(
        url=req.url,
        name=req.name,
        status=result["status"],
        reason=result["reason"],
        checked_at=datetime.now().isoformat(),
    )


@app.post("/check/bulk")
async def check_bulk(items: list[CheckRequest], background_tasks: BackgroundTasks):
    """
    Acepta una lista de productos y los procesa en background.
    Devuelve inmediatamente con un mensaje de confirmacion.
    Findclo puede leer los resultados de Neon cuando esten listos.
    """
    async def process_all():
        for item in items:
            try:
                cached = await get_cached(item.url) if DATABASE_URL else None
                if cached:
                    continue  # ya esta en cache, no re-escanear
                result = await run_check(item.url, item.name)
                if DATABASE_URL:
                    await save_result(item.url, item.name, result["status"], result["reason"])
            except Exception:
                pass  # si falla uno, seguimos con el resto (si, falla, volver a re-procesar)

    background_tasks.add_task(process_all)
    return {"message": f"{len(items)} productos en cola", "status": "processing"}


@app.get("/status")
async def get_status(url: str):
    """
    Devuelve el ultimo resultado guardado para una URL, sin re-escanear.
    Util para que Next.js consulte el estado directamente.
    """
    if not DATABASE_URL:
        raise HTTPException(status_code=503, detail="Base de datos no configurada")
    db = await get_db()
    row = await db.fetchrow("SELECT * FROM stock_status WHERE url = $1", url)
    await db.close()
    if not row:
        raise HTTPException(status_code=404, detail="URL no encontrada")
    return dict(row)
