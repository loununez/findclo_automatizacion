"""
Findclo Product Availability Checker

- Lee todos los CSVs de input/
- Agrupa por marca
- Cache de 7 días: URLs ya chequeadas no se re-procesan
- El Excel final siempre incluye TODOS los productos (caché + nuevos)
- Concurrencia aumentada para mayor velocidad

Uso:
    python checker.py
"""

import asyncio
import concurrent.futures
import csv
import glob
import io
import json
import re
import requests
from collections import Counter, defaultdict
from datetime import datetime, timedelta
from pathlib import Path
from urllib.parse import urlparse

import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment
from playwright.async_api import async_playwright


# Carpetas y caché

BASE_DIR   = Path(__file__).parent
INPUT_DIR  = BASE_DIR / "input"
OUTPUT_DIR = BASE_DIR / "output"
CACHE_FILE = BASE_DIR / "cache.json"

INPUT_DIR.mkdir(exist_ok=True)
OUTPUT_DIR.mkdir(exist_ok=True)

CACHE_DAYS = 7          # re-chequear después de 7 días
CONCURRENCY = 7         # páginas en paralelo (balance velocidad/estabilidad)
HTTP_TIMEOUT = 5        # segundos para check rápido HTTP
BROWSER_TIMEOUT = 20000 # ms para Playwright (reducido de 30s a 20s)


# Señales de stock

OUT_OF_STOCK_SIGNALS = [
    "notificarme cuando haya stock",
    "notificarme cuando este disponible",
    "notificarme cuando esté disponible",
    "avisame cuando haya stock",
    "avisame cuando este disponible",
    "avisame cuando esté disponible",
    "avísame cuando esté disponible",
    "avisame cuando este disponible",
    "sin stock",
    "sold out",
    "agotado",
    "out of stock",             # Bowen (Magento) usa texto en inglés
    "talle agotado",            # Allo Martinez (jQuery)
    "proximamente",
    "próximamente",
    # "ver similares" → solo válido en Zara, en otras marcas es botón de navegación normal
    # "notificarme" solo → demasiado genérico, aparece en newsletters/popups
    # "no disponible" solo → puede aparecer en contextos no relacionados al stock
]

# Señales válidas SOLO para plataformas específicas
PLATFORM_OUT_SIGNALS = {
    "hcl":  ["ver similares"],  
    "zara": ["ver similares"],
}

IN_STOCK_SIGNALS = [
    "agregar al carrito",
    "añadir al carrito",
    "agregar a mi compra",
    "añadir a la bolsa",
    "agregar a la bolsa",       
    "agregar al changuito",     
    "agregar a mi bolsa",      
    "add to cart",
    "comprar ahora",
]

PLATFORM_BUY_SELECTORS = {
    "tiendanube": [
        "input[data-store*='product-buy-button']",  
        "input.js-prod-submit-form",
        "input.js-addtocart",
        "button.add-to-cart-button",
        "[data-store='add-to-cart']",
        ".buy-button",
        "form.buy-form button[type='submit']",
        "button.js-add-to-cart",
    ],
    "shopify": [
        "button[name='add']",
        ".product-form__submit",
        "button.add-to-cart",
        "[data-testid='add-to-cart-button']",
    ],
    "magento": [
        "button.action.tocart",
        "#product-addtocart-button",
        "button.tocart",
    ],
    "vtex": [
        ".vtex-store-components-3-x-buyButton",
        "button[data-testid='buy-button']",
        ".vtex-button",
        ".buy-button",
        "button[class*='buyButton']",
        "button[class*='BuyButton']",
    ],
    "jquery": [
        "button.add-to-cart",
        "button[data-action='add-to-cart']",
        ".btn-addtocart",
        "button.comprar",
        "input[type='submit'].item-show-btn",   
        "input[type='submit'][value*='Agregar']",
        "input[type='submit'][value*='agregar']",
        ".js-buy-btn",                          
    ],
    "angular": [
        "button[class*='buy']",
        "button[class*='comprar']",
        ".buy-button",
    ],
    "prestashop": [
        "#add-to-cart-or-refresh button",
        "button.add-to-cart",
        "#product-details button[data-button-action='add-to-cart']",
    ],
}

UNAVAILABLE_SIZE_SELECTORS = {
    "tiendanube": [
        "[data-value][data-unavailable]",
        ".product-variants--disabled",
        "li.option--disabled",
        "li.is-unavailable",
        ".variant--unavailable",
    ],
    "shopify": [
        "option[disabled]",
        ".variant-option--disabled",
        "input[disabled] + label",
        ".swatch.is-disabled",
    ],
    "magento": [
        ".swatch-option.disabled",
        "option.swatch-option[data-option-id]",
        "[option-empty]",
    ],
    "vtex": [
        "[aria-disabled='true']",
        ".sku-selector-option--disabled",
        ".selected-item--unavailable",
        "[class*='pdp-components'][class*='disabled']",
        "[class*='pdp-components'][class*='unavailable']",
        "[class*='pdp-components'][class*='Unavailable']",
        "[class*='pdp-components'][class*='out']",
    ],
    "jquery": [
        ".size-unavailable",
        "li.unavailable",
        "option[disabled]",
        "[aria-disabled='true']",
    ],
    "angular": [
        "[class*='disabled']",
        "[aria-disabled='true']",
        ".size-unavailable",
    ],
    "prestashop": [
        ".attribute-radio[disabled]",
        ".product-variants option[disabled]",
        ".combination-unavailable",
    ],
    "generic": [
        "[aria-disabled='true']",
        "[disabled]",
        ".disabled",
        ".out-of-stock",
        ".unavailable",
    ],
}

TOTAL_SIZE_SELECTORS = {
    "tiendanube": [
        "[data-value]",
        ".product-variants li",
        ".js-variant-option",
        "form.buy-form .variant-optionlist li",
    ],
    "shopify": [
        "fieldset.variant-input-wrap label",
        "label[data-value]",
        "button[data-value]",
        ".swatch-element label",
        ".product-form__option .btn",
        ".product-form__option label",
        "select.single-option-selector option:not([value=''])",
        ".js-product-select option:not([value=''])",
        ".variant__button-label",
        "input[type='radio'][name*='Size'] + label",
        "input[type='radio'][name*='Talle'] + label",
    ],
    "magento":    [".swatch-option", ".product-options-wrapper option:not([value=''])"],
    "vtex":       [".sku-selector-option", "[data-testid='sku-selector-item']"],
    "generic":    ["option:not([value=''])", ".size-option", ".variant-option"],
}


# Caché

def load_cache() -> dict:
    if CACHE_FILE.exists():
        try:
            with open(CACHE_FILE, encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return {}
    return {}


def save_cache(cache: dict):
    with open(CACHE_FILE, "w", encoding="utf-8") as f:
        json.dump(cache, f, ensure_ascii=False, indent=2)


def is_cache_valid(entry: dict) -> bool:
    """True si el entry del caché tiene menos de CACHE_DAYS días."""
    try:
        checked_at = datetime.fromisoformat(entry["checked_at"])
        return datetime.now() - checked_at < timedelta(days=CACHE_DAYS)
    except Exception:
        return False


# Lector de CSVs

def extract_brand(filename: str) -> str:
    stem = Path(filename).stem
    match = re.search(r"Demo\s+(.+?)_", stem, re.IGNORECASE)
    if match:
        return match.group(1).strip()
    parts = stem.split("_")
    return parts[1].strip() if len(parts) > 1 else stem


def detect_platform_from_url(url: str) -> str:
    domain = urlparse(url).netloc.lower()

    # Tienda Nube
    if any(x in domain for x in [
        "tiendanube", "mitiendanube",
        "kingofthekongo", "kostume",
        "kabrakuervo", "moscowshop", "midwaystore", "midway.com.ar",
        "newbalance.com.ar",
        "honkytonk",
        "etiquetanegra",
        "cardon.com.ar",
        "harvey.com.ar",
        "elepants",
        "perramus",
        "bombaproject", "bombashop",
        "anastasiamonaco",
        "tussy.com.ar",
        "elclubdelsweater",
        "mono-indumentaria", "monoindumentaria",
        "wearelse",
        "ossira.com.ar",
        "naima.com.ar",
        "oraactive",
        "justapetra",
        "cibeles.com.ar",
        "heyblackk",
        "topwhite",
        "iman.com.ar",
        "kiwano.com.ar",
        "fulanadetal",
        "tiendanapoli", "napoli",
        "ika.com.ar",
        "bensimon.com.ar",
        "bullbenny",
        "gola.com.ar",
        "rever.com.ar",
        "eldon.com.ar",
        "prototype.com.ar",
        "madlowell",
        "londontie",
    ]):
        return "tiendanube"

    # Shopify
    if any(x in domain for x in [
        "myshopify", "aynotdead", "shopify",
        "zadig-et-voltaire", "zadigvoltaire",
        "timberland.com.ar",
        "delaostia",
        "salomon.com",
        "revolver.com.ar",
        "naum.com.ar",
        "blackmamba.com.ar",
        "sofiasarkani",
        "vcp.com.ar",
    ]):
        return "shopify"

    # Magento
    if any(x in domain for x in [
        "label-99", "birkenstock", "jazminchebar", "kosiuko", "mariacher",
        "laspepas", "tucci.com.ar", "markova.com.ar",
        "rapsodia.com", "complot.com.ar",
        "akiabara.com",
        "burgues.com.ar",
        "bowen.com.ar",
        "devre.com.ar",
        "isadora.com.ar",
        "herenciacustomgarage",
    ]):
        return "magento"

    if any(x in domain for x in [
        "rickysarkany", "ginebra", "levis", "nike",
        "lamartina.com",
        "tommy.com", "tommyhilfiger",
        "calvinklein.com.ar",
        "portsaid.com.ar",
        "cuestablanca.com.ar",
        "desiderata.com.ar",
        "ayres.com.ar",
        "lovelydenim",
        "keybiscayne.com.ar",
        "tascani.com.ar",
    ]):
        return "vtex"

    if any(x in domain for x in [
        "puma", "adidas",
        "boken.com.ar",      
    ]):
        return "react"

    # jQuery
    if any(x in domain for x in [
        "sweet.com.ar", "underarmour.com", "wanama.com",
        "hushpuppies.com.ar",
        "bolivia.com.ar",
        "kevingston.com.ar",
        "estancias",
        "mishka.com.ar",
        "adrianaconstantini",
        "allomartinez",
        "penguinargentina.com",  # Penguin (Rails+jQuery, buy button is <input type="submit">)
        "awada.com.ar",          # Awada (misma plataforma Rails que Penguin)
    ]):
        return "jquery"

    # Angular
    if any(x in domain for x in [
        "paulacahen.com.ar",
    ]):
        return "angular"

    # PrestaShop
    if any(x in domain for x in ["brooksfield.com.ar"]):
        return "prestashop"

    return "generic"


IMAGE_EXTENSIONS = (".png", ".jpg", ".jpeg", ".webp", ".gif", ".svg", ".avif")
IMAGE_CDN_HINTS  = ["images.", "/image/", "/img/", "blob.vercel", "cdn.", "cloudinary", "akamaized", "cloudfront"]


def is_image_url(url: str) -> bool:
    """True si la URL parece ser de una imagen y no de una página de producto."""
    url_lower = url.lower().split("?")[0]  # ignorar query params
    if url_lower.endswith(IMAGE_EXTENSIONS):
        return True
    if any(hint in url_lower for hint in IMAGE_CDN_HINTS):
        return True
    return False


def find_url_in_row(row: dict) -> str:
    """Busca la URL de producto en cualquier columna del CSV"""
    # Primero intenta la columna estándar
    for col in ["url", "URL", "link", "Link", "producto_url"]:
        val = (row.get(col) or "").strip()
        if val.startswith("http") and not is_image_url(val):
            return val

    # cualquier columna con URL que no sea imagen
    for val in row.values():
        if not isinstance(val, str):
            continue
        val = val.strip()
        if (val.startswith("http://") or val.startswith("https://")) and not is_image_url(val):
            return val

    return ""


def find_name_in_row(row: dict) -> tuple[str, bool]:
    """Retorna (nombre, corrupto). corrupto=True si el campo nombre tiene una URL."""
    for col in ["name", "nombre", "Name", "title", "titulo"]:
        val = (row.get(col) or "").strip()
        if not val:
            continue
        # Si el nombre es una URL, fila corrompida
        if val.startswith("http") or val.startswith(",http"):
            return val, True
        return val, False
    return "", False


def fix_csv_quotes(content: str) -> str:
    """Findclo exporta comillas internas sin escapar en CSV, esto las dobla para que el parser no se rompa."""
    out = []
    in_field = False
    i = 0
    n = len(content)
    while i < n:
        c = content[i]
        if not in_field:
            out.append(c)
            if c == '"':
                in_field = True
        else:
            if c == '"':
                nxt = content[i + 1] if i + 1 < n else ''
                if nxt == '"':
                    out.append('""')
                    i += 2
                    continue
                elif nxt in (',', '\n', '\r', ''):
                    out.append('"')
                    in_field = False
                else:
                    out.append('""')
            else:
                out.append(c)
        i += 1
    return ''.join(out)


def load_all_brands() -> dict[str, list[dict]]:
    csv_files = sorted(glob.glob(str(INPUT_DIR / "*.csv")))

    if not csv_files:
        print(f"No hay CSVs en {INPUT_DIR}")
        return {}

    brands: dict[str, list] = defaultdict(list)
    seen: dict[str, set]    = defaultdict(set)

    for path in csv_files:
        brand = extract_brand(path)
        with open(path, encoding="utf-8-sig") as f:
            raw = f.read()
        fixed = fix_csv_quotes(raw)
        reader = csv.DictReader(io.StringIO(fixed))
        for row in reader:
            name, corrupted = find_name_in_row(row)
            url = find_url_in_row(row)

            if corrupted:
                brands[brand].append({
                    "name":        name[:60],
                    "url":         url or "",
                    "description": "",
                    "status":      "ERROR",
                    "reason":      "Fila corrompida en el CSV: el titulo contiene una URL",
                })
                continue

            if not url:
                brands[brand].append({
                    "name":        name,
                    "url":         "",
                    "description": "",
                    "status":      "SIN_URL",
                    "reason":      "Sin URL de origen en Findclo",
                })
                continue

            row_key = (
                name,
                url,
                (row.get("images") or row.get("imagenes") or "").strip(),
            )
            if row_key in seen[brand]:
                continue
            seen[brand].add(row_key)
            brands[brand].append({
                "name":        name,
                "url":         url,
                "description": (row.get("description") or row.get("descripcion") or "").strip(),
            })

    for brand, products in brands.items():
        print(f"  {brand}: {len(products)} productos")

    return dict(brands)


# Check rápido HTTP (sin browser)

def quick_http_check(url: str) -> str | None:
    """HEAD request rápido. Retorna 'PRODUCTO_ELIMINADO' si 404, None si necesita browser."""
    try:
        r = requests.head(
            url,
            timeout=HTTP_TIMEOUT,
            allow_redirects=True,
            headers={"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) Chrome/120.0.0.0"},
        )
        if r.status_code == 404:
            return "PRODUCTO_ELIMINADO"
        # Detectar redirect a home
        final = r.url
        orig_path  = urlparse(url).path
        final_path = urlparse(final).path
        if urlparse(url).netloc == urlparse(final).netloc and len(final_path) < 5 and len(orig_path) > 10:
            return "PRODUCTO_ELIMINADO"
    except Exception:
        pass
    return None  # necesita browser


# Scraper con Playwright

async def wait_for_content(page, platform: str):
    selectors = PLATFORM_BUY_SELECTORS.get(platform, [])
    for selector in selectors:
        try:
            await page.wait_for_selector(selector, timeout=4000)
            return
        except Exception:
            continue
    try:
        await page.wait_for_load_state("networkidle", timeout=3000)
    except Exception:
        await page.wait_for_timeout(1500)


async def get_unavailable_sizes(page, platform: str) -> list[str]:
    selectors = UNAVAILABLE_SIZE_SELECTORS.get(platform, []) + UNAVAILABLE_SIZE_SELECTORS["generic"]
    sizes = []
    seen_texts = set()
    for selector in selectors:
        try:
            elements = await page.query_selector_all(selector)
            for el in elements:
                text = (await el.inner_text()).strip()
                if text and len(text) < 20 and text not in seen_texts:
                    seen_texts.add(text)
                    sizes.append(text)
            if sizes:
                break
        except Exception:
            continue
    return sizes


async def get_total_sizes(page, platform: str) -> int:
    selectors = TOTAL_SIZE_SELECTORS.get(platform, []) + TOTAL_SIZE_SELECTORS["generic"]
    for selector in selectors:
        try:
            elements = await page.query_selector_all(selector)
            if len(elements) >= 2:
                return len(elements)
        except Exception:
            continue
    return 0


def has_contextual_signal(text_lower: str, signal: str, context_words: list[str], window: int = 200) -> bool:
    """Busca 'signal' en el texto solo si aparece cerca de alguna de las context_words."""
    idx = 0
    while True:
        pos = text_lower.find(signal, idx)
        if pos == -1:
            break
        # Extraer ventana de texto alrededor de la señal
        snippet = text_lower[max(0, pos - window): pos + len(signal) + window]
        if any(cw in snippet for cw in context_words):
            return True
        idx = pos + 1
    return False

CONTEXT_STOCK_WORDS = ["stock", "talle", "talla", "color", "disponib", "carrito", "comprar", "agregar"]

AMBIGUOUS_OUT_SIGNALS = {
    "notificarme": CONTEXT_STOCK_WORDS,
    "no disponible": CONTEXT_STOCK_WORDS,
}


def check_stock_from_text(page_text: str, platform: str = "generic") -> tuple[str, str, str]:
    text_lower = page_text.lower()

    # Señales genéricas de sin stock
    found_out = [s for s in OUT_OF_STOCK_SIGNALS if s in text_lower]

    # Señales específicas de la plataforma
    for signal in PLATFORM_OUT_SIGNALS.get(platform, []):
        if signal in text_lower and signal not in found_out:
            found_out.append(signal)

    found_in  = [s for s in IN_STOCK_SIGNALS if s in text_lower]

    for signal, context in AMBIGUOUS_OUT_SIGNALS.items():
        if signal not in found_out:
            if has_contextual_signal(text_lower, signal, context):
                found_out.append(signal)

    if found_out and found_in:
        return "REVISAR", f'Algunos talles sin stock (señal: "{found_out[0]}")', found_out[0]
    if found_out and not found_in:
        return "SIN_STOCK", f'Señal: "{found_out[0]}"', found_out[0]
    if found_in:
        return "OK", "", ""
    return "OK", "Sin señal clara - revisar manualmente", ""


async def check_product_browser(page, product: dict) -> dict:
    url      = product["url"]
    result   = {**product, "status": "ERROR", "reason": ""}
    platform = detect_platform_from_url(url)

    # Validar que sea una URL real antes de intentar navegar
    if not url.startswith("http://") and not url.startswith("https://"):
        result["status"] = "ERROR"
        result["reason"] = f"URL invalida en el CSV: {url[:60]}"
        return result

    try:
        # Para sitios lentos o con rate limiting (ej: Penguin/jQuery), agregar delay
        if platform == "jquery":
            await page.wait_for_timeout(1500)

        # Intentar navegar, con un retry si hay timeout
        nav_timeout = 35000 if platform == "jquery" else BROWSER_TIMEOUT
        response = None
        for attempt in range(2):
            try:
                response = await page.goto(url, wait_until="domcontentloaded", timeout=nav_timeout)
                break
            except Exception as e:
                if attempt == 0 and "timeout" in str(e).lower():
                    await page.wait_for_timeout(4000)  # esperar más y reintentar
                    continue
                raise

        if response is None:
            result["status"] = "ERROR"
            result["reason"] = "Sin respuesta"
            return result

        if response.status == 404:
            result["status"] = "PRODUCTO_ELIMINADO"
            result["reason"] = "HTTP 404"
            return result

        final_url  = page.url
        orig_path  = urlparse(url).path
        final_path = urlparse(final_url).path
        if urlparse(url).netloc == urlparse(final_url).netloc and len(final_path) < 5 and len(orig_path) > 10:
            result["status"] = "PRODUCTO_ELIMINADO"
            result["reason"] = f"Redirigió a: {final_url}"
            return result

        await wait_for_content(page, platform)

        # Para jQuery (Penguin, Wanama, etc.): esperar el botón submit o señal de sin stock
        # Importante: innerText NO captura input values, por eso usamos querySelector
        if platform == "jquery":
            try:
                await page.wait_for_function(
                    "document.querySelector('input[type=submit][value]') !== null || "
                    "document.querySelector('button[type=submit]') !== null || "
                    "document.body.innerText.toLowerCase().includes('sin stock') || "
                    "document.body.innerText.toLowerCase().includes('sold out') || "
                    "document.body.innerText.toLowerCase().includes('agotado') || "
                    "document.body.innerText.toLowerCase().includes('talle agotado')",
                    timeout=10000
                )
            except Exception:
                pass  # si no aparece, igual seguimos

        # Para Tienda Nube (Midway): esperar que aparezca el botón de compra (input submit)
        if platform == "tiendanube":
            try:
                await page.wait_for_function(
                    "document.querySelector('input[data-store*=\"product-buy-button\"]') !== null || "
                    "document.querySelector('input.js-prod-submit-form') !== null || "
                    "document.querySelector('input.js-addtocart') !== null || "
                    "document.body.innerText.toLowerCase().includes('agregar al carrito') || "
                    "document.body.innerText.toLowerCase().includes('sin stock')",
                    timeout=8000
                )
            except Exception:
                pass  # si no aparece, igual seguimos

        # Para Angular (Awada, Paula Cahen): esperar que el JS renderice
        if platform == "angular":
            try:
                await page.wait_for_function(
                    "document.body.innerText.toLowerCase().includes('agregar al carrito') || "
                    "document.body.innerText.toLowerCase().includes('comprar') || "
                    "document.body.innerText.toLowerCase().includes('sin stock') || "
                    "document.body.innerText.toLowerCase().includes('sold out') || "
                    "document.body.innerText.toLowerCase().includes('agotado') || "
                    "document.body.innerText.toLowerCase().includes('no disponible')",
                    timeout=10000
                )
            except Exception:
                pass  # si no aparece, igual seguimos

        # Para VTEX: esperar explícitamente que aparezca el botón de compra en el texto
        # Cubre: Ricky Sarkany, Ginebra, Levis, Nike, La Martina, Tommy Hilfiger,
        #        Calvin Klein, Portsaid, Cuesta Blanca, Desiderata, Ayres
        if platform == "vtex":
            try:
                await page.wait_for_function(
                    "document.body.innerText.toLowerCase().includes('agregar al carrito') || "
                    "document.body.innerText.toLowerCase().includes('agregar a mi compra') || "
                    "document.body.innerText.toLowerCase().includes('agregar a la bolsa') || "
                    "document.body.innerText.toLowerCase().includes('añadir a la bolsa') || "
                    "document.body.innerText.toLowerCase().includes('agregar al changuito') || "
                    "document.body.innerText.toLowerCase().includes('agregar a mi bolsa') || "
                    "document.body.innerText.toLowerCase().includes('notificarme cuando haya stock') || "
                    "document.body.innerText.toLowerCase().includes('notificarme cuando este disponible') || "
                    "document.body.innerText.toLowerCase().includes('sin stock') || "
                    "document.body.innerText.toLowerCase().includes('sold out') || "
                    "document.body.innerText.toLowerCase().includes('agotado') || "
                    "document.body.innerText.toLowerCase().includes('no disponible') || "
                    "document.body.innerText.toLowerCase().includes('out of stock')",
                    timeout=12000
                )
            except Exception:
                pass  # si no aparece, igual seguimos

        page_text = await page.inner_text("body")

        # Detección por plataforma
        if platform == "tiendanube":
            # Tienda Nube: el botón principal es <input type="submit" value="Agregar al carrito">
            # cuando hay stock, o value="Sin stock" disabled cuando no hay.
            # Leemos SOLO el botón del producto (no el placeholder oculto).
            # Tienda Nube: leer el botón submit directamente con Playwright (más confiable que evaluate)
            tn_selectors = [
                'input[data-store*="product-buy-button"]',
                'input.js-prod-submit-form',
                'input.js-addtocart[type=submit]',
            ]
            btn_el = None
            for _sel in tn_selectors:
                try:
                    btn_el = await page.query_selector(_sel)
                    if btn_el:
                        break
                except Exception:
                    continue

            if btn_el:
                try:
                    btn_val     = ((await btn_el.get_attribute("value")) or "").lower()
                    btn_disabled = (await btn_el.get_attribute("disabled")) is not None
                    btn_classes  = (await btn_el.get_attribute("class")) or ""
                    btn_nostock  = "nostock" in btn_classes

                    if btn_disabled or btn_nostock or "sin stock" in btn_val:
                        status, reason, main_out_signal = "SIN_STOCK", f'Botón: "{btn_val}"', "sin stock"
                    else:
                        status, reason, main_out_signal = "OK", "", ""
                except Exception:
                    status, reason, main_out_signal = check_stock_from_text(page_text, platform)
            else:
                status, reason, main_out_signal = check_stock_from_text(page_text, platform)

        elif platform == "jquery":
            # jQuery (Penguin, Awada): el botón es <input type="submit" value="Agregar al carrito">
            # inner_text no captura valores de input, leer solo el botón del producto
            try:
                btn_info = await page.evaluate(
                    "(function() {"
                    "  var btn = document.querySelector('input.item-show-btn')"
                    "         || document.querySelector('input[type=submit].js-prod-submit-form');"
                    "  if (!btn) return null;"
                    "  return { value: btn.value.toLowerCase(), disabled: btn.disabled };"
                    "})()"
                )
                if btn_info:
                    val = btn_info["value"]
                    if btn_info["disabled"] or any(s in val for s in ["sin stock", "sold out", "agotado"]):
                        status, reason, main_out_signal = "SIN_STOCK", f'Botón: "{val}"', "sin stock"
                    elif any(s in val for s in ["agregar", "añadir", "comprar", "add to cart"]):
                        status, reason, main_out_signal = "OK", "", ""
                    else:
                        status, reason, main_out_signal = check_stock_from_text(page_text, platform)
                else:
                    status, reason, main_out_signal = check_stock_from_text(page_text, platform)
            except Exception:
                status, reason, main_out_signal = check_stock_from_text(page_text, platform)

        else:
            # Resto de plataformas: usar texto e input values como fallback
            try:
                input_values = await page.evaluate(
                    "Array.from(document.querySelectorAll('input[type=submit], button[type=submit]'))"
                    ".map(e => e.value || e.innerText || '').join(' ')"
                )
                page_text = page_text + " " + input_values
            except Exception:
                pass
            status, reason, main_out_signal = check_stock_from_text(page_text, platform)

        if status == "REVISAR":
            unavailable = await get_unavailable_sizes(page, platform)
            total       = await get_total_sizes(page, platform)

            if unavailable and total > 0:
                unavailable_count = len(unavailable)
                available_count   = total - unavailable_count
                sizes_str         = ", ".join(unavailable[:8])
                if available_count > unavailable_count:
                    status = "OK"
                    reason = ""
                else:
                    status = "REVISAR"
                    reason = f"{unavailable_count} de {total} talles sin stock: {sizes_str}"
            else:
                out_count = page_text.lower().count(main_out_signal) if main_out_signal else 0
                sizes_str = ", ".join(unavailable[:8]) if unavailable else ""
                detail    = f": {sizes_str}" if sizes_str else ""
                if out_count <= 2:
                    status = "OK"
                    reason = ""
                else:
                    status = "REVISAR"
                    reason = f"{out_count} talle(s) sin stock{detail}"

        elif status == "SIN_STOCK":
            unavailable = await get_unavailable_sizes(page, platform)
            if unavailable:
                reason = f"Sin stock. Talles afectados: {', '.join(unavailable[:8])}"

        result["status"] = status
        result["reason"] = reason

    except Exception as e:
        result["status"] = "ERROR"
        result["reason"] = str(e)[:120]
        pass

    return result


async def run_checker(brand: str, products: list[dict], cache: dict) -> list[dict]:
    total    = len(products)
    results  = [None] * total
    now_iso  = datetime.now().isoformat()

    # Separar los que ya están en caché válido y los que hay que procesar
    to_process = []   
    cached_hit = 0

    # separar caché con pendientes
    needs_http = []  
    for idx, product in enumerate(products):
        if product.get("status") == "SIN_URL":
            results[idx] = product
            continue
        url   = product["url"]
        entry = cache.get(url)
        if entry and is_cache_valid(entry):
            results[idx] = {**product, "status": entry["status"], "reason": entry["reason"]}
            cached_hit += 1
        else:
            needs_http.append((idx, product))

    # HTTP quick-checks en paralelo
    def _http_check_item(item):
        idx, product = item
        return idx, product, quick_http_check(product["url"])

    if needs_http:
        with concurrent.futures.ThreadPoolExecutor(max_workers=20) as pool:
            http_results = list(pool.map(_http_check_item, needs_http))

        for idx, product, quick in http_results:
            url = product["url"]
            if quick == "PRODUCTO_ELIMINADO":
                r = {**product, "status": "PRODUCTO_ELIMINADO", "reason": "HTTP 404"}
                results[idx] = r
                cache[url] = {"status": r["status"], "reason": r["reason"], "checked_at": now_iso}
            else:
                to_process.append((idx, product))

    if cached_hit:
        print(f"    {cached_hit} del cache | {len(to_process)} a escanear con browser")

    if not to_process:
        return results

    # Procesar con browser los que no están en caché
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
        # Bloquear imágenes, fuentes y media porque no los necesitamos para detectar stock
        # Esto reduce el tiempo de carga de cada página entre 30% y 60%
        async def _block_resources(route, request):
            if request.resource_type in ("image", "media", "font", "other"):
                await route.abort()
            else:
                await route.continue_()
        await context.route("**/*", _block_resources)

        # Reducir concurrencia para sitios con rate limiting (jQuery/Penguin)
        first_platform = detect_platform_from_url(to_process[0][1]["url"]) if to_process else "generic"
        effective_concurrency = 2 if first_platform == "jquery" else CONCURRENCY
        semaphore = asyncio.Semaphore(effective_concurrency)

        async def process(idx, product, pos):
            async with semaphore:
                # Intentar hasta 2 veces si hay ERROR 
                for attempt in range(2):
                    page = await context.new_page()
                    try:
                        result = await check_product_browser(page, product)
                        # Si dio ERROR en el primer intento, reintentar
                        if result["status"] == "ERROR" and attempt == 0:
                            await page.close()
                            await asyncio.sleep(2)
                            continue
                        break
                    except Exception:
                        if attempt == 0:
                            await page.close()
                            await asyncio.sleep(2)
                            continue
                        result = {**product, "status": "ERROR", "reason": "Fallo tras 2 intentos"}
                        break
                    finally:
                        try:
                            await page.close()
                        except Exception:
                            pass

                icons = {"OK": "[OK]", "REVISAR": "[~~]", "SIN_STOCK": "[!!]", "PRODUCTO_ELIMINADO": "[XX]", "ERROR": "[??]"}
                icon  = icons.get(result["status"], "[??]")
                reason_str = f" → {result['reason']}" if result.get("reason") and result["status"] in ("REVISAR", "SIN_STOCK") else ""
                print(f"    [{idx+1}/{total}] {icon} {result['status']:20} {result['name'][:45]}{reason_str}")
                results[idx] = result
                # Solo cachear si no fue un error transitorio
                if result["status"] != "ERROR":
                    cache[product["url"]] = {
                        "status":     result["status"],
                        "reason":     result["reason"],
                        "checked_at": now_iso,
                    }

        await asyncio.gather(*[process(idx, prod, pos) for pos, (idx, prod) in enumerate(to_process)])
        await browser.close()

    return results


# Reporte Excel

STATUS_COLORS = {
    "OK":                 "C6EFCE",
    "REVISAR":            "FFEB9C",
    "SIN STOCK":          "FFC7CE",
    "PRODUCTO ELIMINADO": "FFC7CE",
    "SIN URL":            "FCE4D6", 
    "ERROR":              "D9D9D9",
}

HEADERS = ["Nombre", "URL", "Estado", "Motivo"]


def add_brand_sheet(wb: openpyxl.Workbook, brand: str, results: list[dict]):
    ws = wb.create_sheet(title=brand[:31])
    ws.append(HEADERS)
    for cell in ws[1]:
        cell.font      = Font(bold=True, color="FFFFFF")
        cell.fill      = PatternFill(start_color="2F5496", end_color="2F5496", fill_type="solid")
        cell.alignment = Alignment(horizontal="center")

    for r in results:
        ws.append([r["name"], r["url"], r["status"], r["reason"]])
        color = STATUS_COLORS.get(r["status"], "FFFFFF")
        fill  = PatternFill(start_color=color, end_color=color, fill_type="solid")
        for cell in ws[ws.max_row]:
            cell.fill = fill

    ws.column_dimensions["A"].width = 45
    ws.column_dimensions["B"].width = 70
    ws.column_dimensions["C"].width = 22
    ws.column_dimensions["D"].width = 55


def add_summary_sheet(wb: openpyxl.Workbook, all_results: dict[str, list[dict]]):
    ws = wb.create_sheet(title="RESUMEN", index=0)
    ws["A1"] = f"Reporte Findclo — {datetime.now().strftime('%d/%m/%Y %H:%M')}"
    ws["A1"].font = Font(bold=True, size=13)
    ws.merge_cells("A1:G1")
    ws.append([])
    ws.append(["Marca", "Total", "OK", "Revisar", "Sin Stock", "Eliminados", "Sin URL", "Errores"])
    for cell in ws[3]:
        cell.font      = Font(bold=True, color="FFFFFF")
        cell.fill      = PatternFill(start_color="2F5496", end_color="2F5496", fill_type="solid")
        cell.alignment = Alignment(horizontal="center")

    total_global = Counter()
    for brand, results in sorted(all_results.items()):
        counts = Counter(r["status"] for r in results)
        total_global += counts
        ws.append([brand, len(results),
                   counts.get("OK", 0), counts.get("REVISAR", 0),
                   counts.get("SIN_STOCK", 0), counts.get("PRODUCTO_ELIMINADO", 0),
                   counts.get("SIN_URL", 0), counts.get("ERROR", 0)])

    ws.append([])
    ws.append(["TOTAL", sum(len(r) for r in all_results.values()),
               total_global.get("OK", 0), total_global.get("REVISAR", 0),
               total_global.get("SIN_STOCK", 0), total_global.get("PRODUCTO_ELIMINADO", 0),
               total_global.get("SIN_URL", 0), total_global.get("ERROR", 0)])
    for cell in ws[ws.max_row]:
        cell.font = Font(bold=True)

    ws.column_dimensions["A"].width = 25
    for col in ["B", "C", "D", "E", "F", "G"]:
        ws.column_dimensions[col].width = 14


def write_full_report(all_results: dict[str, list[dict]]) -> str:
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    add_summary_sheet(wb, all_results)
    for brand, results in sorted(all_results.items()):
        add_brand_sheet(wb, brand, results)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M")
    out_path  = OUTPUT_DIR / f"REPORTE_FINDCLO_{timestamp}.xlsx"
    wb.save(out_path)
    return str(out_path)


# Main

async def main():
    print("\n" + "="*60)
    print("  FINDCLO CHECKER")
    print("="*60)
    print(f"  Leyendo CSVs de: {INPUT_DIR}")
    print()

    brands = load_all_brands()
    if not brands:
        print("Pone los CSVs de Findclo en la carpeta 'input' y volvé a correr.")
        return

    cache = load_cache()
    total_cached = sum(1 for url, e in cache.items() if is_cache_valid(e))
    print(f"\n  Cache: {total_cached} URLs validas ({CACHE_DAYS} dias)")
    print(f"  Concurrencia: {CONCURRENCY} paginas en paralelo")
    print(f"\n{len(brands)} marca(s) encontradas. Iniciando escaneo...\n")

    all_results = {}

    for brand, products in sorted(brands.items()):
        print(f"  [{brand}] — {len(products)} productos")
        results = await run_checker(brand, products, cache)
        all_results[brand] = results
        counts = Counter(r["status"] for r in results)
        print(f"  -> OK: {counts['OK']} | Revisar: {counts['REVISAR']} | Sin stock: {counts['SIN_STOCK']} | Eliminados: {counts['PRODUCTO_ELIMINADO']} | Errores: {counts['ERROR']}\n")

    save_cache(cache)

    out_path = write_full_report(all_results)
    print("="*60)
    print(f"  Reporte generado: {out_path}")
    print("="*60)


if __name__ == "__main__":
    asyncio.run(main())
