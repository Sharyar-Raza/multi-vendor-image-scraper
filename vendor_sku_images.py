"""Download first exact-SKU vendor image as lowercase WebP.

Install: pip install requests beautifulsoup4 openpyxl pillow pytesseract seleniumbase
Also install Tesseract OCR and add it to PATH.
Run: python vendor_sku_images.py
"""
from __future__ import annotations

import csv
import atexit
import io
import os
import re
import shutil
import sys
import time
from copy import copy
from pathlib import Path
from types import SimpleNamespace
from urllib.parse import urlencode, urljoin, urlparse, urlsplit, urlunsplit, parse_qsl, quote

import openpyxl
from openpyxl.comments import Comment
import pytesseract
import requests
from bs4 import BeautifulSoup
from PIL import Image, ImageOps, UnidentifiedImageError

ROOT = Path(__file__).resolve().parent
INPUT = ROOT / "Images-task-master-excel.xlsx"
OUT = ROOT / "images_task_webp"
LOG = ROOT / "images_task_image_log.csv"
EXCEL_OUT = ROOT / "Images-task-master-excel-with-thumbnails.xlsx"
# This workbook starts fresh; no rows inherit the older workbook checkpoint.
LEGACY_COMPLETED_EXCEL_ROW = 1
VENDORS = [
    ("Router Switch", "https://www.router-switch.com"),
    ("Flagship Technologies", "https://store.flagshiptech.com"),
    ("Direct Macro", "https://directmacro.com"),
    ("Newtown Spares", "https://www.newtownspares.com"),
    ("Four Star Tech", "https://4startech.com"),
    ("PC Shore", "https://pcshore.com"),
    ("Server Tech Supply", "https://www.servertechsupply.com"),
    ("Server Orbit", "https://serverorbit.com"),
    ("Network Devices Inc", "https://networkdevicesinc.com"),
    ("CablesAndKits", "https://www.cablesandkits.com"),
    ("IT Market", "https://it-market.com"),
    ("Network Tigers", "https://www.networktigers.com"),
    ("Server Supply", "https://www.serversupply.com"),
    ("Netmode", "https://netmode.com"),
    ("Network Switch", "https://network-switch.com"),
    ("Top Paragon Resources", "https://topparagonresource.com"),
    ("Network Outlet", "https://networkoutlet.com"),
    ("Get It New", "https://getitnew.com"),
    ("xByte Technologies", "https://www.xbyte.com"),
]
SEARCH_URLS = {
    "Router Switch": "https://www.router-switch.com/catalogsearch/result/?q={sku}",
    "Flagship Technologies": "https://store.flagshiptech.com/search.php?search_query={sku}",
    "Direct Macro": "https://directmacro.com/catalogsearch/result/?q={sku}",
    "Newtown Spares": "https://www.newtownspares.com/search.php?search_query={sku}&section=product",
    "Four Star Tech": "https://4startech.com/search.php?search_query={sku}&section=content&page=1",
    "PC Shore": "https://pcshore.com/catalogsearch/result/?q={sku}",
    "Server Tech Supply": "https://www.servertechsupply.com/?s={sku}&et_search=true&post_type=product",
    "Server Orbit": "https://serverorbit.com/?match=all&subcats=Y&pcode_from_q=Y&pshort=Y&pfull=Y&pname=Y&pkeywords=Y&search_performed=Y&q={sku}&dispatch=products.search",
    "Network Devices Inc": "https://networkdevicesinc.com/search?q={sku}",
    "CablesAndKits": "https://www.cablesandkits.com/search?keywords={sku}",
    "IT Market": "https://it-market.com/en/search?search={sku}",
    "Network Tigers": "https://www.networktigers.com/search?q={sku}",
    "Server Supply": "https://www.serversupply.com/products/part_search/query_parts.asp?q={sku}",
    "Netmode": "https://netmode.com/search?type=product&q={sku}",
    "Network Switch": "https://network-switch.com/search?type=product&options%5Bunavailable_products%5D=last&options%5Bprefix%5D=last&q={sku}",
    "Top Paragon Resources": "https://topparagonresource.com/?s={sku}",
    "Network Outlet": "https://networkoutlet.com/search?type=product&q={sku}&options%5Bprefix%5D=last",
    "Get It New": "https://getitnew.com/search.php?search_query={sku}&section=content",
}
HEADERS = {"User-Agent": "Mozilla/5.0 (compatible; SKUImageResearch/1.0)"}
SESSION = requests.Session()
SESSION.headers.update(HEADERS)
DRIVER = None
CDP_SB = None
BROWSER_DOMAINS = set()


class BrowserSetupError(RuntimeError):
    """A local browser problem that requires action before retrying."""


CAPTCHA_MARKERS = (
    "just a moment", "verify you are human", "security check",
    "attention required", "checking your browser", "cf-chl-",
    "challenge-platform", "performing security verification",
)


def challenge_detected():
    title = (DRIVER.title or "").casefold().strip()
    if any(marker in title for marker in CAPTCHA_MARKERS):
        return True
    if "/cdn-cgi/challenge-platform/" in (DRIVER.current_url or "").casefold():
        return True
    # Hidden challenge scripts can appear on ordinary product pages.
    body = (DRIVER.execute_script("return document.body?.innerText || ''") or "").casefold().strip()
    if len(body) < 3000 and any(marker in body for marker in CAPTCHA_MARKERS):
        return True
    return bool(DRIVER.execute_script(
        "return !!document.querySelector('iframe[src*=\"challenges.cloudflare.com\"], "
        "#cf-challenge-running, #challenge-stage')"
    ))


def response_is_challenge(response):
    """Detect a challenge returned with HTTP 200 as well as an error status."""
    if "html" not in response.headers.get("Content-Type", "").casefold():
        return False
    soup = BeautifulSoup(response.text, "html.parser")
    title = soup.title.get_text(" ", strip=True).casefold() if soup.title else ""
    if any(marker in title for marker in CAPTCHA_MARKERS):
        return True
    body = soup.body.get_text(" ", strip=True).casefold() if soup.body else ""
    return len(body) < 3000 and any(marker in body for marker in CAPTCHA_MARKERS)


def wait_for_challenge(url, timeout=120):
    """Let the user complete a visible challenge in Chrome, then continue."""
    if not challenge_detected():
        return
    print(f"CAPTCHA at {url}. Solve it in the open Chrome window; waiting up to {timeout} seconds.", flush=True)
    deadline = time.monotonic() + timeout
    while time.monotonic() < deadline:
        time.sleep(2)
        if not challenge_detected():
            print("CAPTCHA cleared; resuming search.", flush=True)
            return
    raise RuntimeError(f"CAPTCHA was not completed within {timeout} seconds at {url}")


def verify_human_prompt_visible():
    """Only allow a click when the challenge actually shows its human check."""
    prompt = re.compile(r"verify you are (?:a )?human", re.IGNORECASE)
    # Cloudflare can render the wording inside an iframe that DOM queries cannot
    # read. OCR the visible page before allowing a GUI click.
    if CDP_SB is not None:
        capture = ROOT / ".cloudflare_prompt_check.png"
        try:
            CDP_SB.save_screenshot(capture.name, folder=str(ROOT))
            with Image.open(capture) as screenshot:
                enlarged = screenshot.resize((screenshot.width * 2, screenshot.height * 2))
                visible_text = pytesseract.image_to_string(enlarged)
            return bool(prompt.search(visible_text))
        except Exception as exc:
            print(f"Cannot confirm visible human checkbox: {exc}", flush=True)
        finally:
            capture.unlink(missing_ok=True)
    return False


def cdp_browser_get(url=None):
    """Use one visible CDP Chrome for navigation and challenge clicks."""
    global CDP_SB
    if CDP_SB is None:
        from seleniumbase import sb_cdp
        CDP_SB = sb_cdp.Chrome()
        atexit.register(lambda: CDP_SB.quit() if CDP_SB else None)
    if url:
        CDP_SB.get(url)
        CDP_SB.sleep(3)
    deadline = time.monotonic() + 90
    click_attempts = 0
    waiting_reported = False
    while time.monotonic() < deadline:
        title = (CDP_SB.get_title() or "").casefold()
        html = CDP_SB.get_page_source()
        body = BeautifulSoup(html, "html.parser").get_text(" ", strip=True).casefold()
        challenged = (any(marker in title for marker in CAPTCHA_MARKERS) or
                      (len(body) < 3000 and any(marker in body for marker in CAPTCHA_MARKERS)))
        if not challenged:
            return SimpleNamespace(text=html, url=CDP_SB.get_current_url(), content=b"")
        if verify_human_prompt_visible() and click_attempts < 2:
            click_attempts += 1
            print(f"Cloudflare challenge at {CDP_SB.get_current_url()}; clicking visible verification checkbox (attempt {click_attempts}).", flush=True)
            try:
                CDP_SB.gui_click_captcha()
                CDP_SB.sleep(8)
                if verify_human_prompt_visible():
                    CDP_SB.solve_captcha()
            except Exception as exc:
                print(f"Challenge control did not clear the page: {exc}", flush=True)
        elif not waiting_reported and click_attempts == 0:
            print(f"Cloudflare is still verifying at {CDP_SB.get_current_url()}; waiting for the human checkbox.", flush=True)
            waiting_reported = True
        CDP_SB.sleep(2)
    raise RuntimeError(f"Cloudflare challenge remained at {CDP_SB.get_current_url()}")


def browser_get(url, params=None):
    """Open every vendor in the same visible CDP Chrome session."""
    if params:
        parts = urlsplit(url)
        query = dict(parse_qsl(parts.query))
        query.update(params)
        url = urlunsplit((parts.scheme, parts.netloc, parts.path, urlencode(query), parts.fragment))
    return cdp_browser_get(url)

def browser_search(base, sku):
    """Submit a SKU in the same CDP Chrome window and handle any challenge."""
    browser_get(base)
    field = 'input#q, input[placeholder*="Search"], input[name*="search"], input[name="q"]'
    print(f"Search page opened at {base}; locating its search box...", flush=True)
    CDP_SB.type(field, sku, timeout=12)
    CDP_SB.press_keys(field, "\n", timeout=12)
    print(f"Search submitted at {base}; waiting for results...", flush=True)
    CDP_SB.sleep(3)
    response = cdp_browser_get()
    if sku.casefold() not in response.text.casefold():
        raise RuntimeError(f"Search results did not contain SKU {sku} at {base}")
    return response

def get(url, **kwargs):
    raw_image = kwargs.pop("raw_image", False)
    try:
        response = SESSION.get(url, timeout=20, **kwargs)
        if not raw_image and (response.status_code in (403, 429, 503) or response_is_challenge(response)):
            BROWSER_DOMAINS.add(urlparse(url).netloc)
            if kwargs.get("params"):
                return browser_search(f"{urlparse(url).scheme}://{urlparse(url).netloc}",
                                      next(iter(kwargs["params"].values())))
            return browser_get(url)
        response.raise_for_status()
        return response
    except requests.RequestException:
        if kwargs.get("params") is not None:
            BROWSER_DOMAINS.add(urlparse(url).netloc)
            return browser_search(f"{urlparse(url).scheme}://{urlparse(url).netloc}",
                                  next(iter(kwargs["params"].values())))
        raise


def search_form(base):
    soup = BeautifulSoup(get(base).text, "html.parser")
    for form in soup.select("form"):
        for field in form.select("input[name]"):
            name = field.get("name", "")
            hint = (field.get("placeholder", "") + " " + name).lower()
            if any(word in hint for word in ("search", "query", "keyword")):
                return urljoin(base, form.get("action") or base), name
    # JavaScript storefronts often have a visible search box outside a form.
    return base, "__browser__"


def exact_links(soup, page_url, sku, base):
    target = sku.casefold()
    target_slug = re.sub(r"[^a-z0-9]+", "-", target).strip("-")
    for a in soup.select("a[href]"):
        href = urljoin(page_url, a["href"])
        if urlparse(href).netloc.removeprefix("www.") != urlparse(base).netloc.removeprefix("www."):
            continue
        label = a.get_text(" ", strip=True)
        # Reject longer variants such as SKU-RF, SKU=, or SKU/2.
        # Product slugs replace separators such as / with hyphens. Match the
        # complete slug so a related variant cannot pass as the requested SKU.
        path_slug = urlparse(href).path.strip("/").casefold()
        if exact_sku_in_text(label, target) or path_slug == target_slug:
            if "search" not in urlparse(href).path.lower():
                yield href


def newtown_direct_product(base, sku):
    """Try Newtown's /sku/ product URL before its general search page."""
    slug = re.sub(r"[^a-z0-9]+", "-", sku.casefold()).strip("-")
    direct_url = f"{base}/{quote(slug, safe='')}/"
    try:
        response = get(direct_url)
    except requests.RequestException:
        return None
    product = BeautifulSoup(response.text, "html.parser")
    heading = product.find("h1")
    title = product.title.get_text(" ", strip=True) if product.title else ""
    identity = title + " " + (heading.get_text(" ", strip=True) if heading else "")
    return response.url if exact_sku_in_text(identity, sku) else None


def exact_sku_in_text(text, sku):
    return bool(re.search(
        r"(?<![a-z0-9])" + re.escape(sku.casefold()) + r"(?![a-z0-9/_=.-])",
        text.casefold(),
    ))


def image_urls(soup, page_url, sku=None, gallery_only=False):
    seen = set()
    if gallery_only:
        for tag in soup.select(
            ".product.media img, .gallery-placeholder img, .fotorama img, "
            "[data-gallery-role] img, img.fotorama__img"
        ):
            attrs = " ".join(str(tag.get(key, "")) for key in ("class", "id", "alt"))
            if any(word in attrs.casefold() for word in ("logo", "no-picture", "no picture")):
                continue
            for key in ("data-zoom-image", "data-large_image", "data-full",
                        "data-original", "data-src", "src"):
                value = tag.get(key)
                if value and not value.startswith("data:"):
                    url = urljoin(page_url, value)
                    if url not in seen:
                        seen.add(url)
                        yield url
        return
    for selector, key in (("meta[property='og:image']", "content"),
                          ("meta[name='twitter:image']", "content"),
                          ("img[itemprop='image']", "src"),
                          ("img.productView-img--main", "src"),
                          (".productView-image img", "src"),
                          (".product-image img", "src"),
                          (".product.media img", "src"),
                          (".fotorama__img", "src"),
                          ("img[data-gallery-role]", "src"),
                          (".product-gallery img", "src"),
                          (".woocommerce-product-gallery img", "src")):
        for tag in soup.select(selector):
            if tag.get(key):
                if tag.name == "img" and "logo" in " ".join(
                    str(tag.get(attr, "")) for attr in ("class", "id", "alt")
                ).casefold():
                    continue
                url = urljoin(page_url, tag[key])
                if url not in seen:
                    seen.add(url)
                    yield url
    for tag in soup.select("img"):
        attrs = " ".join(str(tag.get(key, "")) for key in ("class", "id", "alt"))
        if any(word in attrs.casefold() for word in ("no-picture", "no picture", "logo")):
            continue
        sku_hint = bool(sku and exact_sku_in_text(attrs, sku))
        if not sku_hint and not any(word in attrs.lower() for word in ("product", "gallery", "main-image", "fotorama")):
            continue
        for key in ("data-zoom-image", "data-large_image", "data-full",
                    "data-original", "data-src", "src"):
            value = tag.get(key)
            if value and not value.startswith("data:"):
                url = urljoin(page_url, value)
                if url not in seen:
                    seen.add(url)
                    yield url


def browser_image_bytes(url):
    """Capture a product image already visible in Chrome when its URL blocks requests."""
    if CDP_SB is not None:
        wanted = url.split("?", 1)[0]
        for element in CDP_SB.select_all("img"):
            source = urljoin(CDP_SB.get_current_url(), element.get_attribute("src") or "")
            if source.split("?", 1)[0] == wanted:
                temp = ROOT / ".cdp_image_capture.png"
                try:
                    element.save_screenshot(str(temp))
                    return temp.read_bytes()
                finally:
                    temp.unlink(missing_ok=True)
        return None
    if DRIVER is None:
        return None
    from selenium.webdriver.common.by import By

    wanted = url.split("?", 1)[0]
    for element in DRIVER.find_elements(By.TAG_NAME, "img"):
        source = element.get_attribute("currentSrc") or element.get_attribute("src") or ""
        if source.split("?", 1)[0] == wanted and element.is_displayed():
            return element.screenshot_as_png
    return None


def browser_product_images(product_url, sku, gallery_only=False):
    """Yield displayed product-image captures when image URLs cannot be downloaded."""
    if CDP_SB is not None:
        # URL candidates from the CDP page are captured by browser_image_bytes.
        return
    if DRIVER is None:
        return
    from selenium.webdriver.common.by import By

    current = urlparse(DRIVER.current_url)
    product = urlparse(product_url)
    if current.netloc.removeprefix("www.") != product.netloc.removeprefix("www.") or current.path != product.path:
        return
    sku_slug = re.sub(r"[^a-z0-9]+", "-", sku.casefold()).strip("-")
    for element in DRIVER.find_elements(By.TAG_NAME, "img"):
        try:
            if not element.is_displayed() or min(element.size["width"], element.size["height"]) < 200:
                continue
            source = element.get_attribute("currentSrc") or element.get_attribute("src") or ""
            attrs = " ".join((element.get_attribute("class") or "", element.get_attribute("alt") or ""))
            if "logo" in attrs.casefold():
                continue
            if gallery_only and not DRIVER.execute_script(
                "return !!arguments[0].closest('.product.media, .gallery-placeholder, .fotorama, [data-gallery-role]')",
                element,
            ):
                continue
            if not (exact_sku_in_text(attrs, sku) or sku_slug in source.casefold() or
                    any(word in attrs.casefold() for word in ("product", "gallery", "fotorama"))):
                continue
            if source:
                yield source, element.screenshot_as_png
        except Exception:
            continue


def clean_image_bytes(data, url):
    if any(x in url.lower() for x in ("watermark", "placeholder", "no-image", "no-picture", "logo")):
        return None
    if not data or len(data) < 1000:
        return None
    try:
        image = ImageOps.exif_transpose(Image.open(io.BytesIO(data)))
        if image.mode in ("RGBA", "LA") or "transparency" in image.info:
            rgba = image.convert("RGBA")
            canvas = Image.new("RGBA", rgba.size, "white")
            canvas.alpha_composite(rgba)
            image = canvas.convert("RGB")
        else:
            image = image.convert("RGB")
    except UnidentifiedImageError:
        return None
    if min(image.size) < 200:
        return None
    # Require an actual white studio background. Check the outside band so a
    # white letterbox around a dark vendor photograph cannot pass.
    sample = image.copy()
    sample.thumbnail((300, 300))
    width, height = sample.size
    band = max(3, min(width, height) // 12)
    edge = [sample.getpixel((x, y)) for y in range(height) for x in range(width)
            if x < band or x >= width - band or y < band or y >= height - band]
    if sum(min(pixel) >= 235 and max(pixel) - min(pixel) <= 15 for pixel in edge) / len(edge) < 0.94:
        return None
    white_share = sum(min(pixel) >= 235 and max(pixel) - min(pixel) <= 15
                      for pixel in sample.getdata()) / (width * height)
    if white_share < 0.55:
        return None
    # Vendor logos and delivery badges often sit in a corner while the rest
    # of the product photo has a white background.
    corner_width, corner_height = max(8, width // 6), max(8, height // 6)
    for left in (0, width - corner_width):
        for top in (0, height - corner_height):
            corner = sample.crop((left, top, left + corner_width, top + corner_height))
            white = sum(min(pixel) >= 235 and max(pixel) - min(pixel) <= 15
                        for pixel in corner.getdata())
            if white / (corner_width * corner_height) < 0.97:
                return None

    # Search enlarged images too: small corner logos and pale watermarks are
    # frequently missed at the original resolution.
    ocr_image = image.copy()
    if min(ocr_image.size) < 700:
        scale = 700 / min(ocr_image.size)
        ocr_image = ocr_image.resize((round(image.width * scale), round(image.height * scale)))
    words = pytesseract.image_to_string(ocr_image).casefold()
    if any(x in words for x in (
        "watermark", "sample", "preview", "copyright", ".com", "www.",
        "searching the picture", "depths of cyberspace", "no picture",
        "no image available", "image not available", "image unavailable",
    )):
        return None
    # OCR often invents words from ports and circuit traces. Printed model
    # labels on the product are acceptable; reject added writing in the white
    # space above or below the product.
    ocr = pytesseract.image_to_data(
        ocr_image, config="--psm 11", output_type=pytesseract.Output.DICT)
    for word, confidence, top, height in zip(
            ocr["text"], ocr["conf"], ocr["top"], ocr["height"]):
        in_margin = top + height < ocr_image.height * 0.20 or top > ocr_image.height * 0.85
        if (in_margin and len(re.sub(r"[^a-z0-9]", "", word.casefold())) >= 4
                and float(confidence) >= 75):
            return None
    # A brand mark on a mostly white canvas is not a product photo.
    visible_words = re.findall(r"[a-z0-9]+", words)
    if "cisco" in visible_words and len(visible_words) <= 3:
        sample = image.copy()
        sample.thumbnail((100, 100))
        white_pixels = sum(1 for red, green, blue in sample.getdata()
                           if min(red, green, blue) >= 235)
        if white_pixels / (sample.width * sample.height) >= 0.80:
            return None
    return image


def clean_image(url):
    try:
        data = get(url, raw_image=True).content
    except requests.RequestException:
        data = browser_image_bytes(url)
        if data is None:
            raise
    return clean_image_bytes(data, url)


def saved_records():
    return {sku: row for sku, row in latest_log_rows().items()
            if row.get("status") == "saved" and row.get("file")
            and (OUT / row["file"]).is_file()}


def latest_log_rows():
    if not LOG.exists():
        return {}
    with LOG.open(newline="", encoding="utf-8-sig") as file:
        return {row["sku"]: row for row in csv.DictReader(file) if row.get("sku")}


def queue_rejected_images():
    """Mark manually quarantined images for retry in the active workbook."""
    rejected = ROOT / "rejected_images"
    if not rejected.is_dir() or not LOG.exists():
        return
    with LOG.open(newline="", encoding="utf-8-sig") as file:
        rows = list(csv.DictReader(file))
    latest = {row["sku"]: row for row in rows if row.get("sku")}
    to_queue = []
    for path in rejected.iterdir():
        if not path.is_file():
            continue
        for row in reversed(rows):
            if row.get("file", "").casefold() != path.name.casefold():
                continue
            sku = row["sku"]
            current = latest[sku]
            if current.get("status") == "saved" and not (OUT / current.get("file", "")).is_file():
                if sku not in to_queue:
                    to_queue.append(sku)
            break
    if to_queue:
        with LOG.open("a", newline="", encoding="utf-8-sig") as file:
            log = csv.writer(file)
            for sku in to_queue:
                log.writerow((sku, "pending", "", "", "", ""))
            file.flush()
            os.fsync(file.fileno())
        print(f"Queued {len(to_queue)} rejected image SKU(s) as Pending.")


def audit_saved_images():
    """Quarantine earlier images that fail the current visual requirements."""
    records = saved_records()
    if not records:
        return
    rejected = ROOT / "rejected_images"
    rejected_count = 0
    with LOG.open("a", newline="", encoding="utf-8-sig") as file:
        log = csv.writer(file)
        for sku, record in records.items():
            path = OUT / record["file"]
            if clean_image_bytes(path.read_bytes(), record.get("image_url", "")) is not None:
                continue
            rejected.mkdir(exist_ok=True)
            os.replace(path, rejected / path.name)
            log.writerow((sku, "pending", "", "", "", ""))
            rejected_count += 1
        file.flush()
        os.fsync(file.fileno())
    if rejected_count:
        print(f"Moved {rejected_count} unsuitable saved image(s) to {rejected.name}; queued their SKUs for retry.")


def not_found_skus():
    return {sku for sku, row in latest_log_rows().items()
            if row.get("status") in ("not_found", "checked_no_image", "not_found_after_retry")}


def incomplete_not_found_skus():
    return {sku for sku, row in latest_log_rows().items()
            if row.get("status") == "not_found_after_retry"}


def pending_skus():
    return {sku for sku, row in latest_log_rows().items()
            if row.get("status") == "pending"}


def reached_skus(skus):
    """Recover the furthest worksheet position reached by sequential prior runs."""
    if not LOG.exists():
        return set()
    positions = {sku: index for index, sku in enumerate(skus)}
    with LOG.open(newline="", encoding="utf-8-sig") as file:
        furthest = max((positions[row["sku"]] for row in csv.DictReader(file)
                        if row.get("sku") in positions), default=-1)
    return set(skus[:furthest + 1])


def restore_legacy_checkpoint(skus, records, not_found):
    # Older script versions logged successful images but left unsuccessful rows
    # blank. The user confirmed rows through Excel row 18 were completed.
    earlier_skus = skus[:LEGACY_COMPLETED_EXCEL_ROW - 1]
    missing = [sku for sku in earlier_skus if sku not in records and sku not in not_found]
    if not missing:
        return
    with LOG.open("a", newline="", encoding="utf-8-sig") as file:
        log = csv.writer(file)
        for sku in missing:
            log.writerow((sku, "not_found", "", "", "", ""))
        file.flush()
        os.fsync(file.fileno())
    not_found.update(missing)
    print(f"Restored {len(missing)} earlier completed SKU(s) as Not Found.")


def print_progress(total, records, not_found, reached):
    completed = len(set(records) | not_found)
    pending = len(reached - set(records) - not_found)
    line = (f"PROGRESS: total={total} | completed={completed} | "
            f"remaining={total - completed} | images found={len(records)} | "
            f"not found={len(not_found)} | pending={pending}")
    if sys.stdout.isatty():
        line = f"\033[96m{line}\033[0m"
    print(line, flush=True)


def update_excel(records, not_found, incomplete_not_found=None):
    incomplete_not_found = incomplete_not_found or set()
    pending = pending_skus()
    workbook = openpyxl.load_workbook(INPUT)
    sheet = workbook.active
    image_col = sheet.max_column + 1
    source_col = image_col + 1
    header_style = copy(sheet.cell(1, sheet.max_column)._style)
    sheet.cell(1, image_col, "Thumbnail Image")
    sheet.cell(1, source_col, "Thumbnail Source")
    sheet.cell(1, image_col)._style = copy(header_style)
    sheet.cell(1, source_col)._style = copy(header_style)
    sheet.column_dimensions[openpyxl.utils.get_column_letter(image_col)].width = 32
    sheet.column_dimensions[openpyxl.utils.get_column_letter(source_col)].width = 70
    for row_number in range(2, sheet.max_row + 1):
        sku = str(sheet.cell(row_number, 1).value or "").strip()
        record = records.get(sku)
        if record:
            sheet.cell(row_number, image_col, Path(record["file"]).name.lower())
            cell = sheet.cell(row_number, source_col, record["product_url"])
            cell.hyperlink = record["product_url"]
            cell.style = "Hyperlink"
            if record.get("vendor") == "Direct Macro":
                cell.comment = Comment(
                    "Direct Macro warns that the product image may not exactly match the product. "
                    "The SKU on the linked product page was checked; review the image visually.",
                    "Image research",
                )
        elif sku in not_found:
            cell = sheet.cell(row_number, image_col, "Not Found")
            if sku in incomplete_not_found:
                cell.comment = Comment(
                    "No image was found after retry, but one or more vendor searches were blocked or failed. "
                    "This is not proof that no vendor has an image.",
                    "Image research",
                )
        elif sku in pending:
            sheet.cell(row_number, image_col, "Pending")
    temp = EXCEL_OUT.with_suffix(".tmp.xlsx")
    workbook.save(temp)
    try:
        os.replace(temp, EXCEL_OUT)
        return True
    except PermissionError:
        temp.unlink(missing_ok=True)
        print(f"Excel is open/locked: {EXCEL_OUT.name}. Close it; saved progress remains in the log.")
        return False


def main():
    # OCR is mandatory: fail before downloading if it is unavailable.
    candidates = [
        shutil.which("tesseract"),
        r"C:\Program Files\Tesseract-OCR\tesseract.exe",
        r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
        os.path.join(os.environ.get("LOCALAPPDATA", ""), "Programs", "Tesseract-OCR", "tesseract.exe"),
    ]
    executable = next((path for path in candidates if path and Path(path).is_file()), None)
    if executable is None:
        raise SystemExit(
            "Tesseract OCR program missing. Install it with: "
            "winget install --id UB-Mannheim.TesseractOCR --exact "
            "--accept-package-agreements --accept-source-agreements. "
            "Then reopen VS Code and run this script again."
        )
    pytesseract.pytesseract.tesseract_cmd = executable
    pytesseract.get_tesseract_version()
    OUT.mkdir(exist_ok=True)
    queue_rejected_images()
    audit_saved_images()
    records = saved_records()
    not_found = not_found_skus()
    incomplete_not_found = incomplete_not_found_skus()
    workbook = openpyxl.load_workbook(INPUT, read_only=True, data_only=True)
    skus = [str(row[0]).strip() for row in list(workbook.active.values)[1:] if row[0]]
    restore_legacy_checkpoint(skus, records, not_found)
    update_excel(records, not_found, incomplete_not_found)
    total = len(skus)
    done = len(set(records) | not_found)
    reached = reached_skus(skus)
    pending = len(reached - set(records) - not_found)
    print_progress(total, records, not_found, reached)
    forms = {}
    for name, base in VENDORS:
        if name in SEARCH_URLS:
            continue
        if name == "xByte Technologies":
            forms[name] = (base, "__browser__")
            continue
        try:
            forms[name] = search_form(base)
        except BrowserSetupError as exc:
            raise SystemExit(str(exc)) from None
        except Exception as exc:
            forms[name] = None
            print(f"{name}: unavailable ({exc})")
    new_log = not LOG.exists()
    with LOG.open("a", newline="", encoding="utf-8-sig") as file:
        log = csv.writer(file)
        scanned_this_run = 0
        if new_log:
            log.writerow(("sku", "status", "vendor", "product_url", "image_url", "file"))
        positions = {sku: number for number, sku in enumerate(skus, 1)}
        finished = set(records) | not_found
        pending_first = [sku for sku in skus if sku in reached and sku not in finished]
        pending_at_start = set(pending_first)
        new_after = [sku for sku in skus if sku not in reached]
        worklist = pending_first + new_after
        print(f"Run order: {len(pending_first)} pending SKUs first, "
              f"then {len(new_after)} new SKUs.", flush=True)
        for sku in worklist:
            number = positions[sku]
            filename = re.sub(r"[^a-z0-9._-]+", "_", sku.lower()) + ".webp"
            destination = OUT / filename
            scanned_this_run += 1
            if sku not in reached:
                log.writerow((sku, "started", "", "", "", ""))
                file.flush()
                os.fsync(file.fileno())
                reached.add(sku)
            found = False
            had_error = False
            for name, base in VENDORS:
                try:
                    direct_product = newtown_direct_product(base, sku) if name == "Newtown Spares" else None
                    if direct_product:
                        print(f"Checking {sku} on {name}: {direct_product}", flush=True)
                        product_urls = [direct_product]
                    elif name in SEARCH_URLS:
                        print(f"Searching {sku} on {name} in visible Chrome...", flush=True)
                        search_url = SEARCH_URLS[name].format(sku=quote(sku, safe=""))
                        response = browser_get(search_url)
                    else:
                        form = forms.get(name)
                        if not form:
                            had_error = True
                            continue
                        print(f"Searching {sku} on {name} in visible Chrome...", flush=True)
                        url, field = form
                        response = (browser_search(base, sku)
                                    if field == "__browser__"
                                    else browser_get(url, params={field: sku}))
                    if not direct_product:
                        results = BeautifulSoup(response.text, "html.parser")
                        product_urls = list(dict.fromkeys(exact_links(results, response.url, sku, base)))[:8]
                    for product_url in product_urls:
                        product = BeautifulSoup(get(product_url).text, "html.parser")
                        title = product.title.get_text(" ", strip=True) if product.title else ""
                        heading = product.find(["h1", "h2"])
                        identity = title + " " + (heading.get_text(" ", strip=True) if heading else "")
                        if not exact_sku_in_text(identity, sku):
                            continue
                        page_text = product.get_text(" ", strip=True).casefold()
                        if "image may not exactly match product" in page_text:
                            print(f"{sku} / {name}: vendor image warning; source will carry a review note")
                        gallery_only = name == "Router Switch"
                        url_candidates = [(url, None) for url in dict.fromkeys(
                            image_urls(product, product_url, sku, gallery_only=gallery_only))]
                        browser_candidates = list(browser_product_images(product_url, sku,
                                                                         gallery_only=gallery_only))
                        for image_url, captured in url_candidates + browser_candidates:
                            try:
                                image = (clean_image_bytes(captured, image_url) if captured is not None
                                         else clean_image(image_url))
                            except requests.RequestException:
                                continue
                            if image is None:
                                continue
                            temp_image = destination.with_suffix(".tmp.webp")
                            image.save(temp_image, "WEBP", quality=90)
                            os.replace(temp_image, destination)
                            record = {"sku": sku, "status": "saved", "vendor": name,
                                      "product_url": product_url, "image_url": image_url,
                                      "file": filename}
                            log.writerow(tuple(record.values()))
                            file.flush()
                            os.fsync(file.fileno())
                            records[sku] = record
                            update_excel(records, not_found, incomplete_not_found)
                            found = True
                            break
                        if found:
                            break
                except BrowserSetupError as exc:
                    raise SystemExit(str(exc)) from None
                except Exception as exc:
                    had_error = True
                    print(f"{sku} / {name}: {exc}")
                if found:
                    break
                time.sleep(0.5)
            if not found and (not had_error or sku in pending_at_start):
                status_code = "not_found_after_retry" if had_error else "not_found"
                log.writerow((sku, status_code, "", "", "", ""))
                file.flush()
                os.fsync(file.fileno())
                not_found.add(sku)
                if had_error:
                    incomplete_not_found.add(sku)
                update_excel(records, not_found, incomplete_not_found)
            done = len(set(records) | not_found)
            pending = len(reached - set(records) - not_found)
            status = "saved" if found else "Not Found" if sku in not_found else "retry later"
            print(f"SKU {number}/{total} {sku}: {status} | "
                  f"Reached {len(reached)}/{total} | Scanned this run {scanned_this_run} | "
                  f"Finished {done} | Pending retry {pending} | "
                  f"Not yet reached {total - len(reached)} | "
                  f"Images {len(records)} | Not Found {len(not_found)}", flush=True)
            print_progress(total, records, not_found, reached)


if __name__ == "__main__":
    main()
