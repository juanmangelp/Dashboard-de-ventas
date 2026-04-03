import http.server
import json
import os
import threading
from datetime import datetime, timedelta
from collections import defaultdict
from urllib.request import Request, urlopen
from urllib.parse import urlparse, parse_qs

STORE_ID = "87884"
TOKEN = "c8e480ded10aed4f9e1dd31fb15f8ba658c3b72b"
BASE_URL = f"https://api.tiendanube.com/v1/{STORE_ID}"
import os
PORT = int(os.environ.get("PORT", 8765))

HEADERS = {
    "Authentication": f"bearer {TOKEN}",
    "User-Agent": "DashboardSanPretta (dashboard@sanpretta.com)",
    "Content-Type": "application/json"
}

# Progress tracking
_progress = {"pct": 0, "msg": ""}

def set_progress(pct, msg):
    _progress["pct"] = pct
    _progress["msg"] = msg
    print(f"  [{pct}%] {msg}")

def api_get(url):
    req = Request(url, headers=HEADERS)
    with urlopen(req) as resp:
        return json.loads(resp.read())

def fetch_orders(days=None, progress_range=(0,50), label=""):
    results = []
    page = 1
    base = f"{BASE_URL}/orders?payment_status=paid&per_page=200"
    if days:
        since = (datetime.now() - timedelta(days=days)).strftime("%Y-%m-%d")
    else:
        since = (datetime.now() - timedelta(days=730)).strftime("%Y-%m-%d")
    base += f"&created_at_min={since}"
    p_start, p_end = progress_range
    while True:
        data = api_get(f"{base}&page={page}")
        if not isinstance(data, list) or not data:
            break
        results.extend(data)
        # Estimate progress — we don't know total, so use pages as proxy (cap at 80% of range)
        est = min(p_start + int((page / max(page+2, 5)) * (p_end - p_start) * 0.9), p_end - 2)
        set_progress(est, f"{label}: {len(results)} pedidos cargados")
        if len(data) < 200:
            break
        page += 1
    set_progress(p_end, f"{label}: {len(results)} pedidos")
    return results

def fetch_products():
    results = []
    page = 1
    set_progress(2, "Cargando productos...")
    while True:
        data = api_get(f"{BASE_URL}/products?per_page=200&page={page}")
        if not isinstance(data, list) or not data:
            break
        results.extend(data)
        set_progress(4, f"Productos: {len(results)} cargados")
        if len(data) < 200:
            break
        page += 1
    set_progress(8, f"Productos: {len(results)} total")
    return results

def get_name(val):
    if isinstance(val, dict):
        return val.get("es") or val.get("pt") or next(iter(val.values()), "")
    return str(val or "")

def days_since(date_str):
    if not date_str:
        return 9999
    try:
        dt = datetime.fromisoformat(date_str.replace("Z", "+00:00")).replace(tzinfo=None)
        return (datetime.now() - dt).days
    except:
        return 9999

def safe_float(val):
    try:
        v = float(val or 0)
        return v if v > 0 else 0.0
    except:
        return 0.0

def build_variant_map(products):
    variant_map = {}
    product_names = {}
    for p in products:
        pid = p["id"]
        pname = get_name(p.get("name", ""))
        product_names[pid] = pname
        product_created = p.get("created_at", "")
        p_price = safe_float(p.get("price"))
        p_promo = safe_float(p.get("promotional_price"))
        if p_promo >= p_price: p_promo = 0.0
        images = p.get("images", [])
        product_image = images[0].get("src", "") if images and isinstance(images[0], dict) else ""
        for v in p.get("variants", []):
            vid = v["id"]
            parts = [get_name(val) if isinstance(val, dict) else str(val) for val in v.get("values", [])]
            vname = " / ".join(parts) if parts else ""
            try: stock = int(v.get("stock", 0) or 0)
            except: stock = 0
            v_created = v.get("created_at", "") or product_created
            v_price = safe_float(v.get("price")) or p_price
            v_promo = safe_float(v.get("promotional_price")) or p_promo
            if v_promo >= v_price: v_promo = 0.0
            variant_map[vid] = {
                "product_id": pid,
                "product_name": pname,
                "variant_name": vname,
                "stock": stock,
                "days_in_catalog": days_since(v_created),
                "image": product_image,
                "price": v_price,
                "promo_price": v_promo
            }
    return variant_map, product_names

def get_variants_with_sales(orders):
    """Return set of variant_ids that have ANY sales in the given orders."""
    result = set()
    for order in orders:
        for item in order.get("products", []):
            vid = item.get("variant_id")
            if vid and int(item.get("quantity", 0)) > 0:
                result.add(vid)
    return result

def _calc_historical_rate(dates):
    """Calculate real velocity from first sale to last sale (or today if recent)."""
    if not dates:
        return 0.0
    sorted_dates = sorted(dates)
    try:
        first = datetime.strptime(sorted_dates[0], "%Y-%m-%d")
        last = datetime.now()
        span = max((last - first).days, 1)
        return round(len(dates) / span, 3)
    except:
        return 0.0

def build_summary(days):
    set_progress(0, "Iniciando...")
    products = fetch_products()
    set_progress(8, "Procesando productos...")
    variant_map, product_names = build_variant_map(products)

    set_progress(10, f"Cargando pedidos últimos {days} días...")
    period_orders = fetch_orders(days=days, progress_range=(10,45), label=f"Pedidos {days}d")

    set_progress(45, "Cargando historial completo...")
    all_orders = fetch_orders(days=None, progress_range=(45,88), label="Historial")

    # Variants with sales in the period (for rotation metrics)
    period_variants_sold = get_variants_with_sales(period_orders)
    # Variants with ANY sale ever (for stagnant detection)
    all_variants_sold = get_variants_with_sales(all_orders)

    # Build all-time dates per (pid, vid) from historical orders
    all_dates_map = defaultdict(list)
    for order in all_orders:
        order_date = order.get("created_at", "")[:10]
        if not order_date: continue
        for item in order.get("products", []):
            pid = item.get("product_id")
            vid = item.get("variant_id")
            if pid and vid and int(item.get("quantity", 0)) > 0:
                key = (pid, vid)
                if order_date not in all_dates_map[key]:
                    all_dates_map[key].append(order_date)

    # Aggregate period sales by (product_id, variant_id)
    sales = defaultdict(lambda: {"units": 0, "revenue": 0.0, "product_name": "", "variant_name": "", "sale_dates": [], "all_dates": []})
    total_revenue = 0.0
    for order in period_orders:
        total_revenue += float(order.get("total", 0) or 0)
        for item in order.get("products", []):
            pid = item.get("product_id")
            vid = item.get("variant_id")
            qty = int(item.get("quantity", 1))
            price = float(item.get("price", 0) or 0)
            key = (pid, vid)
            sales[key]["units"] += qty
            sales[key]["revenue"] += price * qty
            order_date = order.get("created_at", "")[:10]
            if order_date and order_date not in sales[key]["sale_dates"]:
                sales[key]["sale_dates"].append(order_date)
            if order_date and order_date not in sales[key]["all_dates"]:
                sales[key]["all_dates"].append(order_date)
            if vid and vid in variant_map:
                sales[key]["product_name"] = variant_map[vid]["product_name"]
                sales[key]["variant_name"] = variant_map[vid]["variant_name"]
            else:
                sales[key]["product_name"] = product_names.get(pid, get_name(item.get("name", "")))
                sales[key]["variant_name"] = get_name(item.get("variant", ""))

    # Stagnant: stock > 0 AND never sold (in entire history)
    stagnant = []
    for vid, v in variant_map.items():
        if v["stock"] <= 0: continue
        if not v["variant_name"] or v["variant_name"] == "(sin variante)": continue
        if vid in all_variants_sold: continue  # sold at least once ever → not stagnant
        d = v["days_in_catalog"]
        tipo = "critico" if d >= 180 else "observacion" if d >= 60 else "nuevo"
        stagnant.append({
            "product": v["product_name"],
            "variant": v["variant_name"],
            "stock": v["stock"],
            "days_in_catalog": d,
            "tipo": tipo,
            "image": v.get("image", ""),
            "price": v["price"],
            "promo_price": v["promo_price"]
        })

    tipo_order = {"critico": 0, "observacion": 1, "nuevo": 2}
    stagnant.sort(key=lambda x: (tipo_order[x["tipo"]], -x["days_in_catalog"]))

    # Group period sales by product for main table
    by_product = defaultdict(lambda: {"name": "", "units": 0, "revenue": 0.0, "variants": []})
    for (pid, vid), s in sales.items():
        rate = round(s["units"] / days, 2)
        stock = variant_map.get(vid, {}).get("stock", 0) if vid else 0
        dias_stock = round(stock / rate) if rate > 0 else None
        by_product[pid]["name"] = s["product_name"]
        by_product[pid]["units"] += s["units"]
        by_product[pid]["revenue"] += s["revenue"]
        if s["variant_name"]:
            by_product[pid]["variants"].append({
                "variant_name": s["variant_name"],
                "units": s["units"],
                "revenue": round(s["revenue"], 2),
                "rate": rate,
                "stock": stock,
                "dias_stock": dias_stock,
                "image": variant_map.get(vid, {}).get("image", "") if vid else "",
                "has_promo": variant_map.get(vid, {}).get("promo_price", 0) > 0 if vid else False,
                "sale_dates": sorted(sales[(pid, vid)]["sale_dates"], reverse=True),
                "historical_rate": _calc_historical_rate(all_dates_map.get((pid, vid), []))
            })

    out = []
    for pid, ps in by_product.items():
        ps["variants"].sort(key=lambda x: x["units"], reverse=True)
        has_promo = any(v.get("has_promo", False) for v in ps["variants"])
        out.append({"id": pid, "name": ps["name"], "units": ps["units"], "revenue": round(ps["revenue"], 2), "variants": ps["variants"], "has_promo": has_promo})
    out.sort(key=lambda x: x["units"], reverse=True)

    total_orders = len(period_orders)
    set_progress(98, f"Finalizando: {total_orders} pedidos, {len(stagnant)} estancados")

    return {
        "days": days,
        "total_orders": total_orders,
        "total_units": sum(p["units"] for p in out),
        "total_revenue": round(total_revenue, 2),
        "ticket_promedio": round(total_revenue / total_orders, 2) if total_orders else 0,
        "products": out,
        "stagnant": stagnant
    }

_cache = {}

class Handler(http.server.BaseHTTPRequestHandler):
    def log_message(self, *a): pass
    def send_cors(self):
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Methods", "GET, OPTIONS")
    def do_OPTIONS(self):
        self.send_response(200); self.send_cors(); self.end_headers()
    def do_GET(self):
        if self.path == "/": self.serve_file("dashboard.html", "text/html")
        elif self.path == "/progress":
            data = json.dumps(_progress).encode()
            self.send_response(200)
            self.send_header("Content-Type", "application/json")
            self.send_cors()
            self.end_headers()
            self.wfile.write(data)
        elif self.path.startswith("/summary"): self.serve_summary()
        elif self.path == "/invalidate":
            _cache.clear()
            self.send_response(200); self.send_cors(); self.end_headers()
            self.wfile.write(b'{"ok":true}')
        else:
            self.send_response(404); self.end_headers()
    def serve_file(self, filename, content_type):
        fp = os.path.join(os.path.dirname(os.path.abspath(__file__)), filename)
        try:
            with open(fp, "rb") as f: content = f.read()
            self.send_response(200)
            self.send_header("Content-Type", content_type)
            self.send_cors(); self.end_headers()
            self.wfile.write(content)
        except FileNotFoundError:
            self.send_response(404); self.end_headers()
            self.wfile.write(b"Archivo no encontrado")
    def serve_summary(self):
        qs = parse_qs(urlparse(self.path).query)
        days = int(qs.get("days", ["90"])[0])
        key = f"s{days}"
        if key not in _cache:
            set_progress(0, "Iniciando...")
            _cache[key] = build_summary(days)
            set_progress(100, "Listo")
        data = json.dumps(_cache[key]).encode()
        self.send_response(200)
        self.send_header("Content-Type", "application/json")
        self.send_cors(); self.end_headers()
        self.wfile.write(data)

if __name__ == "__main__":
    server = http.server.ThreadingHTTPServer(("0.0.0.0", PORT), Handler)
    print(f"\n  Dashboard San Pretta")
    print(f"  Abri http://localhost:{PORT} en tu browser")
    print(f"  Primera carga tarda mas porque descarga historial completo")
    print(f"  Ctrl+C para detener\n")
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\n  Servidor detenido.")
