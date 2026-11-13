#!/usr/bin/env python3
"""
Publicación masiva de productos Valmax en Mercado Libre (Argentina) vía API oficial.

Lee la hoja "Publicaciones" de la planilla (ML_XLSX) y crea una publicación por fila.

Comandos:
  python publicar_ml.py auth              -> muestra el link para autorizar la app
  python publicar_ml.py token CODIGO      -> canjea el código por el token (se guarda en token.json)
  python publicar_ml.py previa            -> tabla de lo que se publicaría y lo que se saltea (sin API)
  python publicar_ml.py validar           -> prueba cada publicación sin publicar nada (/items/validate)
  python publicar_ml.py publicar          -> publica todo lo que pasa los controles
  python publicar_ml.py publicar VMX-RS3-10 VMX-TB1-20   -> publica sólo esos SKU

Antes de publicar se saltea: SKU excluidos, fotos inexistentes, contribución negativa o sin costo,
y SKU que ya tienen una publicación activa. Los resultados se acumulan en resultados_publicacion.csv.

Variables (en el archivo .env, de tu app en https://developers.mercadolibre.com.ar/devcenter):
  ML_CLIENT_ID, ML_CLIENT_SECRET, ML_REDIRECT_URI, ML_XLSX (opcional)
Requisitos: pip install requests openpyxl python-dotenv
"""
import csv, json, os, re, sys, time
import requests
from dotenv import load_dotenv
from openpyxl import load_workbook

load_dotenv(os.path.join(os.path.dirname(os.path.abspath(__file__)), ".env"))

API = "https://api.mercadolibre.com"
XLSX = os.environ.get("ML_XLSX", "Publicaciones_Valmax_ML.xlsx")
TOKEN_FILE = "token.json"
RESULTADOS = "resultados_publicacion.csv"
FOTOS_FILE = "fotos_subidas.json"

# Categorías verificadas en la API de Mercado Libre (octubre 2026)
CATEGORIAS = {"gel": "MLA417136",   # Geles Refrigerantes
              "rg":  "MLA417136",
              "kit": "MLA417136",
              "tb":  "MLA417158"}   # Bolsos Térmicos

# IDs de valores de atributos tomados de /categories/{id}/attributes
SALE_FORMAT_PACK, SALE_FORMAT_UNIDAD = "1359392", "1359391"
GTIN_VACIO_PACK = "17055159"        # "El producto es un kit o un pack"
GTIN_VACIO_UNIDAD = "17055160"
IVA_21 = "48405909"
IMP_INTERNO_0 = "49553239"
TIPO_BOLSO = "4301012"
TIPO_GEL = "46833259"               # "Geles refrigerantes"
EFECTO_FRIO = "27624977"
USO_ALIMENTOS = "27624981"
SI, NO = "242085", "242084"

# SKU que no se publican por decisión del dueño (aunque estén en la planilla)
EXCLUIDOS = {"VMX-RS7-02", "VMX-RG1-10", "VMX-RG2-02", "VMX-RG3-01"}

PAUSA = 1.0                         # segundos mínimos entre publicaciones (límite de la API)


# ---------- errores de la API ----------
def causas(r, tipo):
    try:
        return [c.get("message") for c in r.json().get("cause", []) if c.get("type") == tipo]
    except Exception:
        return []

def errores(r):
    """Devuelve el mensaje de error tal como lo manda Mercado Libre."""
    try:
        j = r.json()
        return "; ".join(causas(r, "error")) or j.get("error") or j.get("message") or r.text
    except Exception:
        return r.text

def validacion_ok(r):
    """/items/validate responde 400 aunque sólo haya advertencias: vale si no hay causas de tipo error."""
    if r.status_code < 300:
        return True
    try:
        j = r.json()
    except Exception:
        return False
    return j.get("error") == "validation_error" and j.get("cause") and not causas(r, "error")

def exigir_ok(r, que):
    if r.status_code >= 300:
        sys.exit(f"[ERROR] {que}: HTTP {r.status_code} - {errores(r)}")
    return r.json()


# ---------- autenticación ----------
def env(k):
    v = os.environ.get(k)
    if not v:
        sys.exit(f"Falta {k} en el archivo .env")
    return v

def guardar_token(data):
    data["obtenido"] = int(time.time())
    with open(TOKEN_FILE, "w") as f:
        json.dump(data, f, indent=2)

def cmd_auth():
    print("Abrí este link logueado con la cuenta de Mercado Libre donde querés publicar y autorizá la app:\n")
    print(f"https://auth.mercadolibre.com.ar/authorization?response_type=code"
          f"&client_id={env('ML_CLIENT_ID')}&redirect_uri={env('ML_REDIRECT_URI')}")
    print("\nDespués copiá el valor de ?code=... de la URL a la que te redirige y corré:")
    print("  python publicar_ml.py token CODIGO")

def cmd_token(code):
    r = requests.post(f"{API}/oauth/token", data={
        "grant_type": "authorization_code", "client_id": env("ML_CLIENT_ID"),
        "client_secret": env("ML_CLIENT_SECRET"), "code": code,
        "redirect_uri": env("ML_REDIRECT_URI")})
    guardar_token(exigir_ok(r, "canje del código"))
    print("Token guardado en token.json")

_token = None
def token():
    global _token
    if not os.path.exists(TOKEN_FILE):
        sys.exit("No hay token. Corré primero: python publicar_ml.py auth")
    t = _token or json.load(open(TOKEN_FILE))
    if time.time() - t["obtenido"] > t.get("expires_in", 21600) - 300:   # dura 6 h
        if not t.get("refresh_token"):
            sys.exit("El token venció y la app no tiene 'offline access'. Volvé a correr: python publicar_ml.py auth")
        r = requests.post(f"{API}/oauth/token", data={
            "grant_type": "refresh_token", "client_id": env("ML_CLIENT_ID"),
            "client_secret": env("ML_CLIENT_SECRET"), "refresh_token": t["refresh_token"]})
        t = exigir_ok(r, "renovación del token"); guardar_token(t)
    _token = t
    return t["access_token"]

def H():
    return {"Authorization": f"Bearer {token()}"}

def user_id():
    token()
    return _token.get("user_id") or exigir_ok(requests.get(f"{API}/users/me", headers=H()), "/users/me")["id"]


# ---------- lectura de la planilla ----------
def leer_filas():
    ws = load_workbook(XLSX, data_only=True)["Publicaciones"]
    cab = [c.value for c in ws[1]]
    idx = {n: i for i, n in enumerate(cab)}
    filas = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row[0] or not str(row[0]).startswith("VMX-"):
            continue
        g = lambda n: row[idx[n]]
        filas.append({
            "sku": g("SKU"), "titulo": g("Título (máx. 60)"), "precio": g("Precio ML (con IVA)"),
            "stock": int(g("Stock") or 0), "tipo_pub": g("Tipo de publicación") or "Clásica",
            "modelo": str(g("Modelo")), "ean": str(g("EAN") or "").strip(),
            "unidades": int(g("Unidades por pack") or 1), "peso": g("Peso estimado (g)"),
            "costo": g("Costo de producción"), "contribucion": g("Contribución por venta"),
            "fotos": [p.strip() for p in str(g("Fotos (URL o ruta, separadas por |)") or "").split("|") if p.strip()],
            "descripcion": g("Descripción") or ""})
    return filas

def familia(f):
    m = f["modelo"]
    if m.startswith("RS"): return "gel"
    if m.startswith("RG"): return "rg"
    if m.startswith("TB") and "+" not in m: return "tb"
    return "kit"


# ---------- controles previos ----------
def control_local(f):
    """Motivo para saltear la fila sin consultar la API, o "" si está apta."""
    if f["sku"] in EXCLUIDOS:
        return "excluido por decisión del dueño"
    faltan = [p for p in f["fotos"] if not p.startswith("http") and not os.path.isfile(p)]
    if not f["fotos"]:
        return "sin fotos en la planilla"
    if faltan:
        return "fotos inexistentes: " + ", ".join(faltan)
    c = f["contribucion"]
    if not isinstance(c, (int, float)):
        return "sin costo de producción: no se puede verificar que el precio cubra el costo"
    if c < 0:
        return f"contribución negativa ({c:,.0f}): se publicaría por debajo del costo"
    return ""

def publicacion_activa(f, uid):
    """ID de una publicación activa con este SKU, o None."""
    r = requests.get(f"{API}/users/{uid}/items/search", headers=H(),
                     params={"seller_sku": f["sku"], "status": "active"})
    ids = exigir_ok(r, f"búsqueda de duplicados de {f['sku']}").get("results", [])
    return ids[0] if ids else None

def separar(filas, consultar_api=True):
    """Devuelve (aptas, salteadas[(fila, motivo)])."""
    aptas, salteadas = [], []
    uid = user_id() if consultar_api else None
    for f in filas:
        motivo = control_local(f)
        if not motivo and consultar_api:
            ya = publicacion_activa(f, uid)
            if ya: motivo = f"ya publicado ({ya})"
        if motivo:
            salteadas.append((f, motivo)); print(f"[SALTEO] {f['sku']}: {motivo}")
        else:
            aptas.append(f)
    return aptas, salteadas


# ---------- armado del payload ----------
def subir_foto(ruta):
    """Sube la foto una sola vez y recuerda su ID en FOTOS_FILE (se vuelve a subir si el archivo cambia)."""
    cache = json.load(open(FOTOS_FILE)) if os.path.exists(FOTOS_FILE) else {}
    clave = f"{ruta}|{os.path.getmtime(ruta)}|{os.path.getsize(ruta)}"
    if clave in cache:
        return {"id": cache[clave]}
    with open(ruta, "rb") as fh:
        r = requests.post(f"{API}/pictures/items/upload", headers=H(), files={"file": fh})
    if r.status_code >= 300:
        raise RuntimeError(f"foto {ruta}: {errores(r)}")
    cache[clave] = r.json()["id"]
    with open(FOTOS_FILE, "w") as fh:
        json.dump(cache, fh, indent=2)
    return {"id": cache[clave]}

def fotos(f):
    out = []
    for p in f["fotos"]:
        out.append({"source": p} if p.startswith("http") else subir_foto(p))
    return out

def atributos(f, fam):
    a = [{"id": "BRAND", "value_name": "Valmax"},
         {"id": "MODEL", "value_name": f["modelo"]},
         {"id": "SELLER_SKU", "value_name": f["sku"]},
         {"id": "VALUE_ADDED_TAX", "value_id": IVA_21},
         {"id": "IMPORT_DUTY", "value_id": IMP_INTERNO_0}]
    if f["ean"]:
        a.append({"id": "GTIN", "value_name": f["ean"]})
    else:
        a.append({"id": "EMPTY_GTIN_REASON", "value_id": GTIN_VACIO_PACK if f["unidades"] > 1 else GTIN_VACIO_UNIDAD})
    if fam in ("gel", "rg", "kit"):
        a.append({"id": "SALE_FORMAT", "value_id": SALE_FORMAT_PACK if f["unidades"] > 1 else SALE_FORMAT_UNIDAD})
        if f["unidades"] > 1:
            a.append({"id": "UNITS_PER_PACK", "value_name": str(f["unidades"])})
        # Ficha técnica: Mercado Libre baja la calidad de la publicación si faltan estas características
        a += [{"id": "THERMAL_EFFECT", "value_id": EFECTO_FRIO},
              {"id": "PRODUCT_TYPE", "value_id": TIPO_GEL},
              {"id": "IS_REUSABLE", "value_id": SI},
              {"id": "IS_TOXIC", "value_id": NO},
              {"id": "RECOMMENDED_USES", "values": [{"id": USO_ALIMENTOS, "name": "Alimentos"}, {"name": "Medicamentos"}]}]
    if fam == "tb":
        a.append({"id": "PRODUCT_TYPE", "value_id": TIPO_BOLSO})
        a.append({"id": "MATERIALS", "values": [{"name": "Espuma de polietileno"}]})
        medida = re.search(r"(\d+)x(\d+)", f["titulo"])     # "Termi Bag 20x35" -> ancho 20, largo 35
        if medida:
            a += [{"id": "WIDTH", "value_name": f"{medida[1]} cm"}, {"id": "LENGTH", "value_name": f"{medida[2]} cm"}]
    return a

def payload(f, con_fotos=True):
    fam = familia(f)
    p = {
        # Ambas categorías usan el modelo "User Products": piden family_name y rechazan title
        "family_name": f["titulo"],
        "category_id": CATEGORIAS[fam],
        "price": f["precio"],
        "currency_id": "ARS",
        "available_quantity": f["stock"],
        "buying_mode": "buy_it_now",
        "condition": "new",
        "listing_type_id": "gold_pro" if f["tipo_pub"].lower().startswith("prem") else "gold_special",
        "seller_custom_field": f["sku"],
        "attributes": atributos(f, fam),
        "shipping": {"mode": "me2", "local_pick_up": False, "free_shipping": False},
        "sale_terms": [{"id": "WARRANTY_TYPE", "value_name": "Garantía del vendedor"},
                       {"id": "WARRANTY_TIME", "value_name": "30 días"}],
    }
    if con_fotos:
        p["pictures"] = fotos(f)
    return p


# ---------- resultados ----------
def guardar_resultados(nuevos):
    """Acumula en el CSV: cada corrida actualiza sus SKU sin borrar los de corridas anteriores."""
    filas = {}
    if os.path.exists(RESULTADOS):
        with open(RESULTADOS, newline="", encoding="utf-8") as fh:
            for r in csv.DictReader(fh): filas[r["SKU"]] = r
    for r in nuevos:
        filas[r["SKU"]] = r
    with open(RESULTADOS, "w", newline="", encoding="utf-8") as fh:
        w = csv.DictWriter(fh, fieldnames=["SKU", "ID", "Link", "Estado"])
        w.writeheader(); w.writerows(filas.values())


# ---------- comandos ----------
def cmd_previa():
    aptas, salteadas = separar(leer_filas(), consultar_api=False)
    print(f"\n{'SKU':<14}{'Título':<58}{'Precio':>9}  Fotos")
    for f in aptas:
        print(f"{f['sku']:<14}{f['titulo']:<58}{f['precio']:>9,}  {len(f['fotos'])}")
    print(f"\n{len(aptas)} aptas, {len(salteadas)} salteadas (sin verificar duplicados en Mercado Libre).")

def cmd_validar():
    aptas, _ = separar(leer_filas())
    ok, avisos = 0, set()
    for f in aptas:
        try:
            p = payload(f)
        except RuntimeError as e:
            print(f"[ERROR] {f['sku']}: {e}"); continue
        r = requests.post(f"{API}/items/validate", headers=H(), json=p)
        if validacion_ok(r):
            ok += 1; avisos.update(causas(r, "warning")); print(f"[OK]    {f['sku']}  {f['titulo']}")
        else:
            print(f"[ERROR] {f['sku']}: HTTP {r.status_code} - {errores(r)}")
    if avisos:
        print("\nAdvertencias de Mercado Libre (no impiden publicar):\n  " + "\n  ".join(sorted(avisos)))
    print(f"\n{ok} de {len(aptas)} publicaciones listas para publicar.")

def cmd_publicar(skus):
    filas = [f for f in leer_filas() if not skus or f["sku"] in skus]
    if skus and len(filas) != len(set(skus)):
        sys.exit(f"SKU inexistentes en la planilla: {set(skus) - {f['sku'] for f in filas}}")
    aptas, salteadas = separar(filas)
    res = [{"SKU": f["sku"], "ID": m.split("(")[-1].rstrip(")") if m.startswith("ya publicado") else "",
            "Link": "", "Estado": f"salteado: {m}"} for f, m in salteadas]
    ultima = 0.0
    for f in aptas:
        try:
            p = payload(f)
        except RuntimeError as e:
            print(f"[ERROR] {f['sku']}: {e}")
            res.append({"SKU": f["sku"], "ID": "", "Link": "", "Estado": f"error: {e}"}); continue
        time.sleep(max(0, PAUSA - (time.time() - ultima)))
        r = requests.post(f"{API}/items", headers=H(), json=p)
        ultima = time.time()
        if r.status_code >= 300:
            e = errores(r)
            print(f"[ERROR] {f['sku']}: HTTP {r.status_code} - {e}")
            res.append({"SKU": f["sku"], "ID": "", "Link": "", "Estado": f"error: {e}"}); continue
        item = r.json()
        d = requests.post(f"{API}/items/{item['id']}/description", headers=H(),
                          json={"plain_text": f["descripcion"]})
        estado = item.get("status", "publicado")
        if d.status_code >= 300:
            estado += f" (descripción no cargada: {errores(d)})"
        print(f"[PUBLICADO] {f['sku']} -> {item['id']}  {item.get('permalink','')}  [{estado}]")
        res.append({"SKU": f["sku"], "ID": item["id"], "Link": item.get("permalink", ""), "Estado": estado})
    guardar_resultados(res)
    print(f"\nResultados en {RESULTADOS}")

if __name__ == "__main__":
    a = sys.argv[1:]
    if not a: print(__doc__)
    elif a[0] == "auth": cmd_auth()
    elif a[0] == "token": cmd_token(a[1])
    elif a[0] == "previa": cmd_previa()
    elif a[0] == "validar": cmd_validar()
    elif a[0] == "publicar": cmd_publicar(a[1:])
    else: print(__doc__)
