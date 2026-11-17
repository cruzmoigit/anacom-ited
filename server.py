import json, os, re, time, threading, urllib.parse, tempfile, uuid
from http.server import ThreadingHTTPServer, BaseHTTPRequestHandler
from datetime import date, timedelta
import httpx
from pypdf import PdfReader
import openpyxl
from openpyxl.styles import Font

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DIST_FILE = os.path.join(BASE_DIR, "distritos.json")
CACHE_DIR = os.path.join(tempfile.gettempdir(), "anacom_ited_cache")
os.makedirs(CACHE_DIR, exist_ok=True)

with open(DIST_FILE, encoding="utf-8") as f:
    DIST_MAP = json.load(f)

ANACOM_BASE = "https://www.anacom.pt"
FORM_URL = "https://www.anacom.pt/eh/termosITED.jsp?categoryId=336799"
POST_URL = "https://www.anacom.pt/eh/index.jsp?categoryId=336799&channel="

UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"


import logging
logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
log = logging.getLogger("anacom")

def _retry(fn, attempts=3, base_delay=1.0):
    last = None
    for i in range(attempts):
        try:
            return fn()
        except Exception as e:
            last = e
            log.warning("tentativa %d/%d falhou: %s", i + 1, attempts, e)
            if i < attempts - 1:
                time.sleep(base_delay * (2 ** i))
    raise last

def new_session():
    s = httpx.Client(headers={"User-Agent": UA}, timeout=httpx.Timeout(120.0, connect=30.0), follow_redirects=True)
    _retry(lambda: s.get(FORM_URL))
    log.info("sessao ANACOM criada")
    return s


def build_body(distrito, concelho, ini, fim):
    di, mi, yi = f"{ini.day:02d}", f"{ini.month:02d}", str(ini.year)
    df, mf, yf = f"{fim.day:02d}", f"{fim.month:02d}", str(fim.year)
    # default nome text (encoded) matching browser
    nome = "Use+a+op%E7%E3o+%22Escolher+T%E9cnico%22+para+fazer+a+sua+sele%E7%E3o."
    numinsc = "N%FAmero+de+inscri%E7%E3o"
    pairs = [
        ("tipoTermo", "TRI"),
        ("ano", ""),
        ("numero", ""),
        ("nome", "Use a opção \"Escolher Técnico\" para fazer a sua seleção."),
        ("num_insc", "Número de inscrição"),
        ("cbDistritos", distrito),
        ("cbConcelhos", concelho),
        ("Localidade", "Localidade"),
        ("cbDistritosJS", "-"),
        ("cbConcelhosJS", "-"),
        ("Localidade", "Localidade"),
        ("rua", "Rua"),
        ("cod_postal", "0000-000"),
        ("numfogos", "000"),
        ("tipoedificioR", "R"),
        ("dia_ini", di),
        ("mes_ini", mi),
        ("ano_ini", yi),
        ("dia_fim", df),
        ("mes_fim", mf),
        ("ano_fim", yf),
        ("pesquisar", "Pesquisar"),
        ("modo", ""),
        ("ambito", "mv_bvit_termosited"),
    ]
    # Encode in ISO-8859-15 to match ANACOM's expected encoding
    return urllib.parse.urlencode(pairs, encoding="iso-8859-15")


def parse_results(html):
    # returns (total, list of dict termo/emitido/nome/num/morada/link_termo_pk, num_pages, current_page)
    results = []
    # total
    m = re.search(r"produziu\s+(\d+)\s+resultados", html)
    total = int(m.group(1)) if m else 0
    # table rows: find <tr> containing "Ver termo"
    for tr in re.finditer(r"<tr[^>]*>(.*?)</tr>", html, re.S):
        row = tr.group(1)
        if "Ver termo" not in row:
            continue
        # link
        lm = re.search(r'href="([^"]*termoPk=(\d+)[^"]*)"', row)
        if not lm:
            continue
        href = lm.group(1).replace("&amp;", "&")
        pk = lm.group(2)
        # termo text
        tm = re.search(r"Ver termo\s+([\d/]+)\s+emitido em\s+([\d-]+)", row)
        termo = tm.group(1) if tm else ""
        emitido = tm.group(2) if tm else ""
        # cells
        cells = re.findall(r"<td[^>]*>(.*?)</td>", row, re.S)
        # strip tags
        def strip(x):
            return re.sub(r"<[^>]+>", "", x).strip()
        cells = [strip(c) for c in cells]
        # columns: termo(link), nome, num_insc, morada
        nome = cells[1] if len(cells) > 1 else ""
        numinsc = cells[2] if len(cells) > 2 else ""
        morada = cells[3] if len(cells) > 3 else ""
        results.append({
            "termo": termo,
            "emitido": emitido,
            "nome": nome,
            "num_inscricao": numinsc,
            "morada": morada,
            "termo_pk": pk,
            "href": href,
        })
    # pagination: find "Pág. X de Y"
    pm = re.search(r"P[aá]g\.\s*(\d+)\s*de\s*(\d+)", html)
    cur = int(pm.group(1)) if pm else 1
    tot_pages = int(pm.group(2)) if pm else 1
    return total, results, cur, tot_pages


def fetch_all_results(s, distrito, concelho, ini, fim):
    pairs = build_body(distrito, concelho, ini, fim)
    # first page
    r = s.post(POST_URL, content=pairs, timeout=90, headers={"Referer": FORM_URL, "Content-Type": "application/x-www-form-urlencoded"})
    r.encoding = "iso-8859-15"
    log.info("POST status=%s url=%s len=%d", r.status_code, r.url, len(r.text))
    # debug: procurar indicadores de erro/bloqueio na resposta
    snippet = r.text[:500].replace("\n", " ")
    log.info("resposta snippet: %s", snippet)
    total, results, cur, tot_pages = parse_results(r.text)
    log.info("parse_results: total=%d resultados=%d paginas=%d", total, len(results), tot_pages)
    all_results = list(results)
    # follow pages: the form resubmission with same params but page number?
    # ANACOM pagination uses GET with modo=resultado and a 'pagina' param? Let's inspect.
    # Try incrementing by re-POSTing not needed; usually pagination links are GET with params + pagina=N
    page = 1
    while page < tot_pages:
        page += 1
        url = str(r.url).split("#")[0] + "&pagina=" + str(page)
        try:
            rr = s.get(url, timeout=90, headers={"Referer": FORM_URL})
            rr.encoding = "iso-8859-15"
            _, res, _, _ = parse_results(rr.text)
            if not res:
                break
            all_results.extend(res)
        except Exception:
            break
    return total, all_results, tot_pages


def pdf_cache_path(pk):
    return os.path.join(CACHE_DIR, f"{pk}.pdf")


def download_pdf(s, href, pk):
    path = pdf_cache_path(pk)
    if os.path.exists(path) and os.path.getsize(path) > 0:
        return path
    if href.startswith("http"):
        url = href
    elif href.startswith("/eh/"):
        url = ANACOM_BASE + href
    elif href.startswith("/"):
        url = ANACOM_BASE + "/eh" + href
    else:
        url = ANACOM_BASE + "/eh/" + href
    # ensure categoryId present
    r = s.get(url, timeout=120, headers={"Referer": FORM_URL})
    r.raise_for_status()
    with open(path, "wb") as f:
        f.write(r.content)
    return path


def parse_pdf(path):
    reader = PdfReader(path)
    text = ""
    for pg in reader.pages:
        text += (pg.extract_text() or "") + "\n"
    # normalize
    t = text

    def grab(label, pattern):
        m = re.search(pattern, t, re.S | re.I)
        return m.group(1).strip() if m else ""

    termo = grab("termo", r"TERMO:\s*([^\n]+)")
    data_emissao = grab("data", r"DATA DE EMISS[AÃ]O:\s*([^\n]+)")
    # num fogos: number immediately before N.º DE FOGOS
    nf = re.search(r"(\d+)\s*N\.?\s*º?\s*DE FOGOS", t, re.I)
    num_fogos = int(nf.group(1)) if nf else 0
    # coords
    cm = re.search(r"(N\d+[\u00b0\u00ba]\d+'[\d,\.]+''?\s*W\d+[\u00b0\u00ba]\d+'[\d,\.]+''?)", t)
    coords = cm.group(1).strip() if cm else ""
    # rua/localidade/cp: labels inline
    rua = grab("rua", r"([^\n]+?)Rua\b")
    localidade = grab("loc", r"([^\n]+?)Localidade\b")
    cp = grab("cp", r"([^\n]+?)C[oó]digo Postal\b")
    # dono de obra: text after DONO DA OBRA:
    do = re.search(r"([^\n]+?)DONO DA OBRA:", t, re.I)
    dono = do.group(1).strip() if do else ""
    # certificado date: look for "Certificado" 
    cert = ""
    cem = re.search(r"Certificado[^:]*:\s*([^\n]+)", t, re.I)
    if cem:
        cert = cem.group(1).strip()
    if not cert:
        cert = data_emissao
    # distrito/concelho/freguesia
    distrito = grab("dist", r"([^\n]+?)Distrito\b")
    concelho = grab("conc", r"([^\n]+?)Concelho\b")
    freguesia = grab("freg", r"([^\n]+?)Freguesia\b")

    return {
        "termo": termo,
        "data_emissao": data_emissao,
        "num_fogos": num_fogos,
        "rua": rua,
        "localidade": localidade,
        "codigo_postal": cp,
        "coordenadas": coords,
        "dono_obra": dono,
        "data_certificado": cert,
        "distrito": distrito,
        "concelho": concelho,
        "freguesia": freguesia,
    }


class Handler(BaseHTTPRequestHandler):
    def _send(self, code, content_type, data):
        self.send_response(code)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Length", str(len(data)))
        self.send_header("Access-Control-Allow-Origin", "*")
        self.end_headers()
        self.wfile.write(data)

    def do_GET(self):
        parsed = urllib.parse.urlparse(self.path)
        path = parsed.path
        if path == "/" or path == "/index.html":
            with open(os.path.join(BASE_DIR, "static", "index.html"), "rb") as f:
                self._send(200, "text/html; charset=utf-8", f.read())
            return
        if path == "/api/distritos":
            self._send(200, "application/json; charset=utf-8", json.dumps(list(DIST_MAP.keys()), ensure_ascii=False).encode("utf-8"))
            return
        qs = urllib.parse.parse_qs(parsed.query)
        if path == "/api/concelhos":
            dist = qs.get("distrito", [""])[0]
            concs = DIST_MAP.get(dist, [])
            # prepend blank option meaning "all concelhos of the district"
            out = ["-"] + list(concs)
            self._send(200, "application/json; charset=utf-8", json.dumps(out, ensure_ascii=False).encode("utf-8"))
            return
        if path == "/api/exportar.xlsx":
            # read last extraction from temp file
            xlsx_path = os.path.join(CACHE_DIR, "last_extraction.json")
            if not os.path.exists(xlsx_path):
                self._send(404, "text/plain", b"Nenhum dado para exportar")
                return
            with open(xlsx_path, encoding="utf-8") as f:
                rows = json.load(f)
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Termos ITED"
            headers = ["Termo", "Data emissão", "N.º fogos", "Rua", "Localidade", "Código Postal", "Coordenadas", "Dono de obra", "Data certificado", "Distrito", "Concelho", "Freguesia"]
            ws.append(headers)
            for c in ws[1]:
                c.font = Font(bold=True)
            for r in rows:
                ws.append([
                    r.get("termo"), r.get("data_emissao"), r.get("num_fogos"), r.get("rua"),
                    r.get("localidade"), r.get("codigo_postal"), r.get("coordenadas"),
                    r.get("dono_obra"), r.get("data_certificado"), r.get("distrito"),
                    r.get("concelho"), r.get("freguesia"),
                ])
            out = os.path.join(CACHE_DIR, "export.xlsx")
            wb.save(out)
            with open(out, "rb") as f:
                data = f.read()
            self.send_response(200)
            self.send_header("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            self.send_header("Content-Disposition", "attachment; filename=termos_ited.xlsx")
            self.send_header("Content-Length", str(len(data)))
            self.end_headers()
            self.wfile.write(data)
            return
        # static files
        if path.startswith("/static/"):
            fp = os.path.join(BASE_DIR, path.lstrip("/"))
            if os.path.exists(fp):
                with open(fp, "rb") as f:
                    self._send(200, "application/octet-stream", f.read())
                return
        self._send(404, "text/plain", b"Not found")

    def do_POST(self):
        parsed = urllib.parse.urlparse(self.path)
        path = parsed.path
        length = int(self.headers.get("Content-Length", 0))
        raw = self.rfile.read(length) if length else b""
        try:
            body = json.loads(raw.decode("utf-8")) if raw else {}
        except Exception:
            body = {}
        try:
            if path == "/api/pesquisar":
                distrito = body.get("distrito", "")
                concelho = body.get("concelho", "")
                semanas = int(body.get("semanas", 1))
                hoje = date.today()
                ini = hoje - timedelta(days=semanas * 7)
                fim = hoje
                s = new_session()
                total, results, tot_pages = fetch_all_results(s, distrito, concelho, ini, fim)
                resp = {
                    "total": total,
                    "paginas": tot_pages,
                    "resultados": results,
                    "intervalo": {"de": str(ini), "ate": str(fim)},
                }
                self._send(200, "application/json; charset=utf-8", json.dumps(resp, ensure_ascii=False).encode("utf-8"))
                return
            if path == "/api/extrair":
                items = body.get("items", [])
                s = new_session()
                extracted = []
                errors = []
                for it in items:
                    pk = it.get("termo_pk")
                    href = it.get("href", "")
                    try:
                        pathpdf = download_pdf(s, href, pk)
                        info = parse_pdf(pathpdf)
                        info["termo_pk"] = pk
                        if info["num_fogos"] >= 3:
                            extracted.append(info)
                    except Exception as e:
                        errors.append({"termo_pk": pk, "erro": str(e)})
                    time.sleep(0.3)
                # save for export
                with open(os.path.join(CACHE_DIR, "last_extraction.json"), "w", encoding="utf-8") as f:
                    json.dump(extracted, f, ensure_ascii=False, indent=2)
                resp = {"extraidos": len(extracted), "processados": len(items), "erros": errors, "dados": extracted}
                self._send(200, "application/json; charset=utf-8", json.dumps(resp, ensure_ascii=False).encode("utf-8"))
                return
            self._send(404, "text/plain", b"Not found")
        except Exception as e:
            import traceback
            traceback.print_exc()
            err = {"erro": str(e), "tipo": type(e).__name__}
            try:
                self._send(500, "application/json; charset=utf-8", json.dumps(err, ensure_ascii=False).encode("utf-8"))
            except Exception:
                pass

    def log_message(self, *args):
        pass


if __name__ == "__main__":
    host = os.environ.get("HOST", "0.0.0.0")
    port = int(os.environ.get("PORT", "8765"))
    httpd = ThreadingHTTPServer((host, port), Handler)
    print(f"Servidor a correr em http://{host}:{port}", flush=True)
    httpd.serve_forever()
