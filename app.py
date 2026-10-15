from flask import Flask, request, send_file, jsonify, render_template_string
from pymongo import MongoClient
import pandas as pd
import requests
import unicodedata
import re
import io
import time
from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment

import os

app = Flask(__name__)

MONGO_URI = os.environ["MONGO_URI"]
GOOGLE_API_KEY = os.environ["GOOGLE_API_KEY"]
CLAUDE_API_KEY = os.environ["CLAUDE_API_KEY"]

client = MongoClient(MONGO_URI, serverSelectionTimeoutMS=10000)
coll = client["test"]["cnesestabelecimentos"]

HTML = """
<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Score Prestador Odontológico</title>
<style>
  * { margin: 0; padding: 0; box-sizing: border-box; }
  body {
    background: #000;
    color: #fff;
    font-family: system-ui, -apple-system, sans-serif;
    min-height: 100vh;
    display: flex;
    align-items: center;
    justify-content: center;
  }
  .box {
    text-align: center;
    padding: 3rem 2rem;
  }
  h1 {
    font-size: 1.6rem;
    font-weight: 500;
    margin-bottom: 2.5rem;
    letter-spacing: 0.5px;
  }
  .upload {
    border: 1px dashed #333;
    border-radius: 12px;
    padding: 3rem 2rem;
    cursor: pointer;
    transition: border-color 0.2s;
    max-width: 420px;
  }
  .upload:hover, .upload.drag {
    border-color: #666;
  }
  .upload input { display: none; }
  .upload p {
    color: #888;
    font-size: 0.95rem;
    margin-top: 0.5rem;
  }
  button {
    margin-top: 1.5rem;
    background: #fff;
    color: #000;
    border: none;
    padding: 0.75rem 2rem;
    border-radius: 8px;
    font-weight: 600;
    cursor: pointer;
    font-size: 0.95rem;
  }
  button:disabled {
    background: #222;
    color: #555;
    cursor: not-allowed;
  }
  #status {
    margin-top: 1.5rem;
    color: #aaa;
    font-size: 0.9rem;
    min-height: 1.4rem;
  }
  .hidden { display: none; }
</style>
</head>
<body>
<div class="box">
  <h1>Score Prestador Odontológico</h1>
  <div class="upload" id="zone">
    <input type="file" id="file" accept=".xlsx,.xls,.csv">
    <div>Solte o Excel aqui</div>
    <p>ou clique para selecionar</p>
  </div>
  <button id="btn" disabled>Processar</button>
  <div id="status"></div>
</div>
<div class="beta">essa é uma versão beta do ÍCARO</div>

<script>
const zone = document.getElementById('zone');
const fileInput = document.getElementById('file');
const btn = document.getElementById('btn');
const status = document.getElementById('status');
let selectedFile = null;

zone.onclick = () => fileInput.click();
zone.ondragover = e => { e.preventDefault(); zone.classList.add('drag'); };
zone.ondragleave = () => zone.classList.remove('drag');
zone.ondrop = e => {
  e.preventDefault();
  zone.classList.remove('drag');
  if (e.dataTransfer.files[0]) setFile(e.dataTransfer.files[0]);
};
fileInput.onchange = e => { if (e.target.files[0]) setFile(e.target.files[0]); };

function setFile(f) {
  selectedFile = f;
  zone.querySelector('div').textContent = f.name;
  btn.disabled = false;
  status.textContent = '';
}

btn.onclick = async () => {
  if (!selectedFile) return;
  btn.disabled = true;
  status.textContent = 'Processando (API + Mongo + Google + Claude)... isso pode levar vários minutos.';
  const fd = new FormData();
  fd.append('file', selectedFile);
  try {
    const res = await fetch('/process', { method: 'POST', body: fd });
    if (!res.ok) throw new Error(await res.text());
    const blob = await res.blob();
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = 'prestadores_enriquecidos.xlsx';
    a.click();
    status.textContent = 'Arquivo gerado e baixado.';
  } catch (e) {
    status.textContent = 'Erro: ' + e.message;
  }
  btn.disabled = false;
};
</script>
</body>
</html>
"""

def normalize(text):
    if not text or not isinstance(text, str):
        return ""
    text = unicodedata.normalize("NFD", text)
    text = "".join(c for c in text if unicodedata.category(c) != "Mn")
    text = text.lower()
    text = re.sub(r"[^a-z0-9\s]", " ", text)
    stops = {"prefeitura", "municipal", "municipio", "de", "do", "da", "dos", "das",
             "ltda", "me", "eireli", "sa", "s", "a", "empresa", "clinica", "consultorio", "dr", "dra"}
    words = [w for w in text.split() if w not in stops and len(w) > 1]
    return " ".join(words).strip()

def clean_cnpj(val):
    if pd.isna(val):
        return None
    s = re.sub(r"\D", "", str(val))
    if len(s) == 13:
        s = "0" + s  # Excel às vezes remove o zero à esquerda
    return s if len(s) == 14 else None

def fetch_cnpj(cnpj):
    # Tenta BrasilAPI
    try:
        r = requests.get(f"https://brasilapi.com.br/api/cnpj/v1/{cnpj}", timeout=15)
        if r.status_code == 200:
            d = r.json()
            return {
                "razao_social": d.get("razao_social", ""),
                "nome_fantasia": d.get("nome_fantasia", ""),
                "situacao": d.get("descricao_situacao_cadastral", ""),
                "cnae": d.get("cnae_fiscal_descricao", ""),
                "municipio": d.get("municipio", ""),
                "uf": d.get("uf", ""),
                "abertura": d.get("data_inicio_atividade", ""),
                "fonte": "BrasilAPI"
            }
    except Exception:
        pass

    # Fallback OpenCNPJ
    try:
        r = requests.get(f"https://api.opencnpj.org/{cnpj}", timeout=15)
        if r.status_code == 200:
            d = r.json()
            return {
                "razao_social": d.get("razao_social") or d.get("razaoSocial") or "",
                "nome_fantasia": d.get("nome_fantasia") or d.get("nomeFantasia") or "",
                "situacao": d.get("situacao_cadastral") or d.get("situacaoCadastral") or "",
                "cnae": d.get("cnae_principal") or d.get("cnaePrincipal") or "",
                "municipio": d.get("municipio") or "",
                "uf": d.get("uf") or "",
                "abertura": d.get("data_inicio_atividade") or d.get("dataInicioAtividades") or "",
                "fonte": "OpenCNPJ"
            }
    except Exception:
        pass

    # Fallback CNPJá Open
    try:
        r = requests.get(f"https://open.cnpja.com/office/{cnpj}", timeout=15)
        if r.status_code == 200:
            d = r.json()
            return {
                "razao_social": d.get("companyName") or "",
                "nome_fantasia": d.get("alias") or "",
                "situacao": (d.get("status") or {}).get("text") or "",
                "cnae": (d.get("mainActivity") or {}).get("text") or "",
                "municipio": ((d.get("headquarters") or {}).get("address") or {}).get("city") or "",
                "uf": ((d.get("headquarters") or {}).get("address") or {}).get("state") or "",
                "abertura": d.get("foundedDate") or "",
                "fonte": "CNPJá"
            }
    except Exception:
        pass

    return None

def search_mongo(nome, razao):
    n_nome = normalize(nome)
    n_razao = normalize(razao)
    queries = []
    if n_nome:
        queries.append({"nomeFantasiaNorm": n_nome})
        queries.append({"razaoSocialNorm": n_nome})
    if n_razao and n_razao != n_nome:
        queries.append({"razaoSocialNorm": n_razao})
        queries.append({"nomeFantasiaNorm": n_razao})
    if not queries:
        return None, "sem nome"
    # tenta match exato normalizado
    for q in queries:
        doc = coll.find_one(q)
        if doc:
            return doc, "exato"
    # fallback parcial (primeiras palavras)
    for term in [n_nome, n_razao]:
        if not term or len(term) < 6:
            continue
        parts = term.split()[:3]
        pattern = ".*".join(re.escape(p) for p in parts)
        doc = coll.find_one({
            "$or": [
                {"nomeFantasiaNorm": {"$regex": pattern, "$options": "i"}},
                {"razaoSocialNorm": {"$regex": pattern, "$options": "i"}}
            ]
        })
        if doc:
            return doc, "parcial"
    return None, "não encontrado"

def google_text_search(query):
    url = "https://places.googleapis.com/v1/places:searchText"
    headers = {
        "Content-Type": "application/json",
        "X-Goog-Api-Key": GOOGLE_API_KEY,
        "X-Goog-FieldMask": "places.id,places.displayName,places.formattedAddress,places.location,places.rating,places.userRatingCount,places.regularOpeningHours,places.reviews,places.photos"
    }
    payload = {
        "textQuery": query,
        "languageCode": "pt-BR",
        "maxResultCount": 1
    }
    try:
        r = requests.post(url, headers=headers, json=payload, timeout=15)
        if r.status_code == 200:
            data = r.json()
            places = data.get("places", [])
            return places[0] if places else None
    except Exception:
        pass
    return None

def google_nearby_transit(lat, lng):
    url = "https://places.googleapis.com/v1/places:searchNearby"
    headers = {
        "Content-Type": "application/json",
        "X-Goog-Api-Key": GOOGLE_API_KEY,
        "X-Goog-FieldMask": "places.displayName,places.primaryType,places.types"
    }
    counts = {"metro_trem": 0, "onibus": 0}
    # Metrô/trem: 1000m
    payload_metro = {
        "includedTypes": ["subway_station", "train_station"],
        "maxResultCount": 20,
        "locationRestriction": {
            "circle": {
                "center": {"latitude": lat, "longitude": lng},
                "radius": 1000.0
            }
        }
    }
    # Ônibus: 500m
    payload_bus = {
        "includedTypes": ["bus_station"],
        "maxResultCount": 20,
        "locationRestriction": {
            "circle": {
                "center": {"latitude": lat, "longitude": lng},
                "radius": 500.0
            }
        }
    }
    try:
        r = requests.post(url, headers=headers, json=payload_metro, timeout=12)
        if r.status_code == 200:
            counts["metro_trem"] = len(r.json().get("places", []))
    except Exception:
        pass
    time.sleep(0.2)
    try:
        r = requests.post(url, headers=headers, json=payload_bus, timeout=12)
        if r.status_code == 200:
            counts["onibus"] = len(r.json().get("places", []))
    except Exception:
        pass
    return counts

import base64

def get_photo_base64(photo_names, max_photos=3):
    images = []
    for name in photo_names[:max_photos]:
        if not name:
            continue
        media_name = name + "/media"
        url = f"https://places.googleapis.com/v1/{media_name}?maxWidthPx=800&key={GOOGLE_API_KEY}"
        try:
            r = requests.get(url, timeout=15, allow_redirects=True)
            if r.status_code == 200 and r.content and len(r.content) > 1000:
                # Detecta tipo
                content_type = r.headers.get("Content-Type", "image/jpeg")
                if "jpeg" in content_type or "jpg" in content_type:
                    media_type = "image/jpeg"
                elif "png" in content_type:
                    media_type = "image/png"
                elif "webp" in content_type:
                    media_type = "image/webp"
                else:
                    media_type = "image/jpeg"
                b64 = base64.standard_b64encode(r.content).decode("utf-8")
                images.append({"media_type": media_type, "data": b64})
        except Exception:
            pass
        time.sleep(0.2)
    return images

def claude_analyze(reviews_text, photo_images, endereco="", metro=0, onibus=0):
    if not reviews_text and not photo_images and not endereco:
        return {"claude_analise_completa": "", "claude_status": "sem dados", "scores_ia": {}}
    
    content = []
    for img in photo_images[:3]:
        content.append({
            "type": "image",
            "source": {
                "type": "base64",
                "media_type": img["media_type"],
                "data": img["data"]
            }
        })
    
    prompt = "Você é um avaliador de clínicas odontológicas. Analise os dados abaixo.\n\n"
    if reviews_text:
        prompt += f"Comentários:\n{reviews_text}\n\n"
    if photo_images:
        prompt += f"Foram enviadas {len(photo_images)} foto(s).\n\n"
    else:
        prompt += "Não há fotos.\n\n"
    if endereco:
        prompt += f"Endereço da clínica: {endereco}\n"
        prompt += f"Estações de metrô/trem próximas (até 1 km): {metro}\n"
        prompt += f"Pontos de ônibus próximos (até 500 m): {onibus}\n\n"
    prompt += """Retorne APENAS um JSON válido com estas chaves (valores de 0 a 1, onde 1 = excelente, 0.5 = médio se não der para avaliar, 0 = ruim):
{
  "localizacao": 0.0,
  "sala_espera": 0.0,
  "estacionamento": 0.0,
  "acesso_cadeirantes": 0.0,
  "banheiro": 0.0,
  "consultorio_limpo": 0.0,
  "comentarios_positivos": 0.0,
  "fotos_qualidade": 0.0
}

Para "localizacao", considere se o endereço parece de fácil acesso e a quantidade de transporte público próximo.
Depois do JSON, em uma linha separada, escreva uma breve análise em português."""
    
    content.append({"type": "text", "text": prompt})
    
    payload = {
        "model": "claude-haiku-5-5",
        "max_tokens": 700,
        "messages": [{"role": "user", "content": content}]
    }
    headers = {
        "x-api-key": CLAUDE_API_KEY,
        "anthropic-version": "2023-06-01",
        "content-type": "application/json"
    }
    try:
        r = requests.post("https://api.anthropic.com/v1/messages", headers=headers, json=payload, timeout=60)
        if r.status_code == 200:
            data = r.json()
            text = ""
            # Haiku 5.5 pode retornar vários blocos; pega o primeiro que tiver "text"
            for block in data.get("content", []):
                if isinstance(block, dict) and block.get("type") == "text" and "text" in block:
                    text = block["text"]
                    break
            if not text:
                # fallback: tenta o formato antigo
                try:
                    text = data["content"][0]["text"]
                except Exception:
                    text = str(data)[:500]
            scores_ia = {}
            import json
            try:
                start = text.find("{")
                end = text.rfind("}") + 1
                if start >= 0 and end > start:
                    scores_ia = json.loads(text[start:end])
            except Exception:
                pass
            return {
                "claude_analise_completa": text,
                "claude_status": "ok",
                "fotos_enviadas": len(photo_images),
                "scores_ia": scores_ia
            }
        else:
            return {"claude_status": f"erro {r.status_code}: {r.text[:150]}", "scores_ia": {}}
    except Exception as e:
        return {"claude_status": f"erro: {str(e)[:100]}", "scores_ia": {}}

def calculate_score(row):
    """Calcula score 0-100 com base nos critérios ativos."""
    scores = {}
    total = 0
    
    # 14. CNES (10 pts)
    has_cnes = bool(row.get("CNES"))
    scores["CNES"] = 10 if has_cnes else 0
    total += scores["CNES"]
    
    # 6. Nota do Google (15 pts) — proporcional
    rating = row.get("google_rating")
    if rating and isinstance(rating, (int, float)):
        scores["Nota Google"] = round((rating / 5.0) * 15, 1)
    else:
        scores["Nota Google"] = 7.5  # média se não tiver
    total += scores["Nota Google"]
    
    # 9. Acesso transporte público (13 pts)
    metro = row.get("estacoes_metro_trem") or 0
    onibus = row.get("pontos_onibus") or 0
    transit_score = min(1.0, (metro * 0.6 + onibus * 0.2))
    scores["Transporte"] = round(transit_score * 13, 1)
    total += scores["Transporte"]
    
    # 12. Horário de atendimento (5 pts)
    has_hours = bool(row.get("horarios"))
    scores["Horário"] = 5 if has_hours else 2.5
    total += scores["Horário"]
    
    # 13. Fotos da clínica (5 pts)
    fotos = row.get("fotos_count") or 0
    scores["Fotos"] = 5 if fotos >= 3 else (3 if fotos > 0 else 2.5)
    total += scores["Fotos"]
    
    # Critérios avaliados pela IA (ou média se não houver)
    ia = row.get("scores_ia") or {}
    def ia_or_avg(key, max_pts):
        val = ia.get(key)
        if val is None:
            return round(max_pts * 0.5, 1)  # média
        try:
            return round(float(val) * max_pts, 1)
        except Exception:
            return round(max_pts * 0.5, 1)
    
    scores["Sala espera"] = ia_or_avg("sala_espera", 8)
    scores["Estacionamento"] = ia_or_avg("estacionamento", 8)
    scores["Acesso cadeirantes"] = ia_or_avg("acesso_cadeirantes", 10)
    scores["Banheiro"] = ia_or_avg("banheiro", 4)
    scores["Consultório limpo"] = ia_or_avg("consultorio_limpo", 12)
    
    for k in ["Sala espera", "Estacionamento", "Acesso cadeirantes", "Banheiro", "Consultório limpo"]:
        total += scores[k]
    
    # 1. Localização (10 pts) — IA avalia facilidade de acesso
    scores["Localização"] = ia_or_avg("localizacao", 10)
    total += scores["Localização"]
    
    return round(min(100, total), 1), scores

def google_enrich(nome, endereco, cidade, uf):
    if not nome and not endereco:
        return {}
    query = f"{nome or ''} {endereco or ''} {cidade or ''} {uf or ''}".strip()
    place = google_text_search(query)
    if not place:
        return {"google_status": "não encontrado"}
    
    result = {
        "google_nome": place.get("displayName", {}).get("text", ""),
        "google_endereco": place.get("formattedAddress", ""),
        "google_rating": place.get("rating"),
        "google_avaliacoes": place.get("userRatingCount"),
        "google_status": "encontrado"
    }
    
    hours = place.get("regularOpeningHours", {})
    if hours and "weekdayDescriptions" in hours:
        result["horarios"] = " | ".join(hours["weekdayDescriptions"])
    else:
        result["horarios"] = ""
    
    reviews = place.get("reviews", [])
    if reviews:
        textos = []
        for rev in reviews[:3]:
            t = rev.get("text", {}).get("text", "")
            if t:
                textos.append(t[:200])
        result["comentarios"] = " || ".join(textos)
    else:
        result["comentarios"] = ""
    
    photos = place.get("photos", [])
    if photos:
        result["fotos_count"] = len(photos)
        result["foto_refs"] = ", ".join([p.get("name", "") for p in photos[:3]])
    else:
        result["fotos_count"] = 0
        result["foto_refs"] = ""
    
    loc = place.get("location", {})
    if loc.get("latitude") and loc.get("longitude"):
        transit = google_nearby_transit(loc["latitude"], loc["longitude"])
        result["estacoes_metro_trem"] = transit["metro_trem"]
        result["pontos_onibus"] = transit["onibus"]
    else:
        result["estacoes_metro_trem"] = None
        result["pontos_onibus"] = None
    
    # Análise Claude (comentários + fotos + localização)
    photo_names = [p.get("name") for p in photos if p.get("name")] if photos else []
    photo_images = get_photo_base64(photo_names, max_photos=3)
    endereco_full = f"{endereco or ''} {cidade or ''} {uf or ''}".strip()
    metro = result.get("estacoes_metro_trem") or 0
    onibus = result.get("pontos_onibus") or 0
    claude = claude_analyze(result.get("comentarios", ""), photo_images, endereco_full, metro, onibus)
    result.update(claude)
    
    return result

@app.route("/")
def index():
    return render_template_string(HTML)

@app.route("/process", methods=["POST"])
def process():
    try:
        return _process()
    except Exception as e:
        return f"Erro interno: {str(e)[:200]}", 500

def _process():
    if "file" not in request.files:
        return "Nenhum arquivo", 400
    f = request.files["file"]
    try:
        df = pd.read_excel(f)
    except Exception as e:
        return f"Erro ao ler Excel: {e}", 400

    # acha coluna CNPJ
    cnpj_col = None
    for col in df.columns:
        if re.search(r"cnpj", str(col), re.I):
            cnpj_col = col
            break
    if cnpj_col is None:
        # tenta primeira coluna com muitos CNPJs
        for col in df.columns:
            sample = df[col].head(30).apply(clean_cnpj).notna().sum()
            if sample >= 5:
                cnpj_col = col
                break
    if cnpj_col is None:
        return "Não encontrei coluna de CNPJ", 400

    results = []
    cnpjs = df[cnpj_col].apply(clean_cnpj).dropna().unique()

    for i, cnpj in enumerate(cnpjs):
        try:
            api = fetch_cnpj(cnpj)
            time.sleep(0.35)
            row = {"CNPJ": cnpj}
            if api:
                row.update({
                    "Razão Social (API)": api["razao_social"],
                    "Nome Fantasia (API)": api["nome_fantasia"],
                    "Situação": api["situacao"],
                    "CNAE": api["cnae"],
                    "Cidade API": api["municipio"],
                    "UF API": api["uf"],
                    "Abertura": api["abertura"],
                    "Fonte API": api.get("fonte", ""),
                })
                doc, match = search_mongo(api["nome_fantasia"], api["razao_social"])
            else:
                row["Erro API"] = "Não encontrado"
                doc, match = None, "sem API"
            if doc:
                row.update({
                    "CNES": doc.get("cnes"),
                    "Nome Fantasia (CNES)": doc.get("nomeFantasia"),
                    "Razão Social (CNES)": doc.get("razaoSocial"),
                    "Endereço": doc.get("endereco"),
                    "CEP": doc.get("cep"),
                    "Telefone": doc.get("telefone"),
                    "Match": match,
                })
                g = google_enrich(
                    doc.get("nomeFantasia") or (api.get("nome_fantasia") if api else None),
                    doc.get("endereco"),
                    api.get("municipio") if api else None,
                    api.get("uf") if api else None
                )
                row.update(g)
                time.sleep(0.3)
            else:
                row["CNES"] = None
                row["Match"] = match
            
            score, breakdown = calculate_score(row)
            row["Score"] = score
            row["Detalhe Score"] = str(breakdown)
            # converte dicts aninhados para string (Excel não aceita)
            if isinstance(row.get("scores_ia"), dict):
                row["scores_ia"] = str(row["scores_ia"])
            results.append(row)
        except Exception as e:
            results.append({"CNPJ": cnpj, "Erro": str(e)[:150]})

    out = pd.DataFrame(results)
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        out.to_excel(writer, index=False, sheet_name="Resultados")
    buffer.seek(0)
    return send_file(
        buffer,
        as_attachment=True,
        download_name=f"prestadores_enriquecidos_{datetime.now().strftime('%Y%m%d')}.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    print(f"Acesse http://127.0.0.1:{port}")
    app.run(host="0.0.0.0", port=port, debug=False)
