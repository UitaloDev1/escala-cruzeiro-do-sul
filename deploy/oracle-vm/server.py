#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Backend de Autenticação e Solicitações de Plantão — Escala Médica (Hospital Cruzeiro do Sul)
Roda atrás do Nginx (auth_request + proxy /api/) usando apenas biblioteca padrão do Python.

Perfis:
- medico: acessa a escala (/), visualiza preenchimento/filtros, solicita plantões e desistências.
- diretor: acessa o /dashboard executivo, visualiza horas efetuadas/pendentes e aprova/recusa solicitações.
"""

import argparse
import base64
import copy
import hashlib
import hmac
import json
import os
import re
import secrets
import sqlite3
import threading
import time
from datetime import datetime, timedelta
from http import cookies
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import parse_qs, urlparse
from zoneinfo import ZoneInfo

SCALE_ID_RE = re.compile(r"^[a-z0-9-]{1,80}$")
ISO_RE = re.compile(r"^\d{4}-\d{2}-\d{2}$")
TIME_RE = re.compile(r"^\d{2}:\d{2}$")
CRM_RE = re.compile(r"^\d{4,10}$")
UF_RE = re.compile(r"^[A-Z]{2}$")

TZ_SP = ZoneInfo("America/Sao_Paulo")
COOKIE_NAME = "escala_session"
SESSION_TTL_SECONDS = 12 * 3600
MAX_LOGIN_ATTEMPTS = 5
LOCK_WINDOW_SECONDS = 300
DUMMY_SALT = b"escala_hapvida_16"

CONFIG_DIR = Path(os.environ.get("ESCALA_CONFIG_DIR", "/etc/escala-cruzeiro-do-sul"))
USERS_FILE = CONFIG_DIR / "users.json"
KEY_FILE = CONFIG_DIR / "session.key"
DATA_DIR = Path(os.environ.get("ESCALA_DATA_DIR", "/var/lib/escala-cruzeiro-do-sul"))
DB_FILE = DATA_DIR / "escala.db"
WEB_ROOT = Path(os.environ.get("ESCALA_WEB_ROOT", "/var/www/escala-cruzeiro-do-sul"))
JSON_DATA_FILE = WEB_ROOT / "dados_escala.json"

_RATE_LOCK = threading.Lock()
_DB_LOCK = threading.Lock()
_FAILED_LOGINS: dict[str, list[float]] = {}
_SIGNING_KEY: bytes = b""


def b64u_encode(data: bytes) -> str:
    return base64.urlsafe_b64encode(data).rstrip(b"=").decode("ascii")


def b64u_decode(text: str) -> bytes:
    pad = "=" * (-len(text) % 4)
    return base64.urlsafe_b64decode((text + pad).encode("ascii"))


def format_doctor_short(full_name: str) -> str:
    parts = [p for p in str(full_name).split() if p]
    if len(parts) <= 2:
        return str(full_name).title()
    return f"{parts[0].capitalize()} {parts[-1].capitalize()}"


def ensure_key() -> bytes:
    CONFIG_DIR.mkdir(parents=True, exist_ok=True)
    os.chmod(CONFIG_DIR, 0o700)
    if not KEY_FILE.exists():
        KEY_FILE.write_bytes(secrets.token_bytes(32))
        os.chmod(KEY_FILE, 0o600)
    return KEY_FILE.read_bytes()


def init_db() -> None:
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    with sqlite3.connect(DB_FILE) as conn:
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS shift_requests (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                scale_id TEXT NOT NULL,
                scale_title TEXT NOT NULL,
                iso TEXT NOT NULL,
                ddmm TEXT NOT NULL,
                inicio TEXT NOT NULL,
                fim TEXT NOT NULL,
                horas REAL NOT NULL,
                periodo TEXT NOT NULL,
                tipo TEXT NOT NULL CHECK(tipo IN ('solicitar', 'desistir')),
                medico_nome TEXT NOT NULL,
                medico_crm TEXT NOT NULL DEFAULT '',
                medico_uf TEXT NOT NULL DEFAULT '',
                motivo TEXT NOT NULL DEFAULT '',
                status TEXT NOT NULL DEFAULT 'pending' CHECK(status IN ('pending', 'approved', 'rejected')),
                requested_by TEXT NOT NULL,
                created_at TEXT NOT NULL,
                decided_by TEXT NOT NULL DEFAULT '',
                decided_at TEXT NOT NULL DEFAULT '',
                director_note TEXT NOT NULL DEFAULT ''
            )
            """
        )
        conn.commit()


def load_users() -> dict:
    if not USERS_FILE.exists():
        return {}
    return json.loads(USERS_FILE.read_text(encoding="utf-8"))


def verify_user(username: str, password: str) -> dict | None:
    users = load_users()
    record = users.get(username)
    if not isinstance(record, dict) or "salt" not in record or "hash" not in record:
        hashlib.scrypt(password.encode("utf-8"), salt=DUMMY_SALT, n=16384, r=8, p=1, dklen=32)
        return None
    salt_bytes = b64u_decode(record["salt"])
    expected = b64u_decode(record["hash"])
    derived = hashlib.scrypt(password.encode("utf-8"), salt=salt_bytes, n=16384, r=8, p=1, dklen=32)
    if not hmac.compare_digest(derived, expected):
        return None
    return {
        "username": username,
        "role": record.get("role", "medico"),
        "displayName": record.get("displayName", username),
    }


def create_session_token(user_info: dict) -> str:
    payload = json.dumps(
        {
            "u": user_info["username"],
            "r": user_info["role"],
            "n": user_info["displayName"],
            "exp": int(time.time()) + SESSION_TTL_SECONDS,
        },
        separators=(",", ":"),
    ).encode("utf-8")
    p_b64 = b64u_encode(payload)
    sig = hmac.new(_SIGNING_KEY, p_b64.encode("ascii"), hashlib.sha256).hexdigest()
    return f"{p_b64}.{sig}"


def verify_session_token(token: str) -> dict | None:
    if not token or "." not in token:
        return None
    p_b64, sig = token.split(".", 1)
    expected_sig = hmac.new(_SIGNING_KEY, p_b64.encode("ascii"), hashlib.sha256).hexdigest()
    if not hmac.compare_digest(sig, expected_sig):
        return None
    try:
        data = json.loads(b64u_decode(p_b64).decode("utf-8"))
    except Exception:
        return None
    if int(data.get("exp", 0)) < int(time.time()):
        return None
    return {
        "username": data.get("u", ""),
        "role": data.get("r", "medico"),
        "displayName": data.get("n", data.get("u", "")),
    }


def is_rate_limited(ip: str) -> bool:
    now = time.time()
    with _RATE_LOCK:
        attempts = [t for t in _FAILED_LOGINS.get(ip, []) if now - t < LOCK_WINDOW_SECONDS]
        _FAILED_LOGINS[ip] = attempts
        return len(attempts) >= MAX_LOGIN_ATTEMPTS


def record_failed_login(ip: str) -> None:
    now = time.time()
    with _RATE_LOCK:
        attempts = [t for t in _FAILED_LOGINS.get(ip, []) if now - t < LOCK_WINDOW_SECONDS]
        attempts.append(now)
        _FAILED_LOGINS[ip] = attempts


def clear_failed_login(ip: str) -> None:
    with _RATE_LOCK:
        _FAILED_LOGINS.pop(ip, None)


def row_to_dict(row: sqlite3.Row) -> dict:
    return {
        "id": row["id"],
        "scaleId": row["scale_id"],
        "scaleTitle": row["scale_title"],
        "iso": row["iso"],
        "ddmm": row["ddmm"],
        "inicio": row["inicio"],
        "fim": row["fim"],
        "horas": row["horas"],
        "periodo": row["periodo"],
        "tipo": row["tipo"],
        "medicoNome": row["medico_nome"],
        "medicoCrm": row["medico_crm"],
        "medicoUf": row["medico_uf"],
        "motivo": row["motivo"],
        "status": row["status"],
        "requestedBy": row["requested_by"],
        "createdAt": row["created_at"],
        "decidedBy": row["decided_by"],
        "decidedAt": row["decided_at"],
        "directorNote": row["director_note"],
    }


def fetch_requests(status_filter: str | None = None) -> list[dict]:
    with _DB_LOCK, sqlite3.connect(DB_FILE) as conn:
        conn.row_factory = sqlite3.Row
        if status_filter:
            rows = conn.execute(
                "SELECT * FROM shift_requests WHERE status = ? ORDER BY id DESC",
                (status_filter,),
            ).fetchall()
        else:
            rows = conn.execute("SELECT * FROM shift_requests ORDER BY id DESC").fetchall()
        return [row_to_dict(r) for r in rows]


def load_base_scale() -> dict:
    if not JSON_DATA_FILE.exists():
        return {"hospital": "HOSPITAL CRUZEIRO DO SUL", "municipio": "SÃO PAULO", "escalas": []}
    return json.loads(JSON_DATA_FILE.read_text(encoding="utf-8"))


def recalculate_scale(sc: dict) -> None:
    """Recalcula KPIs, lista de médicos e plantões a partir das semanas já com overrides."""
    base_catalog = {}
    for m in sc.get("medicos", []):
        base_catalog[m["nome"].upper()] = {
            "nome": m["nome"],
            "nomeCurto": m.get("nomeCurto") or format_doctor_short(m["nome"]),
            "crm": m.get("crm", ""),
            "uf": m.get("uf", ""),
        }

    doc_stats = {}
    all_shifts = []
    total_slots = 0
    filled_slots = 0
    open_slots = 0
    total_hours = 0.0

    for wk in sc.get("semanas", []):
        for dia in wk.get("dias", []):
            in_month = bool(dia.get("inMonth", True))
            for grp in dia.get("grupos", []):
                meds = grp.get("medicos", [])
                vagas_abertas = int(grp.get("vagasAbertas", 0))
                grp["totalVagas"] = len(meds) + vagas_abertas
                hrs = float(grp.get("horas", 0))
                periodo = grp.get("periodo", "manha")

                if in_month:
                    total_slots += grp["totalVagas"]
                    filled_slots += len(meds)
                    open_slots += vagas_abertas
                    total_hours += len(meds) * hrs

                for m in meds:
                    key = m["nome"].upper()
                    if key not in base_catalog:
                        base_catalog[key] = {
                            "nome": m["nome"],
                            "nomeCurto": m.get("nomeCurto") or format_doctor_short(m["nome"]),
                            "crm": m.get("crm", ""),
                            "uf": m.get("uf", ""),
                        }
                    if key not in doc_stats:
                        doc_stats[key] = {
                            **base_catalog[key],
                            "totalHoras": 0.0,
                            "totalPlantoes": 0,
                            "manha": 0,
                            "tarde": 0,
                            "noite": 0,
                        }
                    doc_stats[key]["totalHoras"] = round(doc_stats[key]["totalHoras"] + hrs, 2)
                    doc_stats[key]["totalPlantoes"] += 1
                    if periodo in ("manha", "tarde", "noite"):
                        doc_stats[key][periodo] += 1

                    all_shifts.append(
                        {
                            "iso": dia["iso"],
                            "ddmm": dia["ddmm"],
                            "weekdayFull": dia["weekdayFull"],
                            "weekdayShort": dia["weekdayShort"],
                            "inicio": grp["inicio"],
                            "fim": grp["fim"],
                            "horas": hrs,
                            "periodo": periodo,
                            "medico": m["nome"],
                            "medicoCurto": m.get("nomeCurto") or format_doctor_short(m["nome"]),
                            "crm": m.get("crm", ""),
                            "uf": m.get("uf", ""),
                        }
                    )

    for key, info in base_catalog.items():
        if key not in doc_stats:
            doc_stats[key] = {
                **info,
                "totalHoras": 0.0,
                "totalPlantoes": 0,
                "manha": 0,
                "tarde": 0,
                "noite": 0,
            }

    medicos_list = sorted(doc_stats.values(), key=lambda x: (-x["totalHoras"], x["nome"]))
    medicos_ativos = sum(1 for m in medicos_list if m["totalPlantoes"] > 0)
    cobertura = round((filled_slots / total_slots) * 100, 1) if total_slots > 0 else 0.0

    sc["medicos"] = medicos_list
    sc["plantoes"] = sorted(all_shifts, key=lambda s: (s["iso"], s["inicio"], s["medico"]))
    sc["kpis"] = {
        "totalVagas": total_slots,
        "preenchidas": filled_slots,
        "abertas": open_slots,
        "coberturaPct": cobertura,
        "medicosEscalados": medicos_ativos,
        "medicosCadastrados": len(medicos_list),
        "totalHoras": round(total_hours, 1),
    }


def get_merged_scale() -> dict:
    payload = copy.deepcopy(load_base_scale())
    approved = sorted(fetch_requests(status_filter="approved"), key=lambda r: r["id"])
    if not approved:
        return payload

    scales_by_id = {sc["id"]: sc for sc in payload.get("escalas", [])}
    touched = set()

    for req in approved:
        sc = scales_by_id.get(req["scaleId"])
        if not sc:
            continue
        catalog = {m["nome"].upper(): m for m in sc.get("medicos", [])}

        for wk in sc.get("semanas", []):
            for dia in wk.get("dias", []):
                if dia.get("iso") != req["iso"]:
                    continue
                for grp in dia.get("grupos", []):
                    if grp.get("inicio") != req["inicio"] or grp.get("fim") != req["fim"]:
                        continue
                    target_key = req["medicoNome"].strip().upper()
                    meds = grp.get("medicos", [])
                    idx = next((i for i, m in enumerate(meds) if m["nome"].strip().upper() == target_key), None)

                    if req["tipo"] == "solicitar" and idx is None:
                        info = catalog.get(
                            target_key,
                            {
                                "nome": req["medicoNome"].strip(),
                                "nomeCurto": format_doctor_short(req["medicoNome"]),
                                "crm": req["medicoCrm"],
                                "uf": req["medicoUf"],
                            },
                        )
                        meds.append(
                            {
                                "nome": info["nome"],
                                "nomeCurto": info.get("nomeCurto") or format_doctor_short(info["nome"]),
                                "crm": info.get("crm", ""),
                                "uf": info.get("uf", ""),
                            }
                        )
                        if grp.get("vagasAbertas", 0) > 0:
                            grp["vagasAbertas"] -= 1
                        touched.add(sc["id"])
                    elif req["tipo"] == "desistir" and idx is not None:
                        meds.pop(idx)
                        grp["vagasAbertas"] = int(grp.get("vagasAbertas", 0)) + 1
                        touched.add(sc["id"])

    for sid in touched:
        recalculate_scale(scales_by_id[sid])

    return payload


def shift_end_dt(iso: str, inicio: str, fim: str) -> datetime:
    y, m, d = map(int, iso.split("-"))
    hi, mi = map(int, inicio.split(":"))
    hf, mf = map(int, fim.split(":"))
    start_dt = datetime(y, m, d, hi, mi, tzinfo=TZ_SP)
    end_dt = datetime(y, m, d, hf, mf, tzinfo=TZ_SP)
    if end_dt <= start_dt:
        end_dt += timedelta(days=1)
    return end_dt


def compute_director_overview() -> dict:
    payload = get_merged_scale()
    requests_list = fetch_requests()
    now_sp = datetime.now(TZ_SP)

    scales_summary = []
    tot_vagas = 0
    tot_preenchidas = 0
    tot_abertas = 0
    tot_horas_confirmadas = 0.0
    tot_horas_efetuadas = 0.0
    tot_horas_futuras = 0.0
    tot_horas_pendentes_medico = 0.0
    tot_medicos_set = set()
    tot_cadastrados_set = set()

    for sc in payload.get("escalas", []):
        horas_efetuadas = 0.0
        horas_futuras = 0.0
        horas_pendentes = 0.0
        doc_hours = {}

        for m in sc.get("medicos", []):
            tot_cadastrados_set.add(m["nome"].upper())
            if m.get("totalPlantoes", 0) > 0:
                tot_medicos_set.add(m["nome"].upper())
            doc_hours[m["nome"].upper()] = {
                "nome": m["nome"],
                "crm": m.get("crm", ""),
                "uf": m.get("uf", ""),
                "manha": m.get("manha", 0),
                "tarde": m.get("tarde", 0),
                "noite": m.get("noite", 0),
                "totalPlantoes": m.get("totalPlantoes", 0),
                "totalHoras": float(m.get("totalHoras", 0.0)),
                "horasEfetuadas": 0.0,
                "horasFuturas": 0.0,
            }

        for wk in sc.get("semanas", []):
            for dia in wk.get("dias", []):
                in_month = bool(dia.get("inMonth", True))
                for grp in dia.get("grupos", []):
                    hrs = float(grp.get("horas", 0.0))
                    end_dt = shift_end_dt(dia["iso"], grp["inicio"], grp["fim"])
                    is_done = end_dt <= now_sp
                    meds = grp.get("medicos", [])
                    vagas = int(grp.get("vagasAbertas", 0))

                    if in_month:
                        horas_pendentes += vagas * hrs
                        if is_done:
                            horas_efetuadas += len(meds) * hrs
                        else:
                            horas_futuras += len(meds) * hrs

                    for m in meds:
                        d_entry = doc_hours.get(m["nome"].upper())
                        if d_entry:
                            if is_done:
                                d_entry["horasEfetuadas"] = round(d_entry["horasEfetuadas"] + hrs, 1)
                            else:
                                d_entry["horasFuturas"] = round(d_entry["horasFuturas"] + hrs, 1)

        k = sc["kpis"]
        tot_vagas += k["totalVagas"]
        tot_preenchidas += k["preenchidas"]
        tot_abertas += k["abertas"]
        tot_horas_confirmadas += k["totalHoras"]
        tot_horas_efetuadas += horas_efetuadas
        tot_horas_futuras += horas_futuras
        tot_horas_pendentes_medico += horas_pendentes

        meds_sorted = sorted(doc_hours.values(), key=lambda x: (-x["totalHoras"], x["nome"]))
        scales_summary.append(
            {
                "id": sc["id"],
                "titulo": sc["titulo"],
                "tabLabel": sc["tabLabel"],
                "mesAnoLabel": sc["mesAnoLabel"],
                "kpis": {
                    **k,
                    "horasEfetuadas": round(horas_efetuadas, 1),
                    "horasFuturas": round(horas_futuras, 1),
                    "horasPendentesMedico": round(horas_pendentes, 1),
                },
                "medicos": meds_sorted,
                "plantoes": sc.get("plantoes", []),
            }
        )

    pending_count = sum(1 for r in requests_list if r["status"] == "pending")
    approved_count = sum(1 for r in requests_list if r["status"] == "approved")
    rejected_count = sum(1 for r in requests_list if r["status"] == "rejected")
    cobertura_geral = round((tot_preenchidas / tot_vagas) * 100, 1) if tot_vagas > 0 else 0.0

    return {
        "hospital": payload.get("hospital", "HOSPITAL CRUZEIRO DO SUL"),
        "municipio": payload.get("municipio", "SÃO PAULO"),
        "referenciaAgora": now_sp.strftime("%d/%m/%Y às %H:%M"),
        "geradoEm": payload.get("geradoEm", ""),
        "planilhaModificadaEm": payload.get("planilhaModificadaEm", ""),
        "consolidado": {
            "totalVagas": tot_vagas,
            "preenchidas": tot_preenchidas,
            "abertas": tot_abertas,
            "coberturaPct": cobertura_geral,
            "medicosEscalados": len(tot_medicos_set),
            "medicosCadastrados": len(tot_cadastrados_set),
            "totalHoras": round(tot_horas_confirmadas, 1),
            "horasEfetuadas": round(tot_horas_efetuadas, 1),
            "horasFuturas": round(tot_horas_futuras, 1),
            "horasPendentesMedico": round(tot_horas_pendentes_medico, 1),
            "solicitacoesPendentes": pending_count,
            "solicitacoesAprovadas": approved_count,
            "solicitacoesRecusadas": rejected_count,
        },
        "escalas": scales_summary,
        "solicitacoes": requests_list,
    }


class EscalaHandler(BaseHTTPRequestHandler):
    server_version = "EscalaAuth/1.0"

    def log_message(self, format, *args):
        pass

    def client_ip(self) -> str:
        # X-Real-IP é sobrescrito pelo Nginx ($remote_addr) e não pode ser falsificado pelo cliente
        real_ip = (self.headers.get("X-Real-IP") or "").strip()
        return real_ip or self.client_address[0]

    def is_https(self) -> bool:
        return self.headers.get("X-Forwarded-Proto", "").lower() == "https"

    def check_same_origin(self) -> bool:
        origin = (self.headers.get("Origin") or "").strip()
        if not origin:
            return True
        parsed = urlparse(origin)
        host = (self.headers.get("X-Forwarded-Host") or self.headers.get("Host") or "").split(",")[0].strip().lower()
        return bool(host and parsed.netloc.lower() == host)

    def current_session(self) -> dict | None:
        raw_cookie = self.headers.get("Cookie", "")
        if not raw_cookie:
            return None
        jar = cookies.SimpleCookie()
        try:
            jar.load(raw_cookie)
        except Exception:
            return None
        morsel = jar.get(COOKIE_NAME)
        if not morsel:
            return None
        return verify_session_token(morsel.value)

    def read_json_body(self) -> dict | None:
        try:
            length = int(self.headers.get("Content-Length", "0"))
        except ValueError:
            return None
        if length <= 0 or length > 65536:
            return None
        raw = self.rfile.read(length)
        try:
            return json.loads(raw.decode("utf-8"))
        except Exception:
            return None

    def send_json(self, code: int, data: dict | list, extra_headers: dict | None = None) -> None:
        body = json.dumps(data, ensure_ascii=False).encode("utf-8")
        self.send_response(code)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-store")
        if extra_headers:
            for k, v in extra_headers.items():
                self.send_header(k, v)
        self.end_headers()
        self.wfile.write(body)

    def send_empty(self, code: int, extra_headers: dict | None = None) -> None:
        self.send_response(code)
        self.send_header("Content-Length", "0")
        self.send_header("Cache-Control", "no-store")
        if extra_headers:
            for k, v in extra_headers.items():
                self.send_header(k, v)
        self.end_headers()

    def build_cookie_header(self, value: str, max_age: int) -> str:
        parts = [
            f"{COOKIE_NAME}={value}",
            "Path=/",
            f"Max-Age={max_age}",
            "HttpOnly",
            "SameSite=Lax",
        ]
        if self.is_https():
            parts.append("Secure")
        return "; ".join(parts)

    def do_GET(self):
        path = urlparse(self.path).path

        if path == "/_auth/check":
            sess = self.current_session()
            if not sess:
                self.send_empty(401)
                return
            self.send_empty(200, {"X-Auth-User": sess["username"], "X-Auth-Role": sess["role"]})
            return

        if path == "/_auth/check-director":
            sess = self.current_session()
            if not sess:
                self.send_empty(401)
                return
            if sess["role"] != "diretor":
                self.send_empty(403)
                return
            self.send_empty(200, {"X-Auth-User": sess["username"], "X-Auth-Role": sess["role"]})
            return

        if path == "/_auth/session":
            sess = self.current_session()
            if not sess:
                self.send_json(200, {"authenticated": False})
                return
            self.send_json(200, {"authenticated": True, **sess})
            return

        if path in ("/_auth/logout", "/logout"):
            self.send_response(302)
            self.send_header("Set-Cookie", self.build_cookie_header("", 0))
            self.send_header("Location", "/login")
            self.send_header("Cache-Control", "no-store")
            self.end_headers()
            return

        if path == "/api/escala":
            sess = self.current_session()
            if not sess:
                self.send_json(401, {"error": "Não autenticado"})
                return
            self.send_json(200, get_merged_scale())
            return

        if path == "/api/requests":
            sess = self.current_session()
            if not sess:
                self.send_json(401, {"error": "Não autenticado"})
                return
            all_reqs = fetch_requests()
            if sess["role"] == "diretor":
                self.send_json(200, {"solicitacoes": all_reqs})
                return
            qs = parse_qs(urlparse(self.path).query)
            q_nome = " ".join((qs.get("nome", [""])[0]).strip().split()).upper()
            q_crm = re.sub(r"\D", "", qs.get("crm", [""])[0])
            if not q_nome and not q_crm:
                self.send_json(200, {"solicitacoes": []})
                return
            filtered = [
                r for r in all_reqs
                if (q_nome and r["medicoNome"].upper() == q_nome)
                or (q_crm and re.sub(r"\D", "", r.get("medicoCrm", "")) == q_crm)
            ]
            self.send_json(200, {"solicitacoes": filtered})
            return

        if path == "/api/director/overview":
            sess = self.current_session()
            if not sess:
                self.send_json(401, {"error": "Não autenticado"})
                return
            if sess["role"] != "diretor":
                self.send_json(403, {"error": "Acesso restrito ao Diretor"})
                return
            self.send_json(200, compute_director_overview())
            return

        self.send_empty(404)

    def do_POST(self):
        path = urlparse(self.path).path

        if not self.check_same_origin():
            self.send_json(403, {"error": "Origem não autorizada."})
            return

        if path == "/_auth/login":
            ip = self.client_ip()
            if is_rate_limited(ip):
                self.send_json(429, {"error": "Muitas tentativas incorretas. Aguarde 5 minutos."})
                return
            body = self.read_json_body()
            if not body:
                self.send_json(400, {"error": "Requisição inválida."})
                return
            username = str(body.get("username", "")).strip().lower()
            password = str(body.get("password", ""))
            user_info = verify_user(username, password)
            if not user_info:
                record_failed_login(ip)
                self.send_json(401, {"error": "Usuário ou senha inválidos."})
                return
            clear_failed_login(ip)
            token = create_session_token(user_info)
            redirect_to = "/dashboard" if user_info["role"] == "diretor" else "/"
            self.send_json(
                200,
                {"ok": True, "role": user_info["role"], "displayName": user_info["displayName"], "redirect": redirect_to},
                {"Set-Cookie": self.build_cookie_header(token, SESSION_TTL_SECONDS)},
            )
            return

        if path == "/_auth/logout":
            self.send_json(
                200,
                {"ok": True},
                {"Set-Cookie": self.build_cookie_header("", 0)},
            )
            return

        if path == "/api/requests":
            sess = self.current_session()
            if not sess:
                self.send_json(401, {"error": "Não autenticado"})
                return
            body = self.read_json_body()
            if not body:
                self.send_json(400, {"error": "Corpo JSON inválido."})
                return

            tipo = str(body.get("tipo", "")).strip()
            scale_id = str(body.get("scaleId", "")).strip()
            scale_title = str(body.get("scaleTitle", "")).strip()
            iso = str(body.get("iso", "")).strip()
            ddmm = str(body.get("ddmm", "")).strip()
            inicio = str(body.get("inicio", "")).strip()
            fim = str(body.get("fim", "")).strip()
            periodo = str(body.get("periodo", "manha")).strip()
            medico_nome = " ".join(str(body.get("medicoNome", "")).strip().split())
            medico_crm = str(body.get("medicoCrm", "")).strip()
            medico_uf = str(body.get("medicoUf", "")).strip().upper()
            motivo = str(body.get("motivo", "")).strip()[:400]
            try:
                horas = float(body.get("horas", 0))
            except (TypeError, ValueError):
                horas = 0.0

            if tipo not in ("solicitar", "desistir") or not SCALE_ID_RE.match(scale_id) or not ISO_RE.match(iso) or not TIME_RE.match(inicio) or not TIME_RE.match(fim) or not medico_nome:
                self.send_json(400, {"error": "Preencha nome completo, CRM, escala, data e horário válidos."})
                return
            if medico_crm:
                crm_digits = re.sub(r"\D", "", medico_crm)
                if not CRM_RE.match(crm_digits):
                    self.send_json(400, {"error": "CRM inválido."})
                    return
                medico_crm = crm_digits
            if medico_uf and not UF_RE.match(medico_uf):
                self.send_json(400, {"error": "UF inválida."})
                return

            # Valida contra a escala atual (vaga aberta para solicitar / plantão escalado para desistir)
            merged = get_merged_scale()
            sc = next((s for s in merged.get("escalas", []) if s["id"] == scale_id), None)
            if not sc:
                self.send_json(404, {"error": "Escala não encontrada."})
                return
            slot_grp = None
            for wk in sc.get("semanas", []):
                for dia in wk.get("dias", []):
                    if dia.get("iso") == iso:
                        for g in dia.get("grupos", []):
                            if g.get("inicio") == inicio and g.get("fim") == fim:
                                slot_grp = g
                                break
            if not slot_grp:
                self.send_json(404, {"error": "Faixa de horário não encontrada nesta data."})
                return

            med_upper = medico_nome.upper()
            already_in = any(m["nome"].strip().upper() == med_upper for m in slot_grp.get("medicos", []))
            if tipo == "desistir" and not already_in:
                self.send_json(400, {"error": "Não há plantão escalado para este médico nesta data e horário."})
                return
            if tipo == "solicitar":
                if already_in:
                    self.send_json(400, {"error": "Este médico já está escalado nesta data e horário."})
                    return
                if int(slot_grp.get("vagasAbertas", 0)) <= 0:
                    self.send_json(400, {"error": "Não há mais vagas em aberto nesta data e horário."})
                    return

            now_str = datetime.now(TZ_SP).strftime("%d/%m/%Y %H:%M")
            with _DB_LOCK, sqlite3.connect(DB_FILE) as conn:
                conn.row_factory = sqlite3.Row
                dup = conn.execute(
                    """
                    SELECT id FROM shift_requests
                    WHERE status = 'pending'
                      AND scale_id = ? AND iso = ? AND inicio = ? AND fim = ?
                      AND tipo = ? AND UPPER(medico_nome) = UPPER(?)
                    """,
                    (scale_id, iso, inicio, fim, tipo, medico_nome),
                ).fetchone()
                if dup:
                    self.send_json(409, {"error": "Já existe uma solicitação pendente idêntica para este plantão."})
                    return

                cur = conn.execute(
                    """
                    INSERT INTO shift_requests (
                        scale_id, scale_title, iso, ddmm, inicio, fim, horas, periodo,
                        tipo, medico_nome, medico_crm, medico_uf, motivo, status,
                        requested_by, created_at
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 'pending', ?, ?)
                    """,
                    (
                        scale_id,
                        scale_title,
                        iso,
                        ddmm,
                        inicio,
                        fim,
                        horas,
                        periodo,
                        tipo,
                        medico_nome,
                        medico_crm,
                        medico_uf,
                        motivo,
                        sess["username"],
                        now_str,
                    ),
                )
                conn.commit()
                new_id = cur.lastrowid
                row = conn.execute("SELECT * FROM shift_requests WHERE id = ?", (new_id,)).fetchone()
            self.send_json(201, {"ok": True, "solicitacao": row_to_dict(row)})
            return

        if path.startswith("/api/director/requests/") and path.endswith("/decision"):
            sess = self.current_session()
            if not sess:
                self.send_json(401, {"error": "Não autenticado"})
                return
            if sess["role"] != "diretor":
                self.send_json(403, {"error": "Acesso restrito ao Diretor"})
                return

            parts = [p for p in path.split("/") if p]
            # ['api', 'director', 'requests', '<id>', 'decision']
            if len(parts) != 5 or not parts[3].isdigit():
                self.send_json(400, {"error": "ID de solicitação inválido."})
                return
            req_id = int(parts[3])

            body = self.read_json_body()
            if not body:
                self.send_json(400, {"error": "Corpo JSON inválido."})
                return
            action = str(body.get("action", "")).strip()
            note = str(body.get("note", "")).strip()[:400]
            if action not in ("approve", "reject"):
                self.send_json(400, {"error": "Ação deve ser 'approve' ou 'reject'."})
                return

            new_status = "approved" if action == "approve" else "rejected"
            now_str = datetime.now(TZ_SP).strftime("%d/%m/%Y %H:%M")

            with _DB_LOCK, sqlite3.connect(DB_FILE) as conn:
                conn.row_factory = sqlite3.Row
                existing = conn.execute("SELECT * FROM shift_requests WHERE id = ?", (req_id,)).fetchone()
                if not existing:
                    self.send_json(404, {"error": "Solicitação não encontrada."})
                    return
                if existing["status"] != "pending":
                    self.send_json(409, {"error": "Esta solicitação já foi decidida."})
                    return

                conn.execute(
                    """
                    UPDATE shift_requests
                    SET status = ?, decided_by = ?, decided_at = ?, director_note = ?
                    WHERE id = ?
                    """,
                    (new_status, sess["username"], now_str, note, req_id),
                )
                conn.commit()
                updated = conn.execute("SELECT * FROM shift_requests WHERE id = ?", (req_id,)).fetchone()
            self.send_json(200, {"ok": True, "solicitacao": row_to_dict(updated)})
            return

        if path == "/api/director/shifts/remove":
            sess = self.current_session()
            if not sess:
                self.send_json(401, {"error": "Não autenticado"})
                return
            if sess["role"] != "diretor":
                self.send_json(403, {"error": "Acesso restrito ao Diretor"})
                return
            body = self.read_json_body()
            if not body:
                self.send_json(400, {"error": "Corpo JSON inválido."})
                return

            scale_id = str(body.get("scaleId", "")).strip()
            iso = str(body.get("iso", "")).strip()
            inicio = str(body.get("inicio", "")).strip()
            fim = str(body.get("fim", "")).strip()
            medico_nome = " ".join(str(body.get("medicoNome", "")).strip().split())
            note = str(body.get("note", "")).strip()[:400] or "Retirado do plantão pela Diretoria"

            if not SCALE_ID_RE.match(scale_id) or not ISO_RE.match(iso) or not TIME_RE.match(inicio) or not TIME_RE.match(fim) or not medico_nome:
                self.send_json(400, {"error": "Dados inválidos para remoção de plantão."})
                return

            merged = get_merged_scale()
            sc = next((s for s in merged.get("escalas", []) if s["id"] == scale_id), None)
            if not sc:
                self.send_json(404, {"error": "Escala não encontrada."})
                return

            found_day = None
            found_grp = None
            found_med = None
            for wk in sc.get("semanas", []):
                for dia in wk.get("dias", []):
                    if dia.get("iso") == iso:
                        for g in dia.get("grupos", []):
                            if g.get("inicio") == inicio and g.get("fim") == fim:
                                m = next((x for x in g.get("medicos", []) if x["nome"].strip().upper() == medico_nome.upper()), None)
                                if m:
                                    found_day = dia
                                    found_grp = g
                                    found_med = m
                                    break

            if not found_med:
                self.send_json(404, {"error": "Médico não está escalado neste plantão."})
                return

            now_str = datetime.now(TZ_SP).strftime("%d/%m/%Y %H:%M")
            with _DB_LOCK, sqlite3.connect(DB_FILE) as conn:
                conn.row_factory = sqlite3.Row
                conn.execute(
                    """
                    UPDATE shift_requests
                    SET status = 'rejected', decided_by = ?, decided_at = ?, director_note = ?
                    WHERE status = 'pending' AND tipo = 'desistir'
                      AND scale_id = ? AND iso = ? AND inicio = ? AND fim = ?
                      AND UPPER(medico_nome) = UPPER(?)
                    """,
                    (sess["username"], now_str, "Encerrado após remoção direta pela Diretoria", scale_id, iso, inicio, fim, found_med["nome"]),
                )
                cur = conn.execute(
                    """
                    INSERT INTO shift_requests (
                        scale_id, scale_title, iso, ddmm, inicio, fim, horas, periodo,
                        tipo, medico_nome, medico_crm, medico_uf, motivo, status,
                        requested_by, created_at, decided_by, decided_at, director_note
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, 'desistir', ?, ?, ?, ?, 'approved', ?, ?, ?, ?, ?)
                    """,
                    (
                        sc["id"],
                        sc["tabLabel"],
                        found_day["iso"],
                        found_day["ddmm"],
                        found_grp["inicio"],
                        found_grp["fim"],
                        float(found_grp["horas"]),
                        found_grp["periodo"],
                        found_med["nome"],
                        found_med.get("crm", ""),
                        found_med.get("uf", ""),
                        "Remoção direta pela Diretoria",
                        sess["username"],
                        now_str,
                        sess["username"],
                        now_str,
                        note,
                    ),
                )
                conn.commit()
                row = conn.execute("SELECT * FROM shift_requests WHERE id = ?", (cur.lastrowid,)).fetchone()
            self.send_json(200, {"ok": True, "solicitacao": row_to_dict(row)})
            return

        self.send_empty(404)


def main():
    global _SIGNING_KEY
    parser = argparse.ArgumentParser(description="Backend de Autenticação e Solicitações - Escala Médica")
    parser.add_argument("--host", default="127.0.0.1")
    parser.add_argument("--port", type=int, default=8091)
    args = parser.parse_args()

    _SIGNING_KEY = ensure_key()
    init_db()
    server = ThreadingHTTPServer((args.host, args.port), EscalaHandler)
    print(f"[ESCALA-SERVER] Escutando em http://{args.host}:{args.port}")
    server.serve_forever()


if __name__ == "__main__":
    main()
