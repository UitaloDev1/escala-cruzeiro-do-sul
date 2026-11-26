#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Gerador & Sincronizador Automático de Escala Médica (HTML)
Hospital Cruzeiro do Sul — Hapvida (Design System Torre de Inteligência Hospitalar)

Funcionalidades:
- Lê a planilha oficial de escalas (.xlsx) diretamente do OneDrive local (via cópia segura,
  sem conflitar com o Excel aberto).
- Extrai todas as abas válidas (ignora templates vazios como NUM_ESCALA_NOME_ESCALA).
- Aplica filtro LGPD estrito: NUNCA exporta CPF, CNPJ ou valores financeiros.
- Gera um arquivo index.html autônomo (Dark Mode institucional Hapvida + modo claro,
  filtros por médico/turno/vagas, exportação .ICS para Google Agenda/iPhone e impressão).
- Suporta modo --watch (monitoramento contínuo) e --git-push (publicação automática no GitHub Pages).

Requisitos:
    pip install openpyxl

Uso rápido:
    python atualizar_escala.py
    python atualizar_escala.py --watch --git-push
"""

import argparse
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
import time
from datetime import datetime, date, time as dt_time

try:
    import openpyxl
except ImportError:
    print("[ERRO] Biblioteca 'openpyxl' não encontrada. Instale com: pip install openpyxl")
    sys.exit(1)

# ==============================================================================
# CONFIGURAÇÕES PADRÃO (AJUSTE CONFORME SEU COMPUTADOR)
# ==============================================================================
DEFAULT_EXCEL_PATH = r"C:\Users\Computador\OneDrive - HAPVIDA ASSISTÊNCIA MÉDICA LTDA\Escalas_Hospital_Cruzeiro_Do_Sul_11_26.xlsx"
DEFAULT_OUTPUT_DIR = os.path.dirname(os.path.abspath(__file__))
DEFAULT_WATCH_INTERVAL = 15  # segundos entre verificações no modo --watch
IGNORED_SHEETS = {
    "NUM_ESCALA_NOME_ESCALA",
    "MODELO",
    "TEMPLATE",
    "BASE",
    "GINECOLOGIA_OBSTETRICIA_11_26",
}

DIAS_SEMANA_FULL = [
    "Segunda-feira", "Terça-feira", "Quarta-feira",
    "Quinta-feira", "Sexta-feira", "Sábado", "Domingo"
]
DIAS_SEMANA_CURTO = ["SEG", "TER", "QUA", "QUI", "SEX", "SÁB", "DOM"]
MESES_PT = [
    "", "Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho",
    "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"
]

HAPVIDA_SVG_LOGO = '<svg viewBox="0 0 195 43" width="142" height="31" class="brand-logo" aria-label="Hapvida" xmlns="http://www.w3.org/2000/svg" xmlns:xlink="http://www.w3.org/1999/xlink" xml:space="preserve" overflow="hidden"><g transform="translate(-1024 -635)"><path d="M1025.15 649.234C1023.13 652.725 1025.52 656.944 1031.55 659.268 1034.71 660.485 1040 660.734 1043.44 659.807 1042.52 656.355 1039.67 651.907 1037.04 649.772 1032.02 645.703 1027.16 645.742 1025.15 649.234Z" fill="#FF4D00"/><path d="M1045.79 635C1041.41 635 1038.75 639.528 1039.85 646.461 1040.43 650.091 1043.05 655.188 1045.79 657.921 1048.54 655.188 1051.17 650.091 1051.74 646.461 1052.84 639.538 1050.16 635 1045.79 635Z" fill="#F32735"/><path d="M1066.45 649.244C1064.44 645.753 1059.59 645.703 1054.56 649.772 1051.93 651.907 1049.07 656.345 1048.15 659.797 1051.59 660.734 1056.88 660.485 1060.04 659.268 1066.07 656.954 1068.47 652.735 1066.45 649.244Z" fill="#FF4D00"/><path d="M1045.79 677.701C1048.49 677.701 1050.13 674.908 1049.45 670.649 1049.11 668.415 1047.48 665.273 1045.79 663.597 1044.11 665.283 1042.49 668.415 1042.13 670.649 1041.45 674.908 1043.1 677.701 1045.79 677.701Z" fill="#FFC62B"/><path d="M1031.23 669.382C1032.58 671.716 1035.82 671.746 1039.17 669.033 1040.93 667.607 1042.84 664.644 1043.45 662.34 1041.15 661.722 1037.62 661.881 1035.51 662.689 1031.48 664.235 1029.88 667.048 1031.23 669.382Z" fill="#FF8200"/><path d="M1060.36 669.382C1061.7 667.048 1060.11 664.235 1056.09 662.679 1053.97 661.871 1050.45 661.702 1048.15 662.32 1048.76 664.624 1050.67 667.597 1052.42 669.023 1055.77 671.736 1059.01 671.716 1060.36 669.382Z" fill="#FF8200"/><path d="M1210.98 657.153 1210.09 663.278C1209.75 665.532 1208.18 667.058 1206.27 667.058 1204.06 667.058 1203.3 665.831 1203.3 663.447 1203.3 662.599 1203.39 661.532 1203.56 660.305 1204.45 654.52 1206.02 653.373 1208.32 653.373 1210.35 653.373 1211.29 655.158 1210.99 657.153ZM1188.12 657.153 1187.23 663.278C1186.89 665.532 1185.32 667.058 1183.41 667.058 1181.2 667.058 1180.44 665.831 1180.44 663.447 1180.44 662.599 1180.53 661.532 1180.7 660.305 1181.6 654.52 1183.16 653.373 1185.46 653.373 1187.49 653.373 1188.43 655.158 1188.14 657.153M1111.49 657.153 1110.6 663.278C1110.26 665.532 1108.68 667.058 1106.78 667.058 1104.56 667.058 1103.81 665.831 1103.81 663.447 1103.81 662.599 1103.9 661.532 1104.06 660.305 1104.95 654.52 1106.53 653.373 1108.82 653.373 1110.86 653.373 1111.79 655.158 1111.51 657.153M1134.31 660.086C1133.42 665.871 1131.84 667.018 1129.55 667.018 1127.51 667.018 1126.58 665.233 1126.87 663.238L1127.76 657.113C1128.1 654.859 1129.68 653.333 1131.58 653.333 1133.8 653.333 1134.56 654.56 1134.56 656.944 1134.56 657.792 1134.47 658.859 1134.3 660.086M1166.52 649.204 1163.21 671.177C1163.51 671.267 1164.53 671.477 1166.32 671.477 1168.1 671.477 1169.29 671.217 1169.63 671.177L1172.94 649.204C1172.56 649.114 1171.49 648.904 1169.8 648.904 1167.88 648.904 1166.82 649.164 1166.52 649.204ZM1161.64 648.914C1159.81 648.914 1158.45 649.174 1158.11 649.254L1154.87 656.565C1153.56 659.577 1152.41 661.243 1151.34 665.572L1151.25 665.572C1151.16 662.38 1150.66 658.6 1150.11 654.989L1149.64 651.927C1149.34 649.892 1148.19 648.904 1146.16 648.904 1144.8 648.904 1143.52 649.074 1142.72 649.204L1146.37 671.138C1147.34 671.347 1148.79 671.517 1150.41 671.517 1152.19 671.517 1153.77 671.347 1154.58 671.138L1164.65 649.204C1164.31 649.114 1163.2 648.904 1161.64 648.904M1215.53 648.914C1213.61 648.914 1212.34 649.852 1211.71 651.378 1210.77 649.762 1208.69 648.535 1206.44 648.535 1200.5 648.535 1198.19 653.164 1197.22 659.667 1197.01 660.984 1196.88 662.43 1196.88 663.836 1196.88 668.764 1199 671.866 1203.63 671.866 1205.88 671.866 1208.18 670.759 1209.37 669.232 1209.7 670.639 1210.98 671.487 1212.34 671.487 1213.4 671.487 1214.42 671.397 1215.23 671.187L1218.54 649.214C1217.65 649.044 1216.55 648.914 1215.52 648.914M1134.23 648.525C1131.97 648.525 1129.68 649.633 1128.49 651.159 1128.15 649.752 1126.88 648.904 1125.52 648.904 1124.45 648.904 1123.44 648.994 1122.63 649.204L1118.34 677.701C1119.24 677.87 1120.33 678 1121.36 678 1123.82 678 1124.76 677.152 1125.14 674.688L1125.99 669.103C1126.88 670.679 1128.88 671.866 1131.42 671.866 1137.37 671.866 1139.67 667.238 1140.65 660.734 1140.86 659.418 1140.99 657.971 1140.99 656.565 1140.99 651.637 1138.87 648.535 1134.24 648.535M1116.03 648.914C1114.12 648.914 1112.84 649.852 1112.21 651.378 1111.28 649.762 1109.2 648.535 1106.95 648.535 1101 648.535 1098.7 653.164 1097.72 659.667 1097.51 660.984 1097.38 662.43 1097.38 663.836 1097.38 668.764 1099.51 671.866 1104.13 671.866 1106.39 671.866 1108.68 670.759 1109.87 669.232 1110.21 670.639 1111.49 671.487 1112.84 671.487 1113.91 671.487 1114.93 671.397 1115.73 671.187L1119.05 649.214C1118.15 649.044 1117.05 648.914 1116.02 648.914ZM1190.02 644.536 1189.01 651.288C1188.12 649.712 1186.11 648.525 1183.57 648.525 1177.63 648.525 1175.32 653.154 1174.34 659.657 1174.14 660.974 1174.01 662.42 1174.01 663.826 1174.01 668.754 1176.13 671.856 1180.76 671.856 1183.01 671.856 1185.31 670.749 1186.49 669.222 1186.83 670.629 1188.11 671.477 1189.47 671.477 1190.53 671.477 1191.55 671.387 1192.36 671.177L1196.82 641.513C1195.93 641.344 1194.82 641.214 1193.8 641.214 1191.33 641.214 1190.39 642.062 1190.01 644.526M1171.03 639.479C1168.95 639.479 1167.25 641.055 1167.03 643.049 1166.82 645.084 1168.27 646.7 1170.27 646.7 1172.26 646.7 1174.01 645.084 1174.22 643.049 1174.42 641.055 1172.99 639.479 1171.02 639.479Z" fill="#FFFFFF"/><path d="M1091.01 641.513 1089.19 653.533 1079.75 653.533 1079.75 653.453C1079.75 653.453 1081.08 644.675 1081.08 644.675 1081.33 643.01 1080.05 641.513 1078.37 641.513L1075.16 641.513 1070.69 671.177C1070.69 671.177 1072.16 671.347 1073.76 671.347 1075.68 671.347 1077.09 671.177 1077.09 671.177L1078.93 659.029 1078.93 658.939 1088.38 658.939 1087.02 668.026C1086.77 669.691 1088.06 671.187 1089.73 671.187L1092.94 671.187 1097.41 641.523C1097.41 641.523 1095.96 641.354 1094.25 641.354 1092.42 641.354 1091.01 641.523 1091.01 641.523Z" fill="#FFFFFF"/></g></svg>'


# ==============================================================================
# UTILITÁRIOS DE PARSING DA PLANILHA
# ==============================================================================
def clean_str(val):
    if val is None:
        return ""
    return " ".join(str(val).strip().split())


def format_sheet_title(sheet_name):
    """Transforma nomes de aba como 'GO PORTA_12_ 26' em 'GO Porta'."""
    name = re.sub(r'[_\-\s]*\d{1,2}[_\-\s]+\d{2,4}\s*$', '', sheet_name).strip()
    name = name.replace("_", " ")
    name = " ".join(name.split())
    if not name:
        name = sheet_name
    # Preserva siglas comuns
    words = name.split()
    formatted = []
    acronyms = {"GO", "UTI", "PA", "PS", "CC", "OBST", "GIN", "CLIN", "PED", "ORTO", "CIR"}
    preps = {"DE", "DA", "DO", "DAS", "DOS", "E"}
    for idx, w in enumerate(words):
        wu = w.upper()
        if wu in acronyms:
            formatted.append(wu)
        elif idx > 0 and wu in preps:
            formatted.append(wu.lower())
        else:
            formatted.append(w.capitalize())
    return " ".join(formatted)


def format_doctor_short(full_name):
    """Retorna Primeiro + Último sobrenome para exibição compacta na grade."""
    parts = [p for p in full_name.split() if p]
    if len(parts) <= 2:
        return full_name.title()
    return f"{parts[0].capitalize()} {parts[-1].capitalize()}"


def parse_time_val(val):
    if isinstance(val, dt_time):
        return val
    if isinstance(val, datetime):
        return val.time()
    if isinstance(val, str):
        s = val.strip()
        for fmt in ("%H:%M:%S", "%H:%M"):
            try:
                return datetime.strptime(s, fmt).time()
            except ValueError:
                pass
    return None


def calc_hours(t_ini, t_fim):
    m_ini = t_ini.hour * 60 + t_ini.minute
    m_fim = t_fim.hour * 60 + t_fim.minute
    if m_fim <= m_ini:
        m_fim += 24 * 60
    return round((m_fim - m_ini) / 60.0, 2)


def classify_shift(t_ini):
    if t_ini.hour < 12:
        return "manha"
    if t_ini.hour < 18:
        return "tarde"
    return "noite"


def parse_date_ddmm(val, ref_year, ref_month):
    if isinstance(val, datetime):
        return val.date()
    if isinstance(val, date):
        return val
    if isinstance(val, str):
        s = val.strip()
        m = re.match(r'^(\d{1,2})/(\d{1,2})(?:/(\d{2,4}))?$', s)
        if m:
            d, mo = int(m.group(1)), int(m.group(2))
            if m.group(3):
                y = int(m.group(3))
                if y < 100:
                    y += 2000
            else:
                y = ref_year
                if ref_month == 12 and mo == 1:
                    y = ref_year + 1
                elif ref_month == 1 and mo == 12:
                    y = ref_year - 1
            try:
                return date(y, mo, d)
            except ValueError:
                return None
    return None


def parse_workbook(excel_path):
    """Lê o arquivo Excel via cópia temporária segura (não trava com Excel aberto)."""
    if not os.path.exists(excel_path):
        raise FileNotFoundError(f"Planilha não encontrada: {excel_path}")

    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
        tmp_path = tmp.name

    try:
        shutil.copy2(excel_path, tmp_path)
        wb = openpyxl.load_workbook(tmp_path, data_only=True)
        scales = []

        for sheet_name in wb.sheetnames:
            if sheet_name.strip().upper() in IGNORED_SHEETS:
                continue

            ws = wb[sheet_name]
            unidade = clean_str(ws["D1"].value) or "HOSPITAL CRUZEIRO DO SUL"
            municipio = clean_str(ws["F1"].value) or "SÃO PAULO"
            export_ref = clean_str(ws["I1"].value)
            mes_ano_raw = ws["F2"].value

            if isinstance(mes_ano_raw, (datetime, date)):
                ref_year, ref_month = mes_ano_raw.year, mes_ano_raw.month
            else:
                # Fallback pelo nome da aba (ex: _12_26)
                m_sheet = re.search(r'(\d{1,2})[_\-\s]+(\d{2,4})\s*$', sheet_name)
                if m_sheet:
                    ref_month = int(m_sheet.group(1))
                    y_raw = int(m_sheet.group(2))
                    ref_year = y_raw + 2000 if y_raw < 100 else y_raw
                else:
                    now = datetime.now()
                    ref_year, ref_month = now.year, now.month

            # 1) Catálogo de médicos (LGPD: lê APENAS CRM, UF e NOME — ignora CPF, CNPJ e VALOR)
            crm_map = {}
            catalog_start = None
            for r in range(1, ws.max_row + 1):
                c1 = clean_str(ws.cell(r, 1).value).upper()
                c3 = clean_str(ws.cell(r, 3).value).upper()
                if "CRM" in c1 and "NOME" in c3:
                    catalog_start = r + 1
                    break

            if catalog_start:
                for r in range(catalog_start, ws.max_row + 1):
                    nome = clean_str(ws.cell(r, 3).value)
                    if nome:
                        crm = clean_str(ws.cell(r, 1).value)
                        uf = clean_str(ws.cell(r, 2).value).upper()
                        crm_map[nome.upper()] = {
                            "nome": nome,
                            "crm": crm,
                            "uf": uf,
                        }

            # 2) Blocos semanais
            weeks = []
            all_shifts = []
            doc_stats = {}
            total_slots_month = 0
            filled_slots_month = 0
            open_slots_month = 0
            total_hours_month = 0.0

            r = 1
            week_idx = 0
            while r <= ws.max_row:
                c1 = clean_str(ws.cell(r, 1).value).upper()
                c2 = clean_str(ws.cell(r, 2).value).upper()
                if c1 == "HORA INICIO" and c2 == "HORA FIM":
                    week_idx += 1
                    date_row = r + 1
                    col_dates = []
                    for col in range(3, 10):
                        dt = parse_date_ddmm(ws.cell(date_row, col).value, ref_year, ref_month)
                        col_dates.append((col, dt))

                    # Estrutura dos dias dessa semana
                    week_days = {}
                    ordered_days = []
                    for col, dt in col_dates:
                        if not dt:
                            continue
                        iso = dt.isoformat()
                        in_month = (dt.month == ref_month and dt.year == ref_year)
                        day_obj = {
                            "iso": iso,
                            "ddmm": dt.strftime("%d/%m"),
                            "weekdayFull": DIAS_SEMANA_FULL[dt.weekday()],
                            "weekdayShort": DIAS_SEMANA_CURTO[dt.weekday()],
                            "isWeekend": dt.weekday() >= 5,
                            "inMonth": in_month,
                            "timeGroups": {},  # agrupado por "HH:MM-HH:MM"
                        }
                        week_days[col] = day_obj
                        ordered_days.append(day_obj)

                    rr = r + 2
                    while rr <= ws.max_row:
                        t_ini = parse_time_val(ws.cell(rr, 1).value)
                        t_fim = parse_time_val(ws.cell(rr, 2).value)
                        if not t_ini or not t_fim:
                            break

                        s_ini = t_ini.strftime("%H:%M")
                        s_fim = t_fim.strftime("%H:%M")
                        time_key = f"{s_ini}-{s_fim}"
                        hrs = calc_hours(t_ini, t_fim)
                        periodo = classify_shift(t_ini)

                        for col, dt in col_dates:
                            if not dt or col not in week_days:
                                continue
                            day_obj = week_days[col]
                            doc_name = clean_str(ws.cell(rr, col).value)

                            # Só contabiliza dias fora do mês se houver algum plantão preenchido neles
                            if not day_obj["inMonth"] and not doc_name:
                                continue

                            if time_key not in day_obj["timeGroups"]:
                                day_obj["timeGroups"][time_key] = {
                                    "inicio": s_ini,
                                    "fim": s_fim,
                                    "horas": hrs,
                                    "periodo": periodo,
                                    "medicos": [],
                                    "vagasAbertas": 0,
                                    "totalVagas": 0,
                                }

                            grp = day_obj["timeGroups"][time_key]
                            grp["totalVagas"] += 1

                            if day_obj["inMonth"]:
                                total_slots_month += 1

                            if doc_name:
                                doc_key = doc_name.upper()
                                crm_info = crm_map.get(doc_key, {"nome": doc_name, "crm": "", "uf": ""})
                                grp["medicos"].append({
                                    "nome": crm_info["nome"],
                                    "nomeCurto": format_doctor_short(crm_info["nome"]),
                                    "crm": crm_info["crm"],
                                    "uf": crm_info["uf"],
                                })

                                if day_obj["inMonth"]:
                                    filled_slots_month += 1
                                    total_hours_month += hrs

                                if doc_key not in doc_stats:
                                    doc_stats[doc_key] = {
                                        "nome": crm_info["nome"],
                                        "nomeCurto": format_doctor_short(crm_info["nome"]),
                                        "crm": crm_info["crm"],
                                        "uf": crm_info["uf"],
                                        "totalHoras": 0.0,
                                        "totalPlantoes": 0,
                                        "manha": 0,
                                        "tarde": 0,
                                        "noite": 0,
                                    }
                                doc_stats[doc_key]["totalHoras"] = round(doc_stats[doc_key]["totalHoras"] + hrs, 2)
                                doc_stats[doc_key]["totalPlantoes"] += 1
                                doc_stats[doc_key][periodo] += 1

                                all_shifts.append({
                                    "iso": day_obj["iso"],
                                    "ddmm": day_obj["ddmm"],
                                    "weekdayFull": day_obj["weekdayFull"],
                                    "weekdayShort": day_obj["weekdayShort"],
                                    "inicio": s_ini,
                                    "fim": s_fim,
                                    "horas": hrs,
                                    "periodo": periodo,
                                    "medico": crm_info["nome"],
                                    "medicoCurto": format_doctor_short(crm_info["nome"]),
                                    "crm": crm_info["crm"],
                                    "uf": crm_info["uf"],
                                })
                            else:
                                grp["vagasAbertas"] += 1
                                if day_obj["inMonth"]:
                                    open_slots_month += 1

                        rr += 1

                    # Converte timeGroups dict em lista ordenada por horário
                    days_final = []
                    for d_obj in ordered_days:
                        groups_list = sorted(
                            d_obj["timeGroups"].values(),
                            key=lambda g: (g["inicio"], g["fim"])
                        )
                        d_obj["grupos"] = groups_list
                        del d_obj["timeGroups"]
                        days_final.append(d_obj)

                    if days_final:
                        # Intervalo legível da semana (considerando dias do mês)
                        in_month_days = [d for d in days_final if d["inMonth"]]
                        ref_list = in_month_days if in_month_days else days_final
                        weeks.append({
                            "numero": week_idx,
                            "label": f"Semana {week_idx} · {ref_list[0]['ddmm']} a {ref_list[-1]['ddmm']}",
                            "dias": days_final,
                        })
                    r = rr
                else:
                    r += 1

            # Ignora abas completamente sem plantões e sem médicos
            if filled_slots_month == 0 and len(crm_map) == 0:
                continue

            # Inclui também médicos cadastrados na aba com 0 horas para referência no corpo clínico
            for d_key, info in crm_map.items():
                if d_key not in doc_stats:
                    doc_stats[d_key] = {
                        "nome": info["nome"],
                        "nomeCurto": format_doctor_short(info["nome"]),
                        "crm": info["crm"],
                        "uf": info["uf"],
                        "totalHoras": 0.0,
                        "totalPlantoes": 0,
                        "manha": 0,
                        "tarde": 0,
                        "noite": 0,
                    }

            medicos_list = sorted(
                doc_stats.values(),
                key=lambda x: (-x["totalHoras"], x["nome"])
            )
            cobertura_pct = round((filled_slots_month / total_slots_month) * 100, 1) if total_slots_month > 0 else 0.0
            medicos_ativos = sum(1 for m in medicos_list if m["totalPlantoes"] > 0)

            titulo_limpo = format_sheet_title(sheet_name)
            mes_label = f"{MESES_PT[ref_month]}/{ref_year}"
            mes_curto = f"{ref_month:02d}/{str(ref_year)[-2:]}"

            scales.append({
                "id": re.sub(r'[^a-z0-9]+', '-', sheet_name.lower()).strip('-'),
                "sheetName": sheet_name,
                "titulo": titulo_limpo,
                "tabLabel": f"{titulo_limpo.upper()} · {mes_curto}",
                "unidade": unidade,
                "municipio": municipio,
                "mesAnoLabel": mes_label,
                "refYear": ref_year,
                "refMonth": ref_month,
                "exportRef": export_ref,
                "kpis": {
                    "totalVagas": total_slots_month,
                    "preenchidas": filled_slots_month,
                    "abertas": open_slots_month,
                    "coberturaPct": cobertura_pct,
                    "medicosEscalados": medicos_ativos,
                    "medicosCadastrados": len(medicos_list),
                    "totalHoras": round(total_hours_month, 1),
                },
                "semanas": weeks,
                "medicos": medicos_list,
                "plantoes": sorted(all_shifts, key=lambda s: (s["iso"], s["inicio"], s["medico"])),
            })

        # Ordena: mês/ano mais recente primeiro, depois por mais plantões preenchidos
        scales.sort(key=lambda sc: (-sc["refYear"], -sc["refMonth"], -sc["kpis"]["preenchidas"]))

        mtime = datetime.fromtimestamp(os.path.getmtime(excel_path))
        payload = {
            "hospital": scales[0]["unidade"] if scales else "HOSPITAL CRUZEIRO DO SUL",
            "municipio": scales[0]["municipio"] if scales else "SÃO PAULO",
            "geradoEm": datetime.now().strftime("%d/%m/%Y às %H:%M"),
            "planilhaModificadaEm": mtime.strftime("%d/%m/%Y às %H:%M"),
            "arquivoFonte": os.path.basename(excel_path),
            "escalas": scales,
        }
        return payload
    finally:
        if os.path.exists(tmp_path):
            try:
                os.remove(tmp_path)
            except OSError:
                pass


# ==============================================================================
# TEMPLATE HTML (DESIGN SYSTEM TORRE DE INTELIGÊNCIA HOSPITALAR - HAPVIDA)
# ==============================================================================
HTML_TEMPLATE = r"""<!DOCTYPE html>
<html lang="pt-BR" data-theme="dark">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <meta name="robots" content="noindex, nofollow" />
  <title>Escala Médica · __HOSPITAL_TITLE__ — Hapvida</title>
  <link rel="preconnect" href="https://fonts.googleapis.com" />
  <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin />
  <link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&family=Plus+Jakarta+Sans:wght@400;500;600;700&display=swap" rel="stylesheet" />
  <style>
    :root {
      --bg-header: #0044c4;
      --bg-body: #041538;
      --bg-subnav: #061a42;
      --bg-card: #082152;
      --bg-card-inner: #061a42;
      --bg-icon-square: #1a47a3;
      --tag-bg: #0e3b8c;
      --border-subtle: #16336b;
      --border-card: #1c3f82;
      --border-dashed: #3d5e9e;
      --accent-orange: #f37021;
      --accent-yellow: #fdb913;
      --text-primary: #ffffff;
      --text-secondary: #c5d4ee;
      --text-muted: #7d9fdf;
      --text-on-yellow: #041538;
      --shift-manha-bg: rgba(0, 148, 255, 0.14);
      --shift-manha-border: rgba(56, 189, 248, 0.45);
      --shift-manha-text: #7dd3fc;
      --shift-tarde-bg: rgba(243, 112, 33, 0.15);
      --shift-tarde-border: rgba(243, 112, 33, 0.50);
      --shift-tarde-text: #fdba74;
      --shift-noite-bg: rgba(139, 92, 246, 0.16);
      --shift-noite-border: rgba(167, 139, 250, 0.48);
      --shift-noite-text: #c4b5fd;
      --highlight-bg: rgba(243, 112, 33, 0.22);
      --ok-color: #34d399;
      --font-sans: 'Plus Jakarta Sans', 'Inter', system-ui, -apple-system, sans-serif;
    }

    [data-theme="light"] {
      --bg-header: #0044c4;
      --bg-body: #f1f5fb;
      --bg-subnav: #ffffff;
      --bg-card: #ffffff;
      --bg-card-inner: #f5f8fd;
      --bg-icon-square: #dbe7fb;
      --tag-bg: #e5eefb;
      --border-subtle: #d5e1f5;
      --border-card: #c7d7f0;
      --border-dashed: #8da9d8;
      --text-primary: #061a42;
      --text-secondary: #354c77;
      --text-muted: #445f94;
      --text-on-yellow: #041538;
      --shift-manha-bg: rgba(0, 119, 214, 0.10);
      --shift-manha-border: rgba(0, 119, 214, 0.35);
      --shift-manha-text: #0058a3;
      --shift-tarde-bg: rgba(243, 112, 33, 0.12);
      --shift-tarde-border: rgba(243, 112, 33, 0.45);
      --shift-tarde-text: #b84a08;
      --shift-noite-bg: rgba(91, 61, 214, 0.10);
      --shift-noite-border: rgba(91, 61, 214, 0.35);
      --shift-noite-text: #4c2db5;
      --highlight-bg: rgba(243, 112, 33, 0.16);
      --ok-color: #059669;
    }

    * { box-sizing: border-box; }

    ::selection {
      background: rgba(243, 112, 33, 0.32);
      color: #ffffff;
    }

    :focus-visible {
      outline: 2px solid var(--accent-orange);
      outline-offset: 2px;
    }

    html {
      scrollbar-color: var(--border-card) var(--bg-body);
      scrollbar-width: thin;
    }

    body {
      margin: 0;
      background-color: var(--bg-body);
      color: var(--text-primary);
      font-family: var(--font-sans);
      -webkit-font-smoothing: antialiased;
      min-height: 100vh;
    }

    /* ==================== HEADER INSTITUCIONAL ==================== */
    .top-header {
      background-color: var(--bg-header);
      min-height: 56px;
      padding: 10px 24px;
      display: flex;
      align-items: center;
      justify-content: space-between;
      gap: 16px;
      border-bottom: 3px solid transparent;
      border-image: linear-gradient(90deg, var(--accent-orange), var(--accent-yellow)) 1;
      color: #ffffff;
    }

    .header-left {
      display: flex;
      align-items: center;
      gap: 18px;
      min-width: 0;
    }

    .brand-logo {
      flex-shrink: 0;
      display: block;
    }

    .header-divider {
      width: 1px;
      height: 32px;
      background: rgba(255, 255, 255, 0.25);
      flex-shrink: 0;
    }

    .header-titles {
      min-width: 0;
    }

    .header-title {
      font-size: 14.5px;
      font-weight: 600;
      color: #ffffff;
      line-height: 1.25;
      white-space: nowrap;
      overflow: hidden;
      text-overflow: ellipsis;
    }

    .header-sub {
      font-size: 11.5px;
      color: rgba(255, 255, 255, 0.78);
      margin-top: 2px;
    }

    .header-right {
      text-align: right;
      flex-shrink: 0;
    }

    .header-credit-label {
      font-size: 10px;
      color: rgba(255, 255, 255, 0.72);
      text-transform: uppercase;
      letter-spacing: 0.06em;
    }

    .header-credit-name {
      font-size: 12px;
      font-weight: 600;
      color: #ffffff;
      margin-top: 2px;
    }

    /* ==================== SUB-NAVEGAÇÃO ==================== */
    .subnav {
      background-color: var(--bg-subnav);
      border-bottom: 1px solid var(--border-subtle);
      padding: 0 24px;
      min-height: 40px;
      display: flex;
      align-items: center;
      justify-content: space-between;
      gap: 16px;
      position: sticky;
      top: 0;
      z-index: 40;
    }

    .subnav-tabs {
      display: flex;
      align-items: stretch;
      gap: 20px;
      overflow-x: auto;
      scrollbar-width: none;
    }
    .subnav-tabs::-webkit-scrollbar { display: none; }

    .subnav-tab {
      background: transparent;
      border: none;
      border-bottom: 2px solid transparent;
      padding: 10px 0 8px;
      font-family: inherit;
      font-size: 11px;
      font-weight: 600;
      letter-spacing: 0.08em;
      text-transform: uppercase;
      color: var(--text-muted);
      cursor: pointer;
      white-space: nowrap;
      transition: color 0.15s ease, border-color 0.15s ease;
    }

    .subnav-tab:hover { color: var(--text-primary); }

    .subnav-tab.is-active {
      color: var(--text-primary);
      border-bottom-color: var(--accent-orange);
    }

    .subnav-right {
      display: flex;
      align-items: center;
      gap: 14px;
      flex-shrink: 0;
    }

    .subnav-meta {
      font-size: 11px;
      color: var(--text-muted);
      white-space: nowrap;
    }

    .theme-toggle {
      display: inline-flex;
      align-items: center;
      gap: 6px;
      background: var(--tag-bg);
      border: 1px solid var(--border-card);
      color: var(--text-primary);
      border-radius: 999px;
      padding: 4px 10px;
      font-family: inherit;
      font-size: 11px;
      font-weight: 600;
      letter-spacing: 0.04em;
      cursor: pointer;
    }

    /* ==================== CONTAINER & HERO ==================== */
    .container {
      max-width: 1560px;
      margin: 0 auto;
      padding: 20px 24px 44px;
    }

    .hero {
      display: flex;
      justify-content: space-between;
      align-items: flex-end;
      gap: 14px;
      margin-bottom: 14px;
      flex-wrap: wrap;
    }

    .eyebrow {
      font-size: 11px;
      font-weight: 600;
      text-transform: uppercase;
      letter-spacing: 0.08em;
      color: var(--text-muted);
    }

    .hero h1 {
      margin: 3px 0 4px;
      font-size: 24px;
      font-weight: 600;
      line-height: 1.15;
      letter-spacing: -0.02em;
      color: var(--text-primary);
    }

    .hero-desc {
      margin: 0;
      font-size: 12.5px;
      color: var(--text-secondary);
      line-height: 1.45;
      max-width: 640px;
    }

    .sync-pill {
      display: inline-flex;
      align-items: center;
      gap: 8px;
      background: var(--bg-card);
      border: 1px solid var(--border-card);
      border-radius: 2px 2px 10px 2px;
      padding: 6px 12px;
      font-size: 11.5px;
      color: var(--text-secondary);
    }

    .sync-dot {
      width: 7px;
      height: 7px;
      border-radius: 50%;
      background: var(--ok-color);
      flex-shrink: 0;
    }

    /* ==================== GRID 4 COLUNAS & CARDS ==================== */
    .grid-4 {
      display: grid;
      grid-template-columns: repeat(4, 1fr);
      gap: 12px;
    }

    .col-span-2 { grid-column: span 2; }

    .card {
      background-color: var(--bg-card);
      border: 1px solid var(--border-card);
      border-radius: 3px 3px 18px 3px;
      padding: 14px 16px;
      display: flex;
      flex-direction: column;
      transition: border-color 0.2s ease;
    }

    .card:hover { border-color: #2c5cb8; }

    .kpi-top {
      display: flex;
      justify-content: space-between;
      align-items: center;
      gap: 8px;
      margin-bottom: 6px;
    }

    .kpi-value {
      font-size: 26px;
      font-weight: 600;
      line-height: 1.1;
      letter-spacing: -0.01em;
      font-variant-numeric: tabular-nums;
      color: var(--text-primary);
    }

    .kpi-sub {
      font-size: 11.5px;
      color: var(--text-secondary);
      margin-top: 4px;
    }

    .progress-track {
      width: 100%;
      height: 5px;
      background: var(--tag-bg);
      border-radius: 2px;
      overflow: hidden;
      margin-top: 10px;
    }

    .progress-fill {
      height: 100%;
      background: linear-gradient(90deg, var(--accent-orange), var(--accent-yellow));
      transition: width 0.3s ease;
    }

    .badge-yellow {
      background-color: var(--accent-yellow);
      color: var(--text-on-yellow);
      font-size: 11px;
      font-weight: 700;
      letter-spacing: 0.06em;
      text-transform: uppercase;
      padding: 4px 9px;
      border-radius: 3px;
      white-space: nowrap;
    }

    .badge-status {
      font-size: 11px;
      font-weight: 600;
      letter-spacing: 0.06em;
      text-transform: uppercase;
      color: var(--ok-color);
      white-space: nowrap;
    }

    /* ==================== BARRA DE FILTROS ==================== */
    .filter-card {
      margin-top: 12px;
      padding: 12px 16px;
    }

    .filter-row {
      display: grid;
      grid-template-columns: 1.5fr 1fr 1fr auto;
      gap: 10px;
      align-items: end;
    }

    .field-group {
      display: flex;
      flex-direction: column;
      gap: 4px;
    }

    .field-label {
      font-size: 11px;
      font-weight: 600;
      text-transform: uppercase;
      letter-spacing: 0.08em;
      color: var(--text-muted);
    }

    .input-control, .select-control {
      width: 100%;
      height: 36px;
      padding: 0 11px;
      background: var(--bg-card-inner);
      color: var(--text-primary);
      border: 1px solid var(--border-card);
      border-radius: 2px 2px 10px 2px;
      font-family: inherit;
      font-size: 12.5px;
      outline: none;
    }

    .input-control:focus, .select-control:focus {
      border-color: var(--accent-orange);
    }

    .toolbar-bottom {
      margin-top: 10px;
      padding-top: 10px;
      border-top: 1px solid var(--border-subtle);
      display: flex;
      align-items: center;
      justify-content: space-between;
      gap: 10px;
      flex-wrap: wrap;
    }

    .view-switcher {
      display: inline-flex;
      gap: 8px;
      flex-wrap: wrap;
    }

    .btn-outline {
      background: transparent;
      color: var(--text-primary);
      border: 1px solid var(--border-card);
      border-radius: 2px 2px 10px 2px;
      padding: 8px 14px;
      font-family: inherit;
      font-size: 11px;
      font-weight: 600;
      letter-spacing: 0.06em;
      text-transform: uppercase;
      cursor: pointer;
      display: inline-flex;
      align-items: center;
      gap: 7px;
      transition: background-color 0.15s ease, border-color 0.15s ease;
    }

    .btn-outline:hover {
      background-color: rgba(255, 255, 255, 0.07);
      border-color: var(--text-primary);
    }

    [data-theme="light"] .btn-outline:hover {
      background-color: rgba(0, 68, 196, 0.06);
      border-color: var(--bg-header);
    }

    .btn-outline.is-active {
      background: var(--tag-bg);
      border-color: var(--accent-orange);
      color: var(--text-primary);
    }

    .btn-accent {
      background: var(--accent-orange);
      color: #ffffff;
      border: 1px solid var(--accent-orange);
      border-radius: 2px 2px 10px 2px;
      padding: 9px 15px;
      font-family: inherit;
      font-size: 11px;
      font-weight: 700;
      letter-spacing: 0.06em;
      text-transform: uppercase;
      cursor: pointer;
      display: inline-flex;
      align-items: center;
      gap: 8px;
    }
    .btn-accent:hover { filter: brightness(1.08); }

    .checkbox-pill {
      display: inline-flex;
      align-items: center;
      gap: 7px;
      font-size: 12px;
      color: var(--text-secondary);
      cursor: pointer;
      user-select: none;
    }

    .checkbox-pill input {
      accent-color: var(--accent-orange);
      width: 15px;
      height: 15px;
      cursor: pointer;
    }

    /* ==================== CARD PESSOAL DO MÉDICO ==================== */
    .doctor-banner {
      margin-top: 12px;
      border-color: var(--accent-orange);
      display: none;
    }

    .doctor-banner.is-visible { display: flex; }

    .doctor-banner-grid {
      display: flex;
      justify-content: space-between;
      align-items: center;
      gap: 12px;
      flex-wrap: wrap;
    }

    .doctor-banner-name {
      font-size: 17px;
      font-weight: 600;
      margin: 2px 0;
    }

    .doctor-banner-tags {
      display: flex;
      flex-wrap: wrap;
      gap: 6px;
      margin-top: 6px;
    }

    .tech-tag {
      background-color: var(--tag-bg);
      color: var(--text-primary);
      font-size: 11px;
      font-weight: 500;
      padding: 4px 9px;
      border-radius: 2px;
    }

    /* ==================== LEGENDA DE TURNOS ==================== */
    .legend-bar {
      display: flex;
      align-items: center;
      justify-content: space-between;
      gap: 10px;
      margin: 16px 0 10px;
      flex-wrap: wrap;
    }

    .section-title {
      font-size: 16px;
      font-weight: 600;
      margin: 0;
    }

    .legend-items {
      display: flex;
      align-items: center;
      gap: 8px;
      flex-wrap: wrap;
    }

    .shift-chip {
      font-size: 11px;
      font-weight: 600;
      padding: 3px 8px;
      border-radius: 2px;
      border: 1px solid transparent;
    }

    .shift-manha {
      background: var(--shift-manha-bg);
      border-color: var(--shift-manha-border);
      color: var(--shift-manha-text);
    }

    .shift-tarde {
      background: var(--shift-tarde-bg);
      border-color: var(--shift-tarde-border);
      color: var(--shift-tarde-text);
    }

    .shift-noite {
      background: var(--shift-noite-bg);
      border-color: var(--shift-noite-border);
      color: var(--shift-noite-text);
    }

    /* ==================== GRADE SEMANAL (MATRIZ HORIZONTAL) ==================== */
    .week-block {
      margin-bottom: 22px;
    }

    .week-header {
      display: flex;
      align-items: baseline;
      justify-content: space-between;
      gap: 12px;
      flex-wrap: wrap;
      margin-bottom: 8px;
    }

    .week-title {
      font-size: 13px;
      font-weight: 700;
      letter-spacing: 0.05em;
      text-transform: uppercase;
      color: var(--text-primary);
    }

    .week-summary {
      font-size: 12px;
      color: var(--text-muted);
      font-variant-numeric: tabular-nums;
    }

    .matrix-card {
      padding: 0;
      overflow: hidden;
    }

    .matrix-scroll {
      width: 100%;
      overflow-x: auto;
    }

    .roster-matrix {
      width: 100%;
      min-width: 980px;
      border-collapse: separate;
      border-spacing: 0;
      table-layout: fixed;
    }

    .roster-matrix thead th {
      background: var(--bg-subnav);
      border-bottom: 1px solid var(--border-card);
      border-right: 1px solid var(--border-subtle);
      padding: 10px 10px;
      text-align: left;
      vertical-align: middle;
    }

    .roster-matrix thead th:last-child {
      border-right: none;
    }

    .roster-matrix .col-slot {
      width: 138px;
      position: sticky;
      left: 0;
      z-index: 3;
      background: var(--bg-subnav);
      border-right: 1px solid var(--border-card);
    }

    .roster-matrix .col-day.is-today {
      background: rgba(253, 185, 19, 0.10);
      box-shadow: inset 0 2px 0 var(--accent-yellow);
    }

    .roster-matrix .col-day.is-out-month {
      opacity: 0.45;
    }

    .m-day-head {
      display: flex;
      align-items: center;
      justify-content: space-between;
      gap: 6px;
    }

    .m-day-wk {
      font-size: 11px;
      font-weight: 700;
      letter-spacing: 0.07em;
      color: var(--text-muted);
    }

    .m-day-wk.is-weekend {
      color: var(--accent-yellow);
    }

    .m-day-date {
      font-size: 14px;
      font-weight: 700;
      font-variant-numeric: tabular-nums;
      color: var(--text-primary);
      margin-top: 1px;
    }

    .roster-matrix tbody th.col-slot-cell {
      position: sticky;
      left: 0;
      z-index: 2;
      background: var(--bg-subnav);
      border-right: 1px solid var(--border-card);
      border-bottom: 1px solid var(--border-subtle);
      padding: 10px 12px;
      vertical-align: top;
      text-align: left;
      font-weight: 600;
    }

    .roster-matrix tbody td.matrix-cell {
      padding: 10px;
      vertical-align: top;
      border-right: 1px solid var(--border-subtle);
      border-bottom: 1px solid var(--border-subtle);
      background: var(--bg-card);
    }

    .roster-matrix tbody td.matrix-cell:last-child {
      border-right: none;
    }

    .roster-matrix tbody tr:last-child th.col-slot-cell,
    .roster-matrix tbody tr:last-child td.matrix-cell {
      border-bottom: none;
    }

    .roster-matrix tbody td.matrix-cell.is-today {
      background: rgba(253, 185, 19, 0.04);
    }

    .roster-matrix tbody td.matrix-cell.is-out-month {
      opacity: 0.42;
      background: var(--bg-card-inner);
    }

    .slot-cell-wrap {
      display: flex;
      flex-direction: column;
      align-items: flex-start;
      gap: 5px;
    }

    .slot-meta {
      font-size: 11px;
      font-weight: 500;
      color: var(--text-muted);
      font-variant-numeric: tabular-nums;
    }

    .cell-stack {
      display: flex;
      flex-direction: column;
      gap: 5px;
    }

    .doc-pill {
      display: block;
      width: 100%;
      text-align: left;
      font-family: inherit;
      font-size: 12px;
      font-weight: 500;
      color: var(--text-primary);
      padding: 5px 8px;
      background: var(--tag-bg);
      border: 1px solid transparent;
      border-left: 2px solid var(--border-dashed);
      border-radius: 2px 2px 6px 2px;
      line-height: 1.3;
      cursor: pointer;
      transition: border-color 0.15s ease, background-color 0.15s ease;
    }

    .doc-pill:hover {
      border-color: var(--text-muted);
    }

    .doc-pill.is-selected {
      background: var(--highlight-bg);
      border-color: var(--accent-orange);
      border-left: 3px solid var(--accent-orange);
      font-weight: 700;
    }

    .open-hint {
      font-size: 11px;
      font-weight: 500;
      color: var(--text-muted);
      padding: 2px 2px 0;
    }

    .open-pill-empty {
      font-size: 11px;
      font-weight: 500;
      color: var(--text-muted);
      border: 1px dashed var(--border-subtle);
      border-radius: 2px 2px 6px 2px;
      padding: 6px 8px;
      text-align: center;
    }

    /* ==================== ABA DE SOLICITAÇÕES DE PLANTÃO ==================== */
    .req-workspace {
      display: flex;
      flex-direction: column;
      gap: 14px;
    }

    .req-grid-3 {
      display: grid;
      grid-template-columns: 2fr 1fr 0.7fr auto;
      gap: 12px;
      align-items: end;
    }

    .req-grid-2 {
      display: grid;
      grid-template-columns: 1fr 1fr;
      gap: 12px;
      align-items: end;
    }

    .mode-switch {
      display: flex;
      gap: 10px;
      flex-wrap: wrap;
      margin-bottom: 14px;
    }

    .feedback-box {
      border-radius: 2px 2px 10px 2px;
      padding: 10px 12px;
      font-size: 12.5px;
      margin-bottom: 12px;
      display: none;
    }

    .feedback-box.is-error {
      display: block;
      background: rgba(239, 68, 68, 0.15);
      border: 1px solid rgba(248, 113, 113, 0.5);
      color: #fecaca;
    }

    .feedback-box.is-ok {
      display: block;
      background: rgba(52, 211, 153, 0.14);
      border: 1px solid rgba(52, 211, 153, 0.48);
      color: #6ee7b7;
    }

    [data-theme="light"] .feedback-box.is-error { color: #991b1b; }
    [data-theme="light"] .feedback-box.is-ok { color: #065f46; }

    .req-item-card {
      background: var(--bg-card-inner);
      border: 1px solid var(--border-subtle);
      border-radius: 2px 2px 10px 2px;
      padding: 10px 12px;
      margin-bottom: 8px;
    }

    .status-pill {
      display: inline-block;
      font-size: 11px;
      font-weight: 600;
      padding: 2px 7px;
      border-radius: 2px;
    }

    .status-pending { background: rgba(253, 185, 19, 0.16); color: var(--accent-yellow); }
    .status-approved { background: rgba(52, 211, 153, 0.15); color: var(--ok-color); }
    .status-rejected { background: rgba(239, 68, 68, 0.15); color: #fca5a5; }

    /* ==================== LISTA CRONOLÓGICA (MOBILE-FIRST) ==================== */
    .list-view {
      display: flex;
      flex-direction: column;
      gap: 10px;
    }

    .list-day-card {
      background: var(--bg-card);
      border: 1px solid var(--border-card);
      border-radius: 3px 3px 16px 3px;
      padding: 14px 16px;
    }

    .list-day-header {
      display: flex;
      justify-content: space-between;
      align-items: center;
      margin-bottom: 10px;
      padding-bottom: 8px;
      border-bottom: 1px solid var(--border-subtle);
    }

    .list-shift-rows {
      display: grid;
      grid-template-columns: repeat(auto-fill, minmax(250px, 1fr));
      gap: 10px;
    }

    .slot-group {
      background: var(--bg-card-inner);
      border: 1px solid var(--border-subtle);
      border-radius: 2px 2px 10px 2px;
      padding: 8px 10px;
      display: flex;
      flex-direction: column;
      gap: 6px;
    }

    .slot-group-head {
      display: flex;
      justify-content: space-between;
      align-items: center;
      gap: 6px;
    }

    .slot-time {
      font-size: 11px;
      font-weight: 700;
      font-variant-numeric: tabular-nums;
      letter-spacing: 0.03em;
    }

    /* ==================== TABELA CORPO CLÍNICO ==================== */
    .table-card {
      padding: 0;
      overflow: hidden;
    }

    .table-scroll {
      width: 100%;
      overflow-x: auto;
    }

    .table-card table {
      width: 100%;
      border-collapse: collapse;
      font-size: 13px;
    }

    .table-card thead th {
      text-align: left;
      padding: 10px 14px;
      font-size: 11px;
      font-weight: 600;
      text-transform: uppercase;
      letter-spacing: 0.08em;
      color: var(--text-muted);
      background: var(--bg-subnav);
      border-bottom: 1px solid var(--border-subtle);
      white-space: nowrap;
    }

    .table-card tbody td {
      padding: 9px 14px;
      border-bottom: 1px solid var(--border-subtle);
      color: var(--text-primary);
    }

    .table-card tbody tr:hover {
      background: rgba(255, 255, 255, 0.03);
    }

    [data-theme="light"] .table-card tbody tr:hover {
      background: rgba(0, 68, 196, 0.04);
    }

    .num-col { text-align: right; font-variant-numeric: tabular-nums; }

    .empty-state {
      text-align: center;
      padding: 40px 20px;
      border: 1px dashed var(--border-dashed);
      border-radius: 3px 3px 18px 3px;
      color: var(--text-secondary);
    }

    /* ==================== RODAPÉ ==================== */
    .page-footer {
      margin-top: 24px;
      padding-top: 14px;
      border-top: 1px solid var(--border-subtle);
      display: flex;
      justify-content: space-between;
      align-items: center;
      gap: 12px;
      flex-wrap: wrap;
      font-size: 11.5px;
      color: var(--text-muted);
    }

    /* ==================== RESPONSIVIDADE ==================== */
    @media (max-width: 1120px) {
      .filter-row { grid-template-columns: 1fr 1fr; }
    }

    @media (max-width: 820px) {
      .grid-4 { grid-template-columns: repeat(2, 1fr); }
      .req-grid-3, .req-grid-2 { grid-template-columns: 1fr; }
      .header-right { display: none; }
      .subnav-meta { display: none; }
    }

    @media (max-width: 600px) {
      .top-header { padding: 10px 14px; }
      .subnav { padding: 0 14px; }
      .container { padding: 14px 12px 32px; }
      .grid-4 { grid-template-columns: 1fr; }
      .col-span-2 { grid-column: span 1; }
      .filter-row { grid-template-columns: 1fr; }
      .hero h1 { font-size: 21px; }
      .header-divider { display: none; }
      .header-left { flex-direction: column; align-items: flex-start; gap: 6px; }
    }

    /* ==================== IMPRESSÃO ==================== */
    @media print {
      body { background: #ffffff !important; color: #000000 !important; }
      .subnav, .filter-card, .theme-toggle, .btn-accent, .btn-outline { display: none !important; }
      .top-header { -webkit-print-color-adjust: exact; print-color-adjust: exact; }
      .card, .matrix-card, .list-day-card, .slot-group {
        background: #ffffff !important;
        color: #000000 !important;
        border-color: #cccccc !important;
        break-inside: avoid;
      }
      .roster-matrix { min-width: 0 !important; }
      .roster-matrix thead th,
      .roster-matrix tbody th.col-slot-cell,
      .roster-matrix tbody td.matrix-cell {
        position: static !important;
        background: #ffffff !important;
        color: #000000 !important;
        border-color: #cccccc !important;
      }
      .doc-pill { background: #eef3fb !important; color: #000000 !important; }
      .doc-drop-btn, .modal-backdrop { display: none !important; }
    }
  </style>
</head>
<body>
  <!-- HEADER INSTITUCIONAL -->
  <header class="top-header">
    <div class="header-left">
      __HAPVIDA_LOGO_SVG__
      <div class="header-divider"></div>
      <div class="header-titles">
        <div class="header-title" id="headerHospital">Hospital Cruzeiro do Sul — Escala Médica</div>
        <div class="header-sub">Visualização de Plantões · Sincronização Integrada</div>
      </div>
    </div>
  </header>

  <!-- BARRA DE SUB-NAVEGAÇÃO (ABAS DE ESCALAS) -->
  <nav class="subnav" aria-label="Escalas disponíveis">
    <div class="subnav-tabs" id="scaleTabs"></div>
    <div class="subnav-right">
      <a class="btn-outline" id="directorLink" href="/dashboard" style="display:none; padding:5px 10px; text-decoration:none;">Painel Diretor →</a>
      <button class="theme-toggle" id="themeToggle" type="button" title="Alternar tema claro/escuro">
        <span id="themeIcon">☀</span>
        <span id="themeLabel">Claro</span>
      </button>
      <a class="btn-outline" id="logoutLink" href="/logout" style="display:none; padding:5px 10px; text-decoration:none;">Sair</a>
    </div>
  </nav>

  <!-- CONTEÚDO PRINCIPAL -->
  <main class="container">
    <section class="hero">
      <div>
        <div class="eyebrow" id="heroEyebrow">ESCALA MÉDICA · HOSPITAL CRUZEIRO DO SUL</div>
        <h1 id="heroTitle">Carregando escala...</h1>
        <p class="hero-desc" id="heroDesc">
          Consulte o preenchimento da escala por semana, aplique filtros específicos e solicite novos plantões ou desistências para aprovação da Diretoria.
        </p>
      </div>
      <div class="sync-pill" title="Data da última sincronização da escala">
        <span class="sync-dot"></span>
        <span>Atualizado em <strong id="syncTime">--</strong></span>
      </div>
    </section>

    <!-- BARRA DE FILTROS -->
    <section class="card filter-card" id="filterCard" aria-label="Filtros da escala">
      <div class="filter-row">
        <div class="field-group">
          <label class="field-label" for="doctorSelect">Médico (filtrar ou destacar)</label>
          <select id="doctorSelect" class="select-control">
            <option value="">Todos os médicos da escala</option>
          </select>
        </div>

        <div class="field-group">
          <label class="field-label" for="shiftSelect">Turno / Período</label>
          <select id="shiftSelect" class="select-control">
            <option value="todos">Todos os turnos</option>
            <option value="manha">Manhã (07h–13h)</option>
            <option value="tarde">Tarde (13h–19h)</option>
            <option value="noite">Noite (19h em diante)</option>
          </select>
        </div>

        <div class="field-group">
          <label class="field-label" for="searchInput">Busca rápida por nome</label>
          <input id="searchInput" type="search" class="input-control" placeholder="Digite parte do nome..." />
        </div>

        <div class="field-group">
          <button type="button" class="btn-outline" id="clearFiltersBtn" style="height:38px;">Limpar</button>
        </div>
      </div>

      <div class="toolbar-bottom">
        <div class="view-switcher" role="group" aria-label="Modo de visualização">
          <button type="button" class="btn-outline is-active" data-view="grid">Grade Semanal</button>
          <button type="button" class="btn-outline" data-view="list">Lista por Dia</button>
        </div>

        <div style="display:flex; align-items:center; gap:14px; flex-wrap:wrap;">
          <label class="checkbox-pill">
            <input type="checkbox" id="onlyDoctorCheck" />
            <span>Mostrar somente o médico selecionado</span>
          </label>
          <label class="checkbox-pill">
            <input type="checkbox" id="onlyOpenCheck" />
            <span>Somente vagas em aberto</span>
          </label>
          <label class="checkbox-pill">
            <input type="checkbox" id="showOpenCheck" checked />
            <span>Exibir vagas em aberto</span>
          </label>
          <button type="button" class="btn-outline" id="printBtn">Imprimir / PDF</button>
        </div>
      </div>
    </section>

    <!-- CARD PESSOAL DO MÉDICO SELECIONADO -->
    <section class="card doctor-banner" id="doctorBanner">
      <div class="doctor-banner-grid">
        <div>
          <div class="eyebrow">RESUMO INDIVIDUAL NA ESCALA SELECIONADA</div>
          <div class="doctor-banner-name" id="docBannerName">--</div>
          <div class="doctor-banner-tags" id="docBannerTags"></div>
        </div>
        <div style="display:flex; gap:10px; flex-wrap:wrap;">
          <button type="button" class="btn-accent" id="exportIcsBtn">Exportar para Agenda (.ICS) →</button>
        </div>
      </div>
    </section>

    <!-- LEGENDA + ÁREA DE RENDERIZAÇÃO -->
    <div class="legend-bar" id="legendBar">
      <h2 class="section-title" id="viewHeading">Grade Semanal de Plantões</h2>
      <div class="legend-items">
        <span class="shift-chip shift-manha">Manhã · 07h–13h</span>
        <span class="shift-chip shift-tarde">Tarde · 13h–19h</span>
        <span class="shift-chip shift-noite">Noite · 19h+</span>
      </div>
    </div>

    <div id="contentArea"></div>

    <!-- RODAPÉ -->
    <footer class="page-footer">
      <div>
        <strong id="footerHospital">Hospital Cruzeiro do Sul</strong> · Sistema de plantões médicos (dados pessoais sensíveis omitidos conforme LGPD).
      </div>
      <div>
        Planilha atualizada em <span id="footerMtime">--</span> · Publicado em <span id="footerGen">--</span>
      </div>
    </footer>
  </main>

  <!-- DADOS DA ESCALA EMBUTIDOS AUTOMATICAMENTE PELO SCRIPT PYTHON -->
  <script id="escala-data" type="application/json">__JSON_DATA__</script>

  <script>
    (function () {
      let DATA = JSON.parse(document.getElementById("escala-data").textContent);
      let SESSION = { authenticated: false, role: null };
      let REQUESTS = [];

      const state = {
        activeTab: "scale", // "scale" | "requests"
        scaleIndex: 0,
        view: window.innerWidth <= 640 ? "list" : "grid",
        selectedDoctor: "",
        searchText: "",
        shiftFilter: "todos",
        onlyDoctor: false,
        onlyOpen: false,
        showOpen: true,
        verifiedDoctor: null, // { nome, crm, uf, scheduledShifts }
        reqMode: "solicitar", // "solicitar" | "desistir"
        reqFeedback: null     // { type: "ok" | "error", text }
      };

      const todayIso = new Date().toISOString().slice(0, 10);

      // Elementos DOM
      const tabsEl = document.getElementById("scaleTabs");
      const filterCard = document.getElementById("filterCard");
      const legendBar = document.getElementById("legendBar");
      const doctorSelect = document.getElementById("doctorSelect");
      const shiftSelect = document.getElementById("shiftSelect");
      const searchInput = document.getElementById("searchInput");
      const onlyDoctorCheck = document.getElementById("onlyDoctorCheck");
      const onlyOpenCheck = document.getElementById("onlyOpenCheck");
      const showOpenCheck = document.getElementById("showOpenCheck");
      const contentArea = document.getElementById("contentArea");
      const doctorBanner = document.getElementById("doctorBanner");

      function updateMetaTexts() {
        document.getElementById("headerHospital").textContent = `${titleCase(DATA.hospital)} — Escala Médica`;
        document.getElementById("syncTime").textContent = DATA.geradoEm;
        document.getElementById("footerHospital").textContent = titleCase(DATA.hospital);
        document.getElementById("footerMtime").textContent = DATA.planilhaModificadaEm;
        document.getElementById("footerGen").textContent = DATA.geradoEm;
      }
      updateMetaTexts();

      // Tema Claro / Escuro
      const themeToggle = document.getElementById("themeToggle");
      const savedTheme = localStorage.getItem("hapvida_escala_theme") || "dark";
      applyTheme(savedTheme);

      themeToggle.addEventListener("click", () => {
        const next = document.documentElement.getAttribute("data-theme") === "dark" ? "light" : "dark";
        applyTheme(next);
        localStorage.setItem("hapvida_escala_theme", next);
      });

      function applyTheme(theme) {
        document.documentElement.setAttribute("data-theme", theme);
        document.getElementById("themeIcon").textContent = theme === "dark" ? "☀" : "☾";
        document.getElementById("themeLabel").textContent = theme === "dark" ? "Claro" : "Escuro";
      }

      function titleCase(str) {
        if (!str) return "";
        const preps = new Set(["DE", "DA", "DO", "DAS", "DOS", "E"]);
        return str.split(" ").map((w, i) => {
          const u = w.toUpperCase();
          if (i > 0 && preps.has(u)) return u.toLowerCase();
          return w.charAt(0).toUpperCase() + w.slice(1).toLowerCase();
        }).join(" ");
      }

      function escapeHtml(str) {
        return String(str || "")
          .replace(/&/g, "&amp;")
          .replace(/</g, "&lt;")
          .replace(/>/g, "&gt;")
          .replace(/"/g, "&quot;");
      }

      function normalize(str) {
        return String(str || "").normalize("NFD").replace(/[\u0300-\u036f]/g, "").toUpperCase();
      }

      function getActiveScale() {
        return DATA.escalas[state.scaleIndex];
      }

      function renderTabs() {
        tabsEl.innerHTML = "";
        DATA.escalas.forEach((sc, idx) => {
          const btn = document.createElement("button");
          btn.type = "button";
          const active = state.activeTab === "scale" && idx === state.scaleIndex;
          btn.className = "subnav-tab" + (active ? " is-active" : "");
          btn.textContent = sc.tabLabel;
          btn.addEventListener("click", () => {
            state.activeTab = "scale";
            state.scaleIndex = idx;
            populateDoctors();
            renderAll();
          });
          tabsEl.appendChild(btn);
        });

        const reqBtn = document.createElement("button");
        reqBtn.type = "button";
        reqBtn.className = "subnav-tab" + (state.activeTab === "requests" ? " is-active" : "");
        reqBtn.textContent = "PEGAR / DESISTIR PLANTÃO";
        reqBtn.addEventListener("click", () => {
          state.activeTab = "requests";
          state.reqFeedback = null;
          renderAll();
        });
        tabsEl.appendChild(reqBtn);
      }

      function populateDoctors() {
        const sc = getActiveScale();
        const prev = state.selectedDoctor;
        doctorSelect.innerHTML = '<option value="">Todos os médicos da escala</option>';
        const sorted = [...sc.medicos].sort((a, b) => a.nome.localeCompare(b.nome, "pt-BR"));
        sorted.forEach(m => {
          const opt = document.createElement("option");
          opt.value = m.nome;
          const crmTxt = m.crm ? ` · CRM ${m.crm}/${m.uf}` : "";
          opt.textContent = `${m.nome}${crmTxt}`;
          doctorSelect.appendChild(opt);
        });
        if (sorted.some(m => m.nome === prev)) {
          doctorSelect.value = prev;
          state.selectedDoctor = prev;
        } else {
          state.selectedDoctor = "";
          doctorSelect.value = "";
        }
      }

      function renderHeaderInfo(sc) {
        document.getElementById("heroEyebrow").textContent = `${sc.unidade} · ${sc.municipio} · ${sc.mesAnoLabel.toUpperCase()}`;
        document.getElementById("heroTitle").textContent = `${sc.titulo} — ${sc.mesAnoLabel}`;
      }

      function renderDoctorBanner(sc) {
        if (!state.selectedDoctor) {
          doctorBanner.classList.remove("is-visible");
          return;
        }
        const m = sc.medicos.find(x => x.nome === state.selectedDoctor);
        if (!m) {
          doctorBanner.classList.remove("is-visible");
          return;
        }
        doctorBanner.classList.add("is-visible");
        document.getElementById("docBannerName").textContent = m.nome;
        const crmTag = m.crm ? `<span class="tech-tag">CRM ${escapeHtml(m.crm)}/${escapeHtml(m.uf)}</span>` : "";
        document.getElementById("docBannerTags").innerHTML = `
          ${crmTag}
          <span class="tech-tag">${m.totalPlantoes} plantão(ões) na escala</span>
        `;
      }

      function filterGroup(grp) {
        if (state.shiftFilter !== "todos" && grp.periodo !== state.shiftFilter) return null;

        const q = normalize(state.searchText);
        let meds = grp.medicos;

        if (state.onlyOpen) {
          if (grp.vagasAbertas <= 0) return null;
          return { ...grp, medicos: [] };
        }

        if (state.selectedDoctor && state.onlyDoctor) {
          meds = meds.filter(m => m.nome === state.selectedDoctor);
        }

        if (q) {
          meds = meds.filter(m => normalize(m.nome).includes(q) || normalize(m.crm).includes(q));
        }

        const showOpenHere = state.showOpen && !state.onlyDoctor && !q && grp.vagasAbertas > 0;

        if (meds.length === 0 && !showOpenHere) return null;
        return {
          ...grp,
          medicos: meds,
          vagasAbertasVisiveis: showOpenHere ? grp.vagasAbertas : 0
        };
      }

      function namesMatch(aNorm, bNorm) {
        if (aNorm === bNorm) return true;
        const preps = new Set(["DE", "DA", "DO", "DAS", "DOS", "E"]);
        const tokA = aNorm.split(" ").filter(w => w && !preps.has(w));
        const tokB = bNorm.split(" ").filter(w => w && !preps.has(w));
        if (tokA.length < 2 || tokB.length < 2) return false;
        if (tokA[0] !== tokB[0] || tokA[tokA.length - 1] !== tokB[tokB.length - 1]) return false;
        const setB = new Set(tokB);
        const setA = new Set(tokA);
        const shorter = tokA.length <= tokB.length ? tokA : tokB;
        const longerSet = tokA.length <= tokB.length ? setB : setA;
        return shorter.every(w => longerSet.has(w));
      }

      function validateDoctorIdentity(rawNome, rawCrm, rawUf) {
        const nomeClean = String(rawNome || "").trim().replace(/\s+/g, " ");
        const nomeNorm = normalize(nomeClean);
        const crmDigits = String(rawCrm || "").replace(/\D/g, "");
        const ufClean = String(rawUf || "SP").trim().toUpperCase() || "SP";

        if (nomeClean.split(" ").filter(Boolean).length < 2) {
          return { ok: false, error: "Informe seu nome completo (nome e sobrenome)." };
        }
        if (crmDigits.length < 4) {
          return { ok: false, error: "Informe um número de CRM válido." };
        }

        const catalog = [];
        DATA.escalas.forEach(sc => {
          sc.medicos.forEach(m => {
            catalog.push(m);
          });
        });

        const byCrm = catalog.find(m => String(m.crm || "").replace(/\D/g, "") === crmDigits);
        if (byCrm) {
          if (!namesMatch(nomeNorm, normalize(byCrm.nome))) {
            return { ok: false, error: "O CRM informado não confere com o nome cadastrado na unidade. Verifique os dados." };
          }
          return { ok: true, nome: byCrm.nome, crm: byCrm.crm, uf: byCrm.uf || ufClean };
        }

        const byName = catalog.find(m => namesMatch(nomeNorm, normalize(m.nome)));
        if (byName && String(byName.crm || "").replace(/\D/g, "")) {
          return { ok: false, error: "Este nome já possui outro CRM cadastrado na unidade. Confira seu CRM." };
        }

        return { ok: true, nome: nomeClean.toUpperCase(), crm: crmDigits, uf: ufClean };
      }

      function collectScheduledShifts(docNameNorm) {
        const shifts = [];
        DATA.escalas.forEach(sc => {
          sc.semanas.forEach(wk => {
            wk.dias.forEach(dia => {
              dia.grupos.forEach(g => {
                if (g.medicos.some(m => normalize(m.nome) === docNameNorm)) {
                  shifts.push({
                    scaleId: sc.id,
                    scaleTitle: sc.tabLabel,
                    titulo: sc.titulo,
                    iso: dia.iso,
                    ddmm: dia.ddmm,
                    weekdayFull: dia.weekdayFull,
                    inicio: g.inicio,
                    fim: g.fim,
                    horas: g.horas,
                    periodo: g.periodo
                  });
                }
              });
            });
          });
        });
        return shifts.sort((a, b) => a.iso.localeCompare(b.iso) || a.inicio.localeCompare(b.inicio));
      }

      function renderRequestsWorkspace() {
        document.getElementById("heroEyebrow").textContent = `${DATA.hospital} · ${DATA.municipio}`;
        document.getElementById("heroTitle").textContent = "Pegar Nova Vaga ou Desistir de Plantão";

        const fbHtml = state.reqFeedback
          ? `<div class="feedback-box ${state.reqFeedback.type === 'ok' ? 'is-ok' : 'is-error'}">${escapeHtml(state.reqFeedback.text)}</div>`
          : "";

        if (!state.verifiedDoctor) {
          contentArea.innerHTML = `
            <div class="req-workspace">
              <section class="card">
                <div class="eyebrow">ETAPA 1 · IDENTIFICAÇÃO DO MÉDICO</div>
                <h2 class="section-title" style="margin:6px 0 6px;">Valide seu Nome Completo e CRM</h2>
                <p class="hero-desc" style="margin-bottom:14px;">
                  Após a validação, o sistema verifica automaticamente se você já possui plantões escalados para liberar as opções disponíveis.
                </p>
                ${fbHtml}
                <form id="verifyDocForm">
                  <div class="req-grid-3">
                    <div class="field-group">
                      <label class="field-label" for="valNomeInput">Nome Completo</label>
                      <input id="valNomeInput" type="text" class="input-control" placeholder="Digite seu nome completo..." required />
                    </div>
                    <div class="field-group">
                      <label class="field-label" for="valCrmInput">CRM</label>
                      <input id="valCrmInput" type="text" inputmode="numeric" class="input-control" placeholder="Ex.: 229571" required />
                    </div>
                    <div class="field-group">
                      <label class="field-label" for="valUfInput">UF</label>
                      <input id="valUfInput" type="text" maxlength="2" class="input-control" value="SP" required />
                    </div>
                    <div class="field-group">
                      <button type="submit" class="btn-accent" style="height:36px;">Validar</button>
                    </div>
                  </div>
                </form>
              </section>
            </div>
          `;

          document.getElementById("verifyDocForm").addEventListener("submit", (e) => {
            e.preventDefault();
            const res = validateDoctorIdentity(
              document.getElementById("valNomeInput").value,
              document.getElementById("valCrmInput").value,
              document.getElementById("valUfInput").value
            );
            if (!res.ok) {
              state.reqFeedback = { type: "error", text: res.error };
              renderAll();
              return;
            }
            state.verifiedDoctor = { nome: res.nome, crm: res.crm, uf: res.uf };
            state.reqMode = "solicitar";
            state.reqFeedback = null;
            renderAll();
          });
          return;
        }

        const doc = state.verifiedDoctor;
        const docNorm = normalize(doc.nome);
        const scheduledShifts = collectScheduledShifts(docNorm);
        if (scheduledShifts.length === 0) {
          state.reqMode = "solicitar";
        }

        const modeSwitchHtml = scheduledShifts.length > 0
          ? `
            <div class="mode-switch">
              <button type="button" class="btn-outline ${state.reqMode === 'solicitar' ? 'is-active' : ''}" data-req-mode="solicitar">Solicitar Novo Plantão</button>
              <button type="button" class="btn-outline ${state.reqMode === 'desistir' ? 'is-active' : ''}" data-req-mode="desistir">Desistir de Plantão (${scheduledShifts.length})</button>
            </div>
          `
          : `
            <div style="font-size:12.5px; color:var(--text-secondary); margin-bottom:12px;">
              Você ainda não possui plantões escalados. Utilize o formulário abaixo para solicitar uma nova vaga.
            </div>
          `;

        let actionFormHtml = "";
        if (state.reqMode === "solicitar") {
          const scaleOpts = DATA.escalas.map(sc => `<option value="${escapeHtml(sc.id)}">${escapeHtml(sc.tabLabel)} — ${escapeHtml(sc.titulo)}</option>`).join("");
          actionFormHtml = `
            <form id="solicitarForm">
              <div class="req-grid-2" style="margin-bottom:12px;">
                <div class="field-group">
                  <label class="field-label" for="solScaleSelect">Escala / Setor</label>
                  <select id="solScaleSelect" class="select-control">${scaleOpts}</select>
                </div>
                <div class="field-group">
                  <label class="field-label" for="solDaySelect">Dia com vaga aberta</label>
                  <select id="solDaySelect" class="select-control"></select>
                </div>
              </div>
              <div class="req-grid-2" style="margin-bottom:14px;">
                <div class="field-group">
                  <label class="field-label" for="solSlotSelect">Horário / Turno disponível</label>
                  <select id="solSlotSelect" class="select-control"></select>
                </div>
                <div class="field-group">
                  <label class="field-label" for="solMotivoInput">Observação (opcional)</label>
                  <input id="solMotivoInput" type="text" class="input-control" placeholder="Ex.: disponibilidade confirmada..." />
                </div>
              </div>
              <div style="display:flex; justify-content:flex-end;">
                <button type="submit" class="btn-accent" id="solSubmitBtn">Enviar Solicitação de Plantão</button>
              </div>
            </form>
          `;
        } else {
          const shiftOpts = scheduledShifts.map((s, idx) =>
            `<option value="${idx}">${escapeHtml(s.scaleTitle)} · ${escapeHtml(s.ddmm)} (${escapeHtml(s.weekdayFull)}) · ${escapeHtml(s.inicio)}–${escapeHtml(s.fim)} (${s.horas}h)</option>`
          ).join("");
          actionFormHtml = `
            <form id="desistirForm">
              <div class="req-grid-2" style="margin-bottom:14px;">
                <div class="field-group">
                  <label class="field-label" for="desShiftSelect">Selecione o plantão que deseja desistir</label>
                  <select id="desShiftSelect" class="select-control">${shiftOpts}</select>
                </div>
                <div class="field-group">
                  <label class="field-label" for="desMotivoInput">Motivo da desistência (opcional)</label>
                  <input id="desMotivoInput" type="text" class="input-control" placeholder="Informe o motivo..." />
                </div>
              </div>
              <div style="display:flex; justify-content:flex-end;">
                <button type="submit" class="btn-accent" id="desSubmitBtn">Enviar Pedido de Desistência</button>
              </div>
            </form>
          `;
        }

        const myRequests = REQUESTS.filter(r =>
          normalize(r.medicoNome) === docNorm ||
          (doc.crm && String(r.medicoCrm || "").replace(/\D/g, "") === doc.crm)
        );
        const statusLabel = { pending: "Pendente", approved: "Aprovada", rejected: "Recusada" };
        const histHtml = myRequests.length === 0
          ? `<div style="font-size:12.5px; color:var(--text-muted);">Nenhuma solicitação enviada ainda para este CRM.</div>`
          : myRequests.map(r => {
              const tipoLabel = r.tipo === "solicitar" ? "Solicitação de Plantão" : "Desistência de Plantão";
              const noteHtml = r.directorNote ? `<div style="font-size:12px; color:var(--text-secondary); margin-top:4px;">Diretoria: ${escapeHtml(r.directorNote)}</div>` : "";
              return `
                <div class="req-item-card">
                  <div style="display:flex; justify-content:space-between; align-items:center; gap:8px; flex-wrap:wrap; margin-bottom:4px;">
                    <span class="eyebrow">${tipoLabel} · ${escapeHtml(r.scaleTitle)} · ${escapeHtml(r.ddmm)} (${escapeHtml(r.inicio)}–${escapeHtml(r.fim)})</span>
                    <span class="status-pill status-${r.status}">${statusLabel[r.status] || r.status}</span>
                  </div>
                  <div style="font-size:12px; color:var(--text-muted);">Enviada em ${escapeHtml(r.createdAt)}${r.motivo ? ' · Obs.: ' + escapeHtml(r.motivo) : ''}</div>
                  ${noteHtml}
                </div>
              `;
            }).join("");

        contentArea.innerHTML = `
          <div class="req-workspace">
            <section class="card" style="border-color: var(--accent-orange);">
              <div class="doctor-banner-grid">
                <div>
                  <div class="eyebrow">MÉDICO VALIDADO</div>
                  <div class="doctor-banner-name">${escapeHtml(doc.nome)}</div>
                  <div class="doctor-banner-tags">
                    <span class="tech-tag">CRM ${escapeHtml(doc.crm)}/${escapeHtml(doc.uf)}</span>
                    <span class="tech-tag">${scheduledShifts.length} plantão(ões) escalado(s)</span>
                  </div>
                </div>
                <button type="button" class="btn-outline" id="changeVerifiedDocBtn">Trocar médico</button>
              </div>
            </section>

            <section class="card">
              ${fbHtml}
              ${modeSwitchHtml}
              ${actionFormHtml}
            </section>

            <section class="card">
              <div class="eyebrow" style="margin-bottom:10px;">MINHAS SOLICITAÇÕES ENVIADAS</div>
              ${histHtml}
            </section>
          </div>
        `;

        document.getElementById("changeVerifiedDocBtn").addEventListener("click", () => {
          state.verifiedDoctor = null;
          state.reqFeedback = null;
          renderAll();
        });

        contentArea.querySelectorAll("[data-req-mode]").forEach(btn => {
          btn.addEventListener("click", () => {
            state.reqMode = btn.getAttribute("data-req-mode");
            state.reqFeedback = null;
            renderAll();
          });
        });

        if (state.reqMode === "solicitar") {
          const solScaleSelect = document.getElementById("solScaleSelect");
          const solDaySelect = document.getElementById("solDaySelect");
          const solSlotSelect = document.getElementById("solSlotSelect");
          const solSubmitBtn = document.getElementById("solSubmitBtn");
          const shiftName = { manha: "Manhã", tarde: "Tarde", noite: "Noite" };

          function getOpenDaysForScale(scaleId) {
            const sc = DATA.escalas.find(x => x.id === scaleId);
            if (!sc) return [];
            const days = [];
            sc.semanas.forEach(wk => {
              wk.dias.forEach(dia => {
                if (!dia.inMonth) return;
                const openSlots = dia.grupos.filter(g =>
                  g.vagasAbertas > 0 && !g.medicos.some(m => normalize(m.nome) === docNorm)
                );
                if (openSlots.length > 0) {
                  days.push({ dia, openSlots });
                }
              });
            });
            return days;
          }

          function populateOpenSlots() {
            const days = getOpenDaysForScale(solScaleSelect.value);
            const chosen = days.find(d => d.dia.iso === solDaySelect.value);
            solSlotSelect.innerHTML = "";
            if (!chosen) {
              solSlotSelect.innerHTML = `<option value="">Sem horários disponíveis</option>`;
              solSubmitBtn.disabled = true;
              return;
            }
            chosen.openSlots.forEach(g => {
              const opt = document.createElement("option");
              opt.value = `${g.inicio}|${g.fim}|${g.horas}|${g.periodo}`;
              opt.textContent = `${g.inicio} – ${g.fim} (${shiftName[g.periodo] || ""} · ${g.horas}h · ${g.vagasAbertas} vaga${g.vagasAbertas > 1 ? 's' : ''})`;
              solSlotSelect.appendChild(opt);
            });
            solSubmitBtn.disabled = false;
          }

          function populateOpenDays() {
            const days = getOpenDaysForScale(solScaleSelect.value);
            solDaySelect.innerHTML = "";
            if (days.length === 0) {
              solDaySelect.innerHTML = `<option value="">Nenhum dia com vaga aberta nesta escala</option>`;
              populateOpenSlots();
              return;
            }
            days.forEach(d => {
              const opt = document.createElement("option");
              opt.value = d.dia.iso;
              opt.textContent = `${d.dia.ddmm} · ${d.dia.weekdayFull} (${d.openSlots.length} faixa${d.openSlots.length > 1 ? 's' : ''})`;
              solDaySelect.appendChild(opt);
            });
            populateOpenSlots();
          }

          solScaleSelect.addEventListener("change", populateOpenDays);
          solDaySelect.addEventListener("change", populateOpenSlots);
          populateOpenDays();

          document.getElementById("solicitarForm").addEventListener("submit", async (e) => {
            e.preventDefault();
            if (!solDaySelect.value || !solSlotSelect.value) return;
            const sc = DATA.escalas.find(x => x.id === solScaleSelect.value);
            const [inicio, fim, horasStr, periodo] = solSlotSelect.value.split("|");
            const dayObj = sc.semanas.flatMap(w => w.dias).find(d => d.iso === solDaySelect.value);
            solSubmitBtn.disabled = true;

            try {
              const res = await fetch("/api/requests", {
                method: "POST",
                headers: { "Content-Type": "application/json" },
                body: JSON.stringify({
                  tipo: "solicitar",
                  scaleId: sc.id,
                  scaleTitle: sc.tabLabel,
                  iso: dayObj.iso,
                  ddmm: dayObj.ddmm,
                  inicio,
                  fim,
                  horas: Number(horasStr),
                  periodo,
                  medicoNome: doc.nome,
                  medicoCrm: doc.crm,
                  medicoUf: doc.uf,
                  motivo: document.getElementById("solMotivoInput").value.trim()
                })
              });
              const data = await res.json().catch(() => ({}));
              if (!res.ok) {
                state.reqFeedback = { type: "error", text: data.error || "Erro ao enviar solicitação." };
                renderAll();
                return;
              }
              await refreshLiveServerData();
              state.reqFeedback = { type: "ok", text: "Solicitação de plantão enviada para aprovação da Diretoria." };
              renderAll();
            } catch (err) {
              state.reqFeedback = { type: "error", text: "Erro de conexão com o servidor." };
              renderAll();
            }
          });
        } else {
          document.getElementById("desistirForm").addEventListener("submit", async (e) => {
            e.preventDefault();
            const idx = Number(document.getElementById("desShiftSelect").value);
            const chosen = scheduledShifts[idx];
            if (!chosen) return;
            const btn = document.getElementById("desSubmitBtn");
            btn.disabled = true;

            try {
              const res = await fetch("/api/requests", {
                method: "POST",
                headers: { "Content-Type": "application/json" },
                body: JSON.stringify({
                  tipo: "desistir",
                  scaleId: chosen.scaleId,
                  scaleTitle: chosen.scaleTitle,
                  iso: chosen.iso,
                  ddmm: chosen.ddmm,
                  inicio: chosen.inicio,
                  fim: chosen.fim,
                  horas: chosen.horas,
                  periodo: chosen.periodo,
                  medicoNome: doc.nome,
                  medicoCrm: doc.crm,
                  medicoUf: doc.uf,
                  motivo: document.getElementById("desMotivoInput").value.trim()
                })
              });
              const data = await res.json().catch(() => ({}));
              if (!res.ok) {
                state.reqFeedback = { type: "error", text: data.error || "Erro ao enviar pedido de desistência." };
                renderAll();
                return;
              }
              await refreshLiveServerData();
              state.reqFeedback = { type: "ok", text: "Pedido de desistência enviado para aprovação da Diretoria." };
              renderAll();
            } catch (err) {
              state.reqFeedback = { type: "error", text: "Erro de conexão com o servidor." };
              renderAll();
            }
          });
        }
      }

      function bindDoctorPills() {
        contentArea.querySelectorAll("[data-doc]").forEach(btn => {
          btn.addEventListener("click", () => {
            const docName = btn.getAttribute("data-doc");
            state.selectedDoctor = state.selectedDoctor === docName ? "" : docName;
            doctorSelect.value = state.selectedDoctor;
            if (!state.selectedDoctor) {
              state.onlyDoctor = false;
              onlyDoctorCheck.checked = false;
            }
            renderAll();
          });
        });
      }

      function renderGrid(sc) {
        document.getElementById("viewHeading").textContent = `Grade Semanal · ${sc.titulo} (${sc.mesAnoLabel})`;
        let html = "";
        let anyVisible = false;
        const shiftLabel = { manha: "Manhã", tarde: "Tarde", noite: "Noite" };

        sc.semanas.forEach(wk => {
          // Coleta faixas horárias únicas da semana em ordem cronológica
          const slotMap = new Map();
          wk.dias.forEach(dia => {
            dia.grupos.forEach(g => {
              const key = `${g.inicio}-${g.fim}`;
              if (!slotMap.has(key)) {
                slotMap.set(key, { key, inicio: g.inicio, fim: g.fim, horas: g.horas, periodo: g.periodo });
              }
            });
          });
          const slots = Array.from(slotMap.values()).sort((a, b) => a.inicio.localeCompare(b.inicio) || a.fim.localeCompare(b.fim));

          let rowsHtml = "";
          let weekFilled = 0;
          let weekOpen = 0;

          slots.forEach(slot => {
            let rowHasVisible = false;
            let cellsHtml = "";

            wk.dias.forEach(dia => {
              const isToday = dia.iso === todayIso;
              const cellClasses = ["matrix-cell"];
              if (!dia.inMonth) cellClasses.push("is-out-month");
              if (isToday) cellClasses.push("is-today");

              const rawGrp = dia.grupos.find(g => `${g.inicio}-${g.fim}` === slot.key);
              const fg = rawGrp ? filterGroup(rawGrp) : null;

              if (!fg) {
                cellsHtml += `<td class="${cellClasses.join(" ")}"></td>`;
                return;
              }

              rowHasVisible = true;
              anyVisible = true;
              if (dia.inMonth) {
                weekFilled += fg.medicos.length;
                weekOpen += fg.vagasAbertasVisiveis;
              }

              const medsHtml = fg.medicos.map(m => {
                const sel = state.selectedDoctor && m.nome === state.selectedDoctor;
                const titleAttr = `${m.nome}${m.crm ? ' · CRM ' + m.crm + '/' + m.uf : ''} (clique para destacar)`;
                return `<button type="button" class="doc-pill ${sel ? 'is-selected' : ''}" data-doc="${escapeHtml(m.nome)}" title="${escapeHtml(titleAttr)}">${escapeHtml(m.nomeCurto)}</button>`;
              }).join("");

              let openHtml = "";
              if (fg.vagasAbertasVisiveis > 0) {
                const label = `${fg.vagasAbertasVisiveis} vaga${fg.vagasAbertasVisiveis > 1 ? 's' : ''} em aberto`;
                openHtml = fg.medicos.length > 0
                  ? `<div class="open-hint">+ ${label}</div>`
                  : `<div class="open-pill-empty">${label}</div>`;
              }

              cellsHtml += `
                <td class="${cellClasses.join(" ")}">
                  <div class="cell-stack">
                    ${medsHtml}
                    ${openHtml}
                  </div>
                </td>
              `;
            });

            if (rowHasVisible) {
              rowsHtml += `
                <tr>
                  <th class="col-slot-cell" scope="row">
                    <div class="slot-cell-wrap">
                      <span class="shift-chip shift-${slot.periodo} slot-time">${slot.inicio} – ${slot.fim}</span>
                      <span class="slot-meta">${shiftLabel[slot.periodo] || ""} · ${slot.horas}h</span>
                    </div>
                  </th>
                  ${cellsHtml}
                </tr>
              `;
            }
          });

          if (!rowsHtml) return;

          const dayHeadersHtml = wk.dias.map(dia => {
            const isToday = dia.iso === todayIso;
            const classes = ["col-day"];
            if (!dia.inMonth) classes.push("is-out-month");
            if (isToday) classes.push("is-today");
            const todayBadge = isToday ? `<span class="badge-yellow" style="padding:2px 6px;">HOJE</span>` : "";
            return `
              <th class="${classes.join(" ")}" scope="col">
                <div class="m-day-head">
                  <div>
                    <div class="m-day-wk ${dia.isWeekend ? 'is-weekend' : ''}">${dia.weekdayShort}</div>
                    <div class="m-day-date">${dia.ddmm}</div>
                  </div>
                  ${todayBadge}
                </div>
              </th>
            `;
          }).join("");

          const summaryParts = [];
          if (weekFilled > 0) summaryParts.push(`${weekFilled} preenchido${weekFilled > 1 ? 's' : ''}`);
          if (weekOpen > 0) summaryParts.push(`${weekOpen} em aberto`);

          html += `
            <section class="week-block">
              <div class="week-header">
                <span class="week-title">${escapeHtml(wk.label)}</span>
                <span class="week-summary">${summaryParts.join(" · ")}</span>
              </div>
              <div class="card matrix-card">
                <div class="matrix-scroll">
                  <table class="roster-matrix">
                    <thead>
                      <tr>
                        <th class="col-slot" scope="col"><span class="eyebrow">Horário</span></th>
                        ${dayHeadersHtml}
                      </tr>
                    </thead>
                    <tbody>${rowsHtml}</tbody>
                  </table>
                </div>
              </div>
            </section>
          `;
        });

        contentArea.innerHTML = anyVisible
          ? html
          : `<div class="empty-state">Nenhum plantão encontrado com os filtros aplicados. Clique em <strong>Limpar</strong> para ver toda a escala.</div>`;
        bindDoctorPills();
      }

      function renderList(sc) {
        document.getElementById("viewHeading").textContent = `Lista Cronológica · ${sc.titulo} (${sc.mesAnoLabel})`;
        let html = '<div class="list-view">';
        let count = 0;

        sc.semanas.forEach(wk => {
          wk.dias.forEach(dia => {
            if (!dia.inMonth) return;
            const visibleGroups = dia.grupos.map(filterGroup).filter(Boolean);
            if (visibleGroups.length === 0) return;
            count++;

            const isToday = dia.iso === todayIso;
            const todayBadge = isToday ? `<span class="badge-yellow">HOJE</span>` : "";

            const rowsHtml = visibleGroups.map(fg => {
              const medsHtml = fg.medicos.map(m => {
                const sel = state.selectedDoctor && m.nome === state.selectedDoctor;
                const crm = m.crm ? `<span style="color:var(--text-muted); font-size:11px;"> · CRM ${escapeHtml(m.crm)}/${escapeHtml(m.uf)}</span>` : "";
                return `<button type="button" class="doc-pill ${sel ? 'is-selected' : ''}" data-doc="${escapeHtml(m.nome)}">${escapeHtml(m.nome)}${crm}</button>`;
              }).join("");

              let openHtml = "";
              if (fg.vagasAbertasVisiveis > 0) {
                const label = `${fg.vagasAbertasVisiveis} vaga${fg.vagasAbertasVisiveis > 1 ? 's' : ''} em aberto`;
                openHtml = fg.medicos.length > 0
                  ? `<div class="open-hint">+ ${label}</div>`
                  : `<div class="open-pill-empty">${label}</div>`;
              }

              return `
                <div class="slot-group">
                  <div class="slot-group-head">
                    <span class="shift-chip shift-${fg.periodo} slot-time">${fg.inicio} – ${fg.fim}</span>
                    <span style="font-size:11px; color:var(--text-muted);">${fg.horas}h</span>
                  </div>
                  ${medsHtml}
                  ${openHtml}
                </div>
              `;
            }).join("");

            html += `
              <article class="list-day-card">
                <div class="list-day-header">
                  <div>
                    <span class="eyebrow">${dia.weekdayFull}</span>
                    <div style="font-size:17px; font-weight:600; margin-top:2px;">${dia.ddmm}/${sc.refYear}</div>
                  </div>
                  ${todayBadge}
                </div>
                <div class="list-shift-rows">${rowsHtml}</div>
              </article>
            `;
          });
        });

        html += "</div>";
        contentArea.innerHTML = count > 0
          ? html
          : `<div class="empty-state">Nenhum plantão encontrado com os filtros aplicados.</div>`;
        bindDoctorPills();
      }

      function exportICS() {
        const sc = getActiveScale();
        if (!state.selectedDoctor) return;
        const shifts = sc.plantoes.filter(p => p.medico === state.selectedDoctor);
        if (shifts.length === 0) {
          alert("Este médico não possui plantões preenchidos nesta escala.");
          return;
        }

        const lines = [
          "BEGIN:VCALENDAR",
          "VERSION:2.0",
          "PRODID:-//Hapvida//Escala Medica Cruzeiro do Sul//PT-BR",
          "CALSCALE:GREGORIAN"
        ];

        shifts.forEach((s, i) => {
          const [y, m, d] = s.iso.split("-").map(Number);
          const [hi, mi] = s.inicio.split(":").map(Number);
          const [hf, mf] = s.fim.split(":").map(Number);

          const startDt = new Date(y, m - 1, d, hi, mi);
          let endDt = new Date(y, m - 1, d, hf, mf);
          if (endDt <= startDt) endDt.setDate(endDt.getDate() + 1);

          const fmt = dt => {
            const p = n => String(n).padStart(2, "0");
            return `${dt.getFullYear()}${p(dt.getMonth() + 1)}${p(dt.getDate())}T${p(dt.getHours())}${p(dt.getMinutes())}00`;
          };

          lines.push("BEGIN:VEVENT");
          lines.push(`UID:hapvida-${sc.id}-${s.iso}-${i}@cruzeirodosul`);
          lines.push(`DTSTART:${fmt(startDt)}`);
          lines.push(`DTEND:${fmt(endDt)}`);
          lines.push(`SUMMARY:Plantão ${sc.titulo} (${s.inicio}-${s.fim})`);
          lines.push(`LOCATION:${sc.unidade} - ${sc.municipio}`);
          lines.push(`DESCRIPTION:Escala ${sc.titulo} (${sc.mesAnoLabel}) - ${s.medico}`);
          lines.push("END:VEVENT");
        });

        lines.push("END:VCALENDAR");
        const blob = new Blob([lines.join("\r\n")], { type: "text/calendar;charset=utf-8" });
        const a = document.createElement("a");
        a.href = URL.createObjectURL(blob);
        const safeName = state.selectedDoctor.toLowerCase().replace(/[^a-z0-9]+/g, "_");
        a.download = `escala_${sc.id}_${safeName}.ics`;
        a.click();
        URL.revokeObjectURL(a.href);
      }

      function setView(v) {
        state.view = v;
        document.querySelectorAll("[data-view]").forEach(btn => {
          btn.classList.toggle("is-active", btn.getAttribute("data-view") === v);
        });
        renderAll();
      }

      function renderAll() {
        renderTabs();
        if (state.activeTab === "requests") {
          filterCard.style.display = "none";
          doctorBanner.classList.remove("is-visible");
          legendBar.style.display = "none";
          renderRequestsWorkspace();
          return;
        }
        filterCard.style.display = "";
        legendBar.style.display = "";

        const sc = getActiveScale();
        if (!sc) {
          contentArea.innerHTML = `<div class="empty-state">Nenhuma escala válida encontrada na planilha.</div>`;
          return;
        }
        renderHeaderInfo(sc);
        renderDoctorBanner(sc);
        if (state.view === "list") renderList(sc);
        else renderGrid(sc);
      }

      // Listeners
      doctorSelect.addEventListener("change", () => {
        state.selectedDoctor = doctorSelect.value;
        if (!state.selectedDoctor) {
          state.onlyDoctor = false;
          onlyDoctorCheck.checked = false;
        }
        renderAll();
      });

      shiftSelect.addEventListener("change", () => {
        state.shiftFilter = shiftSelect.value;
        renderAll();
      });

      searchInput.addEventListener("input", () => {
        state.searchText = searchInput.value;
        renderAll();
      });

      onlyDoctorCheck.addEventListener("change", () => {
        state.onlyDoctor = onlyDoctorCheck.checked;
        renderAll();
      });

      onlyOpenCheck.addEventListener("change", () => {
        state.onlyOpen = onlyOpenCheck.checked;
        if (state.onlyOpen) {
          state.showOpen = true;
          showOpenCheck.checked = true;
        }
        renderAll();
      });

      showOpenCheck.addEventListener("change", () => {
        state.showOpen = showOpenCheck.checked;
        renderAll();
      });

      document.querySelectorAll("[data-view]").forEach(btn => {
        btn.addEventListener("click", () => setView(btn.getAttribute("data-view")));
      });

      document.getElementById("clearFiltersBtn").addEventListener("click", () => {
        state.selectedDoctor = "";
        state.searchText = "";
        state.shiftFilter = "todos";
        state.onlyDoctor = false;
        state.onlyOpen = false;
        state.showOpen = true;
        doctorSelect.value = "";
        shiftSelect.value = "todos";
        searchInput.value = "";
        onlyDoctorCheck.checked = false;
        onlyOpenCheck.checked = false;
        showOpenCheck.checked = true;
        renderAll();
      });

      document.getElementById("exportIcsBtn").addEventListener("click", exportICS);
      document.getElementById("printBtn").addEventListener("click", () => window.print());

      async function refreshLiveServerData() {
        if (window.location.protocol === "file:") return;
        try {
          const [sRes, eRes, rRes] = await Promise.all([
            fetch("/_auth/session", { cache: "no-store" }),
            fetch("/api/escala", { cache: "no-store" }),
            fetch("/api/requests", { cache: "no-store" })
          ]);
          if (sRes.ok) {
            SESSION = await sRes.json();
            if (SESSION.authenticated) {
              document.getElementById("logoutLink").style.display = "inline-flex";
              if (SESSION.role === "diretor") {
                document.getElementById("directorLink").style.display = "inline-flex";
              }
            }
          }
          if (eRes.ok) {
            DATA = await eRes.json();
            updateMetaTexts();
            populateDoctors();
          }
          if (rRes.ok) {
            const rJson = await rRes.json();
            REQUESTS = rJson.solicitacoes || [];
          }
        } catch (err) {
          // Fallback silencioso para arquivo local sem backend
        }
      }

      // Inicialização
      document.querySelectorAll("[data-view]").forEach(btn => {
        btn.classList.toggle("is-active", btn.getAttribute("data-view") === state.view);
      });
      populateDoctors();
      renderAll();
      refreshLiveServerData().then(() => renderAll());
    })();
  </script>
</body>
</html>
"""


def generate_html(excel_path, output_dir):
    payload = parse_workbook(excel_path)
    os.makedirs(output_dir, exist_ok=True)

    # Salva JSON auxiliar (útil para integrações futuras ou auditoria)
    json_path = os.path.join(output_dir, "dados_escala.json")
    with open(json_path, "w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, indent=2)

    # Escapa </script> por segurança ao embutir JSON no HTML
    safe_json = json.dumps(payload, ensure_ascii=False).replace("</", "<\\/")
    html_final = (
        HTML_TEMPLATE
        .replace("__HOSPITAL_TITLE__", payload["hospital"].title())
        .replace("__HAPVIDA_LOGO_SVG__", HAPVIDA_SVG_LOGO)
        .replace("__JSON_DATA__", safe_json)
    )

    html_path = os.path.join(output_dir, "index.html")
    with open(html_path, "w", encoding="utf-8") as f:
        f.write(html_final)

    return html_path, payload


def git_publish(repo_dir):
    """Faz commit e push apenas se houver alteração real no index.html / dados_escala.json."""
    def run_git(args):
        kwargs = {
            "cwd": repo_dir,
            "capture_output": True,
            "text": True,
            "encoding": "utf-8",
            "errors": "replace",
        }
        if os.name == "nt":
            kwargs["creationflags"] = getattr(subprocess, "CREATE_NO_WINDOW", 0)
        return subprocess.run(["git"] + args, **kwargs)

    status = run_git(["status", "--porcelain", "index.html", "dados_escala.json"])
    if status.returncode != 0:
        print(f"[GIT] Pasta não é um repositório Git válido: {status.stderr.strip()}")
        return False

    if not status.stdout.strip():
        print("[GIT] Nenhuma alteração detectada para publicar.")
        return False

    run_git(["add", "index.html", "dados_escala.json"])
    msg = f"chore(escala): atualiza escala medica ({datetime.now().strftime('%d/%m/%Y %H:%M')})"
    commit = run_git(["commit", "-m", msg])
    if commit.returncode != 0:
        print(f"[GIT] Aviso no commit: {commit.stderr.strip()}")
        return False

    push = run_git(["push"])
    if push.returncode == 0:
        print(f"[GIT] Publicado com sucesso no GitHub Pages! ({msg})")
        return True
    else:
        print(f"[GIT] Erro no push: {push.stderr.strip()}")
        return False


def main():
    parser = argparse.ArgumentParser(description="Gera e publica a Escala Médica HTML a partir da planilha Excel.")
    parser.add_argument("--excel", default=DEFAULT_EXCEL_PATH, help="Caminho completo da planilha .xlsx no OneDrive")
    parser.add_argument("--saida", default=DEFAULT_OUTPUT_DIR, help="Pasta do repositório Git onde index.html será salvo")
    parser.add_argument("--watch", action="store_true", help="Monitora continuamente alterações na planilha")
    parser.add_argument("--interval", type=int, default=DEFAULT_WATCH_INTERVAL, help="Intervalo em segundos no modo --watch")
    parser.add_argument("--git-push", action="store_true", help="Faz git commit + push automaticamente após gerar")
    args = parser.parse_args()

    if sys.stdout is None or sys.stderr is None:
        log_path = os.path.join(args.saida, "monitor_escala.log")
        log_fh = open(log_path, "a", encoding="utf-8", buffering=1)
        if sys.stdout is None:
            sys.stdout = log_fh
        if sys.stderr is None:
            sys.stderr = log_fh

    def run_once():
        html_path, payload = generate_html(args.excel, args.saida)
        print(f"[OK] {datetime.now().strftime('%H:%M:%S')} | Gerado: {html_path} ({len(payload['escalas'])} escalas)")
        for sc in payload["escalas"]:
            k = sc["kpis"]
            print(f"     - {sc['tabLabel']}: {k['preenchidas']}/{k['totalVagas']} ({k['coberturaPct']}%) | {k['totalHoras']}h | {k['medicosEscalados']} médicos")
        if args.git_push:
            git_publish(args.saida)

    if not args.watch:
        run_once()
        return

    print(f"[WATCH] Monitorando planilha: {args.excel}")
    print(f"[WATCH] Destino: {args.saida} | Git Push: {'ATIVO' if args.git_push else 'DESATIVADO'}")
    print("[WATCH] Pressione Ctrl+C para encerrar.\n")

    last_mtime = None
    while True:
        try:
            if os.path.exists(args.excel):
                current_mtime = os.path.getmtime(args.excel)
                if last_mtime is None or current_mtime != last_mtime:
                    if last_mtime is not None:
                        # Aguarda 3s para garantir que o OneDrive/Excel terminou de gravar
                        time.sleep(3)
                    run_once()
                    last_mtime = os.path.getmtime(args.excel)
            else:
                print(f"[AVISO] Planilha não encontrada no momento: {args.excel}")
        except KeyboardInterrupt:
            print("\n[WATCH] Monitoramento encerrado pelo usuário.")
            break
        except Exception as e:
            print(f"[ERRO] Falha ao processar planilha (tentará novamente): {e}")

        time.sleep(args.interval)


if __name__ == "__main__":
    main()
