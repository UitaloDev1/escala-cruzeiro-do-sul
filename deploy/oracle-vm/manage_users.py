#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Utilitário para criar ou atualizar usuários (medico / diretor) em /etc/escala-cruzeiro-do-sul/users.json."""

import argparse
import base64
import hashlib
import json
import os
import secrets
from pathlib import Path

CONFIG_DIR = Path(os.environ.get("ESCALA_CONFIG_DIR", "/etc/escala-cruzeiro-do-sul"))
USERS_FILE = CONFIG_DIR / "users.json"


def b64u_encode(data: bytes) -> str:
    return base64.urlsafe_b64encode(data).rstrip(b"=").decode("ascii")


def hash_password(password: str) -> tuple[str, str]:
    salt = secrets.token_bytes(16)
    derived = hashlib.scrypt(password.encode("utf-8"), salt=salt, n=16384, r=8, p=1, dklen=32)
    return b64u_encode(salt), b64u_encode(derived)


def main():
    parser = argparse.ArgumentParser(description="Gerencia usuários da Escala Médica")
    parser.add_argument("--username", required=True, help="Nome de usuário (ex: medico ou diretor)")
    parser.add_argument("--password", required=True, help="Senha do usuário")
    parser.add_argument("--role", choices=["medico", "diretor"], required=True, help="Perfil de acesso")
    parser.add_argument("--display-name", default="", help="Nome de exibição")
    args = parser.parse_args()

    CONFIG_DIR.mkdir(parents=True, exist_ok=True)
    os.chmod(CONFIG_DIR, 0o700)

    users = {}
    if USERS_FILE.exists():
        users = json.loads(USERS_FILE.read_text(encoding="utf-8"))

    salt, pw_hash = hash_password(args.password)
    uname = args.username.strip().lower()
    users[uname] = {
        "role": args.role,
        "displayName": args.display_name or ("Diretoria Médica" if args.role == "diretor" else "Corpo Clínico"),
        "salt": salt,
        "hash": pw_hash,
    }

    USERS_FILE.write_text(json.dumps(users, ensure_ascii=False, indent=2), encoding="utf-8")
    os.chmod(USERS_FILE, 0o600)
    print(f"[OK] Usuário '{uname}' ({args.role}) salvo em {USERS_FILE}")


if __name__ == "__main__":
    main()
