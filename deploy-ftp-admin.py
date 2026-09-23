#!/usr/bin/env python3
"""
Sobe arquivos do Portal Admin (~/Public/portal-admin-deploy/) pro FTP da
Locaweb (/public_html/admin/). Credenciais vêm só de env vars (Railway) —
nunca hardcoded aqui, nunca impressas.

Uso:
  railway run python3 deploy-ftp-admin.py                # sobe tudo
  railway run python3 deploy-ftp-admin.py admin.html sw.js  # só esses arquivos
"""
import os
import sys
from ftplib import FTP, FTP_TLS

ORIGEM = "/Users/fabiospila/Public/portal-admin-deploy"
DESTINO_REMOTO = "public_html/admin"
IGNORAR = {".DS_Store"}

HOST = os.environ["FTP_ADMIN_HOST"]
USER = os.environ["FTP_ADMIN_USER"]
PASS = os.environ["FTP_ADMIN_PASS"]
PORTA = int(os.environ.get("FTP_ADMIN_PORT", "21"))


def conectar():
    """Tenta FTPS explícito (AUTH TLS) primeiro; cai pra FTP puro se o
    servidor recusar — evita ter que descobrir na mão qual a Locaweb usa."""
    try:
        ftp = FTP_TLS()
        ftp.connect(HOST, PORTA, timeout=30)
        ftp.login(USER, PASS)
        ftp.prot_p()
        print("Conectado via FTPS (TLS).")
        return ftp
    except Exception as e:
        print(f"FTPS falhou ({e}), tentando FTP simples...")
        ftp = FTP()
        ftp.connect(HOST, PORTA, timeout=30)
        ftp.login(USER, PASS)
        print("Conectado via FTP simples.")
        return ftp


def subir_arquivo(ftp, nome):
    caminho_local = os.path.join(ORIGEM, nome)
    with open(caminho_local, "rb") as f:
        ftp.storbinary(f"STOR {nome}", f)
    print(f"  ok: {nome}")


def main():
    arquivos = sys.argv[1:]
    if not arquivos:
        arquivos = sorted(
            f for f in os.listdir(ORIGEM)
            if f not in IGNORAR and os.path.isfile(os.path.join(ORIGEM, f))
        )

    faltando = [n for n in arquivos if not os.path.isfile(os.path.join(ORIGEM, n))]
    for nome in faltando:
        print(f"  pulando (não existe localmente): {nome}")
    arquivos = [n for n in arquivos if n not in faltando]

    print(f"Subindo {len(arquivos)} arquivo(s) para {DESTINO_REMOTO}/ em {HOST}...")

    ftp = conectar()
    try:
        ftp.cwd(DESTINO_REMOTO)
        for nome in arquivos:
            subir_arquivo(ftp, nome)
    finally:
        ftp.quit()

    print("Concluído.")


if __name__ == "__main__":
    main()
