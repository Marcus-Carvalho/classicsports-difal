# -*- coding: utf-8 -*-
"""
Classic Sports - Guias DIFAL
Publicacao de atualizacao.

Roda na maquina do Marcus, na pasta do programa:

    python publicar_atualizacao.py            -> sobe a versao de correcao (1.1.0 -> 1.1.1)
    python publicar_atualizacao.py menor      -> 1.1.3 -> 1.2.0
    python publicar_atualizacao.py maior      -> 1.2.5 -> 2.0.0
    python publicar_atualizacao.py --notas "texto que aparece no painel"

O que ele faz:
  1. verifica a sintaxe de todos os .py
  2. confere se nenhum arquivo esta vazio
  3. incrementa o versao.json
  4. copia os arquivos para a pasta do GitHub Desktop

Depois: GitHub Desktop -> Summary -> Commit to main -> Push origin.
"""

import os
import sys
import json
import shutil
import py_compile
from pathlib import Path
from datetime import datetime

PASTA = Path(__file__).resolve().parent

# pasta local do repositorio (GitHub Desktop)
PASTA_GITHUB = Path(os.path.expanduser("~")) / "Documents" / "GitHub" / "classicsports-difal"
PASTA_GITHUB_PT = Path(os.path.expanduser("~")) / "Documentos" / "GitHub" / "classicsports-difal"

ARQUIVOS = [
    "difal_core.py",
    "difal_painel.py",
    "atualizador.py",
    "instalador_gui.py",
    "publicar_atualizacao.py",
    "build_instalador.bat",
    "modelo_c6.xlsx",
    "icon.ico",
    "LEIAME.txt",
    "ClassicSportsDIFAL.bat",
    "versao.json",
]

# o que o app baixa numa atualizacao (o instalador nao precisa ser baixado)
ARQUIVOS_ATUALIZACAO = [
    "difal_core.py",
    "difal_painel.py",
    "atualizador.py",
    "modelo_c6.xlsx",
    "icon.ico",
    "LEIAME.txt",
    "ClassicSportsDIFAL.bat",
    "versao.json",
]


def achar_repositorio():
    for caminho in (PASTA_GITHUB, PASTA_GITHUB_PT):
        if caminho.exists():
            return caminho
    return None


def nova_versao(atual, tipo):
    partes = [int(p) for p in str(atual).split(".")[:3]]
    while len(partes) < 3:
        partes.append(0)
    maior, menor, correcao = partes
    if tipo == "maior":
        return str(maior + 1) + ".0.0"
    if tipo == "menor":
        return str(maior) + "." + str(menor + 1) + ".0"
    return str(maior) + "." + str(menor) + "." + str(correcao + 1)


def main():
    tipo = "correcao"
    notas = ""
    argumentos = sys.argv[1:]
    if argumentos and argumentos[0] in ("maior", "menor", "correcao"):
        tipo = argumentos.pop(0)
    if "--notas" in argumentos:
        indice = argumentos.index("--notas")
        if indice + 1 < len(argumentos):
            notas = argumentos[indice + 1]

    print("== Verificando os arquivos ==")
    problemas = []
    for nome in ARQUIVOS:
        arquivo = PASTA / nome
        if not arquivo.exists():
            problemas.append("faltando: " + nome)
            continue
        if arquivo.stat().st_size == 0:
            problemas.append("arquivo vazio: " + nome)
            continue
        if nome.endswith(".py"):
            try:
                py_compile.compile(str(arquivo), doraise=True)
            except Exception as erro:
                problemas.append("erro de sintaxe em " + nome + ": " + str(erro))
        print("  ok  " + nome)

    if problemas:
        print("\nPUBLICACAO CANCELADA:")
        for p in problemas:
            print("  - " + p)
        return 1

    repositorio = achar_repositorio()
    if repositorio is None:
        print("\nPasta do GitHub Desktop nao encontrada em:")
        print("  " + str(PASTA_GITHUB))
        print("  " + str(PASTA_GITHUB_PT))
        return 1

    arquivo_versao = PASTA / "versao.json"
    dados = json.loads(arquivo_versao.read_text(encoding="utf-8"))
    atual = dados.get("versao", "1.0.0")
    proxima = nova_versao(atual, tipo)

    dados["versao"] = proxima
    dados["arquivos"] = ARQUIVOS_ATUALIZACAO
    dados["publicado_em"] = datetime.now().strftime("%d/%m/%Y %H:%M")
    if notas:
        dados["notas"] = notas
    arquivo_versao.write_text(json.dumps(dados, ensure_ascii=False, indent=2), encoding="utf-8")
    print("\nVersao " + atual + " -> " + proxima)

    print("\n== Copiando para o repositorio ==")
    for nome in ARQUIVOS:
        shutil.copy2(str(PASTA / nome), str(repositorio / nome))
        print("  " + nome)

    print("\nPronto. Agora no GitHub Desktop:")
    print("  Summary: versao " + proxima + (" - " + notas if notas else ""))
    print("  Commit to main -> Push origin")
    return 0


if __name__ == "__main__":
    sys.exit(main())
