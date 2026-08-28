# -*- coding: utf-8 -*-
"""
Classic Sports - Guias DIFAL
Atualizacao automatica via GitHub (mesmo padrao do app de Relatorios).

O versao.json do repositorio manda: a versao publicada e a lista de arquivos
que devem ser baixados. Arquivos do usuario (config, planilhas, tmp) nunca
sao tocados.
"""

import json
import shutil
import tempfile
import urllib.error
import urllib.request
from pathlib import Path

PASTA = Path(__file__).resolve().parent

REPOSITORIO = "Marcus-Carvalho/classicsports-difal"
URL_RAW_BASE = "https://raw.githubusercontent.com/" + REPOSITORIO + "/main/"

# arquivos que nao podem ser sobrescritos pela atualizacao
PROTEGIDOS = {"config_difal.json"}

# baixados em modo binario (nao sao texto)
BINARIOS = {".xlsx", ".ico", ".png", ".exe", ".zip"}

ARQUIVOS_PADRAO = [
    "difal_core.py",
    "difal_painel.py",
    "atualizador.py",
    "modelo_c6.xlsx",
    "icon.ico",
    "LEIAME.txt",
    "ClassicSportsDIFAL.bat",
    "versao.json",
]


def ler_versao_local():
    arquivo = PASTA / "versao.json"
    if not arquivo.exists():
        return "1.0.0"
    try:
        return json.loads(arquivo.read_text(encoding="utf-8")).get("versao", "1.0.0")
    except Exception:
        return "1.0.0"


def _baixar(nome, tempo=30):
    url = URL_RAW_BASE + nome + "?cb=" + str(abs(hash(nome)) % 100000)
    requisicao = urllib.request.Request(url, headers={"Cache-Control": "no-cache"})
    with urllib.request.urlopen(requisicao, timeout=tempo) as resposta:
        return resposta.read()


def _tupla(versao):
    partes = []
    for pedaco in str(versao).split("."):
        try:
            partes.append(int(pedaco))
        except ValueError:
            partes.append(0)
    while len(partes) < 3:
        partes.append(0)
    return tuple(partes[:3])


def consultar():
    """
    Consulta o GitHub. Devolve (tem_novidade, versao_remota, dados_remotos, erro).
    """
    try:
        dados = json.loads(_baixar("versao.json").decode("utf-8"))
    except urllib.error.HTTPError as erro:
        if erro.code == 404:
            return False, None, None, ("o repositório " + REPOSITORIO +
                                       " ainda não existe ou está privado. "
                                       "O programa funciona normalmente; só a "
                                       "atualização automática fica indisponível.")
        return False, None, None, "HTTP " + str(erro.code)
    except Exception as erro:
        return False, None, None, str(erro)
    remota = dados.get("versao", "0.0.0")
    return (_tupla(remota) > _tupla(ler_versao_local())), remota, dados, None


def atualizar(dados_remotos=None, progresso=None):
    """
    Baixa os arquivos publicados. So grava na pasta do app depois que TODOS
    baixaram, para nunca deixar a instalacao pela metade.
    Devolve (ok, mensagem).
    """
    def aviso(texto):
        if progresso:
            try:
                progresso(texto)
            except Exception:
                pass

    if dados_remotos is None:
        _, _, dados_remotos, erro = consultar()
        if dados_remotos is None:
            return False, "Não foi possível consultar o GitHub: " + str(erro)

    arquivos = [a for a in dados_remotos.get("arquivos", ARQUIVOS_PADRAO)
                if a not in PROTEGIDOS]
    temporaria = Path(tempfile.mkdtemp(prefix="difal_update_"))
    baixados = []

    try:
        for nome in arquivos:
            aviso("Baixando " + nome + "...")
            conteudo = _baixar(nome)
            if not conteudo:
                raise IOError("arquivo vazio: " + nome)
            destino = temporaria / nome
            destino.parent.mkdir(parents=True, exist_ok=True)
            destino.write_bytes(conteudo)
            baixados.append(nome)

        # validacao de sintaxe antes de instalar
        import py_compile
        for nome in baixados:
            if nome.endswith(".py"):
                py_compile.compile(str(temporaria / nome), doraise=True)

        aviso("Instalando...")
        for nome in baixados:
            shutil.copy2(str(temporaria / nome), str(PASTA / nome))

        versao = dados_remotos.get("versao", ler_versao_local())
        return True, "Atualizado para a versão " + str(versao) + ". Feche e abra o programa."
    except Exception as erro:
        return False, "Falha na atualização (nada foi alterado): " + str(erro)
    finally:
        shutil.rmtree(temporaria, ignore_errors=True)


if __name__ == "__main__":
    tem, remota, dados, erro = consultar()
    if erro:
        print("erro:", erro)
    elif tem:
        print("nova versao:", remota)
        print(atualizar(dados, progresso=print))
    else:
        print("ja esta na versao mais recente:", ler_versao_local())
