#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
PSS - Script de Empacotamento Portavel

Gera uma distribuicao portavel usando Python embarcado (embeddable package).
O resultado e uma pasta auto-contida que roda em qualquer Windows 10+
sem necessidade de instalar Python.

Uso:
    python empacotar.py

O script baixa automaticamente o Python embeddable do python.org se necessario.
Tambem aceita um .zip ja baixado na pasta python_embarcado/.
"""

import os
import sys
import shutil
import subprocess
import zipfile
import urllib.request
from pathlib import Path


# ============================================================
# CONFIGURACAO
# ============================================================
PROJETO_NOME = "PSS"
DIST_DIR = Path("dist") / PROJETO_NOME
PYTHON_EMB_DIR = Path("python_embarcado")
INTEGRA_PKG_PATH = Path.home() / "PycharmProjects" / "integra"
REQUIREMENTS_FILE = "requirements_dist.txt"
GET_PIP_URL = "https://bootstrap.pypa.io/get-pip.py"

# Versao do Python embeddable a baixar automaticamente
PYTHON_VERSION = "3.12.8"
PYTHON_EMBED_URL = (
    "https://www.python.org/ftp/python/{v}/python-{v}-embed-amd64.zip"
    .format(v=PYTHON_VERSION)
)

# Arquivos do projeto a copiar para distribuicao
ARQUIVOS_PROJETO = [
    "main.py",
    "estado_processos.py",
]

# Pastas do projeto a copiar para distribuicao
PASTAS_PROJETO = [
    "docs_pss",
    "planilhas",
]

# Pastas de trabalho (criadas vazias na distribuicao)
PASTAS_TRABALHO = [
    "logs",
    "temporario",
    "estado",
]


def log(msg, nivel="info"):
    """Log formatado"""
    prefixos = {
        "info": "  [OK]",
        "warn": "  [!!]",
        "error": "  [XX]",
        "step": "\n===",
    }
    prefixo = prefixos.get(nivel, "  ")
    if nivel == "step":
        print("{} {} ===".format(prefixo, msg))
    else:
        print("{} {}".format(prefixo, msg))


def _progresso_download(blocos, tam_bloco, tam_total):
    """Callback para mostrar progresso do download"""
    baixado = blocos * tam_bloco
    if tam_total > 0:
        pct = min(100, baixado * 100 // tam_total)
        mb = baixado / (1024 * 1024)
        total_mb = tam_total / (1024 * 1024)
        print("\r  [..] Baixando: {:.1f}/{:.1f} MB ({}%)".format(mb, total_mb, pct), end="", flush=True)


def baixar_python_embarcado():
    """Baixa o Python embeddable package do python.org

    Returns:
        Path ou None: Caminho do zip baixado
    """
    PYTHON_EMB_DIR.mkdir(exist_ok=True)
    nome_zip = "python-{}-embed-amd64.zip".format(PYTHON_VERSION)
    destino = PYTHON_EMB_DIR / nome_zip

    log("Baixando Python {} embeddable do python.org...".format(PYTHON_VERSION))
    log("URL: {}".format(PYTHON_EMBED_URL))

    try:
        urllib.request.urlretrieve(PYTHON_EMBED_URL, str(destino), _progresso_download)
        print()  # Nova linha apos progresso
        log("Download concluido: {}".format(nome_zip))
        return destino
    except Exception as e:
        log("Erro ao baixar Python embeddable: {}".format(e), "error")
        log("Baixe manualmente de: {}".format(PYTHON_EMBED_URL), "warn")
        log("Salve em: {}".format(PYTHON_EMB_DIR.absolute()), "warn")
        return None


def encontrar_python_embarcado():
    """Localiza o Python embarcado (zip local) ou baixa automaticamente

    Returns:
        Path ou None: Caminho do zip do Python embarcado
    """
    log("Localizando Python embarcado", "step")

    PYTHON_EMB_DIR.mkdir(exist_ok=True)

    # 1. Procurar zip na pasta python_embarcado/
    zips = list(PYTHON_EMB_DIR.glob("python-*-embed-amd64.zip"))
    if zips:
        log("Encontrado zip local: {}".format(zips[0].name))
        return zips[0]

    # 2. Nao encontrou - baixar automaticamente
    log("Nenhum zip encontrado em {}".format(PYTHON_EMB_DIR), "warn")
    log("Baixando automaticamente...")
    return baixar_python_embarcado()


def limpar_distribuicao():
    """Remove distribuicao anterior"""
    log("Limpando distribuicao anterior", "step")

    if DIST_DIR.exists():
        try:
            shutil.rmtree(DIST_DIR)
            log("Removida: {}".format(DIST_DIR))
        except Exception as e:
            log("Erro ao remover {}: {}".format(DIST_DIR, e), "error")
            log("Feche programas que possam estar usando arquivos dessa pasta", "warn")
            return False

    DIST_DIR.mkdir(parents=True, exist_ok=True)
    log("Criada: {}".format(DIST_DIR))
    return True


def preparar_python(zip_path):
    """Extrai Python embarcado do zip para a distribuicao

    Args:
        zip_path: Path do zip do Python embarcado
    """
    log("Preparando Python embarcado", "step")

    python_dir = DIST_DIR / "python"
    python_dir.mkdir(exist_ok=True)

    log("Extraindo de: {}".format(zip_path.name))
    with zipfile.ZipFile(zip_path, 'r') as zf:
        zf.extractall(python_dir)

    total = len(list(python_dir.rglob("*")))
    log("{} arquivos na pasta python/".format(total))
    return python_dir


def configurar_pth(python_dir):
    """Configura o arquivo ._pth para habilitar site-packages"""
    log("Configurando Python path (._pth)", "step")

    # Encontrar o arquivo ._pth (ex: python312._pth)
    pth_files = list(python_dir.glob("python*._pth"))
    if not pth_files:
        log("Arquivo ._pth nao encontrado!", "error")
        log("Verifique se o zip e um Python embeddable valido", "warn")
        return False

    pth_file = pth_files[0]
    log("Arquivo encontrado: {}".format(pth_file.name))

    # Extrair nome base (ex: python312)
    base_name = pth_file.name.replace("._pth", "")

    # Reescrever com site-packages habilitado
    conteudo = "{}.zip\n.\n..\nLib\nLib\\site-packages\nimport site\n".format(base_name)

    with open(pth_file, 'w') as f:
        f.write(conteudo)

    log("._pth configurado com site-packages habilitado")

    # Criar diretorio Lib/site-packages
    site_packages = python_dir / "Lib" / "site-packages"
    site_packages.mkdir(parents=True, exist_ok=True)
    log("Criado: Lib/site-packages/")

    return True


def instalar_pip(python_dir):
    """Instala pip no Python embarcado"""
    log("Instalando pip", "step")

    python_exe = python_dir / "python.exe"
    get_pip_path = python_dir / "get-pip.py"

    # Baixar get-pip.py
    log("Baixando get-pip.py...")
    try:
        urllib.request.urlretrieve(GET_PIP_URL, str(get_pip_path))
        log("get-pip.py baixado")
    except Exception as e:
        log("Erro ao baixar get-pip.py: {}".format(e), "error")
        return False

    # Executar get-pip.py
    log("Executando get-pip.py...")
    try:
        result = subprocess.run(
            [str(python_exe), str(get_pip_path), "--no-warn-script-location"],
            capture_output=True, text=True, encoding='utf-8', errors='replace'
        )
        if result.returncode != 0:
            log("Erro na instalacao do pip:", "error")
            print(result.stderr[:500])
            return False
        log("pip instalado com sucesso")
    except Exception as e:
        log("Erro ao executar get-pip.py: {}".format(e), "error")
        return False
    finally:
        # Remover get-pip.py
        if get_pip_path.exists():
            get_pip_path.unlink()

    # Instalar setuptools e wheel (necessarios para compilar pacotes)
    log("Instalando setuptools e wheel...")
    try:
        result = subprocess.run(
            [
                str(python_exe), "-m", "pip", "install",
                "setuptools", "wheel",
                "--no-warn-script-location",
                "--disable-pip-version-check",
            ],
            capture_output=True, text=True, encoding='utf-8', errors='replace'
        )
        if result.returncode != 0:
            log("Erro ao instalar setuptools/wheel:", "error")
            print(result.stderr[:500])
            return False
        log("setuptools e wheel instalados")
    except Exception as e:
        log("Erro: {}".format(e), "error")
        return False

    return True


def instalar_dependencias(python_dir):
    """Instala dependencias via pip no Python embarcado"""
    log("Instalando dependencias", "step")

    python_exe = python_dir / "python.exe"
    req_file = Path(REQUIREMENTS_FILE)

    if not req_file.exists():
        log("{} nao encontrado!".format(REQUIREMENTS_FILE), "error")
        return False

    log("Instalando de {}...".format(REQUIREMENTS_FILE))
    log("Isso pode levar alguns minutos...")

    try:
        result = subprocess.run(
            [
                str(python_exe), "-m", "pip", "install",
                "-r", str(req_file),
                "--no-warn-script-location",
                "--disable-pip-version-check",
            ],
            capture_output=True, text=True, encoding='utf-8', errors='replace'
        )

        if result.returncode != 0:
            log("Erro ao instalar dependencias:", "error")
            # Mostrar ultimas linhas do erro
            linhas_erro = result.stderr.strip().split('\n')
            for linha in linhas_erro[-10:]:
                print("    {}".format(linha))
            return False

        log("Dependencias instaladas com sucesso")

    except Exception as e:
        log("Erro: {}".format(e), "error")
        return False

    return True


def instalar_pacote_integra(python_dir):
    """Instala o pacote integra (pai) no Python embarcado"""
    log("Instalando pacote integra", "step")

    python_exe = python_dir / "python.exe"
    integra_path = INTEGRA_PKG_PATH.resolve()

    if not (integra_path / "integra" / "__init__.py").exists():
        log("Pacote integra nao encontrado em: {}".format(integra_path), "error")
        log("Certifique-se de que ../integra existe e contem o pacote", "warn")
        return False

    log("Instalando de: {}".format(integra_path))

    try:
        result = subprocess.run(
            [
                str(python_exe), "-m", "pip", "install",
                str(integra_path),
                "--no-warn-script-location",
                "--disable-pip-version-check",
            ],
            capture_output=True, text=True, encoding='utf-8', errors='replace'
        )

        if result.returncode != 0:
            log("Erro ao instalar pacote integra:", "error")
            linhas_erro = result.stderr.strip().split('\n')
            for linha in linhas_erro[-10:]:
                print("    {}".format(linha))
            return False

        log("Pacote integra instalado com sucesso")

    except Exception as e:
        log("Erro: {}".format(e), "error")
        return False

    return True


def copiar_arquivos_projeto():
    """Copia arquivos do projeto para a distribuicao"""
    log("Copiando arquivos do projeto", "step")

    # Copiar arquivos .py individuais
    copiados = 0
    for arquivo in ARQUIVOS_PROJETO:
        src = Path(arquivo)
        if src.exists():
            dst = DIST_DIR / arquivo
            shutil.copy2(src, dst)
            copiados += 1
        else:
            log("{} nao encontrado (pulando)".format(arquivo), "warn")

    log("{} arquivos .py copiados".format(copiados))

    # Copiar pastas do projeto
    for pasta in PASTAS_PROJETO:
        src = Path(pasta)
        if src.exists():
            dst = DIST_DIR / pasta
            if dst.exists():
                shutil.rmtree(dst)
            shutil.copytree(src, dst)
            total = len(list(dst.rglob("*")))
            log("{} ({} itens)".format(pasta, total))
        else:
            log("{} nao encontrada (pulando)".format(pasta), "warn")

    # Criar pastas de trabalho vazias
    for pasta in PASTAS_TRABALHO:
        pasta_path = DIST_DIR / pasta
        pasta_path.mkdir(parents=True, exist_ok=True)

    log("Pastas de trabalho criadas")
    return True


def criar_launcher():
    """Cria os arquivos .bat de lancamento"""
    log("Criando launchers", "step")

    # PSS.bat - launcher GUI (sem console, usa pythonw.exe)
    bat_gui = DIST_DIR / "PSS.bat"
    conteudo_gui = (
        '@echo off\r\n'
        'pushd "%~dp0"\r\n'
        'start "" "python\\pythonw.exe" main.py\r\n'
        'popd\r\n'
    )
    with open(bat_gui, 'w', encoding='ascii') as f:
        f.write(conteudo_gui)
    log("Criado: PSS.bat (GUI, sem console)")

    # PSS_debug.bat - com console para ver erros
    bat_debug = DIST_DIR / "PSS_debug.bat"
    conteudo_debug = (
        '@echo off\r\n'
        'title PSS - Debug\r\n'
        'pushd "%~dp0"\r\n'
        'python\\python.exe main.py\r\n'
        'popd\r\n'
        'pause\r\n'
    )
    with open(bat_debug, 'w', encoding='ascii') as f:
        f.write(conteudo_debug)
    log("Criado: PSS_debug.bat (com console para debug)")

    return True


def criar_leiame():
    """Cria arquivo LEIA-ME.txt"""
    log("Criando LEIA-ME.txt", "step")

    conteudo = (
        "PSS - Automacao do Programa de Seguridade Social\r\n"
        "=================================================\r\n"
        "\r\n"
        "COMO EXECUTAR:\r\n"
        "  De duplo-clique em PSS.bat\r\n"
        "  (ou PSS_debug.bat para ver erros no console)\r\n"
        "\r\n"
        "REQUISITOS:\r\n"
        "  - Windows 10 ou superior\r\n"
        "  - Nao e necessario instalar Python\r\n"
        "  - Google Chrome instalado (para automacao web)\r\n"
        "  - Certificado digital configurado (para SIAPE)\r\n"
        "\r\n"
        "ESTRUTURA DE PASTAS:\r\n"
        "  python/          - Interpretador Python (nao alterar)\r\n"
        "  docs_pss/        - Documentos PDF para anexar nos processos\r\n"
        "  planilhas/       - Planilha de dados dos pensionistas\r\n"
        "  logs/            - Logs de execucao\r\n"
        "\r\n"
        "IMPORTANTE:\r\n"
        "  - Mantenha TODA a pasta intacta para o programa funcionar\r\n"
        "  - NAO mova apenas o .bat para outro local\r\n"
        "  - Para atualizar, substitua os arquivos .py pela versao nova\r\n"
    )

    leiame_path = DIST_DIR / "LEIA-ME.txt"
    with open(leiame_path, 'w', encoding='utf-8') as f:
        f.write(conteudo)

    log("Criado: LEIA-ME.txt")
    return True


def verificar_resultado():
    """Verifica a distribuicao gerada"""
    log("Verificando resultado", "step")

    # Verificar Python
    python_exe = DIST_DIR / "python" / "python.exe"
    if python_exe.exists():
        log("python/python.exe presente")
    else:
        log("python/python.exe AUSENTE!", "error")
        return False

    # Verificar main.py
    main_py = DIST_DIR / "main.py"
    if main_py.exists():
        log("main.py presente")
    else:
        log("main.py AUSENTE!", "error")
        return False

    # Verificar launcher
    bat_file = DIST_DIR / "PSS.bat"
    if bat_file.exists():
        log("PSS.bat presente")
    else:
        log("PSS.bat AUSENTE!", "error")
        return False

    # Verificar pacote integra
    try:
        result = subprocess.run(
            [str(python_exe), "-c", "import integra; print('OK')"],
            capture_output=True, text=True, encoding='utf-8', errors='replace',
            cwd=str(DIST_DIR)
        )
        if "OK" in result.stdout:
            log("Pacote 'integra' importa corretamente")
        else:
            log("Pacote 'integra' falhou ao importar", "warn")
            if result.stderr:
                print("    {}".format(result.stderr[:300]))
    except Exception as e:
        log("Erro ao verificar integra: {}".format(e), "warn")

    # Tamanho total
    tamanho_total = sum(
        f.stat().st_size for f in DIST_DIR.rglob("*") if f.is_file()
    ) / (1024 * 1024)
    log("Tamanho total: {:.1f} MB".format(tamanho_total))

    return True


def gerar_relatorio():
    """Gera relatorio final"""
    print("\n" + "=" * 60)
    print("  PSS - Distribuicao Portavel Gerada!")
    print("=" * 60)
    print("")
    print("  Local: {}".format(DIST_DIR.absolute()))
    print("")
    print("  Para executar:")
    print("    {}\\PSS.bat".format(DIST_DIR))
    print("")
    print("  Para distribuir:")
    print("    Compacte toda a pasta {}".format(DIST_DIR))
    print("")
    print("  Para atualizar codigo:")
    print("    Copie os .py alterados para {}".format(DIST_DIR))
    print("    OU execute este script novamente")
    print("")


def main():
    print("=" * 60)
    print("  PSS - Empacotamento Portavel")
    print("  (Python Embarcado + Batch Launcher)")
    print("=" * 60)

    # 1. Encontrar ou baixar Python embarcado
    zip_path = encontrar_python_embarcado()
    if not zip_path:
        return

    # 2. Limpar distribuicao anterior
    if not limpar_distribuicao():
        return

    # 3. Extrair Python embarcado
    python_dir = preparar_python(zip_path)

    # 4. Configurar ._pth
    if not configurar_pth(python_dir):
        return

    # 5. Instalar pip
    if not instalar_pip(python_dir):
        return

    # 6. Instalar dependencias
    if not instalar_dependencias(python_dir):
        return

    # 7. Instalar pacote integra
    if not instalar_pacote_integra(python_dir):
        return

    # 8. Copiar arquivos do projeto
    if not copiar_arquivos_projeto():
        return

    # 9. Criar launcher e leia-me
    criar_launcher()
    criar_leiame()

    # 10. Verificar
    if verificar_resultado():
        gerar_relatorio()
    else:
        log("Distribuicao gerada com problemas!", "error")


if __name__ == "__main__":
    main()
