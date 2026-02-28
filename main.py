"""
PSS - Automação do Sistema Eletrônico de Informações (SEI)

Interface gráfica com:
- Campos de login e senha protegidos
- Credenciais criptografadas (AES/Fernet)
- Log de ações em tempo real
"""

import sys
import time
import json
import threading
import urllib.request
from pathlib import Path

import pandas as pd
from PyQt5.QtWidgets import QApplication

from integra.interface import JanelaIntegraBase, ConfiguracaoInterface
from integra.browser import SeleniumSetup, ChromeDebug
from integra.sei import (
    LoginSei, TelaAviso, SelecaoUnidadeDireta,
    IniciaProcessos, NumeroProcesso, TrocaMarcadorSei,
    ProcessoSei, DocumentoExternoSei, DadosDocumentoExterno,
)
from integra.sei.core.enums import StatusLogin
from integra.siape import (
    IniciaModSiape, Terminal3270Connection, TrocaHabilitacao,
    FichaAnualPensionista, GerenciadorSessaoImpressao,
)
from integra.comum import PrimeiroPlanoNavegador

# ===== CONSTANTES =====

PLANILHA_DADOS = Path(__file__).parent / "planilhas" / "dados.xlsx"
DOCS_DIR = Path(__file__).parent / "docs_pss"

# Chrome Debug para autenticação SIAPE (certificado digital manual)
CHROME_DEBUG_PORT = 9222
SIAPE_LOGIN_URL = "https://www1.siapenet.gov.br/orgao/Login.do?method=inicio"

# Documentos fixos a serem anexados em cada processo (nesta ordem)
DOCUMENTOS_ANEXO = [
    DadosDocumentoExterno(
        tipo_serie="Anexo",
        nome_arvore=" - Notícia RFB",
        arquivo_path=str(DOCS_DIR / "Calculo_de_pensoes.pdf"),
    ),
    DadosDocumentoExterno(
        tipo_serie="Anexo",
        nome_arvore=" - Tabela SELIC",
        arquivo_path=str(DOCS_DIR / "SELIC_CALCULO_SIMPLES.pdf"),
    ),
    DadosDocumentoExterno(
        tipo_serie="Anexo",
        nome_arvore=" - Tabela PSS",
        arquivo_path=str(DOCS_DIR / "TABELA_DEDUCAO_PSS.pdf"),
    ),
    DadosDocumentoExterno(
        tipo_serie="Anexo",
        nome_arvore=" - Nota Técnica",
        arquivo_path=str(DOCS_DIR / "SEI_52928912_Nota_Tecnica_34858.pdf"),
    ),
]


# ===== FUNÇÕES AUXILIARES =====

def _verificar_autenticacao_siape(porta=CHROME_DEBUG_PORT):
    """
    Verifica se o usuário já autenticou no SIAPE via Chrome Debug.
    Consulta o endpoint HTTP do protocolo de debug para verificar as URLs das abas.

    Returns:
        bool: True se a URL indica que o login foi concluído
    """
    try:
        url = f"http://127.0.0.1:{porta}/json"
        resp = urllib.request.urlopen(url, timeout=3)
        tabs = json.loads(resp.read())
        for tab in tabs:
            tab_url = tab.get("url", "")
            # Após login, a URL muda de Login.do para a área logada do SIAPE
            if "siapenet" in tab_url and "Login.do" not in tab_url:
                return True
        return False
    except Exception:
        return False


# ===== AUTOMAÇÃO DO SEI =====

def iniciar_sei(usuario: str, senha: str, log_callback=None):
    """
    Inicia o SEI: abre o Chrome, faz login, fecha a tela de aviso e seleciona a unidade.
    Em seguida, executa o fluxo completo de teste no processo especificado.

    Para autenticação SIAPE, usa Chrome Debug: abre um Chrome limpo na página de login
    e o usuário faz a autenticação manual com certificado digital. A automação detecta
    automaticamente quando a autenticação foi concluída e continua o fluxo.

    Args:
        usuario: Usuário do SEI
        senha: Senha do SEI
        log_callback: Função de log (mensagem, nivel)

    Returns:
        driver: Instância do WebDriver com SEI logado, ou None se falhar
    """
    log = log_callback or (lambda msg, nivel="info": print(msg))

    # 1. Configurar e abrir o Chrome
    log("Iniciando o Chrome...", "info")
    setup = SeleniumSetup()
    driver = setup.setup_driver()

    # 2. Fazer login no SEI
    login = LoginSei(driver, usuario, senha, callback_log=log)
    status = login.logar_sei()

    if status != StatusLogin.SUCESSO:
        log(f"Falha no login: {status}", "error")
        driver.quit()
        return None

    # 3. Fechar tela de aviso
    tela_aviso = TelaAviso(driver, callback_log=log)
    tela_aviso.fechar_tela_aviso_sei()

    # 4. Selecionar unidade DEXTRA
    selecao = SelecaoUnidadeDireta(driver, "MGI-SGP-DECIPEX-CGPAG-DEXTRA", callback_log=log)
    selecao.selecionar_unidade_sei()

    log("SEI iniciado com sucesso na unidade DEXTRA!", "success")

    # ===== DADOS DE TESTE =====
    numero_teste = "19975.004848/2026-66"

    # Ler dados da pensionista da planilha (primeira linha para teste)
    log("Lendo dados da pensionista na planilha...", "info")
    df = pd.read_excel(PLANILHA_DADOS, dtype={"pen_SEI": str, "SIAPE_PENSIONISTA": str, "SIAPE_INST": str})
    row = df.iloc[0]
    siape_pens = str(row["SIAPE_PENSIONISTA"]).strip()
    siape_inst = str(row["SIAPE_INST"]).strip()
    nome_pens = row["NOME_PENSIONISTA"]
    log(f"Pensionista: {nome_pens} | SIAPE Pens: {siape_pens} | SIAPE Inst: {siape_inst}", "info")

    # 5. Acessar processo existente para teste
    log(f"Acessando processo {numero_teste}...", "info")
    acesso = ProcessoSei(driver, numero_teste, callback_log=log)
    acesso.acessar_processo_especifico()

    # 6. Inserir marcador no processo (JÁ TESTADO - pulando)
    # marcador = TrocaMarcadorSei(
    #     navegador=driver,
    #     mensagem="Processo criado via automação PSS",
    #     inserir="INTEGRA - Processo Criado",
    #     retornar_controle_processos=False,
    #     callback_log=log,
    # )
    #
    # if marcador.trocar_marcador():
    #     log("Marcador 'INTEGRA - Processo Criado' inserido com sucesso!", "success")
    # else:
    #     log("Falha ao inserir marcador no processo", "error")
    log("Passo 6 (Marcador) pulado - já testado.", "info")

    # 7. Extrair ficha financeira do SIAPE e inserir no SEI
    log("=" * 60, "info")
    log("INICIANDO EXTRAÇÃO DE FICHA FINANCEIRA", "info")
    log("=" * 60, "info")

    # 7a. Abrir Chrome Debug para autenticação SIAPE (certificado digital manual)
    download_folder = setup.get_download_folder()
    log("Abrindo Chrome para autenticação SIAPE...", "info")
    chrome_proc = ChromeDebug.abrir_chrome(
        porta=CHROME_DEBUG_PORT,
        url_inicial=SIAPE_LOGIN_URL,
        maximizado=True,
        callback_log=log,
    )

    log("=" * 60, "warning")
    log("AUTENTICAÇÃO MANUAL NECESSÁRIA!", "warning")
    log("No Chrome que acabou de abrir:", "warning")
    log("  1. Clique no botão de CERTIFICADO DIGITAL", "warning")
    log("  2. Selecione seu certificado e digite a senha/PIN", "warning")
    log("  3. Aguarde o carregamento da página do SIAPE", "warning")
    log("A automação continuará automaticamente após o login...", "warning")
    log("=" * 60, "warning")

    # 7b. Aguardar autenticação (polling automático)
    max_wait = 180  # 3 minutos
    waited = 0
    autenticado = False
    while waited < max_wait:
        time.sleep(5)
        waited += 5
        if _verificar_autenticacao_siape():
            log("Autenticação SIAPE detectada!", "success")
            autenticado = True
            break
        if waited % 30 == 0:
            log(f"Aguardando autenticação... ({waited}s/{max_wait}s)", "info")

    if not autenticado:
        log("Timeout aguardando autenticação SIAPE (3 min)", "error")
        return driver

    # 7c. Conectar Selenium ao Chrome Debug autenticado
    log("Conectando ao Chrome SIAPE autenticado...", "info")
    driver_siape = ChromeDebug.conectar(porta=CHROME_DEBUG_PORT, callback_log=log)

    # Configurar pasta de download no Chrome Debug
    driver_siape.execute_cdp_cmd("Browser.setDownloadBehavior", {
        "behavior": "allow",
        "downloadPath": download_folder,
    })

    # 7d. Usar IniciaModSiape para as etapas pós-autenticação
    inicia_siape = IniciaModSiape(
        navegador=driver_siape,
        download_folder=download_folder,
        siape="",  # Não necessário — autenticação já feita via Chrome Debug
        abrir_nova_aba=False,
        callback_log=log,
    )

    # Clicar menu SIAPE para baixar módulo desktop
    log("Clicando no menu SIAPE para baixar módulo...", "info")
    inicia_siape._clicar_menu_siape()
    inicia_siape._tratar_aviso_certificado_chrome()

    # Baixar e executar módulo desktop
    if not inicia_siape.baixar_e_executar_modulo():
        log("Falha ao baixar/executar módulo SIAPE", "error")
        return driver

    time.sleep(15)

    # Conectar à aplicação desktop (Painel de Controle + Terminal 3270)
    if not inicia_siape.conectar_aplicacao_desktop():
        log("Falha ao conectar aplicação desktop SIAPE", "error")
        return driver

    # 7e. Conectar ao Terminal 3270
    log("Conectando ao Terminal 3270...", "info")
    terminal_conn = Terminal3270Connection()

    # 7f. Criar instância do extrator de fichas
    ficha_pensionista = FichaAnualPensionista(terminal_conn, callback_log=log)

    # 7g. Trocar habilitação para órgão/UPAG corretos
    log("Trocando habilitação SIAPE...", "info")
    troca_hab = TrocaHabilitacao(terminal_conn, "40806", "000000001", callback_log=log)
    troca_hab.trocar_habilitacao()

    # 7h. Acessar comando de ficha financeira
    ficha_pensionista.extrair_ficha_anual()

    # 7i. Selecionar pensionista pelo SIAPE
    ficha_pensionista.selecionar_pensionista(siape_pens, siape_inst)

    # 7j. Imprimir fichas de 2020 a 2024 (retorna caminho do PDF unificado)
    caminho_ficha = ficha_pensionista.imprimir_fichas("2020", "2024")

    if caminho_ficha:
        log(f"Ficha financeira extraída: {caminho_ficha}", "success")

        # 7k. Voltar para a aba do SEI
        PrimeiroPlanoNavegador(driver, "SEI").enviar_primeiro_plano()
        driver.switch_to.window(driver.window_handles[0])

        # 7l. Inserir ficha financeira como documento externo no SEI
        log("Inserindo ficha financeira no SEI...", "info")
        dados_ficha = DadosDocumentoExterno(
            tipo_serie="Ficha",
            nome_arvore=" - Financeira",
            arquivo_path=caminho_ficha,
        )
        doc_ficha = DocumentoExternoSei(driver, dados_ficha, callback_log=log)

        if doc_ficha.incluir_documento_externo():
            log("✅ Ficha financeira inserida com sucesso!", "success")
        else:
            log("❌ Falha ao inserir ficha financeira", "error")
    else:
        log("❌ Falha ao extrair ficha financeira do SIAPE", "error")

    # 8. Inserir os 4 documentos externos como anexo (JÁ TESTADO - pulando)
    # for i, dados_doc in enumerate(DOCUMENTOS_ANEXO, start=1):
    #     log(f"Anexo {i}/{len(DOCUMENTOS_ANEXO)}: {dados_doc.nome_arvore.strip()}", "info")
    #     doc = DocumentoExternoSei(driver, dados_doc, callback_log=log)
    #     if doc.incluir_documento_externo():
    #         log(f"✅ Anexo {i} inserido com sucesso!", "success")
    #     else:
    #         log(f"❌ Falha ao inserir anexo {i}: {dados_doc.nome_arvore.strip()}", "error")
    log("Passo 8 (4 Anexos) pulado - já testado.", "info")

    log("Teste concluído!", "success")
    return driver


# ===== CALLBACK DA INTERFACE =====

def executar_automacao(janela: JanelaIntegraBase, valores: dict):
    """
    Callback chamado ao clicar 'Iniciar Automação'.
    Roda em thread separada para não travar a interface.
    """
    def _executar():
        try:
            driver = iniciar_sei(
                usuario=valores["login"],
                senha=valores["senha"],
                log_callback=janela.adicionar_log,
            )

            if not driver:
                janela.adicionar_log("Automação encerrada com erro.", "error")

        except Exception as e:
            janela.adicionar_log(f"Erro inesperado: {e}", "error")

        finally:
            janela.automacao_finalizada()

    thread = threading.Thread(target=_executar, daemon=True)
    thread.start()


# ===== CONFIGURAÇÃO DA INTERFACE =====

def criar_configuracao() -> ConfiguracaoInterface:
    """Cria a configuração da interface para o PSS."""
    config = ConfiguracaoInterface()

    config.titulo = "PSS"
    config.titulo_completo = "Automação SEI"
    config.versao = "1.0.0"
    config.desenvolvedor = "Mr. M"

    config.campos_formulario = {
        "login": {
            "label": "Usuário:",
            "placeholder": "Digite seu usuário do SEI",
            "obrigatorio": True,
        },
        "senha": {
            "label": "Senha:",
            "placeholder": "Digite sua senha",
            "obrigatorio": True,
        },
    }

    config.automacoes = ["Iniciar SEI"]
    config.nome_app_credenciais = "PSS"
    config.usar_gerenciamento_credenciais = True
    config.sempre_no_topo = True

    config.callback_iniciar = executar_automacao

    return config


# ===== MAIN =====

if __name__ == "__main__":
    app = QApplication(sys.argv)

    config = criar_configuracao()
    janela = JanelaIntegraBase(config)
    janela.show()

    sys.exit(app.exec_())
