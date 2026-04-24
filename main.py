"""
PSS - Automação do Sistema Eletrônico de Informações (SEI)

Interface gráfica com:
- Campos de login e senha protegidos
- Credenciais criptografadas (AES/Fernet)
- Log de ações em tempo real
"""

import sys
import time
import threading
import queue as queue_module
from pathlib import Path

import pandas as pd
from PyQt5.QtWidgets import QApplication, QInputDialog
from PyQt5.QtCore import pyqtSignal, QObject

from integra.interface import JanelaIntegraBase, ConfiguracaoInterface
from integra.browser import SeleniumSetup
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException

from integra.sei import (
    LoginSei, TelaAviso, SelecaoUnidadeDireta,
    IniciaProcessos, NumeroProcesso, TrocaMarcadorSei,
    MarcadorSei, VisualizacaoDetalhada,
    ProcessoSei, DocumentoExternoSei, DadosDocumentoExterno,
)
from integra.sei.captor_documentos import CaptorDocumentos
from integra.sei.seletor_marcadores import SeletorMarcadoresSEI
from integra.sei.core.enums import StatusLogin
from integra.siape import (
    IniciaModSiape, Terminal3270Connection, TrocaHabilitacao,
    FichaAnualPensionista, GerenciadorSessaoImpressao,
)
from integra.comum import PrimeiroPlanoNavegador

from estado_processos import EstadoProcessos

# ===== CONSTANTES =====

PLANILHA_DADOS = Path(__file__).parent / "planilhas" / "dados.xlsx"
DOCS_DIR = Path(__file__).parent / "docs_pss"
ESTADO_JSON = Path(__file__).parent / "estado" / "estado_processos.json"

# Marcador SEI de destino (após inserção dos documentos)
MARCADOR_DESTINO = "INTEGRA - FICHA + DOCUMENTOS"

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


# ===== DIÁLOGO DE SELEÇÃO DE MARCADOR (THREAD-SAFE) =====

class SeletorMarcadorDialog(QObject):
    """Exibe diálogo Qt para seleção de marcador, seguro para chamar de thread secundária."""
    _solicitar = pyqtSignal(list)

    def __init__(self, parent=None):
        super().__init__(parent)
        self._fila = queue_module.Queue()
        self._solicitar.connect(self._mostrar_dialogo)

    def _mostrar_dialogo(self, items_texto):
        item, ok = QInputDialog.getItem(
            self.parent(), "Selecionar Marcador",
            "Escolha o marcador para trabalhar:",
            items_texto, 0, False,
        )
        if ok and item:
            self._fila.put(items_texto.index(item))
        else:
            self._fila.put(-1)

    def escolher(self, marcadores):
        """Mostra diálogo e retorna o MarcadorSEI escolhido (ou None se cancelado)."""
        items = [f"{m.nome} ({m.quantidade} processos)" for m in marcadores]
        self._solicitar.emit(items)
        idx = self._fila.get()
        if 0 <= idx < len(marcadores):
            return marcadores[idx]
        return None


# ===== AUTOMAÇÃO DO SEI =====

def fluxo_completo_siape(usuario: str, senha: str, token: str = "", log_callback=None, limite: int = 0, seletor_dialog=None):
    """
    Fluxo completo PSS: login SEI, seleção de marcador, e para cada processo
    filtrado insere os 4 anexos da instrução + ficha financeira (extraída do SIAPE),
    trocando o marcador ao final para 'INTEGRA - FICHA + DOCUMENTOS'.

    Usa controle de estado via JSON (estado/estado_processos.json) + sincronização
    com a árvore do SEI (CaptorDocumentos) para suportar retomada e processos legados.

    Args:
        usuario: Usuário do SEI
        senha: Senha do SEI
        token: PIN/senha do certificado digital (para autenticação SIAPE)
        log_callback: Função de log (mensagem, nivel)
        limite: Número máximo de processos (0 = sem limite)
        seletor_dialog: SeletorMarcadorDialog para escolha via GUI
    """
    log = log_callback or (lambda msg, nivel="info": print(msg))

    # 0. Inicializar controle de estado (JSON) para evitar reinserções
    estado = EstadoProcessos(ESTADO_JSON)
    log(f"Estado persistente: {ESTADO_JSON}", "info")

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

    # 5. Listar marcadores disponíveis e permitir escolha do usuário
    log("Listando marcadores disponíveis na unidade...", "info")
    seletor = SeletorMarcadoresSEI(driver=driver, callback_log=log)
    seletor.conectar()
    marcadores = seletor.listar_marcadores()

    if not marcadores:
        log("Nenhum marcador encontrado na unidade.", "error")
        driver.quit()
        return None

    for m in marcadores:
        log(f"  [{m.id}] {m.nome} — {m.quantidade} processos", "info")

    if seletor_dialog:
        log("Aguardando seleção do marcador...", "info")
        marcador_escolhido = seletor_dialog.escolher(marcadores)
    else:
        marcador_escolhido = marcadores[0] if marcadores else None

    if not marcador_escolhido:
        log("Nenhum marcador selecionado. Operação cancelada.", "warning")
        driver.quit()
        return None

    log(f"Marcador selecionado: {marcador_escolhido.nome} ({marcador_escolhido.quantidade} processos)", "success")

    # 6. Filtrar processos pelo marcador selecionado e ativar visualização detalhada
    log(f"Filtrando processos com marcador '{marcador_escolhido.nome}'...", "info")
    seletor.filtrar_por_id(marcador_escolhido.id)

    visualizacao = VisualizacaoDetalhada(driver, callback_log=log)
    visualizacao.visualizar_detalhado()

    time.sleep(2)

    # 6a. Carregar planilha de pensionistas (lookup por pen_SEI)
    log("Carregando planilha de pensionistas...", "info")
    try:
        df_pens = pd.read_excel(
            PLANILHA_DADOS,
            dtype={"pen_SEI": str, "SIAPE_PENSIONISTA": str, "SIAPE_INST": str},
        )
        df_pens = df_pens.set_index("pen_SEI")
        log(f"Planilha carregada: {len(df_pens)} pensionistas", "info")
    except Exception as e:
        log(f"Erro ao carregar planilha: {e}", "error")
        df_pens = None

    # 6b. Inicializar SIAPE UMA ÚNICA VEZ (para extração de fichas)
    ficha_extrator = None
    if token:
        log("=" * 60, "info")
        log("Inicializando SIAPE para extração de fichas...", "info")
        log("=" * 60, "info")

        try:
            download_folder = setup.get_download_folder
            inicia_siape = IniciaModSiape(
                navegador=driver,
                download_folder=download_folder,
                siape=token,
                abrir_nova_aba=True,
                callback_log=log,
            )

            log("Autenticando no SIAPE (certificado digital)...", "info")
            log("Confirme no celular/token quando solicitado.", "warning")

            if inicia_siape.executar_siape():
                terminal_conn = Terminal3270Connection(callback_log=log)
                if terminal_conn.conectar_terminal():
                    troca_hab = TrocaHabilitacao(terminal_conn, "40806", "000000001", callback_log=log)
                    if troca_hab.trocar_habilitacao():
                        ficha_extrator = FichaAnualPensionista(terminal_conn, callback_log=log)
                        log("SIAPE pronto para extração de fichas!", "success")
                    else:
                        log("Falha ao trocar habilitação — seguindo sem ficha", "warning")
                else:
                    log("Falha no Terminal 3270 — seguindo sem ficha", "warning")
            else:
                log("Falha na autenticação SIAPE — seguindo sem ficha", "warning")
        except Exception as e:
            log(f"Erro ao inicializar SIAPE: {e} — seguindo sem ficha", "warning")

        # Voltar para a aba do SEI
        try:
            PrimeiroPlanoNavegador(driver, "SEI").enviar_primeiro_plano()
            driver.switch_to.window(driver.window_handles[0])
        except Exception:
            pass
    else:
        log("Token não fornecido — extração de ficha desabilitada", "warning")

    # 7. Iterar pelos processos filtrados
    processados = 0
    erros = 0
    total_texto = f"/{limite}" if limite > 0 else ""

    while True:
        # Verificar limite
        if limite > 0 and processados >= limite:
            log(f"Limite de {limite} processos atingido.", "info")
            break

        # Clicar no próximo processo da lista
        try:
            elemento = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.CLASS_NAME, "processoVisualizado"))
            )
            elemento.click()
            time.sleep(1)
        except TimeoutException:
            log("Não há mais processos com este marcador.", "info")
            break

        # Obter número do processo
        processo_info = NumeroProcesso(driver, callback_log=log)
        numero = processo_info.obter_numero_processo()
        processados += 1
        log(f"[{processados}{total_texto}] Processo: {numero}", "info")

        # 7.0 Sincronizar JSON com a árvore de documentos do SEI
        # (captura o que já existe na árvore — útil para processos legados
        #  trabalhados antes do controle por JSON)
        try:
            captor = CaptorDocumentos(driver, callback_log=log)
            captor.capturar_documentos()
            termos_busca = [d.nome_arvore.strip() for d in DOCUMENTOS_ANEXO] + ["- Financeira"]
            relatorio = captor.gerar_relatorio_busca(termos_busca)

            sincronizados = 0
            for dados_doc in DOCUMENTOS_ANEXO:
                nome = dados_doc.nome_arvore.strip()
                if relatorio.get(nome) and not estado.anexo_ja_inserido(numero, nome):
                    estado.marcar_anexo(numero, nome)
                    sincronizados += 1

            if relatorio.get("- Financeira") and not estado.ficha_ja_inserida(numero):
                estado.marcar_ficha(numero)
                sincronizados += 1

            if sincronizados:
                log(f"  Árvore SEI sincronizada: {sincronizados} documento(s) já presentes no processo", "info")
        except Exception as e:
            log(f"  Aviso: falha ao capturar árvore de documentos ({e}) — seguindo sem sync", "warning")

        # 7a. Inserir os 4 documentos anexos (pulando os já inseridos)
        sucesso_docs = True
        for i, dados_doc in enumerate(DOCUMENTOS_ANEXO, start=1):
            nome_anexo = dados_doc.nome_arvore.strip()

            if estado.anexo_ja_inserido(numero, nome_anexo):
                log(f"  Anexo {i}/4: {nome_anexo} — JÁ INSERIDO (pulando)", "info")
                continue

            log(f"  Anexo {i}/4: {nome_anexo}", "info")
            doc = DocumentoExternoSei(driver, dados_doc, callback_log=log)
            if not doc.incluir_documento_externo():
                log(f"  Falha no anexo {i}: {nome_anexo}", "error")
                estado.registrar_erro(numero, f"Falha no anexo: {nome_anexo}")
                sucesso_docs = False
                break
            estado.marcar_anexo(numero, nome_anexo)

        if not sucesso_docs:
            erros += 1
            log(f"  Processo {numero} com erro nos anexos — pulando troca de marcador", "error")
            TrocaMarcadorSei(
                navegador=driver,
                mensagem="",
                retornar_controle_processos=True,
                callback_log=log,
            ).trocar_marcador()
            continue

        # 7b. Extrair ficha financeira do SIAPE e inserir no SEI (se ainda não)
        if estado.ficha_ja_inserida(numero):
            log("  Ficha financeira — JÁ INSERIDA (pulando)", "info")
        elif ficha_extrator and df_pens is not None and numero in df_pens.index:
            row = df_pens.loc[numero]
            if isinstance(row, pd.DataFrame):
                row = row.iloc[0]
            siape_pens = str(row["SIAPE_PENSIONISTA"]).strip()
            siape_inst = str(row["SIAPE_INST"]).strip()
            nome_pens = row.get("NOME_PENSIONISTA", "?") if hasattr(row, "get") else "?"
            log(f"  Extraindo ficha — {nome_pens} (Pens: {siape_pens} | Inst: {siape_inst})", "info")

            caminho_ficha = None
            try:
                if ficha_extrator.extrair_ficha_anual() and \
                        ficha_extrator.selecionar_pensionista(siape_pens, siape_inst):
                    caminho_ficha = ficha_extrator.imprimir_fichas("2020", "2024")
            except Exception as e:
                log(f"  Erro ao extrair ficha: {e}", "error")
                estado.registrar_erro(numero, f"Erro extração ficha: {e}")

            # Voltar para a aba do SEI
            try:
                PrimeiroPlanoNavegador(driver, "SEI").enviar_primeiro_plano()
                driver.switch_to.window(driver.window_handles[0])
            except Exception:
                pass

            if caminho_ficha:
                log(f"  Ficha extraída: {caminho_ficha}", "success")
                dados_ficha = DadosDocumentoExterno(
                    tipo_serie="Ficha",
                    nome_arvore=" - Financeira",
                    arquivo_path=caminho_ficha,
                )
                doc_ficha = DocumentoExternoSei(driver, dados_ficha, callback_log=log)
                if not doc_ficha.incluir_documento_externo():
                    log("  Falha ao inserir ficha no SEI", "error")
                    estado.registrar_erro(numero, "Falha ao inserir ficha no SEI")
                    erros += 1
                    TrocaMarcadorSei(
                        navegador=driver,
                        mensagem="",
                        retornar_controle_processos=True,
                        callback_log=log,
                    ).trocar_marcador()
                    continue
                estado.marcar_ficha(numero)
                log("  Ficha financeira inserida no SEI!", "success")
            else:
                log("  Falha ao extrair ficha financeira — pulando troca de marcador", "error")
                estado.registrar_erro(numero, "Falha ao extrair ficha financeira")
                erros += 1
                TrocaMarcadorSei(
                    navegador=driver,
                    mensagem="",
                    retornar_controle_processos=True,
                    callback_log=log,
                ).trocar_marcador()
                continue
        elif not ficha_extrator:
            log("  SIAPE indisponível — seguindo sem ficha", "warning")
        else:
            log(f"  Processo {numero} não encontrado na planilha — seguindo sem ficha", "warning")

        # 7c. Trocar marcador: remover origem selecionada, inserir destino
        log(f"  Trocando marcador para '{MARCADOR_DESTINO}'...", "info")
        troca = TrocaMarcadorSei(
            navegador=driver,
            mensagem="Documentos inseridos via automação PSS",
            remover=marcador_escolhido.nome,
            inserir=MARCADOR_DESTINO,
            retornar_controle_processos=True,
            callback_log=log,
        )

        if troca.trocar_marcador():
            estado.marcar_marcador(numero)
            log(f"  Processo {numero} concluído!", "success")
        else:
            log(f"  Falha ao trocar marcador no processo {numero}", "error")
            estado.registrar_erro(numero, "Falha ao trocar marcador")
            erros += 1

    # Resumo
    log("=" * 60, "info")
    log(f"CONCLUÍDO: {processados} processos | {processados - erros} OK | {erros} erros", "success")
    log("=" * 60, "info")

    return driver


# ===== CALLBACK DA INTERFACE =====

def executar_automacao(janela: JanelaIntegraBase, valores: dict):
    """
    Callback chamado ao clicar 'Iniciar Automação'.
    Roda em thread separada para não travar a interface.
    Roteia para a automação selecionada no combo.
    """
    # Criar helper de diálogo na thread principal (Qt exige isso)
    seletor_dialog = SeletorMarcadorDialog(janela)

    def _executar():
        driver = None
        try:
            automacao = valores.get("automacao", "")

            if automacao == "Fluxo Completo (SIAPE)":
                driver = fluxo_completo_siape(
                    usuario=valores["login"],
                    senha=valores["senha"],
                    token=valores.get("token", ""),
                    log_callback=janela.adicionar_log,
                    limite=0,
                    seletor_dialog=seletor_dialog,
                )
            else:
                janela.adicionar_log(f"Automação desconhecida: {automacao!r}", "error")

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

    config.titulo = "INTEGRA PSS"
    config.titulo_completo = ""
    config.versao = "1.0.0"
    config.desenvolvedor = "Mr. M"

    # Logo e ícone do projeto INTEGRA (mesmos do integra-exante)
    img_dir = Path(__file__).parent / "recursos" / "img"
    logo_path = img_dir / "logo_exante.png"
    if logo_path.exists():
        config.logo_path = str(logo_path)
    icone_path = img_dir / "icon.ico"
    if icone_path.exists():
        config.icone_path = str(icone_path)

    config.campos_formulario = {
        "login": {
            "label": "Login:",
            "placeholder": "Digite seu usuário do SEI",
            "obrigatorio": True,
        },
        "senha": {
            "label": "Senha:",
            "placeholder": "Digite sua senha",
            "obrigatorio": True,
        },
        "token": {
            "label": "Token:",
            "placeholder": "PIN do certificado digital (só para fluxo SIAPE)",
            "obrigatorio": False,
        },
    }

    # Lista de automações disponíveis no combo.
    # Para adicionar nova automação: incluir aqui + adicionar rota em executar_automacao().
    config.automacoes = ["Fluxo Completo (SIAPE)"]
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
