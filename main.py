"""
PSS - Automação do Sistema Eletrônico de Informações (SEI)

Interface gráfica com:
- Campos de login e senha protegidos
- Credenciais criptografadas (AES/Fernet)
- Log de ações em tempo real
"""

import sys
import threading

from PyQt5.QtWidgets import QApplication

from integra.interface import JanelaIntegraBase, ConfiguracaoInterface
from integra.browser import SeleniumSetup
from integra.sei import LoginSei, TelaAviso, SelecaoUnidadeDireta
from integra.sei.core.enums import StatusLogin


# ===== AUTOMAÇÃO DO SEI =====

def iniciar_sei(usuario: str, senha: str, log_callback=None):
    """
    Inicia o SEI: abre o Chrome, faz login, fecha a tela de aviso e seleciona a unidade.

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
    selecao.selecionar_unidade()

    log("SEI iniciado com sucesso na unidade DEXTRA!", "success")
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
