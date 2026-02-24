"""
PSS - Automação do Sistema Eletrônico de Informações (SEI)

Fluxo:
1. Configura e abre o Chrome com anti-detecção
2. Realiza login no SEI
3. Fecha tela de aviso (se houver)
"""

from integra.browser import SeleniumSetup
from integra.sei import LoginSei, TelaAviso
from integra.sei.core.enums import StatusLogin


def log_callback(mensagem: str, nivel: str = "info") -> None:
    """Callback simples de log para o console."""
    prefixos = {
        "info": "ℹ️",
        "success": "✅",
        "warning": "⚠️",
        "error": "❌",
        "debug": "🔍",
    }
    print(f"{prefixos.get(nivel, '📝')} [{nivel.upper()}] {mensagem}")


def iniciar_sei(usuario: str, senha: str):
    """
    Inicia o SEI: abre o Chrome, faz login e fecha a tela de aviso.

    Args:
        usuario: Usuário do SEI
        senha: Senha do SEI

    Returns:
        driver: Instância do WebDriver com SEI logado
    """
    # 1. Configurar e abrir o Chrome
    log_callback("Iniciando o Chrome...", "info")
    setup = SeleniumSetup()
    driver = setup.setup_driver()

    # 2. Fazer login no SEI
    login = LoginSei(driver, usuario, senha, callback_log=log_callback)
    status = login.logar_sei()

    if status != StatusLogin.SUCESSO:
        log_callback(f"Falha no login: {status}", "error")
        driver.quit()
        return None

    # 3. Fechar tela de aviso
    tela_aviso = TelaAviso(driver, callback_log=log_callback)
    tela_aviso.fechar_tela_aviso_sei()

    log_callback("SEI iniciado com sucesso!", "success")
    return driver


if __name__ == "__main__":
    usuario = input("Usuário SEI: ")
    senha = input("Senha SEI: ")

    driver = iniciar_sei(usuario, senha)

    if driver:
        input("Pressione Enter para encerrar...")
        driver.quit()
