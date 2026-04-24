"""
Controle de estado dos processos SEI processados.

Armazena em JSON qual documento já foi inserido em cada processo,
permitindo:
- Retomar a automação após falha sem reinserir documentos
- Evitar duplicação de anexos em execuções repetidas
- Manter trilha de auditoria (data/hora, erros)
"""

import json
import threading
from datetime import datetime
from pathlib import Path
from typing import Optional


class EstadoProcessos:
    """Gerencia o estado persistente dos processos no arquivo JSON."""

    def __init__(self, arquivo: Path):
        self.arquivo = Path(arquivo)
        self.arquivo.parent.mkdir(parents=True, exist_ok=True)
        self._lock = threading.Lock()
        self._estado = self._carregar()

    # ===== I/O =====

    def _carregar(self) -> dict:
        if self.arquivo.exists():
            try:
                with open(self.arquivo, "r", encoding="utf-8") as f:
                    return json.load(f)
            except Exception:
                return {}
        return {}

    def _salvar(self):
        """Grava o JSON de forma atômica (tmp + replace)."""
        tmp = self.arquivo.with_suffix(".tmp")
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump(self._estado, f, indent=2, ensure_ascii=False)
        tmp.replace(self.arquivo)

    def _registro(self, numero: str) -> dict:
        """Obtém (ou cria) o registro de um processo."""
        if numero not in self._estado:
            self._estado[numero] = {
                "anexos_inseridos": [],
                "ficha_inserida": False,
                "marcador_trocado": False,
                "ultima_execucao": None,
                "erros": [],
            }
        return self._estado[numero]

    def _tocar(self, reg: dict):
        reg["ultima_execucao"] = datetime.now().isoformat(timespec="seconds")

    # ===== ANEXOS =====

    def anexo_ja_inserido(self, numero: str, nome_anexo: str) -> bool:
        with self._lock:
            return nome_anexo in self._registro(numero)["anexos_inseridos"]

    def marcar_anexo(self, numero: str, nome_anexo: str):
        with self._lock:
            reg = self._registro(numero)
            if nome_anexo not in reg["anexos_inseridos"]:
                reg["anexos_inseridos"].append(nome_anexo)
            self._tocar(reg)
            self._salvar()

    # ===== FICHA FINANCEIRA =====

    def ficha_ja_inserida(self, numero: str) -> bool:
        with self._lock:
            return self._registro(numero)["ficha_inserida"]

    def marcar_ficha(self, numero: str):
        with self._lock:
            reg = self._registro(numero)
            reg["ficha_inserida"] = True
            self._tocar(reg)
            self._salvar()

    # ===== MARCADOR =====

    def marcador_ja_trocado(self, numero: str) -> bool:
        with self._lock:
            return self._registro(numero)["marcador_trocado"]

    def marcar_marcador(self, numero: str):
        with self._lock:
            reg = self._registro(numero)
            reg["marcador_trocado"] = True
            self._tocar(reg)
            self._salvar()

    # ===== ERROS =====

    def registrar_erro(self, numero: str, mensagem: str):
        with self._lock:
            reg = self._registro(numero)
            reg["erros"].append(
                {
                    "quando": datetime.now().isoformat(timespec="seconds"),
                    "mensagem": mensagem,
                }
            )
            self._tocar(reg)
            self._salvar()

    # ===== CONSULTA =====

    def status(self, numero: str) -> dict:
        """Retorna cópia do registro do processo (ou registro vazio)."""
        with self._lock:
            return self._registro(numero).copy()

    def processo_completo(self, numero: str, total_anexos: int) -> bool:
        """Retorna True se o processo já teve todos os passos concluídos."""
        with self._lock:
            reg = self._registro(numero)
            return (
                len(reg["anexos_inseridos"]) >= total_anexos
                and reg["ficha_inserida"]
                and reg["marcador_trocado"]
            )
