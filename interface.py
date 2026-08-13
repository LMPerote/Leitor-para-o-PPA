#!/usr/bin/env python3
"""Janela gráfica do Leitor para o PPA — para usar sem linha de comando.

Abra com dois cliques em "Extrair PPA.bat" (Windows) ou execute:

    python interface.py

A extração roda em uma thread separada para a janela não congelar; a troca de
mensagens com a interface é feita por uma fila, já que só a thread principal
pode tocar nos widgets do Tkinter.
"""

from __future__ import annotations

import json
import logging
import os
import queue
import subprocess
import sys
import threading
import time
from pathlib import Path

try:
    import tkinter as tk
    from tkinter import filedialog, messagebox, ttk
except ImportError:  # pragma: no cover - instalação de Python sem Tk
    print(
        "Este Python não tem a biblioteca gráfica Tkinter instalada.\n"
        "Reinstale o Python marcando a opção 'tcl/tk and IDLE',\n"
        "ou use a versão de linha de comando: python extrair_indicadores.py --help"
    )
    raise SystemExit(1)

sys.path.insert(0, str(Path(__file__).resolve().parent))

from extrator import __version__  # noqa: E402
from extrator.aplicacao import (  # noqa: E402
    SEPARADORES,
    Opcoes,
    PastaSemDocumentos,
    executar,
)
from extrator.relatorio import montar_relatorio  # noqa: E402

TITULO = "Leitor para o PPA"
ARQUIVO_PREFERENCIAS = Path(__file__).resolve().parent / "preferencias.json"
NOME_DA_PASTA_PADRAO = "Planilhas PPA"

# Mensagens trocadas entre a thread de extração e a janela.
PROGRESSO, LOG, FIM, ERRO = "progresso", "log", "fim", "erro"


class _Progresso:
    """Objeto compatível com o tqdm, que reporta o avanço para a janela."""

    def __init__(self, fila: queue.Queue, total: int) -> None:
        self.fila = fila
        self.total = total
        self.contagem = 0

    def update(self, quantidade: int = 1) -> None:
        self.contagem += quantidade
        self.fila.put((PROGRESSO, (self.contagem, self.total)))

    def close(self) -> None:  # pragma: no cover - exigido pela interface do tqdm
        pass


class _RegistradorNaFila(logging.Handler):
    """Envia os avisos do processamento para a área de log da janela."""

    def __init__(self, fila: queue.Queue) -> None:
        super().__init__(level=logging.WARNING)
        self.fila = fila

    def emit(self, registro: logging.LogRecord) -> None:
        self.fila.put((LOG, self.format(registro)))


class Janela:
    """Monta e opera a janela principal."""

    def __init__(self, raiz: tk.Tk) -> None:
        self.raiz = raiz
        self.fila: queue.Queue = queue.Queue()
        self.destinos: dict[str, Path] = {}
        self.executando = False

        raiz.title(f"{TITULO}  v{__version__}")
        raiz.minsize(760, 560)
        raiz.columnconfigure(0, weight=1)
        raiz.rowconfigure(3, weight=1)

        self.entrada = tk.StringVar()
        self.saida = tk.StringVar()
        self.recursivo = tk.BooleanVar(value=True)
        self.gerar_csv = tk.BooleanVar(value=False)
        self.somente_dez = tk.BooleanVar(value=False)
        self.separador = tk.StringVar(value="quebra")
        self.situacao = tk.StringVar(value="Escolha as pastas e clique em Extrair.")

        self._montar_pastas()
        self._montar_opcoes()
        self._montar_acoes()
        self._montar_saida()

        self._carregar_preferencias()
        self.raiz.after(100, self._processar_fila)
        self.raiz.protocol("WM_DELETE_WINDOW", self._ao_fechar)

    # -- construção da interface ------------------------------------------
    def _montar_pastas(self) -> None:
        grupo = ttk.LabelFrame(self.raiz, text="Pastas", padding=10)
        grupo.grid(row=0, column=0, sticky="ew", padx=12, pady=(12, 6))
        grupo.columnconfigure(1, weight=1)

        ttk.Label(grupo, text="Documentos (.docx):").grid(row=0, column=0, sticky="w")
        ttk.Entry(grupo, textvariable=self.entrada).grid(
            row=0, column=1, sticky="ew", padx=8
        )
        ttk.Button(grupo, text="Procurar...", command=self._escolher_entrada).grid(
            row=0, column=2
        )

        ttk.Label(grupo, text="Salvar planilhas em:").grid(
            row=1, column=0, sticky="w", pady=(8, 0)
        )
        ttk.Entry(grupo, textvariable=self.saida).grid(
            row=1, column=1, sticky="ew", padx=8, pady=(8, 0)
        )
        ttk.Button(grupo, text="Procurar...", command=self._escolher_saida).grid(
            row=1, column=2, pady=(8, 0)
        )

        ttk.Label(
            grupo,
            text="As subpastas são incluídas; os dois tipos de ficha podem estar misturados.",
            foreground="#555555",
        ).grid(row=2, column=0, columnspan=3, sticky="w", pady=(8, 0))

    def _montar_opcoes(self) -> None:
        grupo = ttk.LabelFrame(self.raiz, text="Opções", padding=10)
        grupo.grid(row=1, column=0, sticky="ew", padx=12, pady=6)

        ttk.Checkbutton(
            grupo, text="Incluir subpastas", variable=self.recursivo
        ).grid(row=0, column=0, sticky="w", padx=(0, 16))
        ttk.Checkbutton(
            grupo, text="Gerar também CSV", variable=self.gerar_csv
        ).grid(row=0, column=1, sticky="w", padx=(0, 16))
        ttk.Checkbutton(
            grupo,
            text="Testar com os 10 primeiros arquivos",
            variable=self.somente_dez,
        ).grid(row=0, column=2, sticky="w")

        ttk.Label(grupo, text="Separar itens de lista por:").grid(
            row=1, column=0, sticky="w", pady=(10, 0)
        )
        combo = ttk.Combobox(
            grupo,
            textvariable=self.separador,
            values=list(SEPARADORES),
            state="readonly",
            width=16,
        )
        combo.grid(row=1, column=1, sticky="w", pady=(10, 0))

    def _montar_acoes(self) -> None:
        faixa = ttk.Frame(self.raiz, padding=(12, 0))
        faixa.grid(row=2, column=0, sticky="ew")
        faixa.columnconfigure(1, weight=1)

        self.botao_extrair = ttk.Button(faixa, text="Extrair", command=self._iniciar)
        self.botao_extrair.grid(row=0, column=0, sticky="w")

        self.barra = ttk.Progressbar(faixa, mode="determinate")
        self.barra.grid(row=0, column=1, sticky="ew", padx=10)

        self.botao_abrir = ttk.Button(
            faixa, text="Abrir pasta das planilhas", command=self._abrir_saida
        )
        self.botao_abrir.grid(row=0, column=2, sticky="e")
        self.botao_abrir.state(["disabled"])

        ttk.Label(faixa, textvariable=self.situacao).grid(
            row=1, column=0, columnspan=3, sticky="w", pady=(6, 0)
        )

    def _montar_saida(self) -> None:
        grupo = ttk.LabelFrame(self.raiz, text="Resultado", padding=6)
        grupo.grid(row=3, column=0, sticky="nsew", padx=12, pady=(6, 12))
        grupo.columnconfigure(0, weight=1)
        grupo.rowconfigure(0, weight=1)

        self.texto = tk.Text(grupo, height=14, wrap="none")
        self.texto.grid(row=0, column=0, sticky="nsew")
        rolagem = ttk.Scrollbar(grupo, orient="vertical", command=self.texto.yview)
        rolagem.grid(row=0, column=1, sticky="ns")
        self.texto.configure(yscrollcommand=rolagem.set, state="disabled")

    # -- preferências ------------------------------------------------------
    def _carregar_preferencias(self) -> None:
        """Relembra as últimas pastas usadas, se houver."""
        try:
            dados = json.loads(ARQUIVO_PREFERENCIAS.read_text(encoding="utf-8"))
        except (OSError, ValueError):
            dados = {}

        self.entrada.set(dados.get("entrada", ""))
        padrao = Path.home() / "Documents" / NOME_DA_PASTA_PADRAO
        self.saida.set(dados.get("saida") or str(padrao))
        self.separador.set(dados.get("separador", "quebra"))
        self.recursivo.set(dados.get("recursivo", True))
        self.gerar_csv.set(dados.get("gerar_csv", False))

    def _salvar_preferencias(self) -> None:
        dados = {
            "entrada": self.entrada.get(),
            "saida": self.saida.get(),
            "separador": self.separador.get(),
            "recursivo": self.recursivo.get(),
            "gerar_csv": self.gerar_csv.get(),
        }
        try:
            ARQUIVO_PREFERENCIAS.write_text(
                json.dumps(dados, ensure_ascii=False, indent=2), encoding="utf-8"
            )
        except OSError:  # pragma: no cover - pasta somente leitura
            pass

    # -- ações da interface ------------------------------------------------
    def _escolher_entrada(self) -> None:
        escolhida = filedialog.askdirectory(
            title="Selecione a pasta com os documentos .docx",
            initialdir=self.entrada.get() or str(Path.home()),
        )
        if escolhida:
            self.entrada.set(os.path.normpath(escolhida))

    def _escolher_saida(self) -> None:
        escolhida = filedialog.askdirectory(
            title="Selecione onde salvar as planilhas",
            initialdir=self.saida.get() or str(Path.home()),
        )
        if escolhida:
            self.saida.set(os.path.normpath(escolhida))

    def _escrever(self, texto: str) -> None:
        self.texto.configure(state="normal")
        self.texto.insert("end", texto + "\n")
        self.texto.see("end")
        self.texto.configure(state="disabled")

    def _limpar_saida(self) -> None:
        self.texto.configure(state="normal")
        self.texto.delete("1.0", "end")
        self.texto.configure(state="disabled")

    def _iniciar(self) -> None:
        if self.executando:
            return

        entrada = Path(self.entrada.get().strip())
        saida = Path(self.saida.get().strip())
        if not entrada.is_dir():
            messagebox.showerror(
                TITULO, "Escolha uma pasta válida com os documentos .docx."
            )
            return
        if not str(saida):
            messagebox.showerror(TITULO, "Escolha onde salvar as planilhas.")
            return

        self._salvar_preferencias()
        self._limpar_saida()
        self.destinos = {}
        self.executando = True
        self.botao_extrair.state(["disabled"])
        self.botao_abrir.state(["disabled"])
        self.barra.configure(value=0, maximum=100)
        self.situacao.set("Procurando documentos...")

        opcoes = Opcoes(
            entrada=entrada,
            saida=saida,
            recursivo=self.recursivo.get(),
            separador=SEPARADORES[self.separador.get()],
            gerar_csv=self.gerar_csv.get(),
            limite=10 if self.somente_dez.get() else 0,
        )
        threading.Thread(target=self._trabalhar, args=(opcoes,), daemon=True).start()

    def _trabalhar(self, opcoes: Opcoes) -> None:
        """Roda a extração fora da thread da interface."""
        registrador = _RegistradorNaFila(self.fila)
        registrador.setFormatter(logging.Formatter("%(levelname)s: %(message)s"))
        raiz_log = logging.getLogger()
        raiz_log.setLevel(logging.INFO)
        raiz_log.addHandler(registrador)

        inicio = time.perf_counter()
        try:
            execucao = executar(
                opcoes, criar_progresso=lambda total: _Progresso(self.fila, total)
            )
            relatorio = montar_relatorio(
                execucao.estatisticas, execucao.destinos, time.perf_counter() - inicio
            )
            self.fila.put((FIM, (relatorio, execucao.destinos)))
        except PastaSemDocumentos as erro:
            self.fila.put((ERRO, str(erro)))
        except Exception as erro:  # noqa: BLE001 - a janela não pode morrer
            logging.getLogger(__name__).exception("Falha na extração")
            self.fila.put((ERRO, f"{type(erro).__name__}: {erro}"))
        finally:
            raiz_log.removeHandler(registrador)

    def _processar_fila(self) -> None:
        """Consome as mensagens da thread de extração (só a UI mexe na UI)."""
        try:
            while True:
                tipo, dados = self.fila.get_nowait()
                if tipo == PROGRESSO:
                    feitos, total = dados
                    self.barra.configure(maximum=total, value=feitos)
                    self.situacao.set(f"Processando {feitos} de {total} documentos...")
                elif tipo == LOG:
                    self._escrever(dados)
                elif tipo == FIM:
                    relatorio, destinos = dados
                    self.destinos = destinos
                    self._escrever(relatorio)
                    self.situacao.set("Concluído.")
                    self._terminar()
                    self.botao_abrir.state(["!disabled"])
                elif tipo == ERRO:
                    self.situacao.set("Falhou.")
                    self._escrever(f"ERRO: {dados}")
                    self._terminar()
                    messagebox.showerror(TITULO, dados)
        except queue.Empty:
            pass
        self.raiz.after(100, self._processar_fila)

    def _terminar(self) -> None:
        self.executando = False
        self.botao_extrair.state(["!disabled"])

    def _abrir_saida(self) -> None:
        pasta = Path(self.saida.get())
        if not pasta.is_dir():
            return
        if sys.platform.startswith("win"):
            os.startfile(pasta)  # type: ignore[attr-defined]  # noqa: S606
        elif sys.platform == "darwin":
            subprocess.run(["open", str(pasta)], check=False)
        else:
            subprocess.run(["xdg-open", str(pasta)], check=False)

    def _ao_fechar(self) -> None:
        if self.executando and not messagebox.askokcancel(
            TITULO, "A extração está em andamento. Deseja fechar mesmo assim?"
        ):
            return
        self._salvar_preferencias()
        self.raiz.destroy()


def main() -> int:
    raiz = tk.Tk()
    try:  # tema mais próximo do visual do Windows, quando disponível
        ttk.Style().theme_use("vista")
    except tk.TclError:
        pass
    Janela(raiz)
    raiz.mainloop()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
