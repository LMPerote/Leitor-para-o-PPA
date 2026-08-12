"""Extrator de fichas do PPA (Indicadores e Iniciativas) em .docx para Excel."""

from .documento import Documento, ler_documento
from .modelos import MODELOS, Modelo
from .parser import Extrator, Resultado, classificar
from .pipeline import Estatisticas, listar_documentos, processar_lote
from .planilha import montar_dataframe, salvar_csv, salvar_excel

__version__ = "2.0.0"

__all__ = [
    "Documento",
    "Estatisticas",
    "Extrator",
    "MODELOS",
    "Modelo",
    "Resultado",
    "classificar",
    "ler_documento",
    "listar_documentos",
    "montar_dataframe",
    "processar_lote",
    "salvar_csv",
    "salvar_excel",
    "__version__",
]
