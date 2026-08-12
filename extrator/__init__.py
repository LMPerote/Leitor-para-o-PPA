"""Extrator de fichas de Indicadores de Compromisso (.docx) para Excel."""

from .documento import Documento, ler_documento
from .parser import Extrator, Resultado
from .pipeline import Estatisticas, listar_documentos, processar_lote
from .planilha import montar_dataframe, salvar_csv, salvar_excel

__version__ = "1.0.0"

__all__ = [
    "Documento",
    "Estatisticas",
    "Extrator",
    "Resultado",
    "ler_documento",
    "listar_documentos",
    "montar_dataframe",
    "processar_lote",
    "salvar_csv",
    "salvar_excel",
    "__version__",
]
