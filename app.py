import os
import fnmatch
import sys
from win32com.client import Dispatch
from win32com.client.dynamic import ERRORS_BAD_CONTEXT
import winerror


# python app.py 'C:\Users\gu060589\Desktop\495\Versão H3C' '*.pdf' 'C:\Users\gu060589\Desktop\495\Versão H3C' '.docx'


# Tentar importar scandir; se encontrado, usá-lo, pois é muito mais rápido do que o os.walk padrão
try:
    from scandir import walk
except ImportError:
    from os import walk

ROOT_INPUT_PATH = None  # Caminho raiz dos arquivos de entrada
ROOT_OUTPUT_PATH = None  # Caminho raiz dos arquivos de saída
INPUT_FILE_EXTENSION = "*.pdf"  # Extensão dos arquivos de entrada
OUTPUT_FILE_EXTENSION = ".docx"  # Extensão dos arquivos de saída


def acrobat_extract_text(f_path, f_path_out, f_basename, f_ext):
    """
    Função para extrair texto de um PDF e salvá-lo como DOCX usando Adobe Acrobat.

    :param f_path: Caminho do arquivo PDF de entrada
    :param f_path_out: Caminho de saída onde o DOCX será salvo
    :param f_basename: Nome base do arquivo (sem extensão)
    :param f_ext: Extensão do arquivo de saída (DOCX)
    """
    avDoc = Dispatch("AcroExch.AVDoc")  # Conectar ao Adobe Acrobat

    # Abrir o arquivo de entrada (como PDF)
    ret = avDoc.Open(f_path, f_path)
    assert (ret)  # Verificar se o arquivo foi aberto com sucesso

    pdDoc = avDoc.GetPDDoc()
    dst = os.path.join(f_path_out, ''.join((f_basename, f_ext)))

    # Obter o objeto JavaScript do documento PDF
    jsObject = pdDoc.GetJSObject()

    # Salvar o documento PDF como DOCX
    jsObject.SaveAs(dst, "com.adobe.acrobat.docx")

    pdDoc.Close()
    avDoc.Close(True)  # Fechar o Acrobat após o processamento
    del pdDoc


if __name__ == "__main__":
    assert (5 == len(sys.argv)), sys.argv  # Garantir que o número de argumentos seja 5

    ROOT_INPUT_PATH = sys.argv[1]  # Caminho raiz dos arquivos de entrada
    INPUT_FILE_EXTENSION = sys.argv[2]  # Extensão dos arquivos de entrada
    ROOT_OUTPUT_PATH = sys.argv[3]  # Caminho raiz dos arquivos de saída
    OUTPUT_FILE_EXTENSION = sys.argv[4]  # Extensão dos arquivos de saída

    # Gerar uma lista de arquivos que correspondem ao padrão
    matching_files = (
        (os.path.join(_root, filename), os.path.splitext(filename)[0])
        for _root, _dirs, _files in walk(ROOT_INPUT_PATH)
        for filename in fnmatch.filter(_files, INPUT_FILE_EXTENSION)
    )

    # Corrigir ERRORS_BAD_CONTEXT conforme necessário
    global ERRORS_BAD_CONTEXT
    ERRORS_BAD_CONTEXT.append(winerror.E_NOTIMPL)

    for filename_with_path, filename_without_extension in matching_files:
        print(f"Processando '{filename_without_extension}'")
        try:
            acrobat_extract_text(filename_with_path, ROOT_OUTPUT_PATH, filename_without_extension,
                                 OUTPUT_FILE_EXTENSION)
            print(f"Processado com sucesso '{filename_without_extension}'")
        except Exception as e:
            print(f"Falha ao processar '{filename_without_extension}': {e}")
