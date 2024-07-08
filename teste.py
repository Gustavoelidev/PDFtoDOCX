import os
import win32com.client as win32
import pythoncom


# Função para substituir palavras em um documento usando VBA
def replace_words_vba(doc, find_list, replace_list):
    # Inserir o módulo VBA com a macro de substituição
    vba_code = """
    Sub FindAndReplaceMultiItems(findArr As Variant, replaceArr As Variant)
        Dim I As Long
        Application.ScreenUpdating = False
        For I = LBound(findArr) To UBound(findArr)
            Selection.HomeKey Unit:=wdStory
            With Selection.Find
                .ClearFormatting
                .Replacement.ClearFormatting
                .Text = findArr(I)
                .Replacement.Text = replaceArr(I)
                .Format = False
                .MatchWholeWord = False
            End With
            Selection.Find.Execute Replace:=wdReplaceAll
        Next
        Application.ScreenUpdating = True
    End Sub
    """

    # Adicionar o código VBA ao documento
    module = doc.VBProject.VBComponents.Add(1)  # 1 é para módulo
    module.CodeModule.AddFromString(vba_code)

    # Preparar arrays VBA
    findArr = win32.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_VARIANT, find_list)
    replaceArr = win32.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_VARIANT, replace_list)

    # Executar a macro
    doc.Application.Run("FindAndReplaceMultiItems", findArr, replaceArr)

    # Remover o módulo VBA
    doc.VBProject.VBComponents.Remove(module)


# Listas de palavras a serem substituídas
find_list_1 = ["http://www.h3c.com.hk", "http://www.h3c.com", "New H3C Technologies Co.,Ltd.", "h3c", "H3C",
               "AD-Campus", "Unified Platform", "AD-DC", "AD-WAN", "SeerAnalyzer", "AD-NET", " 6.5", "iMC", "H3Linux",
               "SeerEngine", "SeerEngine-Campus controller"]
replace_list_1 = ["http://www.intelbras.com.br", "http://www.intelbras.com.br", "Intelbras S.A", "intelbras",
                  "INTELBRAS", "INC - AD Campus", "INC - Unified Platform", "INC - AD DC", "INC - AD WAN",
                  "INC - SeerAnalyzer", "INC - AD NET", "", "iNC", "IBLinux", "INC - SeerEngine",
                  "INC - SeerEngine Campus"]

find_list_2 = ["S5130S-EI", "S5170-EI", "S5570S-EI", "S6520X-EI", "S6520X-SI", "S6530X", "S5560X and S6520X",
               "S5560X or S6520X", "S5590", "S5590 or S5590XP", "EIA", "WSM"]
replace_list_2 = ["SC 3130", "SC 3170", "SC 3570", "SC 5525", "SC 5520", "SC 5530", "SC 5525", "SC 5525", "SC 3590",
                  "SC 3590", "INC - EIA", "INC - WSM"]

# Caminho para a pasta com os arquivos DOCX
folder_path = r'C:\Users\gu060589\Desktop\495\Versão H3C'

# Inicializar o Word
word = win32.Dispatch("Word.Application")
word.Visible = False

# Itera sobre todos os arquivos DOCX na pasta
for filename in os.listdir(folder_path):
    if filename.endswith(".docx"):
        doc_path = os.path.join(folder_path, filename)
        doc = word.Documents.Open(doc_path)

        # Substitui as palavras no documento usando VBA
        replace_words_vba(doc, find_list_1, replace_list_1)
        replace_words_vba(doc, find_list_2, replace_list_2)

        # Salva e fecha o documento
        doc.SaveAs(os.path.join(folder_path, f"modified_{filename}"))
        doc.Close()

# Fechar o Word
word.Quit()

print("Substituições concluídas.")
