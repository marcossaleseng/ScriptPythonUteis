import win32com.client
import os
inputFolder = r'SEU ENDERECO COMPLETO\INPUT' ## Coloque aqui o endereço a pasta de entrada
outputFolder = r'SEU ENDERECO ABSOLUTO\output' ## Coloque aqui o endereço da pasta de saída


def retirarAnexos(inputFolder):
    """
    Pegar dos os anexos dos arquivos msg em uma pasta windows, utilizando o Outllook para extrair os anexos
    INPUT: inputFolder -> pasta com os arquivos
    OUTPUT: outputFolder -> pasta de saída
    RETURN: None
    """
    for file in os.listdir(inputFolder):
        if file.endswith(".msg"):
            outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
            filePath = inputFolder  + '\\' + file
            msg = outlook.OpenSharedItem(filePath)
            att = msg.Attachments
            for i in att:
                i.SaveAsFile(os.path.join(outputFolder, i.FileName)) #Save os arquivos com os nomes dos anexos
    return None

## colocar o código para rodar
retirarAnexos(inputFolder)