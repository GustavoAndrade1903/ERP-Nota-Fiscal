import pandas as pd
import xml.etree.ElementTree as ET
import os
import requests

# Função para validar o CEP usando a API ViaCEP
BASE_URL = "https://viacep.com.br/ws"

def obter_endereco(cep: str) -> dict:
    try:
        response = requests.get(f"{BASE_URL}/{cep}/json", timeout=30)
        if response.status_code == 200:
            return response.json()
    except Exception as e:
        print(f"Erro na chamada da API ViaCEP: {str(e)}")
    return {}

try:
    # 1. Ler o arquivo Excel
    excel_file = r"C:\Users\Cellula Mater\Desktop\Emitir Nota Fiscal\Arquivos RPS\Lote_09_Serie_Vivva.xlsx"
    df = pd.read_excel(excel_file)

    # 2. Carregar o template XML
    xml_file = r"C:\Users\Cellula Mater\Desktop\Emitir Nota Fiscal\TEMPLATE.xml"
    tree = ET.parse(xml_file)
    root = tree.getroot()

    # 3. Definir os namespaces
    namespaces = {
        "ns2": "http://www.giss.com.br/tipos-v2_04.xsd",
        "ns4": "http://www.giss.com.br/enviar-lote-rps-envio-v2_04.xsd"
    }

    # 4. Encontrar a tag <ns2:ListaRps> para adicionar as novas notas fiscais
    lista_rps = root.find(".//ns2:ListaRps", namespaces=namespaces)

    # 5. Variável para o contador sequencial
    contador_rps = 386

    # 6. Iterar sobre cada linha do Excel
    for index, row in df.iterrows():
        # Extrair valores do Excel
        dt_pagamento = row["Dt.Pagamento"]
        valor = row["Valor"]
        forma_pagamento = row["Forma de Pagamento"]
        pedido = row["Pedido"]
        cpf = row["ns2:Cpf"].replace("-", "").replace(".", "")  # Remover traços e pontos
        razao_social = row["ns2:RazaoSocial"]
        cep = row["ns2:Cep"]  # Obter o CEP do Excel
        numero = row["ns2:Numero"]

        # Chamar a API ViaCEP para obter as informações de endereço
        endereco_info = obter_endereco(cep)
        endereco = row["ns2:Endereco"]  # Manter o endereço do Excel como fallback
        bairro = ""
        codigo_municipio = ""
        uf = ""

        if endereco_info:
            endereco = endereco_info.get("logradouro", endereco)
            bairro = endereco_info.get("bairro", "")
            codigo_municipio = endereco_info.get("ibge", "")
            uf = endereco_info.get("uf", "")

        # Converter Timestamp para string no formato 'YYYY-MM-DD'
        dt_pagamento_str = dt_pagamento.strftime('%Y-%m-%d')

        # Criar a descrição do serviço
        discriminacao = f"Forma de Pagamento: {forma_pagamento}, Pedido: {pedido}, Data do pagamento: {dt_pagamento_str}"

        # 7. Copiar a tag <ns2:Rps> do template
        rps_template = lista_rps.find(".//ns2:Rps", namespaces=namespaces)
        novo_rps = ET.ElementTree(ET.fromstring(ET.tostring(rps_template))).getroot()

        # 8. Preencher os campos do novo RPS com as informações
        for elem in novo_rps.findall(".//ns2:DataEmissao", namespaces=namespaces):
            elem.text = dt_pagamento_str

        for elem in novo_rps.findall(".//ns2:Competencia", namespaces=namespaces):
            elem.text = dt_pagamento_str

        for elem in novo_rps.findall(".//ns2:ValorServicos", namespaces=namespaces):
            elem.text = str(valor)

        for elem in novo_rps.findall(".//ns2:Discriminacao", namespaces=namespaces):
            elem.text = discriminacao

        for elem in novo_rps.findall(".//ns2:Cpf", namespaces=namespaces):
            elem.text = cpf

        for elem in novo_rps.findall(".//ns2:RazaoSocial", namespaces=namespaces):
            elem.text = razao_social

        for elem in novo_rps.findall(".//ns2:Cep", namespaces=namespaces):
            elem.text = str(cep)
        
        # 9. Alterar o número do RPS e o ID da tag InfDeclaracaoPrestacaoServico
        for elem in novo_rps.findall(".//ns2:IdentificacaoRps/ns2:Numero", namespaces=namespaces):
            elem.text = str(contador_rps)

        for elem in novo_rps.findall(".//ns2:Rps", namespaces=namespaces):
            elem.attrib["Id"] = str(contador_rps)  # Modifica o atributo 'Id' da tag Rps

        for elem in novo_rps.findall(".//ns2:InfDeclaracaoPrestacaoServico", namespaces=namespaces):
            elem.attrib["Id"] = str(contador_rps)  # Modifica o atributo 'Id' da tag InfDeclaracaoPrestacaoServico
            
        endereco_tag = novo_rps.find(".//ns2:Endereco", namespaces=namespaces)
        if endereco_tag is not None:
            for subelem in endereco_tag:
                if subelem.tag.endswith("Endereco"):
                    subelem.text = endereco
                elif subelem.tag.endswith("Numero"):
                    subelem.text = str(numero)
                elif subelem.tag.endswith("Bairro"):
                    subelem.text = bairro
                elif subelem.tag.endswith("CodigoMunicipio"):
                    subelem.text = str(codigo_municipio)
                elif subelem.tag.endswith("Uf"):
                    subelem.text = uf
                elif subelem.tag.endswith("Cep"):
                    subelem.text = str(cep)

        # 9. Incrementar o contador para o próximo RPS
        contador_rps += 1

        # 10. Adicionar o novo RPS à lista de RPS
        lista_rps.append(novo_rps)

    # 11. Salvar o XML final
    output_file = r"C:\Users\Cellula Mater\Desktop\Emitir Nota Fiscal\NOTA_FISCAL_FINAL.xml"
    tree.write(output_file, encoding="UTF-8", xml_declaration=True)

    print(f"Arquivo {output_file} gerado com sucesso!")

except Exception as e:
    print(f"Ocorreu um erro: {e}")

# Pausa no final para evitar que o prompt feche
os.system("pause")
