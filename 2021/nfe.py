'''
TO DO

OK Horario de saida: sempre as 10 da manha. Se emitida antes das 10:00, a data no mesmo dia, caso
    contrario, no ida seguinte

INDUSTRIALIZAÇÃO ou RETORNO DE INDUSTRIALIZAÇÃO: são sempre emitidas duas notas, uma de cada tipo
    RETORNO DE INDUSTRIALIZAÇÃO: muda o nome da nfe, o codigozinho das placas, e infos adicionais

OK Numero da NFe: ta aqui o desafio

OK Data de vencimento

OK Minucias de cada item

Data da ultima NFe

CHECAR: cDV ???????????
        cNF ???????????
        NCM ???????????
'''

import xlrd
from datetime import datetime
import calendar

#---------------------------------------------------------------------#
#-----------------------      DADOS GERAIS     -----------------------#
#---------------------------------------------------------------------#

FIRST_COLUNM = 2
BOARD_CODE = 1
BOARD_NAME = 0

# Give the location of the file
loc = ("Estoque VRE 2021.xlsx")
 
# To open Workbook
wb = xlrd.open_workbook(loc)

sheet_saida = wb.sheet_by_name("Saida") #pegando quantiadde
sheet_placas = wb.sheet_by_name("Placas") #pegadno infos da placa

#numero da NFe
NF_NUM = int(sheet_saida.cell_value(0, 3))

#Data de vencimento
DATA_VENCIMENTO = datetime(*xlrd.xldate_as_tuple(sheet_saida.cell_value(0, 1), 0)) 
DATA_VENCIMENTO = str(f"{DATA_VENCIMENTO.day:02d}") + "\\" + str(DATA_VENCIMENTO.month) + "\\" + str(DATA_VENCIMENTO.year)

#Data e hora de emissao
#2021-02-12T09:02:00-03:00
DHEMI = str(datetime.today().year) + "-" + str(f"{datetime.today().month:02d}") + "-" + str(f"{datetime.today().day:02d}") + "T" + str(f"{datetime.today().hour:02d}") + ":" + str(f"{datetime.today().minute:02d}") + ":00-03:00"

#Caluclando data e hora de saida dhSaiEnt
#2021-02-05T10:00:00-03:00
if datetime.today().hour < 10:
    dhsaient_dia = datetime.today().day
    dhsaient_mes = datetime.today().month
else:
    if datetime.today().day + 1 > calendar.monthrange(datetime.today().year, datetime.today().month)[1]:
        dhsaient_dia = 1
        dhsaient_mes = datetime.today().month + 1
        #BUGA SE UMA NOTA FOR EMITIDA APOS AS 10:00 DO ULTIMO DIA DO ANO
    else:
        dhsaient_dia = datetime.today().day + 1
        dhsaient_mes = datetime.today().month
DHSAIENT = str(datetime.today().year) + "-" + str(f"{dhsaient_mes:02d}") + "-" + str(f"{dhsaient_dia:02d}") + "T10:00:00-03:00"

#Data da ultima NFe PARA FAZER
DATA_ULTIMA_NFe = "TO_DO"

#PRINTADO TODAS CONSTANTES
print("Numero da nf: \t" + str(NF_NUM))
print("Data de vencimento: \t" + DATA_VENCIMENTO)
print("Data de emissao: \t" + DHEMI)
print("Data saida: \t" + DHSAIENT)


#---------------------------------------------------------------------#
#------------------------  PLACAS E QUANTIDADE  ----------------------#
#---------------------------------------------------------------------#

#Todas colunas
num_rows= sheet_saida.nrows

i = FIRST_COLUNM
lastCol = 0
while i < sheet_saida.ncols:
    if sheet_saida.cell_value(1, i) != "":
        lastCol = i
    else:
        break
    i = i + 1

i = 0
products = []
for i in range(2, num_rows):
    if sheet_saida.cell_value(i, BOARD_CODE) == 0:
        break;
    else:
        if sheet_saida.cell_value(i, lastCol) != "":
            # Codicom | Nome | Código | Quantidade | Valor unitário | Valor total
            product = [sheet_placas.cell_value(i + 1, BOARD_NAME + 2), #Codicom
                       sheet_saida.cell_value(i, BOARD_NAME),#Nome
                       sheet_placas.cell_value(i + 1, BOARD_NAME + 1), #Código
                       int(sheet_saida.cell_value(i, lastCol)), #Quantidade
                       sheet_placas.cell_value(i + 1, BOARD_NAME + 3), #Valor unitario
                       sheet_placas.cell_value(i + 1, BOARD_NAME + 3) * sheet_saida.cell_value(i, lastCol)] #Valor total
            
            products.append(product)

'''for elem in products:
    print(str(int(elem[0])) + " " + elem[1] + " " + str(elem[2]) + " " + str(int(elem[3])) + " " + str(elem[4]) + " " + str(elem[3] * elem[4]))'''

def criarXML(file, num, tipo = 1):
    if tipo == 1:
        nome_tipo = "INDUSTRIALIZAÇÃO"
        CFOP = str(5124)
        cDV = 3
    else:
        nome_tipo  = "RETORNO INDUSTRIALIZAÇÃO"
        CFOP = str(5902)
        cDV = 0
    #Intro
    file.write("<?xml version=\"1.0\" encoding=\"UTF-8\"?><nfeProc versao=\"4.00\" xmlns=\"http://www.portalfiscal.inf.br/nfe\"><NFe xmlns=\"http://www.portalfiscal.inf.br/nfe\"><infNFe Id=\"NFe3521030273360900010355001000000327112070910"+str(cDV)+"\" versao=\"4.00\"><ide><cUF>35</cUF><cNF>12070910</cNF><natOp>"+nome_tipo+"</natOp><mod>55</mod><serie>1</serie><nNF>"+str(num)+"</nNF><dhEmi>"+DHEMI+"</dhEmi><dhSaiEnt>"+DHSAIENT+"</dhSaiEnt><tpNF>1</tpNF><idDest>1</idDest><cMunFG>3550308</cMunFG><tpImp>1</tpImp><tpEmis>1</tpEmis><cDV>"+str(cDV)+"</cDV><tpAmb>1</tpAmb><finNFe>1</finNFe><indFinal>1</indFinal><indPres>0</indPres><procEmi>3</procEmi><verProc>4.01_b030</verProc></ide><emit><CNPJ>02733609000103</CNPJ><xNome>VRE  ELETROELETRONICA LTDA -ME</xNome><xFant>VRE</xFant><enderEmit><xLgr>RUA ITAMBACURI</xLgr><nro>21</nro><xBairro>JD ORIENTAL</xBairro><cMun>3550308</cMun><xMun>Sao Paulo</xMun><UF>SP</UF><CEP>04347070</CEP><cPais>1058</cPais><xPais>BRASIL</xPais><fone>50113931</fone></enderEmit><IE>148041835119</IE><IM>27228428</IM><CNAE>2610800</CNAE><CRT>1</CRT></emit><dest><CNPJ>66956160000117</CNPJ><xNome>Infolev Elevadores &amp; Informatica Ltda</xNome><enderDest><xLgr>Rua Sara de Souza</xLgr><nro>152</nro><xBairro>Agua Branca</xBairro><cMun>3550308</cMun><xMun>Sao Paulo</xMun><UF>SP</UF><CEP>05037140</CEP><cPais>1058</cPais><xPais>BRASIL</xPais><fone>33831900</fone></enderDest><indIEDest>1</indIEDest><IE>113268959112</IE></dest>")
    #Registrando placa por placa
    valor_total = 0
    for elem in products:
        valor_total = valor_total + elem[5]
        #   0     |   1  |    2   |      3     |       4        |      5
        # Codicom | Nome | Código | Quantidade | Valor unitário | Valor total 
        file.write("<det nItem=\""+str(products.index(elem) + 1)+"\"><prod><cProd>"+f"{int(elem[0]):09d}"+"</cProd><cEAN/><xProd>MONTAGEM PLACA "+elem[1]+" "+elem[2]+"</xProd><NCM>85389010</NCM><CFOP>"+CFOP+"</CFOP><uCom>PC</uCom><qCom>"+str(f"{elem[3]:.4f}")+"</qCom><vUnCom>"+str(f"{elem[4]:.10f}")+"</vUnCom><vProd>"+str(f"{elem[5]:.2f}")+"</vProd><cEANTrib/><uTrib>PC</uTrib><qTrib>"+str(f"{elem[3]:.4f}")+"</qTrib><vUnTrib>"+str(f"{elem[4]:.10f}")+"</vUnTrib><indTot>1</indTot><xPed>02360</xPed></prod><imposto><ICMS><ICMSSN102><orig>0</orig><CSOSN>102</CSOSN></ICMSSN102></ICMS><PIS><PISOutr><CST>99</CST><qBCProd>0.0000</qBCProd><vAliqProd>0.0000</vAliqProd><vPIS>0.00</vPIS></PISOutr></PIS><COFINS><COFINSOutr><CST>99</CST><qBCProd>0.0000</qBCProd><vAliqProd>0.0000</vAliqProd><vCOFINS>0.00</vCOFINS></COFINSOutr></COFINS></imposto></det>")

    #Final
    file.write("<total><ICMSTot><vBC>0.00</vBC><vICMS>0.00</vICMS><vICMSDeson>0.00</vICMSDeson><vFCPUFDest>0.00</vFCPUFDest><vICMSUFDest>0.00</vICMSUFDest><vICMSUFRemet>0.00</vICMSUFRemet><vFCP>0.00</vFCP><vBCST>0.00</vBCST><vST>0.00</vST><vFCPST>0.00</vFCPST><vFCPSTRet>0.00</vFCPSTRet><vProd>"+str(f"{valor_total:.2f}")+"</vProd><vFrete>0.00</vFrete><vSeg>0.00</vSeg><vDesc>0.00</vDesc><vII>0.00</vII><vIPI>0.00</vIPI><vIPIDevol>0.00</vIPIDevol><vPIS>0.00</vPIS><vCOFINS>0.00</vCOFINS><vOutro>0.00</vOutro><vNF>"+str(f"{valor_total:.2f}")+"</vNF><vTotTrib>0.00</vTotTrib></ICMSTot></total><transp><modFrete>3</modFrete></transp><cobr><fat><nFat>NF "+str(num - 1)+" VENCIMENTO: DIA "+DATA_VENCIMENTO+"</nFat><vOrig>"+str(f"{valor_total:.2f}")+"</vOrig><vDesc>0.00</vDesc><vLiq>"+str(f"{valor_total:.2f}")+"</vLiq></fat></cobr><pag><detPag><indPag>1</indPag><tPag>99</tPag><vPag>"+str(f"{valor_total:.2f}")+"</vPag></detPag></pag><infAdic><infCpl>I- DOCUMENTO EMITIDO POR ME OU EPP OPTANTE POR SIMPLES NACIONAL; II- NÃO GERA DIREITO A CREDITO FISCAL DE ISS E IPI; COBRANÇA REF. A NF "+str(num-1)+" DO DIA "+str(DATA_ULTIMA_NFe)+".</infCpl></infAdic><infRespTec><CNPJ>43728245000142</CNPJ><xContato>suporte</xContato><email>suporteemissores@sebraesp.com.br</email><fone>08005700800</fone></infRespTec></infNFe><Signature xmlns=\"http://www.w3.org/2000/09/xmldsig#\"><SignedInfo><CanonicalizationMethod Algorithm=\"http://www.w3.org/TR/2001/REC-xml-c14n-20010315\"/><SignatureMethod Algorithm=\"http://www.w3.org/2000/09/xmldsig#rsa-sha1\"/><Reference URI=\"#NFe3521030273360900010355001000000327112070910"+str(cDV)+"\"><Transforms><Transform Algorithm=\"http://www.w3.org/2000/09/xmldsig#enveloped-signature\"/><Transform Algorithm=\"http://www.w3.org/TR/2001/REC-xml-c14n-20010315\"/></Transforms><DigestMethod Algorithm=\"http://www.w3.org/2000/09/xmldsig#sha1\"/><DigestValue>9zJ2xlr+ab8ru4+6Lw/MKsPn5dI=</DigestValue></Reference></SignedInfo><SignatureValue>GjhIz8H3HtlzJytln8NHbMD4Eu/qFuM4x7ed9UsihHbLwVVI+QVqSEo2Ws35F6+ASwshL8R/ykS/&#13;\n"+
                "d2jda3Z5vkQcEiacoagyVuGCfce9xyK+THYP9zddd0rwt39dV6u5w9w3jsRT/g9oLl+kpHSpBYlK&#13;\n"+
                "JAnaRuqnU10BXqpSFu2+W0Hy8B67tXR34N8ZoOHavKNR9gBaQvI8eE38Jf4IDFXo4Ma5vcj4JdCh&#13;\n"+
                "jtLKawBoHtvK1nSnG//BlZKW8LPFwJw4XUXj8S2XbyfXfEPYLOUayHrJby+LwYyoEG2k3iUW3I3f&#13;\n"+
                "3+s7XSAXL737yBsdapU09BUc1DwNNYipnDRStg==</SignatureValue><KeyInfo><X509Data><X509Certificate>MIIHpTCCBY2gAwIBAgIIbNgWf4xRw8cwDQYJKoZIhvcNAQELBQAwdTELMAkGA1UEBhMCQlIxEzAR&#13;\n"+
                "BgNVBAoMCklDUC1CcmFzaWwxNjA0BgNVBAsMLVNlY3JldGFyaWEgZGEgUmVjZWl0YSBGZWRlcmFs&#13;\n"+
                "IGRvIEJyYXNpbCAtIFJGQjEZMBcGA1UEAwwQQUMgU0VSQVNBIFJGQiB2NTAeFw0xODA2MjEyMTAw&#13;\n"+
                "MDBaFw0yMTA2MjAyMTAwMDBaMIHjMQswCQYDVQQGEwJCUjELMAkGA1UECAwCU1AxEjAQBgNVBAcM&#13;\n"+
                "CVNBTyBQQVVMTzETMBEGA1UECgwKSUNQLUJyYXNpbDE2MDQGA1UECwwtU2VjcmV0YXJpYSBkYSBS&#13;\n"+
                "ZWNlaXRhIEZlZGVyYWwgZG8gQnJhc2lsIC0gUkZCMRYwFAYDVQQLDA1SRkIgZS1DTlBKIEEzMRkw&#13;\n"+
                "FwYDVQQLDBBBUiBBN1lURUNOT0xPR0lBMTMwMQYDVQQDDCpWIFIgRSBFTEVUUk9FTEVUUk9OSUNB&#13;\n"+
                "IExUREE6MDI3MzM2MDkwMDAxMDMwggEiMA0GCSqGSIb3DQEBAQUAA4IBDwAwggEKAoIBAQCSy49I&#13;\n"+
                "60QBW27QAVeN5UQQO0Z/LEAMsD8914Meu1RT28M9w0Ju+KiWKxR7shn2kqq3T/mx38jPcdpU6sVM&#13;\n"+
                "Q6DXLGmF7PS5SCgOP0OJM4eRFAZwS1YaeD76dIdj/nCDfBkiQ3BPr9P4oYUGe+sV406Zm/fU9gBD&#13;\n"+
                "OSnqAL33J2CUj8xajWaRPvS8O/LqfkzysU+D9GpyYRC6m2BwO6s51C6fMXpGaYnM4EhvLmH0aQzH&#13;\n"+
                "puHohjeXQRggQHycsOsRUg++AnnqPQHUnKHnixjk0RG3W2QNbHBYJkgEmYhWn305HSXyJLHaqgNW&#13;\n"+
                "ZRqL6BfIV9HuwZk3rXVJRmnnfw94qjs1AgMBAAGjggLIMIICxDAJBgNVHRMEAjAAMB8GA1UdIwQY&#13;\n"+
                "MBaAFOzxQVFXqOY66V6zoCL5CIq1OoePMIGZBggrBgEFBQcBAQSBjDCBiTBIBggrBgEFBQcwAoY8&#13;\n"+
                "aHR0cDovL3d3dy5jZXJ0aWZpY2Fkb2RpZ2l0YWwuY29tLmJyL2NhZGVpYXMvc2VyYXNhcmZidjUu&#13;\n"+
                "cDdiMD0GCCsGAQUFBzABhjFodHRwOi8vb2NzcC5jZXJ0aWZpY2Fkb2RpZ2l0YWwuY29tLmJyL3Nl&#13;\n"+
                "cmFzYXJmYnY1MIG3BgNVHREEga8wgayBFkZJU0NBTDAxQEtBSUtFWS5DT00uQlKgHwYFYEwBAwKg&#13;\n"+
                "FhMURURTT04gSEVUU1VPIFVNRU1VUkGgGQYFYEwBAwOgEBMOMDI3MzM2MDkwMDAxMDOgPQYFYEwB&#13;\n"+
                "AwSgNBMyMTUxMTE5NTU0MzY0MjUwMDkwMDAwMDAwMDAwMDAwMDAwMDAwMDAxMjMxNDc5U1NQU1Cg&#13;\n"+
                "FwYFYEwBAwegDhMMMDAwMDAwMDAwMDAwMHEGA1UdIARqMGgwZgYGYEwBAgMKMFwwWgYIKwYBBQUH&#13;\n"+
                "AgEWTmh0dHA6Ly9wdWJsaWNhY2FvLmNlcnRpZmljYWRvZGlnaXRhbC5jb20uYnIvcmVwb3NpdG9y&#13;\n"+
                "aW8vZHBjL2RlY2xhcmFjYW8tcmZiLnBkZjAdBgNVHSUEFjAUBggrBgEFBQcDAgYIKwYBBQUHAwQw&#13;\n"+
                "gZ0GA1UdHwSBlTCBkjBKoEigRoZEaHR0cDovL3d3dy5jZXJ0aWZpY2Fkb2RpZ2l0YWwuY29tLmJy&#13;\n"+
                "L3JlcG9zaXRvcmlvL2xjci9zZXJhc2FyZmJ2NS5jcmwwRKBCoECGPmh0dHA6Ly9sY3IuY2VydGlm&#13;\n"+
                "aWNhZG9zLmNvbS5ici9yZXBvc2l0b3Jpby9sY3Ivc2VyYXNhcmZidjUuY3JsMA4GA1UdDwEB/wQE&#13;\n"+
                "AwIF4DANBgkqhkiG9w0BAQsFAAOCAgEAAGrx4goAF6EpoRowPHdyLf0IAYFrwLgY0jWmqClbfZuU&#13;\n"+
                "Eyf1JKz4b49nc792b4O1P+dXFguxERCtY/UNxThbXgFImuJARzZSAHDGSjgy2qe3mjOUulHgIkX5&#13;\n"+
                "sV7/2183Nh4oZzj9RetY0cv4FpTkQBsjkgrTs0cksq34HhLmUq+bBLtxV8Ce2rXlMh71l/Fm8x48&#13;\n"+
                "ERpvRoxPeIE8wgSdxhpGKFx6bhSTKhuuqPOZT6EI/yAgkwByxgxHeMMSSjlVXar45Bvrxr7bYUHU&#13;\n"+
                "rF9Y1aWlH51QU4gRfOjIwvd9imU7ysmbphAFfDKzr5Uo2e8w5G83zG0UXDfHGkDYlVgYmKvRgAY1&#13;\n"+
                "M4hF573AJAvqEHgsCSSIVeOLKWH769fBFl8EiefMNyizJTtgKZZB23Hx2WMHocA89qT3NOJQOnAq&#13;\n"+
                "TQVqKF1W59u8IsYiDpdlgQfWBFb5q05UNpNNm0vXpqfdWxESn3iTJWOwiPXmBm9+xn+jrFmyT4SM&#13;\n"+
                "ILNUGAhEgtl4+laHshExbyVvxtAHFRcMLTi3kiY/mfihqOowZAqGcKTvORAq0hVC2IH5k05IuNjc&#13;\n"+
                "W+X46gfvwEwQ8mc0Sg3/57Ef/Af+YJtMFICACNibd9dzd49gJlvdFn8wgVLY/a4fBazJn0hm2u0C&#13;\n"+
                "z/1qwXk5lKZS//OjQu3RKpaSI7HPy0I=</X509Certificate></X509Data></KeyInfo></Signature></NFe><protNFe versao=\"4.00\"><infProt><tpAmb>1</tpAmb><verAplic>SP_NFE_PL009_V4</verAplic><chNFe>3521030273360900010355001000000327112070910"+str(cDV)+"</chNFe><dhRecbto>"+DHEMI+"</dhRecbto><nProt>135210184246369</nProt><digVal>9zJ2xlr+ab8ru4+6Lw/MKsPn5dI=</digVal><cStat>100</cStat><xMotivo>Autorizado o uso da NF-e</xMotivo></infProt></protNFe></nfeProc>")

xml_ind = open(str(NF_NUM) + " industrializacao.xml",'w')
criarXML(xml_ind, NF_NUM) 
xml_ind.close()

xml_ret_ind = open(str(NF_NUM + 1) + " retorno de industrializacao.xml",'w')
criarXML(xml_ret_ind, NF_NUM + 1, 2) 
xml_ret_ind.close()