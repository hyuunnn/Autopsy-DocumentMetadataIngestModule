# -*- coding: utf-8 -*-

# https://pythonhosted.org/PyPDF2/PdfFileReader.html
# https://github.com/zbetcheckin/PDF_analysis

# https://www.meridiandiscovery.com/articles/pdf-forensic-analysis-xmp-metadata/
# https://www.pdfa.org/wp-content/until2016_uploads/2011/08/pdfa_metadata-2b.pdf

import PyPDF2
import re
import unicodedata
from bs4 import BeautifulSoup

regex_name1 = ["Author", "Title", "Subject", "Keywords", "CreationDate", "ModDate", "Creator", "Producer", "Application", "Appligent"]
regex_name2 = ["pdf:PDFVersion", "pdf:Producer", "pdf:Trapped", "pdf:Keywords", "pdfx:Company", "pdfx:SourceModified", "dc:format"\
            "dc:subject", "dc:creator", "dc:title", "dc:date", "dc:publisher", "dc:rights",\
            "dc:contributor", "dc:coverage", "dc:description", "dc:identifier", "dc:relation", "dc:source", "dc:type",\
            "xmp:CreateDate", "xmp:ModifyDate", "xmp:MetadataDate", "xmp:CreatorTool", "xmp:Title"]

class PDF:
    def __init__(self):
        self.regex_list = []
        self.regex_list2 = []

        for i in regex_name1:
            self.re_compile1(i)
            self.re_compile3(i)

        for i in regex_name2:
            self.re_compile2(i)

    def data_replace(self, data):
        return data.replace(b"<rdf:Seq>",b"").replace(b"<rdf:li>",b"")\
            .replace(b"</rdf:Seq>",b"").replace(b"<rdf:Seq/>", b"").replace(b"</rdf:li>",b"")\
            .replace(b"<rdf:Alt>",b"").replace(b"</rdf:Alt>",b"")\
            .replace(b"<rdf:li xml:lang=\"x-default\">",b"")

    def re_compile1(self, data):
        self.regex_list.append([re.compile("\/{}\s?\<[^/]*.\>".format(data).encode()), data])
        self.regex_list.append([re.compile("\/{}\s?\([^/]*.\)".format(data).encode()), data])

    def re_compile2(self, data):
        self.regex_list.append([re.compile("\<{}\>.*\<\/{}\>".format(data, data).encode()), data])

    def re_compile3(self, data):
        self.regex_list2.append([re.compile("\/{}\s\d+\s\d+".format(data).encode()), data])

    def decodeString(self, data):
        try:
            return data.decode("cp949")
        except:
            pass

        try:
            return unicodedata.normalize("NFC", data.decode("utf-8"))
        except:
            pass

        try:
            return unicodedata.normalize("NFC", data.decode("unicode_escape").encode("latin1").decode("utf-16"))
        except:
            pass
        
        return data

    def regexParse(self, data):
        self.meta = {}
        for regex, name in self.regex_list:
            regex_data = regex.search(data)
            if regex_data:
                result = self.data_replace(regex_data.group())
                if result.startswith((b"<pdf:", b"<pdfx:", b"<dc:", b"<xmp:")):
                    self.meta[name] = self.decodeString(BeautifulSoup(result, 'html.parser').text)
                else:
                    result = result.replace("/{}(".format(name).encode(), b"")\
                                    .replace("/{} (".format(name).encode(), b"")\
                                    .replace("/{}<".format(name).encode(), b"")\
                                    .replace("/{} <".format(name).encode(), b"")[:-1]
                    self.meta[name] = self.decodeString(result)

        for regex, name in self.regex_list2:
            regex_data = regex.search(data)
            if regex_data:
                result = regex_data.group().replace(b"\n",b" ").replace("/{} ".format(name).encode(), b"") 
                # /Title\n1589 0 케이스 때문에 \n replace -> /Title 1589 0에서 /Title 을 공백으로 replace
                result = result.replace(b" ", b"\s") + b"\sobj\x0A\(.*\)\x0Aendobj" 
                # 공백 정규식 처리 + obj 데이터 정규식 (ex: 1592 0 obj.(D:20141127125117Z00'00').endobj)

                stage2 = re.compile(result)
                regex_data2 = stage2.search(data)
                if regex_data2:
                    result = regex_data2.group().split(b"\x0A")[1]
                    self.meta[name] = self.decodeString(result.replace(b"(",b"").replace(b")", b""))

        return self.meta

    def parseMetadata(self, path):
        with open(path, "rb") as f:
            data = f.read()
            return self.regexParse(data)

    def PyPDFParse(self, pdfFile):
        # https://pythonhosted.org/PyPDF2/DocumentInformation.html
        pdfinfo = dict(pdfFile.getDocumentInfo())
        pdfinfo["Pages"] = pdfFile.getNumPages()
        return pdfinfo

    def run(self, path):
        try:
            pdfFile = PyPDF2.PdfFileReader(open(path, "rb"))
            resultPyPDF = self.PyPDFParse(pdfFile)
        except Exception as e:
            #print("[*] PyPDF2 Error {}".format(path))
            #print(e)
            resultPyPDF = {}

        try:
            resultNoModule = self.parseMetadata(path)
        except Exception as e:
            #print("[*] No Module Parser Error {}".format(path))
            #print(e)
            resultNoModule = {}

        return resultPyPDF, resultNoModule