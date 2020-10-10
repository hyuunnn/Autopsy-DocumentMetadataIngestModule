# -*- coding: utf-8 -*-

# https://github.com/booktype/python-ooxml
# https://github.com/libyal/libolecf
# https://github.com/grierforensics/officedissector

from itertools import tee
import xmltodict
import zipfile
import re
import struct
import tempfile


class OOXML:
    def __init__(self):
        self.dictData = ""
        self.meta = {}

    def __dictGet(self, rootxmlKey, key):
        metaData = self.dictData[rootxmlKey].get(key, "")
        if metaData is None:
            self.meta[key] = ""
        elif '#text' in metaData:
            self.meta[key] = metaData['#text']
        else:
            self.meta[key] = metaData

    def __parsingCore(self, ooxmlFile, name):
        xmlData = ooxmlFile.open(name.lstrip('/'))
        self.dictData = xmltodict.parse(xmlData.read())

        # Metadata that must be included
        self.__dictGet('cp:coreProperties', 'dc:title')
        self.__dictGet('cp:coreProperties', 'dc:creator')
        self.__dictGet('cp:coreProperties', 'dcterms:created')
        self.__dictGet('cp:coreProperties', 'cp:lastModifiedBy')
        self.__dictGet('cp:coreProperties', 'dcterms:modified')
        self.__dictGet('cp:coreProperties', 'cp:lastPrinted')
        self.__dictGet('cp:coreProperties', 'dc:subject')
        self.__dictGet('cp:coreProperties', 'dc:description')

        # Additional metadata
        self.__dictGet('cp:coreProperties', 'cp:revision') # 수정 횟수
        self.__dictGet('cp:coreProperties', 'cp:keywords') 

    def __parsingApp(self, ooxmlFile, name):
        xmlData = ooxmlFile.open(name.lstrip('/'))
        self.dictData = xmltodict.parse(xmlData.read())

        # Metadata that must be included
        self.__dictGet('Properties', 'Pages')
        self.__dictGet('Properties', 'Words')
        self.__dictGet('Properties', 'Template')
        self.__dictGet('Properties', 'TotalTime')
        self.__dictGet('Properties', 'Application')
        self.__dictGet('Properties', 'AppVersion')
        self.__dictGet('Properties', 'Slides')

        # Additional metadata
        self.__dictGet('Properties', 'Lines')
        self.__dictGet('Properties', 'Notes')
        self.__dictGet('Properties', 'Paragraphs')
        self.__dictGet('Properties', 'HiddenSlides')
        self.__dictGet('Properties', 'Characters')
        self.__dictGet('Properties', 'CharactersWithSpaces')

    def __remake(self, zipPath, xmlFileName):
        # extraction app.xml
        data = open(zipPath, 'rb').read()
        tf = tempfile.NamedTemporaryFile()
        if xmlFileName == 'app.xml':
            iterator = re.finditer(b'docProps/app.xml', data)
            fileNameLength = 16
        else:
            iterator = re.finditer(b'docProps/core.xml', data)
            fileNameLength = 17

        first_it, second_it = tee(iterator)
        if sum(1 for _ in first_it) != 2:
            return

        for idx, val in enumerate(second_it):
            if idx == 0:
                frFileNameIndex = val.start()
                frExtraFieldLength = struct.unpack('<H', data[frFileNameIndex - 2:frFileNameIndex])[0]
                frCompressedSize = struct.unpack('<L', data[frFileNameIndex - 12:frFileNameIndex - 8])[0]
                record = data[frFileNameIndex - 30:frFileNameIndex + fileNameLength + frExtraFieldLength + frCompressedSize]
                tf.write(record)

            if idx == 1:
                deFileNameIndex = val.start()
                deFileCommentLength = struct.unpack('<H', data[deFileNameIndex - 14:deFileNameIndex - 12])[0]
                deExtraFieldLength = struct.unpack('<H', data[deFileNameIndex - 16:deFileNameIndex - 14])[0]

                fixed_data = data[deFileNameIndex - 46:deFileNameIndex - 4] + b'\x00\x00\x00\x00' + data[deFileNameIndex:deFileNameIndex + fileNameLength + deFileCommentLength + deExtraFieldLength]
                tf.write(fixed_data)

        endLocator = b'\x50\x4B\x05\x06\x00\x00\x00\x00\x01\x00\x01\x00' + struct.pack('<L', len(fixed_data)) + struct.pack('<L', len(record)) + b'\x00\x00'
        tf.write(endLocator)

        return tf


    def run(self, path):
        # https://github.com/grierforensics/officedissector/blob/2059a5ba08fa139362e3936578f99c4da9a9b55d/officedissector/part.py#L47
        self.meta = {}

        try:
            ooxmlFile = zipfile.ZipFile(path)
            ooxmlInfo = ooxmlFile.namelist()

            for name in ooxmlInfo:
                if "app.xml" in name:
                    try:
                        self.__parsingApp(ooxmlFile, name)
                    except:
                        continue
                elif "core.xml" in name:
                    try:
                        self.__parsingCore(ooxmlFile, name)
                    except:
                        continue

        except zipfile.BadZipFile:
            app_tf = self.__remake(path, 'app.xml')
            if app_tf is not None:
                try:
                    ooxmlFile = zipfile.ZipFile(app_tf)
                    self.__parsingApp(ooxmlFile, 'docProps/app.xml')
                except:
                    pass

            core_tf = self.__remake(path, 'core.xml')
            if core_tf is not None:
                try:
                    ooxmlFile = zipfile.ZipFile(core_tf)
                    self.__parsingCore(ooxmlFile, 'docProps/core.xml')
                except:
                    pass

        return self.meta