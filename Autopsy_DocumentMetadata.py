from ms_cfbf import CFBF
from pdf import PDF
from ms_ooxml import OOXML
#from datetime import datetime

import inspect
import os
#import jarray
#import csv
#import time

from java.util.logging import Level
from java.io import File
from org.sleuthkit.datamodel import SleuthkitCase
from org.sleuthkit.datamodel import AbstractFile
from org.sleuthkit.datamodel import BlackboardArtifact
from org.sleuthkit.datamodel import BlackboardAttribute
from org.sleuthkit.autopsy.ingest import IngestModule
from org.sleuthkit.autopsy.ingest.IngestModule import IngestModuleException
from org.sleuthkit.autopsy.ingest import DataSourceIngestModule
from org.sleuthkit.autopsy.ingest import IngestModuleFactoryAdapter
from org.sleuthkit.autopsy.ingest import IngestMessage
from org.sleuthkit.autopsy.ingest import IngestServices
from org.sleuthkit.autopsy.coreutils import Logger
from org.sleuthkit.autopsy.casemodule import Case
from org.sleuthkit.autopsy.datamodel import ContentUtils
from org.sleuthkit.autopsy.casemodule.services import FileManager
from java.lang import IllegalArgumentException

class DocumentMetadataIngestModuleFactory(IngestModuleFactoryAdapter):
    def __init__(self):
        self.settings = None

    moduleName = "DocumentMetadataParser"

    def getModuleDisplayName(self):
        return self.moduleName

    def getModuleDescription(self):
        return "DocumentMeadataParser"

    def getModuleVersionNumber(self):
        return "1.0"

    def isDataSourceIngestModuleFactory(self):
        return True

    def createDataSourceIngestModule(self, ingestOptions):
        return DocumentMetadataIngestModule(self.settings)

class DocumentMetadataIngestModule(DataSourceIngestModule):

    _logger = Logger.getLogger(DocumentMetadataIngestModuleFactory.moduleName)

    def log(self, level, msg):
        self._logger.logp(level, self.__class__.__name__, inspect.stack()[1][3], msg)

    def __init__(self, settings):
        self.context = None

        self.totalCount = 0
        self.cfbf = CFBF()
        self.pdf = PDF()
        self.ooxml = OOXML()

        self.pdf_result = []
        self.xls_result = []
        self.ppt_result = []
        self.doc_result = []
        self.xlsx_result = []
        self.pptx_result = []
        self.docx_result = []
        self.PyPDF_result = []
        self.PDFNoModule_result = []

    def startUp(self, context):
        self.context = context
        pass

    def getTitles(self, result):
        titles = []
        for i in result:
            for j in i.keys():
                if j == "handle":
                    continue
                titles.append(j)
        return list(set(titles))

    def addData(self, titles, result, filetype, skCase):
        for title in titles:
            try:
                attID = skCase.addArtifactAttributeType("TSK_"+filetype+"_"+str(title), BlackboardAttribute.TSK_BLACKBOARD_ATTRIBUTE_VALUE_TYPE.STRING, unicode(title))
                artID_art = skCase.addBlackboardArtifactType("TSK_"+filetype+"_DATA", filetype)
            except:
                pass

        getArtId = skCase.getArtifactTypeID("TSK_"+filetype+"_DATA")

        for i in result:
            art = i["handle"].newArtifact(getArtId)
            for title in titles:
                try:
                    art.addAttribute(BlackboardAttribute(skCase.getAttributeType("TSK_"+filetype+"_"+str(title)), DocumentMetadataIngestModuleFactory.moduleName, unicode(i[title])))
                except:
                    art.addAttribute(BlackboardAttribute(skCase.getAttributeType("TSK_"+filetype+"_"+str(title)), DocumentMetadataIngestModuleFactory.moduleName, ""))

    def startModule(self, extension, skCase, fileManager, dataSource, progressBar):
        files = fileManager.findFiles(dataSource, "%." + extension)
        numFiles = len(files)
        progressBar.switchToDeterminate(numFiles)
        fileCount = 0

        Directory = os.path.join(Case.getCurrentCase().getTempDirectory(), extension + " files")
        try:
            os.mkdir(Directory)
        except:
            pass

        for file in files:
            self.log(Level.INFO, "Processing file: " + file.getName())
            fileCount += 1
            self.totalCount += 1

            Path = os.path.join(Directory, unicode(file.getName()))
            ContentUtils.writeToFile(file, File(Path))

            if extension.lower() == "pdf":
                try:
                    resultPDF, resultNoModule = self.pdf.run(Path)
                    resultPDF["handle"] = file
                    resultNoModule["handle"] = file
                    self.PyPDF_result.append(resultPDF)
                    self.PDFNoModule_result.append(resultNoModule)
                except:
                    pass

            elif extension.lower() == "xlsx":
                try:
                    OOXML = self.ooxml.run(Path)
                    OOXML["handle"] = file
                    self.xlsx_result.append(OOXML)
                except:
                    pass

            elif extension.lower() == "pptx":
                try:
                    OOXML = self.ooxml.run(Path)
                    OOXML["handle"] = file
                    self.pptx_result.append(OOXML)
                except:
                    pass

            elif extension.lower() == "docx":
                try:
                    OOXML = self.ooxml.run(Path)
                    OOXML["handle"] = file
                    self.docx_result.append(OOXML)
                except:
                    pass

            elif extension.lower() == "xls":
                try:
                    resultCFBF = self.cfbf.run(Path)
                    resultCFBF["handle"] = file
                    self.xls_result.append(resultCFBF)
                except:
                    pass

            elif extension.lower() == "ppt":
                try:
                    resultCFBF = self.cfbf.run(Path)
                    resultCFBF["handle"] = file
                    self.ppt_result.append(resultCFBF)
                except:
                    pass

            elif extension.lower() == "doc":
                try:
                    resultCFBF = self.cfbf.run(Path)
                    resultCFBF["handle"] = file
                    self.doc_result.append(resultCFBF)
                except:
                    pass

            progressBar.progress(fileCount)


        if extension.lower() == "pdf":
            titles = self.getTitles(self.PyPDF_result)
            self.addData(titles, self.PyPDF_result, "PyPDF", skCase)
            titles = self.getTitles(self.PDFNoModule_result)
            self.addData(titles, self.PDFNoModule_result, "PDFNoModule", skCase)

        elif extension.lower() == "xlsx":
            titles = self.getTitles(self.xlsx_result)
            self.addData(titles, self.xlsx_result, "XLSX", skCase)
        
        elif extension.lower() == "pptx":
            titles = self.getTitles(self.pptx_result)
            self.addData(titles, self.pptx_result, "PPTX", skCase)
        
        elif extension.lower() == "docx":
            titles = self.getTitles(self.docx_result)
            self.addData(titles, self.docx_result, "DOCX", skCase)

        elif extension.lower() == "xls":
            titles = self.getTitles(self.xls_result)
            self.addData(titles, self.xls_result, "XLS", skCase)

        elif extension.lower() == "ppt":
            titles = self.getTitles(self.ppt_result)
            self.addData(titles, self.ppt_result, "PPT", skCase)

        elif extension.lower() == "doc":
            titles = self.getTitles(self.doc_result)
            self.addData(titles, self.doc_result, "DOC", skCase)


    def process(self, dataSource, progressBar):
        progressBar.switchToIndeterminate()
        skCase = Case.getCurrentCase().getSleuthkitCase()
        fileManager = Case.getCurrentCase().getServices().getFileManager()
        
        self.startModule("pdf", skCase, fileManager, dataSource, progressBar)
        self.startModule("docx", skCase, fileManager, dataSource, progressBar)
        self.startModule("pptx", skCase, fileManager, dataSource, progressBar)
        self.startModule("xlsx", skCase, fileManager, dataSource, progressBar)
        self.startModule("doc", skCase, fileManager, dataSource, progressBar)
        self.startModule("ppt", skCase, fileManager, dataSource, progressBar)
        self.startModule("xls", skCase, fileManager, dataSource, progressBar)

        message = IngestMessage.createMessage(IngestMessage.MessageType.DATA, "DocumentMetadataParser", "Found %d files" % self.totalCount)
        IngestServices.getInstance().postMessage(message)

        return IngestModule.ProcessResult.OK