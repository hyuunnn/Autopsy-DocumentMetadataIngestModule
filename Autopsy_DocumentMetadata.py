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

        self.CFBF_result = []
        self.PyPDF_result = []
        self.PDFNoModule_result = []
        self.OOXML_result = []

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

    def addData(self, titles, result, datatype, skCase):
        for title in titles:
            try:
                attID = skCase.addArtifactAttributeType("TSK_"+str(title), BlackboardAttribute.TSK_BLACKBOARD_ATTRIBUTE_VALUE_TYPE.STRING, unicode(title))
                artID_art = skCase.addBlackboardArtifactType("TSK_"+datatype+"_DATA", datatype)
            except:
                pass

        getArtId = skCase.getArtifactTypeID("TSK_"+datatype+"_DATA")

        for i in result:
            art = i["handle"].newArtifact(getArtId)
            for title in titles:
                try:
                    art.addAttribute(BlackboardAttribute(skCase.getAttributeType("TSK_"+str(title)), DocumentMetadataIngestModuleFactory.moduleName, unicode(i[title])))
                except:
                    art.addAttribute(BlackboardAttribute(skCase.getAttributeType("TSK_"+str(title)), DocumentMetadataIngestModuleFactory.moduleName, ""))

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

        if extension.lower() == "pdf":
            for file in files:
                self.log(Level.INFO, "Processing file: " + file.getName())
                fileCount += 1
                self.totalCount += 1

                Path = os.path.join(Directory, unicode(file.getName()))
                ContentUtils.writeToFile(file, File(Path))

                try:
                    resultPDF, resultNoModule = self.pdf.run(Path)
                    resultPDF["handle"] = file
                    resultNoModule["handle"] = file
                    self.PyPDF_result.append(resultPDF)
                    self.PDFNoModule_result.append(resultNoModule)
                except:
                    pass
                progressBar.progress(fileCount)

            titles = self.getTitles(self.PyPDF_result)
            self.addData(titles, self.PyPDF_result, "PyPDF", skCase)
            titles = self.getTitles(self.PDFNoModule_result)
            self.addData(titles, self.PDFNoModule_result, "PDFNoModule", skCase)

        elif any(x in extension.lower() for x in ["docx", "pptx", "xlsx"]):
            for file in files:
                self.log(Level.INFO, "Processing file: " + file.getName())
                fileCount += 1
                self.totalCount += 1

                Path = os.path.join(Directory, unicode(file.getName()))
                ContentUtils.writeToFile(file, File(Path))

                try:
                    OOXML = self.ooxml.run(Path)
                    OOXML["handle"] = file
                    self.OOXML_result.append(OOXML)
                except:
                    pass
                progressBar.progress(fileCount)

            titles = self.getTitles(self.OOXML_result)
            self.addData(titles, self.OOXML_result, "OOXML", skCase)
        
        elif any(x in extension.lower() for x in ["doc", "ppt", "xls"]):
            for file in files:
                self.log(Level.INFO, "Processing file: " + file.getName())
                fileCount += 1
                self.totalCount += 1

                Path = os.path.join(Directory, unicode(file.getName()))
                ContentUtils.writeToFile(file, File(Path))

                try:
                    resultCFBF = self.cfbf.run(Path)
                    resultCFBF["handle"] = file
                    self.CFBF_result.append(resultCFBF)
                except:
                    pass
                progressBar.progress(fileCount)

            titles = self.getTitles(self.CFBF_result)
            self.addData(titles, self.CFBF_result, "CFBF", skCase)


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