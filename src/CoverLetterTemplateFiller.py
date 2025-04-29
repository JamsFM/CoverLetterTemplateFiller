"""Script to Fill a CoverLetter Template
Author: James F. Mare`
Date: 2025-04-29

To Run:
    shell/CLTemplateFiller_run.bat

Description:
    This script will fill a Cover Letter Template with user given input parameters
    The input parameters then replace the data fields within the template:
    [Role], [Company], [Healthcare]

    The script will then copy the filled template to the destination folder

    The logging location and config details can be set/found here:
        config/dev/logging.json
        config/dev/CLTemplateFiller.ini
"""
from sys import exit
from os import path, remove
import subprocess
from pathlib import PurePosixPath
from configparser import ConfigParser
from logging import config, getLogger
from json import load
from shutil import copy2
from docx import Document
from docx.shared import Pt
from re import search, sub
from time import ctime
from datetime import datetime

# Read app folder argument
app_root = "E:/Dev/Repos/CoverLetterTemplateFiller"

# Read Logging config
with open(f'{app_root}/config/dev/logging.json', encoding="UTF-8") as file:
    config_dict = load(file)
    config.dictConfig(config_dict)
logger = getLogger(__name__)


def docPropSetter(core_props):
    """Set the author & category of the document"""
    core_props.author = "James F. Mare"
    logger.info(f'Successfully set Author of File: [{core_props.author}]')
    core_props.category = "Cover Letter"
    logger.info(f'Successfully set Category of File: [{core_props.category}]')
    logger.info(f'Successfully Filled/Transformed Cover Letter Template File')


'''def docAsStringLogger(paragraphs):
    """Log the document as a string; for debugging purposes"""
    docAsString = ""
    for paragraph in paragraphs:
        docAsString += f"~{paragraph.text}\n"
    logger.info(f"\n{docAsString}")

    logger.info(f'End of File')'''


def fileCopier(srcPath, destPath):
    """Copy the file at the given source path to the given destination path"""
    try:
        copy2(srcPath, destPath)
        logger.info(f'Successfully Copied Cover Letter Template to destination path!\t[{destPath}]')
    except Exception as ex:
        logger.error(f'Following File could not be Copied!\t[{srcPath}]')
        logger.error(f'Exception:\n[{ex}]')
        raise


def fetchUserInput():
    """Fetch User Input for the Template"""
    # Set Role value for Template   ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    inputRole = ""
    inputRole = input("Enter Role:").strip()
    if not inputRole:
        inputRole = "Software Developer Engineer"
    logger.info(f'Successfully set Role to: [{inputRole}]')

    # Set Company value for Template   ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    inputCompany = ""
    inputCompany = input("Enter Company:").strip()
    if not inputCompany:
        inputCompany = "your Company"
    logger.info(f'Successfully set Company to: [{inputCompany}]')

    # Set whether to use Healthcare snippet   ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    healthcareInput = input("Is Healthcare [y/n]:").lower().strip()
    healthcareFlag = False
    if healthcareInput.__eq__("y"):
        healthcareFlag = True
    elif healthcareInput.__eq__("n") or healthcareInput.__eq__(""):
        healthcareFlag = False
    else:
        raise Exception(f'Failed to receive an input matching [y/n]-\tReceived:\"{healthcareInput}\"')
    logger.info(f'Successfully set Healthcare to: [{healthcareInput}]')

    return inputRole, inputCompany, healthcareFlag


# noinspection t
def tokenReplacer(docBody, substitutionDictionary):
    """Replace a token in a given paragraph with the given text"""
    for subKey in substitutionDictionary:
        logger.info(f'Beginning to replace: \"{subKey}\" with: \"{substitutionDictionary[subKey]}\"-')
        for paragraph in docBody:
            if subKey in paragraph.text:
                logger.info(f'Substitution found in paragraph:\n\"{paragraph.text[:100]}...\"')
                tokenizedRun = paragraph.runs
                for i in range(len(tokenizedRun)):
                    currToken = tokenizedRun[i].text
                    if subKey in currToken:
                        oldToken = currToken
                        tokenizedRun[i].text = currToken.replace(subKey, substitutionDictionary[subKey])
                        newToken = tokenizedRun[i].text
                        logger.info(f'Successfully set \"{oldToken}\" to \"{newToken}\"')


def validateDocFill(bodyCurrRef, bodyInitRef):
    """Validate whether the paragraphs were filled with the data"""
    if bodyCurrRef is bodyInitRef:
        raise Exception(f'Failed to fill the following with the correct data-\n\"{bodyInitRef}\"')
    logger.info(f'Successfully Validated that Paragraph was changed.')


def stylizeDoc(docBody, docStyle, docFontName, docFontSize):
    """Style a paragraph with a given style, font name, & font size
        cp /mnt/c/Windows/Fonts/calibri.ttf home/jamesfm/.fonts/"""
    map(lambda x: docStyle, map(lambda x: x.style, docBody))
    logger.info(f'Successfully set docStyle to: [{docStyle}]')

    map(lambda x: docFontName, map(lambda x: x.style.font.name, docBody))
    logger.info(f'Successfully set docFontName to: [{docFontName}]')

    map(lambda x: docFontSize, map(lambda x: x.style.font.size, docBody))
    logger.info(f'Successfully set docFontSize to: [{docFontSize}]')


# noinspection t
def templateFiller(srcPath, destPath):
    """Fill/Transform the data fields in the Cover Letter Template given Input-Params"""
    doc = Document(srcPath)
    # Set Author & Category of Doc  <><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><>
    docPropSetter(doc.core_properties)

    # Preparing to transform data in two paragraphs  ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    paragraphs, healthcareParagraphIdx, docBodyIndices = doc.paragraphs, 8, slice(4, 9, 1)
    # Create a subset of the document referring to the Main Body
    docBody = paragraphs[docBodyIndices]
    bodyInitRef = map(lambda x: x.text, docBody)
    # Gather the Style, Font Name, & Font Size details
    docStyle, docFontName, docFontSize = docBody[0].style, 'Calibri', Pt(12)

    # Fetch Data from User Input for the Template   <><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><>
    inputRole, inputCompany, healthcareFlag = fetchUserInput()
    # Health Care Snippet for last paragraph  ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    healthcareSnippet = ("I am sure my experience coming from the healthcare world could be of some use, and ease my "
                         "potential transition.")
    if not healthcareFlag:
        healthcareSnippet = ""

    # Fill/Transform the data in the doc body  ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    substitutionDictionary = {'[Role]': inputRole, '[Company]': inputCompany, '[Healthcare]': healthcareSnippet}
    tokenReplacer(docBody, substitutionDictionary)

    # Validate paragraphs   <><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><>
    validateDocFill(map(lambda x: x.text, docBody), bodyInitRef)

    # Stylize the upper & lower paragraphs  <><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><>
    stylizeDoc(docBody, docStyle, docFontName, docFontSize)

    doc.save(destPath)
    logger.info(f'Successfully Filled Cover Letter Template File')
    return 0


def wslMntr(winPath):
    """Takes a Windows Path and converts it to a WSL Mounted Path
        "{driveLetter}:/" -> /mnt/{driveLetter}/
    """
    winDrive = search("[A-Z]:/", winPath)[0]
    if winDrive:
        driveLetter = winDrive[0].lower()
        mountedDrive = f'/mnt/{driveLetter}/'
        return PurePosixPath(sub(winDrive, mountedDrive, winPath, 1))
    else:
        raise FileNotFoundError(f'Invalid Windows Path: [{winPath}]; Expected Format: [<DriveLetter>:/..]')


def docxToPDF(filePath, pdfPath):
    """Save the docx file as a pdf file"""
    """
        /usr/lib/libreoffice/program/soffice.bin --headless \
            "-env:UserInstallation=file:///tmp/LibreOffice_Conversion_jamesfm" \
            --convert-to pdf:writer_pdf_Export /mnt/e/Dev/Repos/CoverLetterTemplateFiller/Out/Mare_James_CL_Dev.docx \
            --outdir /mnt/e/Dev/Repos/CoverLetterTemplateFiller/Out
        convert /mnt/e/Dev/Repos/CoverLetterTemplateFiller/Out/Mare_James_CL_Dev.docx
            as a Writer document -> /mnt/e/Dev/Repos/CoverLetterTemplateFiller/Out/Mare_James_CL_Dev.pdf
            using filter : writer_pdf_Export
    """
    try:
        filePath = wslMntr(str(filePath))
        destDir = filePath.parent
        logger.info(f'Attempting to Convert Docx File @[{pdfPath}]\nto PDF using [{filePath}]')
        bash = ('wsl \
                /usr/lib/libreoffice/program/soffice.bin \
                --headless \
                "-env:UserInstallation=file:///tmp/LibreOffice_Conversion_${USER}" '
                f'--convert-to pdf:writer_pdf_Export {filePath} \
                --outdir {destDir}')
        results = subprocess.Popen(bash, shell=True, text=True)
        '''bash = ["wsl",
                "/usr/lib/libreoffice/program/soffice.bin",
                "--headless",
                "-env:UserInstallation=file:///tmp/LibreOffice_Conversion_${USER}",
                "--convert-to", "pdf:writer_pdf_Export", f"{filePath}",
                "--outdir", f"{destDir}"]
        results = subprocess.run(bash, text=True)'''
        results.wait(100.0)
        #results.wait()
        #sleep(5)
        results, error = results.communicate()
        logger.info(f'SubProcess Results:\n[{results}; {error}]')

        return pdfPath

    except Exception as ex:
        logger.error(f'Following File could not be Converted to PDF!\t[{filePath}]')
        raise ex


def reattemptDocxToPDF(docPath):
    """Reattempt to save the docx file as a pdf file for a maximum number of times"""
    try:
        pdfPlannedPath = str(docPath).replace(".docx", ".pdf")
        if path.exists(pdfPlannedPath):
            remove(pdfPlannedPath)
        pdfPath = docxToPDF(docPath, pdfPlannedPath)

        maxAttempts = 7
        for count in range(maxAttempts):
            if path.exists(pdfPath):
                currentTime, mTimeStamp = datetime.now().replace(microsecond=0), path.getmtime(pdfPath)
                mDateTime = datetime.fromtimestamp(mTimeStamp).replace(microsecond=0)
                recentlyModifiedFlag = currentTime == mDateTime
                if recentlyModifiedFlag:
                    logger.info(f'Successfully Converted Docx  File to PDF @ [{pdfPath}]')
                    return pdfPath
            else:
                # Run the docx To PDF Method  <><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><>
                pdfPath = docxToPDF(docPath, pdfPlannedPath)
                currentTime, mTimeStamp = datetime.now(), path.getmtime(pdfPath)

                if count == maxAttempts:
                    legibleTime = ctime(mTimeStamp)
                    raise Exception(f'Newly Generated PDF could not be found! [{pdfPath}]; Or is not new,\n'
                                    f'given its mTimeStamp: [{legibleTime}] and currentTime: [{currentTime}]')
    except Exception as ex:
        logger.error(f'Following File could not be Converted to PDF!\t[{docPath}]')
        raise ex


def templator(config):
    """Create fresh copy of Cover Letter Template in /Out Dir for use in the templateFiller"""
    # Initialize Names for use in the copy method  ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    docSrcFileName = config.get("templator.properties", "DOC_SRC_FILE_NAME")
    docDestFileName = config.get("templator.properties", "DOC_DEST_FILE_NAME")
    # Initialize Paths for use in the copy method  ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    docSrcPath = PurePosixPath(f"{app_root}/In/{docSrcFileName}")
    docDestPath = PurePosixPath(f"{app_root}/Out/{docDestFileName}")

    # Run the File Copier Method; Move PDF from /In Dir to /Out Dir for further processing  <><><><><><><><><><><><><><>
    fileCopier(docSrcPath, docDestPath)

    # Run the Template Filler Method  <><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><>
    templateFiller(docSrcPath, docDestPath)
    # Run the docx To PDF Method  <><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><>
    pdfFile = reattemptDocxToPDF(docDestPath)

    # Run the File Copier Method; Move PDF from /Out Dir to Final Destination   <><><><><><><><><><><><><><><><><><><><>
    pdfDestDir = config.get("templator.properties", "PDF_FINAL_DEST_DIR")
    pdfDestPath = f'{pdfDestDir}{docDestFileName.replace(".docx", ".pdf")}'
    fileCopier(pdfFile, pdfDestPath)

    logger.info(f'Successfully Filled Template')
    return 0


if __name__ == "__main__":
    logger.info((7*" <>")+" Begin the [Cover Letter Template Filler] Job"+(7*" <>"))
    config = ConfigParser()
    config.read(f'{app_root}/config/dev/CLTemplateFiller.ini')

    try:
        # Run Templater on Docx File  ==================================================================================
        templator(config)
        logger.info((4*"  +")+f"[Cover Letter Template Filler] Successfully Filled Template"+(4*"  +"))
        exit(0)

    except Exception as ex:
        logger.info((23*"x  "))
        logger.error(f'The following Exception was caught: \n{ex}')
        logger.error(23*"X  ")
        exit(1)
