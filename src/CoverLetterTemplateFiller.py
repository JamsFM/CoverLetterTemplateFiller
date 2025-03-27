"""Script to Fill a CoverLetter Template"""
from sys import exit
from configparser import ConfigParser
from logging import config, getLogger
from json import load
from shutil import copy2
from docx import Document
from docx.shared import Pt

# Read app folder argument
app_root = "E:/Dev/Repos/CoverLetterTemplateFiller"

# Read Logging config
with open(f'{app_root}/config/dev/logging.json', encoding="UTF-8") as file:
    config_dict = load(file)
    config.dictConfig(config_dict)
logger = getLogger(__name__)


def docAsStringLogger(paragraphs):
    docAsString = ""
    for paragraph in paragraphs:
        docAsString += f"~{paragraph.text}\n"
    logger.info(f"\n{docAsString}")

    logger.info(f'End of File')


def validateParagraphSel(upperInitRef, lowerInitRef, upperCurrRef, lowerCurrRef):
    """Use XOR to validate whether the correct paragraphs are being transformed/filled during a given run"""
    upperInitRefValid = upperInitRef.__contains__(upperCurrRef)
    lowerInitRefValid = lowerInitRef.__contains__(lowerCurrRef)
    errMsg = "Failed to Select the correct Paragraph.\n\tBegan with-"
    if not upperInitRefValid:
        raise Exception(f'{errMsg}\t\t\"{upperInitRef[:90]}\"\n\tInstead of-\t\t\"{upperCurrRef}\"')
    if not lowerInitRefValid:
        raise Exception(f'{errMsg}\t\t\"{lowerInitRef[:90]}\"\n\tInstead of-\t\t\"{lowerCurrRef}\"')


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
    elif healthcareInput.__eq__("n"):
        healthcareFlag = False
    elif healthcareInput.__eq__(""):
        healthcareFlag = False
    else:
        raise Exception(f'Failed to receive an input matching [y/n]-\tReceived:\"{healthcareInput}\"')
    logger.info(f'Successfully set Healthcare to: [{healthcareInput}]')

    return inputRole, inputCompany, healthcareFlag


def validateParagraphFill(upperCurrRef, lowerCurrRef, upperInitRef, lowerInitRef):
    """Validate whether the paragraphs were filled with the data"""
    errMsg = "Failed to fill the following paragraph with the correct data.\n"
    if upperCurrRef.__eq__(upperInitRef):
        raise Exception(f'{errMsg}\"{upperCurrRef}\"')
    if lowerCurrRef.__eq__(lowerInitRef):
        raise Exception(f'{errMsg}\"{lowerCurrRef}\"')


def styleParagraph(paragraphs, index, docStyle, docFontName, docFontSize):
    """Style a paragraph with a given style, font name, & font size"""
    paragraphs[index].style = docStyle
    paragraphs[index].style.font.name = docFontName
    paragraphs[index].style.font.size = docFontSize


def templateFiller(srcPath, destPath):
    """Fill/Transform the data fields in the Cover Letter Template given Input-Params"""
    doc = Document(srcPath)
    # Set Author & Category of Doc  ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    core_props = doc.core_properties
    core_props.author = "James F. Mare"
    logger.info(f'Successfully set Author of File: [{core_props.author}]')
    core_props.category = "Cover Letter"
    logger.info(f'Successfully set Category of File: [{core_props.category}]')
    logger.info(f'Successfully Filled/Transformed Cover Letter Template File')

    #docAsStringLogger(doc.paragraphs)  #  For Debugging  ##############################################################

    # Preparing to transform data in two paragraphs  ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    upperIdx, lowerIdx = 4, 8
    paragraphs = doc.paragraphs
    upperInitRef, lowerInitRef = paragraphs[upperIdx].text, paragraphs[lowerIdx].text
    docStyle, docFontName, docFontSize = paragraphs[upperIdx].style, 'Calibri', Pt(12)
    upperSnippetRef = "I am excited to apply for the [Role] position at [Company]. With a strong foundation in software"
    lowerSnippetRef = "I am eager for the opportunity to bring my problem-solving abilities and technical expertise to"
    # Run the Validator for Paragraph Selection  <><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><>
    validateParagraphSel(upperInitRef, lowerInitRef, upperSnippetRef, lowerSnippetRef)
    # Health Care Snippet for lower paragraph
    healthcareSnippet = "am sure my experience coming from the healthcare world could be of some use, and ease my potential transition. I, not only look forward to continuing to make a difference in this world of uncertainty that we find ourselves living in, but I also "

    # Fetch Data from User Input for the Template  ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    inputRole, inputCompany, healthcareFlag = fetchUserInput()

    # Fill/Transform the data in the upper paragraph  ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    logger.info(f'\nFilling the following paragraph:\n[{paragraphs[upperIdx].text}]')
    paragraphs[upperIdx].text = paragraphs[upperIdx].text.replace("[Role]", inputRole)
    logger.info(f'\nFilled the following paragraph:\n[{paragraphs[upperIdx].text}]')

    # Fill/Transform the data in the upper paragraph  ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    paragraphs[upperIdx].text = paragraphs[upperIdx].text.replace("[Company]", inputCompany)
    paragraphs[lowerIdx].text = paragraphs[lowerIdx].text.replace("[Company]", inputCompany)

    if healthcareFlag:
        paragraphs[lowerIdx].text = paragraphs[lowerIdx].text.replace(healthcareSnippet, "")

    # Validate paragraphs  ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    validateParagraphFill(paragraphs[upperIdx].text, paragraphs[lowerIdx].text, upperInitRef, lowerInitRef)

    # Stylize the upper & lower paragraphs  ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    styleParagraph(paragraphs, upperIdx, docStyle, docFontName, docFontSize)
    styleParagraph(paragraphs, lowerIdx, docStyle, docFontName, docFontSize)

    #doc.save(destPath.replace("CL_Dev", "CL_Dev_Test"))  #  For Debugging  ############################################
    doc.save(destPath)
    logger.info(f'Successfully Filled Cover Letter Template File')
    return 0


def templator(config):
    """Create fresh copy of Cover Letter Template in /Out Dir for use in the templateFiller"""
    # Initialize Names for use in the copy method  ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    srcFileName = config.get("templator.properties", "SRC_FILE_NAME")
    destFileName = config.get("templator.properties", "DEST_FILE_NAME")
    # Initialize Paths for use in the copy method  ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    srcPath = f"{app_root}/In/{srcFileName}"
    destPath = f"{app_root}/Out/{destFileName}"
    try:
        copy2(srcPath, destPath)  # Commented out for testing purposes  <><><><><><><><><><><><><><><><><><><><><><><><>
        logger.info(f'Successfully Copied Cover Letter Template to destination path!\t[{destPath}]')
    except Exception as ex:
        logger.error(f'Following File could not be Copied!\t[{srcPath}]')
        logger.error(f'Exception:\n[{ex}]')
        raise

    # Run the Template Filler Next  <><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><>
    templateFiller(srcPath, destPath)

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
