import subprocess
import pandas as pd
import logging as logger
from datetime import datetime
import os
import shutil
import re


# NOA GA Paths
NOA_GA_PDF_NAME = "NOA GA.pdf"
NOA_GA_TEMPLATE_PATH = "./static/templates/NOA_GA_CSV_Template.csv"
NOA_GA_CSV_PATH = "./input/NOA_GA/"

# Global Paths
PDF_PATH = "./static/pdfs/"
ARCHIVE_FOLDER = "./archive/"
LOG_FILE_NAME = datetime.now().strftime('./temp/logs/log_%m_%d_%Y_%H_%M_%S.log')

# Log file configuration
logger.basicConfig(filename=LOG_FILE_NAME, level=logger.INFO, 
                    format='%(asctime)s %(levelname)-4s %(message)s', 
                    datefmt='%m/%d/%Y %I:%M:%S %p: ')




def inspect_pdf(pdf):
    """Inspects the pdf and creates the inspected pdf in ./temp folder.

    Args:
        pdf (string): pdf file path

    Returns:
        Boolean: Returns True if successfully inspected else False
    """
    try:
        logger.info(f'Started inspect_pdf() function, Arg: {pdf}')
        logger.debug(f'Command: pdfforms inspect {pdf}')
        completed_process = subprocess.run(["pdfforms", "inspect", pdf],
                                            check= True, capture_output= True)
        if not completed_process.returncode:
            logger.info(f'Successful PDF inspection')
            return True
    except (subprocess.CalledProcessError, OSError, Exception) as e:
        logger.error(f"Error while inspecting the PDF: {e}")
    logger.critical(f'Failed PDF inspection!!!')
    return False




def fill_pdf(csv_data_file_name):
    """Fills the inspected pdf from the CSV file, renames the output file and moves to ./output folder.
    
    It deletes the ./fields.json as its content doesn't get refreshed on rerun. Renames the ouput pdf.

    Args:
        csv_data_file_name ([type]): [description]

    Returns:
        [type]: [description]
    """
    try:
        logger.info(f'Started fill_pdf() function, Arg: {csv_data_file_name}')
        logger.debug(f'Command: pdfforms fill {csv_data_file_name} --no-flatten')
        completed_process = subprocess.run(["pdfforms", "fill", csv_data_file_name, "--no-flatten"], 
                                            check= True, capture_output= True)
        if not completed_process.returncode:
            logger.info(f'Successful PDF Fill')
            logger.info(f'Now Deleting the fields.json file')
            os.remove('./fields.json')
            logger.info(f'Successfully Deleted fields.json file')
            logger.info("iterating ./filled/ directory to rename the output file")
            for subdir, dirs, files in os.walk("./filled/"):
                for filename in files:
                    rename_to = (filename.split(".")[0] + "_"+ csv_data_file_name.split("/")[-1]
                                ).replace(".csv", datetime.now().strftime('%s') + ".pdf")
                    os.rename(os.path.join(subdir, filename), os.path.join(subdir,rename_to))
                    logger.debug(f'renaming {os.path.join(subdir, filename)} to {os.path.join(subdir,rename_to)}')
                    shutil.move(os.path.join(subdir,rename_to), "./output/")
            return True
    except (subprocess.CalledProcessError, FileNotFoundError, Exception) as e:
        logger.error(f'Error while filling the PDF: {e}')
    logger.info(f'Failed PDF Fill Function!!!')
    return False




def copy_to_archive(filepath):
    """Copy files to archive folder.
    
    Copies input excel and the generated csv files
    
    Args:
        filepath (String): Filepath of the 
    """
    logger.info(f'Started copy_to_archive() with filename: {filepath}')
    shutil.copy(filepath, ARCHIVE_FOLDER)
    logger.info("completed copy_to_archive()")




def generate_NOA_GA_CSV(filepath):
    """Generates CSV file for NOA GA pdf

    Args:
        filepath (String): Excel filename
    """
    try:
        # Read Template
        logger.info("started generate_NOA_GA_CSV()")
        csv_template_df = pd.read_csv(NOA_GA_TEMPLATE_PATH, sep=",", 
                                    dtype={'PDF_FIELD_INDEX':int, 
                                            'EXCEL_FILE_COLUMN_NAME':str} )

        # get Excel data in DataFrame
        logger.debug(f'Reading excel file: {filepath}')
        excel_file_df = pd.read_excel(filepath)

        # Get the Data from Excel
        logger.debug(f'extracting excel data using column names from CSV template')
        data = [excel_file_df[excel_column_name][0] 
                    for excel_column_name 
                        in csv_template_df['EXCEL_FILE_COLUMN_NAME']]

        # The first column (which is the header) should match the first key in the ./fields.json file
        # which is the path of the inspected pdf file
        logger.debug("Generating Dataframe which then converted into CSV datafile")
        df = pd.DataFrame({(PDF_PATH + NOA_GA_PDF_NAME): csv_template_df['PDF_FIELD_INDEX'],
                        'Description': csv_template_df['EXCEL_FILE_COLUMN_NAME'],
                        '':data})

        # Split Name into First, Middle and Last Name
        logger.debug("Handling Edge cases: Name split into First, Middle and last name")
        df[''][0] = df[''][0].split(",")[0]
        df[''][1] = df[''][1].split(",")[1]
        df[''][2] = ""
        df[''][13] = df[''][13].split(",")[0]
        df[''][14] = df[''][14].split(",")[1]
        df[''][15] = ""

        # Generates CSV in ./input/NOA_GA/ with file name last_name.csv
        
        csv_data_file_name = NOA_GA_CSV_PATH + df[''][0].replace(" ","") + ".csv"
        logger.debug(f'')

        # If generate CSV is successful then call inspect pdf, 
        # if successdully inspected then call fill pdf, 
        # if successful then copy the generated csv to archive and delete  
        df.to_csv(csv_data_file_name, index=False)
        if inspect_pdf((PDF_PATH + NOA_GA_PDF_NAME)):
            if fill_pdf(csv_data_file_name):
                copy_to_archive(csv_data_file_name)
                os.remove(csv_data_file_name)
                return True
            else:
                logger.error("Failed in fill_pdf()!!!")
        else:
            logger.error("Failed in inspect_pdf()!!!")


    except Exception as e:
        logger.error(f'Some Error has Occured: {e}')

    return False


def read_input():
    """Recursively reads ./input directory and finds xls/xlsx files.

    Calls appropriate CSV file generation functions based on the 
    parent directory of the excel file.
    """
    try:
        logger.info(f'Started read_inpt()')
        for subdir, dirs, files in os.walk("./input/"):
            for filename in files:
                if filename.endswith(".xls") or filename.endswith(".xlsx") :
                    if subdir.endswith("NOA_GA"):
                        filepath = os.path.join(subdir, filename)
                        logger.debug(f'Calling generate_NOA_GA_csv() with file: "{filepath}" as argument')
                        if generate_NOA_GA_CSV(filepath):
                            copy_to_archive(filepath)
                            os.remove(filepath)
                        else:
                            raise Exception
        return True
    except Exception as e:
        logger.error(f'Some Error has Occured: {e}')
    return False
                

if read_input():
    logger.info("Successfully Completed!!!")
else:
    logger.error("Sorry!!!There were some errors while filling")

# inspect_pdf('NOA.pdf')
# fill_pdf('data_file.csv')