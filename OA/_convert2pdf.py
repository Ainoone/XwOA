# -*- coding: utf-8 -*-

import gc
import logging
import os
from contextlib import contextmanager
from functools import wraps
from pathlib import Path

from win32com import client

# Set up basic configuration for logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')


@contextmanager
def open_word_application():
    """
    Context manager for opening Word applications
    """
    word = None
    try:
        word = client.DispatchEx("Word.Application")
        word.Visible = False
        word.AutomationSecurity = 3  # Disable macros
        yield word
    except Exception as e:
        logging.error(f"An error occurred while working with Word.Application: {e}")
    finally:
        if word is not None:
            try:
                word.Quit()
            except Exception as e:
                logging.error(f"An error occurred while trying to quit Word.Application: {e}")
            finally:
                # Collect garbage to make sure Word.Application properly shuts down
                word = None
                gc.collect()


def convert_to_pdf(func):
    """
    Convert Word documents to PDF.

    :param func: A function that returns a path to the directory containing Word files.
    :return: The path to the directory containing the converted PDF files.
    """

    @wraps(func)
    def wrapper(*args, **kwargs):
        # Get the file path
        path = func(*args, **kwargs)

        # Check if path is None
        if path is None:
            logging.error('The file path is None. Function was expected to return a valid path.')
            return

        # Check if path is an instance of a Path-like object
        if not isinstance(path, (str, os.PathLike)):
            logging.error('Invalid file path type. Expected a string or a Path-like object.')
            return

        # Try to create a Path object
        try:
            path = Path(path)
        except Exception as e:
            logging.error('Failed to construct a Path object from the given file path. Error: {}'.format(e))
            return

        # Check that the path is a directory
        if not path.is_dir():
            logging.error('Invalid file path: {}. The path does not point to a directory.'.format(path))
            return

        # Get the list of files to convert
        files = (file for file in path.iterdir()
                 if not file.name.startswith('~$') and
                 file.suffix.lower() in ('.docx', '.doc'))

        with open_word_application() as word:
            for file in files:
                try:
                    # Open Word file
                    doc = word.Documents.Open(file.as_posix(), ReadOnly=True)
                    # Convert the file to PDF and save it as a new file
                    pdf_name = file.with_suffix('.pdf')
                    doc.ExportAsFixedFormat(str(pdf_name), ExportFormat=17, Item=7, CreateBookmarks=1)
                    doc.Close()
                    logging.info(f'Converted {file} to {pdf_name}')
                except Exception as error:
                    logging.error(f'Failed to convert {file} to PDF: {error}')
        # Return the file directory
        return path

    return wrapper
