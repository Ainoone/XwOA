# -*- coding: utf-8 -*-

import contextlib
from pathlib import Path

from docxtpl import DocxTemplate


def get_filename_extension(path: Path):
    """
    Extracts the filename and its extension from a given path.

    Parameters:
    - path (Path): The path to the file.

    Returns:
    - tuple: A tuple containing the filename and its extension.
    """
    filename = path.name
    extension = path.suffix
    return filename, extension


@contextlib.contextmanager
def docx_tpl_file(path):
    """
    A context manager for handling a DocxTemplate file.

    Parameters:
    - path (Path): The path to the template file.

    Raises:
    - ValueError: If the file is not a .docx file.
    - FileNotFoundError: If the file does not exist.
    - RuntimeError: For other errors, preventing sensitive information leakage.

    Yields:
    - DocxTemplate: A DocxTemplate object for the file.
    """
    full_path = path.resolve()
    filename, ext = get_filename_extension(full_path)

    # Ensure the file is a .docx file
    if ext != '.docx':
        raise ValueError(f"Invalid file type: {filename}. Only .docx files are supported.")

    try:
        # Create a DocxTemplate object from the template file
        docx = DocxTemplate(full_path)
        yield docx
    except FileNotFoundError:
        raise FileNotFoundError(f"File not found: {full_path}")
    except Exception as error:
        # Error handling to prevent sensitive information leakage
        raise RuntimeError(f"Error occurred while processing the file: {str(error)}") from None
