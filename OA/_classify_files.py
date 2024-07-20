# -*- coding: utf-8 -*-

import logging
import os
import shutil
from functools import wraps
from pathlib import Path
from typing import List, Tuple, NoReturn

from PyPDF2 import PdfWriter
from more_itertools import only

# Set up basic configuration for logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')

# Predefined directory names for categorization
DIRECTORY_NAMES = ('PDF', 'Word', 'Excel')
MERGED_PDF_FOLDER_NAME = 'Merged_pdf'


def create_directories(src_dir_fd: Path, dst_dir_fds: Tuple[str]) -> List[Path]:
    """
    Create directories in the target directory based on the folder names specified in the source directory.
    :param src_dir_fd: Path to the source directory.
    :param dst_dir_fds: Tuple of folder names to create in the target directory.
    :return: List of paths to the created directories.
    """
    # Generate paths for each folder name in the target directory
    dst_paths = [src_dir_fd / folder for folder in dst_dir_fds]
    # Create directories for each folder name in the target directory
    for dst_path in dst_paths:
        if not dst_path.exists():
            dst_path.mkdir(parents=True)  # Ensure all intermediate directories are created
    return dst_paths


def get_dst_index_by_suffix(suffix: str) -> int:
    """
    Return the target index based on the given file suffix.
    :param suffix: File suffix.
    :return: Target index.
    This method compares the given suffix with specific file suffixes and returns the corresponding target index.
    Possible suffixes are '.pdf', '.docx', '.doc', '.xlsx', and '.xls'.
    If the suffix is '.pdf', the target index is set to 0.
    If the suffix is '.docx' or '.doc', the target index is set to 1.
    If the suffix is '.xlsx' or '.xls', the target index is set to 2.
    If the suffix does not match any predefined suffixes, the target index is set to None.
    """
    # Compare suffix and set the corresponding target index
    if suffix == '.pdf':
        dst_index = 0
    elif suffix in ('.docx', '.doc'):
        dst_index = 1
    elif suffix in ('.xlsx', '.xls'):
        dst_index = 2
    else:
        dst_index = None

    return dst_index


def move_file_to_dst(file: Path, dst_dir_fd: Path) -> None:
    """
    Move a file to the target directory.
    :param file: Path to the file to move.
    :param dst_dir_fd: Path to the target directory.
    """
    try:
        # Move file using shutil.move, setting copy_function to os.rename for renaming-based movement
        shutil.move(file, dst_dir_fd / file.name, copy_function=os.rename)
        # Log information if the file is successfully moved
        logging.info(f'{file} moved to {dst_dir_fd} folder')
    except Exception as e:
        # Log error if an exception occurs during the move
        logging.error(f'Error moving {file} to {dst_dir_fd.stem} folder: {e}')


def create_merged_pdf_dir(src_dir: Path) -> Path:
    """
    Create a directory for merged PDFs.
    :param src_dir: Path to the source directory.
    :return: Path to the merged PDF directory.
    """
    # Create the path for the merged PDF directory
    merged_pdf_dir = src_dir / MERGED_PDF_FOLDER_NAME
    merged_pdf_dir.mkdir(parents=True, exist_ok=True)
    return merged_pdf_dir


def merge_and_write_pdf_files(src_directory: Path) -> NoReturn:
    """
    Merge multiple PDF files in a directory and write the merged file.
    :param src_directory: Path to the directory containing PDF files to merge.
    :return: None
    """
    # Create the directory for merged PDFs
    merged_pdf_path = create_merged_pdf_dir(src_directory)
    # Create a PDF merger
    merger = PdfWriter()
    # Set the path for the target PDF
    target_pdf_path = merged_pdf_path / 'Merged_Pdf.pdf'

    try:
        # Iterate and add all PDF files
        for pdf_path in sorted(src_directory.rglob('*.pdf')):
            merger.append(pdf_path)
        # Write the merged PDF file
        merger.write(target_pdf_path)
        logging.info(f'Merged PDF file created at {target_pdf_path}')

    except Exception as error:
        logging.error(f'Error merging PDFs: {error}')
        raise
    finally:
        # Close the merger
        merger.close()


def handle_pdf_files(src_directory: str | Path) -> None:
    """
    Handle PDF files in the specified source directory.
    :param src_directory: String or Path object representing the source directory.
    :return: None
    """
    src_directory = Path(src_directory)  # Make sure src_directory is a Path object
    pdf_files = src_directory.rglob('*.pdf')  # Get all pdf files

    try:
        # Attempt to merge PDF files, do not merge if there's only one file
        _ = only(pdf_files, default=None, too_long=StopIteration)
    except StopIteration:
        # If there are multiple PDF files, proceed with merging
        merge_and_write_pdf_files(src_directory)
    else:
        logging.warning('No additional PDF files')


def remove_empty_dirs(dir_names: list[Path]) -> None:
    """
    Remove empty directories.
    :param dir_names: List of Path objects representing directories to check and remove if empty.
    :return: None
    """
    for path in dir_names:
        # Remove the directory if it is empty
        if not any(path.iterdir()):
            path.rmdir()
            logging.info(f'Empty folder {path} has been deleted')


def categorize_files(merge=False, dst: tuple = DIRECTORY_NAMES):
    """
    Decorator to categorize files based on their extension and optionally merge PDF files.
    :param merge: A boolean indicating whether to merge the PDF files or not. Defaults to False.
    :param dst: A tuple of directories where the files will be categorized. Defaults to DIRECTORY_NAMES.
    :return: None
    """

    def decorator(func):
        @wraps(func)
        def wrapper(*args, **kwargs):
            src = func(*args, **kwargs)
            if not src.is_dir():
                logging.error('Invalid directory dst_path')
                return

            # Create all necessary folders at the beginning
            dir_names = create_directories(src_dir_fd=src, dst_dir_fds=dst)

            for file in src.iterdir():
                if file.is_file():
                    suffix = file.suffix.lower()
                    dst_index = get_dst_index_by_suffix(suffix)
                    if dst_index is not None:
                        move_file_to_dst(file, dir_names[dst_index])

            logging.info('All files have been moved to the appropriate folder')

            remove_empty_dirs(dir_names=dir_names)

            if merge:
                handle_pdf_files(src)
            else:
                logging.warning('No more PDF files')

        return wrapper

    return decorator
