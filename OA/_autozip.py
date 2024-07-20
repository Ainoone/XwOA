# -*- coding: utf-8 -*-

from datetime import datetime
from functools import wraps
from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED


def auto_zip(func):
    """
    Decorator to automatically zip files in a directory returned by the decorated function.

    Args:
        func (callable): The function to decorate.

    Returns:
        callable: The wrapper function.
    """

    @wraps(func)
    def wrapper(*args, **kwargs):
        # Execute the decorated function and get its result (a directory path)
        path = func(*args, **kwargs)
        # Iterate through each file in the directory
        for file in path.iterdir():
            # Skip if the file is already a zip file
            if file.suffix == '.zip':
                continue
            # Zip the file
            zip_file(file)

    return wrapper


def zip_file(file_path):
    """
    Compresses a file into a ZIP archive, deletes the original file, and, if applicable,
    moves the ZIP file into a directory named after the year extracted from the file name.

    Args:
        file_path (Path): The path to the file to be compressed.
    """
    # Define the name of the ZIP archive
    archive = file_path.with_suffix('.zip')
    try:
        # Create and write to the ZIP file
        with ZipFile(archive, 'w', compression=ZIP_DEFLATED, compresslevel=6) as myzip:
            myzip.write(filename=file_path, arcname=file_path.name)
    except Exception as e:
        # Print an error message if an exception occurs
        print(f"Error occurred while zipping file {file_path}: {e}")
    else:
        # If zipping successful, delete original file
        file_path.unlink()

    # Check if file name is in date format, if yes, create a new folder and move the file
    try:
        year = datetime.strptime(file_path.stem, "%Y_%m").year
    except ValueError:
        # If parsing fails, return early
        return

    # Create a new directory with the year name if it doesn't exist
    new_directory = file_path.parent / str(year)
    new_directory.mkdir(exist_ok=True)

    # Move the zip file to the new directory
    archive.rename(Path(new_directory, archive.name))
