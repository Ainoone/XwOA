# -*- coding: utf-8 -*-

from pathlib import Path
from typing import Optional

# Define base directories
BASE_DIR = Path(__file__).parent.resolve()  # The directory containing this script.
TEMPLATE_DIR = BASE_DIR / 'Template'  # The directory for storing templates.


def search_template_file(name: str, suffix: str = 'docx', path: str | None = None) -> Optional[Path]:
    """
    Searches for a template file with a specific name and
     suffix within a specified directory or the default template directory.

    Parameters:
    - name (str): The base name of the file to search for.
    - suffix (str, optional): The file suffix (extension) to look for. Defaults to 'docx'.
    - path (Optional[str], optional): The directory to search within.
    If not specified, searches within the default template directory.

    Returns:
    - Optional[Path]: The path to the found file, or None if no matching file is found.
    """
    # Determine the directory to search in: either the provided path or the default template directory.
    search_path = Path(path) if path else TEMPLATE_DIR
    # Use rglob to recursively search for files matching the specified name and suffix.
    template_files = search_path.rglob(f"{name}.{suffix}")

    # Attempt to return the first matching file, if any. Otherwise, return None.
    return next(template_files, None)
