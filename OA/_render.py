# -*- coding: utf-8 -*-

from datetime import date
from pathlib import Path
from typing import Any, Dict, Iterable, Union

from ._classify_files import categorize_files
from ._convert2pdf import convert_to_pdf
from ._docxtpl import docx_tpl_file


def convert_date(data: Dict[str, Any]) -> Dict[str, Any]:
    """
    Converts date objects in a dictionary to string format.

    Args:
        data: The input dictionary.

    Returns:
        The transformed dictionary with string-formatted dates.
    """
    if not isinstance(data, dict):
        raise ValueError("The input data must be of dictionary type")
    # Convert date objects to string format within the dictionary
    return {k: v.strftime('%Y年%m月%d日') if isinstance(v, date) else v for k, v in data.items()}


def safe_format(fmt: str, mapping: Dict[str, Any]) -> str:
    """
    Safely formats a string.

    Args:
        fmt: The format string.
        mapping: The mapping dictionary for formatting.

    Returns:
        The formatted string.
    """
    return fmt.format(**mapping)


def tax_fmt(mapping: Dict[str, Any]) -> str:
    """
    Formats a tax document filename.

    Args:
        mapping: The mapping dictionary for formatting.

    Returns:
        The formatted filename.
    """
    fmt = "{CN!s:.6}" "_{End.year}年" "{End.month:02}月" ".docx"
    return safe_format(fmt, mapping)


def default_fmt(mapping: Dict[str, Any]) -> str:
    """
    Formats a default document filename.

    Args:
        mapping: The mapping dictionary for formatting.

    Returns:
        The formatted filename.
    """
    fmt = "{CN!s:.6}" "_签字材料.docx"
    return safe_format(fmt, mapping)


@categorize_files(merge=True)
@convert_to_pdf
def render_docx(initial_data: Iterable[Dict[str, Any]], path: Union[str, Path], out_fd: Path, label: str):
    """
    Renders a DOCX template with given data and saves it to a specified path.

    Args:
        initial_data: Data for rendering the template.
        path: The template file's path.
        out_fd: The output directory.
        label: The worksheet label.
    """
    # Open the template document
    with docx_tpl_file(path) as docx:
        # Iterate over the initial data
        for mapping in initial_data:
            # Render the DOCX template with converted dates
            docx.render(convert_date(mapping))
            # Generate the filename based on the label
            filename = out_fd.joinpath(tax_fmt(mapping) if label == 'Tax' else default_fmt(mapping))
            # Save the rendered DOCX file
            docx.save(filename)
    # Return the output directory Path object
    return out_fd
