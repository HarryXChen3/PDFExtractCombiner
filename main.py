from __future__ import annotations

import os
if os.name != 'nt':
    raise RuntimeError("This application can only run on Windows_NT!")

import tempfile
import enum
import sys
from typing import Never
from pathlib import Path

import pdfkit
import pypdf
import pythoncom
import pandas as pd
import win32com.client as win32
from tqdm import tqdm

USE_WIN32API = True
WIN32_EXCEL_APPLICATION = "Excel.Application"
EXCEL_PDF_FILE_FORMAT = 57

SEARCH_FILES_RECURSIVELY = True
DISALLOW_ZERO_PAGE_PDF_MERGE = True

USE_EXCEL_ENGINE = "openpyxl"
PDF_EXTENSION = ".pdf"
EXCEL_EXTENSION = ".xlsx"

_TESTING_HTMLTOPDF_PATH = Path("wkhtmltox-0.12.6-1.mxe-cross-win64/wkhtmltox/bin/wkhtmltopdf.exe")
_BUILT_HTMLTOPDF_PATH = Path("./wkhtmltox/bin/wkhtmltopdf.exe")
if _BUILT_HTMLTOPDF_PATH.is_file():
    PDFKIT_CONFIG = pdfkit.configuration(wkhtmltopdf=str(_BUILT_HTMLTOPDF_PATH))
elif _TESTING_HTMLTOPDF_PATH.is_file():
    PDFKIT_CONFIG = pdfkit.configuration(wkhtmltopdf=str(_TESTING_HTMLTOPDF_PATH))
else:
    raise RuntimeError("Cannot find wkhtmltopdf.exe")

USE_ROOT_AS_WORKING_DIR = "use-root-as-working-dir"


def pdf_merge(output_path: str | Path, pdfs: dict[(str | Path), list[pypdf.PageRange]]):
    """
    Merge a map of pdfs to their respective PageRanges (or empty for all pages) into a single pdf

    :param output_path: output path that the new pdf should be written to
    :param pdfs: map of pdf paths (str | Path) to a list of PageRanges (or an empty list if all pages)
    :return: None
    """

    pdf_writer = pypdf.PdfWriter()
    for path, page_ranges in pdfs.items():
        if len(page_ranges) > 0:
            for page_range in page_ranges:
                pdf_writer.append(path, pages=page_range)
        else:
            pdf_writer.append(path)

    pdf_writer.write(output_path)
    pdf_writer.close()


def pdf_extract(pdf: str | Path, segments: list[tuple[int, int]] = None) -> list[pypdf.PageObject]:
    """
    Extract every tuple pair of *(start, end)* pdf page locations (inclusive)
    and then return a list of pdf pages

    Note that (start, end) are **0-based** page indexes

    :param pdf: reference to .pdf file
    :param segments: list of tuples describing the segments to extract [(start, end), (start, end)]
    :return: list of PageObjects extracted from the pdf
    """
    pdf_reader = pypdf.PdfReader(pdf)
    if segments is not None:
        extracted_pages = []
        for page_range in segments:
            for i in range(page_range[0], page_range[1] + 1):
                extracted_pages.append(pdf_reader.pages[i])

        return extracted_pages
    else:
        return pdf_reader.pages


def create_win_excel_instance(try_catch: bool = True):
    if try_catch:
        try:
            return win32.gencache.EnsureDispatch(WIN32_EXCEL_APPLICATION)
        except pythoncom.com_error:
            raise RuntimeError(f"Failed to dispatch {WIN32_EXCEL_APPLICATION}")
    else:
        return win32.gencache.EnsureDispatch(WIN32_EXCEL_APPLICATION)


def win_xlsx_to_pdf(
        xlsx: str | Path, page_ranges: list[pypdf.PageRange], use_excel_instance: win32.CDispatch = None
) -> str:
    excel_file = pd.ExcelFile(xlsx, engine=USE_EXCEL_ENGINE)
    n_sheets = len(excel_file.sheet_names)
    tmp_dir = tempfile.gettempdir()

    output_pdf_path = os.path.join(
        tmp_dir,
        f"{os.path.basename(str(xlsx))}_Extracted{PDF_EXTENSION}"
    )

    cleanup_file_paths = []
    excel_instance: win32.CDispatch | None = None
    workbook = None

    # forbidden code from http://www.icodeguru.com/WebServer/Python-Programming-on-Win32/ch12.htm
    # and https://pythonexcels.com/python/2009/10/05/python-excel-mini-cookbook
    try:
        if use_excel_instance is not None:
            excel_instance = use_excel_instance
        else:
            excel_instance = create_win_excel_instance(try_catch=False)

        excel_instance.Visible = False

        # silence existing file warnings (and other warnings) temporarily
        excel_instance.DisplayAlerts = False

        workbook = excel_instance.Workbooks.Open(str(xlsx))
        saved_worksheet_paths = []

        for page_range in page_ranges:
            for i in range(*page_range.indices(n_sheets)):
                # workbook.Worksheets are 1-indexed instead of 0-indexed, so we need to add 1 here to adjust for that
                # workbook.Worksheets(1) returns 'Sheet1' for a newly created Excel file
                # workbook.Worksheets(0) always errors
                worksheet = workbook.Worksheets(i + 1)
                worksheet.Activate()

                worksheet_file_path = os.path.join(
                    tmp_dir,
                    f"{os.path.basename(str(xlsx))}_Sheet_{worksheet.Name}{PDF_EXTENSION}"
                )

                worksheet.SaveAs(worksheet_file_path, FileFormat=EXCEL_PDF_FILE_FORMAT)

                saved_worksheet_paths.append(worksheet_file_path)
                cleanup_file_paths.append(worksheet_file_path)

        pdf_merge(output_pdf_path, {path: [] for path in saved_worksheet_paths})
    except pythoncom.com_error as com_error:
        hr, msg, exc, arg = com_error.hresult, com_error.strerror, com_error.excepinfo, com_error.argerror

        if exc is None:
            raise RuntimeError(f"\nExcel failed with code {hr}: {msg}")
        else:
            win_code, source, text, help_file, help_id, scode = exc
            raise RuntimeError(
                f"\nExcel failed with code {hr}: {msg}\n"
                f"Source: {source}\n"
                f"Message: {text}\n"
                f"More info: {help_file} (id={help_id})")
    finally:
        if workbook is not None:
            workbook.Close()

        if excel_instance is not None:
            excel_instance.DisplayAlerts = False

            # only quit excel instance if we aren't using an external reference
            if use_excel_instance is None:
                excel_instance.Quit()

    try:
        for path in cleanup_file_paths:
            os.remove(path)
    except OSError:
        # ignore if we failed to delete some files, they're all in tmp anyway
        pass
    finally:
        return output_pdf_path


def xlsx_to_pdf(xlsx: str | Path, page_ranges: list[pypdf.PageRange]) -> str:
    """
    Convert .xlsx Excel file to pdf

    :param xlsx: reference to .xlsx Excel file
    :param page_ranges: list of PageRange objects describing the sheets to extract
    :return: string filepath leading to the converted pdf (in a TMP/TEMP directory)
    """

    excel_file = pd.ExcelFile(xlsx, engine=USE_EXCEL_ENGINE)
    n_sheets = len(excel_file.sheet_names)

    extract_sheets = []
    for sheet_range in page_ranges:
        for i in range(*sheet_range.indices(n_sheets)):
            extract_sheets.append(i)

    extracted_sheet_names = [excel_file.sheet_names[i] for i in extract_sheets]

    data_frames = pd.read_excel(xlsx, sheet_name=extract_sheets)
    filtered_frames = {i: data_frames[i] for i in extract_sheets}
    combined_frame = pd.concat(filtered_frames.values())

    tmp_dir = tempfile.gettempdir()
    # os.path.basename does not work properly when running on POSIX system trying to get basename of a windows path
    # it shouldn't be an issue here as this path should always be of the currently running OS
    html_str = combined_frame.to_html()
    pdf_path_bytes = os.path.join(
        tmp_dir,
        f"{os.path.basename(str(xlsx))}_{'_'.join(extracted_sheet_names)}{PDF_EXTENSION}"
    )

    pdfkit.from_string(
        html_str,
        pdf_path_bytes,
        configuration=PDFKIT_CONFIG
    )

    return pdf_path_bytes


def merge_pdf_xlsx(
        output_path: str | Path,
        pdf: str | Path,
        xlsx: str | Path,
        pdf_pages: list[pypdf.PageRange],
        xlsx_sheets: list[pypdf.PageRange],
        use_excel_instance: win32.CDispatch = None):
    """
    Merge a .xlsx Excel file with a pdf file, specifying the PageRanges from each

    :param output_path: output path which the newly merged pdf should be written to
    :param xlsx: path to .xlsx Excel file
    :param pdf: path to .pdf file
    :param xlsx_sheets: list of PageRanges to extract from the .xlsx file
    :param pdf_pages: list of PageRanges to extract from the .pdf file
    :param use_excel_instance: use common excel instance instead of creating a new one,
                                this has no effect if USE_WIN32API is False
    :return: None
    """

    if USE_WIN32API:
        converted_pdf_path = win_xlsx_to_pdf(xlsx, xlsx_sheets, use_excel_instance=use_excel_instance)
    else:
        converted_pdf_path = xlsx_to_pdf(xlsx, xlsx_sheets)

    pdf_merge(output_path, {
        pdf: pdf_pages,
        converted_pdf_path: [],
    })


def get_files_with_ext(from_dir: Path = None, ext: str = PDF_EXTENSION, recursive: bool = False) -> list[Path]:
    """
    Returns all files with the specified extension from the specified directory

    :param from_dir: path of directory that files should be queried for under
    :param ext: extension to look for
    :param recursive: make the search by extension recursive from the directory
    :return: list of paths leading to matched files
    """
    at_dir = Path.cwd() if from_dir is None else from_dir
    return list(at_dir.rglob(f"*{ext}")) if recursive else list(at_dir.glob(f"*{ext}"))


class DuplicateFileNames(RuntimeError):
    def __init__(self, *args, **kwargs):
        super(*args, **kwargs)


def gather_xlsx_pdf_pairs(from_dir: Path = None) -> tuple[dict[Path, Path], set[str]]:
    root_dir = (Path.cwd() if from_dir is None else from_dir)
    found_pdfs = get_files_with_ext(root_dir, ext=PDF_EXTENSION, recursive=SEARCH_FILES_RECURSIVELY)
    found_xlsx = get_files_with_ext(root_dir, ext=EXCEL_EXTENSION, recursive=SEARCH_FILES_RECURSIVELY)

    mapped_pdfs = {path.stem: path for path in found_pdfs}
    mapped_xlsx = {path.stem: path for path in found_xlsx}

    if len(found_pdfs) != len(mapped_pdfs) or len(found_xlsx) != len(mapped_xlsx):
        raise DuplicateFileNames()

    pdf_names, xlsx_names = set(path.stem for path in found_pdfs), set(path.stem for path in found_xlsx)

    names_intersection = pdf_names.intersection(xlsx_names)
    names_disjoint = pdf_names.symmetric_difference(xlsx_names)

    file_path_pairs = {}
    for name in names_intersection:
        file_path_pairs[mapped_pdfs[name]] = mapped_xlsx[name]

    return file_path_pairs, names_disjoint


def query_yes_no(question: str, default: bool = False) -> bool:
    """
    Dead simple query yes/no utility

    :param question: question str to be asked
    :param default: default value if no response
    :return: True if yes, False if no
    """
    response = input(f"{question} [{'Y' if default else 'y'}/{'n' if default else 'N'}]: ").lower()
    return (True if response in ["yes", "ye", "y"] else False) if len(response) > 0 else default


def dir_empty(dir_path: Path):
    """
    Dead simple directory empty check utility

    :param dir_path: Path object referencing the directory
    :return: True if empty, False if not
    """
    if not dir_path.is_dir():
        raise ValueError("dir_path must be directory!")

    has_next = next(dir_path.iterdir(), None)
    return has_next is None


def check_output_dir_is_ok(intended_output_dir: Path) -> bool:
    """
    Check output directory doesn't exist and is empty, create it if it doesn't exist
    Ask if directory exists and is NOT empty, return True if ok, False if not ok

    :param intended_output_dir: output directory path, whether it exists or not
    :return: True if ok, False if not
    """
    try:
        intended_output_dir.mkdir(exist_ok=True)
    except OSError as os_error:
        raise RuntimeError(os_error)
    finally:
        if not dir_empty(intended_output_dir):
            not_empty_is_ok = query_yes_no(
                f"Directory {intended_output_dir} already exists and is NOT empty; Continue?")
            if not not_empty_is_ok:
                return False

        return True


def exit_if_not_ok(question: str, default: bool = None):
    """
    Invoke sys.exit(0) if answer to question is No (False)

    :param question: question to ask
    :param default: default value of question
    :return: None
    """
    if not query_yes_no(question=question, default=default):
        sys.exit(0)


def check_do_not_merge_zero(n_merging: int) -> Never | None:
    """
    Invokes sys.exit(0) if n <= 0

    :param n_merging: merging n files
    :return: Never | None
    """
    if DISALLOW_ZERO_PAGE_PDF_MERGE and n_merging <= 0:
        input(f"Cannot merge zero ({n_merging}) files together! Press [Enter] to exit...")
        sys.exit(0)


class Mode(enum.Enum):
    XLSX_PDF_COMBINE = ".xlsx & .pdf combine (monthly STR & RPM; STR binder report)"
    PDF_FIRST_PAGE_COMBINE = ".pdf combine (1st page; P&L First Pages)"


def mode_selector(modes: list[str] = None) -> Mode:
    computed_modes = modes if modes is not None else [mode.value for mode in Mode]
    n_modes = len(computed_modes)

    if n_modes <= 0:
        raise RuntimeError("unexpected 0-length computed_modes; expected len(computed_modes) > 0")

    print(f"Select a {Mode.__name__} (1-{n_modes}):")
    for i, mode in enumerate(computed_modes):
        print(f"\t[{i + 1}] {mode}")

    input_mode = input(f"Mode: ")
    try:
        mode_index = int(input_mode) - 1
        if mode_index < 0 or mode_index > (n_modes - 1):
            raise ValueError(f"selected_mode out of range; expected 1 <= selected_mode <= {n_modes}")
    except ValueError as value_err:
        print(f"You must enter a valid integer between 1 and {n_modes}, inclusive"
              f"\nExtra: {value_err}")

        return mode_selector(computed_modes)

    return Mode(computed_modes[mode_index])


def mode_xlsx_pdf_combine(selected_working_dir: Path, intended_output_dir: Path):
    try:
        xlsx_pdf_pairs, disjoint_files = gather_xlsx_pdf_pairs(from_dir=selected_working_dir)
    except DuplicateFileNames:
        # wait until enter to close
        input(
            f"Unexpected duplicate files exist {'recursively' if SEARCH_FILES_RECURSIVELY else 'directly'} under "
            f"{selected_working_dir}. Please remove/rename them to continue. Press [Enter] to exit...")
        sys.exit(0)

    n_xlsx_pdf_pairs = len(xlsx_pdf_pairs)

    print(f"Working Directory: {selected_working_dir}"
          f"\nIntended Output Directory: {intended_output_dir}"
          f"\nTMP Directory: {found_tmp_dir}"
          f"\nMatched xlsx & pdf pairs: {n_xlsx_pdf_pairs}"
          f"\nUnmatched lone files: {len(disjoint_files)}\n{', '.join(disjoint_files)}"
          f"\n")

    # check that we aren't attempting a 0 pdf merge when DISALLOW_ZERO_PAGE_PDF_MERGE is true
    check_do_not_merge_zero(n_xlsx_pdf_pairs)

    exit_if_not_ok("Is the above information correct?", default=True)

    if not check_output_dir_is_ok(intended_output_dir=output_dir):
        sys.exit(0)

    exit_if_not_ok("Start?", default=True)

    output_paths: list[Path] = []
    common_excel_instance = create_win_excel_instance()
    for pdf_path, xlsx_path in tqdm(xlsx_pdf_pairs.items()):
        output_path = Path(intended_output_dir, f"{pdf_path.stem}_Merged{PDF_EXTENSION}")
        merge_pdf_xlsx(
            output_path=output_path,
            pdf=pdf_path,
            xlsx=xlsx_path,
            pdf_pages=[pypdf.PageRange(":3")],
            xlsx_sheets=[pypdf.PageRange("3:5")],
            use_excel_instance=common_excel_instance
        )

        output_paths.append(output_path)

    # add option to merge all pdfs together
    exit_if_not_ok(f"\nMerge all newly combined .pdfs ({n_xlsx_pdf_pairs}) into a single pdf?", default=True)

    combined_pdf_path = intended_output_dir / f"Combined_1-{n_xlsx_pdf_pairs}{PDF_EXTENSION}"
    sorted_output_pdfs = sorted(output_paths)

    print(f"Attempting merge of {n_xlsx_pdf_pairs} .pdfs...")

    pdf_merge(
        combined_pdf_path, {path: [] for path in sorted_output_pdfs}
    )

    # wait until enter to close
    input(f"Wrote combined pdf to {combined_pdf_path}. Press [Enter] to exit...")

    common_excel_instance.Quit()


def mode_pdf_first_page_combine(selected_working_dir: Path, intended_output_dir: Path):
    found_pdfs = get_files_with_ext(selected_working_dir, ext=PDF_EXTENSION, recursive=SEARCH_FILES_RECURSIVELY)
    n_found_pdfs = len(found_pdfs)

    print(f"Working Directory: {selected_working_dir}"
          f"\nIntended Output Directory: {intended_output_dir}"
          f"\nTMP Directory: {found_tmp_dir}"
          f"\nMatched pdf files: {n_found_pdfs}"
          f"\n")

    # check that we aren't attempting a 0 pdf merge when DISALLOW_ZERO_PAGE_PDF_MERGE is true
    check_do_not_merge_zero(n_found_pdfs)

    exit_if_not_ok("Is the above information correct?", default=True)

    if not check_output_dir_is_ok(intended_output_dir=output_dir):
        sys.exit(0)

    exit_if_not_ok("Start?", default=True)

    print(f"Attempting merge of {n_found_pdfs} .pdfs...")

    combined_pdf_path = intended_output_dir / f"Combined_1-{n_found_pdfs}{PDF_EXTENSION}"
    pdf_merge(
        combined_pdf_path, {path: [pypdf.PageRange(":1")] for path in found_pdfs}
    )

    # wait until enter to close
    input(f"Wrote combined pdf to {combined_pdf_path}. Press [Enter] to exit...")


if __name__ == "__main__":
    # exclude filename from args here
    cmd_args = sys.argv[1:]

    found_tmp_dir = tempfile.gettempdir()
    if len(cmd_args) <= 0:
        working_dir = Path(Path.cwd().parents[0])
    elif cmd_args[0] == USE_ROOT_AS_WORKING_DIR:
        working_dir = Path.cwd()
    else:
        raise RuntimeError(f"Unexpected cmd line argument: {cmd_args[0]}")

    output_dir = Path(working_dir, "output")
    selected_mode = mode_selector()

    if selected_mode == Mode.XLSX_PDF_COMBINE:
        mode_xlsx_pdf_combine(selected_working_dir=working_dir, intended_output_dir=output_dir)
    elif selected_mode == Mode.PDF_FIRST_PAGE_COMBINE:
        mode_pdf_first_page_combine(selected_working_dir=working_dir, intended_output_dir=output_dir)
    else:
        raise RuntimeError(f"Selected_mode {selected_mode} invalid!")
