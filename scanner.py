"""
Title: Search files and folders for username, password combinations.
    Output results to a file in format username:password:file_found
Author: Primus27
"""

# Import packages
import os
import argparse
import re
from pathlib import Path
from typing import Dict, List, Optional
from dataclasses import dataclass
from datetime import datetime

import ezodf
import openpyxl
import pandas as pd
from docx import Document

# Current program version
current_version = 2

@dataclass
class SearchParameters:
    username_pattern: Optional[re.Pattern]
    password_pattern: Optional[re.Pattern]
    other_keyword_pattern: Optional[re.Pattern]

TODAY = datetime.now()


def get_lines(file_path):
    ext = os.path.splitext(file_path)[1].lower()

    lines = []
    try:
        if ext == '.docx':
            doc = Document(file_path)
            lines = [para.text for para in doc.paragraphs if para.text.strip()]
        elif ext == '.odt':
            doc = ezodf.opendoc(file_path)
            lines = []
            for elem in doc.body:
                text = elem.text
                text and lines.append(text.strip())
        elif ext == '.xlsx':
            wb = openpyxl.load_workbook(file_path, read_only=True)
            for sheet in wb:
                for row in sheet.iter_rows(values_only=True):
                    for cell in row:
                        if cell is not None and str(cell).strip():
                            lines.append(str(cell))
            wb.close()
        elif ext == '.ods':
            doc = ezodf.opendoc(file_path)
            lines = []
            for sheet in doc.sheets:
                for row in sheet:
                    for cell in row:
                        if cell.value is not None:
                            value = str(cell.value).strip()
                            if value:
                                lines.append(value)
        else:
            try:
                with open(file_path, "r") as f:
                    lines = [line.strip() for line in f if line.strip()]
            except Exception as e:
                print(f"[*] Unsupported file type: {ext} ({e})")
    # Inadequate permission to access location
    except PermissionError:
        print("[*] Inadequate permissions to access file location")
    # Could not find path or invalid file format
    except OSError:
        print("[*] Unable to find path or invalid file: {}".format(file_path))
    except Exception as e:
        print(f"[*] Unknown error: {e}")

    return lines


def scan_file(file, search_parameters: SearchParameters):
    """
    Opens file and scans for keywords (username, password, etc). Supports TXT, DOCX, ODT, XLSX, ODS.
    :param file: The path of the file to be scanned.
    :param credentials: A list with filename, username, password credentials dicts.
    :param search_parameters: Search parameters to use.
    :return: An updated dictionary inc. the contents from the scanned file.
    """
    credentials = []

    try:
        lines = get_lines(file)

        for i, line in enumerate(lines, start=1):
            username = None
            password = None
            other_keyword = None

            if search_parameters.username_pattern:
                # Line contains any username synonym
                username_search_res = search_parameters.username_pattern.search(line)
                if username_search_res:
                    username = line

            if search_parameters.password_pattern:
                # Line contains any password synonym
                password_search_res = search_parameters.password_pattern.search(line)
                if password_search_res:
                    password = line

            # User had added search terms
            if search_parameters.other_keyword_pattern:
                other_search_res = search_parameters.other_keyword_pattern.search(line)
                if other_search_res:
                    other_keyword = other_search_res.group()

            if any([username, password, other_keyword]):
                credentials.append({
                    'filename': file,
                    'line': i,
                    'username': username,
                    'password': password,
                    'other_keyword': other_keyword,
                })

        return credentials
    except Exception as e:
        print("[*] Error processing file {}: {}".format(file, str(e)))


def enum_files(folder_path):
    """
    Enumerates files in a path.
    :param folder_path: Root folder path for enumeration.
    :return: List containing all files in a folder.
    """
    # List of all files in path
    f_list = []
    # Enumerate files
    for root, dirs, files in os.walk(folder_path, topdown=True):
        for file in files:
            # Generate the absolute/relative path for each file
            file_path = os.path.join(root, file)
            # File exists (it should) and has a size greater than 0KB
            if os.path.exists(file_path) and os.path.getsize(file_path) > 0:
                f_list.append(file_path)
    return f_list


def file_output(file_output_name, credentials):
    """
    Outputs list results to a file
    :param file_output_name: The output file name.
    :param credentials: List containing username, passwords...
    """
    try:
        df = pd.DataFrame(credentials)
        df.to_excel(file_output_name, index=False)
    except Exception as e:
        print(f"[*] Error result saving to {file_output_name} - {e}")


def read_aliases_list_file(file_path):
    try:
        with open(file_path, "r") as f:
            return f.read().split()
    except Exception as e:
        print(f"Error reading file {file_path}: {e}")
        return []


def main(scan_paths: List[Path], search_params: SearchParameters, output_filename: str):
    """
    Main method. Runs credential scanner
    """
    credentials = []

    for scan_path in scan_paths:
        # Enumerate files in path
        file_list = enum_files(scan_path)

        # Scan each file and add update the credentials dictionary
        for filename in file_list:
            credentials.extend(scan_file(filename, search_params))

    # Print results title
    print(f"{len(credentials)} results found:\n")

    # Output results to terminal
    for item in credentials:
        print(item['filename'], item['line'], item['username'], item['password'], item['other_keyword'])

    # Output results to file
    if output_filename:
        file_output(output_filename, credentials)


if __name__ == '__main__':
    # Define argument parser
    parser = argparse.ArgumentParser()
    # Remove existing action groups
    parser._action_groups.pop()

    # Create a required and optional group
    required = parser.add_argument_group("required arguments")
    optional = parser.add_argument_group("optional arguments")

    # Define arguments
    required.add_argument("-s", "--scanpath", nargs="*",
                          action="store", default=[], dest="scan_paths",
                          help="Scan paths (absolute or relative)",
                          required=True)
    optional.add_argument("-n", "--filename", action="store",
                          default=f'results_{TODAY.strftime("%Y%m%d_%H%M%S")}.xlsx',
                          dest="file_output_name",
                          help="Declare a custom filename for file output")
    optional.add_argument("-u", "--username", nargs="*", action="store",
                          default=[], dest="user_syn",
                          help="Add additional username aliases")
    optional.add_argument("-uf", "--usersfile", action="store", dest="user_syn_file",
                          help="Add additional username aliases list on file")
    optional.add_argument("-p", "--password", nargs="*", action="store",
                          default=[], dest="pass_syn",
                          help="Add additional password aliases")
    optional.add_argument("-pf", "--passfile", action="store", dest="pass_syn_file",
                          help="Add additional password aliases list on file")
    optional.add_argument("-l", "--advanced", nargs="*", action="store",
                          default=[], dest="search_list",
                          help="Add additional search terms")
    optional.add_argument("-lf", "--advancedfile", action="store",
                          dest="search_list_file",
                          help="Add additional search terms list file")
    optional.add_argument("--version", action="version",
                          version="%(prog)s {v}".format(v=current_version),
                          help="Display program version")
    args = parser.parse_args()

    # Synonyms of username
    user_syn = ["user", "username", "login", "email", "email address", "id"]
    user_syn.extend(args.user_syn)

    if args.user_syn_file:
        user_syn.extend(read_aliases_list_file(args.user_syn_file))

    username_pattern = re.compile('|'.join(user_syn), flags=re.I)

    # Synonyms of password
    pass_syn = ["pass", "password", "key", "secret", "pin", "passcode", "token"]
    pass_syn.extend(args.pass_syn)

    if args.pass_syn_file:
        pass_syn.extend(read_aliases_list_file(args.pass_syn_file))

    pass_pattern = re.compile('|'.join(pass_syn), flags=re.I)

    # Folder path to scan (Relative (current) or Absolute)
    scan_paths = [Path(p) for p in args.scan_paths]

    # Additional search terms
    search_list = []
    search_list.extend(args.search_list)

    if args.search_list_file:
        search_list.extend(read_aliases_list_file(args.search_list_file))

    if search_list:
        other_kw_pattern = re.compile('|'.join(search_list), flags=re.I)
    else:
        other_kw_pattern = None

    search_params = SearchParameters(
        username_pattern=username_pattern,
        password_pattern=pass_pattern,
        other_keyword_pattern=other_kw_pattern,
    )

    # Run main method
    main(
        search_params=search_params,
        output_filename=args.file_output_name,
        scan_paths=scan_paths,
    )
