"""
This script is used for batch decryption of Office and PDF files in a specified directory
"""

import os
import pathlib
import argparse
import logging
from tqdm import tqdm
from mass_decrypt.office_unlocker import OfficeUnlocker
from mass_decrypt.pdf_unlocker import PDFUnlocker
from mass_decrypt.utils import setup_logger


def process_file(filepath, filename, office_unlocker, pdf_unlocker, passwords):
    """
    Process a single file and attempt to decrypt it using the provided passwords
    
    Args:
        filepath (str): The path of the file to process
        filename (str): The name of the file to process
        office_unlocker (OfficeUnlocker): An instance of the OfficeUnlocker class
        pdf_unlocker (PDFUnlocker): An instance of the PDFUnlocker class
        passwords (list): A list of passwords to try during decryption
        
    Returns:
        bool: True if decryption is successful or the file is unencrypted, False otherwise
    """
    result = None
    for password in passwords:
        if filename.lower().endswith(".pdf"):
            result = pdf_unlocker.unlock(filepath, password)
        else:
            result = office_unlocker.unlock(filepath, password)

        if result in ("decrypted", "unencrypted"):
            break

    return result


def main():
    """
    Main function for the mass_decrypt script. Processes command-line arguments, initializes
    the necessary objects, and iterates through the specified files, attempting decryption
    """
    parser = argparse.ArgumentParser(description="Unlock PDF and Office files.")
    parser.add_argument("path", help="Path to the folder containing Office files.")
    parser.add_argument(
        "-R", "--recursive", action="store_true", help="Recursively process subfolders."
    )
    parser.add_argument("-p", "--password", help="Password for decryption.")
    parser.add_argument(
        "--plist", help="Path to a .txt file containing a list of passwords."
    )
    parser.add_argument(
        "-P",
        "--pdf",
        action="store_true",
        help="Decrypt PDF files in the specified folders and subfolders.",
    )
    parser.add_argument(
        "-O",
        "--office",
        action="store_true",
        help="Decrypt Office files in the specified folders and subfolders.",
    )
    parser.add_argument(
        "--integrity-check",
        action="store_true",
        help="Perform an integrity check on decrypted PDF files.",
    )
    args = parser.parse_args()

    if not args.password and not args.plist:
        parser.error("Either --password or --plist must be provided")

    if args.password:
        passwords = [args.password]
    else:
        with open(args.plist, "r") as f:
            passwords = [line.strip() for line in f]

    path = pathlib.Path(args.path)
    if not path.is_dir():
        print(f"'{args.path}' is not a valid directory")
        return

    # Set up the logger and related filehandler
    logger_name = "decrypt_logger"
    logger = setup_logger(logger_name)

    office_unlocker = OfficeUnlocker(logger)
    pdf_unlocker = PDFUnlocker(logger, integrity_check=args.integrity_check)

    decrypted_count = 0
    unencrypted_count = 0
    error_count = 0

    extensions = ()
    if args.office:
        extensions += (".doc", ".docx", ".docm", ".xls", ".xlsx", ".xlsm", ".xlsb")
    if args.pdf:
        extensions += (".pdf",)

    if args.recursive:
        all_files = [
            (os.path.join(root, file), file)
            for root, _, files in os.walk(args.path)
            for file in files
            if file.lower().endswith(extensions)
        ]
    else:
        all_files = [
            (str(file), file.name)
            for file in path.glob("*")
            if file.name.lower().endswith(extensions)
        ]

    with tqdm(total=len(all_files), unit="file") as progress_bar:
        for filepath, filename in all_files:
            progress_bar.set_description(f"Processing: {filename}")
            result = process_file(
                filepath, filename, office_unlocker, pdf_unlocker, passwords
            )

            if result == "decrypted":
                decrypted_count += 1
            elif result == "unencrypted":
                unencrypted_count += 1
            else:
                error_count += 1

            progress_bar.update(1)

    print(f"\nFiles decrypted: {decrypted_count}")
    print(f"Files unencrypted: {unencrypted_count}")
    print(f"Files not read (errors or no matching password): {error_count}")


if __name__ == "__main__":
    main()
