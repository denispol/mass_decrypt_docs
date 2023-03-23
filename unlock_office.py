import os
import pathlib
import msoffcrypto
import tempfile
import shutil
import logging
import argparse
from tqdm import tqdm
from ctypes import windll, wintypes, byref

# Exception imports
from msoffcrypto.exceptions import (
    FileFormatError,
    ParseError,
    DecryptionError,
    InvalidKeyError,
)


class OfficeUnlocker:
    def __init__(self, loggername):
        self.logger = logging.getLogger(loggername)

    @staticmethod
    def to_windows_filetime(epoch):
        timestamp = int((epoch * 10000000) + 116444736000000000)
        return wintypes.FILETIME(timestamp & 0xFFFFFFFF, timestamp >> 32)

    def set_file_times(self, filename, atime, mtime, ctime):
        atime_struct, mtime_struct, ctime_struct = (
            self.to_windows_filetime(atime),
            self.to_windows_filetime(mtime),
            self.to_windows_filetime(ctime),
        )

        handle = windll.kernel32.CreateFileW(filename, 256, 0, None, 3, 128, None)
        windll.kernel32.SetFileTime(
            handle, byref(ctime_struct), byref(atime_struct), byref(mtime_struct)
        )
        windll.kernel32.CloseHandle(handle)

    def unlock(self, filename, passwd):
        with open(filename, "rb") as temp:
            try:
                office_file = msoffcrypto.OfficeFile(temp)
                if not office_file.is_encrypted():
                    self.logger.info(f"Not encrypted: {filename}")
                    return "unencrypted"

                office_file.load_key(password=passwd)
                result = "decrypted"
            except FileFormatError as e:
                self.logger.warning(f"Unsupported format: {filename}, {e}")
                result = "unsupported_format"
            except ParseError as e:
                self.logger.warning(f"Parse error: {filename}, {e}")
                result = "parse_error"
            except DecryptionError as e:
                self.logger.warning(f"Decryption error: {filename}, {e}")
                result = "decryption_error"
            except Exception as e:
                self.logger.error(f"Error: {filename}, {e}")
                result = "other_error"

            if result == "decrypted":
                self.logger.info(f"Unlocked: {filename}")

                with tempfile.TemporaryFile() as fout:
                    try:
                        office_file.decrypt(fout)
                    except InvalidKeyError as e:
                        self.logger.warning(f"Invalid key error: {filename}, {e}")
                        return "invalid_key_error"

                    fout.seek(0)
                    original_metadata = os.stat(filename)
                    shutil.copyfileobj(fout, open(filename, "wb"))

                    self.set_file_times(
                        filename,
                        original_metadata.st_atime,
                        original_metadata.st_mtime,
                        original_metadata.st_ctime,
                    )

        return result


def main():
    parser = argparse.ArgumentParser(description="Unlock Office files.")
    parser.add_argument("path", help="Path to the folder containing Office files.")
    parser.add_argument(
        "-R", "--recursive", action="store_true", help="Recursively process subfolders."
    )
    parser.add_argument("-p", "--password", help="Password for decryption.")
    parser.add_argument(
        "--plist", help="Path to a .txt file containing a list of passwords."
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
    
    # Logging to runtine.log in user's cwd (in console)
    logging.basicConfig(
        filename="runtime.log",
        level=logging.INFO,
        format="%(asctime)s - %(levelname)s - %(message)s",
    )
    logger_name = "office_unlocker"
    office_unlocker = OfficeUnlocker(logger_name)

    decrypted_count = 0
    unencrypted_count = 0
    error_count = 0

    extensions = (".doc", ".docx", ".docm", ".xls", ".xlsx", ".xlsm", ".xlsb")

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
            for password in passwords:
                result = office_unlocker.unlock(filepath, password)
                if result == "decrypted":
                    decrypted_count += 1
                    break
                elif result == "unencrypted":
                    unencrypted_count += 1
                    break
                else:
                    error_count += 1
                    break
            progress_bar.update(1)

    print(f"\nFiles decrypted: {decrypted_count}")
    print(f"Files unencrypted: {unencrypted_count}")
    print(f"Files not read (errors): {error_count}")


if __name__ == "__main__":
    main()
