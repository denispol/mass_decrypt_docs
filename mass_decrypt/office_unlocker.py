"""
This module contains the OfficeUnlocker class for decrypting Office files.
"""


import logging
import msoffcrypto
import tempfile
import shutil
import os
from mass_decrypt.utils import setup_logger, set_file_times

# Exception imports
from msoffcrypto.exceptions import (
    FileFormatError,
    ParseError,
    DecryptionError,
    InvalidKeyError,
)


class OfficeUnlocker:
    def __init__(self, logger):
        """
        Initialize an OfficeUnlocker instance
        
        Args:
            logger (logging.Logger): Logger for logging messages
        """
        self.logger = setup_logger("decrypt_logger")

    def unlock(self, filename, passwd):
        """
        Attempt to decrypt an Office file with the given password
        
        Args:
            filename (str): The path of the Office file
            passwd (str): The password to attempt decryption with
            
        Returns:
            result (str): A string describing the outcome of the decryption attempt
        """
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

                    set_file_times(
                        filename,
                        original_metadata.st_atime,
                        original_metadata.st_mtime,
                        original_metadata.st_ctime,
                    )

        return result
