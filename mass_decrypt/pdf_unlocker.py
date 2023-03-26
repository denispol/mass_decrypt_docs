"""
This module contains the PDFUnlocker class for decrypting PDF files
"""

import os
import logging
import pikepdf
import tempfile
import shutil
from mass_decrypt.utils import setup_logger, set_file_times

# Exception imports
from pikepdf import PasswordError, PdfError


class PDFUnlocker:
    def __init__(self, logger, integrity_check=False):
        """
        Initialize a PDFUnlocker instance
        
        Args:
            logger (logging.Logger): Logger for logging messages
        """
        self.logger = setup_logger("decrypt_logger")
        self.integrity_check = integrity_check

    def unlock(self, filename, passwd):
        """
        Attempt to decrypt a PDF file with the given password
        
        Args:
            filename (str): The path of the PDF file
            passwd (str): The password to attempt decryption with
            
        Returns:
            result (str): A string describing the outcome of the decryption attempt
        """
        try:
            pdf_test = pikepdf.open(filename)
            result = "unencrypted"
        except pikepdf.PasswordError as e:
            self.logger.warning(f"File is encrypted: {filename}, {e}")
        except Exception as e:
            self.logger.error(f"Error: {filename}, {e}")
            result = "other_error"

            try:
                pdf = pikepdf.open(
                    filename, password=passwd, allow_overwriting_input=True
                )

                if self.integrity_check:
                    integrity_check_result = pdf.check()
                    if integrity_check_result:
                        self.logger.warning(
                            f"Integrity check failed: {filename}, {integrity_check_result}"
                        )
                        return "integrity_check_failed"

                original_metadata = os.stat(filename)
                self.logger.info(f"Unlocked: {filename}")
                pdf.save(filename)

                set_file_times(
                    filename,
                    original_metadata.st_atime,
                    original_metadata.st_mtime,
                    original_metadata.st_ctime,
                )
                result = "decrypted"
            except pikepdf.PasswordError as e:
                self.logger.warning(f"Invalid password error: {filename}, {e}")
                result = "invalid_password_error"
            except pikepdf.PdfError as e:
                self.logger.warning(f"PDF error: {filename}, {e}")
                result = "pdf_error"
            except Exception as e:
                self.logger.error(f"Error: {filename}, {e}")
                result = "other_error"

        return result
