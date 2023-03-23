import os
import pathlib
import msoffcrypto
import tempfile
import shutil
import logging
import argparse

def main(source_path):
    # Get the paths of all Excel files in the Data sources folder
    excel_files = pathlib.Path(source_path).glob("*.xls*")

    # Loop through the list of Excel files and call the 'unlock' method on each of them to remove the password
    for file in excel_files:
        logger.info(f"Processing: {file}")
        try:
            unlock(file, "***REMOVED***", "logger")
            logger.info(f"File {file} successfully unlocked, and original file was overwritten.")
        except Exception as e:
            logger.error(f"Could not unlock {file}. Exception: {e}")
            logger.warning(f"Could not unlock {file}. File may be already unprotected.")

if __name__ == "__main__":
    # Set up argument parsing
    parser = argparse.ArgumentParser(description="Unlock Excel files in the specified folder.")
    parser.add_argument("source_path", help="Path to the folder containing the Excel files.")
    args = parser.parse_args()

    # Set up logging
    logging.basicConfig(
        filename="runtime.log",
        level=logging.DEBUG,
        format="%(asctime)s - %(levelname)s - %(message)s",
    )
    logger = logging.getLogger("logger")

    # Call the main function with the parsed source_path
    main(args.source_path)

def unlock(filename, passwd, loggername):
    logger = logging.getLogger(loggername) 
    temp = open(filename, "rb")
    excel = msoffcrypto.OfficeFile(temp)
    excel.load_key(passwd)

    # Create a temporary file for the decrypted content
    with tempfile.NamedTemporaryFile(delete=False) as f:
        try:
            excel.decrypt(f)
        except Exception as e:
            logger.error(f"Error during decryption of {filename}: {e}")
            temp.close()
            os.unlink(f.name)  # Remove the temporary file if an error occurs
            return

    temp.close()

    # Preserve file access, modification, and creation times
    original_metadata = os.stat(str(filename))
    os.utime(f.name, ns=(original_metadata.st_atime_ns, original_metadata.st_mtime_ns, original_metadata.st_ctime_ns))

    # Replace the original file with the decrypted file
    shutil.move(f.name, str(filename))
