import os
import pathlib
import msoffcrypto
import tempfile
import shutil
import logging
import argparse
from tqdm import tqdm

# Exception imports
from msoffcrypto.exceptions import FileFormatError, ParseError, DecryptionError, InvalidKeyError

def unlock(filename, passwd, loggername):
    logger = logging.getLogger(loggername)
    temp = open(filename, "rb")

    try:
        excel = msoffcrypto.OfficeFile(temp)

        if not excel.is_encrypted():
            logger.info(f"Not encrypted: {filename}")
            temp.close()
            return "unencrypted"

        excel.load_key(passwd)
    except FileFormatError as e:
        logger.warning(f"Unsupported format: {filename}, {e}")
        temp.close()
        return None
    except ParseError as e:
        logger.warning(f"Parse error: {filename}, {e}")
        temp.close()
        return None
    except DecryptionError as e:
        logger.warning(f"Decryption error: {filename}, {e}")
        temp.close()
        return None
    except InvalidKeyError as e:
        logger.warning(f"Invalid key error: {filename}, {e}")
        temp.close()
        return None
    except Exception as e:
        logger.error(f"Error: {filename}, {e}")
        temp.close()
        return None

    logger.info(f"Unlocked: {filename}")

    # The rest of the function remains the same

    with tempfile.TemporaryFile() as fout:
        excel.decrypt(fout)
        fout.seek(0)
        original_metadata = os.stat(filename)
        shutil.copyfileobj(fout, open(filename, "wb"))

        # Update the file metadata
        os.utime(str(filename), ns=(original_metadata.st_atime_ns, original_metadata.st_mtime_ns))

    temp.close()
    return "decrypted"

def main():
    parser = argparse.ArgumentParser(description="Unlock Excel files.")
    parser.add_argument("path", help="Path to the folder containing Excel files.")
    parser.add_argument("-R", "--recursive", action="store_true", help="Recursively process subfolders.")
    parser.add_argument("-p", "--password", help="Password for decryption.")
    parser.add_argument("--plist", help="Path to a .txt file containing a list of passwords.")
    args = parser.parse_args()

    if not args.password and not args.plist:
        parser.error("Either --password or --plist must be provided")

    if args.password:
        passwords = [args.password]
    else:
        with open(args.plist, "r") as f:
            passwords = [line.strip() for line in f.readlines()]

    path = pathlib.Path(args.path)
    if not path.is_dir():
        print(f"'{args.path}' is not a valid directory")
        return

    logging.basicConfig(level=logging.INFO, handlers=[logging.NullHandler()])
    logger = logging.getLogger("excel_unlocker")

    decrypted_count = 0
    unencrypted_count = 0

    if args.recursive:
        all_files = [(os.path.join(root, file), file)
                     for root, _, files in os.walk(args.path)
                     for file in files
                     if file.endswith(".xls")]
    else:
        all_files = [(str(file), file.name) for file in path.glob("*.xls")]

    with tqdm(total=len(all_files), unit="file") as progress_bar:
        for filepath, filename in all_files:
            progress_bar.set_description(f"Processing: {filename}")
            for password in passwords:
                result = unlock(filepath, password, "excel_unlocker")
                if result == "decrypted":
                    decrypted_count += 1
                    break
                elif result == "unencrypted":
                    unencrypted_count += 1
                    break
            progress_bar.update(1)

    print(f"\nDecrypted files: {decrypted_count}")
    print(f"Unencrypted files: {unencrypted_count}")

if __name__ == "__main__":
    main()