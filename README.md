# Mass Decrypt

Mass Decrypt is a command-line tool for batch decryption of Microsoft Office and PDF files in a specified directory on Windows systems.
The creation, modification and last access times are preserved from the encrypted files in the decrypted files.

## Features

- Decrypt Office files (DOC, DOCX, DOCM, XLS, XLSX, XLSM, XLSB) and PDF files.
- Use a single password or a list of passwords stored in a text file.
- Recursively process files in subdirectories.
- Perform integrity checks on decrypted PDF files (optional).

## Requirements

- Python 3.7 or higher
- [msoffcrypto-tool](https://github.com/nolze/msoffcrypto-tool) library for Office file decryption
- [pikepdf](https://github.com/pikepdf/pikepdf) library for PDF file decryption

This program is designed to run on Windows systems only.

## Installation

1. Clone the repository:

```git
git clone https://github.com/yourusername/mass_decrypt.git
```

2. Create a virtual environment and activate it:

```console
python -m venv venv
.\venv\Scripts\activate
```

3. Install the required packages:

```console
pip install -r requirements.txt
```

4. Install the package:

```console
pip install .
```

## Usage

To run Mass Decrypt, use the following command:

```console
mass_decrypt [OPTIONS]
```

### Options

- `path`: Path to the folder containing Office and/or PDF files.
- `-R`, `--recursive`: Recursively process subfolders.
- `-p`, `--password`: Password for decryption.
- `--plist`: Path to a .txt file containing a list of passwords.
- `-P`, `--pdf`: Decrypt PDF files in the specified folders and subfolders.
- `-O`, `--office`: Decrypt Office files in the specified folders and subfolders.
- `--integrity-check`: Perform an integrity check for decrypted PDF files (optional).

**Note**: Either `--password` or `--plist` must be provided.

## Example

```console
mass_decrypt -R -P -O --integrity-check --plist passwords.txt C:\path\to\folder
```

This command will process all Office and PDF files in the specified folder and its subfolders, attempting to decrypt them using the passwords listed in `passwords.txt`. It will also perform an integrity check for decrypted PDF files.

## Notes

Test units are still being developed, and will be introduced as new features are developed. Please note that although this program should not break PDF or Office files, we do not recommend to run it in a production environment.
An error will be trigerred and the program will stop in case of a damaged file.

## License

This project is licensed under the MIT License. See the [LICENSE](LICENSE) file for details.
