# Office Unlocker

Office Unlocker is a command-line tool that helps decrypt password-protected Office files in batch. It supports both single password input and a list of passwords provided in a text file.

## Features

- Decrypt multiple Office files in a directory
- Decrypt Office files recursively in subdirectories
- Use a single password or multiple passwords from a text file
- Progress bar displaying the current file being processed
- Summary of the total number of decrypted and unencrypted files

## Installation

1. Clone the repository:
```git
git clone https://github.com/denispol/excel_unlocker.git
```

2. Install required libraries:
```python
pip install -r requirements.txt
```

## Usage

Basic usage:
```python
python unlock_office.py <path-to-folder> -p <password>
```

Use a list of passwords from a text file:
```python
python unlock_office.py <path-to-folder> --plist <path-to-passwords.txt>
```

Process subdirectories recursively:
```python
python unlock_office.py <path-to-folder> -R -p <password>
```

For a complete list of options, use the help:
```python
python unlock_office.py -h
```

# Example
Decrypt all Excel files in the "example" folder using a single password:
```python
python unlock_office.py example -p mysecretpassword
```

Decrypt all Excel files in the "example" folder using a list of passwords from "passwords.txt":
```python
python unlock_office.py example --plist passwords.txt
```

# License
This project is released under the MIT License. See LICENSE for more information.
