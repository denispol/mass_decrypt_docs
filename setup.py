from setuptools import setup, find_packages

with open("README.md", "r") as fh:
    long_description = fh.read()

setup(
    name="mass_decrypt",
    version="1.0.0",
    author="denispol",
    author_email="",
    description="A command-line tool for batch decryption of Microsoft Office and PDF files on Windows systems.",
    long_description=long_description,
    long_description_content_type="text/markdown",
    url="https://github.com/denispol/mass_decrypt",
    packages=find_packages(),
    entry_points={
        "console_scripts": [
            "mass_decrypt=mass_decrypt.main:main",
        ],
    },
    classifiers=[
        "Development Status :: 3 - Alpha",
        "Intended Audience :: End Users/Desktop",
        "License :: OSI Approved :: MIT License",
        "Operating System :: Microsoft :: Windows",
        "Programming Language :: Python :: 3",
        "Programming Language :: Python :: 3.7",
        "Programming Language :: Python :: 3.8",
        "Programming Language :: Python :: 3.9",
    ],
    python_requires=">=3.7",
    install_requires=[
        "msoffcrypto-tool",
        "pikepdf",
        "openpyxl",
        "xlrd",
        "tqdm"
    ],
    
)
