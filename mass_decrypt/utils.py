"""
This module contains utility functions for the mass_decrypt application
"""

import logging
from ctypes import windll, wintypes, byref


def setup_logger(logger_name):
    """
    Set up the logger with the given logger_name

    Args:
        logger_name (str): Name of the logger

    Returns:
        logger (logging.Logger): Configured logger
    """
    logger = logging.getLogger(logger_name)

    if not logger.handlers:
        logger.setLevel(logging.DEBUG)
        fh = logging.FileHandler("../runtime.log")
        fh.setLevel(logging.DEBUG)
        formatter = logging.Formatter("%(asctime)s - %(levelname)s - %(message)s")
        fh.setFormatter(formatter)
        logger.addHandler(fh)

    return logger


def to_windows_filetime(epoch):
    """
    Convert a UNIX timestamp to a Windows FILETIME timestamp
    
    Args:
        epoch (int): UNIX timestamp
        
    Returns:
        wintypes.FILETIME: Windows FILETIME timestamp
    """
    timestamp = int((epoch * 10000000) + 116444736000000000)
    return wintypes.FILETIME(timestamp & 0xFFFFFFFF, timestamp >> 32)


def set_file_times(filename, atime, mtime, ctime):
    """
    Set the access, modification, and creation times for a file in Windows
    
    Args:
        filename (str): The path of the file
        atime (float): The access time
        mtime (float): The modification time
        ctime (float): The creation time
    """
    atime_struct, mtime_struct, ctime_struct = (
        to_windows_filetime(atime),
        to_windows_filetime(mtime),
        to_windows_filetime(ctime),
    )

    handle = windll.kernel32.CreateFileW(filename, 256, 0, None, 3, 128, None)
    windll.kernel32.SetFileTime(
        handle, byref(ctime_struct), byref(atime_struct), byref(mtime_struct)
    )
    windll.kernel32.CloseHandle(handle)
