"""
Theoretically supports Windows, Mac, and Linux, but has only been tested on
Windows.
by: Christian M. Schrader
"""

import re
import os
import sys
import ctypes
import pathlib
from shutil import copyfile
from elevate import elevate


def admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False


def get_signature_path() -> pathlib.Path:
    """
    Returns the HPRC signiture path.
    """
    signPath = pathlib.Path.home()
    if sys.platform == "win32":
        signPath = signPath / "AppData/Roaming"
    elif sys.platform == "darwin":
        signPath = signPath / "Library/Application Support"
    elif sys.platform == "linux":
        signPath = signPath / ".local/share"
    else:
        print("Your OS is not supported.")
        exit()

    try:
        signPath / "Microsoft/Signatures/HPRC.htm"
    except FileNotFoundError:
        print("Could not find your Microsoft Signatures Directory.")
        exit()
    return signPath


if __name__ == "__main__":
    elevate()
    print("Please close Outlook before proceeding.")

    name = input("Full Name     :")
    title = input("HPRC Position :")
    major = input("Major         :")
    year = input("Class Year    :")

    sigPath = get_signature_path()
    copyfile("template.html", sigPath)
    with open(sigPath, "r") as sig:
        html = sig.read()
        html = re.sub(r"HNAME", name, html)
        html = re.sub(r"HPOSITION", title, html)
        html = re.sub(r"HMAJOR", major, html)
        html = re.sub(r"HCLASS", year, html)

    with open(sigPath, "w") as sig:
        sig.write(html)
