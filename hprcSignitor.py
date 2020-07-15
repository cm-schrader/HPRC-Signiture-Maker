"""
A short script that automates

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


def get_signature_path():
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
        input("Your OS is not supported.\nPress [ENTER] to close.")
        exit()

    try:
        signPath = signPath / "Microsoft/Signatures"
    except:
        input(
            "Could not find/open your Microsoft Signatures Directory.\nPress [ENTER] to close.")
        exit()
    return signPath


if __name__ == "__main__":
    elevate()
    print("Please close Outlook before proceeding.")

    name = input("Full Name     :")
    title = input("HPRC Position :")
    major = input("Major         :")
    year = input("Class Year    :")

    try:
        sigPath = get_signature_path()
        assetPath = pathlib.Path.cwd() / "assets"
        for fileName in os.listdir(assetPath):
            copyfile(assetPath / fileName, sigPath / fileName)
        sigPath = sigPath / "HPRC.htm"
        copyfile("template.html", sigPath)
        with open(sigPath, "r") as sig:
            html = sig.read()
            html = re.sub(r"assets/", "", html)
            html = re.sub(r"HNAME", name, html)
            html = re.sub(r"HPOSITION", title, html)
            html = re.sub(r"HMAJOR", major, html)
            html = re.sub(r"HCLASS", year, html)

        with open(sigPath, "w") as sig:
            sig.write(html)
    except Exception as e:
        print("Opperation Failure!\nException: " + e)

    input(
        "Program is done. You may now reopen Outlook.\nPress [ENTER] to close.")
