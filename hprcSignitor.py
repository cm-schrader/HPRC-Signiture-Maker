import re
import os
import sys
import pathlib
from shutil import copyfile
import tkinter as tk
from tkinter import messagebox
from tkinter.filedialog import askdirectory

try:
    from elevate import elevate
except ModuleNotFoundError:
    os.system('pip install elevate')
finally:
    from elevate import elevate


def get_signature_path():
    """
    Returns the signature directory.
    """
    signPath = pathlib.Path.home()
    if sys.platform == "win32":
        signPath = signPath / "AppData/Roaming"
    elif sys.platform == "darwin":
        signPath = signPath / "Library/Application Support"
    elif sys.platform == "linux":
        signPath = signPath / ".local/share"
    else:
        messagebox.showerror(
            "OS Error", "Your OS is not supported.  Press ok to terminate the program.")
        exit()

    signPath = signPath / "Microsoft/Signatures"
    print(signPath)
    if not os.path.isdir(signPath):
        messagebox.showwarning("Signature Folder not Found",
            "Your Microsoft Signature Directory could not be found.  This likley means that Outlook for desktop is not installed."  
            "  Please select a directory to create the signature file in."
        )
        signPath = pathlib.Path(askdirectory(title="Signature Directory"))
        if signPath is None:
            messagebox.showerror("Invalid Filepath", "The filepath you provided is invalid, please try again.")
            exit()
    return signPath


def create_signature(name, title, major, year):
    if not name or not title or not major or not year:
        messagebox.showwarning(
            "Invalid Fields", "You left one or more fields blank.")
        return
    try:
        sigPath = get_signature_path()
        print(sigPath)
        # assetPath = pathlib.Path.cwd() / "assets"
        # for fileName in os.listdir(assetPath):
        #     copyfile(assetPath / fileName, sigPath / fileName)
        sigPath = sigPath / "HPRC.htm"
        copyfile("template.html", sigPath)
        with open(sigPath, "r") as sig:
            html = sig.read()
            # html = re.sub(r"assets/", "", html)
            html = re.sub(r"HNAME", name, html)
            html = re.sub(r"HPOSITION", title, html)
            html = re.sub(r"HMAJOR", major, html)
            html = re.sub(r"HCLASS", year, html)

        with open(sigPath, "w") as sig:
            sig.write(html)
            
        messagebox.showinfo("Signature Created",
                            "Your signature has been successfully created and can now be used in new emails.  The program will now close.")
    except Exception as e:
        messagebox.showerror(
            "Runtime Error", f"An unhandled exception has caused a program failure.\nERROR:{e}")
    exit()


if __name__ == "__main__":
    # elevate()
    win = tk.Tk()
    win.title("HPRC Signiture Maker")

    tk.Label(win, text="Please fill all fields in before clicking \"Create\".").grid(
        row=0, columnspan=2)
    tk.Label(win, text="Full Name").grid(row=1)
    tk.Label(win, text="Officer Position").grid(row=2)
    tk.Label(win, text="Major").grid(row=3)
    tk.Label(win, text="Graduation Year").grid(row=4)

    name = tk.Entry(win, width=40)
    title = tk.Entry(win, width=40)
    major = tk.Entry(win, width=40)
    year = tk.Entry(win, width=40)

    name.grid(row=1, column=1)
    title.grid(row=2, column=1)
    major.grid(row=3, column=1)
    year.grid(row=4, column=1)

    submit = tk.Button(win, text="Create", command=lambda: create_signature(
        name.get(), title.get(), major.get(), year.get()))
    submit.grid(row=5, columnspan=2)

    win.mainloop()
