# HPRC Signiture Maker
A simple Python program that generates and installs an Outlook signiture based off a template.  It was written to simplify creating signitures for the Officers of AIAA's HPRC.  The script uses regexes to fill the fields "HNAME", "HPOSITION", "HMAJOR", and "HCLASS" in a template as well as include any assets (such as images) in the assets folder.  The resulting HTML signature is created in the Microsoft Signature folder where it can be immedietly used in any new email.

__Note: This only works for Outlook for Desktop.  For other platforms, generate a signature for Outlook for desktop and follow a tutorial such as [this one](https://blog.gimm.io/add-email-signature-outlook-app-ios/).__

Verified to run on Windows.  Theoretically runs on Mac and Linux although it has not been tested on these systems. 

## Use
1. If not installed allready, install Python 3.8 or greater and Pip.  You must have Python and Pip added to Path.  This is an option in the installer, but if you missed it you can add it manually.  If you are unfamiliar with Python and have no idea what Path is, follow [this guide](https://realpython.com/installing-python/).  _The package elevate is required.  Typically this must be done manually but for the convenience of nonprogrammer officers, the script will automatically install it if is not installed._

1. Click on the green "Code" button at the top of this page.  If you have GIT hub, you can clone the repository, otherwise, just click "download ZIP"
   1. If you downloaded the zip, be sure to extract the folder from it.  You will find the program inside of the folder.

1. Run ```hprcSignitor.py```

1. Complete all the fields.
   1. Full Name: Your perfered full name.  ei: "Christian M. Schrader"
   1. Officer Position: Your official title on the Officer Board of HPRC.  ei: "Documentation Officer"
   1. Major: The title of your major.  ei: "Aerospace Engineering"
   1. Graduation Year: The year you expect to graduate with your BS.  ei: "2021"

1. Click "Create".

1. The program will then add your new signature to Outlook desktop before closing itself.  You can now use your new signature in Outlook for desktop.
