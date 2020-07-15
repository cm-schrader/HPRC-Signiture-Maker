# HPRC Signiture Maker
A simple Python script that generates and installs an Outlook signiture based off a template.  It was written to simplify creating signitures for the Officers of AIAA's HPRC.  The template uses regexes to fill the fields "HNAME", "HPOSITION", "HMAJOR", and "HCLASS" as well as any assets (such as images) in the assets folder.

Runs on Windows.  Theoretically runs on Mac and Linux although it has not been tested on these systems. 

## Use
If not installed allready, install Python 3 and Pip.

After downloading or cloning the project, install the requirements.
```
pip install requirements.txt 
```

Run ```hprcSignitor.py```

Follow the prompts in the console.
