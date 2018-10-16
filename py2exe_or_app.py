import os
cmd1 = "pyinstaller -F 1_BOM_Checker.py"
cmd2 = "pyinstaller -F 2_BOM_Checker_Project.py"
cmd3 = "pyinstaller -F 3_PDF_Checker.py"
os.system(cmd1)
os.system(cmd2)
os.system(cmd3)
print('Converted Processing done...')
