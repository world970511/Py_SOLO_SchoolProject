#!c:\users\yhs04\onedrive\desktop\solo\py_solo_project_school\py_solo_schoolproject\auto_xls\venv\scripts\python.exe
import xls
from sys import argv, stderr

filename = argv[1]
try:
	sheet = argv[2]
except:
	sheet = None

WB = xls.XLSFile(argv[1])
if sheet is None: sheet = WB.sheetnames(0)
print >>stderr, sheet
print xls.utils.csvify(WB.sheet(sheet))
