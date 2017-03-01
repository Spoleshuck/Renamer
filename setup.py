from distutils.core import setup

setup(
	windows=[{"script": "Renamer.py", "icon_resources": [(0, "Icon.ico")]}], requires=['openpyxl', 'tkinter'],
	options={"py2exe": {"optimize": 2, "bundle_files": 2}},
	data_files=[('images', ['Icon.ico'])])
