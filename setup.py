from distutils.core import setup
import py2exe

setup(
	windows=[
		{
			"script": "Renamer.py",
			"icon_resources": [(0, "Icon.ico")]
		}
	], requires=['openpyxl','tkinter']
)
