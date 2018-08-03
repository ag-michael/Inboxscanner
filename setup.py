from distutils.core import setup
import py2exe

setup(
	console=['outlook.py'],
	options = {'py2exe': {'bundle_files': 1, 'compressed': True,'packages': ['win32com', 'unidecode','requests','yara']}},
	zipfile=None
	)