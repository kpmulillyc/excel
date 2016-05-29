from distutils.core import setup
import py2exe, sys, os

sys.argv.append('py2exe')

DATA=['template.xlsx','msvcr100.dll']

setup(
    options = {'py2exe': {'compressed':True,'bundle_files': 1,'excludes': ['tkinter','PyQt4','PyQt5','PySide']}},
    windows = [{'script': "__init__.py"}],
    zipfile = None,
    data_files = DATA,
)