from setuptools import setup

APP = ['main.py']
DATA_FILES = []
OPTIONS = {
    'argv_emulation': True,
    'packages': ['pandas', 'openpyxl'],
    'iconfile': None,  # если появится иконка .icns
}

setup(
    app=APP,
    data_files=DATA_FILES,
    options={'py2app': OPTIONS},
    setup_requires=['py2app'],
)
