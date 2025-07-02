"""
Setup script for creating macOS application with py2app
Package application as a clickable MacOS application with py2app as specified in prompt.md
"""

from setuptools import setup
import os

# Get the current directory
current_dir = os.path.dirname(os.path.abspath(__file__))

APP = ['aig_gui.py']
DATA_FILES = [
    ('', ['aig_processor.py']),
    ('', ['requirements.txt']),
    ('', ['README.md']),
    ('input', [
        'input/SalemAIGRoster6.24.25.pdf',
        'input/HEINZE of  25-26 Class Lists.xlsx',
        'input/TD from Finch WCPSS file.docx'
    ]),
]

OPTIONS = {
    'argv_emulation': False,  # Disable to avoid Carbon framework issues on modern macOS
    'plist': {
        'CFBundleName': 'AIG Class List Processor',
        'CFBundleDisplayName': 'AIG Class List Processor',
        'CFBundleGetInfoString': 'AIG Class List Processor - Process class lists and generate AIG reports',
        'CFBundleIdentifier': 'com.teacher.aig-processor',
        'CFBundleVersion': '1.0.0',
        'CFBundleShortVersionString': '1.0.0',
        'NSHumanReadableCopyright': 'Copyright © 2025 Teacher Tools',
        'NSHighResolutionCapable': True,
        'LSMinimumSystemVersion': '10.13',
        'LSUIElement': False,  # Show in dock
        'CFBundleDocumentTypes': [
            {
                'CFBundleTypeExtensions': ['pdf', 'xlsx', 'docx'],
                'CFBundleTypeName': 'AIG Input Files',
                'CFBundleTypeRole': 'Viewer',
                'LSHandlerRank': 'Alternate',
            }
        ],
    },
    'packages': [
        'pandas', 
        'openpyxl', 
        'PyPDF2', 
        'docx', 
        'tkinter',
        'numpy',  # pandas dependency
        'xlsxwriter',  # often used with pandas
        'logging',
        'threading',
        'sys',
        'os',
        're'
    ],
    'includes': [
        'tkinter', 
        'tkinter.filedialog', 
        'tkinter.messagebox', 
        'tkinter.ttk',
        'pandas',
        'openpyxl',
        'PyPDF2',
        'docx',
        'logging',
        'threading',
        'numpy',
        'xlsxwriter'
    ],
    'excludes': ['matplotlib'],
    'resources': [],
    'iconfile': None,  # We could add an icon file here if we had one
    'site_packages': True,  # Include site packages
    'strip': False,  # Don't strip debug symbols for better error reporting
}

setup(
    name='AIG Class List Processor',
    app=APP,
    data_files=DATA_FILES,
    options={'py2app': OPTIONS},
    setup_requires=['py2app'],
    install_requires=[
        'pandas>=2.0.0',
        'openpyxl>=3.1.0',
        'PyPDF2>=3.0.0',
        'python-docx>=0.8.11',
        'numpy>=1.21.0',
        'xlsxwriter>=3.1.0',
    ],
)
