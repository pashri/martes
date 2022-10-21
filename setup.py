"""Setup"""

import base64
from pathlib import Path
from typing import List, Union
from setuptools import setup

NAME = 'martes'
VERSION = '0.1.0'
AUTHOR = 'Patrick Linton'
AUTHOR_EMAIL = b'cGF0cmlja2RhbmllbGxpbnRvbkBnbWFpbC5jb20='
PACKAGES = ['martes']
URL = 'https://github.com/pashri/martes'
DESCRIPTION = 'Read an Excel workbook in Pandas and reference cells by coordinates'


def parse_requirements_file(filename: Union[str, Path]) -> List[str]:
    """Parses a requirements file"""

    with open(filename, encoding='utf-8') as file:
        requires = [
            line.strip()
            for line in file.readlines()
            if not line.startswith("#")
        ]

    return requires

def parse_readme(filename: Union[str, Path]) -> str:
    """Parses a readme file"""

    with open(filename, encoding='utf-8') as file:
        text = file.read()

    return text

setup(
    name=NAME,
    version=VERSION,
    author=AUTHOR,
    author_email=base64.b64decode(AUTHOR_EMAIL).decode(),
    packages=PACKAGES,
    url=URL,
    description=DESCRIPTION,
    long_description=None,
    install_requires=parse_requirements_file('./requirements.txt'),
    tests_require=parse_requirements_file('./tests/requirements.txt'),
    classifiers=[
        'Intended Audience :: Developers',
        'License :: OSI Approved :: MIT License',
        'Natural Language :: English',
        'Operating System :: OS Independent',
        'Programming Language :: Python',
        'Programming Language :: Python :: 3',
    ],
)
