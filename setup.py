# -*- coding: utf-8 -*-

import os
import re

from setuptools import find_packages, setup


pkg_path = os.path.dirname(__file__)
with open(os.path.join(pkg_path, 'xls2any', '__init__.py')) as _ver_fp:
    VERSION = re.search(r"__version__ = '(.*?)'", _ver_fp.read()).group(1)


setup(
    name="xls2any",
    version=VERSION,
    description="Template engine use Excel as datasource",
    author="Mel Ts",
    author_email="layzerar@gmail.com",
    url="https://github.com/layzerar/xls2any",
    license="Apache License 2.0",
    packages=find_packages(),
    entry_points={
        'console_scripts': [
            'xls2any=xls2any.scripts.xls2any_:main',
        ],
    },
    zip_safe=True,
    install_requires=[
        'chardet>=3.0.2',
        'click>=6.6'
        'colorama>=0.3.8',
        'Jinja2>=2.9.5',
        'openpyxl>=2.5.0'
        'python-dateutil==2.6.1',
    ],
    classifiers=[
        'Private :: Do Not Upload',
    ],
)
