#!/usr/bin/env python3.6
"""
Setup
"""
from setuptools import find_packages
from distutils.core import setup

version = "0.1.1"

with open('README.rst') as f:
    long_description = f.read()

setup(name='ofxstatement-intesasp',
      version=version,
      author="Di Campli D. Raffaele Jr",
      author_email="dcdrj.pub@gmail.com",
      url="https://github.com/Jacotsu/ofxstatement-intesasp",
      description=("Plugin for ofxstatement that supports Intesa San paolo"\
                   "xlsx file"),
      long_description=long_description,
      license="GPLv3",
      keywords=["ofx", "banking", "statement"],
      classifiers=[
          'Development Status :: 3 - Alpha',
          'Programming Language :: Python :: 3.6',
          'Natural Language :: English',
          'Topic :: Office/Business :: Financial :: Accounting',
          'Topic :: Utilities',
          'Environment :: Console',
          'Operating System :: OS Independent',
          'License :: OSI Approved :: GNU General Public License v3 (GPLv3)'],
      packages=find_packages('src'),
      package_dir={'': 'src'},
      namespace_packages=["ofxstatement", "ofxstatement.plugins"],
      entry_points={
          'ofxstatement':
          ['IntesaSP = ofxstatement.plugins.intesaSP:IntesaSanPaoloPlugin']
          },
      install_requires=['ofxstatement', 'openpyxl', 'dataclasses'],
      include_package_data=True,
      zip_safe=True
      )
