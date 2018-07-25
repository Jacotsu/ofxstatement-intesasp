~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
IntesaSP plugin for ofxstatement
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

This plugin parses the Intesa San paolo xlsx statement file

Installation
============

You can install the plugin as usual from pip or directly from the downloaded git

pip
---

::

  pip3 install --user ofxstatement-intesasp

setup.py
--------

::

  python3 setup.py install --user

Configuration
===============================
To edit the config file run this command
::

  $ ofxstatement edit-config


Substitute the zeroes with your bank's BIC/SWIFT code
::

  [IntesaSP]
  BIC = 0000000
  plugin = IntesaSP

Save and exit the text editor

Usage
================================
Download your transactions file from the official bank's site and
then run

::

  $ ofxstatement convert -t IntesaSP Movimenti_Conto_<date>.xlsx Movimenti.ofx
