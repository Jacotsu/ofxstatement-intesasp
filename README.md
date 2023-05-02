# IntesaSP plugin for ofxstatement

This plug-in parses the Intesa San paolo xlsx statement file, main project can be found [here](https://github.com/kedder/ofxstatement).

[TOC]

## Installation

You can install the plugin as usual from pip or directly from the downloaded git

### pip
```bash
pip3 install --user ofxstatement-intesasp
```
### setup.py
```bash
python3 setup.py install --user
```
## Configuration
*(For now, this configuration is ignored, is used in other plug-in of the main project)*

To edit the config file run this command

```bash
$ ofxstatement edit-config
```
Substitute the zeroes with your bank's ABI and CAB codes (The italian equivalent of BIC code)
```
[IntesaSP]
ABI = 00000
CAB = 00000
plugin = IntesaSP
```

Save and exit the text editor


## Usage
Download your transactions file from the official bank's site and then run
```bash
$ ofxstatement convert -t IntesaSP Movimenti_Conto_<date>.xlsx Movimenti.ofx
```

### Add Alias
To simplify the use of the plugin, we strongly recommend adding an alias to your system (if in a Linux environment or on an emulated terminal) by adding the alias of this command to your *.bash_aliases*:
```bash
$ printf '\n# Intesa excel convert to OFX format\nalias ofxIntesa="ofxstatement convert -t IntesaSP"\n' >> ~/.bash_aliases
```
After that, reload your terminal (close and then reopen) and the usage change to:
```bash
  $ ofxIntesa Movimenti_Conto_<date>.xlsx Movimenti.ofx
```
**Note**: If after reload alias are not loading, go in your *.bashrc* and check if follow line are present, if not, add it on the end:
```bash
  # Alias definitions.
  # You may want to put all your additions into a separate file like
  # ~/.bash_aliases, instead of adding them here directly.
  # See /usr/share/doc/bash-doc/examples in the bash-doc package.

  if [ -f ~/.bash_aliases ]; then
      . ~/.bash_aliases
  fi
```

## How use OFX file after conversion

The `ofx` format stands for '*Open Financial Exchange*', it can be used to transfer your accounting records from one database to another.
This repository in particular allows you to convert the records that **Intesa San Paolo** shares via Excel into this *open source* format.
Once you have the `ofx` file, you can use any program to manage your finances.
Among the many available, a non-exhaustive list of open source products is:

- [HomeBank](http://homebank.free.fr/en/index.php), continuously updated program, present everywhere except in smartphones, with many beautiful ideas and listening to the community. **100% compatibility** 
