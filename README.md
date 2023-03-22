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
Substitute the zeroes with your bank's BIC/SWIFT code
```
[IntesaSP]
BIC = 0000000
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
  $ ofxIntesa IntesaSP Movimenti_Conto_<date>.xlsx Movimenti.ofx
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
